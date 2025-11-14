# core/inbox_loader.py
from __future__ import annotations

from typing import Dict
import os

import pandas as pd
import streamlit as st


def _get_csv_url() -> str:
    """
    SHEET_CSV_URL を secrets または環境変数から取得する。
    - .streamlit/secrets.toml に SHEET_CSV_URL="https://docs.google.com/....export?format=csv&gid=0"
    - もしくは環境変数 SHEET_CSV_URL に同じURL
    のどちらかで設定しておく想定。
    """
    url = ""
    try:
        url = st.secrets.get("SHEET_CSV_URL", "")  # type: ignore[attr-defined]
    except Exception:
        url = ""
    if not url:
        url = os.getenv("SHEET_CSV_URL", "").strip()
    if not url:
        raise RuntimeError("SHEET_CSV_URL が secrets か環境変数に設定されていません。")
    return url


def _load_dataframe():
    """
    Google スプレッドシートの CSV (export?format=csv...) を DataFrame で取得。
    """
    url = _get_csv_url()
    df = pd.read_csv(url, dtype=str)
    # NaN を空文字に統一
    df = df.fillna("")
    return df


def load_from_sheet_by_token(token: str) -> Dict[str, str]:
    """
    inbox シート由来の CSV から token 行を 1件だけ取得し、
    Step3 用の辞書（extracted と同じキー構成）を返す。
    """
    df = _load_dataframe()
    if "token" not in df.columns:
        raise RuntimeError('CSV に "token" 列がありません。inbox シートの1列目に token を配置してください。')

    sub = df[df["token"] == token]
    if sub.empty:
        raise KeyError(f"token={token!r} の行が見つかりません。")

    row = sub.iloc[0].to_dict()

    # まずヘッダー名をそのままキーにして取り込む
    rec: Dict[str, str] = {}
    for col, val in row.items():
        if col == "token":
            continue
        v = (val or "").strip()
        rec[col] = v

    # 想定キーを一通り揃えておく（存在しないものは空文字に）
    expected_keys = [
        "管理番号",
        "物件名",
        "住所",
        "窓口会社",
        "メーカー",
        "制御方式",
        "契約種別",
        "受信時刻",
        "現着時刻",
        "完了時刻",
        "通報者",
        "受信内容",
        "現着状況",
        "原因",
        "処置内容",
        "対応者",
        "送信者",
        "完了連絡先1",
        "受付番号",
        "受付URL",
        "現着完了登録URL",
        "所属",
        "処理修理後",
        "作業時間_分",
    ]
    for key in expected_keys:
        rec.setdefault(key, "")

    return rec