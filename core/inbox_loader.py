# core/inbox_loader.py
from __future__ import annotations

from typing import Dict
import os

import pandas as pd
import streamlit as st


def _get_csv_url() -> str:
    """
    SHEET_CSV_URL を secrets または環境変数から取得する。
    - .streamlit/secrets.toml に SHEET_CSV_URL="https://docs.google.com/...export?format=csv&gid=0"
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


def _load_dataframe() -> pd.DataFrame:
    """
    Google スプレッドシートの CSV (export?format=csv...) を DataFrame で取得。
    """
    url = _get_csv_url()
    df = pd.read_csv(url, dtype=str)
    df = df.fillna("")
    return df


def load_from_sheet_by_token(token: str) -> Dict[str, str]:
    """
    inbox シート由来の CSV から token 行を 1件だけ取得し、
    Step3 用の辞書（extracted と同じキー構成）を返す。

    ★ ポイント：
      - 列名が多少ブレていても alias で正規化
      - 既に入っている値を「空文字」で上書きしない（長い方を優先）
      - 処置内容は列名と列位置の両方から最後まで粘って拾う
    """
    df = _load_dataframe()
    if "token" not in df.columns:
        raise RuntimeError('CSV に "token" 列がありません。inbox シートの1列目に token を配置してください。')

    sub = df[df["token"] == token]
    if sub.empty:
        raise KeyError(f"token={token!r} の行が見つかりません。")

    # 単一行の Series
    row_series = sub.iloc[0]

    # ヘッダー名 → 正規化されたキーへのマッピング
    header_aliases: Dict[str, str] = {
        "窓口": "窓口会社",
        "窓口会社": "窓口会社",
        "現着時間": "現着時刻",
        "現着時刻": "現着時刻",
        "完了時間": "完了時刻",
        "完了時刻": "完了時刻",
        "処置": "処置内容",
        "処置内容": "処置内容",
        "詳細はこちら": "受付URL",
        "受付URL": "受付URL",
        "現着・完了登録はこちら": "現着完了登録URL",
        "現着完了登録URL": "現着完了登録URL",
    }

    rec: Dict[str, str] = {}

    # 1) 行全体を alias を使って辞書化（非空のみ・長い方優先）
    for col, val in row_series.to_dict().items():
        if col == "token":
            continue
        v = (val or "").strip()
        col_norm = str(col).strip()
        key = header_aliases.get(col_norm, col_norm)

        if not v:
            # 空はここでは入れない（後で setdefault で補充）
            continue

        # すでに値がある場合は「長い方」を優先（空や短い値で潰さない）
        if key not in rec or len(v) > len(rec.get(key, "")):
            rec[key] = v

    # 2) 処置内容 専用の保険
    #    - まだ空なら DataFrame から直接拾う
    if not (rec.get("処置内容") or "").strip():
        # a) 正確な列名「処置内容」が存在する場合
        if "処置内容" in sub.columns:
            v = str(sub["処置内容"].iloc[0] or "").strip()
            if v:
                rec["処置内容"] = v

        # b) それでも空なら「処置」を含む列を総当たり
        if not (rec.get("処置内容") or "").strip():
            for col in sub.columns:
                if col == "token":
                    continue
                col_str = str(col).strip()
                if "処置" in col_str:
                    v = str(sub[col].iloc[0] or "").strip()
                    if v:
                        rec["処置内容"] = v
                        break

        # c) さらに保険：列の並び順から P列(処置内容) を決め打ちで読む
        #    inbox ヘッダー：
        #      A:token, B:管理番号, ... P:処置内容, ...
        if not (rec.get("処置内容") or "").strip():
            cols = list(sub.columns)
            if len(cols) >= 16:  # 0〜15 で少なくとも16列
                try:
                    # 処置内容は 0始まりで 15 番目（P列）
                    col_name_by_pos = cols[15]
                    v = str(sub[col_name_by_pos].iloc[0] or "").strip()
                    if v:
                        rec["処置内容"] = v
                except Exception:
                    pass

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