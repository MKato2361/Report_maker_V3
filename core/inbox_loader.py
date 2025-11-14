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
    1行目はヘッダーだが、列順だけを信用し、列名はほぼ無視する。
    """
    url = _get_csv_url()
    df = pd.read_csv(url, dtype=str)
    df = df.fillna("")
    return df


def load_from_sheet_by_token(token: str) -> Dict[str, str]:
    """
    inbox シート由来の CSV から token 行を 1件だけ取得し、
    列の“位置”を決め打ちして Step3 用の辞書を返す。

    ★ 前提：inbox の列順が以下で固定されていること
        A: token
        B: 管理番号
        C: 物件名
        D: 住所
        E: 窓口会社
        F: メーカー
        G: 制御方式
        H: 契約種別
        I: 受信時刻
        J: 現着時刻
        K: 完了時刻
        L: 通報者
        M: 受信内容
        N: 現着状況
        O: 原因
        P: 処置内容
        Q: 対応者
        R: 送信者
        S: 完了連絡先1
        T: 受付番号
        U: 受付URL
        V: 現着完了登録URL
        W: 所属
        X: 処理修理後
        Y: 作業時間_分
    """
    df = _load_dataframe()
    if "token" not in df.columns:
        raise RuntimeError('CSV に "token" 列がありません。inbox シートの1列目に token を配置してください。')

    sub = df[df["token"] == token]
    if sub.empty:
        raise KeyError(f"token={token!r} の行が見つかりません。")

    # 単一行の値を「配列」として取得（列名は一切信用しない）
    row = sub.iloc[0]
    values = list(row)

    # 想定キー（列順と1対1対応）
    pos_keys = [
        "token",        # A列
        "管理番号",      # B列
        "物件名",        # C列
        "住所",          # D列
        "窓口会社",      # E列
        "メーカー",      # F列
        "制御方式",      # G列
        "契約種別",      # H列
        "受信時刻",      # I列
        "現着時刻",      # J列
        "完了時刻",      # K列
        "通報者",        # L列
        "受信内容",      # M列
        "現着状況",      # N列
        "原因",          # O列
        "処置内容",      # P列 ★ここが確実に処置内容
        "対応者",        # Q列
        "送信者",        # R列
        "完了連絡先1",    # S列
        "受付番号",      # T列
        "受付URL",       # U列
        "現着完了登録URL",# V列
        "所属",          # W列
        "処理修理後",     # X列
        "作業時間_分",    # Y列
    ]

    # CSV 側の列数が足りなければ空文字で埋める
    if len(values) < len(pos_keys):
        values = values + [""] * (len(pos_keys) - len(values))

    # 余分な列があっても無視（pos_keys の長さに合わせる）
    values = values[:len(pos_keys)]

    # 位置で dict 化
    data_by_pos = {k: (str(v).strip() if v is not None else "") for k, v in zip(pos_keys, values)}

    # token は返さない（必要ならここで返してもよいが、UI 側では使っていない）
    rec: Dict[str, str] = {k: v for k, v in data_by_pos.items() if k != "token"}

    # 念のため、想定キーをすべて埋めておく
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