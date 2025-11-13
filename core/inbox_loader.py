# core/inbox_loader.py
from __future__ import annotations
import pandas as pd
import streamlit as st

REQUIRED_COLS = [
    "token","管理番号","通報者","受信内容","現着状況","原因","処置内容",
    "所属","処理修理後","受信時刻","物件名","住所","メーカー","制御方式",
    "契約種別","受付番号","受付URL","現着完了登録URL"
]

def _get_csv_url() -> str:
    try:
        url = st.secrets["SHEET_CSV_URL"].strip()
    except Exception:
        raise RuntimeError(
            "SHEET_CSV_URL が secrets に未設定です。"
            " `.streamlit/secrets.toml` または Cloud Secrets に "
            'SHEET_CSV_URL="https://docs.google.com/spreadsheets/d/<ID>/gviz/tq?tqx=out:csv&sheet=inbox" '
            "を設定してください。"
        )
    if not url:
        raise RuntimeError("SHEET_CSV_URL が空文字です。正しいCSVエクスポートURLを設定してください。")
    return url

def _read_csv(url: str) -> pd.DataFrame:
    try:
        # すべて文字列で読み込み（NaN→空文字に統一）
        df = pd.read_csv(url, dtype=str).fillna("")
        return df
    except Exception as e:
        raise RuntimeError(f"CSVの読み込みに失敗しました: {e}")

def load_from_sheet_by_token(token: str) -> dict | None:
    """
    Gmail→GASでシートに追記された1行を token で取得し、辞書で返す。
    見つからなければ None。
    """
    if not token:
        return None

    url = _get_csv_url()
    df = _read_csv(url)

    if "token" not in df.columns:
        raise RuntimeError('CSVに "token" 列が見つかりません。シートのヘッダ行をREADME通りにしてください。')

    hit = df.loc[df["token"] == token]
    if hit.empty:
        return None

    rec = hit.iloc[0].to_dict()

    # 期待列が無ければ空文字で補完（テンプレ生成時のKeyErrorを防止）
    for c in REQUIRED_COLS:
        rec.setdefault(c, "")

    return rec

