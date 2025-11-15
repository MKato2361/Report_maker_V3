# core/inbox_loader.py
from __future__ import annotations

from typing import Dict
import os
import unicodedata

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
    1行目はヘッダーとして扱う。
    """
    url = _get_csv_url()
    # UTF-8 BOM 対策で encoding を utf-8-sig にしておく
    df = pd.read_csv(url, dtype=str, encoding="utf-8-sig")
    df = df.fillna("")
    return df


def _norm(s: str) -> str:
    """ヘッダー名の正規化（NFKC＋小文字＋前後空白除去）"""
    return unicodedata.normalize("NFKC", str(s)).strip().lower()


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
        P: 処置内容   ← 問題の列（index 15）
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

    # --- どの列が token なのかを「名前」で特定（BOMや全角を吸収） ---
    norm_cols = [_norm(c) for c in df.columns]
    try:
        token_col = df.columns[norm_cols.index("token")]
    except ValueError:
        raise RuntimeError(
            f'CSV のヘッダーに "token" 列が見つかりません。現在のヘッダー: {list(df.columns)!r}'
        )

    # token で対象行を絞り込む
    sub = df[df[token_col] == token]
    if sub.empty:
        raise KeyError(f"token={token!r} の行が見つかりません。")

    # 単一行
    row = sub.iloc[0]

    # DataFrame 上の「列名」と「値」をそのまま記録しておく（デバッグ用）
    columns = [str(c) for c in sub.columns]
    values_raw = [("" if v is None else str(v)) for v in row.tolist()]

    # 想定キー（inbox の列順と 1:1 対応させる）
    pos_keys = [
        "token",        # index 0  = A列
        "管理番号",      # 1        = B列
        "物件名",        # 2        = C列
        "住所",          # 3        = D列
        "窓口会社",      # 4        = E列
        "メーカー",      # 5        = F列
        "制御方式",      # 6        = G列
        "契約種別",      # 7        = H列
        "受信時刻",      # 8        = I列
        "現着時刻",      # 9        = J列
        "完了時刻",      # 10       = K列
        "通報者",        # 11       = L列
        "受信内容",      # 12       = M列
        "現着状況",      # 13       = N列
        "原因",          # 14       = O列
        "処置内容",      # 15       = P列 ← ここが本命
        "対応者",        # 16       = Q列
        "送信者",        # 17       = R列
        "完了連絡先1",    # 18       = S列
        "受付番号",      # 19       = T列
        "受付URL",       # 20       = U列
        "現着完了登録URL",# 21       = V列
        "所属",          # 22       = W列
        "処理修理後",     # 23       = X列
        "作業時間_分",    # 24       = Y列
    ]

    # values_raw の長さと pos_keys の長さがズレていればここで補正
    if len(values_raw) < len(pos_keys):
        # 足りない分は空文字で埋める
        values = values_raw + [""] * (len(pos_keys) - len(values_raw))
    else:
        # 余っている場合は pos_keys の分だけ使う
        values = values_raw[: len(pos_keys)]

    # 位置で dict 化
    data_by_pos = {
        key: (val.strip() if isinstance(val, str) else "")
        for key, val in zip(pos_keys, values)
    }

    # token は返さない
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

    # --- デバッグ情報を付与（Step3 の「🛠デバッグ」枠で確認用） ---
    # 列名の順番（index付き）
    rec["_DEBUG_COLUMNS"] = " | ".join(f"[{i}]{c}" for i, c in enumerate(columns))
    # 行の値（index付き）
    rec["_DEBUG_VALUES"] = " | ".join(f"[{i}]{v}" for i, v in enumerate(values_raw))

    return rec