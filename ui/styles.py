# report_maker/ui/styles.py
import streamlit as st

def inject_styles():
    st.markdown(
        """
        <style>
        header {visibility: hidden;}
        .block-container {padding-top: 0rem;}
        .edit-toolbar { position: sticky; top: 0; z-index: 50; backdrop-filter: blur(6px);
          background: rgba(30,30,30,0.08); padding: .5rem .75rem; border-radius: .5rem; margin-bottom: .5rem; }
        .edit-toolbar .btn-row { display: flex; gap: .5rem; align-items: center; flex-wrap: wrap; }
        .edit-badge { font-size: .85rem; background: #ffd24d; color: #4a3b00; padding: .15rem .5rem; border-radius: .5rem; margin-left: .25rem; }
        .missing { color: #b00020; font-weight: 600; }
        </style>
        """,
    st.markdown(
        """
        <style>
        /* すべてのボタンを少し薄めに（必要なら値は微調整OK） */
        .stButton > button {
          padding: 4px 10px !important;
          font-size: 0.9rem !important;
          border-radius: 6px !important;
          border: 1px solid #ddd !important;
          background: transparent !important;
          color: inherit !important;
        }
        /* プライマリっぽいもの（保存など）を少しだけ目立たせる */
        .stButton > button[kind="primary"], .stButton > button:has(svg) {
          border-color: #cfcfcf !important;
          background: #f7f7f7 !important;
        }
        </style>
        """,
        unsafe_allow_html=True,
    )
