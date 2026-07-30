# -*- coding: utf-8 -*-
"""
製品構成 逆検索ツール（Streamlit版）
  Excel/VBA版 製品構成逆検索ツール Ver1.04 を Streamlit へ移植したもの。
  検索ロジック(core.py)・マスタ取得層(dbsource.py)は他版と共通。

  (1) 機種コード逆展開検索 : 部品 → 使用する最上位機種コード・使用数・客先品目コード
  (2) 直上親コード検索     : 部品 → 直上親(TD-仮親はスキップ)・名称・代表機種・メーカー
  (3) CSV / Excel ダウンロード

DBへは SELECT のみ（書込みは行わない）。tdread は読取り専用ユーザー。

起動:
  本番   : streamlit run streamlit_app.py
  デモ   : streamlit run streamlit_app.py -- --demo
"""

import os
import io
import csv
import sys
import configparser

import streamlit as st

import core
import dbsource

BASE = os.path.dirname(os.path.abspath(__file__))
cfg = configparser.ConfigParser()
cfg.read(os.path.join(BASE, "config.ini"), encoding="utf-8")

DEMO = ("--demo" in sys.argv) or (
    cfg.get("app", "demo", fallback="false").lower() == "true")


def db_cfg():
    if not cfg.has_section("db"):
        return {}
    return dict(cfg.items("db"))


# ---- マスタは一度読んだらキャッシュ（再実行のたびに再取得しない）----
@st.cache_resource(show_spinner=False)
def load_index():
    if DEMO:
        return dbsource.load_index_demo()
    return dbsource.load_index_from_db(db_cfg())


def _fmt(v):
    if isinstance(v, float):
        return int(v) if v == int(v) else v
    return "" if v is None else v


def rows_to_table(rows, cols):
    keys = [k for k, _ in cols]
    titles = [t for _, t in cols]
    table = []
    for r in rows:
        table.append({t: _fmt(r.get(k, "")) for k, t in zip(keys, titles)})
    return table, keys, titles


def to_csv_bytes(rows, keys, titles):
    buf = io.StringIO()
    w = csv.writer(buf, lineterminator="\r\n")
    w.writerow(titles)
    for r in rows:
        w.writerow([("" if r.get(k) is None else
                     (int(r[k]) if isinstance(r.get(k), float) and r[k] == int(r[k])
                      else r.get(k, ""))) for k in keys])
    return buf.getvalue().encode("cp932", errors="replace")


def to_xlsx_bytes(mode, rows, keys, titles):
    try:
        from openpyxl import Workbook
        from openpyxl.styles import Font, Alignment, PatternFill
        from openpyxl.utils import get_column_letter
    except ImportError:
        return None
    wb = Workbook()
    ws = wb.active
    ws.title = "機種コード" if mode == "kishu" else "直上親コード"
    ws.append(titles)
    fill = PatternFill("solid", fgColor="2980B9")
    font = Font(color="FFFFFF", bold=True)
    for c in ws[1]:
        c.fill = fill
        c.font = font
        c.alignment = Alignment(horizontal="center")
    for r in rows:
        row = []
        for k in keys:
            v = r.get(k, "")
            if isinstance(v, float):
                v = int(v) if v == int(v) else v
            row.append(v)
        ws.append(row)
    for i in range(1, len(keys) + 1):
        ws.column_dimensions[get_column_letter(i)].width = 18
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.dimensions
    bio = io.BytesIO()
    wb.save(bio)
    return bio.getvalue()


# =========================================================================
# 画面
# =========================================================================
st.set_page_config(page_title="製品構成 逆検索ツール", layout="wide")
st.title("製品構成 逆検索ツール")

# マスタ読込（初回のみ。失敗時はエラー表示）
index = None
load_error = None
with st.spinner("マスタを読み込んでいます...（初回のみ／最大2分）"):
    try:
        index = load_index()
    except Exception as e:
        load_error = e

# ステータス表示
if DEMO:
    st.info("● デモモード（DB未接続・内蔵サンプルデータ）", icon="🧪")
elif index is not None:
    n_item = len(index.info_bunr)
    n_prts = sum(len(v) for v in index.dict_fwd.values())
    st.success("● 本番（マスタ取得済）　品目 %d件 / 構成 %d行" % (n_item, n_prts), icon="✅")
else:
    st.error("マスタ読込に失敗しました: %s" % load_error, icon="⚠️")
    st.caption("通信環境（VPNなど）と config.ini の接続情報を確認し、ページを再読込してください。")

# 入力と検索
left, right = st.columns([1, 3], gap="medium")

with left:
    st.markdown("**検索する部品コード**（1行に1コード）")
    parts_text = st.text_area("部品コード", height=320,
                              placeholder="例）\nPART-X\nPART-Y",
                              label_visibility="collapsed")
    c1, c2 = st.columns(2)
    run_kishu = c1.button("① 機種コード逆展開", type="primary",
                          use_container_width=True, disabled=index is None)
    run_choku = c2.button("② 直上親コード", type="primary",
                          use_container_width=True, disabled=index is None)

parts = [s.strip() for s in parts_text.splitlines() if s.strip()] if parts_text else []

# 検索状態をセッションに保持
if "result" not in st.session_state:
    st.session_state.result = None  # (mode, rows, cols)

if (run_kishu or run_choku):
    if not parts:
        st.warning("検索する部品コードを入力してください（1行に1コード）。")
    elif index is not None:
        mode = "kishu" if run_kishu else "chokujou"
        with st.spinner("検索中...（%d件）" % len(parts)):
            if mode == "kishu":
                rows = core.search_kishu(parts, index)
                cols = core.KISHU_COLUMNS
            else:
                rows = core.search_chokujou(parts, index)
                cols = core.CHOKUJOU_COLUMNS
        st.session_state.result = (mode, rows, cols)

with right:
    res = st.session_state.result
    if res is None:
        st.caption("← 部品コードを入力して検索ボタンを押してください。")
    else:
        mode, rows, cols = res
        label = "機種コード逆展開検索" if mode == "kishu" else "直上親コード検索"
        st.markdown("**%s 結果：%d行**" % (label, len(rows)))
        if not rows:
            st.info("該当する結果がありませんでした。")
        else:
            table, keys, titles = rows_to_table(rows, cols)
            st.dataframe(table, use_container_width=True, hide_index=True, height=460)

            base = "逆検索結果_" + ("機種コード" if mode == "kishu" else "直上親コード")
            d1, d2, _ = st.columns([1, 1, 4])
            d1.download_button("CSV保存", data=to_csv_bytes(rows, keys, titles),
                               file_name=base + ".csv", mime="text/csv",
                               use_container_width=True)
            xlsx = to_xlsx_bytes(mode, rows, keys, titles)
            if xlsx is not None:
                d2.download_button(
                    "Excel保存", data=xlsx, file_name=base + ".xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True)

st.caption("DBへは SELECT のみ。マスタを最新化したい場合はページ右上メニューの「Rerun」"
           "またはアプリ再起動を行ってください。")
