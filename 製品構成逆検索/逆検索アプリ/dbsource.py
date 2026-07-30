# -*- coding: utf-8 -*-
"""
製品構成 逆検索ツール - マスタ取得層

TPiCS(SQL Server) から XITEM / XPRTS / XHEAD / XSECT を読取り専用で取得し、
core.BomIndex を構築する。DBへは SELECT のみ（書込みは一切行わない）。

接続は pymssql を優先し、無ければ pyodbc を試す（単価メンテアプリと同方式）。
demo=True のときは DB に接続せず、内蔵のサンプルBOMでインデックスを作る。
"""

import core


# =========================================================================
# DB接続（pymssql優先・pyodbcフォールバック）
# =========================================================================
def _get_conn(cfg):
    server = cfg.get("server", "YOUR_DB_SERVER_IP")
    port = str(cfg.get("port", "1433"))
    dbname = cfg.get("database", "TPICSDB_T")
    user = cfg.get("user", "YOUR_DB_USER")
    pwd = cfg.get("password", "YOUR_DB_PASSWORD")
    auth = str(cfg.get("auth", "sql")).lower()
    try:
        import pymssql
        kw = dict(server=server, port=port, database=dbname,
                  timeout=120, login_timeout=15)
        if auth == "windows":
            return pymssql.connect(**kw)
        return pymssql.connect(user=user, password=pwd, **kw)
    except ImportError:
        import pyodbc
        drv = cfg.get("odbc_driver", "ODBC Driver 17 for SQL Server")
        if auth == "windows":
            cs = (f"DRIVER={{{drv}}};SERVER={server},{port};"
                  f"DATABASE={dbname};Trusted_Connection=yes;")
        else:
            cs = (f"DRIVER={{{drv}}};SERVER={server},{port};"
                  f"DATABASE={dbname};UID={user};PWD={pwd};")
        return pyodbc.connect(cs, timeout=120)


def _fetch_table(cur, table):
    """SELECT * して {"header":[...], "rows":[[...],...]} を返す（VBA同様）。"""
    try:
        cur.execute("SET TRANSACTION ISOLATION LEVEL READ UNCOMMITTED")
    except Exception:
        pass
    cur.execute("SELECT * FROM [%s]" % table)
    header = [d[0] for d in cur.description]
    rows = [list(r) for r in cur.fetchall()]
    return {"header": header, "rows": rows}


def load_index_from_db(cfg, progress=None):
    """DBから4マスタを取得し core.BomIndex を返す。"""
    def _p(msg):
        if progress:
            progress(msg)
    conn = _get_conn(cfg)
    try:
        cur = conn.cursor()
        _p("XITEM を取得中...")
        xitem = _fetch_table(cur, "XITEM")
        _p("XPRTS を取得中...")
        xprts = _fetch_table(cur, "XPRTS")
        _p("XHEAD を取得中...")
        xhead = _fetch_table(cur, "XHEAD")
        _p("XSECT を取得中...")
        xsect = _fetch_table(cur, "XSECT")
    finally:
        try:
            conn.close()
        except Exception:
            pass
    _p("インデックスを構築中...")
    return core.BomIndex(xitem, xprts, xhead, xsect)


# =========================================================================
# デモ用サンプルBOM（DB未接続でも動作確認できる）
# =========================================================================
def _demo_masters():
    # XITEM: BUNR=1 が機種（最上位）
    xitem = {
        "header": ["CODE", "BUNR"],
        "rows": [
            ["MACHINE-A_01", 1], ["MACHINE-A_02", 1], ["MACHINE-B_01", 1],
            ["ASSY-1", 2], ["ASSY-2", 2], ["TD-TEMP1", 9],
            ["PART-X", 9], ["PART-Y", 9],
        ],
    }
    # XPRTS: CODE(親) が KCODE(子) を SIYOU/SIYOUW の比率で使用
    xprts = {
        "header": ["CODE", "KCODE", "SIYOU", "SIYOUW"],
        "rows": [
            ["MACHINE-A_01", "ASSY-1", 2, 1],
            ["MACHINE-A_02", "ASSY-1", 2, 1],
            ["MACHINE-B_01", "ASSY-2", 1, 1],
            ["ASSY-1", "TD-TEMP1", 3, 1],
            ["ASSY-1", "PART-Y", 1, 1],
            ["TD-TEMP1", "PART-X", 4, 1],
            ["ASSY-2", "PART-X", 5, 1],
        ],
    }
    # XHEAD: 名称・代表機種・客先品目・メーカー担当コード
    xhead = {
        "header": ["CODE", "NAME", "DKCODE", "CUCODE", "MAKER"],
        "rows": [
            ["MACHINE-A_01", "装置A 旧", "DK-A", "CU-A", "M01"],
            ["MACHINE-A_02", "装置A 新", "DK-A", "CU-A", "M01"],
            ["MACHINE-B_01", "装置B", "DK-B", "CU-B", "M02"],
            ["ASSY-1", "中間ユニット1", "DK-A", "", "M03"],
            ["ASSY-2", "中間ユニット2", "DK-B", "", "M03"],
            ["TD-TEMP1", "仮アセンブリ", "", "", "M03"],
            ["PART-X", "部品X 抵抗", "", "", "M04"],
            ["PART-Y", "部品Y コネクタ", "", "", "M05"],
        ],
    }
    # XSECT: 製造担当コード→メーカー名
    xsect = {
        "header": ["BUMO", "BNAME"],
        "rows": [
            ["M01", "高畑電子(自社)"], ["M02", "高畑電子(自社)"],
            ["M03", "山田製作所"], ["M04", "ローム(株)"],
            ["M05", "日本モレックス"],
        ],
    }
    return xitem, xprts, xhead, xsect


def load_index_demo(progress=None):
    if progress:
        progress("デモデータでインデックスを構築中...")
    return core.BomIndex(*_demo_masters())
