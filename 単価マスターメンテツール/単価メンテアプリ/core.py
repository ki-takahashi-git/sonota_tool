# -*- coding: utf-8 -*-
"""
単価マスター メンテナンスツール - 中核ロジック（DB非依存・純粋関数）
TPiCS 単価マスター(XTANK) / アイテムマスター(XHEAD/XITEM) を対象。
ここには「日付変換」「更新行/新規行の生成」「警告判定」など、
DB接続なしで単体テスト可能な処理だけを置く。
"""
from datetime import date, timedelta

# ---- 論理項目名 -> 実テーブルの候補列名（スキーマ自動検出に使用）----
# 実DBの列名が候補に無い場合は config.ini で上書きできる。
CANDIDATES = {
    "tank": {
        "CODE":    ["CODE"],
        "VENDOR":  ["VENDOR"],
        "TID":     ["TID"],
        "TDATE":   ["TDATE"],
        "EDATE":   ["EDATE"],
        "TVOL":    ["TVOL"],
        "TEMA":    ["TEMA"],
        "SPRICE":  ["SPRICE"],
        "PRICE":   ["PRICE"],
        "TAXRATE": ["TAXRATE"],
        "TAXKUBU": ["TAXKUBU"],
        "HIMOKU":  ["HIMOKU"],
        # 注文コード（メーカー品番が入っている場合がある）
        "VCODE":   ["VCODE"],
    },
    # アイテム基本マスター：品名など
    "head": {
        "CODE": ["CODE"],
        "NAME": ["NAME", "HINMEI", "HNAME", "ITEMNAME", "PNAME", "ZUNAME", "ZUMEN"],
        # メーカー品目コード／メーカー型式（環境で列名が異なる）
        "MPN": ["MAKERNO", "MAKERCD", "MAKERCODE", "MCODE", "MTYPE", "MKATA",
                "KATA", "KATASHIKI", "MAKERKATA", "MPCODE", "MPN", "MAKER",
                "MITSUCD", "VITEMCD", "VITEMCODE", "SCODE"],
    },
    # 製造担当マスター(XSECT)：発注先名称など
    "sect": {
        "CODE": ["BUMO", "SECT", "SECTCD", "SECTCODE", "CODE", "VENDOR"],
        "NAME": ["BNAME", "NAME", "SECTNAME", "BUMONAME", "BUMNAME", "SNAME", "HNAME", "VENDORNAME"],
    },
    # アイテム詳細マスター：ロットサイズ(MOQ)など
    "item": {
        "CODE": ["CODE"],
        "BUMO": ["BUMO"],
        "LOTS": ["LOTS"],
        "MPN": ["MAKERNO", "MAKERCD", "MAKERCODE", "MCODE", "MTYPE", "MKATA",
                "KATA", "KATASHIKI", "MAKERKATA", "MPCODE", "MPN", "MAKER"],
    },
}


def resolve_columns(available_cols, candidate_map):
    """実在する列名(available_cols)から論理名→実列名の辞書を作る。"""
    avail_upper = {c.upper(): c for c in available_cols}
    resolved = {}
    for logical, cands in candidate_map.items():
        for cand in cands:
            if cand.upper() in avail_upper:
                resolved[logical] = avail_upper[cand.upper()]
                break
    return resolved


# ----------------- 日付（YYYYMMDDS / 9桁）-----------------
def _digits(raw):
    if raw is None:
        return ""
    s = str(raw).strip()
    # 小数点付き等を除去
    if s.endswith(".0"):
        s = s[:-2]
    return "".join(ch for ch in s if ch.isdigit())


def is_infinite_date(raw):
    """無期限終了日か判定（99/99/99 ＝ 99999999... 系）。"""
    d = _digits(raw)
    if not d:
        return False
    # 8桁以上の先頭8桁、または短い場合はそのまま
    head8 = d[:8] if len(d) >= 8 else d
    return head8.startswith("9999") or head8 == "99999999" or d.startswith("9999")


def parse_date(raw):
    """生値 -> {'raw','disp','y','m','d','shift','infinite'} を返す。"""
    d = _digits(raw)
    res = {"raw": str(raw).strip() if raw is not None else "",
           "disp": "", "y": None, "m": None, "d": None, "shift": "1", "infinite": False}
    if not d:
        return res
    if is_infinite_date(raw):
        res["infinite"] = True
        res["disp"] = "無期限(99/99/99)"
        return res
    # 9桁: YYYYMMDDS / 8桁: YYYYMMDD
    if len(d) >= 9:
        ymd, shift = d[:8], d[8]
    elif len(d) == 8:
        ymd, shift = d, "1"
    else:
        res["disp"] = d
        return res
    try:
        y, m, dd = int(ymd[:4]), int(ymd[4:6]), int(ymd[6:8])
        res.update(y=y, m=m, d=dd, shift=shift or "1",
                   disp="%04d/%02d/%02d" % (y, m, dd))
    except ValueError:
        res["disp"] = d
    return res


def to9(yyyymmdd, shift="1"):
    """'2026-07-01' or '20260701' or date -> '202607011'（9桁）。"""
    if isinstance(yyyymmdd, date):
        ymd = "%04d%02d%02d" % (yyyymmdd.year, yyyymmdd.month, yyyymmdd.day)
    else:
        ymd = "".join(ch for ch in str(yyyymmdd) if ch.isdigit())[:8]
    if len(ymd) != 8:
        raise ValueError("日付は8桁(YYYYMMDD)で指定してください: %r" % yyyymmdd)
    return ymd + (str(shift)[:1] if shift else "1")


def prev_day9(yyyymmdd, shift="1"):
    """指定日の前日を9桁で返す（旧単価の適用終了日に使用）。"""
    ymd = "".join(ch for ch in str(yyyymmdd) if ch.isdigit())[:8]
    d = date(int(ymd[:4]), int(ymd[4:6]), int(ymd[6:8])) - timedelta(days=1)
    return to9(d, shift)


# ----------------- 行生成（更新行＋新規行）-----------------
OUTPUT_COLUMNS = ["TID", "CODE", "VENDOR", "TDATE", "EDATE", "TVOL",
                  "TEMA", "PRICE", "SPRICE", "TAXRATE", "TAXKUBU", "HIMOKU"]


def make_rows(current, kaitei_date, new_price,
              new_tvol=None, new_sprice=None, mode="update"):
    """
    current: 現行単価マスター1行 dict（論理名キー: TID,CODE,VENDOR,TDATE,EDATE,TVOL...）
    kaitei_date: 改定日 'YYYY-MM-DD' or 'YYYYMMDD'
    new_price: 改定後の発注単価
    返り値: 出力行 list（更新行＋新規行 / 新規のみ）
    mode: 'update'=価格改定（更新行＋新規行） / 'new'=新規追加のみ
    """
    shift = str(current.get("TDATE_shift") or "1")[:1] or "1"
    rows = []

    if mode == "update":
        # ①更新行：現行TIDで上書き。終了日を改定日の前日へ＝現行単価を完了
        upd = {c: current.get(c, "") for c in OUTPUT_COLUMNS}
        upd["TID"] = current.get("TID", "")
        upd["EDATE"] = prev_day9(kaitei_date, shift)
        upd["_kind"] = "更新(終了日変更)"
        rows.append(upd)

    # ②新規行：TID空欄で新規追加。開始日=改定日、終了日=無期限、単価=改定単価
    new = {c: current.get(c, "") for c in OUTPUT_COLUMNS}
    new["TID"] = ""  # 空欄＝新規追加
    new["TDATE"] = to9(kaitei_date, shift)
    # 無期限終了日：現行行が無期限ならその値を踏襲、無ければ既定
    cur_edate = current.get("EDATE", "")
    new["EDATE"] = str(cur_edate).strip() if is_infinite_date(cur_edate) else "999999991"
    new["PRICE"] = new_price
    if new_tvol not in (None, ""):
        new["TVOL"] = new_tvol
    if new_sprice not in (None, ""):
        new["SPRICE"] = new_sprice
    new["_kind"] = "新規追加"
    rows.append(new)
    return rows


# ----------------- 警告判定 -----------------
def compute_warnings(rows, lots):
    """
    rows: 同一(アイテムコード＋発注先)の現行単価マスター行 list（各row: dict, 論理名キー）
    lots: アイテムマスターのロットサイズ(MOQ)。複数候補ある場合は dict {'min','max','n'} も可
    返り値: 警告コード list（日本語）
    """
    warns = []
    # lots 正規化
    lots_val = None
    lots_multi = False
    if isinstance(lots, dict):
        lots_val = lots.get("max")
        lots_multi = (lots.get("n") or 1) > 1
    else:
        lots_val = lots

    def _num(x):
        try:
            return float(str(x).replace(",", ""))
        except (TypeError, ValueError):
            return None

    tvols = [_num(r.get("TVOL")) for r in rows]
    tvols = [t for t in tvols if t is not None]
    lv = _num(lots_val)

    # MOQ不一致：LOTSに一致する適用数量(TVOL)の現行行が無い
    if lv is not None and tvols and lv not in tvols:
        warns.append("MOQ不一致：ロットサイズ(%s)に一致する適用数量の単価行が無い→0円明細/旧単価のリスク" % _fmt(lots_val))
    if lv is None or (tvols == []):
        warns.append("MOQまたは適用数量が未取得：要確認")
    if lots_multi:
        warns.append("ロットサイズが製造担当により複数：要確認")

    # 無期限終了日（適用終了日=99/99/99）の残存件数
    inf = [r for r in rows if is_infinite_date(r.get("EDATE"))]
    if len(inf) >= 2:
        warns.append("無期限(99/99/99)の単価行が%d件：改定時は全件を終了処理" % len(inf))

    # 0円単価
    zero = [r for r in rows if _num(r.get("PRICE")) == 0]
    if zero:
        warns.append("発注単価0円の行が%d件（終息品の自動作成など）" % len(zero))

    return warns


def _fmt(x):
    try:
        f = float(str(x).replace(",", ""))
        return ("%g" % f)
    except (TypeError, ValueError):
        return str(x)


def clean_num(v):
    """数値文字列の不要な末尾ゼロ・小数点を除去（5000.000->5000, 0.20->0.2, 0.000->0）。
    数値に変換できない場合や空はそのまま返す。日付・区分コードには使わないこと。"""
    if v is None:
        return ""
    sv = str(v).strip()
    if sv == "":
        return ""
    try:
        f = float(sv.replace(",", ""))
    except ValueError:
        return sv
    if f == int(f):
        return str(int(f))
    return ("%f" % f).rstrip("0").rstrip(".")
