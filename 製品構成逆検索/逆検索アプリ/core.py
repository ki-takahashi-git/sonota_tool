# -*- coding: utf-8 -*-
"""
製品構成 逆検索ツール - コアロジック（DB非依存・GUI非依存）

Excel/VBA版 製品構成逆検索ツール Ver1.04 の検索ロジックを移植したもの。
  - Module2 RunBOMSearch          → search_kishu()    （機種コード逆展開検索）
  - Module3 RunDirectParentSearch → search_chokujou() （直上親コード検索）

TPiCS のマスタ構造:
  XITEM : CODE(品目コード), BUNR(品目分類。1=機種/最上位品目)
  XPRTS : CODE(親), KCODE(子=構成部品), SIYOU(使用数), SIYOUW(使用数の分母)
          → 1個あたり使用比率 ratio = SIYOU / SIYOUW
  XHEAD : CODE, NAME(品名), DKCODE(代表機種コード), CUCODE(客先品目コード),
          MAKER(製造担当コード), BNAME(メーカー名: XSECTから結合)
  XSECT : BUMO(製造担当コード), BNAME(名称)

本モジュールは pandas 等に依存せず、純粋な Python のみで動作する（単体テスト可能）。
"""

import sys

# 深いBOMでも再帰が止まらないように上限を引き上げる
sys.setrecursionlimit(100000)


# =========================================================================
# 列名の自動判定（VBA同様、大文字小文字を無視して列見出しから探す）
# =========================================================================
def _find_col(header, name):
    """header(列名リスト)から name に一致する列のインデックスを返す（無ければ -1）。"""
    target = str(name).strip().upper()
    for i, h in enumerate(header):
        if str(h).strip().upper() == target:
            return i
    return -1


def _to_float(v):
    """数値化。空白/非数値は 0.0。"""
    if v is None:
        return 0.0
    try:
        return float(str(v).replace(",", "").strip())
    except (TypeError, ValueError):
        return 0.0


def _s(v):
    return "" if v is None else str(v).strip()


# =========================================================================
# マスタ → 検索用インデックス構築
# =========================================================================
class BomIndex:
    """4つのマスタ（行リスト形式）から検索用の各種辞書を構築して保持する。

    各マスタは {"header": [列名...], "rows": [[値...], ...]} の形式で渡す。
    """

    def __init__(self, xitem, xprts, xhead, xsect):
        # --- XITEM: CODE → BUNR(数値) ---
        self.info_bunr = {}
        h = xitem["header"]
        ci = _find_col(h, "CODE")
        cb = _find_col(h, "BUNR")
        if ci >= 0 and cb >= 0:
            for r in xitem["rows"]:
                code = _s(r[ci])
                if code:
                    self.info_bunr[code] = _to_float(r[cb])

        # --- XSECT: BUMO → BNAME（メーカー名） ---
        sect = {}
        h = xsect["header"]
        cbu = _find_col(h, "BUMO")
        cbn = _find_col(h, "BNAME")
        if cbu >= 0 and cbn >= 0:
            for r in xsect["rows"]:
                k = _s(r[cbu])
                if k:
                    sect[k] = _s(r[cbn])

        # --- XHEAD: CODE → 各種属性 ---
        self.head_name = {}    # 品名
        self.head_dk = {}      # 代表機種コード
        self.head_cu = {}      # 客先品目コード
        self.head_bname = {}   # メーカー名（MAKER→XSECT.BNAME）
        h = xhead["header"]
        cc = _find_col(h, "CODE")
        cn = _find_col(h, "NAME")
        cdk = _find_col(h, "DKCODE")
        ccu = _find_col(h, "CUCODE")
        cmk = _find_col(h, "MAKER")
        cbn = _find_col(h, "BNAME")  # 既にBNAME列があればそれを優先
        if cc >= 0:
            for r in xhead["rows"]:
                code = _s(r[cc])
                if not code:
                    continue
                if cn >= 0:
                    self.head_name[code] = _s(r[cn])
                if cdk >= 0:
                    self.head_dk[code] = _s(r[cdk])
                if ccu >= 0:
                    self.head_cu[code] = _s(r[ccu])
                if cbn >= 0:
                    self.head_bname[code] = _s(r[cbn])
                elif cmk >= 0:
                    self.head_bname[code] = sect.get(_s(r[cmk]), "")

        # --- XPRTS: 逆引き・順引きインデックス ---
        # dict_rev : 子KCODE → [(親CODE, ratio), ...]   （上方向トレース用）
        # dict_fwd : 親CODE  → [(子KCODE, ratio), ...]   （下方向の使用数合計用）
        self.dict_rev = {}
        self.dict_fwd = {}
        h = xprts["header"]
        cp = _find_col(h, "CODE")
        ck = _find_col(h, "KCODE")
        csi = _find_col(h, "SIYOU")
        csw = _find_col(h, "SIYOUW")
        if cp >= 0 and ck >= 0:
            for r in xprts["rows"]:
                parent = _s(r[cp])
                child = _s(r[ck])
                if not parent or not child:
                    continue
                ratio = 0.0
                if csi >= 0 and csw >= 0:
                    w = _to_float(r[csw])
                    if w != 0:
                        ratio = _to_float(r[csi]) / w
                self.dict_rev.setdefault(child, []).append((parent, ratio))
                self.dict_fwd.setdefault(parent, []).append((child, ratio))

    def is_machine(self, code):
        """BUNR==1 なら機種（最上位品目）。"""
        return self.info_bunr.get(code, 0.0) == 1.0


# =========================================================================
# 共通：リビジョン抽出（コード末尾 "_xx" を除いた基底名）
# =========================================================================
def _base_name(code):
    pos = code.rfind("_")
    return code[:pos] if pos > 0 else code


# =========================================================================
# 機種コード 逆展開検索（VBA Module2 RunBOMSearch 相当）
# =========================================================================
def search_kishu(parts, idx):
    """部品コードのリストから、それを使用する最上位の機種コードを逆引きする。

    返り値: 行(dict)のリスト。各行のキー:
      part   : 検索した部品
      model  : 機種コード（該当なしは "該当なし"）
      usage  : 使用数（機種1台あたりの合計）
      cucode : 客先品目コード
      order  : 入力順(ソート用)
    """
    results = []
    order = 0
    for raw in parts:
        tcode = _s(raw)
        if not tcode:
            continue
        order += 1
        found = {}
        _find_first_models(tcode, idx, found)
        latest = _filter_latest_rev_full(found)
        if not latest:
            results.append({"part": tcode, "model": "該当なし",
                            "usage": "", "cucode": "", "order": order})
        else:
            for model in latest:
                usage = _calc_total_usage(model, tcode, idx, 1.0)
                results.append({
                    "part": tcode,
                    "model": model,
                    "usage": usage,
                    "cucode": idx.head_cu.get(model, ""),
                    "order": order,
                })
    # ソート: 入力順 → 機種コード
    results.sort(key=lambda x: (x["order"], x["model"]))
    return results


def _find_first_models(curr, idx, found, _path=None):
    """curr の上位をたどり、最初に出会う機種(BUNR=1)を found に集める。"""
    if _path is None:
        _path = set()
    if curr in _path:          # 循環防止（VBA版には無いが安全対策）
        return
    _path = _path | {curr}
    for parent, _ratio in idx.dict_rev.get(curr, []):
        if idx.is_machine(parent):
            found[parent] = True
        else:
            _find_first_models(parent, idx, found, _path)


def _filter_latest_rev_full(found):
    """基底名ごとに最新（文字列比較で最大）のコードだけを残す。"""
    best = {}
    for code in found.keys():
        b = _base_name(code)
        if b not in best or code > best[b]:
            best[b] = code
    return list(best.values())


def _calc_total_usage(curr, target, idx, mult, _path=None):
    """curr から下方向に展開し、target に到達する経路の使用数を合計する。"""
    if _path is None:
        _path = set()
    if curr in _path:
        return 0.0
    _path = _path | {curr}
    sub = 0.0
    for child, ratio in idx.dict_fwd.get(curr, []):
        nxt = mult * ratio
        if child == target:
            sub += nxt
        sub += _calc_total_usage(child, target, idx, nxt, _path)
    return sub


# =========================================================================
# 直上親コード検索（VBA Module3 RunDirectParentSearch 相当）
# =========================================================================
def search_chokujou(parts, idx):
    """部品の「直上親」（TD-で始まる仮親はスキップした最初の実親）を検索する。

    返り値: 行(dict)のリスト。各行のキー:
      part        : 検索した部品
      model       : 機種コード
      usage_total : 使用数(機種1台あたり合計)
      parent      : 直上親コード
      parent_name : 直上親名称
      usage_direct: 使用数(直上親1個あたり)
      part_name   : 部品名称
      parent_dk   : 代表機種コード(直上親)
      dup         : 重複フラグ "*" or ""
      maker       : メーカー名(機種)
      order       : 入力順(ソート用)
    """
    raw_results = []   # (order, part, model, total_usage, direct_parent, direct_usage)
    order = 0
    for raw in parts:
        tcode = _s(raw)
        if not tcode:
            continue
        order += 1
        before = len(raw_results)
        _trace_up_to_machine(tcode, tcode, 1.0, "", 1.0, idx, raw_results, order)
        if len(raw_results) == before:
            raw_results.append((order, tcode, "該当なし", 0.0, "該当なし", 0.0))

    # --- 集計1: 基底名ごとに最新リビジョンの機種だけ採用（末尾5文字で比較）---
    map_latest = {}
    for (o, part, model, tot, dp, du) in raw_results:
        if model == "該当なし":
            continue
        base = _base_name(model)
        key = part + "|" + base
        if key not in map_latest:
            map_latest[key] = model
        else:
            if model[-5:] > map_latest[key][-5:]:
                map_latest[key] = model

    # --- 集計2: 行の重複排除と使用数合計 ---
    map_total = {}     # part|model → 合計使用数
    map_rows = {}      # 一意キー → (order, part, model, direct_parent, direct_usage)
    dup_check = {}     # part||direct_parent → 件数
    for (o, part, model, tot, dp, du) in raw_results:
        if model == "該当なし":
            map_rows[(o, "NONE")] = (o, part, "該当なし", "該当なし", 0.0)
            continue
        base = _base_name(model)
        if model != map_latest.get(part + "|" + base):
            continue
        key_full = part + "|" + model
        key_row = (o, key_full, dp)
        map_total[key_full] = map_total.get(key_full, 0.0) + tot
        if key_row not in map_rows:
            map_rows[key_row] = (o, part, model, dp, du)
            dk = part + "||" + dp
            dup_check[dk] = dup_check.get(dk, 0) + 1

    # --- 出力行の組み立て ---
    out = []
    for (o, part, model, parent, direct_usage) in map_rows.values():
        row = {
            "part": part, "model": model, "usage_total": "",
            "parent": "", "parent_name": "", "usage_direct": "",
            "part_name": idx.head_name.get(part, ""),
            "parent_dk": "", "dup": "", "maker": "", "order": o,
        }
        if model != "該当なし":
            row["usage_total"] = map_total.get(part + "|" + model, 0.0)
            row["parent"] = parent
            row["parent_name"] = idx.head_name.get(parent, "")
            row["usage_direct"] = direct_usage
            if dup_check.get(part + "||" + parent, 0) > 1:
                row["dup"] = "*"
            if parent != "該当なし":
                row["parent_dk"] = idx.head_dk.get(parent, "")
            row["maker"] = idx.head_bname.get(model, "")
        out.append(row)

    # ソート: 入力順 → 機種コード → 直上親コード
    out.sort(key=lambda x: (x["order"], x["model"], x["parent"]))
    return out


def _trace_up_to_machine(cur, tar, mult, dp, dacc, idx, results, order, _path=None):
    """cur から上位をたどり機種に到達したら結果を記録する（直上親=最初の非TD-親）。

    dp   : 確定済みの直上親コード（""=未確定, "SKIP"=TD-親をスキップ中）
    dacc : 直上親までの使用数(累積比率)
    mult : 機種までの使用数(累積比率)
    """
    if _path is None:
        _path = set()
    if cur in _path:
        return
    _path = _path | {cur}
    for p, q in idx.dict_rev.get(cur, []):
        if dp == "":
            n_acc = q
            n_p = "SKIP" if p[:3] == "TD-" else p
        elif dp == "SKIP":
            n_acc = dacc * q
            n_p = p if p[:3] != "TD-" else "SKIP"
        else:
            n_acc = dacc
            n_p = dp
        f_p = p if n_p == "SKIP" else n_p

        if idx.is_machine(p):
            results.append((order, tar, p, mult * q, f_p, n_acc))
        else:
            _trace_up_to_machine(p, tar, mult * q, n_p, n_acc, idx, results, order, _path)


# =========================================================================
# 出力列定義（Excel版の見出しに合わせる）
# =========================================================================
KISHU_COLUMNS = [
    ("part", "検索した部品"),
    ("model", "機種コード"),
    ("usage", "使用数"),
    ("cucode", "客先品目コード"),
]

CHOKUJOU_COLUMNS = [
    ("part", "検索した部品"),
    ("model", "機種コード"),
    ("usage_total", "使用数(合計)"),
    ("parent", "直上親コード"),
    ("parent_name", "直上親名称"),
    ("usage_direct", "使用数(直上)"),
    ("part_name", "部品名称"),
    ("parent_dk", "代表機種コード(直上)"),
    ("dup", "重複"),
    ("maker", "メーカー名"),
]
