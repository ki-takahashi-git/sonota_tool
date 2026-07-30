# -*- coding: utf-8 -*-
import core

def show(t, ok):
    print(("OK  " if ok else "NG  ") + t)
    assert ok, t

# 日付
d = core.parse_date("202607011")
show("parse 9桁", d["disp"] == "2026/07/01" and d["shift"] == "1")
d = core.parse_date("20260701")
show("parse 8桁→shift1", d["disp"] == "2026/07/01" and d["shift"] == "1")
show("無期限判定 999999991", core.is_infinite_date("999999991"))
show("無期限判定 99999999", core.is_infinite_date(99999999))
show("通常は無期限でない", not core.is_infinite_date("202607011"))
show("to9", core.to9("2026-07-01") == "202607011")
show("prev_day9 月跨ぎ", core.prev_day9("20260701") == "202606301")
show("prev_day9 年跨ぎ", core.prev_day9("20260101") == "202512311")

# 列解決
res = core.resolve_columns(["CODE","VENDOR","TID","TDATE","EDATE","TVOL","TEMA","PRICE","SPRICE","TAXRATE","TAXKUBU","HIMOKU","HOKAN"], core.CANDIDATES["tank"])
show("列解決 tank 全項目", res.get("PRICE")=="PRICE" and res.get("TID")=="TID")
res2 = core.resolve_columns(["CODE","HINMEI"], core.CANDIDATES["head"])
show("列解決 head 品名=HINMEI", res2.get("NAME")=="HINMEI")

# 行生成（価格改定）
cur = {"TID":"100123","CODE":"ABC-001","VENDOR":"V001","TDATE":"202501011",
       "EDATE":"999999991","TVOL":"100","TEMA":"0","PRICE":"50",
       "SPRICE":"50","TAXRATE":"10","TAXKUBU":"1","HIMOKU":"1"}
rows = core.make_rows(cur, "2026-07-01", "48", mode="update")
show("2行生成", len(rows)==2)
upd, new = rows
show("更新行 TID踏襲", upd["TID"]=="100123")
show("更新行 終了日=改定日前日", upd["EDATE"]=="202606301")
show("更新行 単価は現行のまま", upd["PRICE"]=="50")
show("新規行 TID空欄", new["TID"]=="")
show("新規行 開始日=改定日", new["TDATE"]=="202607011")
show("新規行 終了日=無期限踏襲", new["EDATE"]=="999999991")
show("新規行 改定単価", new["PRICE"]=="48")

# 新規のみ
rows2 = core.make_rows(cur, "2026-07-01", "48", new_tvol="200", mode="new")
show("新規のみ1行", len(rows2)==1 and rows2[0]["TID"]=="" and rows2[0]["TVOL"]=="200")

# 警告：MOQ不一致（LOTS=200 だが TVOL に200が無い）
grp = [dict(cur, TVOL="100"), dict(cur, TVOL="500", TID="100124")]
w = core.compute_warnings(grp, {"max":200,"min":200,"n":1})
show("MOQ不一致 検出", any("MOQ不一致" in x for x in w))
# 一致すれば出ない
w2 = core.compute_warnings([dict(cur, TVOL="200")], 200)
show("MOQ一致なら不一致警告なし", not any("MOQ不一致" in x for x in w2))
# 無期限2件
w3 = core.compute_warnings([dict(cur,EDATE="999999991",TVOL="200"), dict(cur,EDATE="999999991",TVOL="500")], 200)
show("無期限2件 検出", any("無期限" in x for x in w3))
# 0円
w4 = core.compute_warnings([dict(cur,PRICE="0",TVOL="200")], 200)
show("0円 検出", any("0円" in x for x in w4))

print("\nALL TESTS PASSED")
