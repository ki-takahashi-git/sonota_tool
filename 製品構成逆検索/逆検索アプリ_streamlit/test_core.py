# -*- coding: utf-8 -*-
"""
core.py の単体テスト（DB不要）。
  python test_core.py
サンプルBOM:
  MACHINE-A_01/_02 → ASSY-1 → {TD-TEMP1 → PART-X, PART-Y}
  MACHINE-B_01     → ASSY-2 → PART-X
比率: A→ASSY-1=2, ASSY-1→TD-TEMP1=3, TD-TEMP1→PART-X=4, ASSY-1→PART-Y=1
      B→ASSY-2=1, ASSY-2→PART-X=5
"""
import core
import dbsource


def build():
    return core.BomIndex(*dbsource._demo_masters())


def test_kishu_partx():
    idx = build()
    rows = core.search_kishu(["PART-X"], idx)
    models = {r["model"]: r["usage"] for r in rows}
    # 最新リビジョン _02 のみ採用、_01 は出ない
    assert "MACHINE-A_01" not in models, models
    assert "MACHINE-A_02" in models, models
    assert "MACHINE-B_01" in models, models
    # 機種1台あたり使用数: A = 2*3*4 = 24, B = 1*5 = 5
    assert models["MACHINE-A_02"] == 24, models
    assert models["MACHINE-B_01"] == 5, models
    # 客先品目コード
    cu = {r["model"]: r["cucode"] for r in rows}
    assert cu["MACHINE-A_02"] == "CU-A", cu
    print("test_kishu_partx OK", models)


def test_kishu_none():
    idx = build()
    rows = core.search_kishu(["NOTHING"], idx)
    assert len(rows) == 1 and rows[0]["model"] == "該当なし", rows
    print("test_kishu_none OK")


def test_chokujou_partx():
    idx = build()
    rows = core.search_chokujou(["PART-X"], idx)
    by_model = {r["model"]: r for r in rows}
    assert "MACHINE-A_01" not in by_model, by_model
    # A経路: TD-TEMP1 はスキップされ直上親=ASSY-1, 直上使用数=4*3=12, 合計=24
    a = by_model["MACHINE-A_02"]
    assert a["parent"] == "ASSY-1", a
    assert a["usage_direct"] == 12, a
    assert a["usage_total"] == 24, a
    assert a["maker"] == "高畑電子(自社)", a
    assert a["parent_name"] == "中間ユニット1", a
    # B経路: 直上親=ASSY-2, 直上=5, 合計=5
    b = by_model["MACHINE-B_01"]
    assert b["parent"] == "ASSY-2", b
    assert b["usage_direct"] == 5, b
    assert b["usage_total"] == 5, b
    print("test_chokujou_partx OK")


def test_chokujou_party():
    idx = build()
    rows = core.search_chokujou(["PART-Y"], idx)
    # PART-Y は ASSY-1 直下（TD-無し）→ MACHINE-A のみ
    by_model = {r["model"]: r for r in rows}
    assert "MACHINE-A_02" in by_model, by_model
    a = by_model["MACHINE-A_02"]
    assert a["parent"] == "ASSY-1", a
    assert a["usage_direct"] == 1, a
    assert a["usage_total"] == 2, a   # MACHINE-A→ASSY-1=2, ASSY-1→PART-Y=1
    print("test_chokujou_party OK")


def test_order_preserved():
    idx = build()
    rows = core.search_kishu(["PART-Y", "PART-X"], idx)
    orders = [r["order"] for r in rows]
    assert orders == sorted(orders), orders
    assert rows[0]["part"] == "PART-Y", rows[0]
    print("test_order_preserved OK")


if __name__ == "__main__":
    test_kishu_partx()
    test_kishu_none()
    test_chokujou_partx()
    test_chokujou_party()
    test_order_preserved()
    print("\n全テスト成功 ✓")
