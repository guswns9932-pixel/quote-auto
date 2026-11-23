"""생성 전 점검의 수량 규칙 — Rack 대수, 의뢰 수량, Rack Concept, 5D, CH당 품목, 같은 종류 중복."""
import pytest
from openpyxl import Workbook, load_workbook

import quote_review as rv
import review_rules as rr


def _it(cat, spec, qty, price=10, role="ITEM"):
    return {"cat": cat, "role": role, "spec": spec, "qty": qty, "price": price, "amt": qty * price}


def _texts(review, level=None):
    return [c.text for c in review.checks if level is None or c.level == level]


def test_rack_ch_by_rule():
    r = rr.defaults()
    assert r.rack_ch("Rack System-3") == (True, 3)
    assert r.rack_ch("Rack System-6 (전기공통 사양)") == (True, 6)
    assert r.rack_ch("Rack (격자보 Rack-1)") == (True, 1)
    assert r.rack_ch("Rack ((Support -1 (전기공통 사양))") == (True, 1)
    for unknown in ("Rack (3단 Booster용)", "Compressor Rack(1단)", "Rack Compressor Rack(2단)"):
        assert r.rack_ch(unknown) == (True, None), unknown
    # Rack 분류에 있어도 Rack 이 아닌 품목은 세지 않는다 — '-1' 이 붙어 있어도
    for not_rack in ("Rack cover", "RACK 표준화_1CH 당", "Heat jacket Controller 지지대-1CH 당", "Rack 집합배관비용"):
        assert r.rack_ch(not_rack) == (False, None), not_rack


@pytest.mark.parametrize("text, want", [
    ("3CH", 3), ("3CH+2CH", 5), ("3 ch", 3), ("3단", None), ("", None), (None, None)])
def test_concept(text, want):
    assert rr.ch_from_concept(text) == want


@pytest.mark.parametrize("name, want", [
    ("PUMP_DRY_LOT_XD1200_3RACK_CLN_R01", 3),
    ("Pump_DRY_LOT_HD2400eH_3Rack_CVD_R01", 3),
    ("PUMP_GAS VENTING_1Rack_LOT_HD600_VER.000", 1),
    ("PUMP_DRY_LOT_XD1200_3RACK_CLN_R01+PUMP_DRY_LOT_XD1200_3RACK_CLN_R01", 6),
    ("PUMP_DRY_LOT_HD2400_2RACK-A3_CVD_R01+PUMP_DRY_LOT_HD2400_1RACK_CVD_R01", 3),
    ("PUMP_DRY_LOT_HD2000EH_3RACK_CVD-TES_R01*2", 6),
    ("HD6001RACK_EXHUPCONCEPT", None),                   # 숫자가 붙어 있어 뜻이 모호 — 읽지 않는다
    ("PUMP_DRY_LOT_XD1200_CLN", None), ("", None)])
def test_5d(name, want):
    assert rr.ch_from_5d(name) == want


def test_counts_all_match_no_alarm():
    items = [_it("PUMP", "XD1200", 3), _it("RACK", "Rack System-3", 1), _it("RACK", "Cable Tray Ass'y & Aaptor", 3),
             _it("RACK", "Rack cover", 1), _it("RACK", "RACK 표준화_1CH 당", 3)]
    r = rv.analyze(items, "국내", 3, rack_concept="3CH", d5="PUMP_DRY_LOT_XD1200_3RACK_CLN_R01")
    assert _texts(r) == []


def test_rack_total_and_request_quantity():
    items = [_it("PUMP", "XD1200", 4), _it("RACK", "Rack System-2", 1), _it("RACK", "Rack (Support -1)", 1)]
    r = rv.analyze(items, "국내", 3, rack_concept="3CH", d5="PUMP_DRY_LOT_XD1200_2RACK_R01")
    red = _texts(r, rv.RED)
    assert "PUMP 수량 4대 ≠ 의뢰 수량 3" in red
    assert any(t.startswith("Rack 합계 3CH ≠ PUMP 4대") for t in red)
    yellow = _texts(r, rv.YELLOW)
    assert not any("Rack Concept" in t for t in yellow)               # 3 = 3CH 이므로 맞다
    assert "Rack 합계 3CH 가 5D 의 2CH 의 배수가 아닙니다" in yellow


def test_5d_multiples_are_fine_and_concept_mismatch():
    # 5D 3RACK 구성을 2세트(6CH) 담은 경우 — 배수이므로 정상
    items = [_it("PUMP", "P", 6), _it("RACK", "Rack System-3", 2)]
    r = rv.analyze(items, "국내", 6, rack_concept="6CH", d5="PUMP_DRY_LOT_HD2000EH_3RACK_CVD-TES_R01")
    assert _texts(r) == []
    r2 = rv.analyze(items, "국내", 6, rack_concept="3CH")
    assert _texts(r2, rv.YELLOW) == ["Rack 합계 6CH ≠ 의뢰의 Rack Concept 3CH"]


def test_unknown_rack_is_left_out_and_flagged():
    items = [_it("PUMP", "P", 3), _it("RACK", "Rack System-3", 1), _it("RACK", "Rack (3단 Booster용)", 1)]
    r = rv.analyze(items, "국내", 3)
    assert _texts(r, rv.RED) == []                                    # 3단 Rack 은 계산에서 빠져 3 = 3
    assert any(t.startswith("확인 필요: Rack (3단 Booster용)") for t in _texts(r, rv.YELLOW))
    only_unknown = rv.analyze([_it("PUMP", "P", 3), _it("RACK", "Compressor Rack(1단)", 1)], "국내", 3)
    assert not any(t.startswith("Rack 합계") for t in _texts(only_unknown))   # 셀 Rack 이 없으면 합계 비교도 안 한다


def test_same_ch_rack_twice_but_combinations_are_fine():
    combo = rv.analyze([_it("PUMP", "P", 9), _it("RACK", "Rack System-3", 1), _it("RACK", "Rack System-6", 1)], "국내", 9)
    assert _texts(combo) == []                                        # 3 + 6 = 9 — 흔한 정상 조합
    dup = rv.analyze([_it("PUMP", "P", 6), _it("RACK", "Rack System-3", 1),
                      _it("RACK", "Rack System-3 (전기공통 사양)", 1)], "국내", 6)
    assert _texts(dup, rv.YELLOW) == [
        "같은 3CH Rack 이 두 가지 담겼습니다: Rack System-3, Rack System-3 (전기공통 사양)"]


def test_per_ch_rules_and_same_type_twice():
    cls = {"Hot N2 Transfer HT-40": "Hot N2 Transfer", "Hot N2 Transfer HT-50": "Hot N2 Transfer"}
    items = [_it("PUMP", "P", 3), _it("RACK", "Rack System-3", 1), _it("RACK", "Cable Tray Ass'y & Aaptor", 2),
             _it("RACK", "Dual seal O-ring  (NW40)", 6),
             _it("RACK", "Hot N2 Transfer HT-40", 2), _it("RACK", "Hot N2 Transfer HT-50", 1)]
    r = rv.analyze(items, "국내", 3, class_of=cls)
    y = _texts(r, rv.YELLOW)
    assert "Cable Tray Ass'y & Aaptor 수량 2 ≠ 3 (CH 마다 1개 × 3CH)" in y
    assert "Dual seal O-ring  (NW40) 수량 6 ≠ 3 (CH 마다 1개 × 3CH)" in y
    assert "Hot N2 Transfer 이 두 가지 이상 담겼습니다: Hot N2 Transfer HT-40, Hot N2 Transfer HT-50" in y


def test_rules_sheet_replaces_defaults_and_reports_bad_rows(tmp_path):
    rows = [("구분", "품목 이름에 포함", "값", "등급"),
            ("Rack", "Rack System-{n}", "n", None),
            ("Rack", "Rack (3단 Booster용)", 2, None),                 # 사람이 CH 를 정해 주면 그대로 센다
            ("CH당", "Fitting & Hose", 3, "판단 필요"),
            ("Rack", "이름에 n 없음", "n", None),
            ("CH당", "Clamp", "많이", None),
            ("모름", "X", 1, None),
            ("# 메모", "", "", None)]
    rules = rr.read_rows(rows)
    assert rules.rack_ch("Rack (3단 Booster용)") == (True, 2)
    assert rules.rack_ch("Rack (Support -1)") == (False, None)          # 시트에 없으면 적용 안 함
    assert rules.per_ch_rule("Fitting & Hose").level == "red"
    assert len(rules.problems) == 3
    r = rv.analyze([_it("PUMP", "P", 3), _it("RACK", "Rack System-1", 1), _it("RACK", "Rack (3단 Booster용)", 1),
                    _it("RACK", "Fitting & Hose", 9)], "국내", 3, rules=rules)
    assert _texts(r) == []                                             # 1 + 2 = 3, Fitting 9 = 3 × 3

    # 설정 시트 만들기: 검토규칙 시트도 같이 나오고, 그대로 읽으면 기본 규칙과 같다
    import quote_config as qc
    out = qc.export_settings_sheet(str(tmp_path / "s.xlsx"))
    wb = load_workbook(out)
    assert wb.sheetnames == [qc.SETTINGS_SHEET, rr.RULES_SHEET]
    back = rr.from_workbook(wb)
    assert back.racks == rr.defaults().racks and back.per_ch == rr.defaults().per_ch and not back.problems
    assert rr.from_workbook(Workbook()).racks == rr.defaults().racks   # 시트가 없으면 기본 규칙
