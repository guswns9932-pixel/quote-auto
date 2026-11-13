"""견적서작성: 실제 통합양식으로 만든 출력물이 기준과 같은지 확인한다."""
import os

import pytest

from conftest import xlsx_fingerprint


def _item(cat, spec, qty, price, role="ITEM", amt=None):
    return {"cat": cat, "role": role, "spec": spec, "qty": qty, "price": price,
            "amt": qty * price if amt is None else amt}


@pytest.fixture()
def env(quote_template, tmp_path, monkeypatch):
    import core
    import excel_io
    monkeypatch.setattr(excel_io, "exe_dir", lambda: str(tmp_path))
    state = core.QuoteState(template_path=quote_template)
    # 양식에 실제로 있는 값을 읽어 쓴다 (테스트 코드에 단가를 적지 않기 위해)
    from openpyxl import load_workbook
    wb = load_workbook(quote_template, read_only=True, data_only=True)
    rows, by_class, price_by_spec, _ = excel_io.parse_items_sheet(wb["품목"])
    req = [c for c in next(wb["견적의뢰복사본"].iter_rows(min_row=2, max_row=2, values_only=True))]
    wb.close()
    qcode = req[5]
    pump = by_class[qcode][0]
    racks = [r for r in rows if r["A"] and not r["A"].upper().startswith("Q") and r["C"] > 0][:6]
    items = [_item("PUMP", pump["B"], 3, pump["C"])] + [_item("RACK", r["B"], 3, r["C"]) for r in racks]
    rd = {"row_idx": 2, "D": req[3], "E": req[4], "F": qcode, "G": req[6], "H": 3,
          "J": "테스트", "K": req[10], "Z": "T001"}
    return {"io": excel_io, "state": state, "items": items, "rd": rd}


def _credits():
    return [_item("RACK", "Pump Credit", 0, 0, role="CREDIT", amt=-1000000),
            _item("RACK", "Rack Credit", 0, 0, role="CREDIT", amt=-500000)]


def test_domestic(env, baseline):
    path = env["io"].generate_quote(env["state"], "국내", env["items"], env["rd"])
    baseline.check("quote_domestic", "quote", xlsx_fingerprint(path))


def test_domestic_with_credit(env, baseline):
    path = env["io"].generate_quote(env["state"], "국내", env["items"] + _credits(), env["rd"])
    baseline.check("quote_domestic_credit", "quote", xlsx_fingerprint(path))


def test_domestic_batch_matches_single(env, baseline):
    """일괄 생성(워크북 재사용)의 3번째 건도 단건 생성과 같은 내용이어야 한다."""
    rds = [dict(env["rd"], Z="T%03d" % i) for i in range(1, 4)]
    res = env["io"].generate_quote_multi(env["state"], env["items"], rds)
    assert [e for _, _, e in res] == ["", "", ""]
    baseline.check("quote_domestic_batch3", "quote", xlsx_fingerprint(res[2][1]))


def test_china(env, baseline):
    env["state"].cn_info = {"line": "X2", "tool": "T001", "maker": "MAKER", "process": "CVD"}
    path = env["io"].generate_quote(env["state"], "중국", env["items"] + _credits(), {})
    baseline.check("quote_china", "quote", xlsx_fingerprint(path))


def test_us(env, baseline):
    env["state"].us_info = {"site": "SITE", "tool": "T001", "maker": "MAKER", "process": "CVD",
                            "exchange": 1400, "base_ym": (2026, 11)}
    path = env["io"].generate_quote(env["state"], "미국", env["items"] + _credits(), {})
    baseline.check("quote_us", "quote", xlsx_fingerprint(path))


def test_cover_and_submit_sheet(env, baseline):
    """갑지 → 제출용 엑셀. (Excel 이 없는 환경이라 수식값은 openpyxl 폴백 경로)"""
    io = env["io"]
    rds = [dict(env["rd"], Z="T%03d" % i) for i in range(1, 3)]
    files = [p for _, p, _ in io.generate_quote_multi(env["state"], env["items"], rds)]
    cover, failed = io.generate_cover(env["state"].template_path, os.path.dirname(files[0]), files)
    assert failed == []
    fp = xlsx_fingerprint(cover)
    submit = io.export_cover_data_sheet(cover)
    baseline.check("quote_cover", "quote", fp)
    baseline.check("quote_submit_sheet", "quote", xlsx_fingerprint(submit))


def test_overflow_stops_without_leftover(env):
    """칸을 넘기면 오류로 멈추고 미완성 파일을 남기지 않는다."""
    io, st = env["io"], env["state"]
    st.cn_info = {"line": "X2", "tool": "T001", "maker": "M", "process": "P"}
    st.us_info = {"site": "S", "tool": "T001", "maker": "M", "process": "P", "exchange": 1400}
    pump = [_item("PUMP", "P%d" % i, 1, 100) for i in range(4)]
    rack = [_item("RACK", "R%d" % i, 1, 10) for i in range(26)]
    for qtype, items in (("중국", pump), ("중국", pump[:1] + rack[:21]),
                         ("미국", pump[:2]), ("미국", pump[:1] + rack[:21]),
                         ("국내", pump[:1] + rack)):
        with pytest.raises(ValueError, match="초과"):
            io.generate_quote(st, qtype, items, env["rd"])
    left = [f for _, _, fs in os.walk(io.exe_dir()) for f in fs]
    assert left == []


def test_template_checks(quote_template):
    import excel_io
    assert excel_io.inspect_template_losses(quote_template) == []
    assert excel_io.check_pump_price_consistency(quote_template) == []


def test_cover_skips_excel_for_current_quotes(env, tmp_path):
    """지금 버전이 만든 견적서는 값이 이미 들어 있어 갑지 생성 때 Excel 이 필요 없다.
    수식만 있고 계산값이 없는 예전 파일만 Excel 이 필요하다."""
    import shutil
    from openpyxl import load_workbook
    io = env["io"]
    path = io.generate_quote(env["state"], "국내", env["items"], env["rd"])
    assert io._needs_excel_values(path) is False

    old = str(tmp_path / "old.xlsx")
    shutil.copy(path, old)
    wb = load_workbook(old)
    wb["견적의뢰복사본"].cell(2, 19).value = "=사양서!F43"     # 예전 버전처럼 수식으로
    wb.save(old)
    assert io._needs_excel_values(old) is True


def test_us_and_china_header_fields(env):
    """미국·중국 견적서 머리 부분: 입력한 값은 들어가고, 비운 값은 양식에 적힌 그대로 남는다."""
    from openpyxl import load_workbook
    io, st = env["io"], env["state"]
    blank_us = {"site": "S", "tool": "T1", "maker": "M", "process": "P", "exchange": 1400, "base_ym": (2026, 11)}
    blank_cn = {"line": "L", "tool": "T1", "maker": "M", "process": "P"}

    # 아무것도 안 넣으면 양식 그대로 (투자자·보증기간은 기본값이 양식과 같다)
    st.us_info, st.cn_info = dict(blank_us), dict(blank_cn)
    us = load_workbook(io.generate_quote(st, "미국", env["items"], {}))["견적서_미국"]
    assert (us["B3"].value, us["G6"].value, us["F8"].value, us["F7"].value) == (
        "채승철 님 / 설비구매그룹", "입력필요", "PR : 입력필요", "Warranty : 2 Years after delivery")
    cn = load_workbook(io.generate_quote(st, "중국", env["items"], {}))["견적서_중국"]
    assert (cn["B3"].value, cn["G6"].value, cn["G10"].value, cn["G7"].value) == (
        "Xi Li / Purchasing Dept", "입력필요", "입력필요", "2 Year after delivery")

    # 넣은 값은 그 칸에 들어간다
    st.investor_name, st.warranty_years = "홍길동", 3
    st.us_info = dict(blank_us, quote_no="LOT-2611-01", pr="7300001234")
    st.cn_info = dict(blank_cn, quote_no="CN-01", delivery="2026-12-20", to="Wang / Purchasing")
    us = load_workbook(io.generate_quote(st, "미국", env["items"], {}))["견적서_미국"]
    assert (us["B3"].value, us["G6"].value, us["F8"].value, us["F7"].value) == (
        "홍길동 님 / 설비구매그룹", "LOT-2611-01", "PR : 7300001234", "Warranty : 3 Years after delivery")
    cn = load_workbook(io.generate_quote(st, "중국", env["items"], {}))["견적서_중국"]
    assert (cn["B3"].value, cn["G6"].value, cn["G10"].value, cn["G7"].value) == (
        "Wang / Purchasing", "CN-01", "2026-12-20", "3 Year after delivery")


def test_pump_only_hides_rack_rows_of_incoming_sheet(env):
    """PUMP 만(또는 PUMP + Credit) 견적하면 입고검수확인서의 RACK 점검 행(20~39)을 숨긴다."""
    from openpyxl import load_workbook
    io, st = env["io"], env["state"]
    pump = [i for i in env["items"] if i["cat"] == "PUMP"]

    def hidden(path):
        ws = load_workbook(path)["입고검수확인서"]
        return [r for r in range(14, 40) if ws.row_dimensions[r].hidden]

    assert hidden(io.generate_quote(st, "국내", pump, env["rd"])) == list(range(20, 40))
    assert hidden(io.generate_quote(st, "국내", pump + _credits(), env["rd"])) == list(range(20, 40))
    assert hidden(io.generate_quote(st, "국내", env["items"], env["rd"])) == []
    # 일괄 생성(워크북 재사용): PUMP 만 한 건 뒤에도 다음 건의 숨김 여부가 따로 정해진다
    res = io.generate_quote_multi(st, pump, [dict(env["rd"], Z="A1")])
    assert hidden(res[0][1]) == list(range(20, 40))
