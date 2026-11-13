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
