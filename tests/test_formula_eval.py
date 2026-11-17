"""수식 계산 — Excel 없이 파일에 적힌 수식을 따라가 값을 구한다."""
import pytest
from openpyxl import Workbook, load_workbook

import formula_eval as fe


def _ws(cells):
    wb = Workbook(); ws = wb.active
    for ref, v in cells.items():
        ws[ref] = v
    return ws


def test_arithmetic_refs_and_functions():
    ws = _ws({"A1": 10, "A2": 3, "A3": "=A1*A2", "A4": "=SUM(A1:A3)", "A5": "=(A4-A1)/A2+2^3", "A6": "=-A1+50%*A2",
              "B1": "=ROUNDDOWN(12345.678,-3)", "B2": "=ROUNDUP(12001,-3)", "B3": "=ROUND(2.5,0)", "B4": "=ROUND(-2.5,0)",
              "B5": "=ROUNDDOWN($A$1/3,2)", "B6": "=MAX(A1:A2)-MIN(A1,A2)+ABS(-1)+INT(7.9)", "C1": "글자", "C2": None,
              "C3": "=SUM(C1:C2,A1)", "C4": "=A1+C2"})
    want = {"A3": 30, "A4": 43, "A5": 19, "A6": -8.5, "B1": 12000, "B2": 13000, "B3": 3, "B4": -3,
            "B5": 3.33, "B6": 15, "C3": 10, "C4": 10}
    for ref, v in want.items():
        assert fe.cell_value(ws, ref) == pytest.approx(v), ref
    assert fe.cell_value(ws, "A1") == 10 and fe.cell_value(ws, "C1") == "글자"       # 수식이 아니면 그대로


def test_rounddown_is_exact_on_round_numbers():
    """부동소수점 꼬리 때문에 천 원이 깎이면 안 된다."""
    ws = _ws({"A1": 14094600, "A2": 1, "A3": "=ROUNDDOWN((A1/A2),-3)", "B1": 0.1, "B2": 0.2,
              "B3": "=ROUNDDOWN((B1+B2)*10,0)", "B4": "=ROUNDUP((B1+B2)*10,0)"})
    assert fe.cell_value(ws, "A3") == 14094000
    assert fe.cell_value(ws, "B3") == 3 and fe.cell_value(ws, "B4") == 3


@pytest.mark.parametrize("formula, why", [
    ("=Sheet2!A1+1", "다른 시트"), ('=A1&"원"', "글자"), ("=IF(A1>0,1,2)", "읽을 수 없는"),
    ("=VLOOKUP(A1,A1:B2,2,0)", "다루지 않는 함수"), ("=A1/0", "0 으로"), ("=A1+", "중간에 끝"),
])
def test_unsupported_formulas_say_so(formula, why):
    with pytest.raises(fe.FormulaError, match=why):
        fe.cell_value(_ws({"A1": 1, "B1": 2, "A2": 3, "B2": 4, "C1": formula}), "C1")


def test_circular_reference_is_reported():
    with pytest.raises(fe.FormulaError, match="자기 자신"):
        fe.cell_value(_ws({"A1": "=A2+1", "A2": "=A1+1"}), "A1")


def test_china_quote_unit_price_matches_excel(quote_template):
    """실제 양식: 견적서_중국!F46(단가)을 계산한 값이 Excel 이 저장해 둔 값과 같다."""
    ws = load_workbook(quote_template)["견적서_중국"]
    cached = load_workbook(quote_template, data_only=True)["견적서_중국"]
    for ref in ("H17", "H20", "H41", "H42", "D46", "F46", "H46"):
        assert cached[ref].value is not None, ref
        assert fe.cell_value(ws, ref) == pytest.approx(cached[ref].value), ref
