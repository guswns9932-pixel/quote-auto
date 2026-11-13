"""CSP UPLOAD: 출력 파일과 양식 읽기."""
import pytest
from openpyxl import Workbook, load_workbook

from conftest import xlsx_fingerprint


@pytest.fixture(scope="module")
def csp():
    import csp_core          # 화면(tkinter) 없이 불러올 수 있어야 한다
    return csp_core


def _template_headers(csp):
    hdr = [h for _, h, _ in csp.COLUMNS]
    hdr[12], hdr[13] = "고객설비", "고객공정"      # 실제 양식의 13·14열 이름
    return ["*" + h if req else h for h, (_, _, req) in zip(hdr, csp.COLUMNS)]


def _make_template(path, headers):
    wb = Workbook()
    ws = wb.active
    ws.title = "Sheet1"
    for i, h in enumerate(headers, 1):
        ws.cell(1, i, h)
    wb.save(path)
    return str(path)


def _rows(csp, n=2):
    return [{k: "%s%d" % (k, i) for k, _, _ in csp.COLUMNS} for i in range(n)]


def test_maker_and_process_go_to_the_right_columns(csp, tmp_path):
    """13열 고객설비 = 설비MAKER(키 N), 14열 고객공정 = 대공정(키 M)."""
    tpl = _make_template(tmp_path / "t.xlsx", _template_headers(csp))
    out = csp.build_output(tpl, _rows(csp), str(tmp_path / "o.xlsx"))
    ws = load_workbook(out).active
    assert (ws.cell(1, 13).value, ws.cell(2, 13).value) == ("고객설비", "N0")
    assert (ws.cell(1, 14).value, ws.cell(2, 14).value) == ("고객공정", "M0")


def test_inserted_column_is_followed(csp, tmp_path):
    """양식 중간에 열이 추가돼도 헤더 글자를 따라 값이 들어간다."""
    hdr = _template_headers(csp)
    tpl = _make_template(tmp_path / "t.xlsx", hdr[:5] + ["신규열"] + hdr[5:])
    out = csp.build_output(tpl, _rows(csp), str(tmp_path / "o.xlsx"))
    ws = load_workbook(out).active
    assert ws.cell(1, 6).value == "신규열" and ws.cell(2, 6).value is None
    assert (ws.cell(1, 26).value, ws.cell(2, 26).value) == ("*통신유형", "Y0")


def test_missing_header_stops(csp, tmp_path):
    hdr = _template_headers(csp)
    hdr[13] = "다른이름"
    tpl = _make_template(tmp_path / "t.xlsx", hdr)
    with pytest.raises(ValueError, match="고객공정"):
        csp.build_output(tpl, _rows(csp), str(tmp_path / "o.xlsx"))


def test_output_uses_shared_strings_and_text_dates(csp, tmp_path):
    import zipfile
    tpl = _make_template(tmp_path / "t.xlsx", _template_headers(csp))
    out = csp.build_output(tpl, _rows(csp), str(tmp_path / "o.xlsx"))
    with zipfile.ZipFile(out) as z:
        assert "xl/sharedStrings.xml" in z.namelist()
        assert b"inlineStr" not in z.read("xl/worksheets/sheet1.xml")
    ws = load_workbook(out).active
    assert [ws.cell(2, c).number_format for c in (7, 10, 20)] == ["@", "@", "@"]


def test_small_helpers(csp):
    assert csp.extract_after_underscore("P1F_CVD") == "CVD"
    assert csp.extract_model_keyword("DRY_PUMP;EQ,LOT,HD4500PW") == "HD4500"
    assert csp.extract_model_keyword("no marker") == ""
    assert csp.parse_int("1,000") == 1000 or csp.parse_int("1000") == 1000


def test_real_template_output(csp, csp_template, tmp_path, baseline):
    out = csp.build_output(csp_template, _rows(csp, 3), str(tmp_path / "o.xlsx"))
    baseline.check("csp_output", "csp", xlsx_fingerprint(out))


def test_real_template_master(csp, csp_template, baseline):
    md = csp.MasterData(csp_template)
    summary = {
        "codes": [md.order_types, md.channels, md.inco_terms, md.currencies, md.comm_types],
        "partners": [len(md.sold_to), len(md.ship_to)],
        "fsc": md.fsc, "note": md.fsc_filter_note, "cip": md.cip_rows,
        "fsc_map": sorted((list(k), v) for k, v in md.fsc_map.items()),
        "fsc_map_5key": sorted((list(k), v) for k, v in md.fsc_map_5key.items()),
        "opts": [md.opt_sites, md.opt_devices, md.opt_processes, md.opt_vendors,
                 md.opt_subprocs, md.opt_des_values, md.fsc_map_qcodes],
        "des": [sorted(md.fsc_map_des_to_qcode.items()), sorted(md.fsc_map_qcode_to_des.items())],
    }
    baseline.check("csp_master", "csp", summary)
