"""견적서작성 설정 — 기본값, 설정 시트 읽기, 잘못된 값 처리, 실제 생성에 반영되는지."""
import os
import re

import pytest
from openpyxl import Workbook, load_workbook

import quote_config as qc

NO_COM = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), "no_com")


def test_defaults_and_helpers():
    cfg = qc.QuoteConfig()
    assert cfg["sheet.spec"] == "사양서" and cfg.overridden == {} and cfg.problems == []
    assert cfg.col("req.col_due") == 27 and cfg.at("dom.col_spec", 17) == "B17"
    assert cfg.cols("copy.cols") == list(range(1, 15)) + [18] + list(range(23, 28))
    assert cfg.capacity("dom.spec_start", "dom.spec_end") == 25
    assert cfg.text("name.cover_file", 폴더="F", 갑지표시=cfg["name.cover_mark"]) == "대외비_F_갑지.xlsx"
    assert qc.parse_cols("b, D:F") == [2, 4, 5, 6]


def test_every_default_is_valid_for_its_kind():
    for key, default, _, _ in qc._SPEC:
        assert qc._coerce(key, default) == default, key


def test_settings_rows_override_and_report_problems():
    rows = [("항목", "값", "설명"),
            ("dom.spec_end", 45.0, None),              # 엑셀은 정수도 실수로 준다
            ("dom.cell_investor", " $c$5 ", None),      # 공백·$·소문자 정리
            ("req.col_hogi", "ab", None),
            ("name.cn_file", "{날짜}_{라인}.xlsx", None),
            ("dom.text_no_class", 1.0, None),           # 숫자로 바뀐 문구
            ("sheet.spec", None, None),                 # 비워 두면 기본값
            ("# 메모", "x", None),
            ("dom.spec_start", "열일곱", None),          # 숫자 아님
            ("dom.cell_exh_size", "27D", None),         # 셀 주소 아님
            ("name.dom_file", "{날짜}_{없는것}.xlsx", None),
            ("없는.항목", 1, None)]
    over, problems = qc.read_settings_rows(rows)
    assert over == {"dom.spec_end": 45, "dom.cell_investor": "C5", "req.col_hogi": "AB",
                    "name.cn_file": "{날짜}_{라인}.xlsx", "dom.text_no_class": "1"}
    assert len(problems) == 4
    assert any("dom.spec_start" in p and "숫자" in p for p in problems)
    assert any("없는것" in p for p in problems) and any("없는.항목" in p for p in problems)
    cfg = qc.QuoteConfig(over, problems)
    assert cfg["dom.spec_end"] == 45 and cfg["sheet.spec"] == "사양서"
    assert "dom.text_no_class" not in cfg.overridden          # 기본값과 같은 값은 '바꾼 것'이 아니다


def test_check_reports_missing_sheets_and_bad_ranges():
    cfg = qc.QuoteConfig({"dom.spec_end": 45, "cn.rack_start": 50})
    names = [v for k, v in qc.DEFAULTS.items() if k.startswith("sheet.") and k != "sheet.cover"]
    out = cfg.check(names)
    assert any("견적서 갑지" in p for p in out)
    assert any("칸 수가 다릅니다" in p for p in out) and any("중국 RACK" in p for p in out)
    assert qc.QuoteConfig().check([v for k, v in qc.DEFAULTS.items() if k.startswith("sheet.")]) == []


def test_export_then_read_back_is_identity(tmp_path):
    path = qc.export_settings_sheet(str(tmp_path / "설정.xlsx"), qc.QuoteConfig({"dom.spec_end": 45}))
    wb = load_workbook(path)
    cfg = qc.from_workbook(wb)
    assert cfg.problems == [] and cfg.overridden == {"dom.spec_end": 45}
    assert wb[qc.SETTINGS_SHEET].max_row == len(qc._SPEC) + 1


def test_load_caches_until_file_changes(tmp_path):
    path = str(tmp_path / "t.xlsx")
    wb = Workbook(); wb.active.title = "품목"; wb.save(path)
    a = qc.load(path)
    assert a is qc.load(path) and a.overridden == {}
    wb = load_workbook(path)
    ws = wb.create_sheet(qc.SETTINGS_SHEET)
    ws.append(["항목", "값"]); ws.append(["default.investor", "홍길동"])
    wb.save(path)
    os.utime(path, (1, 1))                                  # 수정 시각이 달라지면 다시 읽는다
    assert qc.load(path)["default.investor"] == "홍길동"
    assert qc.load(None).overridden == {} and qc.load(str(tmp_path / "없음.xlsx")).overridden == {}


def test_every_key_is_used_by_the_code():
    """설정 항목을 만들어 놓고 코드가 안 읽으면 '고쳐도 안 바뀌는 설정'이 된다."""
    src = ""
    for name in ("excel_io.py", "page_quote.py", "quote_review.py"):
        with open(os.path.join(NO_COM, name), encoding="utf-8") as f:
            src += f.read()
    prefixes = set(re.findall(r'"((?:codemap|copy)\.col_)" \+', src))       # 조립해서 쓰는 키
    prefixes |= {m + "." for m in re.findall(r'_write_lines\(ws, cfg, "(\w+)"', src)}
    unused = [k for k in qc.DEFAULTS
              if f'"{k}"' not in src and not any(k.startswith(p) for p in prefixes)]
    assert unused == []


# ── 실제 양식으로: 설정을 바꾸면 생성물이 따라 바뀌는가 ──────────────────
def _item(cat, spec, qty, price, role="ITEM", amt=None):
    return {"cat": cat, "role": role, "spec": spec, "qty": qty, "price": price,
            "amt": qty * price if amt is None else amt}


def test_overrides_reach_the_generated_files(quote_template, tmp_path, monkeypatch):
    import shutil
    import core
    import excel_io
    monkeypatch.setattr(excel_io, "exe_dir", lambda: str(tmp_path / "out"))
    tpl = str(tmp_path / "양식.xlsx")
    shutil.copy(quote_template, tpl)
    wb = load_workbook(tpl)
    ws = wb.create_sheet(qc.SETTINGS_SHEET)
    for row in (["항목", "값", "설명"],
                ["dom.cell_investor", "C4"], ["dom.investor_text", "{투자자} 귀하"],
                ["dom.cell_equip_qty", "B43"], ["dom.spec_end", 30], ["dom.sign_end", 24],
                ["name.out_dir", "결과"], ["name.dom_file", "{호기}_{PR}.xlsx"],
                ["cover.cell_warranty", "H12"], ["cover.warranty_text", "보증 {보증}년"],
                ["name.cover_mark", "_표지"], ["default.investor", "홍길동"],
                ["us.cell_tool", "C54"], ["cn.rack_end", 25]):
        ws.append(row)
    wb.save(tpl)

    cfg = qc.load(tpl)
    assert cfg.problems == [] and len(cfg.overridden) == 13
    state = core.QuoteState(template_path=tpl, investor_name="", equip_qty=7)
    rd = {"row_idx": 2, "D": "PR1", "E": 10, "F": "Q-none", "H": 2, "J": "의뢰", "K": "L1", "Z": "H001"}
    items = [_item("PUMP", "P", 2, 1000)] + [_item("RACK", "R%d" % i, 1, 10) for i in range(3)]

    path = excel_io.generate_quote(state, "국내", items, rd)
    assert os.path.basename(path) == "H001_PR1.xlsx" and os.path.basename(os.path.dirname(os.path.dirname(path))) == "결과"
    out = load_workbook(path)
    assert qc.SETTINGS_SHEET not in out.sheetnames               # 설정 시트는 생성물에 실리지 않는다
    spec = out["사양서"]
    assert spec["C4"].value == "홍길동 귀하" and spec["B43"].value == 7
    assert spec.row_dimensions[19].hidden is False and spec.row_dimensions[20].hidden is True

    # 칸 수가 설정을 따른다: 사양서 14행(17~30) / 중국 RACK 5행(21~25)
    too_many = [_item("PUMP", "P", 1, 1)] + [_item("RACK", "R%d" % i, 1, 1) for i in range(15)]
    with pytest.raises(ValueError, match=r"14행"):
        excel_io.generate_quote(state, "국내", too_many, rd)
    state.cn_info = {"line": "L", "tool": "T", "maker": "M", "process": "P"}
    with pytest.raises(ValueError, match=r"5행"):
        excel_io.generate_quote(state, "중국", too_many[:7], {})

    state.us_info = {"site": "S", "tool": "T9", "maker": "M", "process": "P", "exchange": 1400}
    us = load_workbook(excel_io.generate_quote(state, "미국", items, {}))["견적서_미국"]
    assert us["C54"].value == "T9"

    cover, failed = excel_io.generate_cover(tpl, os.path.dirname(path), [path], "", warranty_years=3)
    assert failed == [] and cover.endswith("_표지.xlsx")
    cws = load_workbook(cover)["견적서 갑지"]
    assert cws["H12"].value == "보증 3년" and "홍길동" in cws["B7"].value
    submit = excel_io.export_cover_data_sheet(cover, tpl)
    assert "_표지" not in os.path.basename(submit)
