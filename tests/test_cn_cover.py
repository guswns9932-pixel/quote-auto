"""중국 갑지 — 중국 견적서 여러 건을 '중국 견적서 갑지' 한 장에 모은다."""
import datetime as dt
import os

import pytest
from openpyxl import load_workbook


def _item(cat, spec, qty, price):
    return {"cat": cat, "role": "ITEM", "spec": spec, "qty": qty, "price": price, "amt": qty * price}


@pytest.fixture()
def env(quote_template_cn_cover, tmp_path, monkeypatch):
    import core
    import excel_io
    monkeypatch.setattr(excel_io, "exe_dir", lambda: str(tmp_path / "out"))
    return excel_io, core.QuoteState(template_path=quote_template_cn_cover), tmp_path


def _make(io, st, tool, process, qty, pump_price, rack_price, line="X2"):
    st.cn_info = {"line": line, "tool": tool, "maker": "M", "process": process}
    return io.generate_quote(st, "중국", [_item("PUMP", "P", qty, pump_price), _item("RACK", "R", 2, rack_price)], {})


def test_values_are_collected_in_file_name_order(env):
    io, st, tmp = env
    # 이름 순서(숫자 크기대로): T2 < T10. 만든 순서는 일부러 반대로.
    b = _make(io, st, "T10", "ETCH", 2, 5000000, 300000)
    a = _make(io, st, "T2", "CVD", 3, 13900000, 120000)
    folder = os.path.dirname(a)
    out, failed = io.generate_cn_cover(st.template_path, folder, [b, a])
    assert failed == [] and os.path.basename(out) == f"대외비_{os.path.basename(folder)}_중국갑지.xlsx"

    wb = load_workbook(out)
    # 갑지 + 견적서들(시트 이름 = 호기명, 갑지의 줄 순서대로). 단가표 등 다른 시트는 실리지 않는다
    assert wb.sheetnames == ["중국 견적서 갑지", "T2", "T10"]
    ws = wb["중국 견적서 갑지"]

    # 단가 = 견적서_중국!F46 = ROUNDDOWN(((PUMP + RACK) × 1.014) / 수량, -3) — 파일에 계산 결과가 없어도 구한다
    def unit(qty, pump, rack):
        total = (pump * qty + rack * 2) * 1.014
        return int(total / qty // 1000 * 1000)
    assert [ws.cell(16, c).value for c in (2, 3, 4, 5)] == ["CVD", "T2", 3, unit(3, 13900000, 120000)]
    assert [ws.cell(17, c).value for c in (2, 3, 4, 5)] == ["ETCH", "T10", 2, unit(2, 5000000, 300000)]
    assert ws["B18"].value is None and ws["E18"].value is None       # 양식의 예시 값은 남지 않는다

    # Quote number: LOT + 오늘 + SCS - 파일 이름의 라인 - Q1~건수
    assert ws["C9"].value == f"LOT{dt.datetime.now():%y%m%d}SCS-X2-Q1~2"

    # 공정(B열)이 비는 첫 줄부터 45행까지 숨김, 채운 줄과 합계 줄은 보인다
    hidden = [r for r in range(14, 48) if ws.row_dimensions[r].hidden]
    assert hidden == list(range(18, 46))
    # 양식의 그림 유지: 갑지 3개, 견적서 시트마다 3개
    assert io.count_sheet_images(out) == {"중국 견적서 갑지": 3, "T2": 3, "T10": 3}

    # 견적서 시트는 원본 견적서와 같은 내용이다 — 값·수식, 서식, 숨긴 행, 병합, 열 너비, 인쇄 설정
    for path, title in ((a, "T2"), (b, "T10")):
        src = load_workbook(path)["견적서_중국"]
        dst = wb[title]
        cells = lambda w: {c.coordinate: (str(c.value), c.number_format, c.font.b, c.font.color.rgb if c.font.color and c.font.color.type == "rgb" else None,
                                          c.fill.fgColor.rgb if c.fill.fill_type else None, c.alignment.horizontal, c.border.bottom.style)
                           for row in w.iter_rows() for c in row if c.value is not None}
        assert cells(dst) == cells(src), title
        assert dst["B51"].value == title and dst["F46"].value == src["F46"].value == "=ROUNDDOWN((H42/D46),-3)"
        hid = lambda w: sorted(r for r, d in w.row_dimensions.items() if d.hidden)
        assert hid(dst) == hid(src) and hid(dst), title
        assert sorted(map(str, dst.merged_cells.ranges)) == sorted(map(str, src.merged_cells.ranges))
        assert {k: round(v.width or 0, 2) for k, v in dst.column_dimensions.items()} == \
               {k: round(v.width or 0, 2) for k, v in src.column_dimensions.items()}
        assert dst.print_area.split("!")[-1] == src.print_area.split("!")[-1]
        assert (dst.page_setup.orientation, dst.page_setup.scale, dst.page_setup.paperSize) == \
               (src.page_setup.orientation, src.page_setup.scale, src.page_setup.paperSize)
        assert len(dst.conditional_formatting) == len(src.conditional_formatting)
    # 열면 갑지가 선택돼 있다 (여러 시트가 함께 선택되면 Excel 이 '그룹' 으로 연다)
    assert [n for n in wb.sheetnames if wb[n].sheet_view.tabSelected] == ["중국 견적서 갑지"]


def test_quote_number_follows_count_and_lines(env):
    io, st, tmp = env
    one = _make(io, st, "A1", "CVD", 1, 1000000, 0)
    folder = os.path.dirname(one)
    ws = load_workbook(io.generate_cn_cover(st.template_path, folder, [one])[0])["중국 견적서 갑지"]
    assert ws["C9"].value == f"LOT{dt.datetime.now():%y%m%d}SCS-X2-Q1"            # 한 건이면 Q1

    other = _make(io, st, "B1", "CVD", 1, 1000000, 0, line="X1")
    three = _make(io, st, "A2", "CVD", 1, 1000000, 0)
    # 라인이 섞이면 파일 이름 순서대로 함께 적는다 (…_X1_B1 이 …_X2_A1 보다 앞)
    ws = load_workbook(io.generate_cn_cover(st.template_path, str(tmp / "out"), [one, three, other])[0])["중국 견적서 갑지"]
    assert ws["C9"].value.endswith("SCS-X1/X2-Q1~3")
    assert [ws.cell(r, 3).value for r in (16, 17, 18)] == ["B1", "A1", "A2"]

    # 파일 이름이 규칙과 다르면 라인 부분을 뺀다
    odd = str(tmp / "이름이다른파일.xlsx")
    os.replace(one, odd)
    ws = load_workbook(io.generate_cn_cover(st.template_path, str(tmp), [odd])[0])["중국 견적서 갑지"]
    assert ws["C9"].value == f"LOT{dt.datetime.now():%y%m%d}SCS-Q1"


def test_uses_excel_saved_value_when_present(env):
    """사람이 Excel 에서 고쳐 저장한 파일은 그 계산 결과를 쓴다."""
    io, st, tmp = env
    src = str(tmp / "대외비_260702_중국_SCS_X2_TAAL01.xlsx")
    import shutil
    shutil.copy(st.template_path, src)                 # 양식 자체: Excel 이 저장한 계산 결과가 들어 있다
    want = load_workbook(src, data_only=True)["견적서_중국"]
    ws = load_workbook(io.generate_cn_cover(st.template_path, str(tmp), [src])[0])["중국 견적서 갑지"]
    assert [ws.cell(16, c).value for c in (2, 3, 4, 5)] == [
        want["B17"].value, want["B51"].value, want["G17"].value, want["F46"].value]


def test_unreadable_files_are_reported_not_silently_dropped(env):
    io, st, tmp = env
    good = _make(io, st, "A1", "CVD", 1, 1000000, 0)
    from openpyxl import Workbook
    bad = str(tmp / "대외비_260702_중국_SCS_X2_BAD.xlsx")
    wb = Workbook(); wb.active.title = "다른시트"; wb.save(bad)
    out, failed = io.generate_cn_cover(st.template_path, str(tmp), [good, bad])
    assert [n for n, _ in failed] == ["대외비_260702_중국_SCS_X2_BAD.xlsx"] and "시트 없음" in failed[0][1]
    assert load_workbook(out)["중국 견적서 갑지"]["C9"].value.endswith("-Q1")     # 건수는 실제로 들어간 것만

    with pytest.raises(ValueError, match="기록할 견적서가 없습니다"):
        io.generate_cn_cover(st.template_path, str(tmp), [bad])


def test_more_than_capacity_is_refused(env, monkeypatch):
    io, st, tmp = env
    files = [str(tmp / f"f{i}.xlsx") for i in range(31)]
    for f in files:
        open(f, "wb").close()
    before = set(os.listdir(tmp))
    with pytest.raises(ValueError, match=r"30행.*31건"):
        io.generate_cn_cover(st.template_path, str(tmp), files)
    assert set(os.listdir(tmp)) == before              # 갑지 파일을 만들지 않았다


def test_page_button_generates_from_selected_files(env, monkeypatch):
    """화면: 중국갑지 버튼 → 파일 고르기 → 갑지 생성."""
    pytest.importorskip("PySide6.QtWidgets")
    os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")
    import gc
    from PySide6.QtCore import QSettings
    from PySide6.QtWidgets import QApplication, QFileDialog, QMessageBox
    io, st, tmp = env
    a = _make(io, st, "T1", "CVD", 1, 1000000, 0)
    b = _make(io, st, "T2", "CVD", 1, 2000000, 0)
    QSettings.setPath(QSettings.IniFormat, QSettings.UserScope, str(tmp / "settings"))
    app = QApplication.instance() or QApplication([])
    shown = []
    monkeypatch.setattr(QMessageBox, "information", lambda *x, **k: shown.append(x[2]) or QMessageBox.Ok)
    monkeypatch.setattr(QMessageBox, "warning", lambda *x, **k: shown.append(x[2]) or QMessageBox.Ok)
    monkeypatch.setattr(QMessageBox, "critical", lambda *x, **k: shown.append(x[2]) or QMessageBox.Ok)
    import page_quote
    page = page_quote.QuoteBuilderPage()

    page._generate_cn_cover()                                    # 통합양식을 아직 안 불러왔다
    assert "통합양식을 먼저" in shown[-1]

    page.state.template_path = st.template_path
    monkeypatch.setattr(QFileDialog, "getOpenFileNames", lambda *x, **k: ([a, b], ""))
    page._generate_cn_cover()
    assert "중국 갑지 생성 완료" in shown[-1]
    made = [f for f in os.listdir(os.path.dirname(a)) if f.endswith("_중국갑지.xlsx")]
    assert len(made) == 1 and page._list_all_files[0].endswith("_중국갑지.xlsx")
    page.close(); page.deleteLater(); app.processEvents(); del page; gc.collect()


def test_older_template_with_merged_cells_does_not_crash(quote_template, tmp_path, monkeypatch):
    """목록 칸 일부가 병합된 예전 양식(ver.2.0)에서도 알 수 없는 오류로 죽지 않는다 —
    만들어지거나, 병합 때문에 못 적는 칸이 있으면 무엇을 고칠지 알려준다."""
    import core
    import excel_io
    monkeypatch.setattr(excel_io, "exe_dir", lambda: str(tmp_path / "out"))
    st = core.QuoteState(template_path=quote_template)
    src = _make(excel_io, st, "A1", "CVD", 1, 1000000, 0)
    try:
        out, failed = excel_io.generate_cn_cover(quote_template, str(tmp_path), [src])
    except ValueError as e:
        assert "병합돼 있어" in str(e)
        assert not [f for f in os.listdir(tmp_path) if f.endswith("_중국갑지.xlsx")]
    else:
        assert failed == [] and load_workbook(out)["중국 견적서 갑지"]["C9"].value.endswith("-X2-Q1")


def test_quote_sheets_get_safe_unique_names(env):
    """시트 이름: 호기명이 같으면 _2 를 붙이고, 엑셀이 못 쓰는 글자는 바꾸고, 비어 있으면 번호로."""
    io, st, tmp = env
    files = [_make(io, st, "SAME", "CVD", 1, 1000000, 0), _make(io, st, "SAME", "CVD", 1, 2000000, 0),
             _make(io, st, "A/B:1", "CVD", 1, 3000000, 0)]
    # 호기 칸을 비운 견적서
    blank = _make(io, st, "ZZ", "CVD", 1, 4000000, 0)
    wb = load_workbook(blank); wb["견적서_중국"]["B51"] = None; wb.save(blank)
    out, failed = io.generate_cn_cover(st.template_path, os.path.dirname(files[0]), files + [blank])
    names = load_workbook(out).sheetnames
    assert failed == [] and names[0] == "중국 견적서 갑지"
    assert sorted(names[1:]) == sorted(["A_B_1", "SAME", "SAME_2", names[-1]])
    assert names[-1].startswith("견적서") and len(set(n.lower() for n in names)) == 5

    # 수식만 있고 사람이 고친 값도 그대로 따라온다
    edited = _make(io, st, "EDIT", "CVD", 2, 5000000, 0)
    wb = load_workbook(edited); wb["견적서_중국"]["A50"] = "사람이 적은 메모"; wb.save(edited)
    out2, _ = io.generate_cn_cover(st.template_path, str(tmp), [edited])
    assert load_workbook(out2)["EDIT"]["A50"].value == "사람이 적은 메모"


def test_same_logo_is_stored_once(env):
    """견적서 시트마다 같은 로고가 따로 저장되지 않는다 — 그림은 그대로 보이되 파일은 작게."""
    import zipfile
    io, st, tmp = env
    files = [_make(io, st, "T%d" % i, "CVD", 1, 1000000, 0) for i in range(6)]
    out, _ = io.generate_cn_cover(st.template_path, os.path.dirname(files[0]), files)
    counts = io.count_sheet_images(out)
    assert counts["중국 견적서 갑지"] == 3 and all(counts["T%d" % i] == 3 for i in range(6))
    with zipfile.ZipFile(out) as z:
        names = z.namelist()
        media = [n for n in names if n.startswith("xl/media/")]
        assert len(media) <= 6                                   # 21개(7시트 × 3)가 아니라 서로 다른 그림 수만큼
        # 모든 시트의 그림 연결이 실제로 있는 파일을 가리킨다
        import re
        for n in names:
            if n.startswith("xl/drawings/_rels/"):
                for target in re.findall(r'Target="\\.\\./media/([^"]+)"', z.read(n).decode("utf8")):
                    assert "xl/media/" + target in names, (n, target)
        assert z.testzip() is None
    assert os.path.getsize(out) < 1024 * 1024                    # 1MB 미만
    wb = load_workbook(out)                                      # 다시 읽어도 그림이 그대로다
    assert len(wb["T3"]._images) == 3 and len(wb["중국 견적서 갑지"]._images) == 3
