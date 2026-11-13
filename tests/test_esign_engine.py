"""전자서명 엔진 — Excel 없이 확인할 수 있는 부분(합치기·위치 복사·서명 찍기·임시 파일)."""
import os

import pytest

fitz = pytest.importorskip("fitz")

import esign_engine as eng


def _pdf(n_pages, text, w=595, h=842):
    doc = fitz.open()
    for i in range(n_pages):
        page = doc.new_page(width=w, height=h)
        page.insert_text((72, 72), "%s p%d" % (text, i + 1))
    data = doc.tobytes()
    doc.close()
    return data


def _png(color=(255, 0, 0)):
    pix = fitz.Pixmap(fitz.csRGB, fitz.IRect(0, 0, 170, 40))
    pix.set_rect(pix.irect, color)
    return pix.tobytes("png")


def _file(name, sheets):
    # 쪽 글자는 영문 표식(S0, S1 …)으로 — 기본 글꼴에 한글이 없다
    pdf, pages = eng.merge_sheet_pdfs([(s, _pdf(n, "S%d" % k)) for k, (s, n) in enumerate(sheets)])
    return eng.SignFile(path=name, pdf=pdf, pages=pages)


def test_merge_keeps_sheet_of_every_page():
    f = _file("a.xlsx", [("견적서 갑지", 1), ("사양서", 2), ("입고검수확인서", 1)])
    assert [(p.sheet, p.index_in_sheet, p.pages_in_sheet) for p in f.pages] == [
        ("견적서 갑지", 0, 1), ("사양서", 0, 2), ("사양서", 1, 2), ("입고검수확인서", 0, 1)]
    doc = eng.open_pdf(f.pdf)
    assert doc.page_count == 4 and "S1 p2" in doc[2].get_text()
    doc.close()


def test_merge_skips_broken_part():
    pdf, pages = eng.merge_sheet_pdfs([("A", b"not a pdf"), ("B", _pdf(1, "B"))])
    assert [p.sheet for p in pages] == ["B"]
    assert eng.merge_sheet_pdfs([]) == (b"", [])


def test_render_page_size_follows_zoom():
    f = _file("a.xlsx", [("사양서", 1)])
    doc = eng.open_pdf(f.pdf)
    assert eng.page_size(doc, 0) == (595.0, 842.0)
    w, h, stride, samples = eng.render_page(doc, 0, zoom=2.0)
    assert (w, h) == (1190, 1684) and len(samples) == stride * h
    doc.close()


def test_print_scale():
    # 배율을 직접 지정한 시트 (갑지 73%)
    assert eng.print_scale({"zoom": 73}, 595, 842) == 0.73
    # 한 쪽 너비에 맞춤, 높이는 자동 — 여백 뺀 너비 / 인쇄영역 너비, 1% 단위 내림
    info = {"zoom": False, "wide": 1, "tall": False, "left": 50.4, "right": 50.4,
            "top": 54, "bottom": 54, "area_w": 797.0, "area_h": 600.0}
    assert eng.print_scale(info, 595, 842) == 0.62
    # 너비·높이 모두 맞춤이면 더 많이 줄여야 하는 쪽을 따른다
    assert eng.print_scale(dict(info, tall=1, area_h=1500.0), 595, 842) == 0.48
    # 맞춤이어도 늘리지는 않는다 / 설정을 못 읽으면 100%
    assert eng.print_scale(dict(info, area_w=100.0), 595, 842) == 1.0
    assert eng.print_scale({}, 595, 842) == 1.0


def test_sign_size_follows_sheet_scale_and_user_percent():
    assert eng.sign_size(1.0) == (eng.SIGN_W_PT, eng.SIGN_H_PT)
    w, h = eng.sign_size(0.62, 110)
    assert round(w, 2) == round(eng.SIGN_W_PT * 0.62 * 1.1, 2) and round(h / w, 4) == round(eng.SIGN_H_PT / eng.SIGN_W_PT, 4)


def test_merge_records_scale_per_sheet():
    pdf, pages = eng.merge_sheet_pdfs([("갑지", _pdf(1, "A"), {"zoom": 73}), ("사양서", _pdf(2, "B"), {"zoom": 62}),
                                       ("기타", _pdf(1, "C"))])
    assert [p.scale for p in pages] == [0.73, 0.62, 0.62, 1.0]


def test_build_signed_pdf_puts_image_where_placed(tmp_path):
    files = [_file("1.xlsx", [("사양서", 1), ("입고검수확인서", 1)]),
             eng.SignFile(path="broken.xlsx", error="x"),                # 실패한 파일은 건너뜀
             _file("3.xlsx", [("사양서", 1)])]
    placements = {(0, 1): [eng.Placement(0, 100, 700)], (2, 0): [eng.Placement(1, 300, 650)]}
    out = str(tmp_path / "out.pdf")
    seen = []
    n = eng.build_signed_pdf(files, placements, [_png((255, 0, 0)), _png((0, 0, 255))], out,
                             progress_cb=lambda d, t: seen.append((d, t)))
    assert n == 3 and seen[-1] == (3, 3)
    doc = fitz.open(out)
    assert [len(p.get_image_info()) for p in doc] == [0, 1, 1]
    # 글자가 그림으로 바뀌지 않고 그대로 남아 있다
    assert "S0 p1" in doc[0].get_text()
    # 서명이 놓은 자리에 놓은 크기로 들어갔다
    info = doc[1].get_image_info()[0]["bbox"]
    assert [round(v, 1) for v in info] == [100.0, 700.0, 100.0 + eng.SIGN_W_PT, 700.0 + eng.SIGN_H_PT]
    # 그 자리를 그려 보면 서명 색이다
    pix = doc[2].get_pixmap(clip=fitz.Rect(310, 655, 320, 665))
    assert pix.pixel(2, 2)[:3] == (0, 0, 255)
    doc.close()


def test_build_with_nothing_to_save_leaves_no_file(tmp_path):
    out = str(tmp_path / "out.pdf")
    with pytest.raises(ValueError):
        eng.build_signed_pdf([eng.SignFile(path="x.xlsx", error="e")], {}, [], out)
    assert not os.path.exists(out)


class _FakeSetup:
    Zoom, FitToPagesWide, FitToPagesTall = 62, False, False
    LeftMargin = RightMargin = TopMargin = BottomMargin = 50.0
    PrintArea = ""


class _FakeRange:
    Width, Height = 800.0, 600.0


class _FakeSheet:
    def __init__(self, name, visible=True, fail=False):
        self.Name, self.Visible, self._fail = name, (-1 if visible else 0), fail
        self.PageSetup, self.UsedRange = _FakeSetup(), _FakeRange()

    def ExportAsFixedFormat(self, kind, path):
        assert kind == 0
        if self._fail:
            raise RuntimeError("인쇄할 내용 없음")
        with open(path, "wb") as f:
            f.write(_pdf(1, self.Name))


class _FakeSheets(list):
    def __call__(self, name):
        for s in self:
            if s.Name == name:
                return s
        raise KeyError(name)


class _FakeBook:
    def __init__(self, sheets):
        self.Worksheets = _FakeSheets(sheets)
        self.closed = None

    def Close(self, save):
        self.closed = save


class _FakeExcel:
    def __init__(self, book):
        self.book = book
        self.Workbooks = self
        self.opened = None

    def Open(self, path, **kw):
        self.opened = (path, kw)
        return self.book


def test_export_reads_then_deletes_temp_pdfs(tmp_path):
    """Excel 대역으로 확인: 대상·보이는 시트만, 실패한 시트는 건너뛰고, 임시 PDF 는 남지 않는다."""
    xlsx = tmp_path / "견적.xlsx"
    xlsx.write_bytes(b"")
    book = _FakeBook([_FakeSheet("견적서 갑지", visible=False), _FakeSheet("사양서"),
                      _FakeSheet("현업 사인용 사양서", fail=True), _FakeSheet("입고검수확인서"),
                      _FakeSheet("품목")])
    xl = _FakeExcel(book)
    targets = ["견적서 갑지", "사양서", "현업 사인용 사양서", "입고검수확인서"]
    sf = eng.load_workbook_pdf(xl, str(xlsx), targets)
    assert sf.error == "" and [p.sheet for p in sf.pages] == ["사양서", "입고검수확인서"]
    assert [p.scale for p in sf.pages] == [0.62, 0.62]          # 시트의 인쇄 배율을 함께 읽는다
    assert xl.opened[1]["ReadOnly"] is True and book.closed is False
    assert os.listdir(tmp_path) == ["견적.xlsx"]


def test_export_falls_back_to_all_visible_sheets(tmp_path):
    xlsx = tmp_path / "기타.xlsx"
    xlsx.write_bytes(b"")
    xl = _FakeExcel(_FakeBook([_FakeSheet("A"), _FakeSheet("B", visible=False), _FakeSheet("C")]))
    sf = eng.load_workbook_pdf(xl, str(xlsx), ["사양서"])
    assert [p.sheet for p in sf.pages] == ["A", "C"]


def test_stale_temp_cleanup_only_touches_own_files(tmp_path):
    (tmp_path / (eng.TMP_PREFIX + "123_abc_00.pdf")).write_bytes(b"x")
    (tmp_path / "대외비_견적.pdf").write_bytes(b"x")
    (tmp_path / "견적.xlsx").write_bytes(b"x")
    assert eng.cleanup_stale_temp(str(tmp_path)) == 1
    assert sorted(os.listdir(tmp_path)) == ["견적.xlsx", "대외비_견적.pdf"]
