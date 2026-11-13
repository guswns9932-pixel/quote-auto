"""전자서명 화면 — 창을 띄우지 않고(offscreen) 실제 흐름을 돌려 본다.
불러오기 → 서명 놓기 → 다른 파일에 일괄 적용 → 취소 → 저장."""
import gc
import os

import pytest

fitz = pytest.importorskip("fitz")
pytest.importorskip("PySide6.QtWidgets")
os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")

from PySide6.QtCore import QSettings
from PySide6.QtGui import QColor, QImage
from PySide6.QtWidgets import QApplication, QMessageBox

import esign_engine as eng


@pytest.fixture()
def page(tmp_path, monkeypatch):
    QSettings.setPath(QSettings.IniFormat, QSettings.UserScope, str(tmp_path / "settings"))
    app = QApplication.instance() or QApplication([])
    monkeypatch.setattr(QMessageBox, "information", lambda *a, **k: QMessageBox.Ok)
    import page_esign
    p = page_esign.ESignPage()
    p.resize(1400, 900)
    p.show()
    app.processEvents()
    yield p
    p.shutdown()
    assert p._worker.wait(5000)          # Excel 이 없는 환경이라 워커는 바로 끝난다
    # 화면 객체는 여기(메인 스레드)에서 확실히 치운다. 남겨 두면 다음 테스트의
    # 워커 스레드에서 가비지 수집이 돌 때 Qt 객체가 엉뚱한 스레드에서 해제돼 죽는다.
    # (실제 앱은 화면을 한 번만 만들고 초기화도 같은 화면을 비워 쓰므로 해당 없음)
    p._clear_documents()
    p._worker = None
    p.close()
    p.deleteLater()
    app.processEvents()
    del p
    gc.collect()


def _signs():
    out = []
    for color in ("#d00000", "#0000d0"):
        im = QImage(340, 80, QImage.Format_ARGB32)
        im.fill(QColor(color))
        out.append(im)
    return out


def _scaled(name, sheets):
    """(시트, 쪽수, 인쇄배율) 로 파일을 만든다."""
    from test_esign_engine import _pdf
    pdf, pages = eng.merge_sheet_pdfs([(s, _pdf(n, "S%d" % k), {"zoom": z}) for k, (s, n, z) in enumerate(sheets)])
    return eng.SignFile(path=name, pdf=pdf, pages=pages)


def _files():
    three = [("사양서", 1, 62), ("현업 사인용 사양서", 1, 62), ("입고검수확인서", 1, 95)]
    return [_scaled("/x/갑지.xlsx", [("견적서 갑지", 1, 73)]),
            _scaled("/x/1호기.xlsx", three), _scaled("/x/2호기.xlsx", three),
            eng.SignFile(path="/x/깨진파일.xlsx", error="열 수 없음")]


def _wait(thread):
    app = QApplication.instance()
    for _ in range(500):
        app.processEvents()
        if thread.isFinished():
            break
        thread.wait(10)
    app.processEvents()


def test_full_flow(page, tmp_path):
    app = QApplication.instance()
    assert page.btn_save.isEnabled() is False
    assert page.set_sign_images(_signs()) is True
    page.show_files(_files())
    app.processEvents()

    # 목록: 실패한 파일은 사유와 함께, 첫 파일이 선택되고 쪽 미리보기가 채워진다
    assert page.file_list.count() == 4 and "열 수 없음" in page.file_list.item(3).text()
    assert page.file_list.currentRow() == 0 and page.page_list.count() == 1
    assert page.btn_save.isEnabled() is True

    # 1호기로 가서 사양서와 입고검수확인서에 서명
    page.file_list.setCurrentRow(1)
    assert page.page_list.count() == 3 and "사양서" in page.page_list.item(0).text()
    page.place_sign(0, 100, 700)
    page._next_page(); page._next_page()
    assert (page._cur_file, page._cur_page) == (1, 2)
    page.place_sign(1, 300, 650)
    assert "서명 2" in page.file_list.item(1).text()
    assert "✓ 서명 1" in page.page_list.item(2).text() and "서명 없음" in page.page_list.item(1).text()

    # 서명 없는 파일 경고 대상: 갑지는 원래 서명하지 않으므로 빠진다
    assert page.unsigned_files() == ["2호기.xlsx"]

    # 쪽을 넘기면 다음 파일로 이어지고, 실패한 파일은 건너뛴다
    page._next_page()
    assert (page._cur_file, page._cur_page) == (2, 0)
    page._next_page(); page._next_page(); page._next_page()
    assert (page._cur_file, page._cur_page) == (2, 2)
    page._prev_page(); page._prev_page(); page._prev_page()
    assert (page._cur_file, page._cur_page) == (1, 2)

    # 서명 크기는 시트의 인쇄 배율을 따른다 (사양서 62%, 입고검수확인서 95%)
    pl = page.placements()
    assert round(pl[(1, 0)][0].w, 2) == round(eng.SIGN_W_PT * 0.62, 2)
    assert round(pl[(1, 2)][0].w, 2) == round(eng.SIGN_W_PT * 0.95, 2)

    # 서명 크기 % 를 바꾸면 이미 놓은 서명도 가운데를 지키며 바뀌고, 설정에 기억된다
    before = pl[(1, 0)][0]
    page.spin_size.setValue(150)
    after = page.placements()[(1, 0)][0]
    assert round(after.w, 2) == round(eng.SIGN_W_PT * 0.62 * 1.5, 2)
    assert round(after.x + after.w / 2, 2) == round(before.x + before.w / 2, 2)
    assert round(after.y + after.h / 2, 2) == round(before.y + before.h / 2, 2)
    import app_settings
    assert app_settings.get_int(app_settings.Key.ESIGN_SIGN_PCT) == 150
    page.spin_size.setValue(100)

    # 서명 하나를 끌어 옮기면 저장 위치도 따라간다
    page.file_list.setCurrentRow(1)
    moved = page._sign_items[(1, 0)][0]
    moved.setPos(50, 60)
    w0, h0 = eng.sign_size(0.62)
    got = page.placements()[(1, 0)][0]
    assert (got.sign, got.x, got.y) == (0, 50.0, 60.0) and round(got.w, 2) == round(w0, 2)

    # 저장
    out = str(tmp_path / "out.pdf")
    page.save_to(out)
    _wait(page._pdf_thread)
    doc = fitz.open(out)
    assert doc.page_count == 7                                   # 1 + 3 + 3 (실패 파일 제외)
    assert [len(p.get_image_info()) for p in doc] == [0, 1, 0, 1, 0, 0, 0]
    box = doc[1].get_image_info()[0]["bbox"]
    # 화면용 서명 그림이 정수 픽셀이라 높이에 0.1pt 안쪽의 반올림 차이가 있을 수 있다
    assert all(abs(a - b) < 0.2 for a, b in zip(box, (50.0, 60.0, 50.0 + w0, 60.0 + h0)))
    assert doc[1].get_pixmap(clip=fitz.Rect(55, 62, 60, 66)).pixel(2, 2)[:3] == (208, 0, 0)
    doc.close()
    assert page.btn_save.isEnabled() is True

    page.grab().save(str(tmp_path / "screen.png"))
    shot = os.environ.get("ESIGN_SCREENSHOT")
    if shot:
        page.grab().save(shot)


def test_remove_sign_and_reset(page):
    page.set_sign_images(_signs())
    page.show_files(_files())
    page.file_list.setCurrentRow(1)
    page.place_sign(0, 10, 10)
    item = page._sign_items[(1, 0)][0]
    page.scene.removeItem(item)
    page._remove_sign(item)                    # 우클릭 삭제와 같은 경로
    assert page.placements() == {} and "서명" not in page.file_list.item(1).text()

    page.reset_page()
    assert page.file_list.count() == 0 and page.page_list.count() == 0
    assert page._signs == [] and page.btn_save.isEnabled() is False
    assert page.lbl_code_state.text() == "서명 미로드"


def test_open_result_folder(page, tmp_path, monkeypatch):
    """결과 폴더 열기 — 불러오기 전에는 꺼져 있고, 저장하면 그 PDF 의 폴더를 연다."""
    import page_esign
    opened = []
    monkeypatch.setattr(page_esign.os, "startfile", opened.append, raising=False)
    assert page.btn_open.isEnabled() is False
    page.set_sign_images(_signs())
    page.show_files(_files())
    out_dir = tmp_path / "결과"
    out_dir.mkdir()
    page.save_to(str(out_dir / "out.pdf"))
    _wait(page._pdf_thread)
    assert page.btn_open.isEnabled() is True
    page.btn_open.click()
    assert opened == [str(out_dir)]


def test_close_does_not_wait_for_cleanup(page):
    """닫을 때 Excel 종료를 기다리지 않는다 — shutdown 은 곧바로 돌아온다."""
    import time
    t = time.perf_counter()
    page.shutdown()
    assert time.perf_counter() - t < 1.0
