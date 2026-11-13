"""견적서작성 화면 — 창을 띄우지 않고(offscreen) 검토 부품이 실제 흐름에 반응하는지 본다."""
import gc
import os

import pytest

pytest.importorskip("PySide6.QtWidgets")
os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")

from PySide6.QtCore import QSettings
from PySide6.QtWidgets import QApplication, QMessageBox

import quote_review as rv


@pytest.fixture()
def page(tmp_path, monkeypatch):
    QSettings.setPath(QSettings.IniFormat, QSettings.UserScope, str(tmp_path / "settings"))
    app = QApplication.instance() or QApplication([])
    monkeypatch.setattr(QMessageBox, "information", lambda *a, **k: QMessageBox.Ok)
    import page_quote
    p = page_quote.QuoteBuilderPage()
    p.resize(1900, 1000)
    p.show()
    app.processEvents()
    yield p
    p.close()
    p.deleteLater()
    app.processEvents()
    del p
    gc.collect()


def _load_template(page, path):
    """화면이 백그라운드로 하는 일을 그대로, 다만 기다리지 않고 바로."""
    import excel_io
    import quote_config
    from openpyxl import load_workbook
    wb = load_workbook(path, read_only=True)
    cfg = quote_config.from_workbook(wb)
    rows, by_class, price_by_spec, order = excel_io.parse_items_sheet(wb[cfg["sheet.items"]], cfg)
    cmap = excel_io.parse_code_map_sheet(wb[cfg["sheet.code_map"]], cfg)
    wb.close()
    page._on_template_loaded(path, (rows, by_class, price_by_spec, order, cmap), silent=True)
    return rows, cmap


def test_review_panels_follow_the_work(page, quote_template):
    app = QApplication.instance()
    assert "통합양식" in page.step_bar.hint.text() and page.check_panel.list.count() == 1

    rows, cmap = _load_template(page, quote_template)
    page._refresh_review()            # 화면에서는 잠깐(60ms) 모았다가 스스로 갱신한다
    assert "의뢰파일" in page.step_bar.hint.text()

    # 의뢰 한 행 (수량 3CH)
    page.state.request_rows = [{"row_idx": 2, "D": "PR1", "E": 10, "F": "", "G": "", "H": 3, "J": "의뢰",
                                "K": "L1", "N": "", "R": "", "V": "", "X": "", "Y": "", "Z": "H001", "AA": None}]
    page._fill_req_table(page.state.request_rows)
    page.req_table.setCurrentCell(0, 1)

    # PUMP 1종 + 분류가 있는 RACK 품목 몇 개 + 품목 시트에 없는 이름 하나
    by_class = {}
    for r in rows:
        if r["A"] and r["B"] and r["C"] > 0:
            by_class.setdefault(r["A"], r)
    pump = next(r for r in rows if r["A"].upper().startswith("Q") and r["C"] > 0)
    page._add_row("PUMP", pump["B"], 3, pump["C"])
    page._add_row("RACK", by_class["Gate Valve"]["B"], 2, by_class["Gate Valve"]["C"])      # 2 ≠ 3CH
    page._add_row("RACK", by_class["Pump Rack"]["B"], 1, by_class["Pump Rack"]["C"])
    page._add_row("RACK", "양식에 없는 품목", 1, 0)
    page._refresh_review()
    review = page._review

    flow = {s.label: s for s in review.flow}
    assert flow["PUMP"].state == rv.OK and flow["Gate Valve"].state == rv.DIFF
    assert flow["Metal Bellows"].state == rv.BLANK
    assert {s.label: s.state for s in review.base}["Pump Rack"] == rv.OK
    texts = [c.text for c in review.checks]
    assert "단가 0원: 양식에 없는 품목 — 품목 시트에 없는 이름" in texts
    assert any(t.startswith("Gate Valve 수량 2 ≠ 3CH") for t in texts)
    assert "판단이 필요한 항목이 1건" in page.step_bar.hint.text()

    # 금액 패널 = 파일 생성과 같은 계산
    import quote_calc
    want = quote_calc.domestic(page._snapshot_items(), 3)
    assert page.amount_panel.unit.text() == f"{int(want.unit):,} 원"
    assert page.amount_panel.cap_text.text() == "품목 칸 사용  3 / 25행"

    # 도식의 칸을 누르면 표에서 그 품목이 선택된다
    gate_btn = next(b for b in page.diagram.buttons if b.slot.label == "Gate Valve")
    gate_btn.click()
    row = page.step5_table.currentRow()
    assert page.step5_table.item(row, 2).text() == by_class["Gate Valve"]["B"]

    # 0원 품목을 지우면 빨간 항목이 사라지고 단계 표시가 '생성'으로 넘어간다
    for r in range(page.step5_table.rowCount()):
        it = page.step5_table.item(r, 2)
        if it and it.text() == "양식에 없는 품목":
            page.step5_table.removeRow(r)
            break
    app.processEvents()
    page._refresh_review()
    assert not page._review.red and "견적서 생성을 누르세요" in page.step_bar.hint.text()

    shot = os.environ.get("QUOTE_SCREENSHOT")
    if shot:
        app.processEvents()
        page.grab().save(shot)


def test_code_map_suggestion_can_be_added_from_the_diagram(page, quote_template):
    rows, cmap = _load_template(page, quote_template)
    class_of = {r["B"]: r["A"] for r in rows}
    flow_classes = {"Gate Valve", "Metal Bellows", "Hot N2 Transfer", "Pump Rack",
                    "Power Distribution Box", "성형 Bellows"}
    key, entries = next((k, v) for k, v in cmap.items()
                        if any(class_of.get(acc) in flow_classes for acc, _ in v))
    page.state.process, page.state.vendor, page.state.code_5d = key
    page._add_row("PUMP", "P", 1, 1000)
    page._refresh_review()
    miss = [b for b in page.diagram.buttons if b.slot.state == rv.MISS]
    assert miss and any(c.level == rv.INFO for c in page._review.checks)
    spec = miss[0].slot.suggest[0]
    miss[0].click()                                   # 눌러서 담기
    page._refresh_review()
    assert spec in [i["spec"] for i in page._snapshot_items()]
    assert not any(c.suggest and c.suggest[0] == spec for c in page._review.checks)


def test_generate_asks_before_proceeding_with_red_items(page, quote_template, monkeypatch):
    _load_template(page, quote_template)
    page._add_row("RACK", "양식에 없는 품목", 1, 0)
    asked = []
    monkeypatch.setattr(QMessageBox, "question", lambda *a, **k: asked.append(a[2]) or QMessageBox.No)
    page._generate_quote()
    assert len(asked) == 1 and "PUMP 가 담기지 않았습니다" in asked[0] and "단가 0원" in asked[0]
    assert page._gen_qt_thread is None                # '아니오' → 생성하지 않는다


def test_settings_sheet_button(page, quote_template, tmp_path, monkeypatch):
    from PySide6.QtWidgets import QFileDialog
    from openpyxl import load_workbook
    _load_template(page, quote_template)
    out = str(tmp_path / "설정")
    monkeypatch.setattr(QFileDialog, "getSaveFileName", lambda *a, **k: (out, ""))
    page._export_settings_sheet()
    ws = load_workbook(out + ".xlsx")["설정"]
    assert ws["A2"].value == "sheet.items" and ws.max_row > 100


def test_layout_b_request_list_picker_window_and_notices(page, quote_template):
    """좁은 의뢰 목록 + 툴팁, 품목 추가 창, 로그 창 대신 점검 목록의 '참고'."""
    app = QApplication.instance()
    _load_template(page, quote_template)
    rd = {"row_idx": 2, "D": "PR1", "E": 10, "F": "", "G": "DRY_PUMP;EQ,LOT,XD1200", "H": 3, "J": "의뢰자",
          "K": "L1_ETCH", "N": "", "R": "3CH", "V": "5D-NAME", "X": "MAKER", "Y": "MODEL", "Z": "H001", "AA": None}
    page.state.request_rows = [rd]
    page._fill_req_table([rd])

    # 목록에는 호기·수량만 보이고, 나머지는 툴팁에 전부 있다
    visible = [c for c in range(page.req_table.columnCount()) if not page.req_table.isColumnHidden(c)]
    assert visible == [0, 1, 3]
    tip = page.req_table.item(0, 1).toolTip()
    for part in ("설비호기: H001", "투자정보: EQ_XD1200", "수량: 3", "Rack: 3CH", "Maker: MAKER",
                 "설비: MODEL", "5D: 5D-NAME", "PR: PR1", "라인: L1_ETCH"):
        assert part in tip

    # 품목 목록은 별도 창 — 버튼으로 열고, 거기서 더블클릭하면 담긴다
    assert page._picker.isVisible() is False and page._picker.isWindow()
    page.btn_add_item.click()
    app.processEvents()
    assert page._picker.isVisible() is True
    page._do_filter()
    assert page.items_table.rowCount() > 0 and page.items_table.window() is page._picker

    # 로그 창은 없고, 양식 알림은 점검 목록에 참고로 남는다
    assert not hasattr(page, "log_view")
    page._log("[설정 확인] 시트 없음: '예시'")
    page._log("STEP 메시지")
    page._refresh_review()
    assert [c.text for c in page._review.checks] == ["[설정 확인] 시트 없음: '예시'"]
    assert page.lbl_last.text() == "STEP 메시지"

    # 옵션은 오른쪽 탭 — 요약을 누르면 그 탭이 열린다
    assert page.side_tabs.currentIndex() == 0 and "국내" in page.btn_opt_summary.text()
    page.btn_opt_summary.click()
    assert page.side_tabs.currentIndex() == 1
    page._picker.hide()
