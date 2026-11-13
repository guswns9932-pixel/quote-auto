"""검토용 화면 부품 — 기타 옵션 줄의 '외 n개', 우클릭 말풍선."""
import os

import pytest

pytest.importorskip("PySide6.QtWidgets")
os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")

from PySide6.QtCore import QPoint, Qt
from PySide6.QtTest import QTest
from PySide6.QtWidgets import QApplication, QPushButton, QVBoxLayout, QWidget

import quote_review as rv


def _it(cat, spec, qty, price):
    return {"cat": cat, "role": "ITEM", "spec": spec, "qty": qty, "price": price, "amt": qty * price}


def _review(n_others, zero_at=None):
    items = [_it("PUMP", "P", 2, 100)]
    for i in range(n_others):
        items.append(_it("RACK", "옵션 품목 %02d" % i, i + 1, 0 if i == zero_at else 10))
    return rv.analyze(items, "국내", 2)


def test_more_chip_lists_hidden_items_and_selects_on_pick():
    app = QApplication.instance() or QApplication([])
    from quote_panels import DiagramPanel
    panel = DiagramPanel()
    panel.resize(700, 260); panel.show(); app.processEvents()
    panel.show_review(_review(12, zero_at=11))
    app.processEvents()

    chips = [b for b in panel.buttons if getattr(b, "is_other", False)]
    more = panel.more_button
    assert more is not None and 0 < len(chips) < 12
    hidden = more.hidden_slots
    assert len(chips) + len(hidden) == 12 and more.text() == f"외 {len(hidden)}개 ▾"
    # 마우스를 올리면: 숨은 품목의 이름·수량이 전부, 0원인 것은 표시와 함께
    tip = more.toolTip().split("\n")
    assert len(tip) == len(hidden) and tip[0].startswith(hidden[0].text) and "수량" in tip[0]
    assert "단가 0원" in tip[-1] and "#C62828" in more.styleSheet()      # 숨은 쪽에 문제가 있으면 빨간 테두리

    # 누르거나 우클릭하면 뜨는 목록: 고르면 그 품목을 선택해 달라고 알린다
    picked = []
    panel.pickSpec.connect(picked.append)
    menu = panel.more_menu(more)
    assert [a.text().split("   ·   ")[0] for a in menu.actions()] == [s.text for s in hidden]
    menu.actions()[1].trigger()
    assert picked == [hidden[1].text]
    assert more.contextMenuPolicy() == Qt.CustomContextMenu

    # 칩과 '외 n개' 가 실제로 화면에 보인다
    assert all(b.isVisible() for b in chips) and more.isVisible()

    # 가만히 두거나 높이만 바뀌면 다시 그리지 않는다 (끝없이 지웠다 그리며 꿀렁거리던 문제)
    rebuilds = []
    original = panel._fill_others
    panel._fill_others = lambda: (rebuilds.append(1), original())[1]
    for _ in range(5):
        QTest.qWait(40)
    panel.resize(700, 300); QTest.qWait(200)
    assert rebuilds == []

    # 넓어지면 더 많이 보이고, 다 들어가면 '외 n개' 가 사라진다 — 폭이 바뀔 때 한 번씩만 다시 그린다
    panel.resize(1300, 300); QTest.qWait(200)
    wide = [b for b in panel.buttons if getattr(b, "is_other", False)]
    assert len(wide) > len(chips) and len(rebuilds) == 1
    panel.resize(4000, 300); QTest.qWait(200)
    assert panel.more_button is None and len([b for b in panel.buttons if getattr(b, "is_other", False)]) == 12
    assert len(rebuilds) == 2
    QTest.qWait(200)
    assert len(rebuilds) == 2
    panel._fill_others = original
    # 문제가 없으면 회색 테두리
    panel.resize(700, 300); QTest.qWait(200); panel.show_review(_review(12)); app.processEvents()
    assert "#C62828" not in panel.more_button.styleSheet()
    panel.close(); panel.deleteLater(); app.processEvents()


def test_info_popup_closes_on_any_click_without_eating_it():
    """말풍선이 떠 있어도 다른 곳을 누르면 그 클릭이 그대로 동작하고 말풍선은 닫힌다."""
    app = QApplication.instance() or QApplication([])
    from widgets import InfoPopup
    host = QWidget(); lay = QVBoxLayout(host)
    btn = QPushButton("누르기"); lay.addWidget(btn)
    clicks = []
    btn.clicked.connect(lambda: clicks.append(1))
    host.resize(300, 100); host.show(); app.processEvents()

    popup = InfoPopup(host)
    popup.show_at("설비호기: H001\nMaker: MK", host.mapToGlobal(QPoint(10, 10)))
    app.processEvents()
    assert popup.isVisible() and popup.windowType() == Qt.ToolTip      # 입력을 독점하는 Popup 종류가 아니다

    QTest.mouseClick(btn, Qt.LeftButton)
    app.processEvents()
    assert clicks == [1] and not popup.isVisible()

    popup.show_at("다시", host.mapToGlobal(QPoint(10, 10)))
    QTest.keyClick(host, Qt.Key_Escape); app.processEvents()
    assert not popup.isVisible()
    host.close(); host.deleteLater(); app.processEvents()
