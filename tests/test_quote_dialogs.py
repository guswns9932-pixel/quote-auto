"""견적서작성 옵션창 — 입력한 값이 상태에 담기고, 다시 열면 그대로 보이는지."""
import os

import pytest

pytest.importorskip("PySide6.QtWidgets")
os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")

from PySide6.QtWidgets import QApplication


def test_us_and_china_option_dialogs_round_trip():
    app = QApplication.instance() or QApplication([])
    import core
    import page_quote
    st = core.QuoteState()

    d = page_quote.USOptionsDialog(None, st)
    d._tool.setText(" T1 "); d._quote_no.setText(" LOT-01 "); d._pr.setText("7300001234")
    d.apply()
    assert st.us_info["tool"] == "T1" and st.us_info["quote_no"] == "LOT-01" and st.us_info["pr"] == "7300001234"
    again = page_quote.USOptionsDialog(None, st)
    assert again._quote_no.text() == "LOT-01" and again._pr.text() == "7300001234"

    from PySide6.QtCore import QDate
    c = page_quote.ChinaOptionsDialog(None, st)
    # 날짜는 처음에 '입력 안 함' — 그대로 확인하면 비어 있다
    assert c._delivery_on.isChecked() is False and c._delivery.isEnabled() is False
    c.apply()
    assert st.cn_info["delivery"] == ""
    c._quote_no.setText("CN-01"); c._to.setText("Wang")
    c._delivery_on.setChecked(True); c._delivery.setDate(QDate(2026, 12, 20))
    c.apply()
    assert (st.cn_info["quote_no"], st.cn_info["delivery"], st.cn_info["to"]) == ("CN-01", "2026-12-20", "Wang")
    again_cn = page_quote.ChinaOptionsDialog(None, st)
    assert again_cn._delivery_on.isChecked() and again_cn._delivery.date() == QDate(2026, 12, 20)

    # 아래 줄 버튼 순서: 중국갑지 — OK — Cancel (중국갑지가 OK 의 왼쪽)
    c.show(); app.processEvents()
    assert c.btn_cover.text() == "중국갑지"
    assert c.btn_cover.x() < c.btn_ok.x() < c.btn_cancel.x() and c.btn_cover.y() == c.btn_ok.y()
    called = []
    c.on_cover = called.append
    c.btn_cover.click()
    assert called == [c] and c.result() == 0            # 갑지 버튼은 창을 닫지 않는다
    c.hide()
    for w in (d, again, c, again_cn):
        w.deleteLater()
    app.processEvents()
