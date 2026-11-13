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

    c = page_quote.ChinaOptionsDialog(None, st)
    c._quote_no.setText("CN-01"); c._delivery.setText("2026-12-20"); c._to.setText("Wang")
    c.apply()
    assert (st.cn_info["quote_no"], st.cn_info["delivery"], st.cn_info["to"]) == ("CN-01", "2026-12-20", "Wang")
    assert page_quote.ChinaOptionsDialog(None, st)._delivery.text() == "2026-12-20"
    for w in (d, again, c):
        w.deleteLater()
    app.processEvents()
