"""중복 견적 막기(호기·내용이 완전히 같을 때) + 파일 이름 끝의 PUMP 모델."""
import os

import pytest

import quote_config as qc


def _item(cat, spec, qty, price, role="ITEM", amt=None):
    return {"cat": cat, "role": role, "spec": spec, "qty": qty, "price": price,
            "amt": qty * price if amt is None else amt}


ITEMS = [_item("PUMP", "XD1200", 3, 13900000), _item("RACK", "Rack System-3", 1, 1000000)]
RD = {"row_idx": 2, "D": "PR1", "E": 10, "F": "Q890-044700", "G": "", "H": 3, "J": "의뢰자",
      "K": "P34D_ETCH", "Z": "H001"}


@pytest.fixture()
def env(quote_template, tmp_path, monkeypatch):
    import core
    import excel_io
    monkeypatch.setattr(excel_io, "exe_dir", lambda: str(tmp_path / "out"))
    st = core.QuoteState(template_path=quote_template)
    st.cn_info = {"line": "X2", "tool": "T1", "maker": "M", "process": "P"}
    st.us_info = {"site": "Taylor", "tool": "U1", "maker": "M", "process": "P", "exchange": 1400}
    return excel_io, st, tmp_path


def test_pump_model_at_the_end_of_file_names(env):
    io, st, tmp = env
    dom = os.path.basename(io.generate_quote(st, "국내", ITEMS, RD))
    assert dom.startswith("대외비_PR1-10_") and dom.endswith("_LOT베큠_P34D_ETCH_의뢰자_H001_XD1200.xlsx")
    assert os.path.basename(io.generate_quote(st, "중국", ITEMS, {})).endswith("_중국_SCS_X2_T1_XD1200.xlsx")
    assert os.path.basename(io.generate_quote(st, "미국", ITEMS, {})).endswith("_Taylor_U1_XD1200_Quotation.xlsx")
    # PUMP 가 둘이면 이어서, 없으면 붙이지 않는다 ('…_H001_.xlsx' 가 되지 않게)
    two = [_item("PUMP", "LD105 / LD120", 3, 1)] + ITEMS[1:]
    assert os.path.basename(io.generate_quote(st, "국내", two, RD)).endswith("_H001_LD105_LD120.xlsx")
    assert os.path.basename(io.generate_quote(st, "국내", ITEMS[1:], RD)).endswith("_의뢰자_H001.xlsx")
    assert io.pump_label([_item("PUMP", "DD1855H", 1, 1), _item("PUMP", "MK2000", 1, 1),
                          _item("RACK", "Rack System-3", 1, 1)]) == "DD1855H_MK2000"


def test_file_properties_and_quote_list_read_the_tool(env):
    """견적서 LIST: 파일 끝이 PUMP 모델이 되어도 호기를 제대로 읽는다(파일 속성)."""
    pytest.importorskip("PySide6.QtWidgets")
    import quote_fingerprint as qf
    io, st, tmp = env
    dom = io.generate_quote(st, "국내", ITEMS, RD)
    props = qf.read_props(dom)
    assert (props["tool"], props["line"], props["pump"]) == ("H001", "P34D_ETCH", "XD1200")
    assert len(props["quote_content"]) == 64
    os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")
    import page_quote
    tool, label = page_quote.QuoteBuilderPage._parse_quote_label(dom)
    assert tool == "H001" and label.endswith("_H001")
    us = io.generate_quote(st, "미국", ITEMS, {})
    assert page_quote.QuoteBuilderPage._parse_quote_label(us)[0] == "U1"
    # 예전 파일(속성 없음, 끝이 호기) 도 그대로 읽는다
    old = tmp / "out" / "견적서" / "260101_LOT베큠_L_J" / "대외비_P-1_260101_LOT베큠_L_J_OLD7.xlsx"
    old.parent.mkdir(parents=True); old.write_bytes(b"")
    assert page_quote.QuoteBuilderPage._parse_quote_label(str(old))[0] == "OLD7"


def test_duplicates_only_when_tool_and_content_are_identical(env):
    io, st, tmp = env
    assert io.find_duplicates(st, "국내", ITEMS, RD) == []
    made = io.generate_quote(st, "국내", ITEMS, RD)
    assert io.find_duplicates(st, "국내", ITEMS, RD) == [made]

    import copy
    different = [
        ("수량", [_item("PUMP", "XD1200", 3, 13900000), _item("RACK", "Rack System-3", 2, 1000000)], RD),
        ("단가", [_item("PUMP", "XD1200", 3, 13800000), ITEMS[1]], RD),
        ("의뢰 행", ITEMS, dict(RD, D="PR2")),
        ("호기", ITEMS, dict(RD, Z="H002")),
    ]
    for why, items, rd in different:
        assert io.find_duplicates(st, "국내", items, rd) == [], why
    st2 = copy.copy(st); st2.investor_name = "다른 사람"
    assert io.find_duplicates(st2, "국내", ITEMS, RD) == []
    st3 = copy.copy(st); st3.exh_size = 50
    assert io.find_duplicates(st3, "국내", ITEMS, RD) == []
    # 날짜는 내용이 아니다: 다른 날 만든 같은 견적서도 같은 것으로 본다
    assert io.find_duplicates(st, "국내", ITEMS, dict(RD, row_idx=99)) == [made]

    # 중국·미국은 옵션창 값까지 같아야 같은 것
    cn = io.generate_quote(st, "중국", ITEMS, {})
    assert io.find_duplicates(st, "중국", ITEMS, {}) == [cn]
    st.cn_info = dict(st.cn_info, quote_no="Q-1")
    assert io.find_duplicates(st, "중국", ITEMS, {}) == []


def test_batch_generation_writes_each_rows_fingerprint(env):
    io, st, tmp = env
    rds = [dict(RD, Z="B1"), dict(RD, Z="B2")]
    res = io.generate_quote_multi(st, ITEMS, rds)
    assert [e for _, _, e in res] == ["", ""]
    for rd, (_, path, _) in zip(rds, res):
        assert io.find_duplicates(st, "국내", ITEMS, rd) == [path]       # 재사용 워크북이라도 건마다 자기 지문


def test_path_too_long_is_refused_before_anything_is_written(env, monkeypatch):
    io, st, tmp = env
    monkeypatch.setattr(io, "exe_dir", lambda: str(tmp / ("긴폴더이름" * 40)))
    with pytest.raises(ValueError, match="경로가 너무 깁니다"):
        io.generate_quote(st, "국내", ITEMS, RD)
    assert not os.path.exists(tmp / ("긴폴더이름" * 40))


def test_page_asks_before_making_a_duplicate(env, monkeypatch):
    pytest.importorskip("PySide6.QtWidgets")
    os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")
    import gc
    from PySide6.QtCore import QSettings
    from PySide6.QtWidgets import QApplication, QMessageBox
    io, st, tmp = env
    QSettings.setPath(QSettings.IniFormat, QSettings.UserScope, str(tmp / "settings"))
    app = QApplication.instance() or QApplication([])
    import page_quote
    page = page_quote.QuoteBuilderPage()
    page.state.template_path = st.template_path
    made = io.generate_quote(st, "국내", ITEMS, RD)

    chosen = {"text": None}

    def fake_exec(box):
        texts = [b.text() for b in box.buttons()]
        chosen["buttons"] = texts
        chosen["body"] = box.text()
        target = next(b for b in box.buttons() if b.text() == chosen["text"])
        box._clicked = target
        return 0
    monkeypatch.setattr(QMessageBox, "exec", fake_exec)
    monkeypatch.setattr(QMessageBox, "clickedButton", lambda box: box._clicked)
    monkeypatch.setattr(QMessageBox, "information", lambda *a, **k: QMessageBox.Ok)

    # 한 건: 그래도 만들기 / 취소
    chosen["text"] = "취소"
    assert page._ask_about_duplicates("국내", ITEMS, [RD], [0]) == ([], [])
    assert chosen["buttons"] == ["그래도 만들기", "취소"] and os.path.basename(made) in chosen["body"]
    chosen["text"] = "그래도 만들기"
    assert page._ask_about_duplicates("국내", ITEMS, [RD], [0]) == ([RD], [0])

    # 여러 건: 중복은 건너뛰기 / 중복도 모두 만들기 / 취소
    rds = [RD, dict(RD, Z="NEW1")]
    chosen["text"] = "중복은 건너뛰기"
    assert page._ask_about_duplicates("국내", ITEMS, rds, [0, 1]) == ([rds[1]], [1])
    assert "2건 중 1건" in chosen["body"]
    chosen["text"] = "중복도 모두 만들기"
    assert page._ask_about_duplicates("국내", ITEMS, rds, [0, 1]) == (rds, [0, 1])
    chosen["text"] = "취소"
    assert page._ask_about_duplicates("국내", ITEMS, rds, [0, 1]) == ([], [])
    # 중복이 없으면 묻지 않는다
    chosen["buttons"] = None
    assert page._ask_about_duplicates("국내", ITEMS, [dict(RD, Z="NEW2")], [0]) == ([dict(RD, Z="NEW2")], [0])
    assert chosen["buttons"] is None
    page.close(); page.deleteLater(); app.processEvents(); del page; gc.collect()


def test_name_parser_reads_new_and_old_names():
    import excel_io
    cfg = qc.QuoteConfig()
    assert excel_io.parse_file_name(cfg, "name.cn_file", "대외비_261123_중국_SCS_X2_T1_LD105_LD120_2.xlsx") == \
        {"날짜": "261123", "라인": "X2", "호기": "T1", "PUMP": "LD105_LD120"}
    assert excel_io.parse_file_name(cfg, "name.cn_file", "대외비_261123_중국_SCS_X2_T1.xlsx") == \
        {"날짜": "261123", "라인": "X2", "호기": "T1"}
    assert excel_io.parse_file_name(cfg, "name.us_file", "대외비_261123_S_U1_Quotation.xlsx")["호기"] == "U1"
    assert excel_io.parse_file_name(cfg, "name.cn_file", "전혀다른이름.xlsx") is None
