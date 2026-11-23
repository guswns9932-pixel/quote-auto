"""그림 보존 안전장치 — Pillow 가 없으면 openpyxl 이 그림을 조용히 버리는 문제.
(빌드에서 Pillow 를 뺐다가 로고가 빠진 견적서가 실제로 나갔다)"""
import os
import shutil

import pytest
from openpyxl import load_workbook


def _item(cat, spec, qty, price):
    return {"cat": cat, "role": "ITEM", "spec": spec, "qty": qty, "price": price, "amt": qty * price}


ITEMS = [_item("PUMP", "P", 1, 100), _item("RACK", "R", 1, 10)]
RD = {"row_idx": 2, "D": 1, "E": 10, "F": "Q", "H": 1, "J": "a", "K": "b", "Z": "c"}


@pytest.fixture()
def env(quote_template, tmp_path, monkeypatch):
    import core
    import excel_io
    monkeypatch.setattr(excel_io, "exe_dir", lambda: str(tmp_path / "out"))
    st = core.QuoteState(template_path=quote_template)
    st.cn_info = {"line": "L", "tool": "T", "maker": "M", "process": "P"}
    st.us_info = {"site": "S", "tool": "T", "maker": "M", "process": "P", "exchange": 1400}
    return excel_io, st, tmp_path


def _files(root):
    return [os.path.join(r, f) for r, _, fs in os.walk(root) for f in fs]


def test_every_output_keeps_the_template_images(env):
    """정상 환경: 국내·중국·미국·갑지 모두 양식의 그림 수 그대로."""
    io, st, tmp = env
    want = io.count_sheet_images(st.template_path)
    assert want["견적서_중국"] == 3 and want["견적서_미국"] == 2 and want["사양서"] == 1 and want["견적서 갑지"] == 3
    dom = io.generate_quote(st, "국내", ITEMS, RD)
    outs = {"국내": dom, "중국": io.generate_quote(st, "중국", ITEMS, {}),
            "미국": io.generate_quote(st, "미국", ITEMS, {})}
    outs["갑지"], failed = io.generate_cover(st.template_path, os.path.dirname(dom), [dom])
    for kind, path in outs.items():
        got = io.count_sheet_images(path)
        assert got and all(got[name] == want[name] for name in got), kind
    assert io.count_sheet_images(outs["중국"]) == {"견적서_중국": 3}


def test_without_pillow_nothing_is_generated(env, monkeypatch):
    """① Pillow 를 쓸 수 없으면 시작하지 않는다 — 그림 없는 파일을 만들지 않는다."""
    io, st, tmp = env
    import openpyxl.drawing.image as img
    monkeypatch.setattr(img, "PILImage", False)
    for call in (lambda: io.generate_quote(st, "국내", ITEMS, RD),
                 lambda: io.generate_quote(st, "중국", ITEMS, {}),
                 lambda: io.generate_quote(st, "미국", ITEMS, {}),
                 lambda: io.generate_quote_multi(st, ITEMS, [RD]),
                 lambda: io.generate_cover(st.template_path, str(tmp), [st.template_path]),
                 lambda: io.export_cover_data_sheet(st.template_path, st.template_path)):
        with pytest.raises(RuntimeError, match="Pillow"):
            call()
    assert _files(tmp / "out") == []


def test_lost_images_are_caught_after_saving(env, monkeypatch):
    """② 시작 전 확인을 통과했더라도, 저장한 파일에서 그림이 빠졌으면 오류로 멈추고 파일을 남기지 않는다."""
    io, st, tmp = env
    import openpyxl.reader.drawings as reader
    monkeypatch.setattr(reader, "PILImage", False)       # 읽을 때 그림이 버려지는 상황만 만든다
    for qtype, sheet in (("중국", "견적서_중국"), ("미국", "견적서_미국"), ("국내", "사양서")):
        with pytest.raises(RuntimeError, match="그림이 빠졌습니다") as e:
            io.generate_quote(st, qtype, ITEMS, RD)
        assert sheet in str(e.value)
    assert _files(tmp / "out") == []
    res = io.generate_quote_multi(st, ITEMS, [RD])        # 일괄 생성은 건별로 실패를 돌려준다
    assert res[0][1] == "" and "그림이 빠졌습니다" in res[0][2] and _files(tmp / "out") == []


def test_code_map_save_never_leaves_a_template_without_images(env, monkeypatch):
    """'5D Object 생성' 은 통합양식 자체를 다시 저장한다 — 그림이 빠지면 양식을 되돌린다."""
    io, st, tmp = env
    tpl = str(tmp / "양식.xlsx")
    shutil.copy(st.template_path, tpl)
    before = io.count_sheet_images(tpl)

    # 정상: 저장 후에도 그림 수가 같다
    added, _skipped = io.append_code_map_rows(tpl, "P1", "V1", "5D-1", [("ACC-1", 1.0)])
    assert added and io.count_sheet_images(tpl) == before

    # Pillow 를 쓸 수 없으면 양식을 건드리지 않는다
    import openpyxl.drawing.image as img
    import openpyxl.reader.drawings as reader
    stamp = os.path.getmtime(tpl)
    monkeypatch.setattr(img, "PILImage", False)
    with pytest.raises(RuntimeError, match="Pillow"):
        io.append_code_map_rows(tpl, "P1", "V1", "5D-2", [("ACC-2", 1.0)])
    assert os.path.getmtime(tpl) == stamp
    monkeypatch.undo()

    # 저장 과정에서 그림이 빠졌다면 저장 전 사본으로 되돌린다
    monkeypatch.setattr(reader, "PILImage", False)
    with pytest.raises(RuntimeError, match="그림이 빠졌습니다"):
        io.append_code_map_rows(tpl, "P1", "V1", "5D-3", [("ACC-3", 1.0)])
    monkeypatch.undo()
    assert io.count_sheet_images(tpl) == before
    ws = load_workbook(tpl, read_only=True)["코드매핑"]
    codes = {r[2] for r in ws.iter_rows(min_col=1, max_col=5, values_only=True)}
    assert "5D-1" in codes and "5D-3" not in codes       # 되돌려졌으므로 실패한 저장분은 없다
