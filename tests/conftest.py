"""테스트 공통 준비.

통합양식은 대외비라 저장소에 없다. 아래 폴더(기본 tests/_local, 환경변수
QUOTE_AUTO_TEMPLATE_DIR 로 변경)에 xlsx 를 넣어 두면 파일 이름과 상관없이
시트 구성을 보고 견적서 양식 / CSP 양식을 알아서 찾는다. 없으면 양식이
필요한 테스트는 건너뛴다.
"""
import hashlib
import json
import os
import sys

import pytest

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, os.path.join(ROOT, "no_com"))

LOCAL_DIR = os.environ.get("QUOTE_AUTO_TEMPLATE_DIR") or os.path.join(ROOT, "tests", "_local")
BASELINE_PATH = os.path.join(ROOT, "tests", "baseline.json")
UPDATE = os.environ.get("UPDATE_BASELINE") == "1"


def _sha256(path):
    h = hashlib.sha256()
    with open(path, "rb") as f:
        for chunk in iter(lambda: f.read(1 << 20), b""):
            h.update(chunk)
    return h.hexdigest()


def _find_templates():
    found = {}
    if not os.path.isdir(LOCAL_DIR):
        return found
    from openpyxl import load_workbook
    for name in sorted(os.listdir(LOCAL_DIR)):
        if not name.lower().endswith(".xlsx") or name.startswith("~$"):
            continue
        path = os.path.join(LOCAL_DIR, name)
        try:
            wb = load_workbook(path, read_only=True)
            sheets = set(wb.sheetnames)
            wb.close()
        except Exception:
            continue
        if {"사양서", "견적의뢰복사본", "품목"} <= sheets:
            found.setdefault("quote", path)
        elif {"Sheet1", "FSC매핑", "CIP"} <= sheets:
            found.setdefault("csp", path)
    return found


_TEMPLATES = _find_templates()


def _template(kind):
    path = _TEMPLATES.get(kind)
    if not path:
        pytest.skip("통합양식(%s)이 %s 에 없습니다" % (kind, LOCAL_DIR))
    return path


@pytest.fixture(scope="session")
def quote_template():
    return _template("quote")


@pytest.fixture(scope="session")
def csp_template():
    return _template("csp")


class Baseline:
    """출력물 지문(fingerprint)의 해시를 기준값과 비교한다.

    저장소에는 해시만 올린다(지문에는 단가가 들어 있어 대외비).
    지문 전문은 tests/_local/fingerprints 에 남겨, 어긋났을 때 이전 커밋에서
    만든 것과 diff 해 볼 수 있게 한다.
      UPDATE_BASELINE=1 pytest   → 기준값을 지금 결과로 다시 쓴다
    """

    def __init__(self):
        self.data = {"templates": {}, "cases": {}}
        if os.path.exists(BASELINE_PATH):
            with open(BASELINE_PATH, encoding="utf-8") as f:
                self.data = json.load(f)
        self.dirty = False

    def check(self, case, kind, fingerprint):
        tpl_hash = _sha256(_TEMPLATES[kind])
        text = json.dumps(fingerprint, ensure_ascii=False, sort_keys=True, indent=1, default=str)
        digest = hashlib.sha256(text.encode("utf-8")).hexdigest()
        fp_dir = os.path.join(LOCAL_DIR, "fingerprints")
        os.makedirs(fp_dir, exist_ok=True)
        with open(os.path.join(fp_dir, case + ".json"), "w", encoding="utf-8") as f:
            f.write(text)
        if UPDATE:
            self.data["templates"][kind] = tpl_hash
            self.data["cases"][case] = digest
            self.dirty = True
            return
        if self.data["templates"].get(kind) != tpl_hash:
            pytest.skip("기준값을 만든 양식과 다른 파일입니다 — UPDATE_BASELINE=1 로 기준을 다시 만드세요")
        want = self.data["cases"].get(case)
        if want is None:
            pytest.fail("기준값이 없습니다(%s) — UPDATE_BASELINE=1 로 만드세요" % case)
        assert digest == want, (
            "출력이 기준과 다릅니다: %s\n지문 전문: %s\n"
            "(의도한 변경이면 UPDATE_BASELINE=1 로 기준을 갱신)" % (case, os.path.join(fp_dir, case + ".json")))

    def save(self):
        if self.dirty:
            with open(BASELINE_PATH, "w", encoding="utf-8") as f:
                json.dump(self.data, f, ensure_ascii=False, sort_keys=True, indent=1)
                f.write("\n")


@pytest.fixture(scope="session")
def baseline():
    b = Baseline()
    yield b
    b.save()


def xlsx_fingerprint(path):
    """xlsx 한 개의 '내용'을 비교 가능한 dict 로 만든다.
    시트 순서·표시 상태, 모든 칸의 값(수식은 수식 글자)·표시 형식, 숨긴 행,
    그림 개수까지 담는다. 파일 이름·저장 시각처럼 매번 바뀌는 것은 넣지 않는다."""
    import zipfile
    from openpyxl import load_workbook
    wb = load_workbook(path)
    out = {"sheets": []}
    for ws in wb.worksheets:
        cells = {}
        for row in ws.iter_rows():
            for c in row:
                v = c.value
                if v is None:
                    continue
                # 배열 수식은 객체라 그대로 찍으면 메모리 주소가 섞인다 → 범위와 글자로
                text = "%s|%s" % (v.ref, v.text) if hasattr(v, "ref") and hasattr(v, "text") else str(v)
                cells[c.coordinate] = [type(v).__name__, text, c.number_format]
        hidden = sorted(r for r, d in ws.row_dimensions.items() if d.hidden)
        out["sheets"].append({"name": ws.title, "state": ws.sheet_state, "cells": cells,
                              "hidden_rows": hidden, "print_area": ws.print_area})
    wb.close()
    with zipfile.ZipFile(path) as z:
        names = z.namelist()
    out["media"] = sum(1 for n in names if n.startswith("xl/media/"))
    out["vml"] = sum(1 for n in names if "vmlDrawing" in n)
    return out
