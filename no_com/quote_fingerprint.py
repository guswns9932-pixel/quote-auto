"""
quote_fingerprint.py
====================
견적서의 '내용 지문' — 호기와 내용이 완전히 같은 견적서를 두 번 만들지 않도록.

지문에 넣는 것: 견적서 종류, 의뢰 행의 값, 견적 품목(규격·수량·단가·금액·Credit),
옵션(투자자·보증기간·설비수량·실반입라인·Exh Size·중국/미국 옵션창 값).
넣지 않는 것: 만든 날짜, 양식 파일 자체(단가는 품목에 들어 있으므로 단가가 바뀌면 다른 내용).

생성한 견적서의 파일 속성(사용자 지정 속성)에 숨겨 저장한다. 화면·인쇄에는 보이지 않는다.
"""

from __future__ import annotations

import hashlib
import json
import zipfile
from datetime import date, datetime
from typing import Any, Dict, List, Optional

PROP_NAME = "LOT_quote_content"

# 의뢰 행에서 견적서 내용에 들어가는 칸 (row_idx 같은 위치 정보는 빼고)
_RD_KEYS = ("D", "E", "F", "G", "H", "J", "K", "N", "R", "V", "X", "Y", "Z", "AA")


def _plain(v: Any) -> Any:
    if isinstance(v, (datetime, date)):
        return v.isoformat()
    if isinstance(v, float):
        return int(v) if v.is_integer() else round(v, 6)
    if isinstance(v, (list, tuple)):
        return [_plain(x) for x in v]
    if isinstance(v, dict):
        return {str(k): _plain(x) for k, x in sorted(v.items(), key=lambda kv: str(kv[0]))}
    if v is None:
        return ""
    return v if isinstance(v, (int, str, bool)) else str(v)


def make(qtype: str, rd: Dict[str, Any], items: List[Dict[str, Any]], state) -> str:
    content = {
        "type": qtype,
        "request": {k: _plain(rd.get(k)) for k in _RD_KEYS} if qtype == "국내" else {},
        "items": [{k: _plain(r.get(k)) for k in ("role", "cat", "spec", "qty", "price", "amt")} for r in items],
        "options": {
            "investor": _plain(state.investor_name), "warranty": _plain(state.warranty_years),
            "equip_qty": _plain(state.equip_qty), "actual_line": _plain(state.actual_line),
            "exh_size": _plain(state.exh_size),
            "cn": _plain(state.cn_info or {}) if qtype == "중국" else {},
            "us": _plain(state.us_info or {}) if qtype == "미국" else {},
        },
    }
    text = json.dumps(content, ensure_ascii=False, sort_keys=True, separators=(",", ":"))
    return hashlib.sha256(text.encode("utf-8")).hexdigest()


META_PREFIX = "LOT_"          # 함께 적어 두는 정보: LOT_tool(호기), LOT_line(라인), LOT_pump(PUMP 모델)


def write(wb, fingerprint: str, **meta: Any) -> None:
    """워크북의 사용자 지정 속성에 지문과 정보(호기·라인·PUMP)를 적는다.
    이미 있으면 바꾼다 — 일괄 생성은 같은 워크북을 재사용한다.

    호기·라인을 파일 이름에서 다시 읽으면, 라인처럼 '_' 가 들어간 값 때문에 어디서
    끊어야 할지 모호하다. 그래서 만들 때 값을 그대로 적어 둔다."""
    from openpyxl.packaging.custom import StringProperty
    props = wb.custom_doc_props
    props.props = [p for p in props.props if not p.name.startswith(META_PREFIX)]
    props.append(StringProperty(name=PROP_NAME, value=fingerprint))
    for key, value in meta.items():
        text = "" if value is None else str(value)
        if text:
            props.append(StringProperty(name=META_PREFIX + key, value=text))


def read_props(path: str) -> Dict[str, str]:
    """xlsx 에 적어 둔 LOT_* 속성들 {이름(LOT_ 뺀 것): 값}. 파일을 통째로 열지 않고 속성 부분만 읽는다."""
    from xml.etree import ElementTree as ET
    try:
        with zipfile.ZipFile(path) as z:
            if "docProps/custom.xml" not in z.namelist():
                return {}
            root = ET.fromstring(z.read("docProps/custom.xml"))
    except (OSError, zipfile.BadZipFile, ET.ParseError):
        return {}
    out: Dict[str, str] = {}
    for prop in root:
        name = prop.get("name") or ""
        if not name.startswith(META_PREFIX):
            continue
        value = next(iter(prop), None)              # <vt:lpwstr> 등 값 요소 하나
        out[name[len(META_PREFIX):]] = (value.text or "") if value is not None else ""
    return out


def read(path: str) -> Optional[str]:
    """xlsx 의 지문. 없으면(이 기능 이전에 만든 파일 등) None."""
    return read_props(path).get(PROP_NAME[len(META_PREFIX):])
