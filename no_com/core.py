"""
core.py
=======
데이터 모델 / 상태 / 공통 유틸리티
UI에 일절 의존하지 않는다.
"""

from __future__ import annotations

import logging
import os
import sys
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any, Dict, List, Optional, Tuple

logger = logging.getLogger("QuoteApp")


# ──────────────────────────────────────────────
# 환경 유틸
# ──────────────────────────────────────────────
def exe_dir() -> str:
    """실행 파일(또는 스크립트)이 위치한 디렉터리."""
    if getattr(sys, "frozen", False):
        return str(Path(sys.executable).resolve().parent)
    return str(Path(__file__).resolve().parent)


def ensure_dir(path: str) -> str:
    os.makedirs(path, exist_ok=True)
    return path


def unique_path(path: str) -> str:
    """이미 존재하면 _2, _3 … 접미사를 붙여 반환."""
    if not os.path.exists(path):
        return path
    base, ext = os.path.splitext(path)
    n = 2
    while True:
        candidate = f"{base}_{n}{ext}"
        if not os.path.exists(candidate):
            return candidate
        n += 1


def safe_filename(text: str) -> str:
    """파일명으로 쓸 수 없는 문자를 '_'로 치환."""
    t = (text or "").strip()
    if not t:
        return "-"
    for ch in '\\/:*?"<>|':
        t = t.replace(ch, "_")
    return t.strip() or "-"


# ──────────────────────────────────────────────
# 문자열 / 숫자 변환 유틸
# ──────────────────────────────────────────────
def s(v: Any) -> str:
    return "" if v is None else str(v).strip()


def to_float(v: Any, default: float = 0.0) -> float:
    if v is None:
        return default
    if isinstance(v, (int, float)):
        return float(v)
    t = str(v).strip().replace(",", "")
    try:
        return float(t) if t else default
    except ValueError:
        return default


def fmt_qty(x: float) -> str:
    if abs(x - round(x)) < 1e-9:
        return str(int(round(x)))
    return f"{x:.3f}".rstrip("0").rstrip(".")


def fmt_krw(x: float) -> str:
    return f"{int(round(x)):,}"


def normalize_token(t: str) -> str:
    return s(t).upper().replace(" ", "")


def parse_invest_info(g_val: Any) -> str:
    """G열 값에서 '투자형태_모델명' 형식을 추출한다.

    예) 'DRY_PUMP;EQ,LOT,XD1200' → 'EQ_XD1200'
    세미콜론이 있으면 뒤 부분만 사용, 쉼표로 분리해 첫·마지막 토큰을 '_'로 결합.
    """
    t = s(g_val).strip()
    if not t:
        return ""
    if ";" in t:
        t = t.split(";", 1)[1].strip()
    parts = [p.strip() for p in t.split(",") if p.strip()]
    if not parts:
        return ""
    return parts[0] if len(parts) == 1 else f"{parts[0]}_{parts[-1]}"


# ──────────────────────────────────────────────
# 시트 / 셀 상수
# ──────────────────────────────────────────────
class SheetName:
    ITEMS       = "품목"
    CODE_MAP    = "코드매핑"
    REQ_COPY    = "견적의뢰복사본"
    SPEC        = "사양서"
    SIGN_SPEC   = "현업 사인용 사양서"
    INCOMING    = "입고검수확인서"
    QUOTE_CN    = "견적서_중국"
    QUOTE_US    = "견적서_미국"
    COVER_DATA  = "갑지DATA"
    COVER       = "견적서 갑지"

    REQUIRED = [
        ITEMS, CODE_MAP, REQ_COPY, SPEC,
        SIGN_SPEC, INCOMING, QUOTE_CN, QUOTE_US,
    ]
    ESIGN_TARGET = [COVER, SPEC, SIGN_SPEC, INCOMING]


@dataclass(frozen=True)
class DomesticLayout:
    SPEC_START  : int = 17
    SPEC_END    : int = 41
    SIGN_START  : int = 11
    SIGN_END    : int = 35
    COL_SPEC    : str = "B"
    COL_QTY     : str = "E"
    COL_PRICE   : str = "F"
    COL_AMT     : str = "G"


@dataclass(frozen=True)
class ChinaLayout:
    MAKER_CELL      : str = "A17"
    PROCESS_CELL    : str = "B17"
    TOOL_CELL       : str = "B51"
    PUMP_START      : int = 17
    PUMP_END        : int = 19
    RACK_START      : int = 21
    RACK_END        : int = 40
    COL_SPEC        : str = "C"
    COL_PRICE       : str = "F"
    COL_QTY         : str = "G"


@dataclass(frozen=True)
class USLayout:
    MAKER_CELL      : str = "A17"
    PROCESS_CELL    : str = "B17"
    TOOL_CELL       : str = "B54"
    EXCHANGE_CELL   : str = "A53"
    DATE_CELL       : str = "E53"
    PUMP_SPEC       : str = "C17"
    PUMP_PRICE      : str = "E17"
    PUMP_QTY        : str = "G17"
    RACK_START      : int = 21
    RACK_END        : int = 40
    COL_SPEC        : str = "C"
    COL_PRICE       : str = "E"
    COL_QTY         : str = "G"


DOM = DomesticLayout()
CN  = ChinaLayout()
US  = USLayout()


# ──────────────────────────────────────────────
# 앱 상태 (순수 데이터, UI 없음)
# ──────────────────────────────────────────────
@dataclass
class QuoteState:
    quote_type          : str                   = "국내"   # "국내" | "중국" | "미국"
    template_path       : Optional[str]         = None
    request_path        : Optional[str]         = None
    request_sheet_name  : Optional[str]         = None

    process : Optional[str] = None
    vendor  : Optional[str] = None
    code_5d : Optional[str] = None

    # 의뢰파일 행 목록: row_idx / D E F H J K V X Z
    request_rows : List[Dict[str, Any]] = field(default_factory=list)

    # 품목시트
    items_rows      : List[Dict[str, Any]]                              = field(default_factory=list)
    items_by_class  : Dict[str, List[Dict[str, Any]]]                   = field(default_factory=dict)
    price_by_spec   : Dict[str, float]                                  = field(default_factory=dict)

    # 코드매핑: (공정, 설비사, 5D) → [(ACC, 수량)]
    code_map : Dict[Tuple[str, str, str], List[Tuple[str, float]]] = field(default_factory=dict)

    # Credit
    pump_credit : float = 0.0
    rack_credit : float = 0.0

    # 중국/미국 팝업 입력값
    cn_info : Dict[str, str]  = field(default_factory=dict)
    us_info : Dict[str, Any]  = field(default_factory=dict)

    last_output_dir : Optional[str] = None
    investor_name   : str            = "채승철"
    warranty_years  : int            = 2

    # 국내 견적서 하드코딩 옵션 — 값이 있을 때만(0/빈 문자열이면 미기입) 셀에 반영된다.
    equip_qty   : int = 0    # 사양서!A43
    actual_line : str = ""   # 현업 사인용 사양서!B6
    exh_size    : int = 0    # 입고검수확인서!D27
