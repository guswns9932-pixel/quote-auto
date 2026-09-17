"""
app_settings.py
===============
사용자 설정 영속화 (QSettings 기반).

레지스트리(HKCU\\Software\\LOTVacuum\\QuoteAuto)에 저장되므로
EXE 폴더가 읽기 전용이어도 동작한다.

저장 대상
---------
  - 마지막 통합양식 / 의뢰파일 경로 (다음 실행 시 자동 복원)
  - 마지막 출력 폴더 (파일 대화상자 시작 위치)
  - 투자자명 / 보증기간
  - 창 위치·크기
"""

from __future__ import annotations

import logging
import os
from typing import Optional

from PySide6.QtCore import QByteArray, QSettings

logger = logging.getLogger("QuoteApp")

_ORG = "LOTVacuum"
_APP = "QuoteAuto"


def _st() -> QSettings:
    return QSettings(QSettings.IniFormat, QSettings.UserScope, _ORG, _APP)


# ──────────────────────────────────────────────
# 문자열 / 경로
# ──────────────────────────────────────────────
def get_str(key: str, default: str = "") -> str:
    try:
        v = _st().value(key, default)
        return str(v) if v is not None else default
    except Exception:
        logger.debug("설정 읽기 실패: %s", key, exc_info=True)
        return default


def set_str(key: str, value: Optional[str]) -> None:
    try:
        st = _st()
        if value:
            st.setValue(key, value)
        else:
            st.remove(key)
    except Exception:
        logger.debug("설정 쓰기 실패: %s", key, exc_info=True)


def get_path(key: str) -> str:
    """저장된 경로를 반환하되, 더 이상 존재하지 않으면 빈 문자열."""
    p = get_str(key)
    return p if p and os.path.exists(p) else ""


def get_dir(key: str, fallback: str = "") -> str:
    """저장된 폴더를 반환하되, 없으면 fallback."""
    p = get_str(key)
    return p if p and os.path.isdir(p) else fallback


# ──────────────────────────────────────────────
# 정수
# ──────────────────────────────────────────────
def get_int(key: str, default: int = 0) -> int:
    try:
        return int(_st().value(key, default))
    except (TypeError, ValueError):
        return default
    except Exception:
        logger.debug("설정 읽기 실패: %s", key, exc_info=True)
        return default


def set_int(key: str, value: int) -> None:
    try:
        _st().setValue(key, int(value))
    except Exception:
        logger.debug("설정 쓰기 실패: %s", key, exc_info=True)


# ──────────────────────────────────────────────
# 창 지오메트리
# ──────────────────────────────────────────────
def save_geometry(key: str, widget) -> None:
    try:
        _st().setValue(f"geometry/{key}", widget.saveGeometry())
    except Exception:
        logger.debug("지오메트리 저장 실패: %s", key, exc_info=True)


def restore_geometry(key: str, widget) -> bool:
    """저장된 지오메트리를 복원. 성공하면 True."""
    try:
        raw = _st().value(f"geometry/{key}")
        if isinstance(raw, QByteArray) and not raw.isEmpty():
            return bool(widget.restoreGeometry(raw))
    except Exception:
        logger.debug("지오메트리 복원 실패: %s", key, exc_info=True)
    return False


# ──────────────────────────────────────────────
# 키 상수
# ──────────────────────────────────────────────
class Key:
    TEMPLATE_PATH   = "paths/template"
    REQUEST_PATH    = "paths/request"
    OUTPUT_DIR      = "paths/output_dir"
    ESIGN_DIR       = "paths/esign_dir"
    AUTOLOAD_TEMPLATE = "quote/autoload_template"
    LAST_SEEN_UPDATE  = "app/last_seen_update"
