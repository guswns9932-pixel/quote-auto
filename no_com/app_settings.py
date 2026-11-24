"""
app_settings.py
===============
사용자 설정 영속화 (QSettings 기반).

사용자 폴더의 ini(%APPDATA%\\LOTVacuum\\QuoteAuto.ini)에 저장되므로
EXE 폴더가 읽기 전용이어도 동작한다.

%APPDATA% 는 로밍 폴더라 회사 PC 에서는 같은 계정의 다른 PC 와 공유될 수 있다.
PC 마다 다른 값(통합양식 위치 등)은 get_pc_* / set_pc_* 로 PC 이름별로 따로 둔다.

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
import platform
import re
from typing import Optional

from PySide6.QtCore import QByteArray, QSettings

logger = logging.getLogger("QuoteApp")

_ORG = "LOTVacuum"
_APP = "QuoteAuto"

# 프로그램 이름·버전 — 런처·스플래시·각 앱 창 제목에 쓴다. 배포할 때마다 버전을 올린다.
APP_TITLE = "견적업무 자동화"
APP_VERSION = "1.0"
APP_TITLE_FULL = f"{APP_TITLE} ver.{APP_VERSION}"


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
# PC 별 값 — 같은 계정이라도 PC 가 다르면 따로 기억한다
# ──────────────────────────────────────────────
def pc_name() -> str:
    """이 PC 의 이름 (설정 키에 쓸 수 있게 정리). 알 수 없으면 빈 문자열."""
    name = os.environ.get("COMPUTERNAME") or platform.node() or ""
    return re.sub(r"[^0-9A-Za-z가-힣_.-]", "_", name.strip()).upper()


def _pc_key(key: str) -> str:
    name = pc_name()
    return f"pc/{name}/{key}" if name else key


def get_pc_str(key: str) -> str:
    """이 PC 에서 저장한 값. 아직 없으면(업데이트 직후 첫 실행) 예전 공용 값을 쓴다."""
    return get_str(_pc_key(key)) or get_str(key)


def set_pc_str(key: str, value: Optional[str]) -> None:
    set_str(_pc_key(key), value)


def get_pc_path(key: str) -> str:
    """이 PC 에서 저장한 경로를 반환하되, 더 이상 존재하지 않으면 빈 문자열.
    (네트워크 드라이브가 아직 안 붙은 경우 등 — 다른 PC 의 경로로 대신하지 않는다)"""
    p = get_pc_str(key)
    return p if p and os.path.exists(p) else ""


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
    TEMPLATE_PATH   = "paths/template"    # PC 별 (get_pc_path / set_pc_str)
    REQUEST_PATH    = "paths/request"
    OUTPUT_DIR      = "paths/output_dir"
    ESIGN_DIR       = "paths/esign_dir"
    # v2: 기본 크기를 바꾸면서 키를 새로 했다(예전 값에 다시 곱해지지 않도록)
    ESIGN_SIGN_PCT  = "esign/sign_size_pct_v2"
    AUTOLOAD_TEMPLATE = "quote/autoload_template"
    LAST_SEEN_UPDATE  = "app/last_seen_update"
