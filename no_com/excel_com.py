"""
excel_com.py
============
Excel COM(pywin32) 공통 헬퍼 — 지연 로딩, 클립보드 정리, 세션 컨텍스트.
견적서작성(갑지 생성)과 전자서명(시트 캡처)이 함께 쓴다.
openpyxl·Qt 에 의존하지 않는다.
"""

from __future__ import annotations

import logging
import time
from typing import Any, Optional

logger = logging.getLogger("QuoteApp")


def _clear_clipboard() -> None:
    """
    Windows 클립보드를 강제로 비운다.

    CopyPicture()/ExportAsFixedFormat() 후 COM 세션이 종료되면
    클립보드에 죽은 프로세스 소유의 CF_ENHMETAFILE(EMF) 핸들이 남는다.
    이 stale 핸들이 있는 상태에서 별도 Excel 창이 Ctrl+C를 시도하면
    Windows가 이전 클립보드를 초기화하다가 액세스 위반 → Excel 전체 종료.
    COM 세션 종료 전, 또는 CopyPicture 직후에 반드시 호출한다.
    """
    try:
        import ctypes
        u32 = ctypes.windll.user32
        if u32.OpenClipboard(0):
            u32.EmptyClipboard()
            u32.CloseClipboard()
    except Exception:
        pass


# ── COM 지연 로딩 (PDF 변환 전용) ───────────────────────────────
# 모듈 임포트 시 DLL LoadLibrary를 호출하지 않도록 지연.
# GIL 점유로 인한 시작 시 UI 동결 방지 (AV/Defender 환경에서 10~120초 블로킹 가능).
pythoncom: Any = None
win32: Any = None
COM_AVAILABLE: Optional[bool] = None  # None=미확인, True=사용가능, False=불가


def _ensure_com() -> bool:
    """COM DLL을 최초 필요 시점에만 로드한다."""
    global COM_AVAILABLE, pythoncom, win32
    if COM_AVAILABLE is not None:
        return COM_AVAILABLE
    try:
        import pythoncom as _pc
        import win32com.client as _w32
        pythoncom = _pc
        win32 = _w32
        COM_AVAILABLE = True
    except Exception:
        pythoncom = None
        win32 = None
        COM_AVAILABLE = False
    return COM_AVAILABLE


class ExcelCOM:
    """
    Excel COM 세션 컨텍스트 매니저.
    _ExcelWorkerThread(전자서명 시트 캡처)가 공유 Application 객체로 사용한다.
    그 워커는 페이지가 열려 있는 내내 이 세션을 쥐고 있다 — COM 객체는
    만들어진 스레드에서만 안전하므로 생성·사용·종료가 모두 한 스레드다.
    """

    def __init__(self) -> None:
        self._excel = None

    def __enter__(self) -> "ExcelCOM":
        if not _ensure_com():
            raise RuntimeError("pywin32가 설치되어 있지 않습니다.")
        pythoncom.CoInitialize()
        _t0 = time.perf_counter()
        try:
            self._excel = win32.DispatchEx("Excel.Application")
            logger.info("Excel COM 세션 시작: %.0fms", (time.perf_counter() - _t0) * 1000)
            self._excel.Visible = False
            self._excel.DisplayAlerts = False
            # ★ 자동화 스위치(Calculation / ScreenUpdating)는 걸지 않는다 ★
            #
            # 예전에는 걸려고 시도하는 코드가 있었지만, 첫 줄의 Calculation
            # 설정이 com_error 로 실패하면서 뒤의 것들까지 통째로 건너뛰고
            # 있었다. 그 "버그"를 고쳐 실제로 스위치를 걸었더니 캡처가 빈
            # 이미지로 나왔다 — 실패하던 쪽이 오히려 이 앱을 지켜주고 있었다.
            #
            #   · ScreenUpdating=False
            #     CopyPicture(Appearance=xlScreen) 는 "화면에 그려진 모습"을
            #     복사하는 것이라, 화면 갱신을 끄면 빈 그림이 복사된다.
            #   · Calculation=xlCalculationManual
            #     이 앱이 만드는 견적서는 openpyxl 로 저장돼 수식의 캐시값이
            #     없다. 열 때 재계산(fullCalcOnLoad)이 돌아야 값이 보이는데,
            #     수동으로 돌려놓으면 수식 칸이 빈 채로 캡처된다.
            #
            # 즉 이 경로에서 "열기가 느린 것"(재계산)과 "캡처가 제대로 나오는
            # 것"은 맞바꿀 수 없다. 속도를 줄이려면 재계산량 자체를 줄여야
            # 한다(생성물에서 수식·시트를 덜어내는 방향).
        except BaseException:
            # DispatchEx 실패(Excel 미설치·손상) 시 __exit__ 가 호출되지 않으므로
            # 여기서 아파트먼트를 되돌리지 않으면 호출할 때마다 하나씩 누수된다.
            self._excel = None
            try:
                pythoncom.CoUninitialize()
            except Exception:
                pass
            raise
        return self

    def __exit__(self, *_) -> None:
        # 순서 중요: 프록시 해제 → Quit() → 참조 제거 → CoUninitialize()
        # 반대로 하면 보이지 않는 EXCEL.EXE 가 파일 잠금을 쥔 채 남는다.
        import gc
        gc.collect()
        if self._excel is not None:
            try:
                self._excel.Quit()
            except Exception:
                pass
            self._excel = None
            gc.collect()
        try:
            pythoncom.CoUninitialize()
        except Exception:
            pass

    @property
    def app(self):
        return self._excel

    def open(self, path: str, read_only: bool = False):
        return self._excel.Workbooks.Open(
            path, ReadOnly=read_only, UpdateLinks=0, AddToMru=False
        )
