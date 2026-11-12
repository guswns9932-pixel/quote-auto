"""
app_boot.py
===========
모든 진입점이 공유하는 최소 부트스트랩 (로그 설정 + Tk 앱 실행기).
PySide6 를 임포트하지 않는다 — Tk 앱(키워드 검색기·CSP UPLOAD)이 Qt 를
끌어오지 않아야 첫 실행이 빠르다.
"""

from __future__ import annotations

import logging
import logging.handlers
import os
import sys
import traceback

# ──────────────────────────────────────────────
# 로거 초기화
# ──────────────────────────────────────────────
_LOG_MAX_BYTES = 512 * 1024


def _trim_oversized_log(path: str) -> None:
    """이전 버전이 무한 append 로 키운 로그가 남아 있으면 직전본 자리로 밀어낸다
    (RotatingFileHandler 는 다음 기록 때에야 회전하므로, 시작 시 한 번 정리)."""
    try:
        if os.path.getsize(path) > _LOG_MAX_BYTES:
            os.replace(path, path + ".1")
    except OSError:
        pass


def _setup_logging() -> None:
    root = logging.getLogger("QuoteApp")
    root.setLevel(logging.DEBUG)
    if not root.handlers:
        # console=False 빌드는 sys.stdout 이 None 이라 핸들러가 레코드마다 예외를 삼킨다.
        if sys.stdout is not None:
            h = logging.StreamHandler(sys.stdout)
            h.setFormatter(logging.Formatter("[%(levelname)s][%(name)s] %(message)s"))
            root.addHandler(h)

        # 파일 로그 — console=False 빌드에서도 충돌 원인 추적 가능
        try:
            _exe_dir = (
                os.path.dirname(os.path.abspath(sys.executable))
                if getattr(sys, "frozen", False)
                else os.path.dirname(os.path.abspath(__file__))
            )
            # 클라우드(공유 서버)에서 로그가 무한히 쌓여 자원을 먹지 않도록
            # 총량에 상한을 둔다: 파일 1개 최대 _LOG_MAX_BYTES, 직전본 1개만 보관
            # (합계 ≤ 2배). 평소엔 INFO 이상만 남기고, DEBUG 까지 필요하면
            # 환경변수 QUOTE_APP_DEBUG_LOG=1. delay=True 라 기록할 것이 생기기
            # 전에는 파일을 만들지도 않는다.
            log_path = os.path.join(_exe_dir, "quote_app.log")
            _trim_oversized_log(log_path)
            fh = logging.handlers.RotatingFileHandler(
                log_path, maxBytes=_LOG_MAX_BYTES, backupCount=1,
                encoding="utf-8", delay=True,
            )
            fh.setLevel(logging.DEBUG if os.environ.get("QUOTE_APP_DEBUG_LOG") == "1"
                        else logging.INFO)
            fh.setFormatter(logging.Formatter(
                "[%(levelname)s][%(asctime)s][%(name)s] %(message)s",
                datefmt="%Y-%m-%d %H:%M:%S",
            ))
            root.addHandler(fh)
        except Exception:
            pass


logger = logging.getLogger("QuoteApp")


def _setup_exception_hook() -> None:
    """메인 스레드 uncaught exception → 로그 + 대화상자 (silent crash 방지)."""
    def _hook(exc_type, exc_value, exc_tb):
        msg = "".join(traceback.format_exception(exc_type, exc_value, exc_tb))
        logging.getLogger("QuoteApp").critical("치명적 오류:\n%s", msg)
        try:
            from PySide6.QtWidgets import QMessageBox
            QMessageBox.critical(
                None, "치명적 오류",
                f"예상치 못한 오류가 발생했습니다:\n\n{exc_value}\n\n"
                "프로그램 폴더의 quote_app.log 에서 자세한 내용을 확인하세요.",
            )
        except Exception:
            pass
    sys.excepthook = _hook


def _setup_exception_hook_tk(app_label: str = "Tk 앱") -> None:
    """Tk 전용 실행 경로 — QApplication 이 없으므로 QMessageBox 대신 tkinter로."""
    def _hook(exc_type, exc_value, exc_tb):
        msg = "".join(traceback.format_exception(exc_type, exc_value, exc_tb))
        logging.getLogger("QuoteApp").critical("%s 치명적 오류:\n%s", app_label, msg)
        try:
            import tkinter.messagebox as mb
            mb.showerror("치명적 오류", f"예상치 못한 오류가 발생했습니다:\n\n{exc_value}")
        except Exception:
            pass
    sys.excepthook = _hook


def run_tk_app(app_label: str, factory) -> None:
    """Tk 앱 실행. factory() 는 mainloop() 를 가진 Tk 객체를 돌려준다."""
    _setup_logging()
    _setup_exception_hook_tk(app_label)
    logger.info("%s 시작", app_label)
    factory().mainloop()
