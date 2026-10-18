"""
app_runner.py
=============
Qt 페이지 앱(견적서작성 / 전자서명)이 공유하는 단독 실행기 + 페이지 창.

런처(main.py)는 이 모듈을 임포트하지 않는다. 각 앱은 자기 exe 로 따로 뜨고
자기에게 필요한 모듈만 임포트한다 — 첫 실행이 느린 원인이던 "런처가 모든 앱의
의존성을 한 프로세스에 싣는 구조"를 끊기 위한 분리.
"""

from __future__ import annotations

import logging
import sys
from typing import Optional

from PySide6.QtCore import Qt
from PySide6.QtWidgets import (
    QApplication, QMainWindow, QMessageBox, QPushButton, QSizePolicy, QWidget,
)

import app_settings
from app_boot import _setup_exception_hook, _setup_logging

logger = logging.getLogger("QuoteApp")


def run_qt_page(title: str, mod_name: str, cls_name: str) -> None:
    """견적서작성/전자서명 같은 Qt 페이지를 자기 프로세스에서 단독 실행."""
    _setup_logging()
    _setup_exception_hook()
    logger.info("%s 시작", title)

    app = QApplication(sys.argv)
    f = app.font()
    f.setFamily("Malgun Gothic")
    app.setFont(f)

    win = _PageWindow(title, mod_name, cls_name)
    # 창 위치를 저장해 둔 적이 있으면 그대로, 처음이면 최대화
    if win._geo_restored:
        win.show()
    else:
        win.showMaximized()
    sys.exit(app.exec())


# ──────────────────────────────────────────────
# 페이지 창 (견적서작성 / 전자서명 각각 독립 창)
# ──────────────────────────────────────────────
class _PageWindow(QMainWindow):
    """기능 페이지를 독립 창으로 표시. 초기화 버튼은 우측 상단 툴바."""

    def __init__(self, title: str, mod_name: str, cls_name: str,
                 parent: Optional[QWidget] = None) -> None:
        super().__init__(parent, Qt.Window)
        self.setWindowTitle(f"업무자동화 통합 시스템 — {title}")
        self._mod_name = mod_name
        self._cls_name = cls_name

        # ── 초기화 버튼 — 툴바 우측 ──────────────────
        tb = self.addToolBar("controls")
        tb.setMovable(False)
        tb.setFloatable(False)
        spacer = QWidget()
        spacer.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Preferred)
        tb.addWidget(spacer)
        reset_btn = QPushButton("↺  초기화")
        reset_btn.setFixedHeight(34)
        reset_btn.setFixedWidth(110)
        f = reset_btn.font(); f.setPointSize(11); f.setBold(True); reset_btn.setFont(f)
        reset_btn.setStyleSheet("""
            QPushButton {
                background-color: #E53935;
                color: white;
                border-radius: 5px;
                border: none;
            }
            QPushButton:hover  { background-color: #C62828; }
            QPushButton:pressed{ background-color: #B71C1C; }
        """)
        reset_btn.clicked.connect(self._reset)
        tb.addWidget(reset_btn)

        # ── 페이지 위젯 ───────────────────────────────
        import importlib
        mod = importlib.import_module(mod_name)
        cls = getattr(mod, cls_name)
        self._page = cls()
        self.setCentralWidget(self._page)

        self._geo_key = cls_name
        self._geo_restored = app_settings.restore_geometry(self._geo_key, self)

    def closeEvent(self, event) -> None:
        app_settings.save_geometry(self._geo_key, self)
        self._shutdown_page()
        super().closeEvent(event)

    def _shutdown_page(self) -> None:
        """페이지가 백그라운드 자원(전자서명의 Excel COM 세션 등)을 쥐고 있으면
        정리한다. 빠뜨리면 보이지 않는 EXCEL.EXE 가 남아 파일 잠금을 쥔다."""
        fn = getattr(self._page, "shutdown", None)
        if fn is None:
            return
        try:
            fn()
        except Exception:
            logger.warning("페이지 정리 실패", exc_info=True)

    def _reset(self) -> None:
        try:
            if hasattr(self._page, "reset_page"):
                self._page.reset_page()
            else:
                import importlib
                mod = importlib.import_module(self._mod_name)
                cls = getattr(mod, self._cls_name)
                # 새 페이지를 만들기 전에 옛 페이지의 자원을 먼저 반납한다
                # (전자서명은 Excel 세션을 쥐고 있어 두 개가 겹치면 안 된다).
                self._shutdown_page()
                new_page = cls()
                self.setCentralWidget(new_page)
                self._page.deleteLater()
                self._page = new_page
        except Exception as e:
            QMessageBox.critical(self, "초기화 오류", str(e))


