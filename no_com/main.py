"""
main.py
=======
앱 진입점 + 런처(MainWindow) + 페이지 창(_PageWindow)
"""

from __future__ import annotations

import logging
import os
import sys
import threading
import traceback
from typing import Optional

from PySide6.QtCore import Qt, QRect
from PySide6.QtGui import QColor, QFont, QPainter, QPixmap
from PySide6.QtWidgets import (
    QApplication, QHBoxLayout, QMainWindow, QMessageBox,
    QPushButton, QSizePolicy, QSplashScreen, QVBoxLayout, QWidget,
)

import app_settings
from widgets import tint_button


# ──────────────────────────────────────────────
# 로거 초기화
# ──────────────────────────────────────────────
def _setup_logging() -> None:
    root = logging.getLogger("QuoteApp")
    root.setLevel(logging.DEBUG)
    if not root.handlers:
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
            fh = logging.FileHandler(
                os.path.join(_exe_dir, "quote_app.log"),
                encoding="utf-8", mode="a",
            )
            fh.setFormatter(logging.Formatter(
                "[%(levelname)s][%(asctime)s][%(name)s] %(message)s",
                datefmt="%Y-%m-%d %H:%M:%S",
            ))
            root.addHandler(fh)
        except Exception:
            pass


def _setup_exception_hook() -> None:
    """메인 스레드 uncaught exception → 로그 + 대화상자 (silent crash 방지)."""
    def _hook(exc_type, exc_value, exc_tb):
        msg = "".join(traceback.format_exception(exc_type, exc_value, exc_tb))
        logging.getLogger("QuoteApp").critical("치명적 오류:\n%s", msg)
        try:
            QMessageBox.critical(
                None, "치명적 오류",
                f"예상치 못한 오류가 발생했습니다:\n\n{exc_value}\n\n"
                "프로그램 폴더의 quote_app.log 에서 자세한 내용을 확인하세요.",
            )
        except Exception:
            pass
    sys.excepthook = _hook


logger = logging.getLogger("QuoteApp")


# ──────────────────────────────────────────────
# 스플래시 스크린
# ──────────────────────────────────────────────
def _make_splash_pix() -> QPixmap:
    pix = QPixmap(520, 200)
    pix.fill(QColor("#1565C0"))
    p = QPainter(pix)

    p.setPen(QColor("#FFFFFF"))
    f = QFont("Malgun Gothic", 20)
    f.setBold(True)
    p.setFont(f)
    p.drawText(QRect(0, 55, 520, 65), Qt.AlignCenter, "견적/전자서명 통합 시스템")

    p.setPen(QColor("#BBDEFB"))
    f2 = QFont("Malgun Gothic", 11)
    p.setFont(f2)
    p.drawText(QRect(0, 130, 520, 40), Qt.AlignCenter, "초기화 중…")

    p.end()
    return pix


# ──────────────────────────────────────────────
# 페이지 창 (견적서작성 / 전자서명 각각 독립 창)
# ──────────────────────────────────────────────
class _PageWindow(QMainWindow):
    """기능 페이지를 독립 창으로 표시. 초기화 버튼은 우측 상단 툴바."""

    def __init__(self, title: str, mod_name: str, cls_name: str,
                 parent: Optional[QWidget] = None) -> None:
        super().__init__(parent, Qt.Window)
        self.setWindowTitle(f"견적/전자서명 통합 시스템 — {title}")
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
        super().closeEvent(event)

    def _reset(self) -> None:
        try:
            if hasattr(self._page, "reset_page"):
                self._page.reset_page()
            else:
                import importlib
                mod = importlib.import_module(self._mod_name)
                cls = getattr(mod, self._cls_name)
                new_page = cls()
                self.setCentralWidget(new_page)
                self._page.deleteLater()
                self._page = new_page
        except Exception as e:
            QMessageBox.critical(self, "초기화 오류", str(e))


# ──────────────────────────────────────────────
# 런처 윈도우 (메인)
# ──────────────────────────────────────────────
class MainWindow(QMainWindow):
    """두 버튼으로 각 기능 창을 여는 런처."""

    def __init__(self) -> None:
        super().__init__()
        self.setWindowTitle("견적/전자서명 통합 시스템")
        self.setFixedSize(340, 210)
        app_settings.restore_geometry("MainWindow", self)

        self._quote_win: Optional[_PageWindow] = None
        self._esign_win: Optional[_PageWindow] = None

        central = QWidget()
        self.setCentralWidget(central)
        v = QVBoxLayout(central)
        v.setContentsMargins(30, 30, 30, 30)
        v.setSpacing(18)

        btn_quote = QPushButton("견적서작성")
        btn_esign = QPushButton("전자서명")

        for btn, color in [(btn_quote, "#BBDEFB"), (btn_esign, "#C8E6C9")]:
            btn.setMinimumHeight(60)
            f = btn.font(); f.setPointSize(14); f.setBold(True); btn.setFont(f)
            tint_button(btn, color)
            v.addWidget(btn)

        btn_quote.clicked.connect(
            lambda: self._open_page("견적서작성", "pages", "QuoteBuilderPage", "_quote_win"))
        btn_esign.clicked.connect(
            lambda: self._open_page("전자서명", "pages", "ESignPage", "_esign_win"))

    def _open_page(self, title: str, mod: str, cls: str, attr: str) -> None:
        win: Optional[_PageWindow] = getattr(self, attr)
        if win is None or not win.isVisible():
            win = _PageWindow(title, mod, cls, parent=self)
            setattr(self, attr, win)
            if win._geo_restored:
                win.show()
            else:
                win.showMaximized()
        else:
            win.raise_()
            win.activateWindow()

    def closeEvent(self, event) -> None:
        app_settings.save_geometry("MainWindow", self)
        super().closeEvent(event)


# ──────────────────────────────────────────────
# 진입점
# ──────────────────────────────────────────────
def main() -> None:
    _setup_logging()
    _setup_exception_hook()
    logger.info("앱 시작")

    # 런처에서 in-process로 호출되면 QApplication이 이미 존재한다.
    # 그 경우 새로 생성하지 않고 기존 인스턴스를 재사용한다.
    _standalone = QApplication.instance() is None
    app = QApplication(sys.argv) if _standalone else QApplication.instance()

    f = app.font()
    f.setFamily("Malgun Gothic")
    app.setFont(f)

    # ── 스플래시 스크린: 창 보이기 전에 즉시 표시 ─────────────────────
    splash = QSplashScreen(_make_splash_pix(), Qt.WindowStaysOnTopHint)
    splash.show()
    app.processEvents()

    # ── openpyxl 백그라운드 프리로드 (창 생성과 완전히 병렬) ───────────
    # pages.py 가 더 이상 excel_io 를 모듈 수준에서 임포트하지 않으므로
    # 프리로드 완료를 기다리지 않고 즉시 창을 생성할 수 있다.
    # 사용자가 버튼을 누를 시점(체감 2-3초 이후)에는 로드가 완료되어 있다.
    def _preload() -> None:
        try:
            import openpyxl  # noqa: F401 — 순수 Python, GIL 점유 없음
        except Exception:
            pass
        # COM DLL(pythoncom/win32com)은 프리로드하지 않음:
        # LoadLibrary 중 GIL을 점유해 Qt 메인 스레드를 동결시킴.
        # AV/Defender 환경에서 최대 120초 블로킹 가능.

    threading.Thread(target=_preload, daemon=True).start()

    # ── 런처 윈도우 생성 (고정 크기이므로 show) ─────────────────────────
    window = MainWindow()
    splash.finish(window)
    window.show()

    if _standalone:
        sys.exit(app.exec())
    else:
        # 런처 내부 실행: 독립 QEventLoop 사용.
        # 앱 창이 모두 닫히면(lastWindowClosed) 루프를 종료하고
        # 제어를 런처로 반환한다.
        from PySide6.QtCore import QEventLoop
        loop = QEventLoop()
        app.lastWindowClosed.connect(loop.quit)
        loop.exec()
        try:
            app.lastWindowClosed.disconnect(loop.quit)
        except Exception:
            pass


if __name__ == "__main__":
    main()
