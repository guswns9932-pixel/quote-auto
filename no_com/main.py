"""
main.py
=======
앱 진입점 + MainWindow + LeftNav
"""

from __future__ import annotations

import logging
import sys
import threading

from PySide6.QtCore import Qt, QRect
from PySide6.QtGui import QColor, QFont, QPainter, QPixmap
from PySide6.QtWidgets import (
    QApplication, QHBoxLayout, QMainWindow, QMessageBox,
    QPushButton, QSizePolicy, QSplashScreen, QStackedWidget, QVBoxLayout, QWidget,
)

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
# 왼쪽 내비게이션
# ──────────────────────────────────────────────
class LeftNav(QWidget):
    def __init__(self, stack: QStackedWidget, on_reset, on_navigate=None) -> None:
        super().__init__()
        self.setFixedWidth(220)

        v = QVBoxLayout(self)
        v.setContentsMargins(12, 12, 12, 12)
        v.setSpacing(12)

        def _go(idx):
            if on_navigate:
                on_navigate(idx)
            else:
                stack.setCurrentIndex(idx)

        buttons = [
            ("견적서작성", lambda: _go(0), "#BBDEFB"),
            ("RACK구매요청서 작성", lambda: _go(1), "#FFF176"),
            ("전자서명",   lambda: _go(2), "#C8E6C9"),
        ]
        for label, slot, color in buttons:
            btn = self._nav_btn(label, slot)
            tint_button(btn, color)
            v.addWidget(btn)

        v.addStretch(1)

        reset_btn = self._nav_btn("초기화", on_reset)
        tint_button(reset_btn, "#FFCCBC")   # 연주황(리셋 강조)
        v.addWidget(reset_btn)

    @staticmethod
    def _nav_btn(label: str, slot) -> QPushButton:
        btn = QPushButton(label)
        btn.setMinimumHeight(70)
        btn.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Fixed)
        f = btn.font()
        f.setPointSize(12)
        f.setBold(True)
        btn.setFont(f)
        btn.clicked.connect(slot)
        return btn


# ──────────────────────────────────────────────
# 메인 윈도우
# ──────────────────────────────────────────────
class MainWindow(QMainWindow):
    # 인덱스 → (모듈 클래스명, 클래스 참조 캐시)
    _PAGE_DEFS = [
        ("pages", "QuoteBuilderPage"),
        ("pages", "RackPurchaseRequestPage"),
        ("pages", "ESignPage"),
    ]

    def __init__(self) -> None:
        super().__init__()
        self.setWindowTitle("견적/전자서명 통합 시스템")

        central = QWidget()
        self.setCentralWidget(central)
        layout = QHBoxLayout(central)
        layout.setContentsMargins(10, 10, 10, 10)
        layout.setSpacing(12)

        self.stack = QStackedWidget()
        self._page_cache: dict = {}   # index → QWidget (생성된 페이지)

        # 스택에 빈 플레이스홀더를 먼저 채워 인덱스 유지 (pages 1,2는 첫 탐색 시 지연 생성)
        for _ in self._PAGE_DEFS:
            self.stack.addWidget(QWidget())

        self.nav = LeftNav(self.stack, self.reset, self._ensure_page)
        layout.addWidget(self.nav)
        layout.addWidget(self.stack, 1)

        # 첫 페이지(index 0)는 동기 생성 — 창이 최대화될 때 레이아웃이 올바르게 채워지도록
        self._ensure_page(0)

    def _ensure_page(self, index: int) -> None:
        """index 페이지가 아직 생성되지 않았으면 지금 생성해 스택에 교체."""
        if index in self._page_cache:
            self.stack.setCurrentIndex(index)
            return
        mod_name, cls_name = self._PAGE_DEFS[index]
        import importlib
        mod = importlib.import_module(mod_name)
        cls = getattr(mod, cls_name)
        page = cls()
        self._page_cache[index] = page
        # 플레이스홀더와 교체
        old = self.stack.widget(index)
        self.stack.insertWidget(index, page)
        self.stack.removeWidget(old)
        old.deleteLater()
        self.stack.setCurrentIndex(index)

    # ── 초기화 ───────────────────────────────────
    def reset(self) -> None:
        try:
            cur = self.stack.currentWidget()
            if cur is not None and hasattr(cur, "reset_page"):
                cur.reset_page()
                return
            idx = self.stack.currentIndex()
            # 현재 페이지만 재생성
            if idx in self._page_cache:
                old = self._page_cache.pop(idx)
                self.stack.insertWidget(idx, QWidget())
                self.stack.removeWidget(old)
                old.deleteLater()
            self._ensure_page(idx)
        except Exception as e:
            QMessageBox.critical(self, "초기화 오류", str(e))


# ──────────────────────────────────────────────
# 진입점
# ──────────────────────────────────────────────
def main() -> None:
    _setup_logging()

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

    # ── 백그라운드에서 무거운 모듈 프리로드 ────────────────────────────
    # openpyxl 은 excel_io 임포트 시 함께 로드됨 (가장 느린 부분).
    # 메인 스레드를 블로킹하지 않고 QApplication 이벤트 루프를 유지한다.
    _ready = threading.Event()

    def _preload() -> None:
        try:
            import excel_io  # noqa: F401 — openpyxl 포함
        finally:
            _ready.set()

    threading.Thread(target=_preload, daemon=True).start()

    while not _ready.is_set():
        _ready.wait(0.05)
        app.processEvents()

    # ── 메인 윈도우 생성 (이 시점엔 무거운 모듈이 이미 캐시됨) ─────────
    window = MainWindow()
    splash.finish(window)
    window.showMaximized()

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
