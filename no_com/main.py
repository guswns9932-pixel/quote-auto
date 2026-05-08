"""
main.py
=======
앱 진입점 + MainWindow + LeftNav
"""

from __future__ import annotations

import logging
import sys

from PySide6.QtWidgets import (
    QApplication, QHBoxLayout, QMainWindow, QMessageBox,
    QPushButton, QSizePolicy, QStackedWidget, QVBoxLayout, QWidget,
)

from pages import ESignPage, QuoteBuilderPage, RackPurchasePage
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
# 왼쪽 내비게이션
# ──────────────────────────────────────────────
class LeftNav(QWidget):
    def __init__(self, stack: QStackedWidget, on_reset) -> None:
        super().__init__()
        self.setFixedWidth(220)

        v = QVBoxLayout(self)
        v.setContentsMargins(12, 12, 12, 12)
        v.setSpacing(12)

        buttons = [
            ("견적서작성",       lambda: stack.setCurrentIndex(0), "#BBDEFB"),   # 연파랑
            ("전자서명",         lambda: stack.setCurrentIndex(1), "#C8E6C9"),   # 연초록
            ("RACK구매요청서",   lambda: stack.setCurrentIndex(2), "#FFE0B2"),   # 연주황
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
    def __init__(self) -> None:
        super().__init__()
        self.setWindowTitle("견적/전자서명 통합 시스템")

        central = QWidget()
        self.setCentralWidget(central)
        layout = QHBoxLayout(central)
        layout.setContentsMargins(10, 10, 10, 10)
        layout.setSpacing(12)

        self.stack = QStackedWidget()
        self._populate_stack()

        self.nav = LeftNav(self.stack, self.reset)
        layout.addWidget(self.nav)
        layout.addWidget(self.stack, 1)

    # ── 페이지 채우기 ────────────────────────────
    def _populate_stack(self) -> None:
        self.stack.addWidget(QuoteBuilderPage())
        self.stack.addWidget(ESignPage())
        self.stack.addWidget(RackPurchasePage())

    # ── 초기화 ───────────────────────────────────
    def reset(self) -> None:
        try:
            while self.stack.count():
                w = self.stack.widget(0)
                self.stack.removeWidget(w)
                w.deleteLater()
            self._populate_stack()
            self.stack.setCurrentIndex(0)
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

    window = MainWindow()
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
