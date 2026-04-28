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

from pages import ESignPage, QuoteBuilderPage


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
            ("견적서작성", lambda: stack.setCurrentIndex(0)),
            ("전자서명",   lambda: stack.setCurrentIndex(1)),
        ]
        for label, slot in buttons:
            btn = self._nav_btn(label, slot)
            v.addWidget(btn)

        v.addStretch(1)

        reset_btn = self._nav_btn("초기화", on_reset)
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
    app = QApplication(sys.argv)

    # 한글 폰트
    f = app.font()
    f.setFamily("Malgun Gothic")
    app.setFont(f)

    window = MainWindow()
    window.showMaximized()
    sys.exit(app.exec())


if __name__ == "__main__":
    main()
