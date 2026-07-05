"""
page_common.py
==============
공통 유틸리티: 자연 정렬, 평탄 테이블 생성, 스크롤 가능한 오류 다이얼로그, 범용 백그라운드 작업자.
"""
from __future__ import annotations

import re
import traceback
from typing import List

from PySide6.QtCore import Qt, QThread, Signal
from PySide6.QtWidgets import (
    QAbstractItemView, QDialog, QHBoxLayout, QProgressDialog,
    QPushButton, QTableWidget, QTextEdit, QVBoxLayout,
)


def _natural_key(s: str):
    """파일명 자연 정렬 키: 숫자 부분을 수치로 비교 (10 < 20 < 100)."""
    return [int(c) if c.isdigit() else c.lower() for c in re.split(r'(\d+)', s)]


def _make_plain_table(cols: int, headers: List[str]) -> QTableWidget:
    t = QTableWidget(0, cols)
    t.setHorizontalHeaderLabels(headers)
    t.verticalHeader().setVisible(False)
    t.setEditTriggers(QTableWidget.NoEditTriggers)
    t.setSelectionBehavior(QAbstractItemView.SelectRows)
    t.setSelectionMode(QAbstractItemView.SingleSelection)
    t.setWordWrap(False)
    return t


class _ScrollableErrorDialog(QDialog):
    """긴 오류 메시지를 스크롤·드래그로 볼 수 있는 오류 다이얼로그."""

    def __init__(self, parent=None, message: str = "") -> None:
        super().__init__(parent)
        self.setWindowTitle("오류")
        self.resize(640, 400)
        self.setWindowFlags(self.windowFlags() & ~Qt.WindowContextHelpButtonHint)

        layout = QVBoxLayout(self)
        layout.setContentsMargins(12, 12, 12, 12)
        layout.setSpacing(8)

        text = QTextEdit()
        text.setReadOnly(True)
        text.setPlainText(message)
        text.setLineWrapMode(QTextEdit.NoWrap)
        f = text.font()
        f.setFamily("Consolas")
        f.setPointSize(9)
        text.setFont(f)
        layout.addWidget(text)

        btn = QPushButton("확인")
        btn.setFixedWidth(80)
        btn.clicked.connect(self.accept)
        h = QHBoxLayout()
        h.addStretch()
        h.addWidget(btn)
        layout.addLayout(h)


class _BgWorker(QThread):
    """범용 백그라운드 작업 스레드. UI 블로킹 없이 Excel 작업 실행."""
    result = Signal(object)
    error  = Signal(str)

    def __init__(self, fn, *args, parent=None, **kwargs):
        super().__init__(parent)
        self._fn, self._args, self._kwargs = fn, args, kwargs

    def run(self) -> None:
        try:
            self.result.emit(self._fn(*self._args, **self._kwargs))
        except Exception as e:
            import traceback
            self.error.emit(f"{e}\n\n{traceback.format_exc()}")

    @staticmethod
    def run_with_progress(parent, msg: str, fn, *args,
                          on_result=None, on_error=None, **kwargs) -> "_BgWorker":
        """진행 다이얼로그를 띄우고 fn을 백그라운드 실행. 완료 시 콜백 호출."""
        from PySide6.QtWidgets import QProgressDialog
        dlg = QProgressDialog(msg, None, 0, 0, parent)
        dlg.setWindowModality(Qt.WindowModal)
        dlg.setMinimumDuration(0)
        dlg.show()

        worker = _BgWorker(fn, *args, parent=parent, **kwargs)

        def _on_result(r):
            dlg.close()
            if on_result:
                on_result(r)

        def _on_error(e):
            dlg.close()
            if on_error:
                on_error(e)
            else:
                _ScrollableErrorDialog(parent, e).exec()

        worker.result.connect(_on_result)
        worker.error.connect(_on_error)
        worker.finished.connect(worker.deleteLater)
        worker.start()
        return worker
