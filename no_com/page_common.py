"""
page_common.py
==============
공통 유틸리티: 자연 정렬, 평탄 테이블 생성, 스크롤 가능한 오류 다이얼로그, 범용 백그라운드 작업자.
"""
from __future__ import annotations

import re
import traceback
from typing import List, Tuple

from PySide6.QtCore import Qt, QThread, Signal
from PySide6.QtWidgets import (
    QAbstractItemView, QDialog, QFrame, QHBoxLayout, QLabel,
    QProgressDialog, QPushButton, QTableWidget, QTextEdit, QVBoxLayout,
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


def _friendly_error_msg(exc: BaseException) -> Tuple[str, str]:
    """예외 유형을 분석해 (사용자 메시지, 해결 힌트) 한국어 튜플을 반환한다."""
    msg = str(exc)
    type_name = type(exc).__name__

    if isinstance(exc, PermissionError) or "WinError 32" in msg or "being used by another process" in msg:
        return (
            "파일이 다른 프로그램에서 열려 있습니다.",
            "Excel 또는 관련 프로그램을 완전히 닫은 후 다시 시도해 주세요.",
        )
    if isinstance(exc, FileNotFoundError):
        return (
            "파일을 찾을 수 없습니다.",
            "파일 경로를 확인하거나 파일이 삭제·이동되지 않았는지 확인해 주세요.",
        )
    if "pywintypes" in type_name or "com_error" in type_name.lower() or "CoInitialize" in msg:
        return (
            "Excel COM 연결에 실패했습니다.",
            "Excel이 설치되어 있는지 확인하고, Excel을 완전히 종료한 후 다시 시도해 주세요.",
        )
    if "BadZipFile" in type_name or "zipfile" in msg.lower() or "not a zip" in msg.lower():
        return (
            "Excel 파일 형식이 올바르지 않거나 손상되었습니다.",
            "파일을 다시 열어 저장한 후 시도해 주세요.",
        )
    if isinstance(exc, MemoryError):
        return (
            "메모리가 부족합니다.",
            "다른 프로그램을 종료한 후 다시 시도해 주세요.",
        )
    return (
        "예기치 않은 오류가 발생했습니다.",
        "아래 기술 정보를 캡처하여 담당자에게 문의해 주세요.",
    )


class _ScrollableErrorDialog(QDialog):
    """사용자 친화적 메시지 + 토글 가능한 기술 세부 정보 다이얼로그."""

    def __init__(self, parent=None, message: str = "",
                 user_msg: str = "", hint: str = "") -> None:
        super().__init__(parent)
        self.setWindowTitle("오류")
        self.setMinimumWidth(560)
        self.setWindowFlags(self.windowFlags() & ~Qt.WindowContextHelpButtonHint)

        layout = QVBoxLayout(self)
        layout.setContentsMargins(16, 16, 16, 12)
        layout.setSpacing(8)

        if user_msg:
            lbl = QLabel(f"<b>{user_msg}</b>")
            lbl.setWordWrap(True)
            layout.addWidget(lbl)
        if hint:
            lbl_h = QLabel(hint)
            lbl_h.setWordWrap(True)
            layout.addWidget(lbl_h)

        self._details = None
        self._toggle_btn = None
        if message:
            sep = QFrame()
            sep.setFrameShape(QFrame.HLine)
            layout.addWidget(sep)

            self._details = QTextEdit()
            self._details.setReadOnly(True)
            self._details.setPlainText(message)
            self._details.setLineWrapMode(QTextEdit.NoWrap)
            f = self._details.font()
            f.setFamily("Consolas"); f.setPointSize(9)
            self._details.setFont(f)
            self._details.setFixedHeight(220)
            # 친화적 메시지가 없으면 기술 정보를 바로 표시
            self._details.setVisible(not bool(user_msg))
            layout.addWidget(self._details)

        btn_row = QHBoxLayout()
        if message and user_msg:
            self._toggle_btn = QPushButton("기술 정보 보기 ▼")
            self._toggle_btn.setCheckable(True)
            self._toggle_btn.clicked.connect(self._toggle_details)
            btn_row.addWidget(self._toggle_btn)
        btn_row.addStretch()
        ok_btn = QPushButton("확인")
        ok_btn.setFixedWidth(80)
        ok_btn.clicked.connect(self.accept)
        btn_row.addWidget(ok_btn)
        layout.addLayout(btn_row)

        self.adjustSize()

    def _toggle_details(self, checked: bool) -> None:
        if self._details:
            self._details.setVisible(checked)
        if self._toggle_btn:
            self._toggle_btn.setText("기술 정보 숨기기 ▲" if checked else "기술 정보 보기 ▼")
        self.adjustSize()


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
