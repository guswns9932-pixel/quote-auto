"""
widgets.py
==========
재사용 가능한 커스텀 Qt 위젯 모음.
비즈니스 로직 없음 – 순수 UI 컴포넌트.
"""

from __future__ import annotations

from typing import Callable, List, Optional, Tuple

from PySide6.QtCore import Qt, QTimer, QPointF
from PySide6.QtGui import (
    QAction, QBrush, QColor, QDrag, QKeySequence, QPixmap, QShortcut,
)
from PySide6.QtWidgets import (
    QAbstractItemView, QApplication, QCheckBox, QDialog,
    QDialogButtonBox, QGraphicsPixmapItem, QGraphicsScene,
    QGraphicsView, QHBoxLayout, QLabel, QLineEdit, QMenu,
    QMessageBox, QMimeData, QPushButton, QScrollBar,
    QSizePolicy, QTableWidget, QTableWidgetItem, QVBoxLayout,
    QWidget, QFrame,
)

from core import fmt_krw, s, to_float


# ══════════════════════════════════════════════
# 공통 팩토리
# ══════════════════════════════════════════════

def labeled_frame(title: str, min_h: int = 80) -> QFrame:
    """
    제목 레이블이 달린 Box-테두리 QFrame을 만들어 반환.
    내부 layout: QVBoxLayout (title이 첫 번째 위젯).
    """
    frame = QFrame()
    frame.setFrameShape(QFrame.Box)
    frame.setLineWidth(2)
    frame.setMinimumHeight(min_h)

    layout = QVBoxLayout(frame)
    layout.setContentsMargins(12, 12, 12, 12)
    layout.setSpacing(8)

    lbl = QLabel(title)
    lbl.setAlignment(Qt.AlignHCenter | Qt.AlignTop)
    font = lbl.font()
    font.setPointSize(12)
    font.setBold(True)
    lbl.setFont(font)
    lbl.setFixedHeight(32)
    lbl.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Fixed)
    layout.addWidget(lbl, 0, Qt.AlignTop)
    return frame


def bold_label(text: str, size: int = 12) -> QLabel:
    lbl = QLabel(text)
    f = lbl.font()
    f.setPointSize(size)
    f.setBold(True)
    lbl.setFont(f)
    return lbl


def tint_button(btn: QPushButton, bg: str) -> None:
    """버튼에 파스텔 배경색 + hover 어둡게 + disabled 회색 적용."""
    r = max(0, int(bg[1:3], 16) - 28)
    g = max(0, int(bg[3:5], 16) - 28)
    b = max(0, int(bg[5:7], 16) - 28)
    hover = f"#{r:02X}{g:02X}{b:02X}"
    btn.setStyleSheet(
        f"QPushButton {{ background-color: {bg}; border: 1px solid #bbb; border-radius: 4px; padding: 2px 6px; }}"
        f"QPushButton:hover {{ background-color: {hover}; }}"
        f"QPushButton:disabled {{ background-color: #E0E0E0; color: #999; }}"
    )


def info_label(text: str) -> QLabel:
    lbl = QLabel(text)
    lbl.setTextFormat(Qt.PlainText)
    lbl.setWordWrap(False)
    lbl.setAlignment(Qt.AlignLeft | Qt.AlignVCenter)
    lbl.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Fixed)
    lbl.setMinimumHeight(20)
    f = lbl.font()
    f.setFamily("Malgun Gothic")
    f.setPointSize(9)
    lbl.setFont(f)
    return lbl


# ══════════════════════════════════════════════
# 드래그 지원 품목 테이블 (STEP4)
# ══════════════════════════════════════════════

class DraggableItemsTable(QTableWidget):
    """
    규격/단가 2컬럼 테이블.
    행을 드래그하면 "규격\t단가" 텍스트를 mime 데이터로 전달.
    """

    def startDrag(self, _supported) -> None:
        rows = self.selectionModel().selectedRows()
        if not rows:
            return
        r = rows[0].row()
        spec_it  = self.item(r, 0)
        price_it = self.item(r, 1)
        if not spec_it:
            return
        mime = QMimeData()
        mime.setText(f"{spec_it.text().strip()}\t{price_it.text().strip() if price_it else '0'}")
        drag = QDrag(self)
        drag.setMimeData(mime)
        drag.exec(Qt.CopyAction)


# ══════════════════════════════════════════════
# 드롭 지원 견적 테이블 (STEP5)
# ══════════════════════════════════════════════

class DroppableQuoteTable(QTableWidget):
    """
    드래그앤드롭으로 품목을 받아 콜백을 호출하는 테이블.
    drop_callback(spec: str, unit_price: float)
    """

    def __init__(self, parent=None) -> None:
        super().__init__(parent)
        self.drop_callback: Optional[Callable[[str, float], None]] = None
        self.setAcceptDrops(True)
        self.setDragDropMode(QAbstractItemView.DropOnly)

    def dragEnterEvent(self, e):
        e.acceptProposedAction() if e.mimeData().hasText() else e.ignore()

    def dragMoveEvent(self, e):
        e.acceptProposedAction() if e.mimeData().hasText() else e.ignore()

    def dropEvent(self, e):
        if not e.mimeData().hasText():
            e.ignore()
            return
        parts = e.mimeData().text().split("\t")
        spec  = parts[0].strip()
        price = to_float(parts[1], 0.0) if len(parts) >= 2 else 0.0
        if self.drop_callback and spec:
            self.drop_callback(spec, price)
            e.acceptProposedAction()
        else:
            e.ignore()


# ══════════════════════════════════════════════
# 센터 정렬 체크박스 셀 위젯
# ══════════════════════════════════════════════

def centered_checkbox(on_changed: Callable[[int], None]) -> QWidget:
    """체크박스를 가운데 정렬한 셀 위젯 반환."""
    host = QWidget()
    lay  = QHBoxLayout(host)
    lay.setContentsMargins(0, 0, 0, 0)
    lay.setAlignment(Qt.AlignCenter)
    cb = QCheckBox()
    cb.stateChanged.connect(on_changed)
    lay.addWidget(cb)
    return host


def get_checkbox_from_cell(widget: Optional[QWidget]) -> Optional[QCheckBox]:
    if not widget:
        return None
    return widget.findChild(QCheckBox)


# ══════════════════════════════════════════════
# 패스워드 확인 다이얼로그
# ══════════════════════════════════════════════

class PasswordDialog(QDialog):
    def __init__(self, parent, expected: str) -> None:
        super().__init__(parent)
        self.setWindowTitle("패스워드")
        self._expected = expected or ""
        self.verified  = False

        v = QVBoxLayout(self)
        v.addWidget(QLabel("패스워드를 입력하세요"))
        self._edit = QLineEdit()
        self._edit.setEchoMode(QLineEdit.Password)
        v.addWidget(self._edit)

        bb = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        bb.accepted.connect(self._check)
        bb.rejected.connect(self.reject)
        v.addWidget(bb)

        QTimer.singleShot(0, self._edit.setFocus)

    def _check(self):
        if self._edit.text() == self._expected:
            self.verified = True
            self.accept()
        else:
            QMessageBox.warning(self, "오류", "패스워드가 일치하지 않습니다.")
            self._edit.selectAll()
            self._edit.setFocus()


# ══════════════════════════════════════════════
# 서명 아이템 (QGraphicsPixmapItem 확장)
# ══════════════════════════════════════════════

class SignatureItem(QGraphicsPixmapItem):
    """
    씬 위에 올라가는 서명 이미지.
    드래그 이동 + 우클릭 삭제 지원.
    """

    def __init__(self, pixmap: QPixmap, page_index: int) -> None:
        super().__init__(pixmap)
        self.page_index = page_index
        self.setFlags(
            QGraphicsPixmapItem.ItemIsMovable |
            QGraphicsPixmapItem.ItemIsSelectable
        )
        self.setAcceptedMouseButtons(Qt.LeftButton | Qt.RightButton)

    def contextMenuEvent(self, event) -> None:
        menu = QMenu()
        act  = QAction("서명 삭제", menu)
        menu.addAction(act)
        if menu.exec(event.screenPos()) == act:
            scene = self.scene()
            if scene:
                scene.removeItem(self)
        event.accept()


# ══════════════════════════════════════════════
# PDF 뷰어 (PgUp/PgDn + 휠 페이지 전환)
# ══════════════════════════════════════════════

class PdfView(QGraphicsView):
    """
    커스텀 QGraphicsView.
    - 더블클릭: on_double_click(QPointF) 콜백
    - PgUp/PgDn: on_prev / on_next 콜백
    - 휠: 페이지 끝에서 on_prev / on_next 호출
    """

    def __init__(self, parent=None) -> None:
        super().__init__(parent)
        self.on_prev         : Optional[Callable] = None
        self.on_next         : Optional[Callable] = None
        self.on_double_click : Optional[Callable[[QPointF], None]] = None

        self._locked = False
        self._unlock_timer = QTimer(self)
        self._unlock_timer.setSingleShot(True)
        self._unlock_timer.timeout.connect(self._unlock)

        self.setDragMode(QGraphicsView.NoDrag)
        self.setVerticalScrollBarPolicy(Qt.ScrollBarAsNeeded)
        self.setHorizontalScrollBarPolicy(Qt.ScrollBarAsNeeded)

    # ── 잠금 ────────────────────────────────────
    def _lock(self) -> None:
        self._locked = True
        self._unlock_timer.start(120)

    def _unlock(self) -> None:
        self._locked = False

    # ── 이벤트 ──────────────────────────────────
    def mouseDoubleClickEvent(self, event) -> None:
        if event.button() == Qt.LeftButton and self.on_double_click:
            self.on_double_click(self.mapToScene(event.pos()))
            event.accept()
            return
        super().mouseDoubleClickEvent(event)

    def keyPressEvent(self, event) -> None:
        if event.isAutoRepeat():
            event.ignore()
            return
        if event.key() == Qt.Key_PageDown and not self._locked:
            self._lock()
            if self.on_next:
                self.on_next()
            event.accept()
            return
        if event.key() == Qt.Key_PageUp and not self._locked:
            self._lock()
            if self.on_prev:
                self.on_prev()
            event.accept()
            return
        super().keyPressEvent(event)

    def wheelEvent(self, event) -> None:
        if self._locked:
            event.accept()
            return
        dy   = event.angleDelta().y()
        vbar = self.verticalScrollBar()
        if dy < 0 and vbar.value() < vbar.maximum():
            super().wheelEvent(event)
            return
        if dy > 0 and vbar.value() > vbar.minimum():
            super().wheelEvent(event)
            return
        self._lock()
        if dy < 0 and self.on_next:
            self.on_next()
        elif dy > 0 and self.on_prev:
            self.on_prev()
        event.accept()
