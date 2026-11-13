"""
quote_panels.py
===============
견적서작성의 검토용 화면 부품 — 단계 표시, 구성 도식, 금액, 생성 전 점검.
무엇을 어떤 상태로 보여줄지는 quote_review 가 정하고, 여기는 그리기만 한다.
"""

from __future__ import annotations

from typing import List, Optional

from PySide6.QtCore import Qt, Signal
from PySide6.QtWidgets import (
    QFrame, QGridLayout, QHBoxLayout, QLabel, QListWidget, QListWidgetItem,
    QProgressBar, QPushButton, QSizePolicy, QVBoxLayout, QWidget,
)
from PySide6.QtGui import QBrush, QColor

import quote_review as rv
from core import fmt_krw

# 상태별 색: (바탕, 테두리, 글자, 테두리 모양)
_TONE = {
    rv.OK:    ("#E7F4EA", "#2E7D32", "#1B5E20", "solid"),
    rv.DIFF:  ("#FFF4CF", "#9A6700", "#5C3D00", "solid"),
    rv.BAD:   ("#FDEAEA", "#C62828", "#8E1B1B", "solid"),
    rv.MISS:  ("#EEF3FB", "#2E5AA8", "#1F3864", "dashed"),
    rv.BLANK: ("#FFFFFF", "#8A93A3", "#5C6475", "dashed"),
}
_LEGEND = [(rv.OK, "정상"), (rv.DIFF, "수량이 CH와 다름"), (rv.BAD, "단가 0원"),
           (rv.MISS, "코드매핑에 있는데 안 담김"), (rv.BLANK, "없음")]
_CHECK_TONE = {
    rv.RED:    ("판단 필요", "#FDEAEA", "#8E1B1B"),
    rv.YELLOW: ("확인 권장", "#FFF4CF", "#5C3D00"),
    rv.INFO:   ("참고",      "#EEF3FB", "#1F3864"),
}
NAVY = "#1F3864"


def _short(text: str, n: int) -> str:
    return text if len(text) <= n else text[:n - 1] + "…"


# ══════════════════════════════════════════════
# 단계 표시
# ══════════════════════════════════════════════
class StepBar(QWidget):
    """① 데이터 → ② 의뢰 선택 → ③ 구성 확인 → ④ 생성. 어디까지 왔는지 보여주기만 한다."""

    NAMES = ["데이터", "의뢰 선택", "구성 확인", "생성"]

    def __init__(self, parent=None) -> None:
        super().__init__(parent)
        self.setFixedHeight(34)
        h = QHBoxLayout(self); h.setContentsMargins(0, 0, 0, 0); h.setSpacing(6)
        self._labels: List[QLabel] = []
        for i, name in enumerate(self.NAMES):
            lbl = QLabel(f" {i + 1}  {name} ")
            lbl.setAlignment(Qt.AlignCenter)
            lbl.setFixedHeight(28)
            self._labels.append(lbl)
            h.addWidget(lbl)
            if i < len(self.NAMES) - 1:
                arrow = QLabel("›"); arrow.setStyleSheet("color: #8A93A3; font-size: 16px;")
                h.addWidget(arrow)
        self.hint = QLabel("")
        self.hint.setStyleSheet("color: #5C6475;")
        h.addSpacing(10); h.addWidget(self.hint, 1)
        self.set_progress([False] * 4, "")

    def set_progress(self, done: List[bool], hint: str = "") -> None:
        """done[i] = i번째 단계를 마쳤는가. 마치지 않은 첫 단계가 '지금 할 일'이다."""
        current = next((i for i, d in enumerate(done) if not d), len(done) - 1)
        for i, lbl in enumerate(self._labels):
            if i == current:
                css = f"background: {NAVY}; color: white; font-weight: bold;"
            elif done[i]:
                css = "background: #E7F4EA; color: #1B5E20;"
            else:
                css = "background: #ECEEF2; color: #5C6475;"
            lbl.setStyleSheet(f"QLabel {{ {css} border-radius: 14px; padding: 0 10px; }}")
        self.hint.setText(hint)


# ══════════════════════════════════════════════
# 구성 도식
# ══════════════════════════════════════════════
class DiagramPanel(QFrame):
    """설비 → 배기 흐름으로 구성품을 보여준다. 칸을 누르면 그 품목을 고르거나(담긴 것)
    담는다(코드매핑이 제안하는 것)."""

    pickSpec = Signal(str)                 # 이 규격의 행을 선택해 달라
    addSuggest = Signal(str, float)        # 이 규격을 이 수량으로 담아 달라

    def __init__(self, parent=None) -> None:
        super().__init__(parent)
        self.setFrameShape(QFrame.Box); self.setLineWidth(2)
        v = QVBoxLayout(self); v.setContentsMargins(10, 8, 10, 8); v.setSpacing(6)
        top = QHBoxLayout()
        self.title = QLabel("구성 도식"); f = self.title.font(); f.setBold(True); f.setPointSize(11)
        self.title.setFont(f)
        top.addWidget(self.title); top.addStretch(1)
        tip = QLabel("칸을 누르면 아래 표에서 그 품목이 선택됩니다")
        tip.setStyleSheet("color: #5C6475;")
        top.addWidget(tip)
        v.addLayout(top)

        self._flow = QHBoxLayout(); self._flow.setSpacing(4)
        self._base = QHBoxLayout(); self._base.setSpacing(6)
        self._others = QHBoxLayout(); self._others.setSpacing(6)
        v.addLayout(self._flow); v.addLayout(self._base); v.addLayout(self._others)

        legend = QHBoxLayout(); legend.setSpacing(12)
        for state, text in _LEGEND:
            bg, bd, _, style = _TONE[state]
            box = QLabel(); box.setFixedSize(12, 12)
            box.setStyleSheet(f"background: {bg}; border: 1px {style} {bd};")
            cap = QLabel(text); cap.setStyleSheet("color: #3C4557; font-size: 11px;")
            pair = QHBoxLayout(); pair.setSpacing(4); pair.addWidget(box); pair.addWidget(cap)
            legend.addLayout(pair)
        legend.addStretch(1)
        v.addLayout(legend)
        self.show_review(None)

    @staticmethod
    def _clear(layout) -> None:
        while layout.count():
            w = layout.takeAt(0).widget()
            if w is not None:
                # deleteLater 만으로는 다음 이벤트 처리 때까지 화면에 남아 새 칸과 겹쳐 보인다
                w.hide()
                w.setParent(None)
                w.deleteLater()

    def _slot_button(self, slot: rv.Slot, tall: bool) -> QPushButton:
        bg, bd, fg, style = _TONE[slot.state]
        text = f"{slot.label}\n{_short(slot.text, 20)}" + (f"\n{slot.note}" if slot.note else "")
        btn = QPushButton(text)
        btn.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Fixed)
        btn.setFixedHeight(62 if tall else 48)
        btn.setToolTip(f"{slot.label}\n{slot.text}\n{slot.note}".strip())
        btn.setStyleSheet(
            f"QPushButton {{ background: {bg}; color: {fg}; border: 2px {style} {bd}; "
            f"border-radius: 5px; padding: 2px 6px; text-align: left; font-size: 11px; }}"
            f"QPushButton:hover {{ border-color: {NAVY}; }}")
        btn.slot = slot
        if slot.suggest:
            sp, q = slot.suggest
            btn.clicked.connect(lambda _=False, sp=sp, q=q: self.addSuggest.emit(sp, float(q)))
        elif slot.specs:
            btn.clicked.connect(lambda _=False, sp=slot.specs[0]: self.pickSpec.emit(sp))
        else:
            btn.setEnabled(False)
            btn.setStyleSheet(btn.styleSheet() + f"QPushButton:disabled {{ color: {fg}; }}")
        return btn

    @staticmethod
    def _end(text: str) -> QLabel:
        lbl = QLabel(text); lbl.setAlignment(Qt.AlignCenter); lbl.setFixedHeight(62)
        lbl.setStyleSheet("background: #EEF1F6; color: #3C4557; border-radius: 5px; padding: 0 8px;")
        return lbl

    def show_review(self, review: Optional[rv.Review]) -> None:
        for lay in (self._flow, self._base, self._others):
            self._clear(lay)
        self.buttons: List[QPushButton] = []
        if review is None:
            review = rv.analyze([])
        ch = "" if review.ch is None else f"  ·  {rv._fmt_qty(review.ch)}CH"
        self.title.setText(f"구성 도식  ·  {review.qtype}{ch}")

        self._flow.addWidget(self._end("설비"))
        for s in review.flow:
            arrow = QLabel("→"); arrow.setStyleSheet("color: #8A93A3;")
            self._flow.addWidget(arrow)
            b = self._slot_button(s, tall=True); self.buttons.append(b)
            self._flow.addWidget(b, 1)
        arrow = QLabel("→"); arrow.setStyleSheet("color: #8A93A3;")
        self._flow.addWidget(arrow); self._flow.addWidget(self._end("배기"))

        for s in review.base:
            b = self._slot_button(s, tall=False); self.buttons.append(b)
            self._base.addWidget(b, 1)

        cap = QLabel("기타 옵션"); cap.setStyleSheet("color: #3C4557; font-weight: bold;")
        self._others.addWidget(cap)
        if not review.others:
            none = QLabel("없음"); none.setStyleSheet("color: #5C6475;")
            self._others.addWidget(none)
        shown = review.others[:4]
        for s in shown:
            bg, bd, fg, _ = _TONE[s.state]
            chip = QPushButton(_short(s.text, 18) + ("  · 0원" if s.state == rv.BAD else ""))
            chip.setFixedHeight(26); chip.setToolTip(f"{s.label}\n{s.text}\n{s.note}")
            chip.setStyleSheet(f"QPushButton {{ background: {bg}; color: {fg}; border: 1px solid {bd}; "
                               f"border-radius: 13px; padding: 0 10px; font-size: 11px; }}")
            chip.slot = s
            chip.clicked.connect(lambda _=False, sp=s.specs[0]: self.pickSpec.emit(sp))
            self.buttons.append(chip)
            self._others.addWidget(chip)
        if len(review.others) > len(shown):
            more = QLabel(f"외 {len(review.others) - len(shown)}개"); more.setStyleSheet("color: #5C6475;")
            self._others.addWidget(more)
        self._others.addStretch(1)


# ══════════════════════════════════════════════
# 금액
# ══════════════════════════════════════════════
class AmountPanel(QFrame):
    """사양서와 같은 계산으로 견적단가를 미리 보여준다(국내). 칸 사용량도 함께."""

    def __init__(self, parent=None) -> None:
        super().__init__(parent)
        self.setFrameShape(QFrame.Box); self.setLineWidth(2)
        v = QVBoxLayout(self); v.setContentsMargins(10, 8, 10, 8); v.setSpacing(4)
        title = QLabel("금액"); f = title.font(); f.setBold(True); f.setPointSize(11); title.setFont(f)
        v.addWidget(title)
        grid = QGridLayout(); grid.setHorizontalSpacing(8); grid.setVerticalSpacing(2)
        self._rows = {}
        for i, key in enumerate(("pump", "rack", "credit")):
            cap = QLabel(""); cap.setStyleSheet("color: #3C4557;")
            val = QLabel("-"); val.setAlignment(Qt.AlignRight | Qt.AlignVCenter)
            grid.addWidget(cap, i, 0); grid.addWidget(val, i, 1)
            self._rows[key] = (cap, val)
        v.addLayout(grid)
        line = QFrame(); line.setFrameShape(QFrame.HLine); line.setStyleSheet("color: #D4D8E0;")
        v.addWidget(line)
        self.unit_cap = QLabel("CH당 견적단가 (천 원 미만 버림)"); self.unit_cap.setStyleSheet("color: #3C4557;")
        self.unit = QLabel("-")
        f = self.unit.font(); f.setPointSize(17); f.setBold(True); self.unit.setFont(f)
        self.unit.setStyleSheet(f"color: {NAVY};")
        self.total = QLabel(""); self.total.setStyleSheet("color: #3C4557;")
        v.addWidget(self.unit_cap); v.addWidget(self.unit); v.addWidget(self.total)
        self.cap_text = QLabel("")
        self.cap_text.setStyleSheet("color: #3C4557;")
        self.cap_bar = QProgressBar(); self.cap_bar.setTextVisible(False); self.cap_bar.setFixedHeight(8)
        v.addSpacing(4); v.addWidget(self.cap_text); v.addWidget(self.cap_bar)
        self.show_review(None)

    def show_review(self, review: Optional[rv.Review]) -> None:
        a = review.amounts if review else None
        ch = "" if not review or review.ch is None else f" × {rv._fmt_qty(review.ch)}CH"
        labels = {"pump": f"PUMP{ch}", "rack": "RACK·옵션 합계", "credit": "Credit"}
        if a is not None:
            values = {"pump": fmt_krw(a.pump_amt), "rack": fmt_krw(a.rack_amt),
                      "credit": fmt_krw(a.pump_credit_amt + a.rack_credit_amt)}
            self.unit_cap.setText("CH당 견적단가 (천 원 미만 버림)")
            self.unit.setText(f"{fmt_krw(a.unit)} 원")
            self.total.setText(f"견적금액 {fmt_krw(a.total)} 원" if review.ch is not None
                               else "의뢰 행을 고르면 수량(CH)으로 나눠 계산합니다")
        else:
            values = {k: "-" for k in labels}
            if review is not None:          # 중국·미국 — 합계는 양식의 수식(환율·FOB 등)이 정한다
                real = [r for r in (review.flow + review.base + review.others) if r.specs]
                self.unit_cap.setText(f"{review.qtype} 견적서")
                self.unit.setText(f"품목 {len(real)}칸")
                self.total.setText("합계는 양식의 수식으로 계산됩니다")
            else:
                self.unit_cap.setText("CH당 견적단가"); self.unit.setText("-"); self.total.setText("")
        for key, (cap, val) in self._rows.items():
            cap.setText(labels[key]); val.setText(values[key])
        used, cap_n = (review.used, review.capacity) if review else (0, 0)
        self.cap_text.setText(f"품목 칸 사용  {used} / {cap_n}행" if cap_n else "품목 칸 사용")
        self.cap_bar.setRange(0, max(1, cap_n)); self.cap_bar.setValue(min(used, max(1, cap_n)))
        color = "#C62828" if cap_n and used > cap_n else NAVY
        self.cap_bar.setStyleSheet(
            "QProgressBar { background: #E3E6EC; border: none; border-radius: 4px; }"
            f"QProgressBar::chunk {{ background: {color}; border-radius: 4px; }}")


# ══════════════════════════════════════════════
# 생성 전 점검
# ══════════════════════════════════════════════
class CheckPanel(QFrame):
    pickSpec = Signal(str)
    addSuggest = Signal(str, float)

    def __init__(self, parent=None) -> None:
        super().__init__(parent)
        self.setFrameShape(QFrame.Box); self.setLineWidth(2)
        v = QVBoxLayout(self); v.setContentsMargins(10, 8, 10, 8); v.setSpacing(4)
        self.title = QLabel("생성 전 점검"); f = self.title.font(); f.setBold(True); f.setPointSize(11)
        self.title.setFont(f)
        v.addWidget(self.title)
        self.list = QListWidget(); self.list.setWordWrap(True)
        self.list.setStyleSheet("QListWidget { border: none; background: transparent; }"
                                "QListWidget::item { padding: 5px 6px; margin-bottom: 3px; border-radius: 4px; }")
        self.list.itemClicked.connect(self._on_click)
        v.addWidget(self.list, 1)
        tip = QLabel("항목을 누르면 해당 품목으로 이동하거나 담습니다")
        tip.setStyleSheet("color: #5C6475; font-size: 11px;")
        v.addWidget(tip)
        self.show_review(None)

    def show_review(self, review: Optional[rv.Review]) -> None:
        self.list.clear()
        checks = review.checks if review else []
        self.title.setText(f"생성 전 점검  {len(checks)}건" if checks else "생성 전 점검")
        if not checks:
            it = QListWidgetItem("확인할 항목이 없습니다." if review and (review.flow and any(
                s.specs for s in review.flow + review.base + review.others)) else "품목을 담으면 여기서 점검합니다.")
            it.setFlags(Qt.NoItemFlags); it.setForeground(QBrush(QColor("#5C6475")))
            self.list.addItem(it)
            return
        for c in checks:
            tone, bg, fg = _CHECK_TONE[c.level]
            it = QListWidgetItem(f"{tone} · {c.text}" + ("  [눌러서 담기]" if c.suggest else ""))
            it.setBackground(QBrush(QColor(bg))); it.setForeground(QBrush(QColor(fg)))
            it.setData(Qt.UserRole, c)
            self.list.addItem(it)

    def _on_click(self, item: QListWidgetItem) -> None:
        c = item.data(Qt.UserRole)
        if c is None:
            return
        if c.suggest:
            self.addSuggest.emit(c.suggest[0], float(c.suggest[1]))
        elif c.spec:
            self.pickSpec.emit(c.spec)
