"""
page_quote.py
=============
견적서 작성 페이지와 관련 스레드/다이얼로그.
"""
from __future__ import annotations

import logging
import os
import traceback
from datetime import datetime
from typing import Any, Dict, List, Optional, Tuple

from PySide6.QtCore import Qt, QEvent, QThread, QTimer, Signal
from PySide6.QtGui import QBrush, QColor, QKeySequence, QPixmap, QShortcut
from PySide6.QtWidgets import (
    QAbstractItemView, QCheckBox, QComboBox, QCompleter, QDialog,
    QDialogButtonBox, QDoubleSpinBox, QFileDialog,
    QFormLayout, QFrame, QGridLayout, QHBoxLayout,
    QHeaderView, QInputDialog, QLabel, QLineEdit,
    QListWidget, QListWidgetItem, QMessageBox,
    QProgressDialog, QPushButton, QSizePolicy, QSpinBox, QSplitter, QTabWidget,
    QTableWidget, QTableWidgetItem, QTextEdit,
    QVBoxLayout, QWidget,
)

import app_settings
import quote_config
import quote_review
from quote_panels import AmountPanel, CheckPanel, DiagramPanel, StepBar
from core import (
    QuoteState,
    exe_dir, fmt_krw, fmt_qty,
    normalize_token, parse_invest_info, s, to_float, unique_path,
)
from widgets import (
    DraggableItemsTable, DroppableQuoteTable,
    bold_label, centered_checkbox,
    get_checkbox_from_cell, info_label, labeled_frame, tint_button,
)
from page_common import _friendly_error_msg, _make_plain_table, _ScrollableErrorDialog

logger = logging.getLogger("QuoteApp")


PROCESS_CHOICES = ["", "CVD", "DIFF", "IMP", "ETCH", "METAL", "CLEAN", "GCS"]
VENDOR_CHOICES  = [
    "", "AMAT", "ASM", "AXCELIS", "EUGENETECH", "WONIK_IPS",
    "KOKUSAI", "LAM", "SEMES", "TEL", "TES", "ULVAC", "GCS",
]


# ──────────────────────────────────────────────
# 백그라운드 Excel → PDF 변환 스레드
# ──────────────────────────────────────────────
class _TemplateLoaderThread(QThread):
    """STEP1용: 통합양식 Excel을 백그라운드에서 파싱."""
    done    = Signal(object)   # (rows, by_class, price_by_spec, order, cmap) 또는 Exception
    def __init__(self, path: str, parent=None) -> None:
        super().__init__(parent)
        self.path = path
    def run(self) -> None:
        try:
            import excel_io
            from openpyxl import load_workbook
            # read_only=True: parse_items_sheet/parse_code_map_sheet 는
            # iter_rows() 로만 순회한다 — cell(r,c) 임의 접근을 read_only 모드에
            # 섞으면 O(n²) 로 저하되므로 절대 함께 쓰지 않는다.
            wb = load_workbook(self.path, read_only=True, data_only=False)
            cfg = quote_config.from_workbook(wb)      # 시트 이름·열 위치는 양식의 설정을 따른다
            rows, by_class, price_by_spec, order = excel_io.parse_items_sheet(
                wb[cfg["sheet.items"]], cfg)
            cmap = excel_io.parse_code_map_sheet(wb[cfg["sheet.code_map"]], cfg)
            wb.close()
            self.done.emit((rows, by_class, price_by_spec, order, cmap))
        except Exception as e:
            self.done.emit(e)


class _RequestLoaderThread(QThread):
    """STEP2용: 의뢰파일 Excel을 백그라운드에서 파싱 (메인 스레드 동결 방지)."""
    done = Signal(object)  # (sheet_name, rows) tuple 또는 Exception
    def __init__(self, path: str, parent=None, template_path: Optional[str] = None) -> None:
        super().__init__(parent)
        self.path = path
        self.template_path = template_path
    def run(self) -> None:
        try:
            import excel_io
            sheet_name, rows = excel_io.parse_request_xlsx(
                self.path, quote_config.load(self.template_path))
            self.done.emit((sheet_name, rows))
        except Exception as e:
            self.done.emit(e)


class _GenQuoteThread(QThread):
    """견적서 생성을 백그라운드에서 실행."""
    progress = Signal(str)          # 로그 메시지
    # List[(rd, path, error)] 또는 Exception
    #   성공: error="" / 실패: path="" 이고 error 에 사유
    #   ※ 단건 경로와 멀티 경로가 반드시 같은 모양을 내보내야 한다.
    done     = Signal(object)

    def __init__(self, state, qtype: str, items, rds: list, parent=None) -> None:
        super().__init__(parent)
        self.state = state; self.qtype = qtype
        self.items = items; self.rds   = rds

    def run(self) -> None:
        try:
            import excel_io
            if self.qtype == "국내" and len(self.rds) >= 2:
                def _cb(done, total, name):
                    self.progress.emit(f"[{done}/{total}] 저장: {name}")
                results = excel_io.generate_quote_multi(
                    self.state, self.items, self.rds, progress_cb=_cb)
            else:
                rd   = self.rds[0] if self.rds else {}
                path = excel_io.generate_quote(self.state, self.qtype, self.items, rd)
                self.progress.emit(f"저장: {os.path.basename(path)}")
                results = [(rd, path, "")]
            self.done.emit(results)
        except Exception as e:
            self.done.emit(e)


class _GenCoverThread(QThread):
    """갑지 생성을 백그라운드에서 실행."""
    progress = Signal(str)        # 로그 메시지
    step     = Signal(int, int)   # (done, total) — 진행률 다이얼로그용
    done     = Signal(object)     # (저장 경로, 누락 목록) 튜플 또는 Exception

    def __init__(self, template_path: str, folder: str, paths: list,
                 investor_name: str, warranty_years: int = 2, parent=None) -> None:
        super().__init__(parent)
        self.template_path = template_path; self.folder = folder
        self.paths = paths; self.investor_name = investor_name
        self.warranty_years = warranty_years

    def run(self) -> None:
        try:
            import excel_io
            def _cb(done, total, name):
                self.progress.emit(f"[{done}/{total}] 읽는 중: {name}")
                self.step.emit(done, total)
            out = excel_io.generate_cover(
                self.template_path, self.folder, self.paths,
                investor_name=self.investor_name, progress_cb=_cb,
                warranty_years=self.warranty_years)
            self.done.emit(out)
        except Exception as e:
            self.done.emit(e)


# ══════════════════════════════════════════════
# ══════════════════════════════════════════════
# 중국 옵션 다이얼로그
# ══════════════════════════════════════════════

class ChinaOptionsDialog(QDialog):
    def __init__(self, parent, state: QuoteState) -> None:
        super().__init__(parent)
        self.setWindowTitle("중국 옵션 입력")
        self.state = state
        self.setModal(True)
        self.setFixedWidth(380)

        cn = state.cn_info or {}
        root = QVBoxLayout(self)
        form = QFormLayout()
        self._maker   = QLineEdit(cn.get("maker", ""))
        self._process = QLineEdit(cn.get("process", ""))
        self._tool    = QLineEdit(cn.get("tool", ""))
        self._line    = QLineEdit(cn.get("line", ""))
        self._quote_no = QLineEdit(cn.get("quote_no", ""))
        self._delivery = QLineEdit(cn.get("delivery", ""))
        self._to       = QLineEdit(cn.get("to", ""))
        for w in (self._quote_no, self._delivery, self._to):
            w.setPlaceholderText("비우면 양식에 적힌 그대로")
        form.addRow("Maker",    self._maker)
        form.addRow("Process",  self._process)
        form.addRow("설비호기", self._tool)
        form.addRow("라인",     self._line)
        form.addRow("Quote number",  self._quote_no)
        form.addRow("Delivery date", self._delivery)
        form.addRow("받는 사람(To)", self._to)
        root.addLayout(form)
        bb = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        bb.accepted.connect(self.accept); bb.rejected.connect(self.reject)
        root.addWidget(bb)

    def apply(self) -> None:
        self.state.cn_info = {
            "maker": self._maker.text().strip(), "process": self._process.text().strip(),
            "tool":  self._tool.text().strip(),  "line":    self._line.text().strip(),
            "quote_no": self._quote_no.text().strip(), "delivery": self._delivery.text().strip(),
            "to": self._to.text().strip(),
        }


# ══════════════════════════════════════════════
# 미국 옵션 다이얼로그
# ══════════════════════════════════════════════

class USOptionsDialog(QDialog):
    def __init__(self, parent, state: QuoteState) -> None:
        super().__init__(parent)
        self.setWindowTitle("미국 옵션 입력")
        self.state = state
        self.setModal(True)
        self.setFixedWidth(420)

        us = state.us_info or {}
        root = QVBoxLayout(self)
        form = QFormLayout()
        self._maker   = QLineEdit(us.get("maker", ""))
        self._process = QLineEdit(us.get("process", ""))
        self._tool    = QLineEdit(us.get("tool", ""))
        self._exchange = QDoubleSpinBox()
        self._exchange.setRange(0, 1e9); self._exchange.setDecimals(2)
        self._exchange.setValue(float(us.get("exchange", 0) or 0))
        now = datetime.now()
        y, m = us.get("base_ym", (now.year, now.month))
        self._year  = QSpinBox(); self._year.setRange(2000, 2100); self._year.setValue(int(y))
        self._month = QSpinBox(); self._month.setRange(1, 12);      self._month.setValue(int(m))
        self._site  = QComboBox(); self._site.addItems(["Taylor", "Austin"])
        site = us.get("site", "Taylor")
        self._site.setCurrentText(site if site in ("Taylor", "Austin") else "Taylor")
        ym_row = QWidget()
        ym_h = QHBoxLayout(ym_row); ym_h.setContentsMargins(0,0,0,0)
        ym_h.addWidget(self._year); ym_h.addWidget(QLabel("년"))
        ym_h.addWidget(self._month); ym_h.addWidget(QLabel("월"))
        ym_h.addStretch(1)
        form.addRow("Maker",         self._maker)
        form.addRow("Process",       self._process)
        form.addRow("설비호기",      self._tool)
        form.addRow("환율",          self._exchange)
        form.addRow("기준일(년/월)", ym_row)
        form.addRow("Site",          self._site)
        self._quote_no = QLineEdit(us.get("quote_no", ""))
        self._pr       = QLineEdit(us.get("pr", ""))
        for w in (self._quote_no, self._pr):
            w.setPlaceholderText("비우면 양식에 적힌 그대로")
        form.addRow("Quote number",  self._quote_no)
        form.addRow("PR",            self._pr)
        root.addLayout(form)
        bb = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        bb.accepted.connect(self.accept); bb.rejected.connect(self.reject)
        root.addWidget(bb)

    def apply(self) -> None:
        self.state.us_info = {
            "maker": self._maker.text().strip(), "process": self._process.text().strip(),
            "tool":  self._tool.text().strip(),  "exchange": float(self._exchange.value()),
            "base_ym": (int(self._year.value()), int(self._month.value())),
            "site": self._site.currentText(),
            "quote_no": self._quote_no.text().strip(), "pr": self._pr.text().strip(),
        }


# ══════════════════════════════════════════════
# STEP5 테이블 관리 믹스인
# ══════════════════════════════════════════════

class Step5Manager:
    """
    STEP5 테이블의 행 CRUD + TOTAL/CREDIT 행 관리.
    QuoteBuilderPage 가 다중 상속하여 사용한다.

    하위 클래스가 반드시 제공해야 하는 속성:
        self.step5_table      : DroppableQuoteTable
        self.chk_step5_all    : QCheckBox
        self.state            : QuoteState
        self._step4_highlight : set
    하위 클래스가 구현해야 하는 메서드:
        self._render_items_table(scroll_top: bool)
    """

    # ── 행 종류 판별 ────────────────────────────
    def _row_role(self, row: int) -> str:
        it = self.step5_table.item(row, 1)
        return s(it.data(Qt.UserRole)) if it else ""

    def _is_total(self, row: int)  -> bool: return self._row_role(row) == "TOTAL"
    def _is_credit(self, row: int) -> bool: return self._row_role(row) == "CREDIT"
    def _is_item(self, row: int)   -> bool: return self._row_role(row) == "ITEM"

    # ── 셀 setter ───────────────────────────────
    def _set_qty(self, row: int, qty: float, blank: bool = False) -> None:
        it = QTableWidgetItem("" if blank else fmt_qty(qty))
        it.setData(Qt.UserRole, qty); it.setFlags(Qt.ItemIsEnabled | Qt.ItemIsSelectable)
        self.step5_table.setItem(row, 3, it)

    def _set_price(self, row: int, price: float, blank: bool = False) -> None:
        it = QTableWidgetItem("" if blank else fmt_krw(price))
        it.setData(Qt.UserRole, price); it.setFlags(Qt.ItemIsEnabled | Qt.ItemIsSelectable)
        self.step5_table.setItem(row, 4, it)

    def _set_amt(self, row: int, amt: float) -> None:
        it = QTableWidgetItem(fmt_krw(amt))
        it.setData(Qt.UserRole, amt); it.setFlags(Qt.ItemIsEnabled | Qt.ItemIsSelectable)
        self.step5_table.setItem(row, 5, it)

    def _get_float(self, row: int, col: int) -> float:
        it = self.step5_table.item(row, col)
        return float(it.data(Qt.UserRole) or 0.0) if it else 0.0

    # ── TOTAL 행 보장 ────────────────────────────
    def _ensure_total(self) -> None:
        n = self.step5_table.rowCount()
        if n == 0 or not self._is_total(n - 1):
            self._append_total()

    def _append_total(self) -> None:
        r = self.step5_table.rowCount()
        self.step5_table.insertRow(r)
        it_cat = QTableWidgetItem("TOTAL")
        it_cat.setData(Qt.UserRole, "TOTAL"); it_cat.setFlags(Qt.ItemIsEnabled | Qt.ItemIsSelectable)
        self.step5_table.setItem(r, 1, it_cat)
        it_spec = QTableWidgetItem("총 합계")
        it_spec.setFlags(Qt.ItemIsEnabled | Qt.ItemIsSelectable)
        self.step5_table.setItem(r, 2, it_spec)
        self._set_qty(r, 0, blank=True); self._set_price(r, 0, blank=True); self._set_amt(r, 0)

    # ── 합계 재계산 ──────────────────────────────
    def _recalc_totals(self) -> None:
        self._ensure_total()
        total_qty = total_amt = 0.0
        last = self.step5_table.rowCount() - 1
        for r in range(last):
            if self._is_total(r): continue
            total_amt += self._get_float(r, 5)
            if not self._is_credit(r):
                total_qty += self._get_float(r, 3)
        self._set_qty(last, total_qty, blank=True)
        self._set_amt(last, total_amt)

    def _recalc_row(self, row: int) -> None:
        self._set_amt(row, self._get_float(row, 3) * self._get_float(row, 4))

    # ── 행 추가 ─────────────────────────────────
    def _has_spec(self, spec: str) -> bool:
        """STEP5 에 이미 같은 spec 행이 있는지 확인 (TOTAL/CREDIT 행 제외)."""
        last = self.step5_table.rowCount() - 1
        for r in range(last):
            if self._is_total(r) or self._is_credit(r): continue
            it = self.step5_table.item(r, 2)
            if it and it.text().strip() == spec.strip():
                return True
        return False

    def _add_row(self, category: str, spec: str, qty: float, unit_price: float) -> None:
        """같은 spec 이면 수량 누적, 없으면 TOTAL 행 위에 새 행 삽입."""
        self._ensure_total()
        last = self.step5_table.rowCount() - 1
        for r in range(last):
            if self._is_total(r) or self._is_credit(r): continue
            it = self.step5_table.item(r, 2)
            if it and it.text().strip() == spec.strip():
                self._set_qty(r, self._get_float(r, 3) + qty)
                self._recalc_row(r); self._recalc_totals(); return

        ins = last
        self.step5_table.insertRow(ins)
        cb = QCheckBox()
        cb.stateChanged.connect(self._on_step5_check_changed)
        self.step5_table.setCellWidget(ins, 0, cb)
        it_cat = QTableWidgetItem(category)
        it_cat.setData(Qt.UserRole, "ITEM"); it_cat.setFlags(Qt.ItemIsEnabled | Qt.ItemIsSelectable)
        self.step5_table.setItem(ins, 1, it_cat)
        it_spec = QTableWidgetItem(spec)
        it_spec.setFlags(Qt.ItemIsEnabled | Qt.ItemIsSelectable)
        self.step5_table.setItem(ins, 2, it_spec)
        self._set_qty(ins, qty); self._set_price(ins, unit_price)
        self._set_amt(ins, qty * unit_price); self._recalc_totals()
        self._sort_step5_items()

    # ── Credit 행 ────────────────────────────────
    def _clear_credit_rows(self) -> None:
        for r in range(self.step5_table.rowCount() - 1, -1, -1):
            if self._is_credit(r): self.step5_table.removeRow(r)
        self._ensure_total()

    def _add_credit_row(self, label: str, amount: float) -> None:
        self._ensure_total()
        ins = self.step5_table.rowCount() - 1
        self.step5_table.insertRow(ins)
        red = QBrush(QColor(200, 0, 0))
        it_cat = QTableWidgetItem("CREDIT")
        it_cat.setData(Qt.UserRole, "CREDIT"); it_cat.setFlags(Qt.ItemIsEnabled | Qt.ItemIsSelectable)
        self.step5_table.setItem(ins, 1, it_cat)
        it_spec = QTableWidgetItem(label)
        it_spec.setFlags(Qt.ItemIsEnabled | Qt.ItemIsSelectable); it_spec.setForeground(red)
        self.step5_table.setItem(ins, 2, it_spec)
        self.step5_table.setItem(ins, 3, QTableWidgetItem(""))
        self.step5_table.setItem(ins, 4, QTableWidgetItem(""))
        it_amt = QTableWidgetItem(fmt_krw(amount))
        it_amt.setData(Qt.UserRole, float(amount)); it_amt.setFlags(Qt.ItemIsEnabled | Qt.ItemIsSelectable)
        it_amt.setForeground(red); self.step5_table.setItem(ins, 5, it_amt)

    def _refresh_credits(self) -> None:
        self._clear_credit_rows()
        if self.state.pump_credit: self._add_credit_row("Pump Credit", -abs(self.state.pump_credit))
        if self.state.rack_credit: self._add_credit_row("Rack Credit", -abs(self.state.rack_credit))
        self._refresh_options_display()

    # ── STEP5 정렬 (STEP4 품목 순서 기준) ────────
    def _sort_step5_items(self) -> None:
        """ITEM 행을 self._spec_order 순서에 따라 제자리 정렬."""
        spec_order: Dict[str, int] = getattr(self, "_spec_order", {})
        if not spec_order:
            return
        item_rows = [r for r in range(self.step5_table.rowCount()) if self._is_item(r)]
        if len(item_rows) <= 1:
            return
        # 각 행 데이터 스냅샷
        snaps = []
        for r in item_rows:
            cb = self.step5_table.cellWidget(r, 0)
            cat_it  = self.step5_table.item(r, 1)
            spec_it = self.step5_table.item(r, 2)
            snaps.append({
                "checked": cb.isChecked() if isinstance(cb, QCheckBox) else False,
                "cat":     s(cat_it.text())  if cat_it  else "",
                "spec":    s(spec_it.text()) if spec_it else "",
                "qty":     self._get_float(r, 3),
                "price":   self._get_float(r, 4),
                "amt":     self._get_float(r, 5),
            })
        snaps.sort(key=lambda x: (0 if x["cat"] == "PUMP" else 1, spec_order.get(x["spec"], 10**9)))
        # 정렬된 데이터를 같은 위치에 재기입 (행 삽입/삭제 없이)
        self.step5_table.blockSignals(True)
        for r, d in zip(item_rows, snaps):
            cb = self.step5_table.cellWidget(r, 0)
            if isinstance(cb, QCheckBox):
                cb.blockSignals(True); cb.setChecked(d["checked"]); cb.blockSignals(False)
            cat_it  = self.step5_table.item(r, 1)
            spec_it = self.step5_table.item(r, 2)
            if cat_it:  cat_it.setText(d["cat"])
            if spec_it: spec_it.setText(d["spec"])
            self._set_qty(r, d["qty"])
            self._set_price(r, d["price"])
            self._set_amt(r, d["amt"])
        self.step5_table.blockSignals(False)

    # ── STEP5 스냅샷 ─────────────────────────────
    def _snapshot_items(self) -> List[Dict[str, Any]]:
        rows = []
        for r in range(self.step5_table.rowCount()):
            if self._is_total(r): continue
            cat_it  = self.step5_table.item(r, 1)
            spec_it = self.step5_table.item(r, 2)
            rows.append({
                "role":  s(cat_it.data(Qt.UserRole)) if cat_it else "",
                "cat":   s(cat_it.text())             if cat_it else "",
                "spec":  s(spec_it.text())            if spec_it else "",
                "qty":   self._get_float(r, 3),
                "price": self._get_float(r, 4),
                "amt":   self._get_float(r, 5),
            })
        return rows

    # ── 전체선택 동기화 ──────────────────────────
    def _sync_step5_all_chk(self) -> None:
        selectable = checked = 0
        last = self.step5_table.rowCount() - 1
        for r in range(last + 1):
            if self._is_total(r) or self._is_credit(r): continue
            cb = self.step5_table.cellWidget(r, 0)
            if isinstance(cb, QCheckBox):
                selectable += 1
                if cb.isChecked(): checked += 1
        self.chk_step5_all.blockSignals(True)
        if   selectable == 0 or checked == 0: self.chk_step5_all.setCheckState(Qt.Unchecked)
        elif checked == selectable:            self.chk_step5_all.setCheckState(Qt.Checked)
        else:                                  self.chk_step5_all.setCheckState(Qt.PartiallyChecked)
        self.chk_step5_all.blockSignals(False)

    def _on_step5_check_changed(self, _state: int) -> None:
        """개별 체크박스 변경 → 전체선택 상태 동기화."""
        self._sync_step5_all_chk()

    def _on_step5_select_all(self, checked: bool) -> None:
        """전체선택 체크박스 핸들러."""
        last = self.step5_table.rowCount() - 1
        for r in range(last + 1):
            if self._is_total(r) or self._is_credit(r): continue
            cb = self.step5_table.cellWidget(r, 0)
            if isinstance(cb, QCheckBox):
                cb.blockSignals(True); cb.setChecked(checked); cb.blockSignals(False)
        self._sync_step5_all_chk()

    # ── STEP4 하이라이트 동기화 ──────────────────
    def _sync_step4_highlight(self, scroll_top: bool = True) -> None:
        transferred: set = set()
        for r in range(self.step5_table.rowCount()):
            if self._is_total(r) or self._is_credit(r): continue
            it = self.step5_table.item(r, 2)
            if it and it.text().strip():
                transferred.add(it.text().strip())
        self._step4_highlight = transferred
        self._render_items_table(scroll_top=scroll_top)

    def _render_items_table(self, scroll_top: bool = True) -> None:
        """하위 클래스(QuoteBuilderPage)에서 구현."""
        raise NotImplementedError("_render_items_table must be implemented by subclass")


# ══════════════════════════════════════════════
# 견적서 LIST 독립 창
# ══════════════════════════════════════════════

class _QuoteListWindow(QWidget):
    """견적서 LIST를 보여주는 독립(플로팅) 창."""

    def __init__(self, parent=None):
        super().__init__(parent, Qt.Window)
        self.setWindowTitle("견적서 LIST")
        self.resize(480, 620)
        # (tool, label, full_path) — 추가 시점에 한 번만 파싱해 캐시
        self._session_entries: List[Tuple[str, str, str]] = []
        self._build()

    def _build(self) -> None:
        v = QVBoxLayout(self)
        v.setContentsMargins(12, 12, 12, 12); v.setSpacing(6)

        title_row = QHBoxLayout()
        lbl = QLabel("견적서 LIST")
        f = lbl.font(); f.setPointSize(12); f.setBold(True); lbl.setFont(f)
        title_row.addWidget(lbl)
        v.addLayout(title_row)

        filter_row = QHBoxLayout(); filter_row.setSpacing(6)
        lbl_f = QLabel("설비호기:"); lbl_f.setFixedWidth(52)
        self.edit_filter = QLineEdit()
        self.edit_filter.setPlaceholderText("설비호기 입력 → 실시간 필터")
        self.edit_filter.setFixedHeight(24)
        self.edit_filter.setClearButtonEnabled(True)
        self.edit_filter.textChanged.connect(self._apply_filter)
        filter_row.addWidget(lbl_f); filter_row.addWidget(self.edit_filter)
        v.addLayout(filter_row)

        self.list_widget = QListWidget()
        self.list_widget.itemDoubleClicked.connect(self._open_item)
        v.addWidget(self.list_widget, 1)

    def add_item(self, path: str) -> None:
        entry = QuoteBuilderPage._parse_quote_entry(path)
        if entry is not None:
            self._session_entries.insert(0, entry)
        self._apply_filter()

    def add_submit_item(self, path: str) -> None:
        """제출용 엑셀 항목을 '(제출용) 레이블' 형식으로 추가."""
        _, label = QuoteBuilderPage._parse_quote_label(path)
        display_label = f"(제출용) {label}" if label else f"(제출용) {os.path.basename(path)}"
        self._session_entries.insert(0, ("__submit__", display_label, path))
        self._apply_filter()

    def refresh_session(self, files: List[str],
                        submit_paths: frozenset = frozenset()) -> None:
        self._session_entries = []
        for f in files:
            if f in submit_paths:
                _, label = QuoteBuilderPage._parse_quote_label(f)
                display_label = f"(제출용) {label}" if label else f"(제출용) {os.path.basename(f)}"
                self._session_entries.append(("__submit__", display_label, f))
            else:
                e = QuoteBuilderPage._parse_quote_entry(f)
                if e is not None:
                    self._session_entries.append(e)
        self._apply_filter()

    def _apply_filter(self) -> None:
        keyword = self.edit_filter.text().strip().lower()
        self.list_widget.clear()

        filtered = [
            (tool, label, full)
            for tool, label, full in self._session_entries
            if not keyword or keyword in label.lower()
        ]

        # 중복 레이블 카운트: 두 번째부터 (2), (3) … 표시
        label_counts: Dict[str, int] = {}
        for _, label, _ in filtered:
            label_counts[label] = label_counts.get(label, 0) + 1
        label_seen: Dict[str, int] = {}

        for tool, label, full in filtered:
            if label_counts[label] > 1:
                label_seen[label] = label_seen.get(label, 0) + 1
                n = label_seen[label]
                display = label if n == 1 else f"{label} ({n})"
            else:
                display = label
            item = QListWidgetItem(display)
            item.setData(Qt.UserRole, full)
            item.setToolTip(full)
            if tool == "":
                item.setBackground(QBrush(QColor(255, 153, 153)))   # 갑지: 연빨강
            elif tool == "__submit__":
                item.setBackground(QBrush(QColor(255, 243, 160)))   # 제출용: 연노랑
            self.list_widget.addItem(item)

    def _open_item(self, item: QListWidgetItem) -> None:
        path = item.data(Qt.UserRole)
        if path:
            try: os.startfile(path)
            except Exception as e: QMessageBox.critical(self, "오류", str(e))

    def show_and_raise(self) -> None:
        self.show()
        self.raise_()
        self.activateWindow()
        if self.list_widget.count() > 0:
            self.list_widget.scrollToTop()


# ══════════════════════════════════════════════
# 견적서 작성 페이지
# ══════════════════════════════════════════════

class QuoteBuilderPage(Step5Manager, QWidget):

    def __init__(self) -> None:
        QWidget.__init__(self)
        self.state = QuoteState()

        self._step4_highlight     : set                     = set()
        self._filtered_items      : List[Tuple[str, float]] = []
        self._spec_order          : Dict[str, int]          = {}
        self._done_req_indices    : set                     = set()   # 견적 완료된 의뢰행 인덱스
        self._pending_req_indices : List[int]               = []      # 현재 생성 중인 의뢰행 인덱스

        self._filter_timer  : QTimer                          = QTimer(self)
        self._filter_timer.setSingleShot(True)
        self._filter_timer.timeout.connect(self._do_filter)
        self._tpl_thread    : Optional[_TemplateLoaderThread]  = None
        self._req_thread    : Optional[_RequestLoaderThread]   = None
        self._gen_qt_thread : Optional[_GenQuoteThread]        = None
        self._gen_cv_thread  : Optional[_GenCoverThread]        = None
        self._cover_progress : Optional[QProgressDialog]      = None
        self._list_all_files : List[str]                       = []
        self._submit_paths   : set                             = set()
        self._quote_list_window: Optional[_QuoteListWindow]    = None

        self._notices         : List[str]                      = []      # 양식 알림 → 점검 목록의 '참고'
        self._dup_price_specs : set                            = set()   # 품목 시트에 단가가 둘 이상인 규격
        self._class_of        : Dict[str, str]                 = {}      # 규격 → 품목 시트의 분류
        self._review          : Optional[quote_review.Review]  = None
        self._review_timer    : QTimer                         = QTimer(self)
        self._review_timer.setSingleShot(True)
        self._review_timer.setInterval(60)
        self._review_timer.timeout.connect(self._refresh_review)

        self._build_ui()
        self._ensure_total()
        self._log("준비: STEP1 → STEP2 → STEP4/드래그 → STEP5 → 생성")
        # STEP5 가 어떻게 바뀌든(추가·삭제·수량·Credit·정렬) 검토 화면을 다시 계산한다
        m = self.step5_table.model()
        for sig in (m.rowsInserted, m.rowsRemoved, m.dataChanged, m.modelReset):
            sig.connect(self._schedule_review)
        self.req_table.currentCellChanged.connect(self._schedule_review)
        self._refresh_review()

        sc = QShortcut(QKeySequence.Delete, self)
        sc.activated.connect(self._delete_checked_rows)

        self._restore_settings()

    # ══════════════════════════════════════════
    # 설정 복원 / 저장
    # ══════════════════════════════════════════
    def _restore_settings(self) -> None:
        """이전 실행에서 저장한 설정을 복원한다.

        옵션(Credit/보증기간/투자자/설비수량/실반입라인/Exh Size)은 통합양식과
        달리 매 실행마다 초기화 상태로 시작한다 — 의도적으로 복원하지 않는다.
        """
        self.state.last_output_dir = (
            app_settings.get_dir(app_settings.Key.OUTPUT_DIR) or None)

        # 마지막 통합양식이 그대로 있으면 자동으로 다시 읽는다.
        # (백그라운드 스레드이므로 창 표시를 막지 않는다)
        tpl = app_settings.get_path(app_settings.Key.TEMPLATE_PATH)
        if tpl:
            self._log(f"이전 통합양식 자동 로드: {os.path.basename(tpl)}")
            QTimer.singleShot(0, lambda: self._start_template_load(tpl, silent=True))
        else:
            # 파일이 옮겨졌거나(버전 업) 네트워크 드라이브가 아직 안 붙은 경우.
            # 설정은 지우지 않는다 — 다음 실행에 드라이브가 붙으면 다시 쓴다.
            prev = app_settings.get_str(app_settings.Key.TEMPLATE_PATH)
            if prev:
                self._log(f"이전 통합양식을 찾을 수 없어 자동 로드 건너뜀: {prev}")

    # ══════════════════════════════════════════
    # UI 구성
    # ══════════════════════════════════════════

    def _build_ui(self) -> None:
        """배치: [의뢰 목록(좁게)] | [작업 영역: 요약·조건·도식·STEP5] | [금액·점검 / 옵션 탭]

        계속 보는 것(도식·표·금액·점검)에 화면 대부분을 준다. 가끔만 쓰는 것은
        자리를 줄였다 — 의뢰 목록은 호기·수량만 보이고 나머지는 툴팁으로,
        품목 목록(STEP4)은 별도 창으로, 옵션은 오른쪽 탭으로.
        """
        outer = QVBoxLayout(self)
        outer.setContentsMargins(14, 10, 14, 12)
        outer.setSpacing(8)

        # ── 단계 표시 (어디까지 왔고 다음에 무엇을 하는지) ───────────
        self.step_bar = StepBar()
        outer.addWidget(self.step_bar)

        # ── 툴바 (STEP1·2 LOAD + 생성버튼 + LIST 버튼) ──────────────
        outer.addWidget(self._build_toolbar())

        content = QWidget()
        ch = QHBoxLayout(content)
        ch.setContentsMargins(0, 0, 0, 0); ch.setSpacing(10)

        # 왼쪽: 의뢰 목록 (좁게)
        rail = self._build_req_panel()
        rail.setFixedWidth(250)
        ch.addWidget(rail)

        # 가운데: 요약 한 줄 + 조건 입력 한 줄 + 구성 도식 + STEP5 표
        center = QWidget()
        cv = QVBoxLayout(center); cv.setContentsMargins(0, 0, 0, 0); cv.setSpacing(8)
        cv.addWidget(self._build_quote_info_panel())
        self._step3_frame = self._build_step3()
        self._step3_frame.setEnabled(False)
        cv.addWidget(self._step3_frame)
        self.diagram = DiagramPanel()
        self.diagram.pickSpec.connect(self._select_step5_spec)
        self.diagram.addSuggest.connect(self._add_suggested)
        cv.addWidget(self.diagram)
        cv.addWidget(self._build_step5_panel(), 1)
        ch.addWidget(center, 1)

        # 오른쪽: 금액·점검 / 옵션 탭
        self.side_tabs = QTabWidget()
        self.side_tabs.setFixedWidth(350)
        review_tab = QWidget()
        rv_l = QVBoxLayout(review_tab); rv_l.setContentsMargins(6, 8, 6, 6); rv_l.setSpacing(8)
        self.amount_panel = AmountPanel()
        self.check_panel = CheckPanel()
        self.check_panel.pickSpec.connect(self._select_step5_spec)
        self.check_panel.addSuggest.connect(self._add_suggested)
        rv_l.addWidget(self.amount_panel)
        rv_l.addWidget(self.check_panel, 1)
        self.side_tabs.addTab(review_tab, "금액·점검")
        opt_tab = QWidget()
        op_l = QVBoxLayout(opt_tab); op_l.setContentsMargins(6, 8, 6, 6); op_l.setSpacing(8)
        op_l.addWidget(self._build_options_panel())
        op_l.addStretch(1)
        self.side_tabs.addTab(opt_tab, "옵션")
        ch.addWidget(self.side_tabs)

        outer.addWidget(content, 1)

        # 품목 목록(STEP4)은 별도 창 — '+ 품목 추가' 로 연다
        self._picker = QWidget(self, Qt.Window)
        self._picker.setWindowTitle("품목 추가 — 더블클릭하거나 견적작성 대상 표로 끌어 놓으세요")
        pk = QVBoxLayout(self._picker); pk.setContentsMargins(8, 8, 8, 8)
        pk.addWidget(self._build_step4_panel())
        self._picker.resize(520, 760)
        self._picker_placed = False

    def _open_item_picker(self) -> None:
        """품목 목록 창을 연다(이미 열려 있으면 앞으로). 처음에는 본 창 오른쪽 위에 붙인다."""
        if not self._picker_placed:
            top = self.window().frameGeometry()
            self._picker.move(max(0, top.right() - self._picker.width() - 380), top.top() + 120)
            self._picker_placed = True
        self._picker.show(); self._picker.raise_(); self._picker.activateWindow()
        self.ed_filter.setFocus(); self.ed_filter.selectAll()

    @staticmethod
    def _action_btn(label: str, slot) -> QPushButton:
        btn = QPushButton(label); btn.setFixedHeight(52)
        f = btn.font(); f.setPointSize(10); f.setBold(True); btn.setFont(f)
        btn.clicked.connect(slot); return btn

    # ── 툴바 ──────────────────────────────────
    def _build_toolbar(self) -> QWidget:
        bar = QWidget(); bar.setFixedHeight(58)
        h = QHBoxLayout(bar); h.setContentsMargins(0, 0, 0, 0); h.setSpacing(6)

        self.btn_step1 = self._action_btn("STEP 1\n통합양식 LOAD", self._load_template)
        self.btn_step2 = self._action_btn("STEP 2\n의뢰파일 LOAD", self._load_request)
        tint_button(self.btn_step1, "#C8E6C9")
        tint_button(self.btn_step2, "#BBDEFB")
        self.btn_step2.setEnabled(False)

        self.btn_gen_quote  = self._action_btn("견적서 생성",    self._generate_quote)
        self.btn_gen_cover  = self._action_btn("갑지 생성",      self._generate_cover)
        self.btn_submit_xls = self._action_btn("제출용 엑셀 생성", self._export_submit_excel)
        tint_button(self.btn_gen_quote,  "#B3E5FC")
        tint_button(self.btn_gen_cover,  "#F8BBD0")
        tint_button(self.btn_submit_xls, "#FFF9C4")

        btn_list = QPushButton("견적서 LIST"); btn_list.setFixedHeight(52)
        f = btn_list.font(); f.setPointSize(10); f.setBold(True); btn_list.setFont(f)
        tint_button(btn_list, "#E8EAF6")
        btn_list.clicked.connect(self._open_quote_list_window)

        sep = QFrame(); sep.setFrameShape(QFrame.VLine); sep.setFrameShadow(QFrame.Sunken)

        h.addWidget(self.btn_step1)
        h.addWidget(self.btn_step2)
        h.addWidget(sep)
        h.addWidget(self.btn_gen_quote)
        h.addWidget(self.btn_gen_cover)
        h.addWidget(self.btn_submit_xls)
        h.addWidget(btn_list)
        self.lbl_last = QLabel("")          # 방금 한 일 (작업로그 창을 대신한다)
        self.lbl_last.setStyleSheet("color: #5C6475;")
        self.lbl_last.setAlignment(Qt.AlignRight | Qt.AlignVCenter)
        h.addWidget(self.lbl_last, 1)
        btn_cfg = QPushButton("설정 시트 만들기"); btn_cfg.setFixedHeight(34)
        btn_cfg.setToolTip("양식의 시트 이름·칸 위치·문구·파일 이름 규칙을 적은 '설정' 시트를 만듭니다.\n"
                           "통합양식에 이 시트를 복사해 넣고 값을 고치면 프로그램이 그대로 따릅니다.")
        btn_cfg.clicked.connect(self._export_settings_sheet)
        h.addWidget(btn_cfg)
        return bar

    def _export_settings_sheet(self) -> None:
        cfg = quote_config.load(self.state.template_path)
        start = os.path.join(self.state.last_output_dir or exe_dir(), "견적서작성_설정.xlsx")
        path, _ = QFileDialog.getSaveFileName(self, "설정 시트 저장", start, "Excel Files (*.xlsx)")
        if not path:
            return
        if not path.lower().endswith(".xlsx"):
            path += ".xlsx"
        try:
            quote_config.export_settings_sheet(path, cfg)
        except Exception as e:
            QMessageBox.critical(self, "오류", f"설정 시트를 만들지 못했습니다.\n{e}"); return
        self._log(f"설정 시트 생성: {path}")
        QMessageBox.information(
            self, "완료",
            f"설정 시트를 만들었습니다.\n{path}\n\n"
            f"이 파일의 '{quote_config.SETTINGS_SHEET}' 시트를 통합양식에 복사해 넣고,\n"
            "바꿀 항목의 '값'만 고치면 됩니다. (비워 두거나 지운 항목은 기본값)")

    # ── STEP3 ─────────────────────────────────
    def _build_step3(self) -> QFrame:
        """조건 입력 — 한 줄. 의뢰 행을 더블클릭하면 자동으로 채워지는 값이라 자리를 크게 쓰지 않는다."""
        frame = QFrame(); frame.setFrameShape(QFrame.Box); frame.setLineWidth(2)
        row = QHBoxLayout(frame); row.setContentsMargins(10, 6, 10, 6); row.setSpacing(8)
        self.cb_process = QComboBox(); self.cb_process.addItems(PROCESS_CHOICES)
        self.cb_vendor  = QComboBox(); self.cb_vendor.addItems(VENDOR_CHOICES)
        self.ed_code = QLineEdit(); self.ed_code.setPlaceholderText("5D")
        self.btn_apply_map = QPushButton("코드매핑 적용"); self.btn_apply_map.setEnabled(False)
        self.btn_apply_map.setToolTip("이 공정·설비사·5D 에 등록된 RACK 품목을 견적작성 대상에 담습니다.")
        tint_button(self.btn_apply_map, "#FFE0B2")  # 연주황
        self.btn_gen_5d_obj = QPushButton("5D Object 생성")
        tint_button(self.btn_gen_5d_obj, "#C5CAE9")  # 연남색
        for b in (self.btn_apply_map, self.btn_gen_5d_obj):
            b.setFixedHeight(30)
        row.addWidget(bold_label("조건", size=10))
        row.addWidget(QLabel("공정")); row.addWidget(self.cb_process, 2)
        row.addWidget(QLabel("설비사")); row.addWidget(self.cb_vendor, 2)
        row.addWidget(QLabel("5D")); row.addWidget(self.ed_code, 3)
        row.addWidget(self.btn_apply_map); row.addWidget(self.btn_gen_5d_obj)
        self.cb_process.currentTextChanged.connect(self._on_step3_changed)
        self.cb_vendor.currentTextChanged.connect(self._on_step3_changed)
        self.ed_code.textChanged.connect(self._on_step3_changed)
        self.btn_apply_map.clicked.connect(self._apply_code_map)
        self.btn_gen_5d_obj.clicked.connect(self._gen_5d_object)
        return frame

    # ── 의뢰파일 패널 ─────────────────────────
    _REQ_HIDDEN_COLS = (2, 4, 5, 6, 7)      # 투자정보·Rack·Maker·설비·5D → 툴팁으로

    def _build_req_panel(self) -> QFrame:
        frame = QFrame(); frame.setFrameShape(QFrame.Box); frame.setLineWidth(2)
        frame.setMinimumHeight(180)
        v = QVBoxLayout(frame); v.setContentsMargins(8,8,8,8); v.setSpacing(6)
        v.addWidget(bold_label("의뢰 목록", size=11))
        top = QHBoxLayout()
        self.chk_req_all = QCheckBox("전체선택"); self.chk_req_all.clicked.connect(self._on_req_select_all)
        top.addWidget(self.chk_req_all); top.addStretch(1); v.addLayout(top)
        self.req_table = _make_plain_table(
            8, ["", "설비호기", "투자정보", "수량", "Rack", "Maker", "설비", "5D"])
        self.req_table.setColumnWidth(0, 36)
        _rh = self.req_table.horizontalHeader()
        _rh.setSectionResizeMode(0, QHeaderView.Fixed)
        for _ci in range(1, 8):
            _rh.setSectionResizeMode(_ci, QHeaderView.Interactive)
        self.req_table.setHorizontalScrollBarPolicy(Qt.ScrollBarAsNeeded)
        self.req_table.cellDoubleClicked.connect(self._on_req_double_click)
        self._style_req_table()
        # 좁은 목록에는 호기와 수량만 — 나머지는 행에 마우스를 올리면 툴팁으로 보인다
        for _ci in self._REQ_HIDDEN_COLS:
            self.req_table.setColumnHidden(_ci, True)
        _rh.setSectionResizeMode(1, QHeaderView.Stretch)
        v.addWidget(self.req_table, 1)
        tip = QLabel("행에 마우스를 올리면 전체 정보가 보입니다.\n더블클릭하면 그 의뢰로 채워집니다.")
        tip.setStyleSheet("color: #5C6475; font-size: 11px;"); tip.setWordWrap(True)
        v.addWidget(tip)
        return frame

    # ── STEP4 패널 ────────────────────────────
    def _build_step4_panel(self) -> QFrame:
        frame = QFrame(); frame.setFrameShape(QFrame.Box); frame.setLineWidth(2); frame.setMinimumHeight(150)
        v = QVBoxLayout(frame); v.setContentsMargins(12,12,12,12); v.setSpacing(8)
        v.addWidget(bold_label("STEP 4\n품목(규격/단가) + 필터  (더블클릭/드래그 → STEP5)"))
        self.ed_filter = QLineEdit(); self.ed_filter.setPlaceholderText("규격 필터(실시간) - 포함 검색")
        self.ed_filter.textChanged.connect(self._apply_filter)
        self.items_table = DraggableItemsTable(0, 2)
        self.items_table.setHorizontalHeaderLabels(["규격", "단가(원)"])
        self.items_table.verticalHeader().setVisible(False)
        self.items_table.setAlternatingRowColors(True)
        self.items_table.setEditTriggers(DraggableItemsTable.NoEditTriggers)
        self.items_table.setSelectionBehavior(QAbstractItemView.SelectRows)
        self.items_table.setSelectionMode(QAbstractItemView.SingleSelection)
        self.items_table.setDragEnabled(True); self.items_table.setWordWrap(False)
        self.items_table.setStyleSheet("QTableWidget::item:selected{color:black;}")
        h = self.items_table.horizontalHeader()
        h.setSectionResizeMode(0, QHeaderView.Stretch); h.setSectionResizeMode(1, QHeaderView.Fixed)
        self.items_table.setColumnWidth(1, 80)
        self.items_table.cellDoubleClicked.connect(self._on_item_double_click)
        v.addWidget(self.ed_filter); v.addWidget(self.items_table, 1)
        return frame

    # ══════════════════════════════════════════
    # 검토 (구성 도식 · 금액 · 생성 전 점검)
    # ══════════════════════════════════════════
    def _schedule_review(self, *_args) -> None:
        """잠깐 모았다가 한 번만 다시 계산한다(행 하나 추가에도 신호가 여러 번 온다)."""
        self._review_timer.start()

    def _make_review(self) -> quote_review.Review:
        st = self.state
        suggested = st.code_map.get((st.process, st.vendor, st.code_5d), []) if st.code_map else []
        ch = self._pick_rd().get("H") if st.request_rows else None
        review = quote_review.analyze(
            self._snapshot_items(), st.quote_type or "국내", ch,
            cfg=quote_config.load(st.template_path),
            class_of=self._class_of,
            known_specs=set(st.price_by_spec) if st.template_path else None,
            dup_price_specs=self._dup_price_specs,
            suggested=suggested)
        review.checks.extend(quote_review.Check(quote_review.INFO, n) for n in self._notices)
        return review

    def _refresh_review(self) -> None:
        try:
            review = self._make_review()
        except Exception:
            logger.warning("검토 화면 계산 실패", exc_info=True)
            return
        self._review = review
        self._refresh_opt_summary()
        self.diagram.show_review(review)
        self.amount_panel.show_review(review)
        self.check_panel.show_review(review)
        self._refresh_step_bar(review)

    def _refresh_step_bar(self, review: quote_review.Review) -> None:
        st = self.state
        has_items = any(s.specs for s in review.flow + review.base + review.others)
        done = [bool(st.template_path), bool(st.request_rows),
                has_items and not review.red, bool(self._list_all_files)]
        if not done[0]:
            hint = "STEP 1 에서 통합양식을 불러오세요"
        elif not done[1]:
            hint = "STEP 2 에서 의뢰파일을 불러오고, 의뢰 행을 더블클릭하세요"
        elif not has_items:
            hint = "품목을 담으세요 — 의뢰 행 더블클릭, 코드매핑 적용, 또는 STEP 4 에서 더블클릭"
        elif review.red:
            hint = f"판단이 필요한 항목이 {len(review.red)}건 있습니다 — 오른쪽 점검 목록을 확인하세요"
        elif not done[3]:
            hint = "견적서 생성을 누르세요" + (f"  (확인 권장 {len(review.yellow)}건)" if review.yellow else "")
        else:
            hint = "생성 완료 — 다음 의뢰 행을 고르거나 갑지를 생성하세요"
        self.step_bar.set_progress(done, hint)

    def _select_step5_spec(self, spec: str) -> None:
        """도식·점검 목록에서 고른 규격의 행을 STEP5 표에서 선택해 보여준다."""
        for r in range(self.step5_table.rowCount()):
            it = self.step5_table.item(r, 2)
            if it and it.text().strip() == spec.strip() and self._is_item(r):
                self.step5_table.selectRow(r)
                self.step5_table.scrollToItem(it)
                self.step5_table.setFocus()
                return

    def _add_suggested(self, spec: str, qty: float) -> None:
        """코드매핑이 제안했지만 담기지 않은 품목을 담는다."""
        price = float(self.state.price_by_spec.get(spec, 0))
        self._add_row("RACK", spec, qty, price)
        self._sync_step4_highlight(scroll_top=False)
        self._log(f"코드매핑 품목 담음: {spec} × {quote_review._fmt_qty(qty)}")
        self._select_step5_spec(spec)

    def _build_step5_panel(self) -> QFrame:
        step5 = labeled_frame("STEP5 견적작성 대상", min_h=150)
        step5.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Expanding)
        sl = step5.layout()
        top5 = QHBoxLayout()
        self.chk_step5_all = QCheckBox("전체선택")
        self.chk_step5_all.clicked.connect(self._on_step5_select_all)
        top5.addWidget(self.chk_step5_all); top5.addStretch(1)
        self.btn_add_item = QPushButton("+ 품목 추가"); self.btn_add_item.setFixedHeight(30)
        self.btn_add_item.setToolTip("품목 목록 창을 엽니다. 더블클릭하거나 이 표로 끌어 놓아 담습니다.")
        tint_button(self.btn_add_item, "#DCE6F7")
        self.btn_add_item.clicked.connect(self._open_item_picker)
        top5.addWidget(self.btn_add_item)
        sl.addLayout(top5)

        self.step5_table = DroppableQuoteTable()
        self.step5_table.setColumnCount(6)
        self.step5_table.setHorizontalHeaderLabels(["선택","분류","품목","수량","단가(원)","금액(원)"])
        self.step5_table.verticalHeader().setVisible(False)
        self.step5_table.setAlternatingRowColors(True)
        self.step5_table.setSelectionBehavior(QAbstractItemView.SelectRows)
        self.step5_table.setEditTriggers(DroppableQuoteTable.NoEditTriggers)
        self.step5_table.cellDoubleClicked.connect(self._on_step5_double_click)
        self.step5_table.drop_callback = self._on_drop_to_step5
        h5 = self.step5_table.horizontalHeader()
        h5.setSectionResizeMode(0, QHeaderView.Fixed)
        for i in [1, 3, 4, 5]:
            h5.setSectionResizeMode(i, QHeaderView.Fixed)
        h5.setSectionResizeMode(2, QHeaderView.Interactive)
        h5.setStretchLastSection(False)
        self.step5_table.setHorizontalScrollBarPolicy(Qt.ScrollBarAlwaysOff)
        self.step5_table.setColumnWidth(0, 46)
        for c in [1, 3, 4, 5]:
            self.step5_table.resizeColumnToContents(c)
        h5.sectionResized.connect(self._on_step5_col2_resized)
        self.step5_table.viewport().installEventFilter(self)
        QTimer.singleShot(0, self._step5_fit_col2)
        sl.addWidget(self.step5_table, 1)
        return step5

    @staticmethod
    def _vline() -> QFrame:
        ln = QFrame(); ln.setFrameShape(QFrame.VLine); ln.setFrameShadow(QFrame.Sunken)
        return ln

    def _build_options_panel(self) -> QFrame:
        """
        한 줄에 하나씩: 왼쪽 조작(버튼) · 오른쪽 현재 값.
        Credit / 보증기간 / 투자자 / 견적서 타입 / 설비수량 / 실반입라인 / Exh Size
        """
        frame = QFrame()
        frame.setLayout(QVBoxLayout()); frame.layout().setContentsMargins(4, 4, 4, 4)
        row = QVBoxLayout(); row.setSpacing(10)

        def _col(action: QWidget, status: QWidget) -> QHBoxLayout:
            """한 줄: 왼쪽에 조작(버튼), 오른쪽에 지금 값."""
            h = QHBoxLayout(); h.setSpacing(8)
            action.setFixedWidth(120)
            h.addWidget(action); h.addWidget(status, 1)
            return h

        # ── Credit ──────────────────────────────
        self.btn_credit = QPushButton("Credit"); self.btn_credit.setMinimumHeight(30)
        self.btn_credit.clicked.connect(self._open_credit_dialog)
        tint_button(self.btn_credit, "#FFF9C4")   # 연노랑
        self.lbl_credit_status = QLabel(); self.lbl_credit_status.setAlignment(Qt.AlignCenter)
        self.lbl_credit_status.setMinimumHeight(28); self.lbl_credit_status.setWordWrap(True)
        row.addLayout(_col(self.btn_credit, self.lbl_credit_status), 1)

        # ── 보증기간 ────────────────────────────
        self.btn_warranty = QPushButton("보증기간 변경"); self.btn_warranty.setMinimumHeight(30)
        self.btn_warranty.clicked.connect(self._change_warranty)
        tint_button(self.btn_warranty, "#E0F2F1")   # 연민트
        self.lbl_warranty_disp = info_label(""); self.lbl_warranty_disp.setAlignment(Qt.AlignCenter)
        row.addLayout(_col(self.btn_warranty, self.lbl_warranty_disp), 1)

        # ── 투자자 ──────────────────────────────
        self.btn_investor = QPushButton("투자자 변경"); self.btn_investor.setMinimumHeight(30)
        self.btn_investor.clicked.connect(self._change_investor)
        tint_button(self.btn_investor, "#E8EAF6")   # 연라벤더
        self.lbl_investor_disp = info_label(""); self.lbl_investor_disp.setAlignment(Qt.AlignCenter)
        row.addLayout(_col(self.btn_investor, self.lbl_investor_disp), 1)

        # ── 견적서 타입 ─────────────────────────
        lbl_qtype_hdr = bold_label("견적서 타입", size=9)
        lbl_qtype_hdr.setAlignment(Qt.AlignCenter)
        qtype_row = QHBoxLayout(); qtype_row.setSpacing(6); qtype_row.addStretch(1)
        self.chk_qt_kr = QCheckBox("국내"); self.chk_qt_cn = QCheckBox("중국"); self.chk_qt_us = QCheckBox("미국")
        self.chk_qt_kr.setChecked(True)
        for chk, qt in [(self.chk_qt_kr,"국내"),(self.chk_qt_cn,"중국"),(self.chk_qt_us,"미국")]:
            chk.clicked.connect(lambda checked, t=qt: self._on_qt_checked(t, checked))
            qtype_row.addWidget(chk)
        qtype_row.addStretch(1)
        qtype_host = QWidget(); qtype_host.setLayout(qtype_row)
        row.addLayout(_col(lbl_qtype_hdr, qtype_host), 1)

        # ── 설비수량 (국내 사양서!A43) ───────────
        self.btn_equip_qty = QPushButton("설비수량"); self.btn_equip_qty.setMinimumHeight(30)
        self.btn_equip_qty.clicked.connect(self._change_equip_qty)
        tint_button(self.btn_equip_qty, "#FFE0B2")   # 연주황
        self.lbl_equip_qty_disp = info_label(""); self.lbl_equip_qty_disp.setAlignment(Qt.AlignCenter)
        row.addLayout(_col(self.btn_equip_qty, self.lbl_equip_qty_disp), 1)

        # ── 실반입라인 (국내 현업 사인용 사양서!B6) ──
        self.btn_actual_line = QPushButton("실반입라인"); self.btn_actual_line.setMinimumHeight(30)
        self.btn_actual_line.clicked.connect(self._change_actual_line)
        tint_button(self.btn_actual_line, "#D1C4E9")   # 연보라
        self.lbl_actual_line_disp = info_label(""); self.lbl_actual_line_disp.setAlignment(Qt.AlignCenter)
        row.addLayout(_col(self.btn_actual_line, self.lbl_actual_line_disp), 1)

        # ── Exh Size (국내 입고검수확인서!D27) ───
        self.btn_exh_size = QPushButton("Exh Size"); self.btn_exh_size.setMinimumHeight(30)
        self.btn_exh_size.clicked.connect(self._change_exh_size)
        tint_button(self.btn_exh_size, "#B2DFDB")   # 연청록
        self.lbl_exh_size_disp = info_label(""); self.lbl_exh_size_disp.setAlignment(Qt.AlignCenter)
        row.addLayout(_col(self.btn_exh_size, self.lbl_exh_size_disp), 1)

        frame.layout().addLayout(row)
        self._refresh_options_display()
        return frame

    def _change_equip_qty(self) -> None:
        n, ok = QInputDialog.getInt(
            self, "설비수량 변경", "설비수량:",
            self.state.equip_qty, 0, 999999,
        )
        if ok:
            self.state.equip_qty = n
            self._log(f"설비수량: {n}")
            self._refresh_options_display()

    def _change_actual_line(self) -> None:
        text, ok = QInputDialog.getText(
            self, "실반입라인 변경", "실반입라인:",
            text=self.state.actual_line,
        )
        if ok:
            self.state.actual_line = text.strip()
            self._log(f"실반입라인: {self.state.actual_line or '-'}")
            self._refresh_options_display()

    def _change_exh_size(self) -> None:
        n, ok = QInputDialog.getInt(
            self, "Exh Size 변경", "Exh Size:",
            self.state.exh_size, 0, 999999,
        )
        if ok:
            self.state.exh_size = n
            self._log(f"Exh Size: {n}")
            self._refresh_options_display()

    def _refresh_options_display(self) -> None:
        """옵션 패널 2번째 줄(현재 상태 표기)을 state 값에 맞춰 갱신한다."""
        self.lbl_warranty_disp.setText(f"{self.state.warranty_years} Year after delivery")
        self.lbl_investor_disp.setText(self.state.investor_name or "-")
        self.lbl_equip_qty_disp.setText(str(self.state.equip_qty) if self.state.equip_qty else "-")
        self.lbl_actual_line_disp.setText(self.state.actual_line or "-")
        self.lbl_exh_size_disp.setText(str(self.state.exh_size) if self.state.exh_size else "-")

        applied = bool(self.state.pump_credit or self.state.rack_credit)
        if applied:
            parts = []
            if self.state.pump_credit:
                parts.append(f"Pump {fmt_krw(self.state.pump_credit)}")
            if self.state.rack_credit:
                parts.append(f"Rack {fmt_krw(self.state.rack_credit)}")
            self.lbl_credit_status.setText("✓ Credit 적용됨\n" + " / ".join(parts))
            self.lbl_credit_status.setStyleSheet(
                "background:#E8F5E9; border:1px solid #A5D6A7; border-radius:4px; padding:2px;")
        else:
            self.lbl_credit_status.setText("Credit 미적용")
            self.lbl_credit_status.setStyleSheet(
                "background:#F5F5F5; border:1px solid #DDD; border-radius:4px; padding:2px;")
        self._refresh_opt_summary()

    def _build_quote_info_panel(self) -> QFrame:
        """가운데 맨 위 한 줄: 고른 의뢰의 정보 + 지금 옵션 요약(누르면 옵션 탭)."""
        frame = QFrame(); frame.setFrameShape(QFrame.Box); frame.setLineWidth(2)
        h = QHBoxLayout(frame); h.setContentsMargins(10, 4, 10, 4); h.setSpacing(14)
        self.lbl_pr = info_label("PR: -"); self.lbl_item = info_label("항번: -")
        self.lbl_line = info_label("라인공정: -"); self.lbl_investor = info_label("투자자: -")
        self.lbl_more = info_label("")
        for lbl in (self.lbl_pr, self.lbl_item, self.lbl_line, self.lbl_investor, self.lbl_more):
            lbl.setSizePolicy(QSizePolicy.Preferred, QSizePolicy.Fixed)
            h.addWidget(lbl)
        h.addStretch(1)
        self.btn_opt_summary = QPushButton("옵션")
        self.btn_opt_summary.setFixedHeight(28)
        self.btn_opt_summary.setToolTip("지금 적용된 옵션입니다. 누르면 옵션 탭이 열립니다.")
        self.btn_opt_summary.setStyleSheet(
            "QPushButton { background: #F0F1F4; border: 1px solid #C3C9D4; border-radius: 14px; padding: 0 12px; }"
            "QPushButton:hover { border-color: #1F3864; }")
        self.btn_opt_summary.clicked.connect(lambda: self.side_tabs.setCurrentIndex(1))
        h.addWidget(self.btn_opt_summary)
        return frame

    def _refresh_opt_summary(self) -> None:
        if not hasattr(self, "btn_opt_summary"):
            return
        st = self.state
        credit = "Credit 적용" if (st.pump_credit or st.rack_credit) else "Credit 미적용"
        self.btn_opt_summary.setText(
            f"{st.quote_type or '국내'}  ·  {credit}  ·  보증 {st.warranty_years}년  ·  "
            f"{st.investor_name or '-'}   ▸ 옵션")

    def _style_req_table(self) -> None:
        self.req_table.setAlternatingRowColors(False)
        # item background 는 지정하지 않아야 _mark_req_done 의 setBackground 가 반영된다
        self.req_table.setStyleSheet(
            "QTableWidget{background:transparent;border:none;}"
            "QTableWidget::item:selected{color:black;background:rgba(180,200,230,140);}")

    # ══════════════════════════════════════════
    # STEP1
    # ══════════════════════════════════════════

    def _load_template(self) -> None:
        start = app_settings.get_path(app_settings.Key.TEMPLATE_PATH) or ""
        path, _ = QFileDialog.getOpenFileName(
            self, "통합양식 선택", start, "Excel Files (*.xlsx)")
        if not path:
            return
        self._start_template_load(path)

    def _start_template_load(self, path: str, silent: bool = False) -> None:
        """통합양식을 백그라운드로 읽는다. silent=True 면 자동 복원 경로."""
        if self._tpl_thread and self._tpl_thread.isRunning():
            return
        self._notices = []          # 양식 알림은 양식을 새로 읽을 때마다 다시 모은다
        # 시트명 사전 검증 (빠름 — 메타만 읽음)
        try:
            from openpyxl import load_workbook
            wb_meta = load_workbook(path, read_only=True, data_only=False)
            cfg = quote_config.from_workbook(wb_meta)
            sheetnames = list(wb_meta.sheetnames)
            wb_meta.close()
            missing = [n for n in cfg.required_sheets if n not in sheetnames]
            # 설정 시트의 잘못된 값, 없는 시트 등 — 생성 도중이 아니라 지금 알려준다
            for w in cfg.check(sheetnames):
                self._log(f"[설정 확인] {w}")
            if cfg.overridden:
                self._log(f"양식의 '{quote_config.SETTINGS_SHEET}' 시트에서 설정 {len(cfg.overridden)}개를 읽었습니다")
        except Exception as e:
            if not silent:
                QMessageBox.critical(self, "오류", f"파일을 열 수 없습니다.\n{e}")
            else:
                self._log(f"통합양식 자동 로드 실패: {e}")
            return
        if missing:
            if not silent:
                QMessageBox.warning(self, "통합양식 오류",
                                    "필수 시트 없음:\n" + "\n".join(missing))
            else:
                self._log("통합양식 자동 로드 실패: 필수 시트 없음")
            return

        # 통합양식 사전 점검 — 생성물에서 조용히 어긋나는 것을 미리 알린다
        try:
            import excel_io
            losses = excel_io.inspect_template_losses(path)
            price_gaps = excel_io.check_pump_price_consistency(path)
        except Exception:
            losses, price_gaps = [], []

        for w in losses:
            self._log(f"[통합양식 경고] {w}")
        for w in price_gaps:
            self._log(f"[단가 불일치] {w}")

        if (losses or price_gaps) and not silent:
            msg = ""
            if losses:
                msg += ("생성된 견적서에서 사라지는 요소:\n"
                        + "\n".join(f"  • {w}" for w in losses)
                        + "\n\n엑셀 라이브러리(openpyxl)가 보존하지 못하는 형식입니다.\n"
                          "벡터 이미지는 엑셀에서 PNG 그림으로 교체하면 해결됩니다.\n")
            if price_gaps:
                if msg:
                    msg += "\n" + "─" * 40 + "\n\n"
                msg += ("PUMP 단가가 사양서와 앱에서 다릅니다:\n"
                        + "\n".join(f"  • {w}" for w in price_gaps)
                        + "\n\n해당 Q-Code 로 견적을 만들면 인쇄되는 사양서 금액과\n"
                          "갑지 금액이 서로 달라집니다. 통합양식의 품목/Pump 단가표를\n"
                          "확인해 주세요.\n")
            msg += "\n생성은 그대로 진행됩니다."
            QMessageBox.warning(self, "통합양식 주의", msg)

        self.btn_step1.setEnabled(False)
        self.btn_step1.setText("STEP 1\n로딩 중…")
        self._tpl_thread = _TemplateLoaderThread(path, self)
        self._tpl_thread.done.connect(
            lambda result: self._on_template_loaded(path, result, silent))
        self._tpl_thread.start()

    def _on_template_loaded(self, path: str, result, silent: bool = False) -> None:
        self.btn_step1.setEnabled(True)
        self.btn_step1.setText("STEP 1\n통합양식 LOAD")
        if isinstance(result, Exception):
            logger.error("통합양식 파싱 오류", exc_info=result)
            if silent:
                self._log(f"통합양식 자동 로드 실패: {result}")
            else:
                QMessageBox.critical(self, "파싱 오류", str(result))
            return
        rows, by_class, price_by_spec, order, cmap = result
        self.state.template_path = path
        self.state.items_rows = rows
        self.state.items_by_class = by_class
        self.state.price_by_spec = price_by_spec
        self.state.code_map = cmap
        self._spec_order = order
        self._class_of = {r["B"]: r["A"] for r in rows if r["B"] and r["A"]}
        _prices: Dict[str, set] = {}
        for r in rows:
            if r["B"]:
                _prices.setdefault(r["B"], set()).add(r["C"])
        self._dup_price_specs = {sp for sp, ps in _prices.items() if len(ps) > 1}
        self._schedule_review()
        self._apply_filter()
        self.btn_step2.setEnabled(True)
        self._step3_frame.setEnabled(True)
        app_settings.set_str(app_settings.Key.TEMPLATE_PATH, path)
        self._log(f"STEP1 완료: {path}  (품목 {len(rows)}건, 코드매핑 키 {len(cmap)}개)")
        if not silent:
            QMessageBox.information(self, "완료", "통합양식 LOAD 완료")

    # ══════════════════════════════════════════
    # STEP2
    # ══════════════════════════════════════════

    def _load_request(self) -> None:
        if not self.state.template_path:
            QMessageBox.warning(self, "순서 오류", "STEP1을 먼저 완료하세요."); return
        if self._req_thread and self._req_thread.isRunning():
            return
        start = os.path.dirname(
            app_settings.get_path(app_settings.Key.REQUEST_PATH)) or ""
        path, _ = QFileDialog.getOpenFileName(
            self, "의뢰파일 선택", start, "Excel Files (*.xlsx)")
        if not path: return
        self._req_thread = _RequestLoaderThread(path, self, self.state.template_path)
        self._req_thread.done.connect(lambda result: self._on_request_loaded(path, result))
        self._req_thread.start()

    def _on_request_loaded(self, path: str, result) -> None:
        if isinstance(result, Exception):
            logger.error("의뢰파일 파싱 오류", exc_info=result)
            QMessageBox.critical(self, "오류", f"파일을 읽을 수 없습니다.\n{result}"); return
        sheet_name, rows = result
        if not rows:
            QMessageBox.warning(self, "의뢰파일", "읽을 데이터가 없습니다."); return
        self.state.request_path = path; self.state.request_sheet_name = sheet_name
        app_settings.set_str(app_settings.Key.REQUEST_PATH, path)
        self.state.request_rows = rows
        self._fill_req_table(rows); self._step3_frame.setEnabled(True)
        self._log(f"STEP2 완료: {path}  ({len(rows)}건, 시트={sheet_name})")
        QMessageBox.information(self, "완료", "의뢰파일 LOAD 완료")

    def _fill_req_table(self, rows: List[Dict[str, Any]]) -> None:
        self._done_req_indices = set()   # 새 의뢰파일 로드 시 완료 이력 초기화
        self.req_table.setRowCount(0)
        self.chk_req_all.blockSignals(True); self.chk_req_all.setCheckState(Qt.Unchecked); self.chk_req_all.blockSignals(False)
        for i, rd in enumerate(rows):
            self.req_table.insertRow(i)
            host = centered_checkbox(lambda _s, _r=i: self._on_req_check())
            self.req_table.setCellWidget(i, 0, host)
            invest_info = parse_invest_info(rd.get("G"))
            h_val = rd.get("H")
            qty_str = fmt_qty(to_float(h_val)) if h_val is not None else ""
            # 목록에는 호기·수량만 보이므로 나머지는 행 툴팁으로 전부 보여준다
            tooltip = "\n".join(f"{name}: {val or '-'}" for name, val in [
                ("설비호기", s(rd.get("Z"))), ("투자정보", invest_info), ("수량", qty_str),
                ("Rack", s(rd.get("R"))), ("Maker", s(rd.get("X"))), ("설비", s(rd.get("Y"))),
                ("5D", s(rd.get("V"))), ("PR", s(rd.get("D"))), ("항번", s(rd.get("E"))),
                ("라인", s(rd.get("K"))), ("의뢰자", s(rd.get("J"))),
            ])
            for col, val in [
                (1, s(rd.get("Z"))),
                (2, invest_info),
                (3, qty_str),
                (4, s(rd.get("R"))),   # Rack (R열=col18)
                (5, s(rd.get("X"))),   # Maker (X열=col24)
                (6, s(rd.get("Y"))),   # 설비 (Y열=col25)
                (7, s(rd.get("V"))),   # 5D (V열=col22)
            ]:
                it = QTableWidgetItem(val)
                it.setFlags(Qt.ItemIsEnabled | Qt.ItemIsSelectable)
                it.setToolTip(tooltip)
                self.req_table.setItem(i, col, it)
        self.req_table.setColumnWidth(0, 36)
        self.req_table.setColumnWidth(3, 48)
        self._style_req_table(); self._update_quote_info()
        self._schedule_review()

    def _is_req_checked(self, row: int) -> bool:
        cb = get_checkbox_from_cell(self.req_table.cellWidget(row, 0))
        return bool(cb and cb.isChecked())

    def _checked_req_rows(self) -> List[int]:
        return [r for r in range(self.req_table.rowCount()) if self._is_req_checked(r)]

    def _on_req_check(self) -> None:
        n = sum(1 for r in range(self.req_table.rowCount()) if self._is_req_checked(r))
        total = self.req_table.rowCount()
        self.chk_req_all.blockSignals(True)
        if   n == 0:     self.chk_req_all.setCheckState(Qt.Unchecked)
        elif n == total: self.chk_req_all.setCheckState(Qt.Checked)
        else:            self.chk_req_all.setCheckState(Qt.PartiallyChecked)
        self.chk_req_all.blockSignals(False); self._update_quote_info()

    def _on_req_select_all(self, checked: bool) -> None:
        for r in range(self.req_table.rowCount()):
            cb = get_checkbox_from_cell(self.req_table.cellWidget(r, 0))
            if cb: cb.blockSignals(True); cb.setChecked(checked); cb.blockSignals(False)
        self._on_req_check()

    def _update_quote_info(self) -> None:
        checked = self._checked_req_rows()
        if not checked or not self.state.request_rows:
            self.lbl_pr.setText("PR: -"); self.lbl_item.setText("항번: -")
            self.lbl_line.setText("라인공정: -"); self.lbl_investor.setText("투자자: -")
            self.lbl_more.setText(""); return
        first = self.state.request_rows[checked[0]]
        self.lbl_pr.setText(f"PR: {s(first.get('D')) or '-'}")
        self.lbl_item.setText(f"항번: {s(first.get('E')) or '-'}")
        self.lbl_line.setText(f"라인공정: {s(first.get('K')) or '-'}")
        self.lbl_investor.setText(f"투자자: {s(first.get('J')) or '-'}")
        self.lbl_more.setText(f"외 {len(checked)-1}건" if len(checked) > 1 else "")

    def _on_req_double_click(self, row: int, _col: int) -> None:
        if row >= len(self.state.request_rows): return
        rd = self.state.request_rows[row]

        def _match(text: str, choices: List[str]) -> Optional[str]:
            t = normalize_token(text)
            if not t: return None
            for c in sorted((x for x in choices if x), key=len, reverse=True):
                if normalize_token(c) in t: return c
            return None

        p = _match(s(rd.get("K")), PROCESS_CHOICES)
        if p:
            idx = self.cb_process.findText(p)
            if idx >= 0: self.cb_process.setCurrentIndex(idx)
        v = _match(s(rd.get("X")), VENDOR_CHOICES)
        if v:
            idx = self.cb_vendor.findText(v)
            if idx >= 0: self.cb_vendor.setCurrentIndex(idx)
        vcode = s(rd.get("V"))
        if vcode: self.ed_code.setText(vcode)

        f_class = s(rd.get("F")); qty = to_float(rd.get("H"))
        if f_class:
            matches = self.state.items_by_class.get(f_class, [])
            added, skipped = 0, 0
            for row_data in matches:
                spec = s(row_data.get("B")); price = float(row_data.get("C", 0))
                if not spec:
                    continue
                # _add_row 는 "같은 spec 이면 수량 누적"이 기본 동작이다(STEP5에서
                # 사용자가 수동으로 수량을 늘릴 때 쓰라고 만든 기능). 그런데 이
                # 더블클릭 자동 추가는 같은 의뢰행을 두 번 클릭하거나, 같은
                # Q-Code(=같은 PUMP 모델)를 쓰는 다른 의뢰행을 확인차 또 클릭해도
                # 그대로 호출되므로, 누적을 그대로 적용하면 PUMP 수량이 조용히
                # 배가되어 실제로 견적금액이 2배로 나간 사고가 있었다. 자동
                # 추가는 이미 있는 spec 이면 새로 더하지 않고 건너뛴다.
                if self._has_spec(spec):
                    skipped += 1
                    continue
                self._add_row("PUMP", spec, qty, price)
                added += 1
            if added:
                self._log(f"PUMP 자동: 분류={f_class}, {added}개, 수량={fmt_qty(qty)}")
                self._sync_step4_highlight()
            if skipped:
                self._log(f"PUMP 자동 건너뜀: 분류={f_class}, 이미 담긴 품목 {skipped}개 "
                           f"— 같은 항목을 두 번 담지 않도록 수량 누적을 막았습니다")
        self._log(f"더블클릭 → 공정={self.cb_process.currentText()} / 설비사={self.cb_vendor.currentText()} / 5D={self.ed_code.text()}")

    # ══════════════════════════════════════════
    # STEP3
    # ══════════════════════════════════════════

    def _on_step3_changed(self) -> None:
        self.state.process = s(self.cb_process.currentText()) or None
        self.state.vendor  = s(self.cb_vendor.currentText())  or None
        self.state.code_5d = s(self.ed_code.text())           or None
        ready = bool(self.state.process and self.state.vendor and self.state.code_5d)
        self.btn_apply_map.setEnabled(ready and bool(self.state.code_map))
        self._schedule_review()

    def _apply_code_map(self) -> None:
        key = (self.state.process, self.state.vendor, self.state.code_5d)
        entries = self.state.code_map.get(key)
        if not entries:
            near = sum(1 for k in self.state.code_map if k[:2] == key[:2])
            QMessageBox.warning(self, "코드매핑 없음", f"키: {key}\n동일 공정/설비사 키 수: {near}"); return
        missing = []
        for acc, qty in entries:
            if acc not in self.state.price_by_spec:
                missing.append(acc)
            price = float(self.state.price_by_spec.get(acc, 0))
            self._add_row("RACK", acc, qty, price)
        self._log(f"코드매핑 적용 {key} → {len(entries)}개"); self._sync_step4_highlight()
        if missing:
            # 품목 시트에 없는 이름은 단가를 찾지 못해 0원으로 담긴다.
            # 모르고 지나가면 그 품목이 빠진 금액으로 견적이 나간다.
            self._log(f"코드매핑 품목 중 품목 시트에 없는 것 {len(missing)}개 → 단가 0원: "
                      + ", ".join(missing))
            QMessageBox.warning(
                self, "단가 확인 필요",
                "아래 품목은 품목 시트에 없어 단가 0원으로 담겼습니다.\n"
                "단가를 직접 입력하거나 품목 시트/코드매핑의 이름을 맞춰주세요.\n\n"
                + "\n".join(missing))

    def _gen_5d_object(self) -> None:
        """현재 STEP5의 RACK 품목/수량을 공정·설비사·5D 키로 코드매핑 시트에 저장한다."""
        if not self.state.template_path:
            QMessageBox.warning(self, "오류", "STEP1 통합양식을 먼저 LOAD 하세요."); return
        process, vendor = self.state.process, self.state.vendor
        if not process or not vendor:
            QMessageBox.warning(self, "오류", "STEP3에서 공정/설비사를 먼저 선택하세요."); return

        # 코드매핑은 STEP5(RACK)에만 적용되는 개념이라 RACK 품목만 대상으로 한다.
        # (코드매핑 적용 버튼도 항상 RACK으로만 STEP5에 되돌려 넣는다.)
        rack_items = [
            (it["spec"], it["qty"]) for it in self._snapshot_items()
            if it["role"] == "ITEM" and it["cat"] == "RACK" and it["spec"] and it["qty"]
        ]
        if not rack_items:
            QMessageBox.warning(
                self, "오류",
                "STEP5에 RACK 품목이 없습니다. (코드매핑은 RACK 품목/수량만 저장됩니다)")
            return

        code5d, ok = QInputDialog.getText(
            self, "5D Object 생성", "5D를 입력하세요:", text=self.state.code_5d or "")
        if not ok:
            return
        code5d = code5d.strip()
        if not code5d:
            QMessageBox.warning(self, "오류", "5D 값을 입력하세요."); return

        try:
            import excel_io
            added, skipped = excel_io.append_code_map_rows(
                self.state.template_path, process, vendor, code5d, rack_items)
        except Exception as e:
            logger.error("5D Object 생성 실패", exc_info=e)
            tb = traceback.format_exc()
            user_msg, hint = _friendly_error_msg(e)
            _ScrollableErrorDialog(self, tb, user_msg=user_msg, hint=hint).exec()
            return

        if not added:
            QMessageBox.warning(
                self, "중복된 값",
                f"이미 완전히 동일한 코드매핑이 있어 저장하지 않았습니다.\n"
                f"공정/설비사/5D: {process}/{vendor}/{code5d}  (품목 {len(skipped)}건 모두 중복)")
            return

        key = (process, vendor, code5d)
        self.state.code_map.setdefault(key, []).extend(added)
        # STEP3의 5D 입력칸을 방금 만든 값으로 맞춰, "코드매핑 적용" 버튼으로
        # 바로 이어서 쓸 수 있게 한다 (_on_step3_changed 가 활성화 여부도 갱신).
        self.ed_code.setText(code5d)

        msg = f"코드매핑에 {len(added)}건을 추가했습니다."
        if skipped:
            msg += f" (중복 {len(skipped)}건은 건너뜀)"
        QMessageBox.information(self, "완료", msg)
        self._log(f"5D Object 생성 {key} → {len(added)}건 추가"
                   + (f", 중복 {len(skipped)}건 제외" if skipped else ""))

    # ══════════════════════════════════════════
    # STEP4
    # ══════════════════════════════════════════

    def _apply_filter(self) -> None:
        self._filter_timer.start(200)

    def _do_filter(self) -> None:
        key = s(self.ed_filter.text()).lower()
        self._filtered_items = [
            (s(r.get("B")), float(r.get("C", 0)))
            for r in self.state.items_rows
            if s(r.get("B")) and key in s(r.get("B")).lower()
        ]
        self._render_items_table(scroll_top=False)

    def _render_items_table(self, scroll_top: bool = True) -> None:
        """Step5Manager 추상 메서드 구현: STEP4 테이블 재렌더링."""
        highlighted = self._step4_highlight

        def order(spec: str) -> int:
            return self._spec_order.get(spec, 10**9)

        sorted_rows = sorted(
            self._filtered_items,
            key=lambda x: (0 if x[0] in highlighted else 1, order(x[0])))

        self.items_table.blockSignals(True)
        self.items_table.setRowCount(len(sorted_rows))
        for i, (spec, price) in enumerate(sorted_rows):
            for col, val in [(0, spec), (1, fmt_krw(price))]:
                it = QTableWidgetItem(val); it.setFlags(Qt.ItemIsEnabled | Qt.ItemIsSelectable)
                if spec in highlighted: it.setBackground(QBrush(QColor(220, 220, 220)))
                self.items_table.setItem(i, col, it)
        self.items_table.blockSignals(False)
        self.items_table.setColumnWidth(1, 80)
        if scroll_top: self.items_table.scrollToTop()

    def _on_item_double_click(self, row: int, _col: int) -> None:
        spec_it  = self.items_table.item(row, 0); price_it = self.items_table.item(row, 1)
        if not spec_it: return
        self._ask_qty_add(spec_it.text().strip(), to_float(price_it.text() if price_it else "0"))

    def _on_drop_to_step5(self, spec: str, price: float) -> None:
        self._ask_qty_add(spec, price)

    def _ask_qty_add(self, spec: str, unit_price: float) -> None:
        qty, ok = QInputDialog.getDouble(self, "수량 입력", f"수량 ({spec})", 1.0, 0, 1_000_000, 3)
        if ok: self._add_row("RACK", spec, qty, unit_price); self._sync_step4_highlight()

    # ══════════════════════════════════════════
    # STEP5 이벤트
    # ══════════════════════════════════════════

    def _on_step5_double_click(self, row: int, col: int) -> None:
        if self._is_total(row) or self._is_credit(row): return
        if col == 1:
            it = self.step5_table.item(row, 1)
            if it:
                it.setText("RACK" if it.text().strip() == "PUMP" else "PUMP")
                self._sort_step5_items()
                self._recalc_totals(); self._sync_step4_highlight(scroll_top=False)
        elif col == 3:
            cur = self._get_float(row, 3)
            val, ok = QInputDialog.getDouble(self, "수량 수정", "수량:", cur, 0, 1_000_000, 3)
            if ok: self._set_qty(row, val); self._recalc_row(row); self._recalc_totals(); self._sync_step4_highlight()

    def _delete_checked_rows(self) -> None:
        to_del = [
            r for r in range(self.step5_table.rowCount() - 1)
            if not self._is_total(r) and not self._is_credit(r)
            and isinstance(self.step5_table.cellWidget(r, 0), QCheckBox)
            and self.step5_table.cellWidget(r, 0).isChecked()
        ]
        if not to_del:
            sel = self.step5_table.selectionModel().selectedRows()
            if sel:
                r = sel[0].row()
                if not self._is_total(r) and not self._is_credit(r): to_del = [r]
        if not to_del: return
        reply = QMessageBox.question(self, "삭제 확인", f"{len(to_del)}개 품목을 삭제할까요?",
                                     QMessageBox.Yes | QMessageBox.No)
        if reply != QMessageBox.Yes: return
        for r in sorted(to_del, reverse=True): self.step5_table.removeRow(r)
        self._ensure_total(); self._refresh_credits(); self._sort_step5_items(); self._recalc_totals()
        self.chk_step5_all.blockSignals(True); self.chk_step5_all.setCheckState(Qt.Unchecked); self.chk_step5_all.blockSignals(False)
        self._sync_step4_highlight()

    # ══════════════════════════════════════════
    # 옵션
    # ══════════════════════════════════════════

    def _on_qt_checked(self, selected: str, checked: bool) -> None:
        chk_map = {"국내": self.chk_qt_kr, "중국": self.chk_qt_cn, "미국": self.chk_qt_us}
        if not checked:
            # 이미 선택된 항목 재클릭 → 체크 복원 후 옵션창
            chk = chk_map[selected]; chk.blockSignals(True); chk.setChecked(True); chk.blockSignals(False)
            if selected in ("중국", "미국"):
                self._open_options(focus={"중국": "cn", "미국": "us"}[selected])
            return
        for name, chk in chk_map.items():
            chk.blockSignals(True); chk.setChecked(name == selected); chk.blockSignals(False)
        self.state.quote_type = selected; self._log(f"견적서 타입 = {selected}")
        self._schedule_review()
        if selected in ("중국", "미국"): self._open_options(focus={"중국": "cn", "미국": "us"}[selected])

    def _open_options(self, focus: str = "credit") -> None:
        if focus == "cn":
            dlg = ChinaOptionsDialog(self, self.state)
        elif focus == "us":
            dlg = USOptionsDialog(self, self.state)
        else:
            return
        if dlg.exec() == QDialog.Accepted:
            dlg.apply(); self._refresh_credits(); self._recalc_totals()
            self._sync_step4_highlight(scroll_top=False); self._log("[옵션] 입력값 반영")

    # ── STEP5 컬럼 너비 관리 ─────────────────────────────────────────
    _S5_ADJ = [1, 3, 4, 5]   # 품목 변경 시 비례 조정되는 컬럼들
    _S5_ADJ_MIN = 28          # 비례 컬럼 최소 너비(px)
    _S5_COL2_MIN = 60         # 품목 최소 너비(px)

    def _step5_fit_col2(self) -> None:
        """초기화 시 품목(col 2) 너비를 364px로 설정하고 나머지 칼럼들을 비례 배분."""
        total     = self.step5_table.viewport().width()
        # 품목 열에 표 너비의 절반을 준다(이름이 길다). 나머지를 분류·수량·단가·금액이 나눈다.
        target    = max(200, int(total * 0.5))
        col0_w    = self.step5_table.columnWidth(0)
        min_adj   = len(self._S5_ADJ) * self._S5_ADJ_MIN
        col2      = max(self._S5_COL2_MIN, min(target, total - col0_w - min_adj))
        remaining = total - col0_w - col2
        old_ws    = [self.step5_table.columnWidth(c) for c in self._S5_ADJ]
        s         = sum(old_ws) or 1
        new_ws    = [max(self._S5_ADJ_MIN, round(w / s * remaining)) for w in old_ws]
        new_ws[-1] = max(self._S5_ADJ_MIN, remaining - sum(new_ws[:-1]))
        h5 = self.step5_table.horizontalHeader()
        h5.blockSignals(True)
        for c, w in zip(self._S5_ADJ, new_ws):
            self.step5_table.setColumnWidth(c, w)
        self.step5_table.setColumnWidth(2, col2)
        h5.blockSignals(False)

    def _on_step5_col2_resized(self, logical: int, _old: int, new_size: int) -> None:
        """품목(col 2) 드래그 시 나머지 컬럼들을 현재 비율 그대로 비례 조정."""
        if logical != 2:
            return
        total    = self.step5_table.viewport().width()
        col0_w   = self.step5_table.columnWidth(0)
        min_adj  = self._S5_ADJ_MIN
        # 품목 너비 범위 제한
        max_col2 = total - col0_w - len(self._S5_ADJ) * min_adj
        clamped  = max(self._S5_COL2_MIN, min(new_size, max_col2))
        remaining = total - col0_w - clamped

        old_ws = [self.step5_table.columnWidth(c) for c in self._S5_ADJ]
        s = sum(old_ws) or 1
        new_ws = [max(min_adj, round(w / s * remaining)) for w in old_ws]
        # 마지막 컬럼이 반올림 오차 흡수
        new_ws[-1] = max(min_adj, remaining - sum(new_ws[:-1]))

        h5 = self.step5_table.horizontalHeader()
        h5.blockSignals(True)
        for c, w in zip(self._S5_ADJ, new_ws):
            self.step5_table.setColumnWidth(c, w)
        if clamped != new_size:
            self.step5_table.setColumnWidth(2, clamped)
        h5.blockSignals(False)

    def eventFilter(self, obj, event) -> bool:
        if (hasattr(self, 'step5_table')
                and obj is self.step5_table.viewport()
                and event.type() == QEvent.Resize):
            self._step5_fit_col2()
        return super().eventFilter(obj, event)

    def _change_warranty(self) -> None:
        dlg = QDialog(self)
        dlg.setWindowTitle("보증기간 변경")
        dlg.setFixedWidth(300)
        v = QVBoxLayout(dlg)
        v.setSpacing(10)

        form = QFormLayout()
        spin = QSpinBox()
        spin.setRange(1, 10)
        spin.setValue(self.state.warranty_years)
        spin.setSuffix(" Year after delivery")
        form.addRow("보증기간:", spin)
        v.addLayout(form)

        preview = QLabel(f"{self.state.warranty_years} Year after delivery")
        preview.setStyleSheet("color: #555; font-size: 11px;")
        spin.valueChanged.connect(lambda n: preview.setText(f"{n} Year after delivery"))
        v.addWidget(preview)

        btns = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        btns.accepted.connect(dlg.accept)
        btns.rejected.connect(dlg.reject)
        v.addWidget(btns)

        if dlg.exec() == QDialog.Accepted:
            self.state.warranty_years = spin.value()
            self._log(f"보증기간: {self.state.warranty_years} Year after delivery")
            self._refresh_options_display()

    def _change_investor(self) -> None:
        name, ok = QInputDialog.getText(
            self, "투자자 변경", "투자자 이름:",
            text=self.state.investor_name or "채승철",
        )
        if ok:
            self.state.investor_name = name.strip() or "채승철"
            self._log(f"투자자: {self.state.investor_name}")
            self._refresh_options_display()

    def _open_credit_dialog(self) -> None:
        try:
            self._open_credit_dialog_impl()
        except Exception:
            QMessageBox.critical(self, "Credit 오류", traceback.format_exc())

    def _open_credit_dialog_impl(self) -> None:
        dlg = QDialog(self); dlg.setWindowTitle("Credit 입력"); dlg.setFixedWidth(500)
        v = QVBoxLayout(dlg); v.setSpacing(10)

        h_val = to_float(self._pick_rd().get("H")) if self.state.request_rows else 0.0

        # ── 카테고리별 품목 합계 ──────────────────────────────────
        def _cat_sum(is_pump: bool) -> float:
            total = 0.0
            for r in range(self.step5_table.rowCount()):
                if self._is_item(r):
                    cat_it = self.step5_table.item(r, 1)
                    if cat_it and (s(cat_it.text()) == "PUMP") == is_pump:
                        total += self._get_float(r, 5)
            return total

        pump_sum = _cat_sum(True)
        rack_sum = _cat_sum(False)

        # ── 섹션 빌더 ────────────────────────────────────────────
        # 반환: (ed_credit, ed_target, ed_ch)
        def _make_section(label: str, item_sum: float, init_credit: float):
            frame = QFrame(); frame.setFrameShape(QFrame.StyledPanel)
            fv    = QVBoxLayout(frame); fv.setContentsMargins(8, 6, 8, 6); fv.setSpacing(6)
            hdr = QLabel(f"<b>{label}</b>  (품목 합계: {fmt_krw(item_sum)} 원)")
            fv.addWidget(hdr)

            form = QFormLayout(); form.setVerticalSpacing(6)
            fv.addLayout(form)

            ed_credit = QLineEdit(); ed_credit.setPlaceholderText("0")
            ed_target = QLineEdit(); ed_target.setPlaceholderText(fmt_krw(item_sum))
            ed_ch     = QLineEdit(); ed_ch.setPlaceholderText(
                fmt_krw(item_sum / h_val) if h_val > 0 else "CH 값 없음"
            )
            ed_ch.setEnabled(h_val > 0)

            form.addRow("Credit 금액(원):", ed_credit)
            form.addRow("목표가(원):",      ed_target)
            form.addRow("CH당 단가(원/CH):", ed_ch)

            if init_credit:
                ed_credit.setText(fmt_krw(init_credit))

            updating = [False]

            def _from_credit():
                if updating[0]: return
                updating[0] = True
                try:
                    c = to_float(ed_credit.text())
                    t = item_sum - c
                    ed_target.setText(fmt_krw(max(0.0, t)))
                    if h_val > 0:
                        ed_ch.setText(fmt_krw(max(0.0, t) / h_val))
                finally:
                    updating[0] = False

            def _from_target():
                if updating[0]: return
                updating[0] = True
                try:
                    t = to_float(ed_target.text())
                    c = max(0.0, item_sum - t)
                    ed_credit.setText(fmt_krw(c))
                    if h_val > 0:
                        ed_ch.setText(fmt_krw(t / h_val))
                finally:
                    updating[0] = False

            def _from_ch():
                if updating[0]: return
                if h_val <= 0: return
                updating[0] = True
                try:
                    ch = to_float(ed_ch.text())
                    t  = ch * h_val
                    c  = max(0.0, item_sum - t)
                    ed_credit.setText(fmt_krw(c))
                    ed_target.setText(fmt_krw(max(0.0, t)))
                finally:
                    updating[0] = False

            ed_credit.textEdited.connect(_from_credit)
            ed_target.textEdited.connect(_from_target)
            ed_ch.textEdited.connect(_from_ch)

            return frame, ed_credit, ed_target, ed_ch

        grp_pump, ed_pump_c, ed_pump_t, ed_pump_ch = _make_section(
            "PUMP", pump_sum, self.state.pump_credit)
        grp_rack, ed_rack_c, ed_rack_t, ed_rack_ch = _make_section(
            "RACK", rack_sum, self.state.rack_credit)

        v.addWidget(grp_pump)
        v.addWidget(grp_rack)

        # ── 결과 요약 ─────────────────────────────────────────────
        sep = QFrame(); sep.setFrameShape(QFrame.HLine); sep.setFrameShadow(QFrame.Sunken)
        v.addWidget(sep)
        lbl_summary = QLabel(); lbl_summary.setStyleSheet("font-weight:bold;")
        lbl_ch_total = QLabel()
        v.addWidget(lbl_summary); v.addWidget(lbl_ch_total)

        def _refresh_summary() -> None:
            pc    = to_float(ed_pump_c.text())
            rc    = to_float(ed_rack_c.text())
            after = pump_sum + rack_sum - pc - rc
            lbl_summary.setText(
                f"Credit 반영 후 총액: {fmt_krw(after)} 원"
                f"  (감액 Pump {fmt_krw(pc)} + Rack {fmt_krw(rc)} = {fmt_krw(pc + rc)} 원)"
            )
            if h_val > 0:
                lbl_ch_total.setText(
                    f"전체 CH당 단가: {fmt_krw(after / h_val)} 원/CH  (H: {fmt_qty(h_val)} CH)"
                )
            else:
                lbl_ch_total.setText("전체 CH당 단가: H열 값 없음")

        for ed in (ed_pump_c, ed_pump_t, ed_pump_ch, ed_rack_c, ed_rack_t, ed_rack_ch):
            ed.textEdited.connect(_refresh_summary)

        bb = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel); v.addWidget(bb)

        def _ok() -> None:
            self.state.pump_credit = to_float(ed_pump_c.text())
            self.state.rack_credit = to_float(ed_rack_c.text())
            self._refresh_credits(); self._recalc_totals(); dlg.accept()

        bb.accepted.connect(_ok); bb.rejected.connect(dlg.reject)
        _refresh_summary(); dlg.exec()

    # ══════════════════════════════════════════
    # STEP6 – 생성
    # ══════════════════════════════════════════

    def _set_gen_buttons(self, enabled: bool) -> None:
        self.btn_gen_quote.setEnabled(enabled)
        self.btn_gen_cover.setEnabled(enabled)
        self.btn_submit_xls.setEnabled(enabled)

    def _generate_quote(self) -> None:
        if not self.state.template_path:
            QMessageBox.warning(self, "오류", "STEP1 통합양식을 먼저 LOAD 하세요."); return
        if self._gen_qt_thread and self._gen_qt_thread.isRunning():
            QMessageBox.warning(self, "진행 중", "이미 생성 작업이 진행 중입니다."); return
        self._refresh_credits(); self._recalc_totals()
        items = self._snapshot_items(); qtype = self.state.quote_type or "국내"
        if qtype == "중국" and not self.state.cn_info.get("tool"):
            self._open_options("cn")
            if not self.state.cn_info.get("tool"): return
        if qtype == "미국" and not self.state.us_info.get("tool"):
            self._open_options("us")
            if not self.state.us_info.get("tool"): return
        # 생성 전 점검 — 판단이 필요한 항목이 남아 있으면 한 번 더 묻는다
        review = self._make_review()
        if review.red:
            lines = "\n".join("  • " + c.text for c in review.red[:10])
            more = f"\n  … 외 {len(review.red) - 10}건" if len(review.red) > 10 else ""
            if QMessageBox.question(
                    self, "생성 전 확인",
                    f"판단이 필요한 항목이 {len(review.red)}건 있습니다.\n\n{lines}{more}\n\n그래도 생성할까요?"
            ) != QMessageBox.Yes:
                return
        checked = self._checked_req_rows()
        if qtype == "국내" and self.state.request_rows and len(checked) >= 2:
            valid = [i for i in checked if i < len(self.state.request_rows)]
            rds = [self.state.request_rows[i] for i in valid]
            self._pending_req_indices = valid
        else:
            idx = self._pick_rd_index()
            rds = [self.state.request_rows[idx]] if (idx >= 0 and self.state.request_rows) else [{}]
            self._pending_req_indices = [idx] if idx >= 0 else []
        self._log(f"견적서 생성 시작 ({qtype}, {len(rds)}건) …")
        self._set_gen_buttons(False)
        self._gen_qt_thread = _GenQuoteThread(self.state, qtype, items, rds, self)
        self._gen_qt_thread.progress.connect(self._log)
        self._gen_qt_thread.done.connect(self._on_gen_quote_done)
        self._gen_qt_thread.start()

    def _on_gen_quote_done(self, result) -> None:
        self._set_gen_buttons(True)
        if isinstance(result, Exception):
            logger.error("견적서 생성 실패", exc_info=result)
            tb = "".join(traceback.format_exception(type(result), result, result.__traceback__))
            user_msg, hint = _friendly_error_msg(result)
            _ScrollableErrorDialog(self, tb, user_msg=user_msg, hint=hint).exec()
            return
        failures = []
        for i, (rd, path, err) in enumerate(result):
            if path:
                self._add_done(path)
                self.state.last_output_dir = os.path.dirname(path)
                app_settings.set_str(app_settings.Key.OUTPUT_DIR,
                                     self.state.last_output_dir)
                if i < len(self._pending_req_indices):
                    idx = self._pending_req_indices[i]
                    self._done_req_indices.add(idx)
                    self._mark_req_done(idx)
            else:
                failures.append((s(rd.get("Z")) or f"{i + 1}번째 행", err))
        ok = len(result) - len(failures)
        self._log(f"견적서 생성 완료: {ok}/{len(result)}건")

        if failures:
            # 실패를 '완료'로 알리면 사용자가 누락을 알아챌 방법이 없다.
            detail = "\n".join(f"  • {name}: {err}" for name, err in failures[:15])
            if len(failures) > 15:
                detail += f"\n  … 외 {len(failures) - 15}건"
            QMessageBox.warning(
                self, "일부 실패",
                f"견적서 {ok}/{len(result)}건 생성 완료\n"
                f"{len(failures)}건 실패:\n\n{detail}",
            )
        else:
            QMessageBox.information(self, "완료", f"견적서 {ok}건 생성 완료")
        self._open_quote_list_window()

    def _generate_cover(self) -> None:
        if not self.state.template_path:
            QMessageBox.warning(self, "오류", "STEP1 통합양식을 먼저 LOAD 하세요."); return
        if self._gen_cv_thread and self._gen_cv_thread.isRunning():
            QMessageBox.warning(self, "진행 중", "이미 갑지 생성 작업이 진행 중입니다."); return
        start = self.state.last_output_dir or exe_dir()
        paths, _ = QFileDialog.getOpenFileNames(self, "갑지 생성 대상 엑셀파일 (다중)", start, "Excel Files (*.xlsx)")
        if not paths: return
        folder = os.path.commonpath(paths)
        if os.path.isfile(folder): folder = os.path.dirname(folder)
        self._log(f"갑지 생성 시작 ({len(paths)}건) …")
        self._set_gen_buttons(False)
        self._gen_cv_thread = _GenCoverThread(
            self.state.template_path, folder, paths, self.state.investor_name,
            self.state.warranty_years, self)
        self._gen_cv_thread.progress.connect(self._log)
        self._gen_cv_thread.done.connect(self._on_gen_cover_done)

        self._cover_progress = QProgressDialog("갑지 생성 중…", None, 0, len(paths), self)
        self._cover_progress.setWindowTitle("갑지 생성")
        self._cover_progress.setWindowModality(Qt.WindowModal)
        self._cover_progress.setMinimumDuration(500)
        self._cover_progress.setValue(0)
        self._gen_cv_thread.step.connect(
            lambda d, t: self._cover_progress.setValue(d) if self._cover_progress else None
        )

        self._gen_cv_thread.start()

    def _on_gen_cover_done(self, result) -> None:
        if self._cover_progress:
            self._cover_progress.close()
            self._cover_progress = None
        self._set_gen_buttons(True)
        if isinstance(result, Exception):
            logger.error("갑지 생성 실패", exc_info=result)
            tb = "".join(traceback.format_exception(type(result), result, result.__traceback__))
            user_msg, hint = _friendly_error_msg(result)
            _ScrollableErrorDialog(self, tb, user_msg=user_msg, hint=hint).exec()
            return
        out, failed = result
        self.state.last_output_dir = os.path.dirname(out)
        app_settings.set_str(app_settings.Key.OUTPUT_DIR, self.state.last_output_dir)
        self._add_done_cover(out)
        self._log(f"갑지 생성 완료: {out} (누락 {len(failed)}건)")

        if failed:
            # 읽기에 실패한 원본은 갑지에서 빠진다. 조용히 넘어가면
            # 갑지가 한 건 모자란 채 정상으로 보인다.
            detail = "\n".join(f"  • {name}: {why}" for name, why in failed[:15])
            if len(failed) > 15:
                detail += f"\n  … 외 {len(failed) - 15}건"
            QMessageBox.warning(
                self, "일부 누락",
                f"갑지 생성 완료\n{os.path.basename(out)}\n\n"
                f"다음 {len(failed)}건은 읽지 못해 갑지에서 제외되었습니다:\n\n{detail}",
            )
        else:
            QMessageBox.information(self, "완료", f"갑지 생성 완료\n{os.path.basename(out)}")

    def _export_submit_excel(self) -> None:
        """갑지DATA 시트만 추출해 제출용 엑셀 파일로 저장."""
        import re as _re
        path, _ = QFileDialog.getOpenFileName(
            self, "갑지를 선택하세요", self.state.last_output_dir or exe_dir(),
            "Excel Files (*.xlsx)")
        if not path:
            return
        stem = os.path.splitext(os.path.basename(path))[0]
        if not _re.search(r"_갑지(?:_\d+)?$", stem):
            QMessageBox.warning(self, "갑지가 아닙니다",
                                "선택한 파일의 이름에 '_갑지'가 없습니다.\n갑지 파일을 선택해 주세요.")
            return
        try:
            import excel_io
            out = excel_io.export_cover_data_sheet(path, self.state.template_path)
        except Exception as e:
            logger.error("제출용 엑셀 생성 실패", exc_info=True)
            QMessageBox.critical(self, "오류", f"제출용 엑셀 생성 실패:\n{e}")
            return
        self._add_done_submit(out)
        self._log(f"제출용 엑셀 생성 완료: {out}")
        QMessageBox.information(self, "완료", f"제출용 엑셀 생성 완료\n{os.path.basename(out)}")

    def _pick_rd_index(self) -> int:
        """클릭(선택)된 행 → 체크된 첫 번째 행 → 0 순으로 의뢰행 인덱스 반환."""
        cur = self.req_table.currentRow()
        if 0 <= cur < len(self.state.request_rows):
            return cur
        checked = self._checked_req_rows()
        if checked and checked[0] < len(self.state.request_rows):
            return checked[0]
        return 0 if self.state.request_rows else -1

    def _pick_rd(self) -> Dict[str, Any]:
        idx = self._pick_rd_index()
        if idx >= 0 and idx < len(self.state.request_rows):
            return self.state.request_rows[idx]
        return {}

    def _mark_req_done(self, row: int) -> None:
        """견적 완료된 의뢰행을 연두색으로 하이라이트."""
        color = QBrush(QColor(160, 220, 160))
        for col in range(1, self.req_table.columnCount()):
            it = self.req_table.item(row, col)
            if it:
                it.setBackground(color)

    def _add_done(self, path: str) -> None:
        self._list_all_files.insert(0, path)
        if self._quote_list_window:
            self._quote_list_window.add_item(path)

    def _add_done_cover(self, path: str) -> None:
        self._list_all_files.insert(0, path)
        if self._quote_list_window:
            self._quote_list_window.add_item(path)

    def _add_done_submit(self, path: str) -> None:
        self._submit_paths.add(path)
        self._list_all_files.insert(0, path)
        if self._quote_list_window:
            self._quote_list_window.add_submit_item(path)

    @staticmethod
    def _parse_quote_label(full_path: str) -> Tuple[Optional[str], str]:
        """파일 경로에서 (설비호기, 표시레이블) 추출.

        반환값:
        - (None, '')    : 패턴 불일치 → 목록에서 제외
        - ('', label)   : 갑지 파일  → label 끝이 '_갑지', 붉은색으로 표시
        - (tool, label) : 일반 견적서 → 라인_공정_설비호기 형식
        """
        import re as _re
        fn = os.path.basename(full_path)
        if not fn.endswith('.xlsx'):
            return None, ""
        stem   = fn[:-5]
        folder = os.path.basename(os.path.dirname(full_path))

        # ── 갑지: tool='' 로 반환해 붉은색 음영 처리 신호 ───────────
        if _re.search(r'_갑지(?:_\d+)?$', stem):
            m_dom = _re.match(r'^\d{6}_LOT베큠_(.+)$', folder)
            if m_dom:
                return "", f"{m_dom.group(1)}_갑지"
            m_cn = _re.match(r'^\d{6}_중국_SCS_(.+)$', folder)
            if m_cn:
                return "", f"중국_{m_cn.group(1)}_갑지"
            return "", "갑지"

        # ── 미국 Quotation: ..._{tool}_Quotation ─────────────────────
        m_us = _re.search(r'_([^_]+)_Quotation$', stem)
        if m_us:
            tool = m_us.group(1)
            mf   = _re.match(r'^\d{6}_미국_(.+)$', folder)
            return tool, f"미국_{mf.group(1) if mf else '미국'}_{tool}"

        # ── 국내·중국: 마지막 _구분자 이후 = 설비호기 ───────────────
        parts = stem.split('_')
        tool  = parts[-1] if parts else ""
        if not tool:
            return None, ""

        m_dom = _re.match(r'^\d{6}_LOT베큠_(.+)$', folder)
        if m_dom:
            lp_parts = m_dom.group(1).split('_')
            lineproc = '_'.join(lp_parts[:-1]) if len(lp_parts) >= 2 else m_dom.group(1)
            return tool, f"{lineproc}_{tool}"

        m_cn = _re.match(r'^\d{6}_중국_SCS_(.+)$', folder)
        if m_cn:
            return tool, f"중국_{m_cn.group(1)}_{tool}"

        return tool, tool

    @staticmethod
    def _parse_quote_entry(full_path: str) -> Optional[Tuple[str, str, str]]:
        """_parse_quote_label 을 호출하고 (tool, label, full_path) 튜플로 감싼다.
        패턴 불일치(tool is None)이면 None 반환."""
        tool, label = QuoteBuilderPage._parse_quote_label(full_path)
        if tool is None:
            return None
        return tool, label, full_path

    def _open_quote_list_window(self) -> None:
        """견적서 LIST 독립 창을 열거나 앞으로 가져온다."""
        if self._quote_list_window is None:
            self._quote_list_window = _QuoteListWindow(self)
            self._quote_list_window.refresh_session(
                self._list_all_files, frozenset(self._submit_paths))
        self._quote_list_window.show_and_raise()

    _NOTICE_PREFIXES = ("[설정 확인]", "[통합양식 경고]", "[단가 불일치]")

    def _log(self, text: str) -> None:
        """작업 기록. 화면의 로그 창은 없앴다 — 방금 한 일은 위쪽 한 줄에 잠깐 보이고,
        양식에 대한 알림은 '생성 전 점검'에 참고 항목으로 남는다."""
        logger.info(text)
        if text.startswith(self._NOTICE_PREFIXES):
            if text not in self._notices:
                self._notices.append(text)
                self._schedule_review()
            return
        if hasattr(self, "lbl_last"):
            self.lbl_last.setText(text if len(text) <= 70 else text[:69] + "…")
            self.lbl_last.setToolTip(text)
