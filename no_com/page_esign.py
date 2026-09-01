"""
page_esign.py
=============
전자서명 페이지와 Excel 시트 캡처 스레드.
"""
from __future__ import annotations

import glob
import logging
import os
import shutil
import traceback
from typing import List, Optional

from PySide6.QtCore import Qt, QBuffer, QByteArray, QPointF, QThread, QTimer, Signal
from PySide6.QtGui import QImage, QPixmap
from PySide6.QtWidgets import (
    QApplication, QDialog, QFileDialog, QFrame, QHBoxLayout, QLabel,
    QListWidget, QListWidgetItem, QMessageBox, QProgressDialog, QPushButton, QSizePolicy,
    QSplitter, QVBoxLayout, QWidget,
)

import app_settings
from core import ensure_dir, unique_path
from widgets import PdfView, SignatureItem, PasswordDialog, tint_button
from page_common import _friendly_error_msg, _natural_key, _ScrollableErrorDialog

logger = logging.getLogger("QuoteApp")


class _ExcelLoaderThread(QThread):
    """전자서명 페이지용: Excel 시트를 CopyPicture로 캡처 (ExportAsFixedFormat 미사용 → RenameFile 없음)."""
    progress = Signal(int, int, str)   # (완료수, 전체수, 현재파일명)

    def __init__(self, paths: List[str], tmp_dir: str, parent=None) -> None:
        super().__init__(parent)
        self.paths      = paths
        self.tmp_dir    = tmp_dir
        self.sheet_pngs : List[List[str]] = []   # 파일별 PNG 경로 리스트
        self._cancel    = False

    def cancel(self) -> None:
        self._cancel = True

    def run(self) -> None:
        import excel_io
        total = len(self.paths)
        com_ctx = excel_io.ExcelCOM()
        xl_app = None
        try:
            com_ctx.__enter__()
            xl_app = com_ctx.app
        except Exception as e:
            logger.error("Excel COM 초기화 실패: %s", e, exc_info=True)

        try:
            for i, xlsx in enumerate(self.paths):
                if self._cancel:
                    break
                self.progress.emit(i, total, os.path.basename(xlsx))
                try:
                    pngs = excel_io.excel_capture_sheets_to_pngs(
                        xlsx, self.tmp_dir, i + 1, xl_app)
                    self.sheet_pngs.append(pngs)
                except Exception as e:
                    logger.error("시트 캡처 실패 (%s): %s", xlsx, e, exc_info=True)
                    self.sheet_pngs.append([])
        finally:
            try:
                com_ctx.__exit__(None, None, None)
            except Exception:
                pass

        self.progress.emit(len(self.sheet_pngs), total, "완료")


# ══════════════════════════════════════════════
# 전자서명 페이지
# ══════════════════════════════════════════════

class ESignPage(QWidget):

    SIGN_W = 170
    SIGN_H = 40

    def __init__(self) -> None:
        super().__init__()
        # ESignPage 전용 Qt 클래스: 이 페이지가 최초 생성될 때만 임포트
        global QGraphicsScene, QGraphicsPixmapItem, QPointF, QBuffer, QByteArray, QIODevice
        from PySide6.QtCore import QPointF, QBuffer, QByteArray, QIODevice
        from PySide6.QtWidgets import QGraphicsScene, QGraphicsPixmapItem
        self._code        : str  = ""
        self._signs       : List = []
        self._files       : List[str] = []
        self._base_folder : str  = ""
        self._sheet_pngs  : List[List[str]] = []
        self._cur_pngs    : List[str] = []
        self._cur_file    : int  = 0
        self._cur_page    : int  = 0
        self._sign_items    : dict = {}
        self._render_sz     : dict = {}
        self._bg_item              = None
        self._shown_key     : Optional[tuple] = None
        self._loader_thread    : Optional[_ExcelLoaderThread] = None
        self._load_progress    : Optional[QProgressDialog]   = None
        self._tmp_dir          : Optional[str]               = None
        self._com_init_timer   : Optional[QTimer]            = None
        self._com_init_ok      : bool                        = False
        self._build_ui()

    def _build_ui(self) -> None:
        outer = QVBoxLayout(self); outer.setContentsMargins(14,14,14,14); outer.setSpacing(10)
        top = QHBoxLayout()
        self.btn_code   = QPushButton("승인코드 LOAD")
        self.btn_excel  = QPushButton("엑셀 LOAD")
        self.btn_save   = QPushButton("PDF 저장")
        tint_button(self.btn_code,  "#DCEDC8")   # 연초록
        tint_button(self.btn_excel, "#B3E5FC")   # 연하늘
        tint_button(self.btn_save,  "#FFE0B2")   # 연주황
        self.lbl_status = QLabel("준비")
        for b in (self.btn_code, self.btn_excel, self.btn_save): top.addWidget(b); top.addSpacing(8)
        top.addStretch(1); top.addWidget(self.lbl_status); outer.addLayout(top)
        mid = QHBoxLayout()
        self.file_list = QListWidget(); self.file_list.setFixedWidth(360); mid.addWidget(self.file_list)
        self.scene = QGraphicsScene(self)
        self.view  = PdfView(self); self.view.setScene(self.scene); self.view.setAlignment(Qt.AlignCenter)
        self.view.setFocusPolicy(Qt.StrongFocus); self.view.setFocus(); mid.addWidget(self.view, 1)
        outer.addLayout(mid, 1)
        self.btn_code.clicked.connect(self._load_code)
        self.btn_excel.clicked.connect(self._load_excels)
        self.btn_save.clicked.connect(self._save_pdf)
        self.file_list.currentRowChanged.connect(self._on_select_file)
        self.view.on_prev = self._prev_page; self.view.on_next = self._next_page
        self.view.on_double_click = self._add_sign

    def _load_code(self) -> None:
        path, _ = QFileDialog.getOpenFileName(self, "승인코드 TXT", "", "Text Files (*.txt)")
        if not path: return
        try:
            with open(path, encoding="utf-8") as f:
                self._code = f.read().strip()
        except Exception as e:
            QMessageBox.critical(self, "오류", str(e))
            return
        folder = os.path.dirname(path)
        import glob as _glob
        imgs = sorted(
            [p for ext in ("*.png","*.jpg","*.jpeg","*.bmp","*.webp")
             for p in _glob.glob(os.path.join(folder, ext))
             + _glob.glob(os.path.join(folder, ext.upper()))],
            key=lambda p: os.path.basename(p).lower()
        )
        def _pick(kws):
            for kw in kws:
                for p in imgs:
                    if kw in os.path.basename(p).lower(): return p
            return None
        p1 = _pick(["서명1","sign1","signature1","stamp1"]); p2 = _pick(["서명2","sign2","signature2","stamp2"])
        if not p1 or not p2:
            p1 = p1 or (imgs[0] if imgs else None); p2 = p2 or (imgs[1] if len(imgs)>1 else None)
        if not p1 or not p2:
            QMessageBox.warning(self, "안내", "서명 이미지 2개를 찾지 못했습니다.")
            self._signs = []; self.lbl_status.setText("승인코드 OK, 서명 없음"); return
        pm1 = QPixmap(p1).scaled(self.SIGN_W,self.SIGN_H,Qt.IgnoreAspectRatio,Qt.SmoothTransformation)
        pm2 = QPixmap(p2).scaled(self.SIGN_W,self.SIGN_H,Qt.IgnoreAspectRatio,Qt.SmoothTransformation)
        if pm1.isNull() or pm2.isNull(): QMessageBox.critical(self,"오류","서명 이미지 로드 실패"); return
        self._signs = [pm1, pm2]
        self.lbl_status.setText(f"승인코드 OK / {os.path.basename(p1)}, {os.path.basename(p2)}")
        QMessageBox.information(self, "완료", "승인코드 LOAD 완료\n전자서명 ON")

    def _load_excels(self) -> None:
        import excel_io
        if not excel_io._ensure_com():
            QMessageBox.critical(self, "오류", "Excel COM이 없습니다.")
            return
        if self._loader_thread and self._loader_thread.isRunning():
            QMessageBox.information(self, "안내", "이미 로딩 중입니다.")
            return
        start = app_settings.get_dir(app_settings.Key.ESIGN_DIR)
        paths, _ = QFileDialog.getOpenFileNames(
            self, "엑셀 선택(다중)", start, "Excel Files (*.xlsx)")
        if not paths:
            return
        paths = sorted(paths, key=lambda p: (0 if "갑지" in os.path.basename(p).lower() else 1, _natural_key(os.path.basename(p))))
        base = os.path.commonpath(paths)
        if os.path.isfile(base):
            base = os.path.dirname(base)
        self._base_folder = base
        app_settings.set_str(app_settings.Key.ESIGN_DIR, base)
        self._files = paths
        self._sheet_pngs = []
        self._cur_pngs   = []
        self._sign_items.clear()

        self.file_list.blockSignals(True)
        self.file_list.clear()
        for p in paths:
            it = QListWidgetItem(os.path.basename(p))
            it.setData(Qt.UserRole, p)
            self.file_list.addItem(it)
        self.file_list.blockSignals(False)

        self._cleanup_tmp()
        tmp = ensure_dir(os.path.join(base, "_esign_tmp_pdf"))
        self._tmp_dir = tmp

        self._load_progress = QProgressDialog("변환 준비 중...", "취소", 0, len(paths), self)
        self._load_progress.setWindowTitle("엑셀 → PDF 변환")
        self._load_progress.setWindowModality(Qt.WindowModal)
        self._load_progress.setMinimumDuration(0)
        self._load_progress.setValue(0)

        self._loader_thread = _ExcelLoaderThread(paths, tmp, self)
        self._loader_thread.progress.connect(self._on_load_progress)
        self._loader_thread.finished.connect(self._on_load_finished)
        self._load_progress.canceled.connect(self._loader_thread.cancel)

        # COM DispatchEx 무한 블로킹 방지: 30초 내 첫 progress 없으면 강제 종료
        self._com_init_ok = False
        self._com_init_timer = QTimer(self)
        self._com_init_timer.setSingleShot(True)
        self._com_init_timer.timeout.connect(self._on_com_init_timeout)
        self._com_init_timer.start(30_000)

        self.btn_code.setEnabled(False)
        self.btn_excel.setEnabled(False)
        self.btn_save.setEnabled(False)
        self._loader_thread.start()

    def _on_load_progress(self, done: int, total: int, fname: str) -> None:
        if self._load_progress is None:
            return
        # COM init 성공 확인 → 타임아웃 타이머 해제
        if not self._com_init_ok:
            self._com_init_ok = True
            if self._com_init_timer:
                self._com_init_timer.stop()
        self._load_progress.setValue(done)
        # setValue(max) 가 QProgressDialog 자동 닫기(hide)를 트리거하고,
        # hide() 중 Qt 이벤트가 재진입해 _on_load_finished 가 동기 실행될 수 있다.
        # 그 경우 _load_progress 가 None 으로 바뀌므로 재확인 후 접근한다.
        if self._load_progress is None:
            return
        if done < total:
            self._load_progress.setLabelText(f"변환 중 ({done + 1}/{total}): {fname}")
        else:
            self._load_progress.setLabelText("변환 완료")

    def _on_com_init_timeout(self) -> None:
        """Excel COM DispatchEx가 30초 내에 응답하지 않으면 스레드를 강제 종료한다."""
        if self._com_init_ok:
            return
        logger.error("Excel COM 초기화 30초 타임아웃 — 스레드 강제 종료")
        if self._loader_thread and self._loader_thread.isRunning():
            self._loader_thread.terminate()
            self._loader_thread.wait(3000)
        if self._load_progress:
            self._load_progress.close()
            self._load_progress = None
        self.btn_code.setEnabled(True)
        self.btn_excel.setEnabled(True)
        self.btn_save.setEnabled(True)
        QMessageBox.critical(
            self, "Excel COM 타임아웃",
            "Excel COM 초기화가 30초를 초과했습니다.\n\n"
            "가능한 원인:\n"
            "  • Office 활성화 창이 백그라운드에서 대기 중\n"
            "  • 이전 Excel 충돌 복구 대화창 열려 있음\n"
            "  • COM 등록 손상\n\n"
            "Excel을 직접 열어 완료한 뒤 다시 시도하세요."
        )

    def _on_load_finished(self) -> None:
        if self._com_init_timer:
            self._com_init_timer.stop()
        if self._load_progress:
            self._load_progress.close()
            self._load_progress = None
        self.btn_code.setEnabled(True)
        self.btn_excel.setEnabled(True)
        self.btn_save.setEnabled(True)
        self._sheet_pngs = self._loader_thread.sheet_pngs
        while len(self._sheet_pngs) < len(self._files):
            self._sheet_pngs.append([])
        self.btn_excel.setEnabled(True)
        self.lbl_status.setText(f"{len(self._files)}개 로드 완료")
        if self.file_list.count() > 0:
            self.file_list.setCurrentRow(0)

    def _on_select_file(self, row: int) -> None:
        if row < 0 or row >= len(self._files): return
        self._cur_file = row; self._cur_page = 0; self._load_sheets(); self._render()

    def _load_sheets(self) -> None:
        self._cur_pngs = (self._sheet_pngs[self._cur_file]
                          if self._cur_file < len(self._sheet_pngs) else [])
        if not self._cur_pngs:
            self.lbl_status.setText("표시할 시트 없음(스킵)"); self.scene.clear()

    def _render(self) -> None:
        if not self._cur_pngs: return
        self._cur_page = max(0, min(self._cur_page, len(self._cur_pngs) - 1))
        pm = QPixmap(self._cur_pngs[self._cur_page])
        if pm.isNull(): return
        self._render_sz[(self._cur_file, self._cur_page)] = (pm.width(), pm.height())
        if self._shown_key is not None:
            for it in list(self._sign_items.get(self._shown_key, [])):
                try:
                    if it.scene() is self.scene: self.scene.removeItem(it)
                except RuntimeError: pass
        if self._bg_item is not None:
            try:
                if self._bg_item.scene() is self.scene: self.scene.removeItem(self._bg_item)
            except Exception: pass
        bg = QGraphicsPixmapItem(pm); bg.setZValue(0); bg.setAcceptedMouseButtons(Qt.NoButton)
        self.scene.addItem(bg); self._bg_item = bg; self._shown_key = (self._cur_file, self._cur_page)
        for it in list(self._sign_items.get(self._shown_key, [])):
            try: self.scene.addItem(it); it.setZValue(10)
            except RuntimeError: pass
        self.scene.setSceneRect(bg.boundingRect()); self.view.resetTransform()
        vp_w = max(1, self.view.viewport().width())
        scale = vp_w / max(1, pm.width()); self.view.scale(scale, scale); self.view.setFocus()
        self.view.verticalScrollBar().setValue(self.view.verticalScrollBar().minimum())
        self.lbl_status.setText(f"파일 {self._cur_file+1}/{len(self._files)} / 시트 {self._cur_page+1}/{len(self._cur_pngs)}")

    def _next_page(self) -> None:
        if not self._cur_pngs: return
        if self._cur_page+1 < len(self._cur_pngs): self._cur_page += 1; self._render()
        elif self._cur_file+1 < len(self._files): self.file_list.setCurrentRow(self._cur_file+1)

    def _prev_page(self) -> None:
        if not self._cur_pngs: return
        if self._cur_page-1 >= 0: self._cur_page -= 1; self._render()
        elif self._cur_file-1 >= 0:
            self.file_list.setCurrentRow(self._cur_file-1)
            if self._cur_pngs: self._cur_page = max(0, len(self._cur_pngs)-1); self._render()

    def _add_sign(self, scene_pos: QPointF) -> None:
        if not self._signs:
            QMessageBox.information(self, "안내", "승인코드 LOAD 후 서명 이미지가 필요합니다."); return
        if not self._cur_pngs: return
        dlg = PasswordDialog(self, self._code)
        if dlg.exec() != QDialog.Accepted or not dlg.verified: return
        idx = min(1 if (QApplication.keyboardModifiers() & Qt.ShiftModifier) else 0, len(self._signs)-1)
        key = (self._cur_file, self._cur_page)
        item = SignatureItem(self._signs[idx], self._cur_page,
                             on_delete=self._remove_sign)
        item.setZValue(10)
        item.setPos(QPointF(scene_pos.x()-self.SIGN_W/2, scene_pos.y()-self.SIGN_H/2))
        self.scene.addItem(item)
        self._sign_items.setdefault(key, []).append(item)
        self.scene.update()

    def _remove_sign(self, item) -> None:
        """SignatureItem 우클릭 삭제 콜백 — _sign_items 에서도 제거한다.

        여기서 빼주지 않으면 씬에서만 사라지고 목록에는 남아
        페이지를 다시 그릴 때 되살아나며, _build_pdf 가 그대로 PDF 에 찍는다.
        """
        for key, lst in list(self._sign_items.items()):
            if item in lst:
                lst.remove(item)
                if not lst:
                    self._sign_items.pop(key, None)
                break

    def _cleanup_tmp(self) -> None:
        self._cur_pngs = []
        if self._tmp_dir and os.path.isdir(self._tmp_dir):
            try:
                shutil.rmtree(self._tmp_dir)
            except Exception as e:
                logger.warning("tmp 폴더 삭제 실패: %s", e)
            self._tmp_dir = None

    def _save_pdf(self) -> None:
        if not any(self._sheet_pngs) or not self._files:
            QMessageBox.information(self, "안내", "먼저 엑셀을 LOAD 하세요."); return
        folder_name = os.path.basename(self._base_folder.rstrip("\\/"))
        out = unique_path(os.path.join(self._base_folder, f"대외비_{folder_name}.pdf"))
        try:
            self._build_pdf(out)
            self._cleanup_tmp()
            QMessageBox.information(self, "완료", f"저장 완료:\n{out}"); self.lbl_status.setText("PDF 저장 완료")
        except Exception as e:
            logger.error("PDF 저장 실패", exc_info=True)
            tb = traceback.format_exc()
            user_msg, hint = _friendly_error_msg(e)
            _ScrollableErrorDialog(self, tb, user_msg=user_msg, hint=hint).exec()

    def _build_pdf(self, out: str) -> None:
        import fitz
        from PySide6.QtGui import QPainter, QImage
        A4_W, A4_H = 595.0, 842.0       # A4 in points (1pt = 1/72 inch)
        MAX_PX = int(A4_W * 150 / 72)   # 방안A: A4 150 DPI 기준 최대 너비 ≈ 1240px
        final = fitz.open()

        for fi, pngs in enumerate(self._sheet_pngs):
            for pno, png_path in enumerate(pngs):
                if not os.path.exists(png_path):
                    continue
                base_pm = QPixmap(png_path)
                if base_pm.isNull():
                    continue
                iw, ih = base_pm.width(), base_pm.height()
                signs = self._sign_items.get((fi, pno), [])

                if signs:
                    composed = QPixmap(iw, ih)
                    composed.fill(Qt.white)
                    painter = QPainter(composed)
                    painter.drawPixmap(0, 0, base_pm)
                    for it in list(signs):
                        try:
                            x, y = float(it.pos().x()), float(it.pos().y())
                        except RuntimeError:
                            continue
                        painter.drawPixmap(int(x), int(y), it.pixmap())
                    painter.end()
                    final_pm = composed
                else:
                    final_pm = base_pm

                # 방안A: 해상도 상한 (너비 MAX_PX 초과 시 축소)
                if final_pm.width() > MAX_PX:
                    final_pm = final_pm.scaled(
                        MAX_PX, MAX_PX * 10,
                        Qt.KeepAspectRatio, Qt.SmoothTransformation)

                # 방안B: 입고검수확인서만 그레이스케일 변환
                if "입고검수확인서" in png_path:
                    gray_img = final_pm.toImage().convertToFormat(
                        QImage.Format_Grayscale8)
                    final_pm = QPixmap.fromImage(gray_img)

                iw, ih = final_pm.width(), final_pm.height()

                # QPixmap → JPEG bytes
                ba = QByteArray(); buf = QBuffer(ba); buf.open(QIODevice.WriteOnly)
                final_pm.save(buf, "JPEG", 45); buf.close()

                # 고정 A4 페이지, 이미지를 비율 유지하며 중앙+상단 배치
                page = final.new_page(width=A4_W, height=A4_H)
                scale = min(A4_W / max(1, iw), A4_H / max(1, ih))
                pw, ph = iw * scale, ih * scale
                x0 = (A4_W - pw) / 2
                y0 = 0.0  # 상단 정렬
                page.insert_image(fitz.Rect(x0, y0, x0 + pw, y0 + ph), stream=bytes(ba))

        final.save(out, deflate=True, garbage=4); final.close()
