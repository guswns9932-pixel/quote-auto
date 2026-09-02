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
import threading
import traceback
from typing import List, Optional

from PySide6.QtCore import (
    Qt, QBuffer, QByteArray, QIODevice, QPointF, QRectF, QThread, QTimer, Signal,
)
from PySide6.QtGui import QImage, QPainter, QPixmap
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


class _ImageCache:
    """캡처한 PNG 를 QImage 로 캐싱한다.

    페이지 이동(_render)과 PDF 빌드(_PdfBuildThread)가 같은 파일을 각자
    디스크에서 다시 읽던 것을 없애기 위한 공유 캐시. PDF 빌드는 백그라운드
    스레드에서 이 캐시를 읽고(캐시 미스 시) 채워 넣으므로 락으로 보호한다.
    QImage 자체는 Qt 문서상 어느 스레드에서 만들고 다뤄도 안전하다
    (QPixmap 과 달리 GUI 스레드 전용이 아니다).
    """

    def __init__(self) -> None:
        self._data: dict = {}
        self._lock = threading.Lock()

    def get(self, path: str) -> QImage:
        with self._lock:
            img = self._data.get(path)
        if img is not None:
            return img
        img = QImage(path)
        with self._lock:
            self._data.setdefault(path, img)
            return self._data[path]

    def clear(self) -> None:
        with self._lock:
            self._data.clear()


class _ExcelLoaderThread(QThread):
    """전자서명 페이지용: Excel 시트를 CopyPicture로 캡처 (ExportAsFixedFormat 미사용 → RenameFile 없음)."""
    progress       = Signal(int, int, str)             # (완료 파일수, 전체 파일수, 현재파일명)
    sheet_progress = Signal(int, int, str, int, int, str)
    # (파일idx, 파일전체, 파일명, 완료시트수, 전체시트수, 시트명) — 파일 안에서의 세부 진행

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
                fname = os.path.basename(xlsx)

                def _sheet_cb(done, sheet_total, sheet_name, _i=i, _fname=fname):
                    self.sheet_progress.emit(_i, total, _fname, done, sheet_total, sheet_name)

                try:
                    pngs = excel_io.excel_capture_sheets_to_pngs(
                        xlsx, self.tmp_dir, i + 1, xl_app,
                        progress_cb=_sheet_cb,
                        should_cancel=lambda: self._cancel)
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
        # 서명 배치 시 매번 다시 타이핑하지 않도록, 마지막으로 인증에 성공한
        # 값을 세션(메모리)에만 기억해 다음 창에 미리 채운다. 그래도 창은
        # 매번 뜨고 확인은 필요하다 — 그냥 건너뛰는 게 아니다.
        # app_settings 에 저장하지 않으므로 앱을 껐다 켜면 자동으로 비워진다.
        self._last_password: str = ""
        self._signs       : List = []
        self._files       : List[str] = []
        self._base_folder : str  = ""
        self._sheet_pngs  : List[List[str]] = []
        self._cur_pngs    : List[str] = []
        self._cur_file    : int  = 0
        self._cur_page    : int  = 0
        self._sign_items    : dict = {}
        self._image_cache   : _ImageCache = _ImageCache()
        self._bg_item              = None
        self._shown_key     : Optional[tuple] = None
        self._loader_thread    : Optional[_ExcelLoaderThread] = None
        self._load_progress    : Optional[QProgressDialog]   = None
        self._pdf_thread        : Optional["_PdfBuildThread"] = None
        self._pdf_progress      : Optional[QProgressDialog]   = None
        self._tmp_dir          : Optional[str]               = None
        self._com_init_timer   : Optional[QTimer]            = None
        self._com_init_ok      : bool                        = False
        self._build_ui()

    def _build_ui(self) -> None:
        outer = QVBoxLayout(self); outer.setContentsMargins(14,14,14,14); outer.setSpacing(10)
        top = QHBoxLayout()
        # 순서: 엑셀 LOAD → PDF 저장 → 승인코드 LOAD (+ 로드 상태/서명 미리보기)
        self.btn_excel  = QPushButton("엑셀 LOAD")
        self.btn_save   = QPushButton("PDF 저장")
        self.btn_code   = QPushButton("승인코드 LOAD")
        tint_button(self.btn_excel, "#B3E5FC")   # 연하늘
        tint_button(self.btn_save,  "#FFE0B2")   # 연주황
        tint_button(self.btn_code,  "#DCEDC8")   # 연초록
        top.addWidget(self.btn_excel); top.addSpacing(8)
        top.addWidget(self.btn_save);  top.addSpacing(8)
        top.addWidget(self.btn_code);  top.addSpacing(6)

        # 승인코드 LOAD 바로 옆 — 로드 성공/실패와 실제 찍힐 서명(대표로 서명1
        # 하나만)을 즉시 눈으로 확인할 수 있게. 실제 문서에 찍히는 크기
        # (SIGN_W x SIGN_H)와 동일하게 보여줘야 알아보기 쉽다 — 축소된
        # 썸네일은 작아서 잘 안 보인다.
        self.code_status_frame = QFrame()
        self.code_status_frame.setFrameShape(QFrame.StyledPanel)
        self._set_code_status_style(ok=None)
        csf = QHBoxLayout(self.code_status_frame)
        csf.setContentsMargins(8, 3, 8, 3)
        csf.setSpacing(8)
        self.lbl_code_state = QLabel("승인코드 미로드")
        self.lbl_sign_preview = QLabel()
        self.lbl_sign_preview.setFixedSize(self.SIGN_W, self.SIGN_H)
        self.lbl_sign_preview.setAlignment(Qt.AlignCenter)
        self.lbl_sign_preview.setStyleSheet(
            "background: white; border: 1px solid #CCC;")
        self.lbl_sign_preview.setToolTip(
            "서명1 미리보기 (더블클릭 시 배치)\nShift+더블클릭: 서명2 배치")
        csf.addWidget(self.lbl_code_state)
        csf.addWidget(self.lbl_sign_preview)
        top.addWidget(self.code_status_frame)

        top.addStretch(1)
        self.lbl_status = QLabel("준비")
        top.addWidget(self.lbl_status)
        outer.addLayout(top)
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

    def _set_code_status_style(self, ok: Optional[bool]) -> None:
        """승인코드 상태 프레임 배경색. ok=None(미로드/회색) True(성공/연초록) False(실패/연빨강)."""
        color = {"None": "#F5F5F5", "True": "#E8F5E9", "False": "#FFEBEE"}[str(ok)]
        border = {"None": "#DDD", "True": "#A5D6A7", "False": "#EF9A9A"}[str(ok)]
        self.code_status_frame.setStyleSheet(
            f"QFrame {{ background: {color}; border: 1px solid {border}; border-radius: 4px; }}")

    def _build_sign_preview_pixmap(self, pm: QPixmap) -> QPixmap:
        """승인코드 LOAD 옆에 붙일 서명 미리보기 — 대표로 서명1 하나만,
        실제 문서에 찍히는 크기(SIGN_W x SIGN_H)와 같게 보여준다."""
        return pm.scaled(self.SIGN_W, self.SIGN_H,
                         Qt.KeepAspectRatio, Qt.SmoothTransformation)

    def _reset_code_status(self, text: str, ok: Optional[bool] = None) -> None:
        self.lbl_code_state.setText(text)
        self.lbl_sign_preview.clear()
        self.lbl_sign_preview.setToolTip("")
        self._set_code_status_style(ok)

    def _load_code(self) -> None:
        path, _ = QFileDialog.getOpenFileName(self, "승인코드 TXT", "", "Text Files (*.txt)")
        if not path: return
        try:
            with open(path, encoding="utf-8") as f:
                self._code = f.read().strip()
        except Exception as e:
            self._reset_code_status("⚠ 승인코드 읽기 실패", ok=False)
            QMessageBox.critical(self, "오류", str(e))
            return
        # 코드를 새로 불러오면 이전 세션에서 기억해둔 값을 지운다 —
        # 새 코드로 처음 서명할 땐 다시 직접 입력해야 한다.
        self._last_password = ""
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
            self._signs = []
            self._reset_code_status("⚠ 서명 이미지 없음", ok=False)
            QMessageBox.warning(self, "안내", "서명 이미지 2개를 찾지 못했습니다.")
            self.lbl_status.setText("승인코드 OK, 서명 없음")
            return
        pm1 = QPixmap(p1).scaled(self.SIGN_W,self.SIGN_H,Qt.IgnoreAspectRatio,Qt.SmoothTransformation)
        pm2 = QPixmap(p2).scaled(self.SIGN_W,self.SIGN_H,Qt.IgnoreAspectRatio,Qt.SmoothTransformation)
        if pm1.isNull() or pm2.isNull():
            self._reset_code_status("⚠ 서명 이미지 로드 실패", ok=False)
            QMessageBox.critical(self,"오류","서명 이미지 로드 실패"); return
        self._signs = [pm1, pm2]

        # 승인코드 LOAD 버튼 옆에 상태 + 실제 찍힐 서명 이미지를 바로 보여준다.
        self.lbl_code_state.setText("✓ 승인코드 로드됨")
        self._set_code_status_style(ok=True)
        self.lbl_sign_preview.setPixmap(self._build_sign_preview_pixmap(pm1))
        self.lbl_sign_preview.setToolTip(
            f"서명1(대표): {os.path.basename(p1)}  (더블클릭)\n"
            f"서명2: {os.path.basename(p2)}  (Shift+더블클릭)")

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
        self._image_cache.clear()

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
        self._loader_thread.sheet_progress.connect(self._on_sheet_progress)
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

    def _on_sheet_progress(self, file_idx: int, file_total: int, fname: str,
                            sheet_done: int, sheet_total: int, sheet_name: str) -> None:
        """파일 안에서의 시트 단위 진행 — 라벨만 갱신, setValue 는 건드리지 않는다
        (진행 다이얼로그의 숫자 범위는 파일 개수 기준을 그대로 유지)."""
        if self._load_progress is None:
            return
        self._load_progress.setLabelText(
            f"변환 중 ({file_idx + 1}/{file_total}): {fname} — 시트 {sheet_done}/{sheet_total}: {sheet_name}")

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
        path = self._cur_pngs[self._cur_page]
        # 캐시 경유 — 페이지를 앞뒤로 오가도 같은 파일을 디스크에서 다시 읽지 않는다.
        # PDF 저장 단계도 이 캐시를 공유해 이미 본 페이지는 재사용한다.
        img = self._image_cache.get(path)
        if img.isNull(): return
        pm = QPixmap.fromImage(img)
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
        dlg = PasswordDialog(self, self._code, prefill=self._last_password)
        if dlg.exec() != QDialog.Accepted or not dlg.verified: return
        self._last_password = self._code   # 다음 서명부터는 미리 채워서 뜬다
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

    def _collect_build_plan(self) -> list:
        """서명 오버레이를 QImage 로 미리 뽑아 (fi, pno, png_path, overlays) 목록으로 만든다.

        SignatureItem.pixmap()/pos() 는 GUI 스레드에서만 접근 가능하므로,
        백그라운드 스레드(_PdfBuildThread)로 넘기기 전에 여기서 전부 값으로
        떠 둔다. overlays 의 각 항목은 (QImage, x, y) — QImage 는 스레드
        경계를 넘나들어도 안전하다.
        """
        plan = []
        for fi, pngs in enumerate(self._sheet_pngs):
            for pno, png_path in enumerate(pngs):
                overlays = []
                for it in list(self._sign_items.get((fi, pno), [])):
                    try:
                        x, y = float(it.pos().x()), float(it.pos().y())
                        img = it.pixmap().toImage()
                    except RuntimeError:
                        continue
                    overlays.append((img, x, y))
                plan.append((fi, pno, png_path, overlays))
        return plan

    def _save_pdf(self) -> None:
        if not any(self._sheet_pngs) or not self._files:
            QMessageBox.information(self, "안내", "먼저 엑셀을 LOAD 하세요."); return
        if self._pdf_thread and self._pdf_thread.isRunning():
            QMessageBox.information(self, "안내", "이미 PDF 저장 중입니다."); return

        folder_name = os.path.basename(self._base_folder.rstrip("\\/"))
        out = unique_path(os.path.join(self._base_folder, f"대외비_{folder_name}.pdf"))
        plan = self._collect_build_plan()

        # 취소 불가 대기 다이얼로그 — page_common._BgWorker.run_with_progress 와
        # 같은 패턴(cancelButtonText=None). fitz 문서 작성 도중 취소하면 PDF가
        # 반쯤 쓰인 상태로 남는 처리가 새로 필요해져 범위를 늘리므로 지금은 두지 않는다.
        self._pdf_progress = QProgressDialog("PDF 저장 준비 중…", None, 0, len(plan), self)
        self._pdf_progress.setWindowTitle("PDF 저장")
        self._pdf_progress.setWindowModality(Qt.WindowModal)
        self._pdf_progress.setMinimumDuration(0)
        self._pdf_progress.setValue(0)

        self._pdf_thread = _PdfBuildThread(plan, self._image_cache, out, self)
        self._pdf_thread.progress.connect(self._on_pdf_progress)
        self._pdf_thread.done.connect(lambda result: self._on_pdf_done(result, out))

        self.btn_code.setEnabled(False)
        self.btn_excel.setEnabled(False)
        self.btn_save.setEnabled(False)
        self._pdf_thread.start()

    def _on_pdf_progress(self, done: int, total: int) -> None:
        if self._pdf_progress is None:
            return
        self._pdf_progress.setValue(done)
        # setValue(max) 가 다이얼로그 자동 닫기를 트리거하고, 그 과정에서
        # Qt 이벤트가 재진입해 _on_pdf_done 이 동기 실행될 수 있다
        # (_on_load_progress 에서 이미 겪은 것과 같은 패턴 — 재확인 필수).
        if self._pdf_progress is None:
            return
        self._pdf_progress.setLabelText(f"PDF 저장 중 ({done}/{total})")

    def _on_pdf_done(self, result, out: str) -> None:
        if self._pdf_progress:
            self._pdf_progress.close()
            self._pdf_progress = None
        self.btn_code.setEnabled(True)
        self.btn_excel.setEnabled(True)
        self.btn_save.setEnabled(True)
        if isinstance(result, Exception):
            logger.error("PDF 저장 실패", exc_info=result)
            tb = "".join(traceback.format_exception(type(result), result, result.__traceback__))
            user_msg, hint = _friendly_error_msg(result)
            _ScrollableErrorDialog(self, tb, user_msg=user_msg, hint=hint).exec()
            return
        self._cleanup_tmp()
        QMessageBox.information(self, "완료", f"저장 완료:\n{out}")
        self.lbl_status.setText("PDF 저장 완료")


class _PdfBuildThread(QThread):
    """서명이 찍힌 PDF 합성·저장을 백그라운드에서 수행.

    QPixmap/QPainter-on-widget 은 GUI 스레드 전용이라 여기서는 전부 QImage 로
    처리한다(Qt 문서: QImage 는 어느 스레드에서 다뤄도 안전). fitz(PyMuPDF)
    문서 작성도 이 스레드 하나로만 국한되므로(다른 스레드와 동시에 같은
    문서를 건드리지 않음) 안전하다.

    다운스케일을 먼저 하고 그 위에 서명을 합성한다(기존엔 원본 해상도로
    합성한 뒤 축소) — 서명의 위치·크기도 같은 배율로 줄여서 최종 결과물의
    상대적 위치/크기는 기존과 동일하게 유지한다.
    """
    progress = Signal(int, int)   # (완료 페이지수, 전체 페이지수)
    done     = Signal(object)     # None(성공) 또는 Exception

    A4_W, A4_H = 595.0, 842.0                  # A4, pt (1pt = 1/72 inch)
    MAX_PX = int(A4_W * 150 / 72)              # 150 DPI 기준 최대 너비 ≈ 1240px

    def __init__(self, plan: list, image_cache: "_ImageCache", out_path: str,
                 parent=None) -> None:
        super().__init__(parent)
        self.plan = plan
        self.image_cache = image_cache
        self.out_path = out_path

    def run(self) -> None:
        try:
            self._build()
            self.done.emit(None)
        except Exception as e:
            logger.error("PDF 빌드 스레드 실패", exc_info=True)
            self.done.emit(e)

    def _build(self) -> None:
        import fitz

        final = fitz.open()
        total = len(self.plan)

        for done_i, (fi, pno, png_path, overlays) in enumerate(self.plan):
            self.progress.emit(done_i, total)
            if not os.path.exists(png_path):
                continue
            base = self.image_cache.get(png_path)
            if base.isNull():
                continue

            # 다운스케일 먼저 — 원본 해상도로 합성한 뒤 축소하던 것을 뒤집었다.
            # 캡처가 고DPI 환경일수록(1600~2400px) 절약 폭이 커진다.
            orig_w = base.width()
            if orig_w > self.MAX_PX:
                scaled = base.scaledToWidth(self.MAX_PX, Qt.SmoothTransformation)
            else:
                scaled = base
            scale = (scaled.width() / orig_w) if orig_w else 1.0

            if overlays:
                composed = QImage(scaled.size(), QImage.Format_ARGB32_Premultiplied)
                composed.fill(Qt.white)
                painter = QPainter(composed)
                painter.drawImage(0, 0, scaled)
                for sign_img, x, y in overlays:
                    sw = sign_img.width() * scale
                    sh = sign_img.height() * scale
                    painter.drawImage(QRectF(x * scale, y * scale, sw, sh), sign_img)
                painter.end()
                final_img = composed
            else:
                final_img = scaled

            if "입고검수확인서" in png_path:
                final_img = final_img.convertToFormat(QImage.Format_Grayscale8)

            iw, ih = final_img.width(), final_img.height()
            ba = QByteArray(); buf = QBuffer(ba); buf.open(QIODevice.WriteOnly)
            final_img.save(buf, "JPEG", 45); buf.close()

            # 고정 A4 페이지, 이미지를 비율 유지하며 중앙+상단 배치
            page = final.new_page(width=self.A4_W, height=self.A4_H)
            pscale = min(self.A4_W / max(1, iw), self.A4_H / max(1, ih))
            pw, ph = iw * pscale, ih * pscale
            x0 = (self.A4_W - pw) / 2
            page.insert_image(fitz.Rect(x0, 0.0, x0 + pw, ph), stream=bytes(ba))

        self.progress.emit(total, total)
        final.save(self.out_path, deflate=True, garbage=4)
        final.close()
