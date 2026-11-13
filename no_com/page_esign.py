"""
page_esign.py
=============
전자서명 화면.

엑셀을 PDF 로 바꿔 불러오고(esign_engine), 쪽 위에 서명을 놓은 뒤 제출용 PDF 로
저장한다. 변환·합성 규칙은 전부 esign_engine 에 있고 여기는 화면만 다룬다.

화면 좌표 = PDF 포인트. 쪽 그림은 선명하게 보이도록 RENDER_ZOOM 배로 그린 뒤
1/RENDER_ZOOM 로 줄여 올리므로, 서명 아이템의 pos() 가 곧 PDF 위의 위치다.
"""
from __future__ import annotations

import glob
import logging
import os
import shutil
import threading
import traceback
from collections import OrderedDict
from typing import Dict, List, Optional, Tuple

from PySide6.QtCore import (
    Qt, QBuffer, QByteArray, QIODevice, QPointF, QSize, QThread, QTimer, Signal,
)
from PySide6.QtGui import QIcon, QImage, QPixmap
from PySide6.QtWidgets import (
    QApplication, QDialog, QFileDialog, QFrame, QGraphicsPixmapItem, QGraphicsScene,
    QHBoxLayout, QLabel, QListWidget, QListWidgetItem, QMessageBox, QProgressDialog,
    QPushButton, QVBoxLayout, QWidget,
)

import app_settings
import esign_engine as engine
from core import SheetName, unique_path
from widgets import PdfView, SignatureItem, PasswordDialog, tint_button
from page_common import _friendly_error_msg, _natural_key, _ScrollableErrorDialog

logger = logging.getLogger("QuoteApp")

RENDER_ZOOM = 3.0        # 화면용 쪽 그림 배율 (72dpi × 3 = 216dpi)
THUMB_ZOOM = 0.2         # 쪽 미리보기 배율
PAGE_CACHE_MAX = 8       # 쪽 그림을 몇 장까지 기억할지 (장당 십수 MB)


class _ExcelWorkerThread(QThread):
    """전자서명 화면이 살아 있는 동안 Excel COM 세션을 쥐고 있는 워커.

    화면이 열리자마자 백그라운드로 Excel 을 띄워 두고, 사용자가 엑셀을 고르면
    이미 준비된 세션으로 바로 변환을 시작한다(기동에만 1초 남짓 걸린다).

    [왜 스레드를 나누지 않는가]
    COM 객체는 만들어진 스레드에 묶여 있어서, 다른 스레드에서 쓰려면 인터페이스를
    마샬링해야 한다. 실패하면 보이지 않는 EXCEL.EXE 가 남는 골치 아픈 경로가
    생기므로, 이 스레드 하나가 세션 생성 → 변환 → 종료까지 전부 맡는다.
    """
    ready          = Signal(bool)                      # COM 세션 준비 완료(성공 여부)
    progress       = Signal(int, int, str)             # (완료 파일수, 전체 파일수, 현재파일명)
    sheet_progress = Signal(int, int, str, int, int, str)
    # (파일idx, 파일전체, 파일명, 완료시트수, 전체시트수, 시트명)
    job_done       = Signal(object)                    # List[engine.SignFile]

    def __init__(self, parent=None) -> None:
        super().__init__(parent)
        self._job    : Optional[List[str]] = None
        self._wake   = threading.Event()
        self._lock   = threading.Lock()
        self._stop   = False
        self._cancel = False

    def submit(self, paths: List[str]) -> None:
        """변환할 파일 목록을 워커에 넘긴다(GUI 스레드에서 호출)."""
        with self._lock:
            self._job = list(paths)
        self._cancel = False
        self._wake.set()

    def cancel(self) -> None:
        self._cancel = True

    def shutdown(self) -> None:
        """루프를 깨워 종료시킨다 → run() 의 finally 에서 Excel 이 Quit 된다."""
        self._stop = True
        self._cancel = True
        self._wake.set()

    def run(self) -> None:
        import excel_com
        com_ctx = excel_com.ExcelCOM()
        xl_app = None
        try:
            com_ctx.__enter__()
            xl_app = com_ctx.app
        except Exception as e:
            logger.error("Excel COM 초기화 실패: %s", e, exc_info=True)
        # 실패해도 계속 간다 — 파일마다 자체 세션을 여는 폴백이 있다(느릴 뿐 동작함).
        self.ready.emit(xl_app is not None)

        try:
            while not self._stop:
                self._wake.wait()
                self._wake.clear()
                if self._stop:
                    break
                with self._lock:
                    job, self._job = self._job, None
                if job is None:
                    continue
                self._convert_all(excel_com, xl_app, job)
        finally:
            try:
                com_ctx.__exit__(None, None, None)
            except Exception:
                pass

    def _convert_all(self, excel_com, xl_app, paths: List[str]) -> None:
        files: List[engine.SignFile] = []
        total = len(paths)
        targets = list(SheetName.ESIGN_TARGET)
        try:
            for i, xlsx in enumerate(paths):
                if self._cancel:
                    break
                fname = os.path.basename(xlsx)
                self.progress.emit(i, total, fname)

                def _sheet_cb(done, sheet_total, sheet_name, _i=i, _fname=fname):
                    self.sheet_progress.emit(_i, total, _fname, done, sheet_total, sheet_name)

                kw = dict(progress_cb=_sheet_cb, should_cancel=lambda: self._cancel)
                if xl_app is not None:
                    files.append(engine.load_workbook_pdf(xl_app, xlsx, targets, **kw))
                    continue
                try:
                    with excel_com.ExcelCOM() as xl:
                        files.append(engine.load_workbook_pdf(xl.app, xlsx, targets, **kw))
                except Exception as e:
                    logger.error("PDF 변환 실패 (%s): %s", xlsx, e, exc_info=True)
                    files.append(engine.SignFile(path=xlsx, error=str(e)))
        finally:
            self.progress.emit(len(files), total, "완료")
            self.job_done.emit(files)


class _PdfBuildThread(QThread):
    """서명이 찍힌 제출용 PDF 저장을 백그라운드에서 수행."""
    progress = Signal(int, int)   # (완료 쪽수, 전체 쪽수)
    done     = Signal(object)     # None(성공) 또는 Exception

    def __init__(self, files, placements, sign_images, out_path: str, parent=None) -> None:
        super().__init__(parent)
        self._args = (files, placements, sign_images, out_path)

    def run(self) -> None:
        try:
            engine.build_signed_pdf(*self._args, progress_cb=self.progress.emit)
            self.done.emit(None)
        except Exception as e:
            logger.error("PDF 저장 스레드 실패", exc_info=True)
            self.done.emit(e)


# ══════════════════════════════════════════════
# 전자서명 화면
# ══════════════════════════════════════════════

class ESignPage(QWidget):

    def __init__(self) -> None:
        super().__init__()
        self._code          : str = ""
        # 서명을 놓을 때마다 다시 타이핑하지 않도록, 마지막으로 인증에 성공한
        # 값을 세션(메모리)에만 기억해 다음 창에 미리 채운다. 그래도 창은
        # 매번 뜨고 확인은 필요하다. 저장하지 않으므로 앱을 끄면 사라진다.
        self._last_password : str = ""
        self._signs         : List[QPixmap] = []     # 화면에 올릴 서명 그림 (RENDER_ZOOM 배)
        self._sign_bytes    : List[bytes]   = []     # PDF 에 넣을 서명 그림 (PNG)
        self._files         : List[engine.SignFile] = []
        self._docs          : Dict[int, object] = {}             # 파일 순번 → 열린 PDF
        self._page_cache    : "OrderedDict[Tuple[int, int], QPixmap]" = OrderedDict()
        self._base_folder   : str = ""
        self._cur_file      : int = 0
        self._cur_page      : int = 0
        self._sign_items    : Dict[Tuple[int, int], List[SignatureItem]] = {}
        self._last_bulk     : List[SignatureItem] = []           # 방금 일괄 적용으로 놓은 것
        self._bg_item                                 = None
        self._shown_key     : Optional[Tuple[int, int]] = None
        self._worker        : Optional[_ExcelWorkerThread] = None
        self._load_progress : Optional[QProgressDialog]    = None
        self._pdf_thread    : Optional[_PdfBuildThread]    = None
        self._pdf_progress  : Optional[QProgressDialog]    = None
        self._com_init_timer: Optional[QTimer]             = None
        self._com_init_ok   : bool                         = False
        self._build_ui()
        # 화면이 뜨자마자 Excel 을 백그라운드로 띄워 둔다. 사용자가 서명을 고르고
        # 엑셀을 선택하는 동안 기동이 끝난다.
        self._start_worker()

    # ── Excel 워커 (COM 세션 예열) ─────────────────────────────
    def _start_worker(self) -> None:
        """Excel COM 세션을 쥘 워커를 띄운다. 이미 살아 있으면 아무것도 안 한다."""
        if self._worker is not None and self._worker.isRunning():
            return
        self._com_init_ok = False
        w = _ExcelWorkerThread()        # 부모 없이 — 창이 먼저 닫혀도 종료를 끝까지 마치도록
        w.ready.connect(self._on_com_ready)
        w.progress.connect(self._on_load_progress)
        w.sheet_progress.connect(self._on_sheet_progress)
        w.job_done.connect(self._on_job_done)
        self._worker = w
        # DispatchEx 가 무한 블로킹하는 경우(Office 활성화 창 대기 등)를 잡는다.
        # 예열 중에는 기다리는 사람이 없으므로 대화창은 띄우지 않고,
        # 실제로 "엑셀 Load" 를 눌렀을 때만 알린다.
        self._arm_com_timer()
        w.start()

    def _arm_com_timer(self) -> None:
        if self._com_init_timer is None:
            self._com_init_timer = QTimer(self)
            self._com_init_timer.setSingleShot(True)
            self._com_init_timer.timeout.connect(self._on_com_init_timeout)
        self._com_init_timer.start(30_000)

    def _stop_com_timer(self) -> None:
        if self._com_init_timer is not None:
            self._com_init_timer.stop()

    def _on_com_ready(self, ok: bool) -> None:
        self._com_init_ok = True
        self._stop_com_timer()
        if not ok:
            logger.warning("Excel COM 예열 실패 — 파일별 세션으로 동작합니다")

    # ── 종료 ────────────────────────────────────────────────
    # Excel 을 끄는 데(Quit) 환경에 따라 수십 초가 걸린다(실측 50초). 예전에는
    # 창이 그동안 멈춰 있었다. 이제 창은 바로 닫고, 워커가 Excel 종료를 마칠
    # 때까지 프로세스만 뒤에 남는다(app_runner 가 background_busy 를 보고 기다림).
    def shutdown(self) -> None:
        """Excel 세션 정리를 '시작'만 한다. 끝나길 기다리지 않는다."""
        self._stop_com_timer()
        if self._worker is not None:
            self._worker.shutdown()
        self._close_docs()

    def background_busy(self) -> bool:
        """아직 끝나지 않은 뒷정리(Excel 종료)가 있는가."""
        return self._worker is not None and self._worker.isRunning()

    def force_stop(self) -> None:
        """뒷정리가 너무 오래 걸릴 때 마지막 수단."""
        w, self._worker = self._worker, None
        if w is not None and w.isRunning():
            logger.warning("Excel 워커가 제때 끝나지 않아 강제 종료합니다")
            w.terminate()
            w.wait(3000)

    def reset_page(self) -> None:
        """초기화 — 불러온 것과 서명을 모두 비운다. Excel 세션은 그대로 둔다
        (껐다 켜면 그 종료 대기 수십 초를 초기화할 때마다 치르게 된다)."""
        if self._load_progress is not None or (self._pdf_thread and self._pdf_thread.isRunning()):
            QMessageBox.information(self, "안내", "작업이 끝난 뒤 초기화하세요.")
            return
        self._clear_documents()
        self._code = ""
        self._last_password = ""
        self._signs = []
        self._sign_bytes = []
        self._reset_code_status("서명 미로드")
        self.lbl_status.setText("준비")
        self._start_worker()

    # ── 화면 구성 ───────────────────────────────────────────
    @staticmethod
    def _action_btn(label: str, slot) -> QPushButton:
        """견적서작성 화면과 크기·폰트를 통일한다."""
        btn = QPushButton(label); btn.setFixedHeight(52)
        f = btn.font(); f.setPointSize(10); f.setBold(True); btn.setFont(f)
        btn.clicked.connect(slot); return btn

    def _build_ui(self) -> None:
        outer = QVBoxLayout(self); outer.setContentsMargins(14, 14, 14, 14); outer.setSpacing(10)

        # 1행: 버튼 — 작업 순서대로(서명 준비 → 문서 로드 → 저장)
        top = QHBoxLayout()
        self.btn_code  = self._action_btn("① 서명 Load",             self._load_code)
        self.btn_excel = self._action_btn("② 원본 견적(엑셀) Load",   self._load_excels)
        self.btn_save  = self._action_btn("③ 제출용 견적(PDF) Print", self._save_pdf)
        tint_button(self.btn_code,  "#DCEDC8")   # 연초록
        tint_button(self.btn_excel, "#B3E5FC")   # 연하늘
        tint_button(self.btn_save,  "#FFE0B2")   # 연주황
        top.addWidget(self.btn_code); top.addSpacing(8)
        top.addWidget(self.btn_excel); top.addSpacing(8)
        top.addWidget(self.btn_save)
        top.addStretch(1)
        self.lbl_status = QLabel("준비")
        top.addWidget(self.lbl_status)
        outer.addLayout(top)

        # 2행: 서명 상태 + 서명 미리보기 + 일괄 적용
        code_row = QHBoxLayout()
        self.code_status_frame = QFrame()
        self.code_status_frame.setFrameShape(QFrame.StyledPanel)
        self._set_code_status_style(ok=None)
        csf = QHBoxLayout(self.code_status_frame)
        csf.setContentsMargins(8, 6, 8, 6)
        csf.setSpacing(8)
        self.lbl_code_state = QLabel("서명 미로드")
        csf.addWidget(self.lbl_code_state)
        self.lbl_sign_previews: List[QLabel] = []
        for n in (1, 2):
            cap = QLabel(f"서명{n}")
            cap.setStyleSheet("color: #555; border: none; background: transparent;")
            lbl = QLabel()
            lbl.setFixedSize(170, 40)
            lbl.setAlignment(Qt.AlignCenter)
            lbl.setStyleSheet("background: white; border: 1px solid #CCC;")
            csf.addWidget(cap); csf.addWidget(lbl)
            self.lbl_sign_previews.append(lbl)
        code_row.addWidget(self.code_status_frame)
        code_row.addStretch(1)
        self.btn_bulk = QPushButton("이 파일의 서명 위치를 나머지 파일에도 적용")
        self.btn_bulk.setFixedHeight(34)
        self.btn_bulk.setToolTip("지금 보고 있는 파일에 놓은 서명을, 다른 파일의 같은 시트 같은 자리에 놓습니다.\n"
                                 "이미 서명이 있는 쪽은 건드리지 않습니다.")
        self.btn_bulk.clicked.connect(self._apply_to_others)
        self.btn_bulk_undo = QPushButton("일괄 적용 취소")
        self.btn_bulk_undo.setFixedHeight(34)
        self.btn_bulk_undo.clicked.connect(self._undo_bulk)
        code_row.addWidget(self.btn_bulk)
        code_row.addWidget(self.btn_bulk_undo)
        outer.addLayout(code_row)

        # 3행: 왼쪽(파일 목록 + 쪽 미리보기) / 오른쪽(문서)
        mid = QHBoxLayout()
        left = QVBoxLayout(); left.setSpacing(6)
        left.addWidget(QLabel("파일"))
        self.file_list = QListWidget(); self.file_list.setFixedWidth(360)
        left.addWidget(self.file_list, 3)
        left.addWidget(QLabel("쪽"))
        self.page_list = QListWidget(); self.page_list.setFixedWidth(360)
        self.page_list.setIconSize(QSize(84, 119))
        left.addWidget(self.page_list, 4)
        mid.addLayout(left)

        self.scene = QGraphicsScene(self)
        self.view  = PdfView(self); self.view.setScene(self.scene); self.view.setAlignment(Qt.AlignCenter)
        self.view.setFocusPolicy(Qt.StrongFocus); self.view.setFocus()
        self.view.setStyleSheet("background: #6E7480;")
        mid.addWidget(self.view, 1)
        outer.addLayout(mid, 1)

        hint = QLabel("더블클릭: 서명1 놓기   ·   Shift+더블클릭: 서명2 놓기   ·   끌기: 이동   ·   "
                      "우클릭: 삭제   ·   PgUp/PgDn·휠: 쪽 넘기기")
        hint.setStyleSheet("color: #555;")
        outer.addWidget(hint)

        self.file_list.currentRowChanged.connect(self._on_select_file)
        self.page_list.currentRowChanged.connect(self._on_select_page)
        self.view.on_prev = self._prev_page; self.view.on_next = self._next_page
        self.view.on_double_click = self._add_sign
        self._update_buttons()

    def _set_code_status_style(self, ok: Optional[bool]) -> None:
        """서명 로드 상태 프레임 배경색. ok=None(미로드/회색) True(성공/연초록) False(실패/연빨강)."""
        color = {"None": "#F5F5F5", "True": "#E8F5E9", "False": "#FFEBEE"}[str(ok)]
        border = {"None": "#DDD", "True": "#A5D6A7", "False": "#EF9A9A"}[str(ok)]
        self.code_status_frame.setStyleSheet(
            f"QFrame {{ background: {color}; border: 1px solid {border}; border-radius: 4px; }}")

    def _reset_code_status(self, text: str, ok: Optional[bool] = None) -> None:
        self.lbl_code_state.setText(text)
        for lbl in self.lbl_sign_previews:
            lbl.clear(); lbl.setToolTip("")
        self._set_code_status_style(ok)

    def _busy(self) -> bool:
        return self._load_progress is not None or bool(self._pdf_thread and self._pdf_thread.isRunning())

    def _update_buttons(self) -> None:
        busy = self._busy()
        has_pages = any(f.pages for f in self._files)
        self.btn_code.setEnabled(not busy)
        self.btn_excel.setEnabled(not busy)
        self.btn_save.setEnabled(not busy and has_pages)
        cur_signed = any(self._sign_items.get((self._cur_file, p))
                         for p in range(len(self._cur_pages())))
        self.btn_bulk.setEnabled(not busy and cur_signed and len(self._files) > 1)
        self.btn_bulk_undo.setVisible(bool(self._last_bulk))

    # ── ① 서명 Load ─────────────────────────────────────────
    def _load_code(self) -> None:
        path, _ = QFileDialog.getOpenFileName(self, "승인코드 TXT", "", "Text Files (*.txt)")
        if not path: return
        try:
            with open(path, encoding="utf-8") as f:
                self._code = f.read().strip()
        except Exception as e:
            self._reset_code_status("⚠ 서명 읽기 실패", ok=False)
            QMessageBox.critical(self, "오류", str(e))
            return
        # 코드를 새로 불러오면 기억해둔 값을 지운다 — 새 코드로는 다시 직접 입력.
        self._last_password = ""
        folder = os.path.dirname(path)
        imgs = sorted(
            {p for ext in ("*.png", "*.jpg", "*.jpeg", "*.bmp", "*.webp")
             for p in glob.glob(os.path.join(folder, ext))
             + glob.glob(os.path.join(folder, ext.upper()))},
            key=lambda p: os.path.basename(p).lower()
        )
        def _pick(kws):
            for kw in kws:
                for p in imgs:
                    if kw in os.path.basename(p).lower(): return p
            return None
        p1 = _pick(["서명1", "sign1", "signature1", "stamp1"])
        p2 = _pick(["서명2", "sign2", "signature2", "stamp2"])
        if not p1 or not p2:
            p1 = p1 or (imgs[0] if imgs else None); p2 = p2 or (imgs[1] if len(imgs) > 1 else None)
        if not p1 or not p2:
            self._signs = []; self._sign_bytes = []
            self._reset_code_status("⚠ 서명 이미지 없음", ok=False)
            QMessageBox.warning(self, "안내", "서명 이미지 2개를 찾지 못했습니다.")
            self._update_buttons()
            return
        if not self.set_sign_images([QImage(p1), QImage(p2)]):
            self._reset_code_status("⚠ 서명 이미지 로드 실패", ok=False)
            QMessageBox.critical(self, "오류", "서명 이미지 로드 실패"); return

        self.lbl_code_state.setText("✓ 서명 로드됨")
        self._set_code_status_style(ok=True)
        self.lbl_sign_previews[0].setToolTip(f"{os.path.basename(p1)}  (더블클릭으로 놓기)")
        self.lbl_sign_previews[1].setToolTip(f"{os.path.basename(p2)}  (Shift+더블클릭으로 놓기)")
        QMessageBox.information(self, "완료", "서명 Load 완료\n전자서명 ON")

    def set_sign_images(self, images: List[QImage]) -> bool:
        """서명 그림 2개를 받아 화면용·PDF용으로 준비한다. 하나라도 못 읽으면 False."""
        if len(images) < 2 or any(im.isNull() for im in images):
            self._signs = []; self._sign_bytes = []
            return False
        w = int(round(engine.SIGN_W_PT * RENDER_ZOOM)); h = int(round(engine.SIGN_H_PT * RENDER_ZOOM))
        signs, blobs = [], []
        for im in images[:2]:
            # 서명 칸 크기에 맞춰 늘린다(예전과 같은 동작 — 가로세로 비율 무시)
            signs.append(QPixmap.fromImage(im).scaled(w, h, Qt.IgnoreAspectRatio, Qt.SmoothTransformation))
            ba = QByteArray(); buf = QBuffer(ba); buf.open(QIODevice.WriteOnly)
            im.save(buf, "PNG"); buf.close()
            blobs.append(bytes(ba))
        self._signs, self._sign_bytes = signs, blobs
        for lbl, pm in zip(self.lbl_sign_previews, signs):
            lbl.setPixmap(pm.scaled(170, 40, Qt.IgnoreAspectRatio, Qt.SmoothTransformation))
        self._update_buttons()
        return True

    # ── ② 엑셀 Load ─────────────────────────────────────────
    def _load_excels(self) -> None:
        import excel_com
        if not excel_com._ensure_com():
            QMessageBox.critical(self, "오류", "Excel COM이 없습니다.")
            return
        if self._load_progress is not None:
            QMessageBox.information(self, "안내", "이미 로딩 중입니다.")
            return
        start = app_settings.get_dir(app_settings.Key.ESIGN_DIR)
        paths, _ = QFileDialog.getOpenFileNames(
            self, "엑셀 선택(다중)", start, "Excel Files (*.xlsx)")
        if not paths:
            return
        paths = sorted(paths, key=lambda p: (0 if "갑지" in os.path.basename(p).lower() else 1,
                                             _natural_key(os.path.basename(p))))
        base = os.path.commonpath(paths)
        if os.path.isfile(base):
            base = os.path.dirname(base)
        self._clear_documents()
        self._base_folder = base
        app_settings.set_str(app_settings.Key.ESIGN_DIR, base)
        self._files = [engine.SignFile(path=p) for p in paths]
        self._fill_file_list()

        # 흔적 정리: 비정상 종료로 남은 임시 PDF, 구버전이 만들던 임시 폴더
        for folder in {os.path.dirname(os.path.abspath(p)) for p in paths}:
            engine.cleanup_stale_temp(folder)
        old_tmp = os.path.join(base, "_esign_tmp_pdf")
        if os.path.isdir(old_tmp):
            try:
                shutil.rmtree(old_tmp)
            except Exception as e:
                logger.warning("구버전 tmp 폴더 삭제 실패: %s", e)

        self._load_progress = QProgressDialog("변환 준비 중...", "취소", 0, len(paths), self)
        self._load_progress.setWindowTitle("엑셀 → PDF 변환")
        self._load_progress.setWindowModality(Qt.WindowModal)
        self._load_progress.setMinimumDuration(0)
        self._load_progress.setValue(0)

        # 워커는 화면이 열릴 때 이미 떠 있다. 예열이 실패해 죽어 있으면 다시 살린다.
        self._start_worker()
        self._load_progress.canceled.connect(self._worker.cancel)

        # 예열이 아직 안 끝났거나 첫 파일 열기가 무한 블로킹하는 경우를 잡는다.
        self._com_init_ok = False
        self._arm_com_timer()

        self._update_buttons()
        self._worker.submit(paths)

    def _on_load_progress(self, done: int, total: int, fname: str) -> None:
        if self._load_progress is None:
            return
        if not self._com_init_ok:               # 첫 progress 도착 → 감시 해제
            self._com_init_ok = True
            self._stop_com_timer()
        self._load_progress.setValue(done)
        # setValue(max) 가 다이얼로그 자동 닫기를 트리거하고, 그 과정에서 Qt 이벤트가
        # 재진입해 _on_job_done 이 먼저 실행될 수 있다. None 이 됐는지 다시 확인한다.
        if self._load_progress is None:
            return
        if done < total:
            self._load_progress.setLabelText(f"변환 중 ({done + 1}/{total}): {fname}")
        else:
            self._load_progress.setLabelText("변환 완료")

    def _on_sheet_progress(self, file_idx: int, file_total: int, fname: str,
                           sheet_done: int, sheet_total: int, sheet_name: str) -> None:
        if self._load_progress is None:
            return
        self._load_progress.setLabelText(
            f"변환 중 ({file_idx + 1}/{file_total}): {fname} — 시트 {sheet_done}/{sheet_total}: {sheet_name}")

    def _on_com_init_timeout(self) -> None:
        """Excel COM 이 30초 내에 응답하지 않으면 워커를 강제 종료한다.

        예열 중(사용자가 아직 아무것도 안 누른 상태)이면 조용히 워커만 정리한다 —
        실제로 엑셀을 로드할 때 다시 시도된다. 로드 중이었다면 안내까지 띄운다.
        """
        if self._com_init_ok:
            return
        waiting = self._load_progress is not None
        logger.error("Excel COM 초기화 30초 타임아웃 — 워커 강제 종료 (로드 중=%s)", waiting)
        w, self._worker = self._worker, None
        if w is not None and w.isRunning():
            w.terminate()
            w.wait(3000)
        if not waiting:
            return
        self._load_progress.close()
        self._load_progress = None
        self._update_buttons()
        QMessageBox.critical(
            self, "Excel COM 타임아웃",
            "Excel COM 초기화가 30초를 초과했습니다.\n\n"
            "가능한 원인:\n"
            "  • Office 활성화 창이 백그라운드에서 대기 중\n"
            "  • 이전 Excel 충돌 복구 대화창 열려 있음\n"
            "  • COM 등록 손상\n\n"
            "Excel을 직접 열어 완료한 뒤 다시 시도하세요."
        )

    def _on_job_done(self, files: list) -> None:
        self._stop_com_timer()
        if self._load_progress:
            self._load_progress.close()
            self._load_progress = None
        self.show_files(list(files))

    def show_files(self, files: List[engine.SignFile]) -> None:
        """변환이 끝난 파일들을 화면에 올린다. 취소로 빠진 파일은 '취소됨'으로 채운다."""
        done = {f.path: f for f in files}
        wanted = [f.path for f in self._files] or [f.path for f in files]
        self._files = [done.get(p) or engine.SignFile(path=p, error="취소됨") for p in wanted]
        if not self._base_folder and self._files:
            self._base_folder = os.path.dirname(os.path.abspath(self._files[0].path))
        self._fill_file_list()
        ok = sum(1 for f in self._files if f.pages)
        bad = len(self._files) - ok
        self.lbl_status.setText(f"{ok}개 로드 완료" + (f" · {bad}개 실패" if bad else ""))
        self._update_buttons()
        first = next((i for i, f in enumerate(self._files) if f.pages), -1)
        if first >= 0:
            self.file_list.setCurrentRow(first)

    # ── 문서 상태 ───────────────────────────────────────────
    def _close_docs(self) -> None:
        for d in self._docs.values():
            try:
                d.close()
            except Exception:
                pass
        self._docs = {}

    def _clear_documents(self) -> None:
        self.scene.clear()
        self._bg_item = None
        self._shown_key = None
        self._sign_items = {}
        self._last_bulk = []
        self._page_cache.clear()
        self._close_docs()
        self._files = []
        self._cur_file = self._cur_page = 0
        for lst in (self.file_list, self.page_list):
            lst.blockSignals(True); lst.clear(); lst.blockSignals(False)
        self._update_buttons()

    def _doc(self, fi: int):
        d = self._docs.get(fi)
        if d is None:
            d = self._docs[fi] = engine.open_pdf(self._files[fi].pdf)
        return d

    def _cur_pages(self) -> List[engine.PageRef]:
        return self._files[self._cur_file].pages if self._cur_file < len(self._files) else []

    def _render_pixmap(self, fi: int, pno: int, zoom: float) -> QPixmap:
        w, h, stride, samples = engine.render_page(self._doc(fi), pno, zoom)
        # samples 를 QImage 가 복사해 갖도록 copy() — 원본 bytes 는 곧 사라진다
        return QPixmap.fromImage(QImage(samples, w, h, stride, QImage.Format_RGB888).copy())

    def _page_pixmap(self, fi: int, pno: int) -> QPixmap:
        key = (fi, pno)
        pm = self._page_cache.get(key)
        if pm is None:
            pm = self._render_pixmap(fi, pno, RENDER_ZOOM)
            self._page_cache[key] = pm
            while len(self._page_cache) > PAGE_CACHE_MAX:
                self._page_cache.popitem(last=False)
        else:
            self._page_cache.move_to_end(key)
        return pm

    # ── 목록 ────────────────────────────────────────────────
    def _file_label(self, fi: int) -> str:
        f = self._files[fi]
        if f.error:
            return f"⚠ {f.name} — {f.error}"
        if not f.pages:
            return f.name
        n = sum(len(self._sign_items.get((fi, p), [])) for p in range(len(f.pages)))
        return f"{f.name}   (서명 {n})" if n else f.name

    @staticmethod
    def _page_title(ref: engine.PageRef) -> str:
        return ref.sheet if ref.pages_in_sheet <= 1 else f"{ref.sheet} ({ref.index_in_sheet + 1}/{ref.pages_in_sheet})"

    def _page_label(self, fi: int, pno: int) -> str:
        n = len(self._sign_items.get((fi, pno), []))
        title = self._page_title(self._files[fi].pages[pno])
        return f"{pno + 1}. {title}" + (f"\n✓ 서명 {n}" if n else "\n서명 없음")

    def _fill_file_list(self) -> None:
        self.file_list.blockSignals(True)
        self.file_list.clear()
        for fi, f in enumerate(self._files):
            it = QListWidgetItem(self._file_label(fi))
            it.setData(Qt.UserRole, f.path)
            it.setToolTip(f.path)
            self.file_list.addItem(it)
        self.file_list.blockSignals(False)

    def _fill_page_list(self) -> None:
        self.page_list.blockSignals(True)
        self.page_list.clear()
        fi = self._cur_file
        for pno in range(len(self._cur_pages())):
            it = QListWidgetItem(QIcon(self._render_pixmap(fi, pno, THUMB_ZOOM)), self._page_label(fi, pno))
            self.page_list.addItem(it)
        self.page_list.blockSignals(False)

    def _refresh_marks(self) -> None:
        """서명 개수 표시(파일 목록·쪽 목록)와 버튼 상태를 지금 상태에 맞춘다."""
        for fi in range(min(self.file_list.count(), len(self._files))):
            self.file_list.item(fi).setText(self._file_label(fi))
        for pno in range(min(self.page_list.count(), len(self._cur_pages()))):
            self.page_list.item(pno).setText(self._page_label(self._cur_file, pno))
        self._update_buttons()

    # ── 보기 ────────────────────────────────────────────────
    def _on_select_file(self, row: int) -> None:
        if row < 0 or row >= len(self._files): return
        self._cur_file = row; self._cur_page = 0
        self._fill_page_list()
        if not self._cur_pages():
            self._detach_shown()
            self.lbl_status.setText("표시할 시트 없음(스킵)")
            self._update_buttons()
            return
        self._render()

    def _on_select_page(self, row: int) -> None:
        if row < 0 or row >= len(self._cur_pages()) or row == self._cur_page: return
        self._cur_page = row
        self._render()

    def _detach_shown(self) -> None:
        """지금 화면에 올라가 있는 쪽 그림과 서명을 씬에서 떼어낸다(서명은 보관)."""
        if self._shown_key is not None:
            for it in list(self._sign_items.get(self._shown_key, [])):
                try:
                    if it.scene() is self.scene: self.scene.removeItem(it)
                except RuntimeError: pass
        if self._bg_item is not None:
            try:
                if self._bg_item.scene() is self.scene: self.scene.removeItem(self._bg_item)
            except RuntimeError: pass
        self._bg_item = None
        self._shown_key = None

    def _render(self) -> None:
        pages = self._cur_pages()
        if not pages: return
        self._cur_page = max(0, min(self._cur_page, len(pages) - 1))
        fi, pno = self._cur_file, self._cur_page
        pm = self._page_pixmap(fi, pno)
        self._detach_shown()
        bg = QGraphicsPixmapItem(pm); bg.setZValue(0); bg.setAcceptedMouseButtons(Qt.NoButton)
        bg.setTransformationMode(Qt.SmoothTransformation)
        bg.setScale(1.0 / RENDER_ZOOM)                     # 씬 좌표 = PDF 포인트
        self.scene.addItem(bg); self._bg_item = bg; self._shown_key = (fi, pno)
        for it in list(self._sign_items.get(self._shown_key, [])):
            try: self.scene.addItem(it); it.setZValue(10)
            except RuntimeError: pass
        w_pt, h_pt = engine.page_size(self._doc(fi), pno)
        self.scene.setSceneRect(0, 0, w_pt, h_pt); self.view.resetTransform()
        vp_w = max(1, self.view.viewport().width())
        scale = vp_w / max(1.0, w_pt); self.view.scale(scale, scale); self.view.setFocus()
        self.view.verticalScrollBar().setValue(self.view.verticalScrollBar().minimum())
        self.page_list.blockSignals(True); self.page_list.setCurrentRow(pno); self.page_list.blockSignals(False)
        self.lbl_status.setText(
            f"파일 {fi + 1}/{len(self._files)} / 쪽 {pno + 1}/{len(pages)} · {self._page_title(pages[pno])}")
        self._update_buttons()

    def _step_file(self, step: int) -> int:
        """step 방향으로 쪽이 있는 다음 파일의 순번. 없으면 -1."""
        fi = self._cur_file + step
        while 0 <= fi < len(self._files):
            if self._files[fi].pages:
                return fi
            fi += step
        return -1

    def _next_page(self) -> None:
        if not self._cur_pages(): return
        if self._cur_page + 1 < len(self._cur_pages()):
            self._cur_page += 1; self._render()
        else:
            fi = self._step_file(+1)
            if fi >= 0: self.file_list.setCurrentRow(fi)

    def _prev_page(self) -> None:
        if not self._cur_pages(): return
        if self._cur_page - 1 >= 0:
            self._cur_page -= 1; self._render()
        else:
            fi = self._step_file(-1)
            if fi >= 0:
                self.file_list.setCurrentRow(fi)
                self._cur_page = max(0, len(self._cur_pages()) - 1); self._render()

    # ── 서명 ────────────────────────────────────────────────
    def _confirm_password(self) -> bool:
        dlg = PasswordDialog(self, self._code, prefill=self._last_password)
        if dlg.exec() != QDialog.Accepted or not dlg.verified:
            return False
        self._last_password = self._code   # 다음 서명부터는 미리 채워서 뜬다
        return True

    def _make_item(self, sign_idx: int, pno: int, x: float, y: float) -> SignatureItem:
        item = SignatureItem(self._signs[sign_idx], pno, on_delete=self._remove_sign)
        item.sign_idx = sign_idx
        item.setTransformationMode(Qt.SmoothTransformation)
        item.setScale(1.0 / RENDER_ZOOM)
        item.setZValue(10)
        item.setPos(QPointF(x, y))
        return item

    def place_sign(self, sign_idx: int, x: float, y: float) -> None:
        """지금 보고 있는 쪽의 (x, y)pt 에 서명을 놓는다. (x, y) 는 서명의 왼쪽 위."""
        key = (self._cur_file, self._cur_page)
        item = self._make_item(sign_idx, self._cur_page, x, y)
        self.scene.addItem(item)
        self._sign_items.setdefault(key, []).append(item)
        self.scene.update()
        self._refresh_marks()

    def _add_sign(self, scene_pos: QPointF) -> None:
        if not self._signs:
            QMessageBox.information(self, "안내", "서명 Load 후 서명 이미지가 필요합니다."); return
        if not self._cur_pages(): return
        if not self._confirm_password(): return
        idx = min(1 if (QApplication.keyboardModifiers() & Qt.ShiftModifier) else 0, len(self._signs) - 1)
        self.place_sign(idx, scene_pos.x() - engine.SIGN_W_PT / 2, scene_pos.y() - engine.SIGN_H_PT / 2)

    def _remove_sign(self, item) -> None:
        """서명 우클릭 삭제 콜백 — 보관 목록에서도 뺀다.

        빼주지 않으면 씬에서만 사라지고 목록에는 남아, 쪽을 다시 그릴 때
        되살아나고 PDF 에도 그대로 찍힌다.
        """
        for key, lst in list(self._sign_items.items()):
            if item in lst:
                lst.remove(item)
                if not lst:
                    self._sign_items.pop(key, None)
                break
        if item in self._last_bulk:
            self._last_bulk.remove(item)
        self._refresh_marks()

    def placements(self) -> Dict[Tuple[int, int], List[engine.Placement]]:
        """화면에 놓인 서명을 엔진이 쓰는 값으로 떠 낸다(GUI 스레드에서만 호출)."""
        out: Dict[Tuple[int, int], List[engine.Placement]] = {}
        for key, items in self._sign_items.items():
            got = []
            for it in items:
                try:
                    got.append(engine.Placement(int(getattr(it, "sign_idx", 0)),
                                                float(it.pos().x()), float(it.pos().y())))
                except RuntimeError:
                    continue
            if got:
                out[key] = got
        return out

    def _apply_to_others(self) -> None:
        if not self._signs:
            QMessageBox.information(self, "안내", "서명 Load 후 서명 이미지가 필요합니다."); return
        added = engine.copy_placements(self._files, self.placements(), self._cur_file)
        if not added:
            QMessageBox.information(self, "안내", "적용할 곳이 없습니다.\n"
                                    "(다른 파일에 같은 시트가 없거나, 이미 서명이 있습니다)")
            return
        if not self._confirm_password(): return
        n = self.apply_placements(added)
        self.lbl_status.setText(f"{len({k[0] for k in added})}개 파일 · {len(added)}쪽에 서명 {n}개를 놓았습니다")

    def apply_placements(self, added: Dict[Tuple[int, int], List[engine.Placement]]) -> int:
        """주어진 위치에 서명을 놓는다(일괄 적용). 반환: 놓은 서명 수."""
        self._last_bulk = []
        for (fi, pno), pls in added.items():
            for pl in pls:
                item = self._make_item(min(pl.sign, len(self._signs) - 1), pno, pl.x, pl.y)
                self._sign_items.setdefault((fi, pno), []).append(item)
                self._last_bulk.append(item)
                if (fi, pno) == self._shown_key:
                    self.scene.addItem(item)
        self._refresh_marks()
        return len(self._last_bulk)

    def _undo_bulk(self) -> None:
        items, self._last_bulk = self._last_bulk, []
        for item in items:
            for key, lst in list(self._sign_items.items()):
                if item in lst:
                    lst.remove(item)
                    if not lst:
                        self._sign_items.pop(key, None)
            try:
                if item.scene() is self.scene: self.scene.removeItem(item)
            except RuntimeError: pass
        self.lbl_status.setText(f"일괄 적용한 서명 {len(items)}개를 취소했습니다")
        self._refresh_marks()

    # ── ③ PDF 저장 ──────────────────────────────────────────
    def _save_pdf(self) -> None:
        if not any(f.pages for f in self._files):
            QMessageBox.information(self, "안내", "먼저 엑셀을 LOAD 하세요."); return
        if self._pdf_thread and self._pdf_thread.isRunning():
            QMessageBox.information(self, "안내", "이미 PDF 저장 중입니다."); return
        unsigned = [f.name for fi, f in enumerate(self._files) if f.pages
                    and not any(self._sign_items.get((fi, p)) for p in range(len(f.pages)))]
        if unsigned and self._sign_items:
            names = "\n".join("  • " + n for n in unsigned[:8]) + ("\n  …" if len(unsigned) > 8 else "")
            if QMessageBox.question(
                    self, "확인", f"서명이 하나도 없는 파일이 {len(unsigned)}개 있습니다.\n{names}\n\n그대로 저장할까요?"
            ) != QMessageBox.Yes:
                return
        folder_name = os.path.basename(self._base_folder.rstrip("\\/"))
        out = unique_path(os.path.join(self._base_folder, f"대외비_{folder_name}.pdf"))
        self.save_to(out)

    def save_to(self, out: str) -> None:
        total = sum(len(f.pages) for f in self._files)
        # 취소 불가 대기 다이얼로그 — 도중에 끊으면 반쯤 쓰인 PDF 처리가 새로 필요하다.
        self._pdf_progress = QProgressDialog("제출용 견적(PDF) 준비 중…", None, 0, total, self)
        self._pdf_progress.setWindowTitle("제출용 견적(PDF) Print")
        self._pdf_progress.setWindowModality(Qt.WindowModal)
        self._pdf_progress.setMinimumDuration(0)
        self._pdf_progress.setValue(0)

        self._pdf_thread = _PdfBuildThread(list(self._files), self.placements(),
                                           list(self._sign_bytes), out, self)
        self._pdf_thread.progress.connect(self._on_pdf_progress)
        self._pdf_thread.done.connect(lambda result: self._on_pdf_done(result, out))
        self._pdf_thread.start()
        self._update_buttons()

    def _on_pdf_progress(self, done: int, total: int) -> None:
        if self._pdf_progress is None:
            return
        self._pdf_progress.setValue(done)
        if self._pdf_progress is None:      # setValue(max) 재진입 대비 (_on_load_progress 와 같은 이유)
            return
        self._pdf_progress.setLabelText(f"제출용 견적(PDF) 저장 중 ({done}/{total})")

    def _on_pdf_done(self, result, out: str) -> None:
        if self._pdf_progress:
            self._pdf_progress.close()
            self._pdf_progress = None
        if self._pdf_thread is not None:
            self._pdf_thread.wait(2000)
        self._update_buttons()
        if isinstance(result, Exception):
            logger.error("PDF 저장 실패", exc_info=result)
            tb = "".join(traceback.format_exception(type(result), result, result.__traceback__))
            user_msg, hint = _friendly_error_msg(result)
            _ScrollableErrorDialog(self, tb, user_msg=user_msg, hint=hint).exec()
            return
        self.lbl_status.setText("제출용 견적(PDF) Print 완료")
        QMessageBox.information(self, "완료", f"저장 완료:\n{out}")
