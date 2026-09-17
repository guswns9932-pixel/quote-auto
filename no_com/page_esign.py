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
from core import unique_path
from widgets import PdfView, SignatureItem, PasswordDialog, tint_button
from page_common import _friendly_error_msg, _natural_key, _ScrollableErrorDialog

logger = logging.getLogger("QuoteApp")


class _ImageCache:
    """캡처한 PNG 를 QImage 로 캐싱한다.

    페이지 이동(_render)과 PDF 빌드(_PdfBuildThread)가 같은 이미지를 각자
    다시 디코드하던 것을 없애기 위한 공유 캐시. PDF 빌드는 백그라운드
    스레드에서 이 캐시를 읽고(캐시 미스 시) 채워 넣으므로 락으로 보호한다.
    QImage 자체는 Qt 문서상 어느 스레드에서 만들고 다뤄도 안전하다
    (QPixmap 과 달리 GUI 스레드 전용이 아니다).

    키는 보통 캡처 스레드가 메모리에 담아 둔 PNG bytes 의 키(mem://…)다.
    blobs 에 없는 키는 파일 경로로 보고 디스크에서 읽는다.
    """

    def __init__(self) -> None:
        self._data: dict = {}
        self._blobs: dict = {}
        self._lock = threading.Lock()

    def set_blobs(self, blobs: dict) -> None:
        with self._lock:
            self._blobs = blobs or {}

    def get(self, key: str) -> QImage:
        with self._lock:
            img = self._data.get(key)
            blob = None if img is not None else self._blobs.get(key)
        if img is not None:
            return img
        if blob is not None:
            img = QImage()
            img.loadFromData(QByteArray(blob), "PNG")
        else:
            img = QImage(key)
        with self._lock:
            self._data.setdefault(key, img)
            return self._data[key]

    def clear(self) -> None:
        with self._lock:
            self._data.clear()
            self._blobs = {}


class _ExcelWorkerThread(QThread):
    """전자서명 페이지가 살아 있는 동안 Excel COM 세션을 쥐고 있는 워커.

    페이지가 열리자마자 백그라운드로 Excel 을 띄워 두고, 사용자가 엑셀을
    고르면 이미 준비된 세션으로 바로 캡처를 시작한다. 예전에는 "엑셀 Load"
    를 누른 뒤에야 DispatchEx 를 했고, 그 기동에만 실측 1.1~1.2초가
    진행 다이얼로그 앞에 그대로 얹혀 있었다.

    [왜 스레드를 나누지 않는가]
    COM 객체는 만들어진 아파트먼트(=스레드)에 묶여 있어서, 다른 스레드에서
    쓰려면 인터페이스를 마샬링해야 한다. "미리 띄우는 스레드"와 "캡처하는
    스레드"를 따로 두면 그 마샬링이 필요해지고, 실패하면 보이지 않는
    EXCEL.EXE 가 남는 골치 아픈 경로가 생긴다. 그래서 이 스레드 하나가
    세션 생성 → 캡처 → 종료까지 전부 맡고, 작업이 없을 때는 이벤트에서
    잠들어 있는다.

    캡처는 CopyPicture 로 한다(ExportAsFixedFormat/PrintOut 미사용 →
    RenameFile 없음).
    """
    ready          = Signal(bool)                      # COM 세션 준비 완료(성공 여부)
    progress       = Signal(int, int, str)             # (완료 파일수, 전체 파일수, 현재파일명)
    sheet_progress = Signal(int, int, str, int, int, str)
    # (파일idx, 파일전체, 파일명, 완료시트수, 전체시트수, 시트명) — 파일 안에서의 세부 진행
    job_done       = Signal(object, object)            # (sheet_pngs, png_blobs)

    def __init__(self, parent=None) -> None:
        super().__init__(parent)
        self._job    : Optional[tuple] = None
        self._wake   = threading.Event()
        self._lock   = threading.Lock()
        self._stop   = False
        self._cancel = False

    def submit(self, paths: List[str], tmp_dir: str) -> None:
        """캡처할 파일 목록을 워커에 넘긴다(GUI 스레드에서 호출)."""
        with self._lock:
            self._job = (list(paths), tmp_dir)
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
        import excel_io
        com_ctx = excel_io.ExcelCOM()
        xl_app = None
        try:
            com_ctx.__enter__()
            xl_app = com_ctx.app
        except Exception as e:
            logger.error("Excel COM 초기화 실패: %s", e, exc_info=True)
        # 실패해도 계속 간다 — excel_capture_sheets_to_pngs 가 xl_app=None 이면
        # 파일마다 자체 ExcelCOM 을 쓰는 폴백 경로를 갖고 있다(느릴 뿐 동작함).
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
                self._capture_all(excel_io, xl_app, *job)
        finally:
            try:
                com_ctx.__exit__(None, None, None)
            except Exception:
                pass

    def _capture_all(self, excel_io, xl_app, paths: List[str], tmp_dir: str) -> None:
        sheet_pngs: List[List[str]] = []
        # 캡처한 PNG 를 디스크가 아니라 여기에 bytes 로 담는다. 임시 폴더를
        # 쓰던 시절엔 네트워크 드라이브에 썼다가 곧바로 다시 읽었는데,
        # 이미지는 이미 메모리에 있으므로 그 왕복이 통째로 불필요했다.
        # PNG 1장이 30~100KB 수준이라 수백 장이어도 수십 MB 다.
        png_blobs: dict = {}
        total = len(paths)
        try:
            for i, xlsx in enumerate(paths):
                if self._cancel:
                    break
                fname = os.path.basename(xlsx)
                self.progress.emit(i, total, fname)

                def _sheet_cb(done, sheet_total, sheet_name, _i=i, _fname=fname):
                    self.sheet_progress.emit(_i, total, _fname, done, sheet_total, sheet_name)

                try:
                    pngs = excel_io.excel_capture_sheets_to_pngs(
                        xlsx, tmp_dir, i + 1, xl_app,
                        progress_cb=_sheet_cb,
                        should_cancel=lambda: self._cancel,
                        blob_sink=png_blobs)
                    sheet_pngs.append(pngs)
                except Exception as e:
                    logger.error("시트 캡처 실패 (%s): %s", xlsx, e, exc_info=True)
                    sheet_pngs.append([])
        finally:
            self.progress.emit(len(sheet_pngs), total, "완료")
            self.job_done.emit(sheet_pngs, png_blobs)


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
        self._worker           : Optional[_ExcelWorkerThread] = None
        self._load_progress    : Optional[QProgressDialog]   = None
        self._pdf_thread        : Optional["_PdfBuildThread"] = None
        self._pdf_progress      : Optional[QProgressDialog]   = None
        self._tmp_dir          : Optional[str]               = None
        self._com_init_timer   : Optional[QTimer]            = None
        self._com_init_ok      : bool                        = False
        self._build_ui()
        # 페이지가 뜨자마자 Excel 을 백그라운드로 띄워 둔다. 사용자가 서명을
        # 고르고 엑셀을 선택하는 동안 기동이 끝나므로, "엑셀 Load" 앞에
        # 붙던 1초 남짓의 COM 기동이 체감에서 사라진다.
        self._start_worker()

    # ── Excel 워커 (COM 세션 예열) ─────────────────────────────
    def _start_worker(self) -> None:
        """Excel COM 세션을 쥘 워커를 띄운다. 이미 살아 있으면 아무것도 안 한다."""
        if self._worker is not None and self._worker.isRunning():
            return
        self._com_init_ok = False
        w = _ExcelWorkerThread(self)
        w.ready.connect(self._on_com_ready)
        w.progress.connect(self._on_load_progress)
        w.sheet_progress.connect(self._on_sheet_progress)
        w.job_done.connect(self._on_job_done)
        self._worker = w
        # DispatchEx 가 무한 블로킹하는 경우(Office 활성화 창 대기 등)를 잡는다.
        # 예열 중에는 사용자가 기다리고 있는 게 아니므로 대화창은 띄우지 않고,
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
            # 세션 없이도 파일마다 자체 COM 으로 동작은 한다 — 느릴 뿐이다.
            logger.warning("Excel COM 예열 실패 — 파일별 세션으로 동작합니다")

    def shutdown(self) -> None:
        """창이 닫히거나 페이지가 교체될 때 Excel 세션을 정리한다.

        이걸 빠뜨리면 보이지 않는 EXCEL.EXE 가 남아 파일 잠금을 쥔다.
        """
        self._stop_com_timer()
        w, self._worker = self._worker, None
        if w is None:
            return
        w.shutdown()
        if not w.wait(10_000):
            logger.warning("Excel 워커가 10초 내에 종료되지 않음 — 강제 종료")
            w.terminate()
            w.wait(3000)

    @staticmethod
    def _action_btn(label: str, slot) -> QPushButton:
        """견적서작성 페이지(page_quote._action_btn)와 크기·폰트를 통일한다."""
        btn = QPushButton(label); btn.setFixedHeight(52)
        f = btn.font(); f.setPointSize(10); f.setBold(True); btn.setFont(f)
        btn.clicked.connect(slot); return btn

    def _build_ui(self) -> None:
        outer = QVBoxLayout(self); outer.setContentsMargins(14,14,14,14); outer.setSpacing(10)

        # 1행: 버튼 — 작업 순서대로(서명 준비 → 문서 로드 → 저장)
        top = QHBoxLayout()
        self.btn_code   = self._action_btn("서명 Load",           self._load_code)
        self.btn_excel  = self._action_btn("원본 견적(엑셀) Load", self._load_excels)
        self.btn_save   = self._action_btn("제출용 견적(PDF) Print", self._save_pdf)
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

        # 2행: 서명 로드 상태 + 서명 미리보기 — 버튼 줄과 높이가 안 맞아
        # (미리보기가 실제 서명 크기 170x40) 별도 줄로 분리했다.
        code_row = QHBoxLayout()
        self.code_status_frame = QFrame()
        self.code_status_frame.setFrameShape(QFrame.StyledPanel)
        self._set_code_status_style(ok=None)
        csf = QHBoxLayout(self.code_status_frame)
        csf.setContentsMargins(8, 6, 8, 6)
        csf.setSpacing(8)
        self.lbl_code_state = QLabel("서명 미로드")
        self.lbl_sign_preview = QLabel()
        self.lbl_sign_preview.setFixedSize(self.SIGN_W, self.SIGN_H)
        self.lbl_sign_preview.setAlignment(Qt.AlignCenter)
        self.lbl_sign_preview.setStyleSheet(
            "background: white; border: 1px solid #CCC;")
        self.lbl_sign_preview.setToolTip(
            "서명1 미리보기 (더블클릭 시 배치)\nShift+더블클릭: 서명2 배치")
        csf.addWidget(self.lbl_code_state)
        csf.addWidget(self.lbl_sign_preview)
        code_row.addWidget(self.code_status_frame)
        code_row.addStretch(1)
        outer.addLayout(code_row)

        mid = QHBoxLayout()
        self.file_list = QListWidget(); self.file_list.setFixedWidth(360); mid.addWidget(self.file_list)
        self.scene = QGraphicsScene(self)
        self.view  = PdfView(self); self.view.setScene(self.scene); self.view.setAlignment(Qt.AlignCenter)
        self.view.setFocusPolicy(Qt.StrongFocus); self.view.setFocus(); mid.addWidget(self.view, 1)
        outer.addLayout(mid, 1)
        # 버튼 클릭 연결은 _action_btn() 생성 시 이미 처리했다(중복 연결 금지).
        self.file_list.currentRowChanged.connect(self._on_select_file)
        self.view.on_prev = self._prev_page; self.view.on_next = self._next_page
        self.view.on_double_click = self._add_sign

    def _set_code_status_style(self, ok: Optional[bool]) -> None:
        """서명 로드 상태 프레임 배경색. ok=None(미로드/회색) True(성공/연초록) False(실패/연빨강)."""
        color = {"None": "#F5F5F5", "True": "#E8F5E9", "False": "#FFEBEE"}[str(ok)]
        border = {"None": "#DDD", "True": "#A5D6A7", "False": "#EF9A9A"}[str(ok)]
        self.code_status_frame.setStyleSheet(
            f"QFrame {{ background: {color}; border: 1px solid {border}; border-radius: 4px; }}")

    def _build_sign_preview_pixmap(self, pm: QPixmap) -> QPixmap:
        """서명 Load 옆에 붙일 서명 미리보기 — 대표로 서명1 하나만,
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
            self._reset_code_status("⚠ 서명 읽기 실패", ok=False)
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
            return
        pm1 = QPixmap(p1).scaled(self.SIGN_W,self.SIGN_H,Qt.IgnoreAspectRatio,Qt.SmoothTransformation)
        pm2 = QPixmap(p2).scaled(self.SIGN_W,self.SIGN_H,Qt.IgnoreAspectRatio,Qt.SmoothTransformation)
        if pm1.isNull() or pm2.isNull():
            self._reset_code_status("⚠ 서명 이미지 로드 실패", ok=False)
            QMessageBox.critical(self,"오류","서명 이미지 로드 실패"); return
        self._signs = [pm1, pm2]

        # 서명 Load 버튼 옆에 상태 + 실제 찍힐 서명 이미지를 바로 보여준다.
        self.lbl_code_state.setText("✓ 서명 로드됨")
        self._set_code_status_style(ok=True)
        self.lbl_sign_preview.setPixmap(self._build_sign_preview_pixmap(pm1))
        self.lbl_sign_preview.setToolTip(
            f"서명1(대표): {os.path.basename(p1)}  (더블클릭)\n"
            f"서명2: {os.path.basename(p2)}  (Shift+더블클릭)")

        QMessageBox.information(self, "완료", "서명 Load 완료\n전자서명 ON")

    def _load_excels(self) -> None:
        import excel_io
        if not excel_io._ensure_com():
            QMessageBox.critical(self, "오류", "Excel COM이 없습니다.")
            return
        # 워커는 페이지가 열려 있는 동안 계속 살아 있으므로 isRunning() 으로는
        # 로딩 중인지 알 수 없다 — 진행 다이얼로그의 존재로 판단한다.
        if self._load_progress is not None:
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
        # 캡처 PNG 는 이제 메모리에만 담으므로 임시 폴더를 만들지 않는다.
        # (예전엔 로컬 디스크 쓰기가 막힌 환경이라 선택한 엑셀 옆에 만들었는데,
        #  그 네트워크 쓰기가 캡처 시간의 20% 였다.) 경로는 캡처 키 이름을
        # 만들 때만 쓰이고 실제로 생성되지는 않는다.
        tmp = os.path.join(base, "_esign_tmp_pdf")
        self._tmp_dir = None
        # 구버전이 남긴 임시 폴더가 있으면 정리한다(이제 다시 만들지 않으므로
        # 여기서 지워 주지 않으면 네트워크 드라이브에 계속 남는다).
        if os.path.isdir(tmp):
            try:
                shutil.rmtree(tmp)
            except Exception as e:
                logger.warning("구버전 tmp 폴더 삭제 실패: %s", e)

        self._load_progress = QProgressDialog("변환 준비 중...", "취소", 0, len(paths), self)
        self._load_progress.setWindowTitle("엑셀 → PDF 변환")
        self._load_progress.setWindowModality(Qt.WindowModal)
        self._load_progress.setMinimumDuration(0)
        self._load_progress.setValue(0)

        # 워커는 페이지가 열릴 때 이미 떠 있다. 예열이 실패해 스레드가 죽어
        # 있으면(초기화 타임아웃 등) 여기서 한 번 더 살려 본다.
        self._start_worker()
        self._load_progress.canceled.connect(self._worker.cancel)

        # 예열이 아직 안 끝났거나 첫 파일 열기가 무한 블로킹하는 경우를 잡는다.
        # 이번엔 "첫 progress" 를 기준으로 다시 감시한다.
        self._com_init_ok = False
        self._arm_com_timer()

        self.btn_code.setEnabled(False)
        self.btn_excel.setEnabled(False)
        self.btn_save.setEnabled(False)
        self._worker.submit(paths, tmp)

    def _on_load_progress(self, done: int, total: int, fname: str) -> None:
        if self._load_progress is None:
            return
        # 첫 progress 도착 → 타임아웃 타이머 해제
        if not self._com_init_ok:
            self._com_init_ok = True
            self._stop_com_timer()
        self._load_progress.setValue(done)
        # setValue(max) 가 QProgressDialog 자동 닫기(hide)를 트리거하고,
        # hide() 중 Qt 이벤트가 재진입해 _on_job_done 이 동기 실행될 수 있다.
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
        """Excel COM 이 30초 내에 응답하지 않으면 워커를 강제 종료한다.

        예열 중(사용자가 아직 아무것도 안 누른 상태)이라면 조용히 워커만
        정리한다 — 기다리는 사람이 없는데 대화창을 띄울 이유가 없고, 실제로
        엑셀을 로드할 때 다시 시도된다. 로드 중이었다면 안내까지 띄운다.
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

    def _on_job_done(self, sheet_pngs: list, png_blobs: dict) -> None:
        self._stop_com_timer()
        if self._load_progress:
            self._load_progress.close()
            self._load_progress = None
        self.btn_code.setEnabled(True)
        self.btn_excel.setEnabled(True)
        self.btn_save.setEnabled(True)
        self._sheet_pngs = sheet_pngs
        # 캡처 스레드가 메모리에 담아 둔 PNG bytes 를 캐시에 넘긴다.
        # 이후 _render 와 PDF 빌드는 전부 이 캐시만 보고 디스크를 건드리지 않는다.
        self._image_cache.set_blobs(png_blobs)
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
            QMessageBox.information(self, "안내", "서명 Load 후 서명 이미지가 필요합니다."); return
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
        self._pdf_progress = QProgressDialog("제출용 견적(PDF) 준비 중…", None, 0, len(plan), self)
        self._pdf_progress.setWindowTitle("제출용 견적(PDF) Print")
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
        self._pdf_progress.setLabelText(f"제출용 견적(PDF) 저장 중 ({done}/{total})")

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
        self.lbl_status.setText("제출용 견적(PDF) Print 완료")


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
            # 캡처 결과는 메모리 캐시(mem:// 키)에 있으므로 존재 확인은
            # 디코드 결과로 한다 — 경로가 아니라 키라 os.path.exists 는 못 쓴다.
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
