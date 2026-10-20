"""
main.py
=======
런처 — 앱별 버튼만 가진 얇은 창.

각 앱(견적서작성·전자서명·키워드 검색기·CSP UPLOAD 자동화)은 별도 exe
(`<런처이름>_quote.exe` 등)로 떠서 자기에게 필요한 모듈만 임포트한다.
런처는 PySide6 와 업데이트 안내만 쓰므로 openpyxl·fitz·PIL·selenium 을 싣지
않고, 그래서 첫 실행이 빠르다. 개발 환경(비 frozen)에서는 run_*.py 를 직접 실행한다.
"""

from __future__ import annotations

import logging
import os
import subprocess
import sys
from typing import Optional

from PySide6.QtCore import Qt, QEvent, QRect
from PySide6.QtGui import QColor, QFont, QPainter, QPixmap
from PySide6.QtWidgets import (
    QApplication, QDialog, QHBoxLayout, QLabel, QMainWindow, QMessageBox,
    QPushButton, QScrollArea, QScrollBar, QSizePolicy, QSplashScreen,
    QTabBar, QTabWidget, QVBoxLayout, QWidget,
)

import app_settings
from app_boot import _setup_exception_hook, _setup_logging
from widgets import tint_button

logger = logging.getLogger("QuoteApp")


# ──────────────────────────────────────────────
# 앱 실행 (앱마다 별도 프로세스 / 별도 exe)
# ──────────────────────────────────────────────
# 키: 런처 버튼 → (라벨, exe 접미사, 개발 환경 스크립트)
# frozen 일 때 exe 는 런처 exe 와 같은 폴더의 "<런처이름>_<접미사>.exe" 이고,
# 없으면(단일 exe 빌드) 런처 exe 자신을 "--<접미사>" 인자로 재실행한다.
_APPS = {
    "quote":  ("견적서작성",        "quote",  "run_quote.py"),
    "esign":  ("전자서명",          "esign",  "run_esign.py"),
    "search": ("키워드 검색기",     "search", "run_search.py"),
    "csp":    ("CSP UPLOAD 자동화", "csp",    "run_csp.py"),
}

# subprocess.Popen 은 자식이 창을 띄울 때까지 기다리지 않고 즉시 반환한다.
# 창이 뜨기 전에 버튼을 여러 번 누르면 프로세스가 계속 늘어나므로, 마지막으로
# 띄운 프로세스가 아직 살아 있으면(poll() is None) 새로 띄우지 않는다.
_launched_procs: dict = {}


def _app_command(key: str) -> tuple:
    """(실행 명령, 존재해야 하는 파일 경로).

    frozen 이면 런처 옆의 "<런처이름>_<접미사>.exe"(앱별 exe 빌드)를 쓰고,
    없으면 자기 자신을 "--<접미사>" 로 재실행한다(단일 exe 빌드 — 해시 1개)."""
    _label, suffix, script = _APPS[key]
    if getattr(sys, "frozen", False):
        exe = os.path.abspath(sys.executable)
        stem = os.path.splitext(os.path.basename(exe))[0]
        sibling = os.path.join(os.path.dirname(exe), f"{stem}_{suffix}.exe")
        if os.path.exists(sibling):
            return [sibling], sibling
        return [exe, f"--{suffix}"], exe
    path = os.path.join(os.path.dirname(os.path.abspath(__file__)), script)
    return [sys.executable, path], path


def _launch_app(key: str, parent: Optional[QWidget] = None) -> None:
    label = _APPS[key][0]
    proc = _launched_procs.get(key)
    if proc is not None and proc.poll() is None:
        logger.info("%s가 이미 실행 중이라 재실행을 건너뜀", label)
        return
    cmd, must_exist = _app_command(key)
    if not os.path.exists(must_exist):
        logger.error("%s 실행 파일 없음: %s", label, must_exist)
        QMessageBox.critical(parent, f"{label} 오류",
                             f"실행 파일을 찾을 수 없습니다.\n{must_exist}")
        return
    try:
        _launched_procs[key] = subprocess.Popen(cmd)
    except Exception as e:
        logger.error("%s 실행 실패", label, exc_info=True)
        QMessageBox.critical(parent, f"{label} 오류", f"실행할 수 없습니다.\n{e}")


# ──────────────────────────────────────────────
# 업데이트 안내 — 굵직한 변경사항만, 실행할 때마다
# ──────────────────────────────────────────────
# id 앞 10자(YYYY-MM-DD)를 날짜로 그대로 쓴다. 자잘한 버그 수정·성능
# 개선·내부 리팩터링은 이 목록에 올리지 않는다 — 사용자가 체감할 만한
# 기능 추가/변경만 적는다.
#
# desc 는 한 문장, 60자 이내. 어떻게 구현했는지·조건이 무엇인지는 쓰지
# 않고 사용자가 뭘 할 수 있게 됐는지만 적는다. 한 번에 여러 개를 바꿨어도
# 나열하지 말고 가장 체감되는 것 하나로 줄인다.
# app: 어느 런처 버튼(앱)의 소식인지. 업데이트 안내 창에서 앱별 탭으로
# 나눠 보여주는 데 쓴다. "launcher"는 특정 앱 하나가 아니라 런처 자체
# 또는 여러 앱에 걸친 변경사항.
WHATS_NEW = [
    {
        "id": "2026-10-20-local-install",
        "app": "launcher",
        "title": "내 PC에서 빠르게 실행",
        "desc": "PC에 설치해 실행하고, 견적서는 원하는 클라우드 폴더에 저장할 수 있습니다.",
    },
    {
        "id": "2026-08-05-window-split",
        "app": "quote",
        "title": "창 분리 + 컬럼 확장",
        "desc": "견적서작성 창을 독립 창으로 분리하고 Rack/Maker/설비/5D 컬럼을 추가했습니다.",
    },
    {
        "id": "2026-08-26-submit-excel",
        "app": "quote",
        "title": "제출용 엑셀 생성 기능 추가",
        "desc": "갑지에서 갑지DATA 시트만 추출해 제출용 파일을 바로 만들 수 있습니다.",
    },
    {
        "id": "2026-09-01-autoload-template",
        "app": "quote",
        "title": "통합양식 자동 로드",
        "desc": "마지막에 쓴 통합양식·출력 폴더를 다음 실행 때 자동으로 불러옵니다.",
    },
    {
        "id": "2026-09-01-credit-fix",
        "app": "quote",
        "title": "Credit 계산 정확도 개선",
        "desc": "Pump/Rack Credit이 사양서·갑지 금액에 정확히 반영되도록 계산 방식을 바로잡았습니다.",
    },
    {
        "id": "2026-09-02-keyword-search",
        "app": "search",
        "title": "키워드 검색기 앱 추가",
        "desc": "런처에서 폴더 안 파일명·내용을 검색하는 키워드 검색기를 바로 실행할 수 있습니다.",
    },
    {
        "id": "2026-09-02-esign-speed",
        "app": "esign",
        "title": "전자서명 속도·편의성 개선",
        "desc": "PDF 저장 중에도 창이 멈추지 않고, 비밀번호는 한 번만 입력하면 됩니다.",
    },
    {
        "id": "2026-09-02-options-panel-redesign",
        "app": "quote",
        "title": "옵션 패널 2단 구조로 재정리",
        "desc": "Credit·보증기간·투자자·견적서 타입을 항목별로 나누고 적용 상태를 바로 보여줍니다.",
    },
    {
        "id": "2026-09-02-spec-hardcode-options",
        "app": "quote",
        "title": "사양서 입력 옵션 추가",
        "desc": "설비수량·실반입라인·Exh Size를 입력하면 국내 견적서에 자동으로 채워집니다.",
    },
    {
        "id": "2026-09-04-5d-object-gen",
        "app": "quote",
        "title": "5D Object 생성 버튼 추가",
        "desc": "STEP5에 담은 RACK 품목을 통합양식 코드매핑 시트에 바로 저장합니다.",
    },
    {
        "id": "2026-09-06-csp-upload",
        "app": "csp",
        "title": "CSP UPLOAD 자동화 앱 추가",
        "desc": "런처에서 CSP 주문접수 업로드 파일 생성기를 바로 실행할 수 있습니다.",
    },
    {
        "id": "2026-09-07-launcher-dedup",
        "app": "launcher",
        "title": "런처 버튼 중복 실행 방지",
        "desc": "런처 버튼을 여러 번 눌러도 창이 하나만 뜹니다.",
    },
    {
        "id": "2026-09-09-csp-cip-alarm",
        "app": "csp",
        "title": "AS-IS FSC 검토 알람",
        "desc": "옵션 조건이 CIP의 AS-IS 자재코드와 겹치면 검토 필요로 표시합니다.",
    },
    {
        "id": "2026-09-10-csp-fsc-recommend",
        "app": "csp",
        "title": "FSC 추천 기능 추가",
        "desc": "옵션 조건에 맞는 자재코드를 선택 창 맨 위에 추천하고, 규격으로 Q-code를 찾을 수 있습니다.",
    },
    {
        "id": "2026-09-11-csp-fsc-map-writeback",
        "app": "csp",
        "title": "FSC매핑 시트 자동 갱신",
        "desc": "엑셀 생성 시 새로 쓴 조합과 자재코드를 FSC매핑 시트에 자동으로 기록합니다.",
    },
    {
        "id": "2026-09-15-csp-load-speed",
        "app": "csp",
        "title": "통합양식 로드 속도 개선",
        "desc": "통합양식을 불러올 때 창이 멈추던 시간이 절반 이하로 줄었습니다.",
    },
    {
        "id": "2026-09-15-quote-output-slim",
        "app": "quote",
        "title": "견적서 파일 용량 감소",
        "desc": "견적서 파일에서 내부 단가표 등을 빼 용량이 1MB에서 80KB로 줄었습니다.",
    },
    {
        "id": "2026-09-16-esign-load-speed",
        "app": "esign",
        "title": "문서 로드 속도 개선",
        "desc": "엑셀을 불러와 미리보기가 뜨기까지 걸리는 시간이 절반으로 줄었습니다.",
    },
    {
        "id": "2026-09-16-pump-dup-fix",
        "app": "quote",
        "title": "PUMP 중복 담김 방지",
        "desc": "같은 PUMP 품목을 두 번 담아 견적금액이 배로 나가는 문제를 막았습니다.",
    },
    {
        "id": "2026-10-18-split-apps",
        "app": "launcher",
        "title": "앱별 분리로 빠른 실행",
        "desc": "각 앱이 따로 실행되어 첫 실행이 빨라졌습니다.",
    },
]


def _entry_date(entry: dict) -> str:
    """id 앞 10자를 날짜(YYYY-MM-DD)로 뽑아낸다."""
    return entry["id"][:10]


# 앱별 탭 표시 이름. 런처 버튼과 같은 이름을 써서 어느 버튼 소식인지
# 바로 알아볼 수 있게 한다. "launcher"는 버튼이 따로 없는 공용 탭.
_APP_TAB_LABELS = {
    "quote": "견적서작성",
    "esign": "전자서명",
    "search": "키워드 검색기",
    "csp": "CSP UPLOAD",
    "launcher": "런처",
}


class _WhatsNewDialog(QDialog):
    """굵직한 업데이트 안내 — 앱별 탭으로 나눠 보여준다.
    탭/스크롤바가 아닌 곳을 클릭하거나 X를 누르면 바로 닫힌다."""

    def __init__(self, entries: list, parent: Optional[QWidget] = None) -> None:
        super().__init__(parent)
        self.setWindowTitle("업데이트 소식")
        self.setFixedSize(460, 480)
        self.setWindowModality(Qt.ApplicationModal)

        # 앱별로 묶고, 각 그룹 안에서는 최신순(내림차순) 정렬.
        by_app: dict[str, list] = {}
        for e in entries:
            by_app.setdefault(e.get("app", "launcher"), []).append(e)
        for group in by_app.values():
            group.sort(key=_entry_date, reverse=True)

        # 탭 순서: 가장 최근에 업데이트된 앱이 맨 앞(기본 선택 탭)에 오게.
        app_keys = sorted(by_app.keys(), key=lambda a: _entry_date(by_app[a][0]),
                          reverse=True)
        latest_date = max(_entry_date(e) for e in entries) if entries else ""

        v = QVBoxLayout(self)
        v.setContentsMargins(20, 18, 20, 14)
        v.setSpacing(10)

        head = QLabel(f"새로워진 점  <span style='color:#888; font-weight:normal; "
                      f"font-size:11px;'>(최근 업데이트: {latest_date})</span>")
        head.setTextFormat(Qt.RichText)
        hf = head.font(); hf.setPointSize(13); hf.setBold(True); head.setFont(hf)
        v.addWidget(head)

        tabs = QTabWidget()
        v.addWidget(tabs, 1)
        for app_key in app_keys:
            tabs.addTab(self._build_tab(by_app[app_key]),
                       _APP_TAB_LABELS.get(app_key, app_key))

        hint = QLabel("탭을 눌러 다른 앱 소식도 볼 수 있습니다 · 그 외 클릭하면 닫힙니다")
        hint.setStyleSheet("color: #999; font-size: 10px;")
        hint.setAlignment(Qt.AlignRight)
        v.addWidget(hint)

        # 배경이든 라벨 위든 어디를 클릭해도 닫히게 한다 — 단, 탭 헤더나
        # 스크롤바를 클릭한 거라면 그 조작(탭 전환/스크롤)이 되게 둔다.
        # Qt 마우스 이벤트는 부모로 자동 전파되지 않으므로, 다이얼로그와
        # 모든 자식 위젯에 이벤트 필터를 직접 걸어 가로챈다.
        # (창의 X 버튼은 창 관리자 기본 동작으로 이미 닫힘 — 별도 처리 불필요)
        self.installEventFilter(self)
        for w in self.findChildren(QWidget):
            w.installEventFilter(self)

    @staticmethod
    def _build_tab(group_entries: list) -> QWidget:
        """앱 한 개 분량의 항목 목록을 스크롤 가능한 탭 페이지로 만든다."""
        content = QWidget()
        cv = QVBoxLayout(content)
        cv.setContentsMargins(4, 8, 4, 8)
        cv.setSpacing(12)
        for e in group_entries:
            item = QLabel(
                f"<b>• [{_entry_date(e)}] {e['title']}</b><br>"
                f"<span style='color:#555;'>{e['desc']}</span>"
            )
            item.setTextFormat(Qt.RichText)
            item.setWordWrap(True)
            cv.addWidget(item)
        cv.addStretch(1)

        scroll = QScrollArea()
        scroll.setWidget(content)
        scroll.setWidgetResizable(True)
        scroll.setFrameShape(QScrollArea.NoFrame)
        return scroll

    def eventFilter(self, obj, event) -> bool:
        if event.type() == QEvent.MouseButtonPress:
            if isinstance(obj, (QTabBar, QScrollBar)):
                return super().eventFilter(obj, event)
            self.accept()
            return True
        return super().eventFilter(obj, event)


def _show_whats_new(parent: QWidget) -> None:
    """굵직한 업데이트 안내. 실행할 때마다 전체 목록을 보여준다."""
    if not WHATS_NEW:
        return
    _WhatsNewDialog(WHATS_NEW, parent).exec()
    # 가장 최근에 본 업데이트 일자를 기록(참고용 — 노출 여부와는 무관).
    latest_id = max(e["id"] for e in WHATS_NEW)
    app_settings.set_str(app_settings.Key.LAST_SEEN_UPDATE, latest_id)


# ──────────────────────────────────────────────
# 스플래시 스크린
# ──────────────────────────────────────────────
def _make_splash_pix() -> QPixmap:
    pix = QPixmap(520, 200)
    pix.fill(QColor("#1565C0"))
    p = QPainter(pix)

    p.setPen(QColor("#FFFFFF"))
    f = QFont("Malgun Gothic", 20)
    f.setBold(True)
    p.setFont(f)
    p.drawText(QRect(0, 55, 520, 65), Qt.AlignCenter, "업무자동화 통합 시스템")

    p.setPen(QColor("#BBDEFB"))
    f2 = QFont("Malgun Gothic", 11)
    p.setFont(f2)
    p.drawText(QRect(0, 130, 520, 40), Qt.AlignCenter, "초기화 중…")

    p.end()
    return pix


# ──────────────────────────────────────────────
# 런처 윈도우 (메인)
# ──────────────────────────────────────────────
class MainWindow(QMainWindow):
    """앱별 버튼으로 각 앱을 별도 프로세스로 띄우는 런처."""

    def __init__(self) -> None:
        super().__init__()
        self.setWindowTitle("업무자동화 통합 시스템")
        self.setFixedSize(340, 354)
        app_settings.restore_geometry("MainWindow", self)

        central = QWidget()
        self.setCentralWidget(central)
        v = QVBoxLayout(central)
        v.setContentsMargins(30, 30, 30, 30)
        v.setSpacing(18)

        for key, color in [("quote", "#BBDEFB"), ("esign", "#C8E6C9"),
                           ("search", "#FFE0B2"), ("csp", "#D1C4E9")]:
            btn = QPushButton(_APPS[key][0])
            btn.setMinimumHeight(60)
            f = btn.font(); f.setPointSize(14); f.setBold(True); btn.setFont(f)
            tint_button(btn, color)
            v.addWidget(btn)
            btn.clicked.connect(lambda _=False, k=key: _launch_app(k, self))

    def closeEvent(self, event) -> None:
        # 각 앱은 독립 프로세스라 런처를 닫아도 계속 실행된다.
        app_settings.save_geometry("MainWindow", self)
        super().closeEvent(event)


# ──────────────────────────────────────────────
# 진입점
# ──────────────────────────────────────────────
def main() -> None:
    _setup_logging()
    _setup_exception_hook()
    logger.info("런처 시작")

    app = QApplication(sys.argv)
    f = app.font()
    f.setFamily("Malgun Gothic")
    app.setFont(f)

    # 스플래시: 창이 보이기 전에 즉시 표시
    splash = QSplashScreen(_make_splash_pix(), Qt.WindowStaysOnTopHint)
    splash.show()
    app.processEvents()

    window = MainWindow()
    splash.finish(window)
    window.show()
    _show_whats_new(window)
    sys.exit(app.exec())


if __name__ == "__main__":
    main()
