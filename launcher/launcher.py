"""
launcher.py
===========
모드 1 (기본 실행 / GUI): App/ 폴더를 스캔해 프로그램 목록 표시.
  더블클릭 / 실행 버튼 → 같은 프로세스 내에서 pyz 앱을 직접 로드·실행.
  앱 종료 후 런처 창이 다시 나타난다.

※ subprocess를 쓰지 않는 이유
  EXE 내부에서 subprocess로 같은 EXE를 재실행하면 PyInstaller의
  모듈 탐색 경로(sys.path)가 재구성되면서 openpyxl 등 의존성이
  사라지는 문제가 발생한다.  같은 프로세스에서 직접 import하면
  이미 초기화된 Python 환경을 그대로 공유하므로 이 문제가 없다.
"""

from __future__ import annotations

import hashlib
import importlib
import json
import os
import sys

# ────────────────────────────────────────────────────────────────────
# PyInstaller 번들링 힌트
# launcher.py 자체는 아래 패키지를 직접 사용하지 않지만,
# 런타임에 pyz 앱(pages.py 등)이 import하므로 EXE에 반드시 포함해야 함.
# 이 import 문이 있어야 PyInstaller 정적 분석기가 패키지를 감지해
# EXE 빌드 시 자동으로 번들링한다.
# ────────────────────────────────────────────────────────────────────
try:
    import openpyxl            # noqa: F401
    import openpyxl.styles     # noqa: F401
    import openpyxl.utils      # noqa: F401
    import openpyxl.reader     # noqa: F401
    import openpyxl.writer     # noqa: F401
    import fitz                # noqa: F401
    import pythoncom           # noqa: F401
    import win32com.client     # noqa: F401
except ImportError:
    pass


# ════════════════════════════════════════════════
# GUI 런처 모드  (기본)
# ════════════════════════════════════════════════

from PySide6.QtCore import Qt, QSize
from PySide6.QtGui import QColor, QFont
from PySide6.QtWidgets import (
    QApplication, QFrame, QHBoxLayout, QLabel,
    QListWidget, QListWidgetItem, QMainWindow,
    QMessageBox, QPushButton, QSizePolicy,
    QStyle, QStyledItemDelegate,
    QVBoxLayout, QWidget,
)


# ── 유틸 ─────────────────────────────────────────

def _base_dir() -> str:
    if getattr(sys, "frozen", False):
        return os.path.dirname(sys.executable)
    return os.path.dirname(os.path.abspath(__file__))


def _scan_apps(base: str) -> list[dict]:
    """App/ 하위 폴더를 스캔해 앱 정보 목록 반환."""
    app_root = os.path.join(base, "App")
    result = []
    if not os.path.isdir(app_root):
        return result
    for name in sorted(os.listdir(app_root)):
        folder = os.path.join(app_root, name)
        if not os.path.isdir(folder):
            continue
        manifest_path = os.path.join(folder, "manifest.json")
        if not os.path.exists(manifest_path):
            continue
        try:
            manifest = json.loads(open(manifest_path, encoding="utf-8").read())
        except Exception:
            continue
        pyz = os.path.join(folder, manifest.get("pyz", ""))
        if not os.path.exists(pyz):
            continue
        result.append({
            "folder_name": name,
            "name":        manifest.get("name",        name),
            "description": manifest.get("description", ""),
            "version":     manifest.get("version",     "-"),
            "built_at":    manifest.get("built_at",    ""),
            "folder":      folder,
            "manifest":    manifest,
        })
    return result


def _load_and_run(app_info: dict) -> None:
    """
    pyz를 같은 프로세스에서 직접 로드·실행한다.
    subprocess 없이 실행하므로 PyInstaller 번들 의존성을 그대로 공유한다.
    """
    folder   = app_info["folder"]
    manifest = app_info["manifest"]

    pyz_name = manifest.get("pyz", "")
    pyz_path = os.path.join(folder, pyz_name)
    if not os.path.exists(pyz_path):
        raise RuntimeError(f"앱 파일 없음:\n{pyz_path}")

    expected = manifest.get("sha256", "")
    if expected:
        actual = hashlib.sha256(open(pyz_path, "rb").read()).hexdigest()
        if actual != expected:
            raise RuntimeError(f"파일 무결성 오류:\n{pyz_name}")

    entry = manifest.get("entry", "main")

    # 같은 pyz가 아직 sys.path에 없을 때만 추가
    if pyz_path not in sys.path:
        sys.path.insert(0, pyz_path)

    # 모듈 캐시 초기화 (앱 재실행 시 fresh import)
    for key in list(sys.modules.keys()):
        if key in ("main", "pages", "core", "excel_io", "widgets"):
            del sys.modules[key]

    mod = importlib.import_module(entry)
    mod.main()  # main.py 의 main() 호출 → 내부에서 이벤트 루프 관리


# ── 커스텀 아이템 델리게이트 ──────────────────────

class _AppDelegate(QStyledItemDelegate):
    """앱 이름(굵게) + 설명·버전·날짜(회색 작은 글씨) 2줄 렌더링."""

    LINE1_SIZE = 13
    LINE2_SIZE = 10
    PADDING    = 10

    def sizeHint(self, option, index) -> QSize:
        return QSize(option.rect.width(), 64)

    def paint(self, painter, option, index) -> None:
        app = index.data(Qt.UserRole)
        if not app:
            super().paint(painter, option, index)
            return

        painter.save()

        # 선택 배경
        if option.state & QStyle.StateFlag.State_Selected:
            painter.fillRect(option.rect, QColor("#BBDEFB"))
        elif index.row() % 2 == 1:
            painter.fillRect(option.rect, QColor("#F9F9F9"))

        x = option.rect.x() + self.PADDING
        y = option.rect.y()

        # 앱 이름
        f1 = QFont(); f1.setPointSize(self.LINE1_SIZE); f1.setBold(True)
        f1.setFamily("Malgun Gothic")
        painter.setFont(f1)
        painter.setPen(QColor("#1A1A1A"))
        painter.drawText(x, y + 24, app["name"])

        # 설명 + 버전 + 날짜
        parts = []
        if app["description"]:
            parts.append(app["description"])
        parts.append(f"v{app['version']}")
        if app["built_at"]:
            parts.append(app["built_at"])
        line2 = "   ".join(parts)

        f2 = QFont(); f2.setPointSize(self.LINE2_SIZE)
        f2.setFamily("Malgun Gothic")
        painter.setFont(f2)
        painter.setPen(QColor("#777777"))
        painter.drawText(x + 2, y + 46, line2)

        painter.restore()


# ── 메인 윈도우 ───────────────────────────────────

class LauncherWindow(QMainWindow):

    def __init__(self) -> None:
        super().__init__()
        self.setWindowTitle("앱 런처")
        self.setMinimumSize(560, 460)
        self._base = _base_dir()
        self._apps: list[dict] = []
        self._build_ui()
        self._refresh()

    def _build_ui(self) -> None:
        central = QWidget()
        self.setCentralWidget(central)
        root = QVBoxLayout(central)
        root.setContentsMargins(18, 18, 18, 18)
        root.setSpacing(12)

        # ── 헤더 ──
        header = QHBoxLayout()
        lbl_title = QLabel("앱 런처")
        f = lbl_title.font(); f.setPointSize(18); f.setBold(True); lbl_title.setFont(f)
        self.btn_refresh = QPushButton("새로고침")
        self.btn_refresh.setFixedHeight(32)
        self.btn_refresh.setStyleSheet(
            "QPushButton{background:#E8EAF6;border:1px solid #bbb;border-radius:4px;padding:0 10px;}"
            "QPushButton:hover{background:#C5CAE9;}"
        )
        self.btn_refresh.clicked.connect(self._refresh)
        header.addWidget(lbl_title)
        header.addStretch(1)
        header.addWidget(self.btn_refresh)
        root.addLayout(header)

        # ── 구분선 ──
        sep = QFrame(); sep.setFrameShape(QFrame.HLine); sep.setFrameShadow(QFrame.Sunken)
        root.addWidget(sep)

        # ── 앱 목록 ──
        self.list_widget = QListWidget()
        self.list_widget.setItemDelegate(_AppDelegate())
        self.list_widget.setSpacing(2)
        self.list_widget.setStyleSheet(
            "QListWidget{border:1px solid #ddd;border-radius:6px;}"
            "QListWidget::item{border-bottom:1px solid #eee;}"
        )
        self.list_widget.itemDoubleClicked.connect(self._on_launch)
        root.addWidget(self.list_widget, 1)

        # ── 하단 ──
        bottom = QHBoxLayout()
        self.lbl_status = QLabel("")
        self.lbl_status.setStyleSheet("color:#888;font-size:11px;")

        self.btn_run = QPushButton("▶   실행")
        self.btn_run.setMinimumHeight(44)
        self.btn_run.setMinimumWidth(130)
        f2 = self.btn_run.font(); f2.setPointSize(12); f2.setBold(True); self.btn_run.setFont(f2)
        self.btn_run.setStyleSheet(
            "QPushButton{background:#B3E5FC;border:1px solid #bbb;border-radius:6px;}"
            "QPushButton:hover{background:#81D4FA;}"
            "QPushButton:disabled{background:#EEE;color:#AAA;}"
        )
        self.btn_run.clicked.connect(self._on_launch)

        bottom.addWidget(self.lbl_status, 1)
        bottom.addWidget(self.btn_run)
        root.addLayout(bottom)

    def _refresh(self) -> None:
        self._apps = _scan_apps(self._base)
        self.list_widget.clear()
        for app in self._apps:
            item = QListWidgetItem()
            item.setData(Qt.UserRole, app)
            item.setSizeHint(QSize(0, 64))
            self.list_widget.addItem(item)
        n = len(self._apps)
        self.lbl_status.setText(f"App/ 폴더에서 {n}개 프로그램 발견" if n else "App/ 폴더에 프로그램 없음")

    def _on_launch(self, *_) -> None:
        items = self.list_widget.selectedItems()
        if not items:
            QMessageBox.information(self, "안내", "실행할 프로그램을 선택하세요.")
            return
        app_info = items[0].data(Qt.UserRole)

        try:
            # 런처 숨기고 앱 실행 (같은 프로세스 내 직접 로드)
            self.lbl_status.setText(f"'{app_info['name']}' 실행 중...")
            self.hide()

            qa = QApplication.instance()
            qa.setQuitOnLastWindowClosed(False)   # 앱 창 닫아도 런처 프로세스 유지
            try:
                _load_and_run(app_info)
            except SystemExit:
                pass
            finally:
                qa.setQuitOnLastWindowClosed(True)

        except Exception as e:
            QMessageBox.critical(None, "실행 오류", str(e))
        finally:
            self.lbl_status.setText("")
            self.show()
            self._refresh()


# ── 진입점 ────────────────────────────────────────

def main() -> None:
    app = QApplication(sys.argv)
    f = app.font(); f.setFamily("Malgun Gothic"); app.setFont(f)
    w = LauncherWindow()
    w.show()
    sys.exit(app.exec())


if __name__ == "__main__":
    main()
