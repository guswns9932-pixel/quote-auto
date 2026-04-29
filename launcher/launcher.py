"""
launcher.py
===========
모드 1 (기본 실행): GUI 앱 목록
  - 실행 위치의 App/ 폴더를 스캔
  - 발견된 프로그램을 리스트로 표시
  - 더블클릭 / 실행 버튼 → subprocess로 앱 기동

모드 2 (--run <folder>): 특정 앱 직접 실행
  - launcher.exe --run App/견적자동화
  - manifest.json 읽기 → pyz 로드 → entry.main() 실행
"""

from __future__ import annotations

import hashlib
import importlib
import json
import os
import subprocess
import sys


# ════════════════════════════════════════════════
# 앱 실행 모드  (subprocess에서 --run <folder> 로 호출)
# ════════════════════════════════════════════════

if "--run" in sys.argv:
    _idx = sys.argv.index("--run")
    _app_dir = sys.argv[_idx + 1]

    def _fatal_run(msg: str) -> None:
        print(f"[launcher] ERROR: {msg}", file=sys.stderr)
        try:
            from PySide6.QtWidgets import QApplication, QMessageBox
            _a = QApplication.instance() or QApplication(sys.argv[:1])
            QMessageBox.critical(None, "실행 오류", msg)
        except Exception:
            pass
        sys.exit(1)

    _manifest_path = os.path.join(_app_dir, "manifest.json")
    if not os.path.exists(_manifest_path):
        _fatal_run(f"manifest.json 없음:\n{_manifest_path}")

    try:
        _manifest = json.loads(open(_manifest_path, encoding="utf-8").read())
    except Exception as e:
        _fatal_run(f"manifest.json 읽기 실패:\n{e}")

    _pyz_name = _manifest.get("pyz", "")
    _pyz_path = os.path.join(_app_dir, _pyz_name)
    if not os.path.exists(_pyz_path):
        _fatal_run(f"앱 파일 없음:\n{_pyz_path}")

    _expected = _manifest.get("sha256", "")
    if _expected:
        _actual = hashlib.sha256(open(_pyz_path, "rb").read()).hexdigest()
        if _actual != _expected:
            _fatal_run(f"파일 무결성 오류:\n{_pyz_name}")

    sys.path.insert(0, _pyz_path)
    _entry = _manifest.get("entry", "main")

    try:
        _mod = importlib.import_module(_entry)
        _mod.main()
    except Exception as e:
        _fatal_run(f"앱 실행 실패:\n{e}")

    sys.exit(0)


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
        })
    return result


def _launch_app(folder: str) -> None:
    """subprocess로 앱 실행 (런처 자신을 --run 모드로 호출)."""
    exe = sys.executable
    script = os.path.abspath(__file__)
    if getattr(sys, "frozen", False):
        cmd = [exe, "--run", folder]
    else:
        cmd = [exe, script, "--run", folder]
    subprocess.Popen(cmd)


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
        w = option.rect.width() - self.PADDING * 2

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
        app = items[0].data(Qt.UserRole)
        try:
            _launch_app(app["folder"])
            self.lbl_status.setText(f"'{app['name']}' 실행 중...")
        except Exception as e:
            QMessageBox.critical(self, "실행 오류", str(e))


# ── 진입점 ────────────────────────────────────────

def main() -> None:
    app = QApplication(sys.argv)
    f = app.font(); f.setFamily("Malgun Gothic"); app.setFont(f)
    w = LauncherWindow()
    w.show()
    sys.exit(app.exec())


if __name__ == "__main__":
    main()
