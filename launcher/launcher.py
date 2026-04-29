"""
launcher.py
===========
EXE로 컴파일되는 진입점.
- Python 런타임 + 무거운 의존성(PySide6, openpyxl, fitz, win32com)만 포함
- 실제 앱 코드는 manifest.json이 가리키는 .pyz에서 로드
- 코드 변경 시 EXE 재빌드 없이 .pyz + manifest.json 교체만으로 배포 완료
"""

from __future__ import annotations

import hashlib
import json
import os
import sys


def _base_dir() -> str:
    """EXE 실행 시 exe 위치, 스크립트 실행 시 launcher.py 위치."""
    if getattr(sys, "frozen", False):
        return os.path.dirname(sys.executable)
    return os.path.dirname(os.path.abspath(__file__))


def _verify_hash(path: str, expected: str) -> bool:
    if not expected:
        return True
    sha = hashlib.sha256(open(path, "rb").read()).hexdigest()
    return sha == expected


def main() -> None:
    base = _base_dir()
    manifest_path = os.path.join(base, "manifest.json")

    if not os.path.exists(manifest_path):
        _fatal("manifest.json 파일을 찾을 수 없습니다.\n\n" + manifest_path)

    try:
        manifest = json.loads(open(manifest_path, encoding="utf-8").read())
    except Exception as e:
        _fatal(f"manifest.json 읽기 실패:\n{e}")

    pyz_name = manifest.get("pyz")
    if not pyz_name:
        _fatal("manifest.json에 'pyz' 항목이 없습니다.")

    pyz_path = os.path.join(base, pyz_name)
    if not os.path.exists(pyz_path):
        _fatal(f"앱 파일을 찾을 수 없습니다:\n{pyz_path}")

    expected_hash = manifest.get("sha256", "")
    if expected_hash and not _verify_hash(pyz_path, expected_hash):
        _fatal(f"앱 파일 무결성 검증 실패.\n{pyz_name} 파일이 손상됐을 수 있습니다.")

    # .pyz를 sys.path 최상위에 삽입 → import 시 pyz 내부 코드 우선 탐색
    sys.path.insert(0, pyz_path)

    try:
        import main as app_main  # noqa: PLC0415
        app_main.main()
    except ImportError as e:
        _fatal(f"앱 코드 로드 실패:\n{e}\n\n{pyz_path}")


def _fatal(msg: str) -> None:
    """GUI 없이도 오류를 표시하고 종료."""
    try:
        from PySide6.QtWidgets import QApplication, QMessageBox
        _app = QApplication.instance() or QApplication(sys.argv)
        QMessageBox.critical(None, "실행 오류", msg)
    except Exception:
        print("ERROR:", msg, file=sys.stderr)
    sys.exit(1)


if __name__ == "__main__":
    main()
