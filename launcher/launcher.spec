# launcher.spec
# =============
# PyInstaller 빌드 명세 파일.
# EXE 재빌드가 필요한 경우에만 사용:
#   pyinstaller launcher.spec
#
# 포함 대상: Python 런타임 + 무거운 의존성
# 미포함:    앱 코드 (main.py 등) → .pyz로 분리 배포

import sys
from pathlib import Path

BASE = Path(SPEC).parent  # launcher/ 폴더

a = Analysis(
    [str(BASE / "launcher.py")],
    pathex=[str(BASE.parent)],   # quote-auto/ → 의존성 탐색용
    binaries=[],
    datas=[],
    hiddenimports=[
        # PySide6 필수 모듈
        "PySide6.QtCore",
        "PySide6.QtGui",
        "PySide6.QtWidgets",
        # 앱 의존성 (앱 코드는 pyz지만 런타임은 EXE에 포함)
        "openpyxl",
        "openpyxl.styles",
        "openpyxl.utils",
        "fitz",           # PyMuPDF
        "pythoncom",
        "win32com",
        "win32com.client",
    ],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[
        # 앱 코드 제외 (pyz에서 로드)
        "main",
        "pages",
        "core",
        "excel_io",
        "widgets",
    ],
    noarchive=False,
    optimize=1,
)

pyz = PYZ(a.pure)

exe = EXE(
    pyz,
    a.scripts,
    [],
    exclude_binaries=True,
    name="QuoteAuto",
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    console=False,          # 콘솔창 숨김
    disable_windowed_traceback=False,
    # icon="icon.ico",      # 아이콘 경로 (있으면 주석 해제)
)

coll = COLLECT(
    exe,
    a.binaries,
    a.datas,
    strip=False,
    upx=True,
    upx_exclude=[],
    name="QuoteAuto",
)
