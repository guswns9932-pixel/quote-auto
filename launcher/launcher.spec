# launcher.spec
# =============
# PyInstaller 빌드 명세 파일.
# EXE 재빌드가 필요한 경우에만 사용:
#   cd launcher/
#   pyinstaller launcher.spec
#
# 포함 대상: Python 런타임 + 무거운 의존성 (PySide6, openpyxl, fitz, win32com)
# 미포함:    앱 코드 (main.py 등) → App/<name>/*.pyz 로 분리 배포
#
# 배포 폴더 구조:
#   dist/QuoteAuto/
#   ├── QuoteAuto.exe    ← 이 파일
#   └── App/
#       └── 견적자동화/
#           ├── manifest.json
#           └── app_v1.0.1.pyz

from pathlib import Path
import os
from PyInstaller.utils.hooks import collect_all, collect_submodules, copy_metadata

BASE = Path(SPEC).parent  # launcher/ 폴더
ONEFILE = os.environ.get("QUOTEAUTO_ONEFILE", "0") == "1"
EXE_NAME = os.environ.get("QUOTEAUTO_EXE_NAME", "QuoteAuto")

# collect_all: 서브모듈 + 데이터파일 + 바이너리 전체 수집
# hiddenimports=["openpyxl"] 만으로는 서브모듈이 누락될 수 있음
openpyxl_d, openpyxl_b, openpyxl_h = collect_all("openpyxl")
fitz_d,     fitz_b,     fitz_h     = collect_all("fitz")
openpyxl_submods = collect_submodules("openpyxl")
openpyxl_meta = copy_metadata("openpyxl")

a = Analysis(
    [str(BASE / "launcher.py")],
    pathex=[str(BASE.parent)],
    binaries=[
        *openpyxl_b,
        *fitz_b,
    ],
    datas=[
        # App 폴더 전체를 배포물에 포함
        (str(BASE / "App"), "App"),
        *openpyxl_d,
        *fitz_d,
        *openpyxl_meta,
    ],
    hiddenimports=[
        "PySide6.QtCore",
        "PySide6.QtGui",
        "PySide6.QtWidgets",
        *openpyxl_h,
        *openpyxl_submods,
        *fitz_h,
        "pythoncom",
        "pywintypes",
        "win32api",
        "win32com",
        "win32com.client",
        "win32com.server",
        "win32con",
    ],
    hookspath=[],
    runtime_hooks=[],
    excludes=[
        # 앱 코드 제외 (pyz에서 로드)
        "main", "pages", "core", "excel_io", "widgets",
    ],
    noarchive=False,
    optimize=1,
)

pyz = PYZ(a.pure)

if ONEFILE:
    exe = EXE(
        pyz,
        a.scripts,
        a.binaries,
        a.datas,
        [],
        exclude_binaries=False,
        name=EXE_NAME,
        debug=False,
        strip=False,
        upx=True,
        console=False,
        # icon="icon.ico",
    )
else:
    exe = EXE(
        pyz,
        a.scripts,
        [],
        exclude_binaries=True,
        name=EXE_NAME,
        debug=False,
        strip=False,
        upx=True,
        console=False,
        # icon="icon.ico",
    )

    coll = COLLECT(
        exe,
        a.binaries,
        a.datas,
        strip=False,
        upx=True,
        upx_exclude=[],
        name=EXE_NAME,
    )
