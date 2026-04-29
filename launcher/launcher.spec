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
#           └── app_v1.0.0.pyz

from pathlib import Path

BASE = Path(SPEC).parent  # launcher/ 폴더

a = Analysis(
    [str(BASE / "launcher.py")],
    pathex=[str(BASE.parent)],
    binaries=[],
    datas=[
        # App 폴더 전체를 배포물에 포함
        (str(BASE / "App"), "App"),
    ],
    hiddenimports=[
        "PySide6.QtCore",
        "PySide6.QtGui",
        "PySide6.QtWidgets",
        "openpyxl",
        "openpyxl.styles",
        "openpyxl.utils",
        "fitz",
        "pythoncom",
        "win32com",
        "win32com.client",
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

exe = EXE(
    pyz,
    a.scripts,
    [],
    exclude_binaries=True,
    name="QuoteAuto",
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
    name="QuoteAuto",
)
