"""
build.py
========
no_com/ 앱을 배포용 단일 EXE 로 빌드하는 스크립트.

동작 방식 (2단계)
-----------------
 Stage 1 — 앱 본체: onedir 로 빌드 (빠른 실행, 추출 불필요)
 Stage 2 — 런처 스텁: bundle.zip + 소형 런처를 onefile 로 묶음

배포 EXE 실행 흐름
------------------
 최초 실행:
   AppName.exe 실행
   → %LOCALAPPDATA%\\AppName\\ 에 bundle.zip 압축 해제 (1회)
   → AppName.exe (onedir) 실행

 이후 실행:
   AppName.exe 실행
   → 이미 설치됨 감지 (해시 일치)
   → AppName.exe (onedir) 즉시 실행  ← 빠름

사용법
------
  python build.py                 # 이름 프롬프트
  python build.py 견적자동화      # → dist/견적자동화.exe
  python build.py 견적자동화 icon.ico
"""

import os
import sys
import textwrap
import subprocess
import zipfile

HERE = os.path.dirname(os.path.abspath(__file__))

# ── 런처 스텁 코드 템플릿 ──────────────────────────────────────────
# PLACEHOLDER_APP_NAME 은 빌드 시 실제 앱 이름으로 치환됨
STUB_TEMPLATE = r'''
import os, sys, zipfile, subprocess

APP_NAME   = "PLACEHOLDER_APP_NAME"
INSTALL_DIR = os.path.join(
    os.environ.get("LOCALAPPDATA", os.path.expanduser("~")), APP_NAME
)
VER_FILE   = os.path.join(INSTALL_DIR, ".ver")


def _bundle_path():
    base = getattr(sys, "_MEIPASS", os.path.dirname(os.path.abspath(__file__)))
    return os.path.join(base, "bundle.zip")


def _bundle_sig(path):
    """파일 크기 + 수정일시로 번들 고유 서명 생성 (SHA256 대비 즉각적)."""
    st = os.stat(path)
    return f"{st.st_size}:{int(st.st_mtime)}"


def _already_installed(bundle_path):
    try:
        with open(VER_FILE, encoding="utf-8") as f:
            return f.read().strip() == _bundle_sig(bundle_path)
    except Exception:
        return False


def _install(bundle_path):
    import shutil
    if os.path.exists(INSTALL_DIR):
        shutil.rmtree(INSTALL_DIR, ignore_errors=True)
    os.makedirs(INSTALL_DIR, exist_ok=True)
    with zipfile.ZipFile(bundle_path) as zf:
        zf.extractall(INSTALL_DIR)
    with open(VER_FILE, "w", encoding="utf-8") as f:
        f.write(_bundle_sig(bundle_path))


def _show_msg(title, msg, icon=0x40):
    # ctypes 로 MessageBox 표시 (외부 의존 없음)
    try:
        import ctypes
        ctypes.windll.user32.MessageBoxW(0, msg, title, icon)
    except Exception:
        print(msg)


def main():
    bundle = _bundle_path()

    if not _already_installed(bundle):
        _show_msg(
            APP_NAME,
            "처음 실행 시 초기 설치가 진행됩니다.\n\n"
            "확인을 누르면 설치가 시작되며 잠시 후 자동으로 실행됩니다.",
            0x40,
        )
        try:
            _install(bundle)
        except Exception as e:
            _show_msg(APP_NAME, f"설치 실패:\n{e}", 0x10)
            sys.exit(1)

    exe = os.path.join(INSTALL_DIR, APP_NAME, APP_NAME + ".exe")
    if not os.path.exists(exe):
        _show_msg(APP_NAME, f"실행 파일을 찾을 수 없습니다:\n{exe}", 0x10)
        sys.exit(1)

    subprocess.Popen([exe] + sys.argv[1:])


if __name__ == "__main__":
    main()
'''


# ── 인자 처리 ──────────────────────────────────────────────────────

def _get_args():
    args = sys.argv[1:]
    if args:
        app_name = args[0].strip()
        icon     = args[1].strip() if len(args) >= 2 else ""
    else:
        app_name = input("프로세스명(EXE 파일명, 예: 견적자동화): ").strip()
        icon     = input("아이콘 경로 (없으면 엔터): ").strip()

    if not app_name:
        print("오류: 이름을 입력하세요.")
        sys.exit(1)

    if icon and not os.path.isabs(icon):
        icon = os.path.join(HERE, icon)
    if icon and not os.path.exists(icon):
        print(f"경고: 아이콘 파일 없음 — {icon}")
        icon = ""

    return app_name, icon


# ── Stage 1: 앱 본체 spec (onedir) ────────────────────────────────

def _make_app_spec(app_name: str, icon: str) -> str:
    icon_line = f'icon=r"{icon}",' if icon else ""
    return textwrap.dedent(f"""\
        # Stage 1 — app onedir spec (auto-generated)
        import glob, site
        from pathlib import Path
        from PyInstaller.utils.hooks import collect_all

        HERE = Path(r"{HERE}")

        openpyxl_d, openpyxl_b, openpyxl_h = collect_all("openpyxl")
        fitz_d,     fitz_b,     fitz_h     = collect_all("fitz")

        _pywin32_dlls = []
        for _sp_dir in site.getsitepackages():
            _pywin32_dlls += [
                (f, ".")
                for f in glob.glob(str(Path(_sp_dir) / "pywin32_system32" / "*.dll"))
            ]

        a = Analysis(
            [str(HERE / "main.py")],
            pathex=[str(HERE)],
            binaries=[*openpyxl_b, *fitz_b, *_pywin32_dlls],
            datas=[*openpyxl_d, *fitz_d],
            hiddenimports=[
                "PySide6.QtCore", "PySide6.QtGui", "PySide6.QtWidgets",
                "PySide6.QtSvg", "PySide6.QtPrintSupport",
                "openpyxl.cell._writer",
                *openpyxl_h, *fitz_h,
                "pythoncom", "pywintypes", "win32api",
                "win32com", "win32com.client", "win32com.server",
                "win32con", "win32timezone",
            ],
            hookspath=[],
            runtime_hooks=[],
            excludes=[
                "tkinter", "matplotlib", "numpy", "scipy", "PIL",
                "unittest", "doctest", "pdb", "profile", "cProfile",
                "difflib", "pickletools", "ftplib",
                "multiprocessing.pool",
            ],
            noarchive=False,  # .pyc를 PYZ 아카이브로 압축 → 첫 실행 시 Defender 스캔 파일 수 최소화
            optimize=2,       # 2: docstring 제거 → .pyc 크기 추가 감소, 임포트 속도 향상
        )

        pyz = PYZ(a.pure)

        exe = EXE(
            pyz, a.scripts, [],
            exclude_binaries=True,
            name="{app_name}",
            debug=False, strip=False, upx=False, console=False,
            {icon_line}
        )

        coll = COLLECT(
            exe, a.binaries, a.datas,
            strip=False, upx=True, upx_exclude=[],
            name="{app_name}",
        )
    """)


# ── Stage 2: 런처 스텁 spec (onefile, 초소형) ─────────────────────

def _make_stub_spec(app_name: str, stub_py: str, bundle_zip: str, icon: str) -> str:
    icon_line = f'icon=r"{icon}",' if icon else ""
    return textwrap.dedent(f"""\
        # Stage 2 — launcher stub spec (auto-generated)
        a = Analysis(
            [r"{stub_py}"],
            pathex=[],
            binaries=[],
            datas=[(r"{bundle_zip}", ".")],
            hiddenimports=[],
            hookspath=[],
            runtime_hooks=[],
            excludes=[],
            noarchive=False,
        )

        pyz = PYZ(a.pure)

        exe = EXE(
            pyz, a.scripts, a.binaries, a.datas,
            name="{app_name}",
            debug=False,
            bootloader_ignore_signals=False,
            strip=False,
            upx=True,
            upx_exclude=[],
            runtime_tmpdir=None,
            console=False,
            {icon_line}
        )
    """)


# ── 유틸 ──────────────────────────────────────────────────────────

def _run(spec_path: str, dist_dir: str, work_dir: str) -> None:
    cmd = [
        sys.executable, "-m", "PyInstaller", spec_path,
        "--distpath", dist_dir,
        "--workpath", work_dir,
        "--noconfirm", "--clean",
    ]
    result = subprocess.run(cmd, cwd=HERE)
    if result.returncode != 0:
        print(f"\n빌드 실패 (exit={result.returncode})")
        sys.exit(result.returncode)


def _zip_folder(folder: str, out_zip: str) -> None:
    """onedir 폴더 → ZIP (압축 레벨 6)."""
    parent = os.path.dirname(folder)
    total = sum(
        os.path.getsize(os.path.join(r, f))
        for r, _, fs in os.walk(folder) for f in fs
    ) / 1024 / 1024
    print(f"  압축 대상: {total:.1f} MB")
    with zipfile.ZipFile(out_zip, "w", zipfile.ZIP_DEFLATED, compresslevel=6) as zf:
        for root, _, files in os.walk(folder):
            for fname in files:
                fp = os.path.join(root, fname)
                zf.write(fp, os.path.relpath(fp, parent))
    size_mb = os.path.getsize(out_zip) / 1024 / 1024
    print(f"  bundle.zip: {size_mb:.1f} MB")


# ── 메인 ──────────────────────────────────────────────────────────

def main():
    app_name, icon = _get_args()

    tmp_dir  = os.path.join(HERE, "_build_tmp")
    os.makedirs(tmp_dir, exist_ok=True)

    # ── Stage 1: 앱 본체 onedir 빌드 ────────────────────────────
    print(f"\n{'='*54}")
    print(f"  [1/3] 앱 본체 빌드 (onedir) …")
    print(f"{'='*54}")
    app_dist = os.path.join(tmp_dir, "app_dist")
    app_work = os.path.join(tmp_dir, "app_work")
    app_spec = os.path.join(tmp_dir, f"{app_name}_app.spec")
    with open(app_spec, "w", encoding="utf-8") as f:
        f.write(_make_app_spec(app_name, icon))
    _run(app_spec, app_dist, app_work)

    # ── Stage 2: bundle.zip 생성 ────────────────────────────────
    print(f"\n{'='*54}")
    print(f"  [2/3] bundle.zip 압축 중 …")
    print(f"{'='*54}")
    bundle_zip = os.path.join(tmp_dir, "bundle.zip")
    _zip_folder(os.path.join(app_dist, app_name), bundle_zip)

    # ── Stage 3: 런처 스텁 onefile 빌드 ─────────────────────────
    print(f"\n{'='*54}")
    print(f"  [3/3] 런처 스텁 빌드 (onefile) …")
    print(f"{'='*54}")
    stub_py = os.path.join(tmp_dir, f"{app_name}_stub.py")
    stub_content = STUB_TEMPLATE.replace("PLACEHOLDER_APP_NAME", app_name)
    with open(stub_py, "w", encoding="utf-8") as f:
        f.write(stub_content)
    stub_spec = os.path.join(tmp_dir, f"{app_name}.spec")
    with open(stub_spec, "w", encoding="utf-8") as f:
        f.write(_make_stub_spec(app_name, stub_py, bundle_zip, icon))
    dist_dir  = os.path.join(HERE, "dist")
    stub_work = os.path.join(tmp_dir, "stub_work")
    _run(stub_spec, dist_dir, stub_work)

    # ── 결과 출력 ───────────────────────────────────────────────
    out = os.path.join(dist_dir, f"{app_name}.exe")
    if os.path.exists(out):
        size_mb = os.path.getsize(out) / 1024 / 1024
        print(f"\n{'='*54}")
        print(f"  완료: {out}")
        print(f"  크기: {size_mb:.1f} MB")
        print()
        print(f"  최초 실행: %LOCALAPPDATA%\\{app_name}\\ 에 자동 설치")
        print(f"  이후 실행: 즉시 시작 (재추출 없음)")
        print(f"  업데이트: 새 EXE 배포 시 해시 감지 → 자동 재설치")
        print(f"{'='*54}\n")
    else:
        print("\n빌드 실패: 출력 파일 없음")
        sys.exit(1)


if __name__ == "__main__":
    main()
