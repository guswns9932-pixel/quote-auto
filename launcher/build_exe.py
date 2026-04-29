"""Windows EXE builder for launcher.

Usage:
  python launcher/build_exe.py
  python launcher/build_exe.py --clean
  python launcher/build_exe.py --noconfirm
"""

from __future__ import annotations

import argparse
import shutil
import subprocess
import sys
from pathlib import Path


def run(cmd: list[str], cwd: Path) -> None:
    print("$", " ".join(cmd))
    result = subprocess.run(cmd, cwd=str(cwd))
    if result.returncode != 0:
        raise SystemExit(result.returncode)


def main() -> None:
    parser = argparse.ArgumentParser(description="Build QuoteAuto launcher EXE with PyInstaller")
    parser.add_argument("--clean", action="store_true", help="remove build/dist before building")
    parser.add_argument("--noconfirm", action="store_true", help="pass --noconfirm to pyinstaller")
    args = parser.parse_args()

    base = Path(__file__).resolve().parent
    dist_dir = base / "dist"
    build_dir = base / "build"
    spec_file = base / "launcher.spec"

    if sys.platform != "win32":
        print("[경고] 이 스크립트는 Windows 빌드를 전제로 작성되었습니다.")

    if not spec_file.exists():
        raise SystemExit(f"[오류] spec 파일이 없습니다: {spec_file}")

    if args.clean:
        for target in (build_dir, dist_dir):
            if target.exists():
                print(f"[clean] 삭제: {target}")
                shutil.rmtree(target)

    cmd = [sys.executable, "-m", "PyInstaller", "launcher.spec"]
    if args.noconfirm:
        cmd.append("--noconfirm")
    run(cmd, cwd=base)

    exe_path = dist_dir / "QuoteAuto" / "QuoteAuto.exe"
    if not exe_path.exists():
        raise SystemExit(f"[오류] 빌드 결과 EXE를 찾을 수 없습니다: {exe_path}")

    print("\n[완료] 빌드 성공")
    print(f"- EXE: {exe_path}")
    print(f"- App 포함 경로: {dist_dir / 'QuoteAuto' / 'App'}")
    print("\n[권장 확인]")
    print("1) QuoteAuto.exe 실행 후 앱 목록이 표시되는지")
    print("2) 견적자동화 실행 시 openpyxl import 오류가 사라졌는지")


if __name__ == "__main__":
    main()
