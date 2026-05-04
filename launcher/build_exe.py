"""Windows EXE builder for launcher.

Usage:
  python launcher/build_exe.py
  python launcher/build_exe.py --clean
  python launcher/build_exe.py --noconfirm
  python launcher/build_exe.py --onefile --name MyLauncher
"""

from __future__ import annotations

import argparse
import os
import shutil
import subprocess
import sys
from pathlib import Path


def run(cmd: list[str], cwd: Path, env: dict[str, str] | None = None) -> None:
    print("$", " ".join(cmd))
    result = subprocess.run(cmd, cwd=str(cwd), env=env)
    if result.returncode != 0:
        raise SystemExit(result.returncode)


def main() -> None:
    parser = argparse.ArgumentParser(description="Build QuoteAuto launcher EXE with PyInstaller")
    parser.add_argument("--clean", action="store_true", help="remove build/dist before building")
    parser.add_argument("--noconfirm", action="store_true", help="pass --noconfirm to pyinstaller")
    parser.add_argument("--onefile", action="store_true", help="build a single EXE file")
    parser.add_argument("--name", help="output EXE base name (without .exe)")
    args = parser.parse_args()

    base = Path(__file__).resolve().parent
    dist_dir = base / "dist"
    build_dir = base / "build"
    spec_file = base / "launcher.spec"

    if sys.platform != "win32":
        print("[경고] 이 스크립트는 Windows 빌드를 전제로 작성되었습니다.")

    if not spec_file.exists():
        raise SystemExit(f"[오류] spec 파일이 없습니다: {spec_file}")

    exe_name = (args.name or "").strip()
    if not exe_name:
        exe_name = input("생성할 EXE 파일 이름을 입력하세요 (기본: QuoteAuto): ").strip() or "QuoteAuto"

    if args.clean:
        for target in (build_dir, dist_dir):
            if target.exists():
                print(f"[clean] 삭제: {target}")
                shutil.rmtree(target)

    cmd = [sys.executable, "-m", "PyInstaller", "launcher.spec"]
    if args.noconfirm:
        cmd.append("--noconfirm")

    env = dict(os.environ)
    env["QUOTEAUTO_ONEFILE"] = "1" if args.onefile else "0"
    env["QUOTEAUTO_EXE_NAME"] = exe_name
    run(cmd, cwd=base, env=env)

    if args.onefile:
        exe_path = dist_dir / f"{exe_name}.exe"
    else:
        exe_path = dist_dir / exe_name / f"{exe_name}.exe"
    if not exe_path.exists():
        raise SystemExit(f"[오류] 빌드 결과 EXE를 찾을 수 없습니다: {exe_path}")

    print("\n[완료] 빌드 성공")
    print(f"- EXE: {exe_path}")
    if args.onefile:
        print("- 단일 EXE 모드: App 리소스는 실행 시 임시 폴더로 자동 압축해제됩니다.")
    else:
        print(f"- App 포함 경로: {dist_dir / exe_name / 'App'}")
    print("\n[권장 확인]")
    print(f"1) {exe_name}.exe 실행 후 앱 목록이 표시되는지")
    print("2) 견적자동화 실행 시 openpyxl import 오류가 사라졌는지")


if __name__ == "__main__":
    main()
