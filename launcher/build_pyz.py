"""
build_pyz.py
============
앱 코드(.py)를 .pyz로 패키징하고 manifest.json을 갱신한다.
코드 변경 시 이 스크립트만 실행하면 배포 준비 완료 (3~5초).

사용법:
    python build_pyz.py [version]
    python build_pyz.py 1.3.1

버전 미지정 시 manifest.json의 현재 버전에서 패치 번호를 자동 증가.
"""

from __future__ import annotations

import compileall
import hashlib
import json
import os
import sys
import zipfile
from pathlib import Path
from datetime import datetime

# ── 설정 ────────────────────────────────────────
SRC_DIR = Path(__file__).resolve().parent.parent   # quote-auto/
OUT_DIR = Path(__file__).resolve().parent          # quote-auto/launcher/

APP_FILES = [
    "main.py",
    "pages.py",
    "core.py",
    "excel_io.py",
    "widgets.py",
]

MANIFEST_PATH = OUT_DIR / "manifest.json"
# ────────────────────────────────────────────────


def _next_version(current: str) -> str:
    parts = current.split(".")
    parts[-1] = str(int(parts[-1]) + 1)
    return ".".join(parts)


def _current_version() -> str:
    if MANIFEST_PATH.exists():
        try:
            return json.loads(MANIFEST_PATH.read_text(encoding="utf-8")).get("version", "1.0.0")
        except Exception:
            pass
    return "1.0.0"


def build(version: str) -> None:
    pyz_name = f"app_v{version}.pyz"
    pyz_path = OUT_DIR / pyz_name

    print(f"[build_pyz] 버전: {version}")
    print(f"[build_pyz] 소스: {SRC_DIR}")
    print(f"[build_pyz] 출력: {pyz_path}")
    print()

    # ── 1. .py → .pyc 사전 컴파일 ──────────────
    print("① .pyc 사전 컴파일 중...")
    compileall.compile_dir(str(SRC_DIR), quiet=1, force=True)

    # ── 2. .pyz 패키징 ──────────────────────────
    print("② .pyz 패키징 중...")
    missing = [f for f in APP_FILES if not (SRC_DIR / f).exists()]
    if missing:
        print(f"  [오류] 파일 없음: {missing}")
        sys.exit(1)

    with zipfile.ZipFile(pyz_path, "w", zipfile.ZIP_DEFLATED, compresslevel=6) as z:
        for fname in APP_FILES:
            src = SRC_DIR / fname
            z.write(src, fname)
            print(f"  + {fname}")

            # __pycache__ 에서 .pyc 찾아 함께 추가
            stem = Path(fname).stem
            cache_dir = SRC_DIR / "__pycache__"
            pyc_candidates = sorted(cache_dir.glob(f"{stem}.*.pyc")) if cache_dir.exists() else []
            if pyc_candidates:
                pyc = pyc_candidates[-1]   # 최신 버전 우선
                z.write(pyc, pyc.name)
                print(f"    └─ {pyc.name}")

    size_kb = pyz_path.stat().st_size // 1024
    print(f"\n  → {pyz_name}  ({size_kb} KB)")

    # ── 3. SHA-256 해시 ─────────────────────────
    print("③ SHA-256 계산 중...")
    sha256 = hashlib.sha256(pyz_path.read_bytes()).hexdigest()
    print(f"  → {sha256[:16]}...")

    # ── 4. manifest.json 갱신 ───────────────────
    print("④ manifest.json 갱신 중...")
    manifest = {
        "version": version,
        "pyz": pyz_name,
        "sha256": sha256,
        "built_at": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        "launcher_min": "1.0.0",
    }
    MANIFEST_PATH.write_text(json.dumps(manifest, indent=2, ensure_ascii=False), encoding="utf-8")
    print(f"  → manifest.json 저장 완료")

    # ── 5. 이전 버전 .pyz 정리 (선택) ──────────
    old_pyzs = [f for f in OUT_DIR.glob("app_v*.pyz") if f.name != pyz_name]
    if old_pyzs:
        print(f"\n⑤ 이전 버전 정리: {[f.name for f in old_pyzs]}")
        for f in old_pyzs:
            f.unlink()
            print(f"  삭제: {f.name}")

    print(f"\n✓ 빌드 완료 → {pyz_name}")
    print(f"  배포 파일: {pyz_name}  +  manifest.json")


if __name__ == "__main__":
    if len(sys.argv) >= 2:
        ver = sys.argv[1]
    else:
        current = _current_version()
        ver = _next_version(current)
        print(f"[build_pyz] 버전 자동 증가: {current} → {ver}")

    build(ver)
