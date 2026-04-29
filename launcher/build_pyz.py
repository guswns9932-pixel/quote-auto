"""
build_pyz.py
============
앱 코드(.py)를 .pyz로 패키징하고 App/<name>/manifest.json을 갱신한다.
코드 변경 시 이 스크립트만 실행하면 배포 준비 완료 (3~5초).

사용법:
    python build_pyz.py                          # 버전 자동 증가
    python build_pyz.py 1.3.1                   # 버전 직접 지정
    python build_pyz.py 1.3.1 --name 견적자동화  # 앱 이름 지정

앱 이름 미지정 시 config 섹션의 APP_NAME 사용.
"""

from __future__ import annotations

import compileall
import hashlib
import json
import os
import sys
import zipfile
from datetime import datetime
from pathlib import Path


# ════════════════════════════════════════════════
# ★ 앱별 설정 (프로그램마다 이 부분을 수정)
# ════════════════════════════════════════════════

APP_NAME    = "견적자동화"
APP_DESC    = "견적서 작성 및 전자서명 통합 관리 시스템"
APP_ENTRY   = "main"          # pyz 안에서 실행할 모듈명 (main.main() 호출)

SRC_DIR = Path(__file__).resolve().parent.parent   # quote-auto/
OUT_ROOT = Path(__file__).resolve().parent / "App" # quote-auto/launcher/App/

APP_FILES = [
    "main.py",
    "pages.py",
    "core.py",
    "excel_io.py",
    "widgets.py",
]

# ════════════════════════════════════════════════


def _resolve_app_name() -> str:
    for i, arg in enumerate(sys.argv):
        if arg == "--name" and i + 1 < len(sys.argv):
            return sys.argv[i + 1]
    return APP_NAME


def _resolve_version(app_name: str) -> str:
    """CLI 버전 지정 → 없으면 현재 manifest 버전에서 패치 자동 증가."""
    for arg in sys.argv[1:]:
        if arg.startswith("--"):
            continue
        parts = arg.split(".")
        if all(p.isdigit() for p in parts) and len(parts) >= 2:
            return arg
    manifest_path = OUT_ROOT / app_name / "manifest.json"
    if manifest_path.exists():
        try:
            cur = json.loads(manifest_path.read_text(encoding="utf-8")).get("version", "1.0.0")
            parts = cur.split(".")
            parts[-1] = str(int(parts[-1]) + 1)
            return ".".join(parts)
        except Exception:
            pass
    return "1.0.0"


def build(app_name: str, version: str) -> None:
    out_dir  = OUT_ROOT / app_name
    out_dir.mkdir(parents=True, exist_ok=True)

    pyz_name = f"app_v{version}.pyz"
    pyz_path = out_dir / pyz_name

    print(f"{'─'*50}")
    print(f"  앱 이름 : {app_name}")
    print(f"  버전    : {version}")
    print(f"  소스    : {SRC_DIR}")
    print(f"  출력    : {out_dir}")
    print(f"{'─'*50}")

    # ── ① .py → .pyc 사전 컴파일 ──────────────
    print("① .pyc 사전 컴파일...")
    compileall.compile_dir(str(SRC_DIR), quiet=1, force=True)

    # ── ② .pyz 패키징 ──────────────────────────
    print("② .pyz 패키징...")
    missing = [f for f in APP_FILES if not (SRC_DIR / f).exists()]
    if missing:
        print(f"  [오류] 파일 없음: {missing}"); sys.exit(1)

    with zipfile.ZipFile(pyz_path, "w", zipfile.ZIP_DEFLATED, compresslevel=6) as z:
        for fname in APP_FILES:
            src = SRC_DIR / fname
            z.write(src, fname)
            print(f"  + {fname}")
            cache_dir = SRC_DIR / "__pycache__"
            stem = Path(fname).stem
            candidates = sorted(cache_dir.glob(f"{stem}.*.pyc")) if cache_dir.exists() else []
            if candidates:
                pyc = candidates[-1]
                z.write(pyc, pyc.name)
                print(f"    └─ {pyc.name}")

    size_kb = pyz_path.stat().st_size // 1024
    print(f"\n  → {pyz_name}  ({size_kb} KB)")

    # ── ③ SHA-256 ───────────────────────────────
    print("③ SHA-256 계산...")
    sha256 = hashlib.sha256(pyz_path.read_bytes()).hexdigest()
    print(f"  → {sha256[:16]}...")

    # ── ④ manifest.json ─────────────────────────
    print("④ manifest.json 갱신...")
    manifest = {
        "name":         app_name,
        "description":  APP_DESC,
        "version":      version,
        "pyz":          pyz_name,
        "sha256":       sha256,
        "built_at":     datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        "entry":        APP_ENTRY,
        "launcher_min": "1.0.0",
    }
    manifest_path = out_dir / "manifest.json"
    manifest_path.write_text(json.dumps(manifest, indent=2, ensure_ascii=False), encoding="utf-8")
    print(f"  → {manifest_path}")

    # ── ⑤ 이전 버전 .pyz 정리 ───────────────────
    old = [f for f in out_dir.glob("app_v*.pyz") if f.name != pyz_name]
    if old:
        print(f"⑤ 이전 버전 정리: {[f.name for f in old]}")
        for f in old:
            f.unlink(); print(f"  삭제: {f.name}")

    print(f"\n✓ 완료  →  App/{app_name}/{pyz_name}")
    print(f"  배포 파일: App/{app_name}/manifest.json  +  App/{app_name}/{pyz_name}")
    print(f"{'─'*50}")


if __name__ == "__main__":
    _name = _resolve_app_name()
    _ver  = _resolve_version(_name)
    if "--name" not in sys.argv:
        print(f"[build_pyz] 앱: {_name}  버전 자동 증가 → {_ver}")
    build(_name, _ver)
