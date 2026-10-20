"""Create a Windows local-install distribution from a completed onedir build."""
import hashlib
import json
import shutil
import tempfile
import zipfile
from pathlib import Path


def package_local(build_dir, app_name):
    build_dir = Path(build_dir).resolve()
    if Path(app_name).name != app_name or any(c in app_name for c in '/\\:*?"<>|'):
        raise ValueError("앱 이름은 파일 이름이어야 합니다.")
    launcher = app_name + ".exe"
    if not (build_dir / launcher).is_file():
        raise FileNotFoundError(build_dir / launcher)
    destination = build_dir.parent / (app_name + "-install")
    destination.mkdir(exist_ok=True)
    with tempfile.TemporaryDirectory(dir=destination) as tmp:
        archive = Path(tmp) / "payload.zip"
        with zipfile.ZipFile(archive, "w", zipfile.ZIP_DEFLATED) as z:
            for path in sorted(build_dir.rglob("*")):
                if path.is_symlink():
                    raise ValueError(f"배포 파일에 심볼릭 링크를 포함할 수 없습니다: {path}")
                if path.is_file():
                    z.write(path, path.relative_to(build_dir).as_posix())
        hasher = hashlib.sha256()
        with archive.open("rb") as stream:
            for chunk in iter(lambda: stream.read(1024 * 1024), b""):
                hasher.update(chunk)
        digest = hasher.hexdigest()
        package_name = f"package-{digest}.zip"
        archive.replace(destination / package_name)
        manifest = {"schema": 1, "sha256": digest, "package": package_name,
                    "launcher": launcher, "app_name": app_name}
        manifest_path = Path(tmp) / "release.json"
        manifest_path.write_text(json.dumps(manifest, ensure_ascii=False, indent=2), encoding="utf-8")
        installer = Path(__file__).with_name("install_local.ps1")
        shutil.copyfile(installer, destination / installer.name)
        # cmd contains only ASCII so it works with the Windows system code page.
        (destination / "Install.cmd").write_bytes(
            b'@echo off\r\npowershell.exe -NoProfile -File "%~dp0install_local.ps1"\r\n'
            b'if errorlevel 1 (pause & exit /b 1)\r\npause\r\n')
        # Publish the manifest last; never point at an incomplete archive.
        manifest_path.replace(destination / "release.json")
    return destination
