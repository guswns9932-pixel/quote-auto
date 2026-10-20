"""Writable user data is independent of the installed executable directory."""
import os
from pathlib import Path


def local_data_dir() -> Path:
    if os.name == "nt":
        root = os.environ.get("LOCALAPPDATA")
        if not root:
            raise RuntimeError("LOCALAPPDATA가 없어 로컬 저장 위치를 찾을 수 없습니다.")
        return Path(root) / "QuoteAuto"
    return Path(os.environ.get("XDG_DATA_HOME", str(Path.home() / ".local/share"))) / "QuoteAuto"


def default_output_dir() -> str:
    return str(local_data_dir() / "documents")


def log_dir() -> Path:
    path = local_data_dir() / "logs"
    path.mkdir(parents=True, exist_ok=True)
    return path
