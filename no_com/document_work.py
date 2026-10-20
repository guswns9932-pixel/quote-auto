"""Work locally; publish a completed file without overwriting another user's file."""
import os
import logging
import shutil
import tempfile
from dataclasses import replace
from pathlib import Path

from app_paths import local_data_dir


class DocumentPublishError(OSError):
    def __init__(self, path, cause):
        self.local_path = str(path)
        super().__init__(f"선택한 폴더에 저장하지 못했습니다. 완성된 로컬 파일: {path}\n{cause}")


def prepare_quote(state):
    root = local_data_dir() / "work"
    root.mkdir(parents=True, exist_ok=True)
    work = Path(tempfile.mkdtemp(prefix="quote-", dir=root))
    local = replace(state, output_root=str(work / "output"))
    try:
        for attr in ("template_path", "request_path"):
            source = getattr(state, attr)
            if source:
                target = work / (attr + Path(source).suffix)
                shutil.copyfile(source, target)
                setattr(local, attr, str(target))
    except BaseException:
        shutil.rmtree(work, ignore_errors=True)
        raise
    return work, local


def publish_document(source, destination):
    """Copy once, then expose the completed file. Never replace an existing file.

    Windows rename is non-overwriting. POSIX uses an exclusive hard link because
    rename would overwrite. Unsupported filesystem operations fail safely.
    Cloud synchronization beyond the mounted filesystem is provider-managed.
    """
    source, destination = Path(source), Path(destination)
    destination.parent.mkdir(parents=True, exist_ok=True)
    fd, temporary = tempfile.mkstemp(prefix=".quote-upload-", suffix=".tmp", dir=destination.parent)
    os.close(fd)
    try:
        shutil.copyfile(source, temporary)
        for number in range(1, 10001):
            candidate = destination if number == 1 else destination.with_name(
                f"{destination.stem}_{number}{destination.suffix}")
            try:
                if os.name == "nt":
                    os.rename(temporary, candidate)
                else:
                    os.link(temporary, candidate)
                return str(candidate)
            except FileExistsError:
                continue
        raise FileExistsError("같은 이름의 결과 파일이 너무 많습니다.")
    finally:
        try:
            os.unlink(temporary)
        except FileNotFoundError:
            pass
        except OSError:
            logging.getLogger("QuoteApp").warning("업로드 임시 파일 정리 실패: %s", temporary)


def publish_quote(path, local_state, output_root):
    relative = Path(path).relative_to(local_state.output_root)
    try:
        return publish_document(path, Path(output_root) / relative)
    except Exception as exc:
        raise DocumentPublishError(path, exc) from exc
