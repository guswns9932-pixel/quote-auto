"""폴더 순회와 파일 검색을 담당하는 GUI 비의존 스캔 엔진.

os.scandir로 디렉터리를 순회하며 발견 즉시 스트리밍하고, 파일 하나하나의
검색 작업은 스레드 풀에 맡겨 동시에 여러 파일을 검사한다(대부분 시간을
디스크/파서 라이브러리의 I/O 대기에 쓰기 때문에 스레드 병렬화 효과가 크다).
"""

import os
import queue
from concurrent.futures import ThreadPoolExecutor
from pathlib import Path

from .parsers import SUPPORTED_EXTENSIONS, file_contains_excluded_text, search_file_contents
from .text_utils import find_matches

DEFAULT_MAX_WORKERS = 8


def iter_paths(folders, include_subfolders, extensions, is_cancelled):
    """여러 폴더 아래 파일 경로를 찾는 즉시 하나씩 내보낸다(전체 목록을 먼저 모으지 않음)."""
    stack = list(folders)

    while stack:
        current = stack.pop()
        try:
            entries = os.scandir(current)
        except OSError:
            continue

        with entries:
            for entry in entries:
                if is_cancelled():
                    return
                try:
                    if entry.is_dir(follow_symlinks=False):
                        if include_subfolders:
                            stack.append(entry.path)
                        continue
                    if not entry.is_file(follow_symlinks=False):
                        continue
                except OSError:
                    continue

                path = Path(entry.path)
                if extensions and path.suffix.lower() not in extensions:
                    continue
                yield path


def search_one_file(path, pattern, search_filename, search_contents, exclude_pattern=None):
    """
    파일 하나를 검사해 (위치, 일치한 키워드, 스니펫) 목록을 반환한다.

    exclude_pattern이 주어지면, 검색 범위(파일명/내용) 안에 그 패턴이 하나라도
    있는 파일은 검색 키워드와 무관하게 통째로 제외한다(빈 리스트 반환).
    """
    if exclude_pattern is not None:
        if search_filename and exclude_pattern.search(path.name):
            return []
        if (
            search_contents
            and path.suffix.lower() in SUPPORTED_EXTENSIONS
            and file_contains_excluded_text(path, exclude_pattern)
        ):
            return []

    rows = []

    if search_filename:
        for _, _, keyword, snippet in find_matches(path.name, pattern):
            rows.append(("파일명", keyword, snippet))

    if search_contents and path.suffix.lower() in SUPPORTED_EXTENSIONS:
        rows.extend(search_file_contents(path, pattern))

    return rows


def scan(
    folders,
    pattern,
    include_subfolders,
    search_filename,
    search_contents,
    extensions,
    is_cancelled,
    on_file_start,
    on_file_result,
    on_error,
    exclude_pattern=None,
    max_workers=DEFAULT_MAX_WORKERS,
):
    """
    여러 폴더를 병렬로 검색한다. folders는 폴더 경로의 리스트다.

    on_file_start(processed_count, path), on_file_result(path, location, keyword, snippet),
    on_error(message) 콜백은 스캔이 실행되는 스레드에서 호출된다.
    exclude_pattern이 주어지면, 검색 범위(파일명/내용) 안 어딘가에 그 패턴이
    하나라도 있는 파일은 통째로 결과에서 제외한다.
    반환값: (processed_count, total_hits, matched_file_count)
    """
    processed = 0
    total_hits = 0
    matched_files = set()
    completed = queue.SimpleQueue()
    paths_iter = iter_paths(folders, include_subfolders, extensions, is_cancelled)

    executor = ThreadPoolExecutor(max_workers=max_workers)
    pending = 0

    def submit_next():
        nonlocal pending
        for path in paths_iter:
            future = executor.submit(
                search_one_file, path, pattern, search_filename, search_contents, exclude_pattern
            )
            future.path = path
            future.add_done_callback(completed.put)
            pending += 1
            return True
        return False

    try:
        for _ in range(max_workers):
            if not submit_next():
                break

        while pending > 0:
            if is_cancelled():
                break

            future = completed.get()
            pending -= 1
            path = future.path
            processed += 1
            on_file_start(processed, path)

            try:
                rows = future.result()
            except PermissionError:
                on_error(f"권한 없음: {path}")
                rows = []
            except Exception as exc:
                on_error(f"{path} | {exc}")
                rows = []

            for location, keyword, snippet in rows:
                total_hits += 1
                matched_files.add(path)
                on_file_result(path, location, keyword, snippet)

            submit_next()
    finally:
        executor.shutdown(wait=not is_cancelled(), cancel_futures=is_cancelled())

    return processed, total_hits, len(matched_files)
