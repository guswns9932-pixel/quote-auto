"""전자서명 — 단독 실행 진입점 (런처 없이 이 앱만 띄운다)."""

from __future__ import annotations

from app_runner import run_qt_page


def main() -> None:
    run_qt_page("전자서명", "page_esign", "ESignPage")


if __name__ == "__main__":
    main()
