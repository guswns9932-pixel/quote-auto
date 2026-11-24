"""키워드 검색기 — 단독 실행 진입점 (런처 없이 이 앱만 띄운다)."""

from __future__ import annotations

from app_boot import run_tk_app


def _factory():
    from content_search.gui import ContentSearchApp
    return ContentSearchApp()


def main() -> None:
    run_tk_app("키워드 검색기", _factory)


if __name__ == "__main__":
    main()
