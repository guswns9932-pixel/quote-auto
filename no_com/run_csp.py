"""CSP UPLOAD 자동화 — 단독 실행 진입점 (런처 없이 이 앱만 띄운다)."""

from __future__ import annotations

from app_boot import run_tk_app


def _factory():
    from csp_order_maker import App
    return App()


def main() -> None:
    run_tk_app("CSP UPLOAD 자동화", _factory)


if __name__ == "__main__":
    main()
