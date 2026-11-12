"""
app_main.py
===========
단일 exe 빌드(build_onedir.py --single)의 진입점.

exe 하나(해시 하나)로 런처와 4개 앱을 모두 실행한다. 런처가 이 exe 를
`--quote` 등의 인자로 다시 실행하면 해당 앱만 띄운다. [중요] 인자를 보기 전에
PySide6 같은 무거운 모듈을 임포트하지 않는다 — Tk 앱이 Qt 를 로드하면 단일 exe
방식의 이점이 사라진다.
"""

from __future__ import annotations

import sys

# 인자 → run_*.py 모듈. 임포트는 분기 이후에 한다.
_MODES = {
    "--quote":  "run_quote",
    "--esign":  "run_esign",
    "--search": "run_search",
    "--csp":    "run_csp",
}


def main() -> None:
    for flag, mod_name in _MODES.items():
        if flag in sys.argv:
            import importlib
            importlib.import_module(mod_name).main()
            return
    import main as launcher
    launcher.main()


if __name__ == "__main__":
    main()
