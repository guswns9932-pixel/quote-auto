"""
app_main.py
===========
기본 빌드(build_onedir.py <이름>)의 진입점 — exe 하나에 런처와 앱 전부가 들어간다.

exe 하나(해시 하나)로 런처와 2개 앱을 모두 실행한다. 런처가 이 exe 를
`--quote` 등의 인자로 다시 실행하면 해당 앱만 띄운다. [중요] 인자를 보기 전에
PySide6 같은 무거운 모듈을 임포트하지 않는다 — 앱마다 필요한 모듈만 불러와야
단일 exe 방식의 이점이 산다.
"""

from __future__ import annotations

import sys

# 인자 → run_*.py 모듈. 임포트는 분기 이후에 한다.
_MODES = {
    "--quote":  "run_quote",
    "--esign":  "run_esign",
}


def main() -> None:
    for flag, mod_name in _MODES.items():
        if flag in sys.argv:
            import importlib
            importlib.import_module(mod_name).main()
            return
    import main as launcher
    launcher.ALL_IN_ONE = True      # 앱을 띄울 때 이 exe 를 --<앱> 으로 다시 실행하게 한다
    launcher.main()


if __name__ == "__main__":
    main()
