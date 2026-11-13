"""빌드 설정 검사 — PyInstaller 를 실제로 돌리지 않고(Windows 전용) 다음을 확인한다.
  · 만들어지는 exe 구성 (전체 / 앱 하나)
  · 앱마다: 넣겠다고 적은 우리 모듈이 실제로 있는지
  · 앱마다: 빼겠다고 적은 라이브러리·모듈을 그 앱의 코드가 실제로는 쓰고 있지 않은지
  · 앱마다: 이름으로만 불러오는 모듈(문자열)이 빠지지 않게 명시돼 있는지
"""
import ast
import os
import sys
import types

import pytest

NO_COM = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), "no_com")

import build_onedir


class _Rec:
    def __init__(self, kind, args, kw):
        self.kind, self.args, self.kw = kind, args, kw
        self.pure, self.scripts, self.binaries, self.datas = [], [], [], []


def _run_spec(only=""):
    """spec 을 가짜 PyInstaller 로 실행해 Analysis/EXE/COLLECT 호출을 받아 낸다."""
    calls = []

    def fake(kind):
        def f(*a, **k):
            r = _Rec(kind, a, k)
            calls.append(r)
            return r
        return f

    hooks = types.ModuleType("PyInstaller.utils.hooks")
    hooks.collect_all = lambda name: ([("d:" + name, ".")], [("b:" + name, ".")], ["h:" + name])
    saved = {k: sys.modules.get(k) for k in ("PyInstaller", "PyInstaller.utils", "PyInstaller.utils.hooks")}
    sys.modules["PyInstaller"] = types.ModuleType("PyInstaller")
    sys.modules["PyInstaller.utils"] = types.ModuleType("PyInstaller.utils")
    sys.modules["PyInstaller.utils.hooks"] = hooks
    try:
        g = {k: fake(k) for k in ("Analysis", "PYZ", "EXE", "COLLECT")}
        exec(compile(build_onedir._make_spec("견적자동화", "", only), "<spec>", "exec"), g)
    finally:
        for k, v in saved.items():
            if v is None:
                sys.modules.pop(k, None)
            else:
                sys.modules[k] = v
    return calls, g["_PROFILES"]


def test_full_build_makes_five_exes_in_one_folder():
    calls, profiles = _run_spec()
    exes = [c.kw["name"] for c in calls if c.kind == "EXE"]
    assert exes == ["견적자동화", "견적자동화_quote", "견적자동화_esign", "견적자동화_search", "견적자동화_csp"]
    assert [c.kw["name"] for c in calls if c.kind == "COLLECT"] == ["견적자동화"]
    assert [p["key"] for p in profiles] == list(build_onedir.APP_KEYS)
    assert build_onedir.output_names("견적자동화") == ("견적자동화", [e + ".exe" for e in exes])


@pytest.mark.parametrize("key", ["launcher", "quote", "esign", "search", "csp"])
def test_single_app_build_is_a_folder_of_its_own(key):
    calls, profiles = _run_spec(key)
    exes = [c.kw["name"] for c in calls if c.kind == "EXE"]
    want = "견적자동화" if key == "launcher" else f"견적자동화_{key}"
    assert exes == [want] and [c.kw["name"] for c in calls if c.kind == "COLLECT"] == [want]
    assert build_onedir.output_names("견적자동화", key) == (want, [want + ".exe"])


def test_command_line_parsing(monkeypatch):
    assert build_onedir._split_args(["이름", "icon.ico", "--app", "Quote"]) == (["이름", "icon.ico"], "quote")
    assert build_onedir._split_args(["이름", "--app=esign"]) == (["이름"], "esign")
    assert build_onedir._split_args(["이름", "--single"]) == (["이름"], "")       # 예전 옵션은 무시
    with pytest.raises(SystemExit):
        build_onedir._split_args(["이름", "--app", "없는앱"])


# ── 앱의 코드가 실제로 불러오는 것 ───────────────────────────────
def _local_modules():
    mods = {f[:-3] for f in os.listdir(NO_COM) if f.endswith(".py")}
    mods |= {d for d in os.listdir(NO_COM) if os.path.isfile(os.path.join(NO_COM, d, "__init__.py"))}
    return mods


def _module_path(name):
    parts = name.split(".")
    p = os.path.join(NO_COM, *parts)
    if os.path.isfile(p + ".py"):
        return p + ".py"
    if os.path.isfile(os.path.join(p, "__init__.py")):
        return os.path.join(p, "__init__.py")
    return None


def _scan_one(mod):
    """mod 한 개를 훑는다: (불러오는 우리 모듈, 바깥 라이브러리, 문자열로만 적힌 우리 모듈).

    함수 안의 import 도 본다. 다만 try 안에서 불러오는 것은 '있으면 쓰는 것'이라 세지 않는다
    (예: app_boot 의 오류 창 — Qt 앱은 Qt 로, Tk 앱은 tkinter 로 띄우며 실패하면 그냥 넘어간다).
    """
    local = _local_modules()
    path = _module_path(mod)
    direct, third, by_string = set(), set(), set()
    if path is None:
        return direct, third, by_string
    with open(path, encoding="utf-8") as f:
        tree = ast.parse(f.read())
    guarded = set()
    for node in ast.walk(tree):
        if isinstance(node, ast.Try):
            for stmt in node.body:
                guarded.update(id(n) for n in ast.walk(stmt))
    is_pkg = os.path.basename(path) == "__init__.py"
    base = mod if is_pkg else (mod.rsplit(".", 1)[0] if "." in mod else None)
    for node in ast.walk(tree):
        names = []
        if isinstance(node, ast.Import):
            names = [a.name for a in node.names]
        elif isinstance(node, ast.ImportFrom):
            if node.level and base:                                # 패키지 안의 상대 import
                names = [f"{base}.{node.module}"] if node.module else [f"{base}.{a.name}" for a in node.names]
            elif node.module:
                names = [node.module]
        elif (isinstance(node, ast.Constant) and isinstance(node.value, str)
              and node.value in local and node.value != mod):
            by_string.add(node.value); direct.add(node.value)
        for n in names:
            top = n.split(".")[0]
            if top in local:
                direct.add(n if _module_path(n) else top)
            elif id(node) not in guarded:
                third.add(top)
    return direct, third, by_string


def _closure(entry, stop=()):
    """entry 에서 닿는 우리 모듈 전부와, 그 코드가 불러오는 바깥 라이브러리."""
    seen, third, by_string, todo = set(), set(), set(), [entry]
    while todo:
        m = todo.pop()
        if m in seen or m in stop or _module_path(m) is None:
            continue
        seen.add(m)
        d, t3, bs = _scan_one(m)
        third |= t3; by_string |= bs - set(stop)
        todo.extend(d)
    return seen, third, by_string


@pytest.mark.parametrize("key", ["launcher", "quote", "esign", "search", "csp"])
def test_profile_matches_what_the_app_really_imports(key):
    _calls, profiles = _run_spec()
    prof = next(p for p in profiles if p["key"] == key)
    local = _local_modules()
    entry = prof["script"][:-3]
    assert _module_path(entry), prof["script"]
    # 런처는 다른 앱을 '파일 이름'으로 실행할 뿐 불러오지 않는다 → 그 스크립트는 따라가지 않는다
    stop = ("run_quote", "run_esign", "run_search", "run_csp") if key == "launcher" else ()
    seen, third, by_string = _closure(entry, stop)

    hidden = set(prof["hidden"])
    excludes = list(prof["excludes"])

    # 1) 넣겠다고 적은 우리 모듈은 실제로 있다
    for h in hidden:
        if h.split(".")[0] in local:
            assert _module_path(h), f"{key}: hidden 에 적힌 모듈이 없습니다 — {h}"

    # 2) 빼겠다고 적은 것을 실제로는 쓰고 있지 않다
    def excluded(name):
        return any(name == e or name.startswith(e + ".") for e in excludes)
    used_local = {m.split(".")[0] for m in seen}
    assert not [m for m in used_local if excluded(m)], f"{key}: 쓰는 모듈을 제외했습니다"
    bad = sorted(t for t in third if excluded(t))
    assert not bad, f"{key}: 쓰는 라이브러리를 제외했습니다 — {bad}"

    # 3) 이름(문자열)으로만 불러오는 우리 모듈은 PyInstaller 가 스스로 못 찾는다 → hidden 에 있어야 한다
    missing = sorted(m for m in by_string if m not in hidden)
    assert not missing, f"{key}: 이름으로 불러오는 모듈이 hidden 에 없습니다 — {missing}"

    # 4) 이 앱이 쓰는 우리 모듈과 바깥 라이브러리 (기록용 — 실패 시 원인 파악에 쓴다)
    assert entry in seen


def test_quote_app_no_longer_ships_unused_libraries():
    """견적서작성이 무거웠던 이유 — 쓰지 않는 PDF·그림·브라우저 자동화 라이브러리를 싣고 있었다."""
    calls, profiles = _run_spec("quote")
    a = next(c for c in calls if c.kind == "Analysis")
    for lib in ("fitz", "PIL", "selenium", "webdriver_manager", "pyautogui"):
        assert lib in a.kw["excludes"], lib
        assert not any(lib in str(x) for x in a.kw["hiddenimports"]), lib
        assert not any(("b:" + lib) == b[0] for b in a.kw["binaries"]), lib
    assert ("b:openpyxl", ".") in a.kw["binaries"]


def test_launcher_finds_apps_next_to_it_or_in_a_sibling_folder(tmp_path, monkeypatch):
    pytest.importorskip("PySide6.QtWidgets")
    os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")
    import main as launcher
    root = tmp_path / "dist"
    (root / "견적자동화").mkdir(parents=True)
    exe = root / "견적자동화" / "견적자동화.exe"
    exe.write_bytes(b"")
    monkeypatch.setattr(sys, "frozen", True, raising=False)
    monkeypatch.setattr(sys, "executable", str(exe))

    same = root / "견적자동화" / "견적자동화_quote.exe"
    same.write_bytes(b"")
    assert launcher._app_command("quote") == ([str(same)], str(same))          # 같은 폴더(전체 빌드)

    apart = root / "견적자동화_esign" / "견적자동화_esign.exe"
    apart.parent.mkdir(); apart.write_bytes(b"")
    assert launcher._app_command("esign") == ([str(apart)], str(apart))        # 옆 폴더(따로 빌드)

    cmd, must = launcher._app_command("csp")                                    # 없으면 없는 경로를 돌려준다
    assert not os.path.exists(must) and cmd == [must]
