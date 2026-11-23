"""
review_rules.py
===============
'생성 전 점검'의 규칙 — 어떤 품목이 Rack 이고 몇 CH 분인지, 어떤 품목이 CH 마다
몇 개씩 들어가야 하는지. 화면(Qt)에 의존하지 않는다.

값의 출처
---------
1. 이 파일의 기본 규칙 (통합양식 ver.2.1 의 품목·코드매핑 기준)
2. 통합양식 안의 `검토규칙` 시트 — 있으면 기본 규칙 대신 그 시트의 규칙을 쓴다
   (시트에서 지운 규칙은 적용되지 않는다)

`검토규칙` 시트 모양 (1행 머리글):
    구분      | 품목 이름에 포함          | 값   | 등급
    Rack      | Rack System-{n}          | n    |            ← 이름의 숫자만큼 CH
    Rack      | Rack (3단 Booster용)     | 확인 |            ← Rack 이지만 CH 를 모름: 대수 계산에서 빼고 '확인 필요'
    CH당      | Cable Tray Ass'y         | 1    | 확인 권장  ← CH 마다 1개
  · '품목 이름에 포함' 의 {n} 은 숫자 자리. 대소문자·공백 차이는 무시한다.
  · 등급: 판단 필요 / 확인 권장 (비우면 확인 권장)

CH당 기본 규칙은 코드매핑 227건 중 거의 예외 없이(94% 이상) 지켜진 것만 넣었다.
Fitting & Hose·Clamp 처럼 CH당 개수가 여러 가지인 품목은 넣지 않았다(오경보가 많다).
Gate Valve·Metal Bellows·Hot N2 는 구성 도식이 이미 CH 와 대조하므로 넣지 않았다.
"""

from __future__ import annotations

import re
from dataclasses import dataclass, field
from typing import Any, Iterable, List, Optional, Tuple

RULES_SHEET = "검토규칙"

RED_WORD, YELLOW_WORD = "판단 필요", "확인 권장"


@dataclass(frozen=True)
class RackRule:
    pattern: str                 # 품목 이름에 포함되는 말 ({n} = 숫자 자리)
    ch: Optional[str]            # "n" = 이름의 숫자, "3" 같은 고정값, None = 모름(확인 필요)


@dataclass(frozen=True)
class PerChRule:
    pattern: str
    per_ch: float
    level: str = "yellow"        # red / yellow


@dataclass
class ReviewRules:
    racks: List[RackRule] = field(default_factory=list)
    per_ch: List[PerChRule] = field(default_factory=list)
    problems: List[str] = field(default_factory=list)
    from_sheet: bool = False

    # ── 판정 ────────────────────────────────────
    def rack_ch(self, spec: str) -> Tuple[bool, Optional[int]]:
        """(Rack 인가, 몇 CH 인가 — 모르면 None)."""
        for rule in self.racks:
            m = _match(rule.pattern, spec)
            if m is None:
                continue
            if rule.ch is None:
                return True, None
            if rule.ch == "n":
                return True, int(m) if m is not None and m != "" else None
            return True, int(float(rule.ch))
        return False, None

    def per_ch_rule(self, spec: str) -> Optional[PerChRule]:
        for rule in self.per_ch:
            if _match(rule.pattern, spec) is not None:
                return rule
        return None


DEFAULT_RACKS = [
    RackRule("Rack System-{n}", "n"),
    RackRule("격자보 Rack-{n}", "n"),
    RackRule("Support -{n}", "n"),
    RackRule("Rack (3단 Booster용)", None),
    RackRule("Compressor Rack(1단)", None),
    RackRule("Compressor Rack(2단)", None),
]
DEFAULT_PER_CH = [
    PerChRule("Cable Tray Ass'y", 1),
    PerChRule("Dual seal O-ring  (NW40)", 1),
    PerChRule("Device Net Interface Module", 1),
    PerChRule("외장형 Silencer (방열판 Type)", 1),
    PerChRule("RACK 표준화_1CH 당", 1),
]


def defaults() -> ReviewRules:
    return ReviewRules(list(DEFAULT_RACKS), list(DEFAULT_PER_CH))


def _norm(text: str) -> str:
    return re.sub(r"\s+", "", str(text or "")).lower()


_pattern_cache: dict = {}


def _match(pattern: str, spec: str) -> Optional[str]:
    """pattern 이 spec 에 들어 있으면 {n} 자리의 숫자(없으면 '')를, 아니면 None."""
    rx = _pattern_cache.get(pattern)
    if rx is None:
        body = re.escape(_norm(pattern)).replace(re.escape("{n}"), r"(\d+)")
        rx = _pattern_cache[pattern] = re.compile(body)
    m = rx.search(_norm(spec))
    if m is None:
        return None
    return m.group(1) if m.groups() else ""


# ──────────────────────────────────────────────
# 읽기 · 쓰기
# ──────────────────────────────────────────────
def read_rows(rows: Iterable[tuple]) -> ReviewRules:
    """`검토규칙` 시트의 행(값만)에서 규칙을 읽는다. 잘못된 줄은 건너뛰고 사유를 남긴다."""
    out = ReviewRules(from_sheet=True)
    for n, row in enumerate(rows, start=1):
        if not row or all(v is None or str(v).strip() == "" for v in row):
            continue
        kind = str(row[0] or "").strip()
        pattern = str(row[1] or "").strip() if len(row) > 1 else ""
        value = row[2] if len(row) > 2 else None
        level = str(row[3] or "").strip() if len(row) > 3 else ""
        if n == 1 and kind in ("구분", "종류"):
            continue
        if kind.startswith("#") or not kind:
            continue
        if not pattern:
            out.problems.append(f"검토규칙 {n}행: '품목 이름에 포함' 이 비어 있어 건너뜁니다")
            continue
        val = str(value if value is not None else "").strip()
        if kind.lower() == "rack":
            if val.lower() == "n":
                if "{n}" not in pattern:
                    out.problems.append(f"검토규칙 {n}행: 값이 n 인데 이름에 {{n}} 자리가 없습니다")
                    continue
                out.racks.append(RackRule(pattern, "n"))
            elif val in ("", "확인", "확인필요", "확인 필요", "?"):
                out.racks.append(RackRule(pattern, None))
            else:
                try:
                    out.racks.append(RackRule(pattern, str(int(float(val)))))
                except ValueError:
                    out.problems.append(f"검토규칙 {n}행: Rack 의 값 '{val}' 을 읽을 수 없습니다 (n, 숫자, 확인)")
        elif kind.replace(" ", "") in ("CH당", "ch당"):
            try:
                per = float(val)
            except ValueError:
                out.problems.append(f"검토규칙 {n}행: CH당 개수 '{val}' 가 숫자가 아닙니다")
                continue
            out.per_ch.append(PerChRule(pattern, per, "red" if level == RED_WORD else "yellow"))
        else:
            out.problems.append(f"검토규칙 {n}행: 구분 '{kind}' 을 모릅니다 (Rack 또는 CH당)")
    return out


def from_workbook(wb) -> ReviewRules:
    """열려 있는 통합양식에서 규칙을 읽는다. `검토규칙` 시트가 없으면 기본 규칙."""
    if RULES_SHEET not in wb.sheetnames:
        return defaults()
    return read_rows(wb[RULES_SHEET].iter_rows(max_col=4, values_only=True))


_cache: dict = {}


def load(template_path: Optional[str]) -> ReviewRules:
    """통합양식 파일에서 규칙을 읽는다(파일이 바뀌기 전까지 기억)."""
    import os
    if not template_path:
        return defaults()
    try:
        st = os.stat(template_path)
    except OSError:
        return defaults()
    stamp = (st.st_mtime, st.st_size)
    hit = _cache.get(template_path)
    if hit and hit[0] == stamp:
        return hit[1]
    from openpyxl import load_workbook
    try:
        wb = load_workbook(template_path, read_only=True, data_only=True)
    except Exception:
        return defaults()
    try:
        rules = from_workbook(wb)
    finally:
        wb.close()
    _cache[template_path] = (stamp, rules)
    return rules


def write_sheet(wb, rules: Optional[ReviewRules] = None) -> None:
    """wb 에 지금 규칙을 적은 `검토규칙` 시트를 만든다."""
    from openpyxl.styles import Font
    rules = rules or defaults()
    ws = wb.create_sheet(RULES_SHEET)
    ws.append(["구분", "품목 이름에 포함", "값", "등급", "설명"])
    for c in ws[1]:
        c.font = Font(bold=True)
    for r in rules.racks:
        why = ("이름의 숫자만큼 CH" if r.ch == "n" else
               "CH 를 알 수 없음 — 대수 계산에서 빼고 '확인 필요' 로 표시" if r.ch is None else f"{r.ch} CH")
        ws.append(["Rack", r.pattern, r.ch or "확인", "", why])
    for r in rules.per_ch:
        per = int(r.per_ch) if float(r.per_ch).is_integer() else r.per_ch
        ws.append(["CH당", r.pattern, per, RED_WORD if r.level == "red" else YELLOW_WORD,
                   f"CH 마다 {per}개"])
    ws.column_dimensions["A"].width = 8
    ws.column_dimensions["B"].width = 36
    ws.column_dimensions["C"].width = 8
    ws.column_dimensions["D"].width = 10
    ws.column_dimensions["E"].width = 48
    ws.freeze_panes = "A2"


# ──────────────────────────────────────────────
# 글자에서 CH 수 읽기
# ──────────────────────────────────────────────
def ch_from_concept(text: Any) -> Optional[int]:
    """의뢰의 Rack Concept('3CH', '3CH+2CH') → CH 합. 숫자+CH 형식이 아니면 None."""
    found = re.findall(r"(\d+)\s*CH", str(text or ""), flags=re.IGNORECASE)
    return sum(int(x) for x in found) if found else None


def ch_from_5d(name: Any) -> Optional[int]:
    """5D 이름 → CH 수. 'nRACK' 이 n CH, '+' 는 더하고, 끝의 '*k' 는 k 배.
    예) …_3RACK_…+…_3RACK_… → 6,  …_3RACK_…*2 → 6.  읽을 수 없으면 None."""
    text = str(name or "").strip()
    if not text:
        return None
    total = 0
    for part in text.split("+"):
        mult = 1
        m = re.search(r"\*\s*(\d+)\s*$", part)
        if m:
            mult = int(m.group(1)); part = part[:m.start()]
        found = re.findall(r"(?<![0-9])(\d{1,2})\s*RACK", part, flags=re.IGNORECASE)
        if len(found) != 1:
            return None                       # 숫자가 없거나 여러 개(뜻이 모호) — 읽지 않는다
        total += int(found[0]) * mult
    return total or None
