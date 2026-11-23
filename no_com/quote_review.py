"""
quote_review.py
===============
"견적작성 대상"을 생성 전에 검토하기 위한 판단 — 금액, 칸 사용량, 점검 목록,
구성 도식에 무엇을 어떤 상태로 그릴지. 화면(Qt)·엑셀에 의존하지 않는다.

도식의 칸은 품목 시트의 '분류'로 정해진다(설정 diagram.flow / diagram.base).
"""

from __future__ import annotations

from dataclasses import dataclass, field
from typing import Any, Dict, List, Optional, Sequence, Tuple

import quote_calc
import quote_config
import review_rules

# 상태: ok 정상 / diff 수량이 CH 와 다름 / bad 단가 0원 / miss 코드매핑에 있는데 안 담김 / blank 없음
OK, DIFF, BAD, MISS, BLANK = "ok", "diff", "bad", "miss", "blank"
# 점검 등급: red 판단 필요 / yellow 확인 권장 / info 참고
RED, YELLOW, INFO = "red", "yellow", "info"


@dataclass
class Slot:
    label: str
    state: str = BLANK
    text: str = "없음"                 # 담긴 규격(여러 개면 ' + ' 로 이음) 또는 제안 규격
    note: str = ""                     # 수량·경고 한 줄
    specs: List[str] = field(default_factory=list)          # 눌렀을 때 선택할 품목
    suggest: Optional[Tuple[str, float]] = None             # (규격, 수량) — miss 일 때 눌러서 담기


@dataclass
class Check:
    level: str
    text: str
    spec: str = ""                                          # 눌렀을 때 선택할 품목
    suggest: Optional[Tuple[str, float]] = None             # 눌러서 담을 품목


@dataclass
class Review:
    qtype: str
    ch: Optional[float]                                     # 수량(CH). 의뢰 행이 없으면 None
    flow: List[Slot] = field(default_factory=list)
    base: List[Slot] = field(default_factory=list)
    others: List[Slot] = field(default_factory=list)        # 도식에 자리가 없는 품목
    checks: List[Check] = field(default_factory=list)
    amounts: Optional[quote_calc.DomesticAmounts] = None
    used: int = 0                                           # 품목 칸 사용 수
    capacity: int = 0                                       # 품목 칸 수

    @property
    def red(self) -> List[Check]:
        return [c for c in self.checks if c.level == RED]

    @property
    def yellow(self) -> List[Check]:
        return [c for c in self.checks if c.level == YELLOW]


def _fmt_qty(v: float) -> str:
    return str(int(v)) if float(v).is_integer() else ("%g" % v)


def _parse_slots(names: Sequence[str]) -> List[Tuple[str, str, bool]]:
    """['Gate Valve*', '@PUMP*'] → [(표시이름, 분류 또는 '@PUMP', CH 대조 여부)]"""
    out = []
    for raw in names:
        per_ch = raw.endswith("*")
        name = raw[:-1].strip() if per_ch else raw.strip()
        out.append(("PUMP" if name == "@PUMP" else name, name, per_ch))
    return out


def analyze(items: List[Dict[str, Any]],
            qtype: str = "국내",
            ch: Any = None,
            cfg: Optional[quote_config.QuoteConfig] = None,
            class_of: Optional[Dict[str, str]] = None,
            known_specs: Optional[set] = None,
            dup_price_specs: Optional[set] = None,
            suggested: Sequence[Tuple[str, float]] = (),
            rack_concept: Any = None,
            d5: Any = None,
            rules: Optional[review_rules.ReviewRules] = None) -> Review:
    """견적작성 대상(items)을 검토한다.

    ch              의뢰 행의 수량. 없으면 CH 대조를 하지 않는다.
    class_of        규격 → 품목 시트의 분류
    known_specs     품목 시트에 있는 규격 (없는 규격의 0원을 가려내는 데 쓴다)
    dup_price_specs 품목 시트에 단가가 둘 이상인 규격
    suggested       코드매핑이 이 조건에 제안하는 (규격, 수량)
    rack_concept    의뢰의 Rack Concept (예: '3CH')
    d5              5D 이름 (예: 'PUMP_DRY_LOT_XD1200_3RACK_…')
    rules           Rack·CH당 규칙 (통합양식의 검토규칙 시트, 없으면 기본 규칙)
    """
    cfg = cfg or quote_config.QuoteConfig()
    class_of = class_of or {}
    try:
        ch_val: Optional[float] = float(ch) if ch not in (None, "") else None
    except (TypeError, ValueError):
        ch_val = None
    if ch_val is not None and ch_val <= 0:
        ch_val = None

    rv = Review(qtype=qtype, ch=ch_val)
    real = [r for r in items if r.get("role") == "ITEM"]
    credits = [r for r in items if r.get("role") == "CREDIT"]
    pumps = [r for r in real if r.get("cat") == "PUMP"]
    racks = [r for r in real if r.get("cat") != "PUMP"]
    have = {r["spec"] for r in real}
    missing = [(sp, q) for sp, q in suggested if sp not in have]

    # ── 도식 ──────────────────────────────────
    placed: set = set()
    # 같은 분류라도 보조 품목(예: Safety 용 Gate Valve)은 CH 마다 하나씩 붙는 것이 아니다.
    # 칸에 넣으면 수량 대조가 어긋나므로 '기타 옵션'으로 보낸다.
    aux = [k.lower() for k in cfg.items("diagram.aux_keywords")]

    def is_aux(spec: str) -> bool:
        low = spec.lower()
        return any(k in low for k in aux)

    def slot(label: str, key: str, per_ch: bool) -> Slot:
        if key == "@PUMP":
            rows = pumps
        else:
            rows = [r for r in racks if class_of.get(r["spec"]) == key and not is_aux(r["spec"])]
        placed.update(id(r) for r in rows)
        if rows:
            qty = sum(r["qty"] for r in rows)
            s = Slot(label, OK, " + ".join(r["spec"] for r in rows), "수량 " + _fmt_qty(qty),
                     [r["spec"] for r in rows])
            if any(not r["price"] for r in rows):
                s.state, s.note = BAD, "단가 0원"
            elif per_ch and ch_val is not None and qty != ch_val:
                s.state, s.note = DIFF, f"수량 {_fmt_qty(qty)} ≠ {_fmt_qty(ch_val)}CH"
            return s
        sug = next(((sp, q) for sp, q in missing if class_of.get(sp) == key), None)
        if sug:
            return Slot(label, MISS, sug[0], "코드매핑에 있음 · 눌러서 담기", suggest=sug)
        return Slot(label)

    flow_def = _parse_slots(cfg.items("diagram.flow"))
    base_def = _parse_slots(cfg.items("diagram.base"))
    rv.flow = [slot(*d) for d in flow_def]
    rv.base = [slot(*d) for d in base_def]
    for r in racks:
        if id(r) not in placed:
            rv.others.append(Slot(class_of.get(r["spec"]) or "기타", OK if r["price"] else BAD,
                                  r["spec"], "수량 " + _fmt_qty(r["qty"]), [r["spec"]]))

    # ── 금액 · 칸 ─────────────────────────────
    if qtype == "국내":
        rv.amounts = quote_calc.domestic(items, ch_val or 1, cfg["credit.pump_keyword"])
        rv.capacity = cfg.capacity("dom.spec_start", "dom.spec_end")
        rv.used = len(racks) + len(credits)
        over = [("사양서 품목", rv.used, rv.capacity)]
    else:
        kw = cfg["credit.pump_keyword"]
        pump_cr = [c for c in credits if kw in c.get("spec", "")]
        if qtype == "중국":
            rv.capacity = cfg.capacity("cn.rack_start", "cn.rack_end")
            rv.used = len(racks) + len(credits) - len(pump_cr)
            over = [("중국 견적서 RACK", rv.used, rv.capacity),
                    ("중국 견적서 PUMP", len(pumps) + len(pump_cr),
                     cfg.capacity("cn.pump_start", "cn.pump_end"))]
        else:
            rv.capacity = cfg.capacity("us.rack_start", "us.rack_end")
            rv.used = len(racks) + len(credits)
            over = [("미국 견적서 RACK", rv.used, rv.capacity),
                    ("미국 견적서 PUMP", len(pumps), 1)]

    # ── 점검 목록 ─────────────────────────────
    add = rv.checks.append
    if not real:
        return rv                                   # 아직 아무것도 담지 않았다 — 지적할 것이 없다
    if not pumps:
        add(Check(RED, "PUMP 가 담기지 않았습니다"))
    for what, used, cap in over:
        if used > cap:
            add(Check(RED, f"{what} 칸 {cap}행 초과 ({used}행) — 이대로는 생성되지 않습니다"))
    for r in real:
        if not r["price"]:
            why = " — 품목 시트에 없는 이름" if known_specs is not None and r["spec"] not in known_specs else ""
            add(Check(RED, f"단가 0원: {r['spec']}{why}", spec=r["spec"]))
    for s in rv.flow + rv.base:
        if s.state == DIFF:
            add(Check(YELLOW, f"{s.label} {s.note}", spec=s.specs[0]))
    for r in real:
        if dup_price_specs and r["spec"] in dup_price_specs:
            add(Check(YELLOW, f"품목 시트에 단가가 둘 이상인 규격: {r['spec']}", spec=r["spec"]))
    _check_counts(rv, pumps, racks, ch_val, rack_concept, d5, rules or review_rules.defaults(), cfg)
    for sp, q in missing:
        add(Check(INFO, f"코드매핑에 있는데 안 담김: {sp} × {_fmt_qty(q)}", suggest=(sp, q)))
    return rv


def _check_counts(rv: Review, pumps: List[Dict[str, Any]], racks: List[Dict[str, Any]],
                  ch_val: Optional[float], rack_concept: Any, d5: Any,
                  rules: review_rules.ReviewRules, cfg: quote_config.QuoteConfig) -> None:
    """수량끼리 맞는지: PUMP·의뢰·Rack·Rack Concept·5D, CH당 품목, 같은 종류 중복."""
    add = rv.checks.append
    pump_qty = sum(r["qty"] for r in pumps)

    # ── Rack 대수 ──
    rack_ch, counted, unknown, by_n = 0, False, [], {}
    for r in racks:
        is_rack, n = rules.rack_ch(r["spec"])
        if not is_rack:
            continue
        if n is None:
            unknown.append(r)
            continue
        counted = True
        rack_ch += n * r["qty"]
        by_n.setdefault(n, []).append(r["spec"])
    for r in unknown:
        add(Check(YELLOW, f"확인 필요: {r['spec']} — 몇 CH 인지 알 수 없어 Rack 대수 계산에서 뺐습니다",
                  spec=r["spec"]))

    if pumps and ch_val is not None and pump_qty != ch_val:
        add(Check(RED, f"PUMP 수량 {_fmt_qty(pump_qty)}대 ≠ 의뢰 수량 {_fmt_qty(ch_val)}",
                  spec=pumps[0]["spec"]))
    if counted and pumps and rack_ch != pump_qty:
        add(Check(RED, f"Rack 합계 {_fmt_qty(rack_ch)}CH ≠ PUMP {_fmt_qty(pump_qty)}대 "
                       f"(Rack System 등 이름의 숫자 × 수량)", spec=next(iter(by_n.values()))[0]))
    concept = review_rules.ch_from_concept(rack_concept)
    if counted and concept is not None and rack_ch != concept:
        add(Check(YELLOW, f"Rack 합계 {_fmt_qty(rack_ch)}CH ≠ 의뢰의 Rack Concept {concept}CH"))
    d5_ch = review_rules.ch_from_5d(d5)
    if counted and d5_ch and rack_ch % d5_ch != 0:
        add(Check(YELLOW, f"Rack 합계 {_fmt_qty(rack_ch)}CH 가 5D 의 {d5_ch}CH 의 배수가 아닙니다"))

    # 같은 CH 수의 Rack 이 이름만 다르게 둘 (예: Rack System-3 과 Rack System-3 (전기공통 사양))
    for n, specs in by_n.items():
        if len(set(specs)) > 1:
            add(Check(YELLOW, f"같은 {n}CH Rack 이 두 가지 담겼습니다: " + ", ".join(sorted(set(specs))),
                      spec=specs[0]))

    # ── CH 마다 들어가는 품목 ──
    base = ch_val if ch_val is not None else (pump_qty or None)
    if base:
        seen_rule: Dict[str, float] = {}
        first: Dict[str, str] = {}
        rule_of: Dict[str, review_rules.PerChRule] = {}
        for r in racks:
            rule = rules.per_ch_rule(r["spec"])
            if rule is None:
                continue
            seen_rule[rule.pattern] = seen_rule.get(rule.pattern, 0) + r["qty"]
            first.setdefault(rule.pattern, r["spec"]); rule_of[rule.pattern] = rule
        for pat, qty in seen_rule.items():
            rule = rule_of[pat]
            want = rule.per_ch * base
            if qty != want:
                add(Check(RED if rule.level == "red" else YELLOW,
                          f"{first[pat]} 수량 {_fmt_qty(qty)} ≠ {_fmt_qty(want)} "
                          f"(CH 마다 {_fmt_qty(rule.per_ch)}개 × {_fmt_qty(base)}CH)", spec=first[pat]))

    # ── 같은 종류가 두 가지 (도식 칸 기준, 보조 품목 제외) ──
    for s_ in rv.flow + rv.base:
        names = sorted(set(s_.specs))
        if len(names) > 1 and s_.label != "PUMP" and not any(rules.rack_ch(n)[0] for n in names):
            add(Check(YELLOW, f"{s_.label} 이 두 가지 이상 담겼습니다: " + ", ".join(names), spec=names[0]))
