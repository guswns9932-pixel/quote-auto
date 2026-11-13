"""
quote_calc.py
=============
국내 견적서의 금액 계산 — 화면(미리 보여주는 금액)과 파일 생성이 같은 식을 쓰도록
한곳에 둔다. 엑셀·Qt 에 의존하지 않는다.

통합양식의 수식을 그대로 따른다:
    사양서!G15 = E15*F15                 PUMP 총액 (수량 × CH당 단가)
    사양서!G16 = SUM(G17:G41)            Rack system 합계
    사양서!F43 = ROUNDDOWN((G16+G15)/수량, -3)                견적단가
    견적의뢰복사본!T2 = S2*H2            견적금액
    견적의뢰복사본!U2 = ROUNDUP((P2*H2+Q2)/H2, -3)            견적단가(Check)

[Pump Credit]
Pump Credit 을 Rack 합계(G16)에 넣으면 PUMP 값이어야 할 금액이 Rack 쪽에서 빠지는
것처럼 보인다. 그래서 PUMP 는 credit 을 반영한 CH당 단가를 F15 에 직접 적어 G15 에서
차감되게 하고, Credit 줄은 표시만 한다(금액 0). Rack Credit 은 G16 에서 차감한다.
"""

from __future__ import annotations

import math
from dataclasses import dataclass, field
from typing import Any, Dict, List, Optional


@dataclass
class DomesticAmounts:
    qty: float                       # 수량(CH)
    pump_items: List[Dict[str, Any]] = field(default_factory=list)
    rack_items: List[Dict[str, Any]] = field(default_factory=list)
    pump_credits: List[Dict[str, Any]] = field(default_factory=list)
    rack_credits: List[Dict[str, Any]] = field(default_factory=list)
    pump_spec: Optional[str] = None
    pump_amt: float = 0.0            # PUMP 금액 합
    rack_amt: float = 0.0            # RACK·옵션 금액 합
    pump_credit_amt: float = 0.0     # 통상 음수
    rack_credit_amt: float = 0.0     # 통상 음수
    p_unit_net: int = 0              # Pump Credit 을 반영한 CH당 PUMP 단가
    pump_credit_unit: int = 0        # 표시용 CH당 Pump Credit (음수)
    p_unit: float = 0                # 견적의뢰복사본 P (=사양서 F15)
    g15: float = 0.0                 # PUMP 총액
    g16: float = 0.0                 # Rack system 합계 (Rack Credit 반영)
    unit: int = 0                    # 견적단가 S (천 원 미만 버림)
    total: float = 0.0               # 견적금액 T
    unit_check: int = 0              # 견적단가(Check) U (천 원 미만 올림)

    @property
    def credits(self) -> List[Dict[str, Any]]:
        return self.pump_credits + self.rack_credits


def domestic(items: List[Dict[str, Any]], qty: Any, pump_keyword: str = "Pump") -> DomesticAmounts:
    """items(견적작성 대상)와 수량으로 국내 견적서 금액을 계산한다.

    items 의 각 항목: {"cat": "PUMP"|…, "role": "ITEM"|"CREDIT", "spec", "qty", "price", "amt"}
    qty 가 0 이하이거나 숫자가 아니면 1 로 본다.
    """
    try:
        h = float(qty)
    except (TypeError, ValueError):
        h = 0.0
    if not (h > 0):
        h = 1.0

    a = DomesticAmounts(qty=h)
    credits = [r for r in items if r["role"] == "CREDIT"]
    a.pump_credits = [r for r in credits if pump_keyword in r.get("spec", "")]
    a.rack_credits = [r for r in credits if pump_keyword not in r.get("spec", "")]
    a.pump_items = [r for r in items if r["cat"] == "PUMP" and r["role"] != "CREDIT"]
    a.rack_items = [r for r in items if r["cat"] != "PUMP" and r["role"] != "CREDIT"]
    a.pump_spec = a.pump_items[0]["spec"] if a.pump_items else None
    a.pump_amt = sum(r["amt"] for r in a.pump_items)
    a.rack_amt = sum(r["amt"] for r in a.rack_items)
    a.pump_credit_amt = sum(c["amt"] for c in a.pump_credits)
    a.rack_credit_amt = sum(c["amt"] for c in a.rack_credits)

    # PUMP CH당 단가 — credit 반영 후 정수(원)로 확정한다. 소수점이 남으면
    # 견적서에 13,900,000.333 같은 값이 찍히고 G15 도 원 단위로 떨어지지 않는다.
    a.p_unit_net = round((a.pump_amt + a.pump_credit_amt) / h)
    a.pump_credit_unit = round(a.pump_credit_amt / h) if a.pump_credits else 0

    # ※ p_unit_net 은 이미 정수로 반올림했으므로 g15 도 그 값에서 직접 만든다.
    #   (A/h)*h 식으로 되돌리면 부동소수점 오차로 버림이 1,000원 낮게 떨어질 수 있다.
    if a.pump_credits:
        a.p_unit = a.p_unit_net
        a.g15 = a.p_unit_net * h
    else:
        a.p_unit = round(a.pump_amt / h)
        a.g15 = a.pump_amt
    a.g16 = a.rack_amt + a.rack_credit_amt

    a.unit = math.floor((a.g15 + a.g16) / h / 1000) * 1000
    a.total = a.unit * h
    a.unit_check = math.ceil((a.p_unit * h + a.g16) / h / 1000) * 1000
    return a
