"""
formula_eval.py
===============
엑셀 수식의 값을 파일만 보고 계산한다 (Excel 없이).

openpyxl 로 저장한 파일에는 수식의 계산 결과가 들어 있지 않다. Excel 에서 한 번
열어 저장해야 생긴다. 그래서 이 프로그램이 만든 견적서에서 수식 칸(예: 중국
견적서의 단가)을 읽으면 빈 값이 나온다. Excel 을 띄워 읽을 수도 있지만 느리고
(이 환경에서는 종료에만 수십 초) 무겁다.

여기서는 그 칸의 수식을 따라가며 직접 계산한다. 수식을 코드에 옮겨 적지 않고
파일에 적힌 수식을 읽으므로, 양식의 수식이 바뀌어도 따라간다.

다루는 것: 숫자, 같은 시트의 셀·범위, + - * / ^, 괄호, %, 비교는 하지 않음.
함수: SUM, ROUND, ROUNDDOWN, ROUNDUP, ABS, MIN, MAX, AVERAGE, INT.
그 밖의 것(다른 시트 참조, 문자열 연산, IF 등)을 만나면 FormulaError 를 낸다 —
틀린 값을 조용히 내는 것보다 못 읽었다고 알리는 편이 낫다.
"""

from __future__ import annotations

import math
import re
from typing import Any, List, Optional, Set

from openpyxl.utils import column_index_from_string, get_column_letter


class FormulaError(ValueError):
    """수식을 계산할 수 없다."""


_TOKEN = re.compile(r"""
      (?P<num>   \d+\.?\d*(?:[eE][+-]?\d+)? | \.\d+ )
    | (?P<range> \$?[A-Za-z]{1,3}\$?\d+ : \$?[A-Za-z]{1,3}\$?\d+ )
    | (?P<cell>  \$?[A-Za-z]{1,3}\$?\d+ (?![A-Za-z0-9_(]) )
    | (?P<func>  [A-Za-z_][A-Za-z0-9_.]* (?=\s*\() )
    | (?P<op>    [-+*/^(),%] )
    | (?P<ws>    \s+ )
""", re.VERBOSE)


def _tokenize(text: str) -> List[tuple]:
    out, pos = [], 0
    while pos < len(text):
        m = _TOKEN.match(text, pos)
        if not m:
            raise FormulaError(f"읽을 수 없는 수식입니다: …{text[pos:pos + 20]}")
        pos = m.end()
        kind = m.lastgroup
        if kind != "ws":
            out.append((kind, m.group(kind)))
    return out


def _round_half_away(x: float, digits: int) -> float:
    """엑셀의 ROUND — 5 는 0 에서 먼 쪽으로."""
    k = 10.0 ** digits
    return math.copysign(math.floor(abs(x) * k + 0.5), x) / k


def _round_dir(x: float, digits: int, up: bool) -> float:
    k = 10.0 ** digits
    v = abs(x) * k
    # 0.1+0.2 같은 부동소수점 꼬리 때문에 딱 떨어지는 값이 한 칸 밀리지 않게 한다
    r = round(v)
    if abs(v - r) < 1e-9:
        v = r
    v = math.ceil(v) if up else math.floor(v)
    return math.copysign(v, x) / k


class _Parser:
    def __init__(self, ws, tokens: List[tuple], seen: Set[str]) -> None:
        self.ws, self.toks, self.i, self.seen = ws, tokens, 0, seen

    # ── 토큰 ──────────────────────────────────
    def peek(self) -> Optional[tuple]:
        return self.toks[self.i] if self.i < len(self.toks) else None

    def take(self, value: Optional[str] = None) -> tuple:
        tok = self.peek()
        if tok is None or (value is not None and tok[1] != value):
            raise FormulaError(f"수식 모양이 맞지 않습니다 ('{value}' 가 있어야 할 자리)")
        self.i += 1
        return tok

    # ── 문법: expr = term (('+'|'-') term)* … ──
    def expr(self) -> float:
        v = self.term()
        while self.peek() and self.peek()[1] in "+-":
            op = self.take()[1]
            r = self.term()
            v = v + r if op == "+" else v - r
        return v

    def term(self) -> float:
        v = self.power()
        while self.peek() and self.peek()[1] in "*/":
            op = self.take()[1]
            r = self.power()
            if op == "/":
                if r == 0:
                    raise FormulaError("0 으로 나누는 수식입니다 (#DIV/0!)")
                v = v / r
            else:
                v = v * r
        return v

    def power(self) -> float:
        v = self.unary()
        if self.peek() and self.peek()[1] == "^":
            self.take()
            v = v ** self.power()
        return v

    def unary(self) -> float:
        tok = self.peek()
        if tok and tok[1] in "+-":
            self.take()
            v = self.unary()
            return -v if tok[1] == "-" else v
        v = self.atom()
        while self.peek() and self.peek()[1] == "%":
            self.take()
            v = v / 100.0
        return v

    def atom(self) -> float:
        tok = self.peek()
        if tok is None:
            raise FormulaError("수식이 중간에 끝났습니다")
        kind, text = tok
        if kind == "num":
            self.take()
            return float(text)
        if kind == "cell":
            self.take()
            return _cell_number(self.ws, text.replace("$", "").upper(), self.seen)
        if kind == "range":
            raise FormulaError(f"범위({text})는 SUM 같은 함수 안에서만 쓸 수 있습니다")
        if kind == "func":
            return self.call()
        if text == "(":
            self.take()
            v = self.expr()
            self.take(")")
            return v
        raise FormulaError(f"다룰 수 없는 수식입니다: '{text}'")

    def args(self) -> List[float]:
        """함수 인자 — 범위는 그 안의 숫자들로 펼친다."""
        out: List[float] = []
        self.take("(")
        if self.peek() and self.peek()[1] == ")":
            self.take()
            return out
        while True:
            tok = self.peek()
            if tok and tok[0] == "range":
                self.take()
                out.extend(_range_numbers(self.ws, tok[1], self.seen))
            else:
                out.append(self.expr())
            if self.peek() and self.peek()[1] == ",":
                self.take()
                continue
            self.take(")")
            return out

    def call(self) -> float:
        name = self.take()[1].upper()
        a = self.args()

        def need(n: int) -> None:
            if len(a) != n:
                raise FormulaError(f"{name} 의 인자 수가 맞지 않습니다")

        if name == "SUM":
            return float(sum(a))
        if name in ("ROUND", "ROUNDDOWN", "ROUNDUP"):
            need(2)
            d = int(a[1])
            if name == "ROUND":
                return _round_half_away(a[0], d)
            return _round_dir(a[0], d, up=(name == "ROUNDUP"))
        if name == "ABS":
            need(1)
            return abs(a[0])
        if name == "INT":
            need(1)
            return float(math.floor(a[0]))
        if name in ("MIN", "MAX", "AVERAGE"):
            if not a:
                raise FormulaError(f"{name} 에 값이 없습니다")
            return float(min(a) if name == "MIN" else max(a) if name == "MAX" else sum(a) / len(a))
        raise FormulaError(f"다루지 않는 함수입니다: {name}")


def _range_numbers(ws, text: str, seen: Set[str]) -> List[float]:
    a, b = text.replace("$", "").upper().split(":")
    ma, mb = re.fullmatch(r"([A-Z]+)(\d+)", a), re.fullmatch(r"([A-Z]+)(\d+)", b)
    c1, c2 = sorted((column_index_from_string(ma.group(1)), column_index_from_string(mb.group(1))))
    r1, r2 = sorted((int(ma.group(2)), int(mb.group(2))))
    out: List[float] = []
    for r in range(r1, r2 + 1):
        for c in range(c1, c2 + 1):
            ref = f"{get_column_letter(c)}{r}"
            v = ws[ref].value
            if v is None or (isinstance(v, str) and not v.startswith("=")):
                continue                      # SUM 은 빈칸과 글자를 건너뛴다
            out.append(_cell_number(ws, ref, seen))
    return out


def _cell_number(ws, ref: str, seen: Set[str]) -> float:
    v = ws[ref].value
    if v is None or v == "":
        return 0.0                            # 빈칸은 계산에서 0
    if isinstance(v, bool):
        return float(v)
    if isinstance(v, (int, float)):
        return float(v)
    if isinstance(v, str) and v.startswith("="):
        if ref in seen:
            raise FormulaError(f"수식이 자기 자신을 다시 가리킵니다 ({ref})")
        return _eval_text(ws, v, seen | {ref})
    try:
        return float(str(v).replace(",", ""))
    except ValueError:
        raise FormulaError(f"{ref} 칸이 숫자가 아닙니다: '{v}'")


def _eval_text(ws, formula: str, seen: Set[str]) -> float:
    text = formula.strip()
    if text.startswith("="):
        text = text[1:]
    if "!" in text:
        raise FormulaError("다른 시트를 가리키는 수식은 계산하지 않습니다")
    if '"' in text or "&" in text:
        raise FormulaError("글자를 다루는 수식은 계산하지 않습니다")
    p = _Parser(ws, _tokenize(text), seen)
    v = p.expr()
    if p.peek() is not None:
        raise FormulaError(f"수식 뒤에 남는 것이 있습니다: '{p.peek()[1]}'")
    return v


def cell_value(ws, ref: str) -> Any:
    """ref 칸의 값. 수식이면 계산한 숫자를, 아니면 적힌 값을 그대로 돌려준다.
    ws 는 수식이 살아 있는(data_only 가 아닌) 워크시트여야 한다."""
    ref = ref.replace("$", "").upper()
    v = ws[ref].value
    if isinstance(v, str) and v.startswith("="):
        return _eval_text(ws, v, {ref})
    return v
