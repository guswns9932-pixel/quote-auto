"""
quote_config.py
===============
견적서작성의 "양식에 따라 달라지는 값" — 시트 이름, 셀 위치, 열 문자, 행 범위,
문구, 파일 이름 규칙 — 을 한곳에 모은다. 화면(Qt)에 의존하지 않는다.

값의 출처
---------
1. 이 파일의 기본값 (통합양식 ver.2.0 기준)
2. 통합양식 안의 `설정` 시트 — 있으면 거기 적힌 항목만 기본값을 덮어쓴다

양식의 칸이 옮겨지거나 시트 이름이 바뀌면 코드를 고치지 않고 `설정` 시트만
고치면 된다. `설정` 시트가 없는 양식은 기본값 그대로 동작한다.

`설정` 시트 모양: A열 = 항목(아래 표의 키), B열 = 값, C열 = 설명(읽지 않음).
1행은 머리글. export_settings_sheet() 가 지금 값으로 채운 시트를 만들어 준다.

계산 규칙(Credit 처리, 반올림 등)은 여기 두지 않는다 — 그건 코드의 일이다.
"""

from __future__ import annotations

import logging
import os
import re
from typing import Any, Dict, Iterable, List, Optional, Tuple

logger = logging.getLogger("QuoteApp")

SETTINGS_SHEET = "설정"

# (키, 기본값, 종류, 설명)
#   종류: sheet=시트 이름, cell=셀 주소, col=열 문자, cols=열 범위 목록,
#         row=행 번호, int=정수, text=문구, name=이름 규칙({자리표시} 사용)
_SPEC: List[Tuple[str, Any, str, str]] = [
    # ── 시트 이름 ──────────────────────────────────────────────
    ("sheet.items",      "품목",              "sheet", "품목(분류·규격·단가) 시트"),
    ("sheet.code_map",   "코드매핑",           "sheet", "코드매핑 시트"),
    ("sheet.req_copy",   "견적의뢰복사본",      "sheet", "의뢰 행을 옮겨 적는 시트"),
    ("sheet.spec",       "사양서",             "sheet", "사양서 시트"),
    ("sheet.sign_spec",  "현업 사인용 사양서",  "sheet", "현업 사인용 사양서 시트"),
    ("sheet.incoming",   "입고검수확인서",      "sheet", "입고검수확인서 시트"),
    ("sheet.quote_cn",   "견적서_중국",        "sheet", "중국 견적서 시트"),
    ("sheet.quote_us",   "견적서_미국",        "sheet", "미국 견적서 시트"),
    ("sheet.cover_data", "갑지DATA",           "sheet", "갑지 원본 데이터 시트"),
    ("sheet.cover",      "견적서 갑지",         "sheet", "갑지 시트"),
    ("sheet.pump_price", "Pump 단가표",        "sheet", "PUMP 단가표 시트"),
    ("sheet.acc_price",  "악세서리 단가표",     "sheet", "악세서리 단가표 시트"),

    # ── 품목 시트 ─────────────────────────────────────────────
    ("items.col_class", "A", "col", "품목: 분류 열"),
    ("items.col_spec",  "B", "col", "품목: 규격 열"),
    ("items.col_price", "C", "col", "품목: 단가 열"),

    # ── 코드매핑 시트 ──────────────────────────────────────────
    ("codemap.col_process", "A", "col", "코드매핑: 공정 열"),
    ("codemap.col_vendor",  "B", "col", "코드매핑: 설비사 열"),
    ("codemap.col_code5d",  "C", "col", "코드매핑: 5D 열"),
    ("codemap.col_acc",     "D", "col", "코드매핑: ACC(규격) 열"),
    ("codemap.col_qty",     "E", "col", "코드매핑: 수량 열"),

    # ── 의뢰파일 / 견적의뢰복사본 (같은 열 구성) ─────────────────
    ("req.col_pr",        "D",  "col", "의뢰: Purchase Requisition 열"),
    ("req.col_item",      "E",  "col", "의뢰: Item of Requisition 열"),
    ("req.col_qcode",     "F",  "col", "의뢰: Material(Q-code) 열"),
    ("req.col_desc",      "G",  "col", "의뢰: Material Description 열"),
    ("req.col_qty",       "H",  "col", "의뢰: 수량 열"),
    ("req.col_requester", "J",  "col", "의뢰: 의뢰자 열"),
    ("req.col_line",      "K",  "col", "의뢰: 라인 열"),
    ("req.col_subproc",   "N",  "col", "의뢰: 세부공정 열"),
    ("req.col_concept",   "R",  "col", "의뢰: Rack Concept 열"),
    ("req.col_5d",        "V",  "col", "의뢰: 비고(5D명) 열"),
    ("req.col_maker",     "X",  "col", "의뢰: 설비Maker 열"),
    ("req.col_model",     "Y",  "col", "의뢰: 설비Model 열"),
    ("req.col_hogi",      "Z",  "col", "의뢰: 설비호기 열 (비어 있으면 그 행은 건너뜀)"),
    ("req.col_due",       "AA", "col", "의뢰: 희망 납품일 열"),
    ("copy.row",            2,   "row", "견적의뢰복사본: 의뢰 행을 적는 행"),
    ("copy.cols",  "A:N, R, W:AA", "cols", "견적의뢰복사본: 의뢰파일에서 그대로 옮기는 열"),
    ("copy.col_pump_model", "O", "col", "견적의뢰복사본: PUMP 메인모델 열"),
    ("copy.col_pump_price", "P", "col", "견적의뢰복사본: PUMP 단가 열"),
    ("copy.col_rack_sum",   "Q", "col", "견적의뢰복사본: Rack+액세서리 합계 열"),
    ("copy.col_unit",       "S", "col", "견적의뢰복사본: 견적단가 열"),
    ("copy.col_amount",     "T", "col", "견적의뢰복사본: 견적금액 열"),
    ("copy.col_check",      "U", "col", "견적의뢰복사본: 견적단가(Check) 열"),

    # ── 국내 견적서 (사양서 · 현업 사인용 사양서 · 입고검수확인서) ──
    ("dom.pump_row",   15, "row", "사양서: PUMP 행"),
    ("dom.spec_start", 17, "row", "사양서: 품목 첫 행"),
    ("dom.spec_end",   41, "row", "사양서: 품목 마지막 행 (칸 수 = 끝-첫+1)"),
    ("dom.sign_start", 11, "row", "현업 사인용 사양서: 품목 첫 행"),
    ("dom.sign_end",   35, "row", "현업 사인용 사양서: 품목 마지막 행"),
    ("dom.col_class", "A", "col", "사양서: 분류(PUMP 행에서는 모델명) 열"),
    ("dom.col_spec",  "B", "col", "사양서: 규격 열"),
    ("dom.col_qty",   "E", "col", "사양서: 수량 열"),
    ("dom.col_price", "F", "col", "사양서: 단가 열"),
    ("dom.col_amt",   "G", "col", "사양서: 금액 열"),
    ("dom.cell_investor",    "B4",  "cell", "사양서: 투자자 문구 칸"),
    ("dom.cell_equip_qty",   "A43", "cell", "사양서: 설비수량 칸"),
    ("dom.cell_actual_line", "B6",  "cell", "현업 사인용 사양서: 실반입라인 칸"),
    ("dom.cell_exh_size",    "D27", "cell", "입고검수확인서: Exh Size 칸"),
    ("dom.cell_dry",         "D17", "cell", "입고검수확인서: DRY/BOOSTER 칸"),
    ("dom.cell_cap",         "D18", "cell", "입고검수확인서: 용량 칸"),
    ("dom.cell_motor",       "D19", "cell", "입고검수확인서: MOTOR TYPE 칸"),
    ("dom.investor_text", "{투자자} 님 / 설비구매그룹", "name", "사양서 투자자 문구"),
    ("dom.text_no_model", "미등록 Q-Code", "text", "단가표에 Q-code 가 없을 때 모델명 칸에 적는 말"),
    ("dom.text_no_spec",  "-",            "text", "단가표에 Q-code 가 없을 때 규격 칸에 적는 말"),
    ("dom.text_no_class", "1",            "text", "악세서리 단가표에 규격이 없을 때 분류 칸에 적는 말"),
    ("dom.text_no_dry",   "의뢰파일DATA",  "text", "단가표에 없을 때 DRY/BOOSTER 칸에 적는 말"),
    ("dom.text_no_cap",   "확인필요",      "text", "단가표에 없을 때 용량 칸에 적는 말"),
    ("dom.text_no_motor", "확인필요",      "text", "단가표에 없을 때 MOTOR TYPE 칸에 적는 말"),

    # ── PUMP 단가표 / 악세서리 단가표 ───────────────────────────
    ("pump.col_code",  "C",  "col", "PUMP 단가표: 자재코드(Q-code) 열"),
    ("pump.col_spec",  "E",  "col", "PUMP 단가표: 규격 열"),
    ("pump.col_model", "F",  "col", "PUMP 단가표: 모델명 열"),
    ("pump.col_price", "K",  "col", "PUMP 단가표: 단가 열"),
    ("pump.col_dry",   "W",  "col", "PUMP 단가표: DRY/BOOSTER 열"),
    ("pump.col_cap",   "Z",  "col", "PUMP 단가표: 용량 열"),
    ("pump.col_motor", "AA", "col", "PUMP 단가표: MOTOR TYPE 열"),
    ("acc.col_spec",   "D",  "col", "악세서리 단가표: 규격 열"),
    ("acc.col_class",  "Q",  "col", "악세서리 단가표: 분류 열"),

    # ── 중국 견적서 ────────────────────────────────────────────
    ("cn.cell_maker",   "A17", "cell", "중국: 설비사 칸"),
    ("cn.cell_process", "B17", "cell", "중국: 공정 칸"),
    ("cn.cell_tool",    "B51", "cell", "중국: 설비호기 칸"),
    ("cn.cell_to",       "B3",  "cell", "중국: 받는 사람 칸 (옵션창에서 비우면 양식 그대로)"),
    ("cn.to_text",       "{받는사람}", "name", "중국: 받는 사람 문구"),
    ("cn.cell_quote_no", "G6",  "cell", "중국: Quote number 칸 (비우면 양식 그대로)"),
    ("cn.quote_no_text", "{견적번호}", "name", "중국: Quote number 문구"),
    ("cn.cell_delivery", "G10", "cell", "중국: Delivery date 칸 (비우면 양식 그대로)"),
    ("cn.delivery_text", "{납기}", "name", "중국: Delivery date 문구"),
    ("cn.cell_warranty", "G7",  "cell", "중국: 보증기간 칸"),
    ("cn.warranty_text", "{보증} Year after delivery", "name", "중국: 보증기간 문구"),
    ("cn.pump_start", 17, "row", "중국: PUMP 첫 행"),
    ("cn.pump_end",   19, "row", "중국: PUMP 마지막 행"),
    ("cn.rack_start", 21, "row", "중국: RACK 첫 행"),
    ("cn.rack_end",   40, "row", "중국: RACK 마지막 행"),
    ("cn.col_spec",  "C", "col", "중국: 규격 열"),
    ("cn.col_price", "F", "col", "중국: 단가 열"),
    ("cn.col_qty",   "G", "col", "중국: 수량 열"),

    # ── 미국 견적서 ────────────────────────────────────────────
    ("us.cell_maker",    "A17", "cell", "미국: 설비사 칸"),
    ("us.cell_process",  "B17", "cell", "미국: 공정 칸"),
    ("us.cell_tool",     "B54", "cell", "미국: 설비호기 칸"),
    ("us.cell_exchange", "A53", "cell", "미국: 환율 칸"),
    ("us.cell_date",     "E53", "cell", "미국: 기준월 칸"),
    ("us.cell_investor", "B3",  "cell", "미국: 투자자 문구 칸"),
    ("us.investor_text", "{투자자} 님 / 설비구매그룹", "name", "미국: 투자자 문구"),
    ("us.cell_quote_no", "G6",  "cell", "미국: Quote number 칸 (비우면 양식 그대로)"),
    ("us.quote_no_text", "{견적번호}", "name", "미국: Quote number 문구"),
    ("us.cell_pr",       "F8",  "cell", "미국: PR 칸 (비우면 양식 그대로)"),
    ("us.pr_text",       "PR : {PR}", "name", "미국: PR 문구"),
    ("us.cell_warranty", "F7",  "cell", "미국: 보증기간 칸"),
    ("us.warranty_text", "Warranty : {보증} Years after delivery", "name", "미국: 보증기간 문구"),
    ("us.pump_row",   17, "row", "미국: PUMP 행 (한 줄)"),
    ("us.rack_start", 21, "row", "미국: RACK 첫 행"),
    ("us.rack_end",   40, "row", "미국: RACK 마지막 행"),
    ("us.col_spec",  "C", "col", "미국: 규격 열"),
    ("us.col_price", "E", "col", "미국: 단가 열"),
    ("us.col_qty",   "G", "col", "미국: 수량 열"),

    # ── 갑지 ──────────────────────────────────────────────────
    ("cover.cell_investor", "B7",  "cell", "갑지: 투자자 문구 칸"),
    ("cover.cell_warranty", "I12", "cell", "갑지: 보증기간 문구 칸"),
    ("cover.investor_text", "삼성전자 ㈜ / {투자자} 님", "name", "갑지 투자자 문구"),
    ("cover.warranty_text", "{보증} Year after delivery", "name", "갑지 보증기간 문구"),
    ("cover.row_start",     20,  "row", "갑지: 견적 목록 첫 행"),
    ("cover.data_row_start", 2,  "row", "갑지DATA: 첫 데이터 행"),
    ("cover.rows",          30,  "int", "갑지에 실을 수 있는 견적 건수"),
    ("cover.data_last_col", "AA", "col", "갑지DATA: 견적의뢰복사본에서 옮겨 오는 마지막 열"),

    # ── 기본값 · 판별 ─────────────────────────────────────────
    ("default.investor",    "채승철", "text", "투자자 이름을 비워 두었을 때 쓰는 이름"),
    ("credit.pump_keyword", "Pump",  "text", "Credit 규격에 이 말이 있으면 Pump Credit 으로 본다"),

    # ── 파일 · 폴더 이름 ───────────────────────────────────────
    ("name.out_dir",     "견적서", "text", "생성물을 모으는 폴더 이름(프로그램 폴더 아래)"),
    ("name.dom_folder",  "{날짜}_LOT베큠_{라인}_{의뢰자}", "name", "국내 견적서 폴더 이름"),
    ("name.dom_file",    "대외비_{PR}-{항번}_{날짜}_LOT베큠_{라인}_{의뢰자}_{호기}.xlsx", "name", "국내 견적서 파일 이름"),
    ("name.cn_folder",   "{날짜}_중국_SCS_{라인}", "name", "중국 견적서 폴더 이름"),
    ("name.cn_file",     "대외비_{날짜}_중국_SCS_{라인}_{호기}.xlsx", "name", "중국 견적서 파일 이름"),
    ("name.us_folder",   "{날짜}_미국_{사이트}", "name", "미국 견적서 폴더 이름"),
    ("name.us_file",     "대외비_{날짜}_{사이트}_{호기}_Quotation.xlsx", "name", "미국 견적서 파일 이름"),
    ("name.cover_file",  "대외비_{폴더}{갑지표시}.xlsx", "name", "갑지 파일 이름"),
    ("name.cover_mark",  "_갑지", "text", "갑지 파일 이름 끝에 붙는 표시 (제출용 엑셀은 이걸 뗀 이름)"),
]

# 이름 규칙·문구에 쓸 수 있는 자리표시
_PLACEHOLDERS = {
    "dom.investor_text":   ("투자자",),
    "cover.investor_text": ("투자자",),
    "cover.warranty_text": ("보증",),
    "cn.to_text": ("받는사람",), "cn.quote_no_text": ("견적번호",),
    "cn.delivery_text": ("납기",), "cn.warranty_text": ("보증",),
    "us.investor_text": ("투자자",), "us.quote_no_text": ("견적번호",),
    "us.pr_text": ("PR",), "us.warranty_text": ("보증",),
    "name.dom_folder": ("날짜", "PR", "항번", "라인", "의뢰자", "호기"),
    "name.dom_file":   ("날짜", "PR", "항번", "라인", "의뢰자", "호기"),
    "name.cn_folder":  ("날짜", "라인", "호기"),
    "name.cn_file":    ("날짜", "라인", "호기"),
    "name.us_folder":  ("날짜", "사이트", "호기"),
    "name.us_file":    ("날짜", "사이트", "호기"),
    "name.cover_file": ("폴더", "갑지표시"),
}

DEFAULTS: Dict[str, Any] = {k: v for k, v, _, _ in _SPEC}
KIND: Dict[str, str] = {k: kind for k, _, kind, _ in _SPEC}
DESCRIPTION: Dict[str, str] = {k: d for k, _, _, d in _SPEC}

_CELL_RE = re.compile(r"^[A-Z]{1,3}[1-9][0-9]{0,6}$")
_COL_RE = re.compile(r"^[A-Z]{1,3}$")


# ──────────────────────────────────────────────
# 열 문자 유틸
# ──────────────────────────────────────────────
def col_index(letter: str) -> int:
    """열 문자 → 1부터 세는 번호. 'A' → 1, 'AA' → 27."""
    n = 0
    for ch in letter:
        n = n * 26 + (ord(ch) - 64)
    return n


def parse_cols(text: str) -> List[int]:
    """'A:N, R, W:AA' → [1..14, 18, 23..27] (1부터 세는 열 번호)."""
    out: List[int] = []
    for part in str(text).split(","):
        part = part.strip().upper()
        if not part:
            continue
        a, _, b = part.partition(":")
        a, b = a.strip(), (b.strip() or a.strip())
        if not (_COL_RE.match(a) and _COL_RE.match(b)):
            raise ValueError(f"열 범위를 읽을 수 없습니다: '{part}'")
        i, j = col_index(a), col_index(b)
        if i > j:
            raise ValueError(f"열 범위의 앞뒤가 바뀌었습니다: '{part}'")
        out.extend(range(i, j + 1))
    if not out:
        raise ValueError("열 범위가 비어 있습니다")
    return out


# ──────────────────────────────────────────────
# 값 변환
# ──────────────────────────────────────────────
def _coerce(key: str, raw: Any) -> Any:
    """설정 시트의 값을 종류에 맞게 바꾼다. 맞지 않으면 ValueError."""
    kind = KIND[key]
    if kind in ("row", "int"):
        if isinstance(raw, bool):
            raise ValueError("숫자가 아닙니다")
        if isinstance(raw, float) and raw.is_integer():
            raw = int(raw)
        if not isinstance(raw, int):
            try:
                raw = int(str(raw).strip())
            except ValueError:
                raise ValueError("숫자가 아닙니다")
        if raw < 1:
            raise ValueError("1 이상이어야 합니다")
        return raw
    text = str(raw)
    if isinstance(raw, float) and raw.is_integer():
        text = str(int(raw))            # 엑셀이 숫자로 바꿔 버린 '1' 같은 문구
    if kind == "cell":
        text = text.strip().upper().replace("$", "")
        if not _CELL_RE.match(text):
            raise ValueError("셀 주소가 아닙니다 (예: B4)")
        return text
    if kind == "col":
        text = text.strip().upper()
        if not _COL_RE.match(text):
            raise ValueError("열 문자가 아닙니다 (예: A, AA)")
        return text
    if kind == "cols":
        text = text.strip().upper()
        parse_cols(text)
        return text
    if kind == "sheet":
        text = text.strip()
        if not text:
            raise ValueError("시트 이름이 비어 있습니다")
        return text
    if kind == "name":
        try:
            text.format(**{p: "x" for p in _PLACEHOLDERS[key]})
        except (KeyError, IndexError, ValueError) as e:
            allowed = ", ".join("{%s}" % p for p in _PLACEHOLDERS[key])
            raise ValueError(f"쓸 수 없는 자리표시입니다: {e} (쓸 수 있는 것: {allowed})")
        return text
    return text                          # text — 앞뒤 공백도 뜻이 있을 수 있어 그대로


class QuoteConfig:
    """견적서작성 설정. cfg["dom.spec_start"] 처럼 꺼내 쓴다."""

    def __init__(self, overrides: Optional[Dict[str, Any]] = None,
                 problems: Optional[List[str]] = None) -> None:
        self._values = dict(DEFAULTS)
        self.overridden: Dict[str, Any] = {}
        self.problems: List[str] = list(problems or [])
        for k, v in (overrides or {}).items():
            self._values[k] = v
            if v != DEFAULTS[k]:
                self.overridden[k] = v

    def __getitem__(self, key: str) -> Any:
        return self._values[key]

    # 자주 쓰는 꺼내기 ──────────────────────────
    def col(self, key: str) -> int:
        """열 설정을 1부터 세는 번호로."""
        return col_index(self._values[key])

    def cols(self, key: str) -> List[int]:
        return parse_cols(self._values[key])

    def at(self, col_key: str, row: int) -> str:
        """열 설정 + 행 번호 → 셀 주소. at('dom.col_spec', 17) → 'B17'"""
        return f"{self._values[col_key]}{row}"

    def text(self, key: str, **values: Any) -> str:
        """문구·이름 규칙에 값을 채운다."""
        return self._values[key].format(**values)

    @property
    def required_sheets(self) -> List[str]:
        """통합양식에 반드시 있어야 하는 시트."""
        return [self._values[k] for k in (
            "sheet.items", "sheet.code_map", "sheet.req_copy", "sheet.spec",
            "sheet.sign_spec", "sheet.incoming", "sheet.quote_cn", "sheet.quote_us")]

    def capacity(self, start_key: str, end_key: str) -> int:
        return self._values[end_key] - self._values[start_key] + 1

    def check(self, sheetnames: Iterable[str]) -> List[str]:
        """양식과 맞지 않는 점을 사람이 읽을 문장으로 돌려준다(없으면 빈 목록).
        생성 도중이 아니라 양식을 읽는 시점에 알려 주기 위한 것이다."""
        out = list(self.problems)
        names = set(sheetnames)
        for key, _, kind, desc in _SPEC:
            if kind == "sheet" and self._values[key] not in names:
                out.append(f"시트 없음: '{self._values[key]}' ({desc}) — 설정 항목 {key}")
        for a, b, what in (("dom.spec_start", "dom.spec_end", "사양서 품목 행"),
                           ("dom.sign_start", "dom.sign_end", "현업 사인용 사양서 품목 행"),
                           ("cn.pump_start", "cn.pump_end", "중국 PUMP 행"),
                           ("cn.rack_start", "cn.rack_end", "중국 RACK 행"),
                           ("us.rack_start", "us.rack_end", "미국 RACK 행")):
            if self._values[a] > self._values[b]:
                out.append(f"{what}의 첫 행이 마지막 행보다 큽니다 ({a}={self._values[a]}, {b}={self._values[b]})")
        if self.capacity("dom.spec_start", "dom.spec_end") != self.capacity("dom.sign_start", "dom.sign_end"):
            out.append("사양서와 현업 사인용 사양서의 품목 칸 수가 다릅니다 "
                       f"({self.capacity('dom.spec_start', 'dom.spec_end')}행 / "
                       f"{self.capacity('dom.sign_start', 'dom.sign_end')}행)")
        return out


# ──────────────────────────────────────────────
# 읽기
# ──────────────────────────────────────────────
def read_settings_rows(rows: Iterable[tuple]) -> Tuple[Dict[str, Any], List[str]]:
    """설정 시트의 행들(값만) → (덮어쓸 값, 문제 목록). 1행 머리글은 알아서 건너뛴다."""
    overrides: Dict[str, Any] = {}
    problems: List[str] = []
    for n, row in enumerate(rows, start=1):
        if not row or row[0] is None:
            continue
        key = str(row[0]).strip()
        value = row[1] if len(row) > 1 else None
        if not key or key.startswith("#"):
            continue
        if key not in DEFAULTS:
            if n > 1:                      # 1행은 머리글일 수 있다
                problems.append(f"설정 {n}행: 모르는 항목 '{key}' — 무시합니다")
            continue
        if value is None or (isinstance(value, str) and value == ""):
            continue                       # 비워 두면 기본값
        try:
            overrides[key] = _coerce(key, value)
        except ValueError as e:
            problems.append(f"설정 {n}행 '{key}' = '{value}': {e} — 기본값 '{DEFAULTS[key]}' 을 씁니다")
    return overrides, problems


def from_workbook(wb) -> QuoteConfig:
    """열려 있는 통합양식(openpyxl)에서 설정을 읽는다. 워크북에 기억해 두고 재사용한다."""
    cached = getattr(wb, "_lot_quote_config", None)
    if cached is not None:
        return cached
    overrides: Dict[str, Any] = {}
    problems: List[str] = []
    if SETTINGS_SHEET in wb.sheetnames:
        ws = wb[SETTINGS_SHEET]
        overrides, problems = read_settings_rows(ws.iter_rows(max_col=2, values_only=True))
    cfg = QuoteConfig(overrides, problems)
    try:
        wb._lot_quote_config = cfg
    except Exception:
        pass
    return cfg


_cache: Dict[str, Tuple[Tuple[float, int], QuoteConfig]] = {}


def load(template_path: Optional[str]) -> QuoteConfig:
    """통합양식 파일에서 설정을 읽는다(파일이 바뀌기 전까지 기억). 경로가 없으면 기본값."""
    if not template_path:
        return QuoteConfig()
    try:
        st = os.stat(template_path)
    except OSError:
        return QuoteConfig()
    stamp = (st.st_mtime, st.st_size)
    hit = _cache.get(template_path)
    if hit and hit[0] == stamp:
        return hit[1]
    from openpyxl import load_workbook
    try:
        wb = load_workbook(template_path, read_only=True, data_only=True)
    except Exception as e:
        logger.warning("설정을 읽지 못해 기본값을 씁니다 (%s): %s", template_path, e)
        return QuoteConfig()
    try:
        cfg = from_workbook(wb)
    finally:
        wb.close()
    _cache[template_path] = (stamp, cfg)
    return cfg


# ──────────────────────────────────────────────
# 설정 시트 만들기
# ──────────────────────────────────────────────
def export_settings_sheet(out_path: str, cfg: Optional[QuoteConfig] = None) -> str:
    """지금 값으로 채운 `설정` 시트 한 장짜리 xlsx 를 만든다.
    이 시트를 통합양식에 복사해 넣고 필요한 값만 고치면 된다."""
    from openpyxl import Workbook
    from openpyxl.styles import Font
    cfg = cfg or QuoteConfig()
    wb = Workbook()
    ws = wb.active
    ws.title = SETTINGS_SHEET
    ws.append(["항목", "값", "설명"])
    for c in ws[1]:
        c.font = Font(bold=True)
    for key, _, _, desc in _SPEC:
        ws.append([key, cfg[key], desc])
    ws.column_dimensions["A"].width = 26
    ws.column_dimensions["B"].width = 52
    ws.column_dimensions["C"].width = 60
    for row in ws.iter_rows(min_row=2, min_col=2, max_col=2):
        row[0].number_format = "@" if isinstance(row[0].value, str) else "0"
    ws.freeze_panes = "A2"
    wb.save(out_path)
    return out_path


if __name__ == "__main__":
    # python quote_config.py [통합양식.xlsx] [만들파일.xlsx]
    #   통합양식을 주면 그 양식의 지금 설정으로, 안 주면 기본값으로 `설정` 시트를 만든다.
    import sys
    _args = sys.argv[1:]
    _src = _args[0] if _args and os.path.exists(_args[0]) else None
    _out = _args[-1] if _args and _args[-1] != _src else "견적서작성_설정.xlsx"
    _cfg = load(_src)
    for _p in _cfg.problems:
        print("확인:", _p)
    print("만들었습니다:", export_settings_sheet(_out, _cfg))
