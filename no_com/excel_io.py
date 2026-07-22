"""
excel_io.py  [no-COM edition]
==============================
Excel 입출력 전담 모듈.

- openpyxl  : 읽기 + 쓰기 전담
              (파싱, 견적서/갑지 생성 모두 openpyxl)
- Excel COM : PDF 변환 전용 (excel_to_merged_pdf 만 사용)

핵심 전략
---------
템플릿의 이미지·서식은 shutil.copy2 로 원본 ZIP 구조를 통째로
복사한 뒤 openpyxl 로 셀 값만 수정한다.
수식 재계산은 wb.calculation.fullCalcOnLoad = True 를 설정해
파일을 열 때 Excel 이 자동 처리하도록 위임한다.
"""

from __future__ import annotations

import logging
import os
import re
import shutil
from datetime import datetime
from typing import Any, Callable, Dict, List, Optional, Tuple

from openpyxl import load_workbook
from openpyxl.worksheet.worksheet import Worksheet

from core import (
    CN, DOM, US, SheetName, QuoteState,
    ensure_dir, safe_filename, s, to_float, unique_path,
    exe_dir, parse_invest_info,
)

logger = logging.getLogger("QuoteApp")


def _clear_clipboard() -> None:
    """
    Windows 클립보드를 강제로 비운다.

    CopyPicture()/ExportAsFixedFormat() 후 COM 세션이 종료되면
    클립보드에 죽은 프로세스 소유의 CF_ENHMETAFILE(EMF) 핸들이 남는다.
    이 stale 핸들이 있는 상태에서 별도 Excel 창이 Ctrl+C를 시도하면
    Windows가 이전 클립보드를 초기화하다가 액세스 위반 → Excel 전체 종료.
    COM 세션 종료 전, 또는 CopyPicture 직후에 반드시 호출한다.
    """
    try:
        import ctypes
        u32 = ctypes.windll.user32
        if u32.OpenClipboard(0):
            u32.EmptyClipboard()
            u32.CloseClipboard()
    except Exception:
        pass


# ── COM 지연 로딩 (PDF 변환 전용) ───────────────────────────────
# 모듈 임포트 시 DLL LoadLibrary를 호출하지 않도록 지연.
# GIL 점유로 인한 시작 시 UI 동결 방지 (AV/Defender 환경에서 10~120초 블로킹 가능).
pythoncom: Any = None
win32: Any = None
COM_AVAILABLE: Optional[bool] = None  # None=미확인, True=사용가능, False=불가


def _ensure_com() -> bool:
    """COM DLL을 최초 필요 시점에만 로드한다."""
    global COM_AVAILABLE, pythoncom, win32
    if COM_AVAILABLE is not None:
        return COM_AVAILABLE
    try:
        import pythoncom as _pc
        import win32com.client as _w32
        pythoncom = _pc
        win32 = _w32
        COM_AVAILABLE = True
    except Exception:
        pythoncom = None
        win32 = None
        COM_AVAILABLE = False
    return COM_AVAILABLE


# ═══════════════════════════════════════════════════════════════
# 읽기 (openpyxl) — 기존과 동일
# ═══════════════════════════════════════════════════════════════

def parse_items_sheet(ws: Worksheet) -> Tuple[
    List[Dict[str, Any]],
    Dict[str, List[Dict[str, Any]]],
    Dict[str, float],
    Dict[str, int],
]:
    """품목 시트 파싱. 반환: (rows, by_class, price_by_spec, spec_order_index)"""
    def _is_header(a: str, b: str, c: str) -> bool:
        h = (a + b + c).replace(" ", "").lower()
        return "분류" in h and ("단가" in h or "가격" in h)

    start = 2 if _is_header(
        s(ws.cell(1, 1).value), s(ws.cell(1, 2).value), s(ws.cell(1, 3).value)
    ) else 1

    rows: List[Dict[str, Any]] = []
    by_class: Dict[str, List[Dict[str, Any]]] = {}
    price_by_spec: Dict[str, float] = {}
    order_index: Dict[str, int] = {}
    order = 0

    for r in range(start, ws.max_row + 1):
        a = s(ws.cell(r, 1).value)
        b = s(ws.cell(r, 2).value)
        c = ws.cell(r, 3).value
        if not (a or b or c):
            continue
        price = to_float(c)
        row = {"A": a, "B": b, "C": price}
        rows.append(row)
        if b and b not in order_index:
            order_index[b] = order
            order += 1
        if a:
            by_class.setdefault(a, []).append(row)
        if b:
            price_by_spec[b] = price

    return rows, by_class, price_by_spec, order_index


def parse_code_map_sheet(ws: Worksheet) -> Dict[Tuple[str, str, str], List[Tuple[str, float]]]:
    """코드매핑 시트 파싱. 반환: {(공정, 설비사, 5D): [(ACC, 수량), ...]}"""
    def _is_header(a, b, c, d, e) -> bool:
        h = (a + b + c + d + e).replace(" ", "").lower()
        return "공정" in h and "설비" in h and ("acc" in h or "수량" in h or "5d" in h)

    start = 2 if _is_header(
        s(ws.cell(1, 1).value), s(ws.cell(1, 2).value), s(ws.cell(1, 3).value),
        s(ws.cell(1, 4).value), s(ws.cell(1, 5).value),
    ) else 1

    cmap: Dict[Tuple[str, str, str], List[Tuple[str, float]]] = {}
    for r in range(start, ws.max_row + 1):
        proc   = s(ws.cell(r, 1).value)
        vendor = s(ws.cell(r, 2).value)
        code5d = s(ws.cell(r, 3).value)
        acc    = s(ws.cell(r, 4).value)
        qty    = to_float(ws.cell(r, 5).value)
        if proc and vendor and code5d and acc:
            cmap.setdefault((proc, vendor, code5d), []).append((acc, qty))
    return cmap


def parse_request_xlsx(path: str) -> Tuple[str, List[Dict[str, Any]]]:
    """
    의뢰파일 파싱.
    반환: (active_sheet_name, row_dicts)
    row_dict keys: row_idx, D, E, F, G, H, J, K, N, V, X, Y, Z, AA
    """
    wb = load_workbook(path, data_only=True)
    ws = wb.active
    sheet_name = ws.title

    def _is_header() -> bool:
        checks = [
            s(ws.cell(1, 26).value).replace(" ", "").lower(),
            s(ws.cell(1, 11).value).replace(" ", "").lower(),
            s(ws.cell(1, 24).value).replace(" ", "").lower(),
            s(ws.cell(1, 22).value).replace(" ", "").lower(),
        ]
        return any("설비" in c or "공정" in c or "5d" in c or "코드" in c for c in checks)

    start = 2 if _is_header() else 1
    rows: List[Dict[str, Any]] = []
    for r in range(start, ws.max_row + 1):
        z = ws.cell(r, 26).value
        if s(z) == "":
            continue
        rows.append({
            "row_idx": r,
            "D": ws.cell(r,  4).value,
            "E": ws.cell(r,  5).value,
            "F": ws.cell(r,  6).value,
            "G": ws.cell(r,  7).value,
            "H": ws.cell(r,  8).value,
            "J": ws.cell(r, 10).value,
            "K": ws.cell(r, 11).value,
            "N": ws.cell(r, 14).value,
            "V": ws.cell(r, 22).value,
            "X": ws.cell(r, 24).value,
            "Y": ws.cell(r, 25).value,
            "Z": z,
            "AA": ws.cell(r, 27).value,
        })
    return sheet_name, rows


# ═══════════════════════════════════════════════════════════════
# openpyxl 공통 헬퍼
# ═══════════════════════════════════════════════════════════════

def _hide_rows(ws: Worksheet, start: int, end: int, filled_count: int) -> None:
    """start~end 행 중 filled_count 이후 행을 숨김."""
    for i in range(end - start + 1):
        ws.row_dimensions[start + i].hidden = (i >= filled_count)


def _show_only_sheets(wb, keep: List[str]) -> None:
    """keep 목록에 없는 시트를 veryHidden으로 설정."""
    keep_set = set(keep)
    for name in wb.sheetnames:
        wb[name].sheet_state = "visible" if name in keep_set else "veryHidden"


def _write_req_row_openpyxl(wb_copy, state: QuoteState, rd: Dict[str, Any],
                             req_ws_cache: Optional[Any] = None) -> None:
    """
    의뢰파일의 특정 행을 견적의뢰복사본 시트 2행에 복사 (openpyxl).
    req_ws_cache: 외부에서 미리 열어 전달된 Worksheet 객체(멀티 생성 시 재사용).
                  None이면 여기서 직접 open/close 한다.
    """
    req_path = state.request_path
    if not req_path or not req_path.lower().endswith(".xlsx"):
        return
    try:
        row_idx = int(rd.get("row_idx", 0))
    except (TypeError, ValueError):
        return
    if row_idx <= 0:
        return

    def _copy_row(req_ws) -> None:
        copy_ws = wb_copy[SheetName.REQ_COPY]
        for c in range(1, 15):
            copy_ws.cell(2, c).value = None
        copy_ws.cell(2, 18).value = None
        for c in range(23, 28):
            copy_ws.cell(2, c).value = None
        for c in range(1, 15):
            copy_ws.cell(2, c).value = req_ws.cell(row_idx, c).value
        copy_ws.cell(2, 18).value = req_ws.cell(row_idx, 18).value
        for c in range(23, 28):
            copy_ws.cell(2, c).value = req_ws.cell(row_idx, c).value

    if req_ws_cache is not None:
        _copy_row(req_ws_cache)
        return

    req_wb = load_workbook(req_path, data_only=True)
    try:
        sheet_name = state.request_sheet_name
        req_ws = (
            req_wb[sheet_name]
            if sheet_name and sheet_name in req_wb.sheetnames
            else req_wb.active
        )
        _copy_row(req_ws)
    finally:
        req_wb.close()


# ═══════════════════════════════════════════════════════════════
# 견적서 생성 (openpyxl 기반, COM 없음)
# ═══════════════════════════════════════════════════════════════

def _fill_domestic(state: QuoteState,
                   rd: Dict[str, Any],
                   items: List[Dict[str, Any]],
                   req_ws_cache: Optional[Any] = None) -> str:
    """
    국내 견적서: 템플릿을 복사한 뒤 openpyxl 로 데이터 기입.
    반환: 저장된 xlsx 경로.
    """
    ymd        = datetime.now().strftime("%y%m%d")
    pr         = safe_filename(s(rd.get("D")))
    itemno     = safe_filename(s(rd.get("E")))
    lineproc   = safe_filename(s(rd.get("K")))
    investor_fn = safe_filename(s(rd.get("J")))
    tool       = safe_filename(s(rd.get("Z")))
    folder = ensure_dir(
        os.path.join(exe_dir(), "견적서", f"{ymd}_LOT베큠_{lineproc}_{investor_fn}")
    )
    path = unique_path(
        os.path.join(folder, f"{pr}-{itemno}_{ymd}_LOT베큠_{lineproc}_{investor_fn}_{tool}.xlsx")
    )

    # ① 템플릿 전체를 복사 — 이미지·서식·수식 구조 모두 보존
    shutil.copy2(state.template_path, path)

    wb      = load_workbook(path)
    ws_spec = wb[SheetName.SPEC]
    ws_sign = wb[SheetName.SIGN_SPEC]

    # ② 의뢰파일 행 → 견적의뢰복사본 시트 복사
    _write_req_row_openpyxl(wb, state, rd, req_ws_cache)

    # ③ 투자자명
    investor = s(state.investor_name) or "채승철"
    ws_spec["B4"] = f"{investor} 님 / 설비구매그룹"

    rack_items = [r for r in items if r["cat"] != "PUMP" and r["role"] != "CREDIT"]
    credits    = [r for r in items if r["role"] == "CREDIT"]

    # ④ 기존 데이터 클리어
    for rr in range(DOM.SPEC_START, DOM.SPEC_END + 1):
        for col in (DOM.COL_SPEC, DOM.COL_QTY, DOM.COL_PRICE, DOM.COL_AMT):
            ws_spec[f"{col}{rr}"] = None

    # ⑤ 품목 기입
    out_list = rack_items[:25]
    for cr in credits:
        out_list.append({"spec": cr["spec"], "qty": 0.0, "price": 0.0, "amt": cr["amt"]})
    out_list = out_list[:25]

    for i, r in enumerate(out_list):
        rr = DOM.SPEC_START + i
        ws_spec[f"{DOM.COL_SPEC}{rr}"] = r["spec"]
        if r.get("qty") == 0.0 and r.get("price") == 0.0 and "amt" in r:
            ws_spec[f"{DOM.COL_AMT}{rr}"] = r["amt"]
        else:
            ws_spec[f"{DOM.COL_QTY}{rr}"]   = r["qty"]
            ws_spec[f"{DOM.COL_PRICE}{rr}"]  = r["price"]
            ws_spec[f"{DOM.COL_AMT}{rr}"]    = r["amt"]

    # ⑥ 수식 재계산을 파일 열 때 Excel 에 위임
    wb.calculation.fullCalcOnLoad = True

    # ⑦ 빈 행 숨기기 (SPEC / SIGN_SPEC 모두 같은 개수)
    filled = len(out_list)
    _hide_rows(ws_spec, DOM.SPEC_START, DOM.SPEC_END, filled)
    _hide_rows(ws_sign, DOM.SIGN_START, DOM.SIGN_END, filled)

    # ⑧ 불필요한 시트 숨기기
    _show_only_sheets(wb, [
        SheetName.SPEC, SheetName.SIGN_SPEC,
        SheetName.INCOMING, SheetName.REQ_COPY,
    ])

    wb.save(path)
    wb.close()
    logger.info("국내 견적서 생성: %s", path)
    return path


def _fill_china(state: QuoteState, items: List[Dict[str, Any]]) -> str:
    """중국 견적서: 템플릿 복사 후 openpyxl 로 데이터 기입."""
    info   = state.cn_info
    ymd    = datetime.now().strftime("%y%m%d")
    line   = safe_filename(info.get("line", ""))
    tool   = safe_filename(info.get("tool", ""))
    folder = ensure_dir(
        os.path.join(exe_dir(), "견적서", f"{ymd}_중국_SCS_{line}")
    )
    path = unique_path(
        os.path.join(folder, f"{ymd}_중국_SCS_{line}_{tool}.xlsx")
    )

    shutil.copy2(state.template_path, path)
    wb = load_workbook(path)
    ws = wb[SheetName.QUOTE_CN]

    ws[CN.MAKER_CELL]   = info.get("maker", "")
    ws[CN.PROCESS_CELL] = info.get("process", "")
    ws[CN.TOOL_CELL]    = info.get("tool", "")

    pump    = [r for r in items if r["role"] != "CREDIT" and r["cat"] == "PUMP"]
    others  = [r for r in items if r["role"] != "CREDIT" and r["cat"] != "PUMP"]
    credits = [r for r in items if r["role"] == "CREDIT"]
    pump_credits = [r for r in credits if "Pump" in r.get("spec", "")]
    rack_credits = [r for r in credits if "Pump" not in r.get("spec", "")]

    # 펌프 구간 클리어 & 기입 (Pump Credit 포함)
    for rr in range(CN.PUMP_START, CN.PUMP_END + 1):
        ws[f"{CN.COL_SPEC}{rr}"]  = None
        ws[f"{CN.COL_PRICE}{rr}"] = None
        ws[f"{CN.COL_QTY}{rr}"]   = None

    pump_out = list(pump[:3])
    for cr in pump_credits:
        pump_out.append({"spec": cr["spec"], "qty": 0.0, "price": cr["amt"]})
    for i, r in enumerate(pump_out[:3]):
        rr = CN.PUMP_START + i
        ws[f"{CN.COL_SPEC}{rr}"]  = r["spec"]
        ws[f"{CN.COL_PRICE}{rr}"] = r["price"]
        ws[f"{CN.COL_QTY}{rr}"]   = None if float(r.get("qty", 0)) == 0 else r["qty"]

    # 랙 구간 클리어 & 기입 (Rack Credit만 포함)
    for rr in range(CN.RACK_START, CN.RACK_END + 1):
        ws[f"{CN.COL_SPEC}{rr}"]  = None
        ws[f"{CN.COL_PRICE}{rr}"] = None
        ws[f"{CN.COL_QTY}{rr}"]   = None

    out_list = others[:20]
    for cr in rack_credits:
        out_list.append({"spec": cr["spec"], "qty": 0.0, "price": cr["amt"]})
    out_list = out_list[:20]

    for i, r in enumerate(out_list):
        rr = CN.RACK_START + i
        ws[f"{CN.COL_SPEC}{rr}"]  = r["spec"]
        ws[f"{CN.COL_PRICE}{rr}"] = r["price"]
        ws[f"{CN.COL_QTY}{rr}"]   = None if float(r.get("qty", 0)) == 0 else r["qty"]

    wb.calculation.fullCalcOnLoad = True
    _hide_rows(ws, CN.RACK_START, CN.RACK_END, len(out_list))
    _show_only_sheets(wb, [SheetName.QUOTE_CN])

    wb.save(path)
    wb.close()
    logger.info("중국 견적서 생성: %s", path)
    return path


def _fill_us(state: QuoteState, items: List[Dict[str, Any]]) -> str:
    """미국 견적서: 템플릿 복사 후 openpyxl 로 데이터 기입."""
    info   = state.us_info
    ymd    = datetime.now().strftime("%y%m%d")
    site   = safe_filename(info.get("site", ""))
    tool   = safe_filename(info.get("tool", ""))
    folder = ensure_dir(
        os.path.join(exe_dir(), "견적서", f"{ymd}_미국_{site}")
    )
    path = unique_path(
        os.path.join(folder, f"{ymd}_{site}_{tool}_Quotation.xlsx")
    )

    shutil.copy2(state.template_path, path)
    wb = load_workbook(path)
    ws = wb[SheetName.QUOTE_US]

    ws[US.MAKER_CELL]    = info.get("maker", "")
    ws[US.PROCESS_CELL]  = info.get("process", "")
    ws[US.TOOL_CELL]     = info.get("tool", "")
    ws[US.EXCHANGE_CELL] = float(info.get("exchange", 0))
    y, m = info.get("base_ym", (datetime.now().year, datetime.now().month))
    ws[US.DATE_CELL] = f"{int(y)}-{int(m):02d}-1"

    pump    = [r for r in items if r["role"] != "CREDIT" and r["cat"] == "PUMP"]
    others  = [r for r in items if r["role"] != "CREDIT" and r["cat"] != "PUMP"]
    credits = [r for r in items if r["role"] == "CREDIT"]
    pump_credits = [r for r in credits if "Pump" in r.get("spec", "")]
    rack_credits = [r for r in credits if "Pump" not in r.get("spec", "")]

    # 펌프 단일 행 (Pump Credit은 RACK 구간 맨 앞 줄로 표기)
    ws[US.PUMP_SPEC]  = None
    ws[US.PUMP_PRICE] = None
    ws[US.PUMP_QTY]   = None
    if pump:
        ws[US.PUMP_SPEC]  = pump[0]["spec"]
        ws[US.PUMP_PRICE] = pump[0]["price"]
        ws[US.PUMP_QTY]   = pump[0]["qty"]

    # 랙 구간 클리어 & 기입 (Pump Credit → 맨 앞, Rack Credit → 맨 뒤)
    for rr in range(US.RACK_START, US.RACK_END + 1):
        ws[f"{US.COL_SPEC}{rr}"]  = None
        ws[f"{US.COL_PRICE}{rr}"] = None
        ws[f"{US.COL_QTY}{rr}"]   = None

    out_list = []
    for cr in pump_credits:
        out_list.append({"spec": cr["spec"], "qty": 0.0, "price": cr["amt"]})
    out_list += others[:20 - len(pump_credits)]
    for cr in rack_credits:
        out_list.append({"spec": cr["spec"], "qty": 0.0, "price": cr["amt"]})
    out_list = out_list[:20]

    for i, r in enumerate(out_list):
        rr = US.RACK_START + i
        ws[f"{US.COL_SPEC}{rr}"]  = r["spec"]
        ws[f"{US.COL_PRICE}{rr}"] = r["price"]
        ws[f"{US.COL_QTY}{rr}"]   = None if float(r.get("qty", 0)) == 0 else r["qty"]

    wb.calculation.fullCalcOnLoad = True
    _hide_rows(ws, US.RACK_START, US.RACK_END, len(out_list))
    _show_only_sheets(wb, [SheetName.QUOTE_US])

    wb.save(path)
    wb.close()
    logger.info("미국 견적서 생성: %s", path)
    return path


# ═══════════════════════════════════════════════════════════════
# 공개 API: 견적서 생성
# ═══════════════════════════════════════════════════════════════

def generate_quote(
    state   : QuoteState,
    qtype   : str,
    items   : List[Dict[str, Any]],
    rd      : Dict[str, Any],
) -> str:
    """견적서 1건 생성 → 저장 → 경로 반환."""
    if qtype == "중국":
        return _fill_china(state, items)
    if qtype == "미국":
        return _fill_us(state, items)
    return _fill_domestic(state, rd, items)


def generate_quote_multi(
    state        : QuoteState,
    items        : List[Dict[str, Any]],
    request_rows : List[Dict[str, Any]],
    progress_cb  : Optional[Callable[[int, int, str], None]] = None,
) -> List[Tuple[Dict[str, Any], str]]:
    """
    국내 견적서 여러 건을 순서대로 생성.
    Excel 을 전혀 실행하지 않으므로 COM 버전 대비 크게 빠르다.
    반환: [(rd, saved_path), ...]  실패 시 path=""
    """
    results: List[Tuple[Dict[str, Any], str]] = []
    total = len(request_rows)

    # 의뢰파일을 한 번만 열어 전 행에 재사용 — 반복 open/close 방지
    req_ws_cache = None
    req_wb_cache = None
    req_path = state.request_path
    if req_path and req_path.lower().endswith(".xlsx"):
        try:
            req_wb_cache = load_workbook(req_path, data_only=True)
            sn = state.request_sheet_name
            req_ws_cache = (
                req_wb_cache[sn]
                if sn and sn in req_wb_cache.sheetnames
                else req_wb_cache.active
            )
        except Exception as e:
            logger.warning("의뢰파일 사전 로드 실패 (행별 열기로 대체): %s", e)
            req_wb_cache = None
            req_ws_cache = None

    try:
        for i, rd in enumerate(request_rows):
            try:
                path = _fill_domestic(state, rd, items, req_ws_cache)
                results.append((rd, path))
                if progress_cb:
                    progress_cb(i + 1, total, os.path.basename(path))
            except Exception as e:
                logger.error("멀티 생성 실패: %s", e, exc_info=True)
                results.append((rd, ""))
                if progress_cb:
                    progress_cb(i + 1, total, f"[실패] {e}")
    finally:
        if req_wb_cache is not None:
            try:
                req_wb_cache.close()
            except Exception:
                pass

    return results


def generate_cover(
    template_path : str,
    folder        : str,
    source_files  : List[str],
    investor_name : str = "채승철",
    progress_cb   : Optional[Callable[[int, int, str], None]] = None,
    warranty_years: int = 2,
) -> str:
    """
    갑지 생성.
    source_files 각각의 견적의뢰복사본 2행 데이터를 openpyxl 로 읽어
    갑지DATA 시트에 쌓은 뒤 저장.
    """
    folder_name = os.path.basename(folder.rstrip("\\/"))
    out_path = unique_path(os.path.join(folder, f"{folder_name}_갑지.xlsx"))

    input_abs = {os.path.abspath(p).lower() for p in source_files}
    while os.path.abspath(out_path).lower() in input_abs:
        out_path = unique_path(out_path)

    def _natural_key(s: str):
        return [int(c) if c.isdigit() else c.lower() for c in re.split(r'(\d+)', s)]

    out_base = os.path.basename(out_path).lower()
    clean = sorted(
        [p for p in source_files if os.path.basename(p).lower() != out_base],
        key=lambda p: _natural_key(os.path.basename(p)),
    )
    if not clean:
        raise RuntimeError("처리할 엑셀 파일이 없습니다.")

    # 템플릿 복사 후 openpyxl 로 편집
    shutil.copy2(template_path, out_path)
    wb       = load_workbook(out_path)
    ws_data  = wb[SheetName.COVER_DATA]
    ws_cover = wb[SheetName.COVER]

    # 갑지DATA 초기화 — iter_rows 공개 API 사용 (내부 _cells 직접 접근 금지)
    max_data_row = min(ws_data.max_row or 1, 2000)
    if max_data_row >= 2:
        for row_cells in ws_data.iter_rows(min_row=2, max_row=max_data_row,
                                            min_col=1, max_col=27):
            for cell in row_cells:
                if cell.value is not None:
                    cell.value = None

    # 각 source 파일의 견적의뢰복사본 2행 읽기
    write_row = 2
    total_src = len(clean)
    for idx, path in enumerate(clean):
        if progress_cb:
            progress_cb(idx + 1, total_src, os.path.basename(path))
        try:
            src_wb = load_workbook(path, data_only=True)
            src_ws = src_wb[SheetName.REQ_COPY]
            for c in range(1, 28):  # A:AA
                ws_data.cell(write_row, c).value = src_ws.cell(2, c).value
            src_wb.close()
            write_row += 1
        except Exception as e:
            logger.warning("갑지 원본 읽기 실패 (%s): %s", path, e)

    # 투자자명
    ws_cover["B7"] = f"삼성전자 ㈜ / {investor_name} 님"

    # 보증기간 (기본값: 2 Year after delivery)
    ws_cover["I12"] = f"{warranty_years} Year after delivery"

    # 갑지DATA A2:A31 을 직접 확인하여 빈 행을 COVER(20~49) + COVER_DATA(2~31) 양쪽에 숨기기
    for i in range(30):
        a_val = ws_data.cell(2 + i, 1).value
        blank = a_val is None or str(a_val).strip() == ""
        ws_cover.row_dimensions[20 + i].hidden = blank
        ws_data.row_dimensions[2 + i].hidden   = blank

    wb.calculation.fullCalcOnLoad = True
    _show_only_sheets(wb, [SheetName.COVER_DATA, SheetName.COVER])
    wb.save(out_path)
    wb.close()
    logger.info("갑지 생성: %s", out_path)
    return out_path


# ═══════════════════════════════════════════════════════════════
# COM 헬퍼 — PDF 변환 전용 (변경 없음)
# ═══════════════════════════════════════════════════════════════

class ExcelCOM:
    """
    Excel COM 세션 컨텍스트 매니저.
    excel_to_merged_pdf 에서만 사용한다.
    """

    def __init__(self) -> None:
        self._excel = None

    def __enter__(self) -> "ExcelCOM":
        if not _ensure_com():
            raise RuntimeError("pywin32가 설치되어 있지 않습니다.")
        pythoncom.CoInitialize()
        self._excel = win32.DispatchEx("Excel.Application")
        self._excel.Visible = False
        self._excel.DisplayAlerts = False
        return self

    def __exit__(self, *_) -> None:
        try:
            self._excel.Quit()
        except Exception:
            pass
        try:
            pythoncom.CoUninitialize()
        except Exception:
            pass

    @property
    def app(self):
        return self._excel

    def open(self, path: str, read_only: bool = False):
        return self._excel.Workbooks.Open(
            path, ReadOnly=read_only, UpdateLinks=0, AddToMru=False
        )


# ═══════════════════════════════════════════════════════════════
# 공개 API: 전자서명용 임시 PDF 생성 (COM 유지)
# ═══════════════════════════════════════════════════════════════

def excel_to_merged_pdf(xlsx_path: str, tmp_dir: str, file_index: int,
                        xl_app=None) -> str:
    """
    xlsx 의 ESIGN_TARGET 시트(원래 Visible인 것)만 PDF 로 내보낸다.
    반환: PDF 경로 (대상 시트 없으면 "")

    워크북 단위 ExportAsFixedFormat 1회 호출:
      - 비대상 시트를 일시적으로 VeryHidden 처리 → workbook 전체 export
      - Close(False) 로 디스크 파일은 변경하지 않음
      - RenameFile 발생 횟수: 파일당 1회 (기존 시트 수만큼 → 1회로 감소)

    xl_app: 공유 Excel.Application COM 객체. None이면 자체 ExcelCOM 사용.
    """
    import fitz

    out_pdf = os.path.join(tmp_dir, f"tmp_{file_index:03d}.pdf")

    def _process(app) -> str:
        wb = app.Workbooks.Open(xlsx_path, ReadOnly=False, UpdateLinks=0, AddToMru=False)
        try:
            target_set = set(SheetName.ESIGN_TARGET)
            to_hide = []

            # 원래 Visible 이면서 ESIGN_TARGET에 속하는 시트만 유지
            has_visible_target = False
            for ws in wb.Worksheets:
                is_target  = ws.Name in target_set
                was_visible = int(ws.Visible) == -1   # xlSheetVisible = -1
                if is_target and was_visible:
                    has_visible_target = True
                else:
                    to_hide.append(ws)

            if not has_visible_target:
                return ""

            # 비대상 시트를 VeryHidden (xlVeryHidden = 2)
            for ws in to_hide:
                try:
                    ws.Visible = 2
                except Exception:
                    pass

            # 워크북 전체를 PDF 1회 export (ExportAsFixedFormat 1회 = RenameFile 1회)
            wb.ExportAsFixedFormat(0, out_pdf)
            return out_pdf if os.path.exists(out_pdf) else ""
        finally:
            _clear_clipboard()   # ExportAsFixedFormat 이 클립보드를 사용하는 경우 대비
            try:
                wb.Close(False)   # 디스크 파일 변경 없음
            except Exception:
                pass

    if xl_app is not None:
        return _process(xl_app)

    with ExcelCOM() as xl:
        return _process(xl.app)


def _print_area_range(ws):
    """PrintArea(R1C1 또는 A1 형식)를 Range 객체로 반환. 없으면 UsedRange."""
    import re
    pa = ws.PageSetup.PrintArea
    if not pa:
        return ws.UsedRange
    if "!" in pa:
        pa = pa.split("!")[-1]
    # R1C1 형식 감지: "R숫자C숫자:R숫자C숫자" 또는 단일 "R숫자C숫자"
    m = re.match(r'R(\d+)C(\d+)(?::R(\d+)C(\d+))?$', pa.strip())
    if m:
        r1, c1 = int(m.group(1)), int(m.group(2))
        r2 = int(m.group(3)) if m.group(3) else r1
        c2 = int(m.group(4)) if m.group(4) else c1
        return ws.Range(ws.Cells(r1, c1), ws.Cells(r2, c2))
    # A1 형식
    try:
        return ws.Range(pa)
    except Exception:
        return ws.UsedRange


def excel_capture_sheets_to_pngs(xlsx_path: str, tmp_dir: str, file_index: int,
                                  xl_app=None) -> List[str]:
    """
    xlsx ESIGN_TARGET 가시 시트를 클립보드로 캡처 → PNG 저장.
    ExportAsFixedFormat/PrintOut 미사용 → RenameFile 없음.

    xl_app: 공유 Excel.Application COM 객체. None이면 자체 ExcelCOM 사용.
    반환: 저장된 PNG 경로 리스트 (순서 = ESIGN_TARGET 순서)
    """
    try:
        from PIL import ImageGrab
    except ImportError:
        logger.error("Pillow(PIL) 없음 — pip install pillow")
        return []

    png_paths: List[str] = []

    def _process(app) -> List[str]:
        wb = app.Workbooks.Open(xlsx_path, ReadOnly=True, UpdateLinks=0, AddToMru=False)
        try:
            # ESIGN_TARGET 시트 우선, 없으면 보이는 시트 전체 캡처
            target_names = []
            for name in SheetName.ESIGN_TARGET:
                try:
                    ws = wb.Worksheets(name)
                    if int(ws.Visible) == -1:
                        target_names.append(name)
                except Exception:
                    continue
            if not target_names:
                for ws in wb.Worksheets:
                    try:
                        if int(ws.Visible) == -1:
                            target_names.append(ws.Name)
                    except Exception:
                        continue

            for idx, name in enumerate(target_names):
                try:
                    ws = wb.Worksheets(name)
                except Exception:
                    continue
                if int(ws.Visible) != -1:
                    continue
                safe_name = "".join(c if c not in r'\/:*?"<>|' else "_" for c in name)
                png_path = os.path.join(
                    tmp_dir, f"cap_{file_index:03d}_{idx:02d}_{safe_name}.png")
                try:
                    ws.Activate()
                    # 페이지 나누기 미리보기 → 기본(기본 보기)으로 전환 후 캡처
                    try:
                        app.ActiveWindow.View = 1  # xlNormalView
                    except Exception:
                        pass
                    rng = _print_area_range(ws)
                    rng.CopyPicture(Appearance=1, Format=2)  # xlScreen, xlBitmap
                    img = ImageGrab.grabclipboard()
                    # ── 클립보드 즉시 해제 ──────────────────────────────────────
                    # CopyPicture 후 Windows 클립보드에 CF_ENHMETAFILE(EMF) 핸들이
                    # 남아 있는 상태에서 COM 세션(xl.Quit)이 종료되면 stale 핸들이
                    # 발생한다. 이후 사용자가 별도 Excel에서 Ctrl+C를 시도하면
                    # Windows가 이 핸들을 해제하려다 액세스 위반 → Excel 전체 종료.
                    # 이미지를 읽은 직후 클립보드를 비워 이 경로를 차단한다.
                    _clear_clipboard()
                    if img is not None:
                        img.save(png_path, "PNG")
                        png_paths.append(png_path)
                    else:
                        logger.warning("클립보드 캡처 실패 (%s / %s)", xlsx_path, name)
                except Exception as e:
                    logger.warning("시트 캡처 실패 (%s / %s): %s", xlsx_path, name, e)
        finally:
            _clear_clipboard()   # 워크북 닫기 전 최종 클립보드 해제
            try:
                wb.Close(False)
            except Exception:
                pass
        return png_paths

    if xl_app is not None:
        return _process(xl_app)

    with ExcelCOM() as xl:
        return _process(xl.app)
