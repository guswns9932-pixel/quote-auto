"""
excel_io.py  [no-COM edition]
==============================
Excel 입출력 전담 모듈.

- openpyxl  : 읽기 + 쓰기 전담
              (파싱, 견적서/갑지 생성 모두 openpyxl)
- Excel COM : 갑지 생성 시 수식값 읽기에만 사용 (공통 헬퍼는 excel_com, 전자서명은 esign_engine)

핵심 전략
---------
템플릿의 이미지·서식은 shutil.copy2 로 원본 ZIP 구조를 통째로
복사한 뒤 openpyxl 로 셀 값만 수정한다.
수식 재계산은 wb.calculation.fullCalcOnLoad = True 를 설정해
파일을 열 때 Excel 이 자동 처리하도록 위임한다.
"""

from __future__ import annotations

import io
import logging
import math
import os
import re
import shutil
import stat
import threading
import time
from datetime import datetime
from typing import Any, Callable, Dict, List, Optional, Tuple

from openpyxl import load_workbook
from openpyxl.worksheet.worksheet import Worksheet

import excel_com
import quote_config
from excel_com import _clear_clipboard
from core import (
    QuoteState,
    ensure_dir, safe_filename, s, to_float, unique_path,
    exe_dir,
)

logger = logging.getLogger("QuoteApp")


# ──────────────────────────────────────────────
# 생성 중 출력 파일 추적 — 실패 시 미완성 파일 정리
# ──────────────────────────────────────────────
# 템플릿을 출력 경로로 복사한 뒤 편집 도중 예외가 나면, 가공되지 않은
# 템플릿 사본이 '대외비_….xlsx' 이름으로 출력 폴더에 남는다.
# 사용자가 그걸 갑지 원본으로 다시 선택하면 문제가 연쇄된다.
_inflight = threading.local()


def _copy_template(src: str, dst: str) -> str:
    """템플릿을 출력 경로로 복사하고 실패 시 정리 대상으로 등록한다."""
    shutil.copy2(src, dst)
    # copy2 는 읽기전용 속성까지 복사한다. 공유 드라이브의 마스터 템플릿이
    # 읽기전용이면 사본도 읽기전용이 되어 wb.save() 가 PermissionError 로 죽는다.
    try:
        os.chmod(dst, os.stat(dst).st_mode | stat.S_IWRITE)
    except OSError:
        pass
    _inflight.path = dst
    return dst


def _discard_inflight() -> None:
    """생성에 실패한 미완성 출력 파일을 삭제한다."""
    p = getattr(_inflight, "path", None)
    _inflight.path = None
    if p and os.path.exists(p):
        try:
            os.remove(p)
            logger.warning("생성 실패 — 미완성 파일 삭제: %s", p)
        except OSError as e:
            logger.warning("미완성 파일 삭제 실패 (%s): %s", p, e)


def _clear_inflight() -> None:
    """생성 성공 — 추적 해제."""
    _inflight.path = None


def inspect_template_losses(path: str) -> List[str]:
    """
    openpyxl 왕복(load→save) 시 소실되는 파트를 미리 찾아 사람이 읽을 목록으로 반환.

    openpyxl 은 아래를 보존하지 못하고 조용히 버린다. 생성된 견적서에서
    그림·머리글 로고·인쇄 설정이 사라져도 아무 오류가 나지 않으므로,
    통합양식을 읽는 시점에 미리 알려준다.

      - EMF/WMF 벡터 이미지 → 도형이 통째로 사라진다. 복원 불가.
                              엑셀에서 PNG 그림으로 교체해야 한다.

    ※ 머리글/바닥글 이미지(VML)도 openpyxl 이 버리지만
      restore_header_footer_images() 로 저장 후 되살리므로 경고하지 않는다.
    ※ 인쇄 설정은 용지·배율·방향·여백·printOptions 가 모두 보존된다.
      printerSettings*.bin(프린터 드라이버 고유값)만 빠지며 실사용에 영향 없다.

    반환: 경고 문구 리스트 (비어 있으면 소실 없음)
    """
    import zipfile
    from xml.etree import ElementTree as ET

    _RELNS = "{http://schemas.openxmlformats.org/package/2006/relationships}Relationship"
    out: List[str] = []
    try:
        with zipfile.ZipFile(path) as z:
            names = z.namelist()

            # ① EMF/WMF 이미지 — 어느 시트에서 쓰이는지까지 찾아 알려준다
            vec = [n for n in names
                   if n.startswith("xl/media/") and n.lower().endswith((".emf", ".wmf"))]
            if vec:
                def _rels(p):
                    try:
                        return {r.get("Id"): (r.get("Target") or "")
                                for r in ET.fromstring(z.read(p)).iter(_RELNS)}
                    except Exception:
                        return {}
                wbrel = _rels("xl/_rels/workbook.xml.rels")
                sheets = re.findall(
                    r'<sheet name="([^"]+)"[^>]*r:id="(rId\d+)"',
                    z.read("xl/workbook.xml").decode("utf8", "replace"))
                users = {}
                for sname, rid in sheets:
                    fn = (wbrel.get(rid) or "").split("/")[-1]
                    for tgt in _rels(f"xl/worksheets/_rels/{fn}.rels").values():
                        if "drawings/" not in tgt or not tgt.endswith(".xml"):
                            continue
                        dn = tgt.split("/")[-1]
                        for v in _rels(f"xl/drawings/_rels/{dn}.rels").values():
                            base = v.split("/")[-1]
                            if any(base == m.split("/")[-1] for m in vec):
                                users.setdefault(base, []).append(sname)
                for m in vec:
                    base = m.split("/")[-1]
                    where = ", ".join(users.get(base, [])) or "(사용 시트 불명)"
                    out.append(f"벡터 이미지 {base} → {where} 에서 사라집니다")


    except Exception as e:
        logger.warning("템플릿 소실 검사 실패 (%s): %s", path, e)
    return out


def _at(row: tuple, col: int) -> Any:
    """values_only 로 읽은 한 행에서 col 번째(1부터) 값. 행이 짧으면 None."""
    return row[col - 1] if 0 < col <= len(row) else None


def check_pump_price_consistency(path: str) -> List[str]:
    """
    앱이 쓰는 '품목' 단가와 템플릿 '사양서'가 VLOOKUP 하는 'Pump 단가표' 단가를 대조.

    [왜 필요한가]
    PUMP 단가가 두 경로로 갈라져 있다.
      · 앱   : 품목 시트의 분류(A열)=Q-Code 로 찾은 행들을 STEP5 에 PUMP 로 추가
      · 템플릿: 사양서!F15 = VLOOKUP(A15,'Pump 단가표'!F:M,6,0)
    둘이 어긋나면 인쇄되는 사양서 금액과 갑지 금액이 조용히 달라진다.
    (ver.1.8 실측: 102개 중 3개 불일치, 최대 2,840만원/CH 차이)

    반환: 불일치 설명 리스트 (비어 있으면 정상)
    """
    out: List[str] = []
    try:
        wb = load_workbook(path, read_only=True, data_only=True)
        try:
            cfg = quote_config.from_workbook(wb)
            sh_items, sh_pump = cfg["sheet.items"], cfg["sheet.pump_price"]
            if sh_items not in wb.sheetnames or sh_pump not in wb.sheetnames:
                return out

            i_cls, i_price = cfg.col("items.col_class"), cfg.col("items.col_price")
            by_class: Dict[str, List[float]] = {}
            for row in wb[sh_items].iter_rows(max_col=max(i_cls, i_price), values_only=True):
                key = s(_at(row, i_cls))
                if key:
                    by_class.setdefault(key, []).append(to_float(_at(row, i_price)))

            # Pump 단가표: C열=자재코드, F열=모델명, K열=단가
            # 템플릿 수식 체인을 그대로 흉내낸다:
            #   A15 = VLOOKUP(Q-Code, C:F, 4, 0)   → 모델명 (첫 일치 행)
            #   F15 = VLOOKUP(A15,    F:M, 6, 0)   → 단가   (첫 일치 행)
            # VLOOKUP 은 첫 일치만 쓰므로 중복 행이 있어도 첫 행 기준으로 본다.
            code_to_model: Dict[str, str] = {}
            model_to_price: Dict[str, float] = {}
            p_code, p_model, p_price = (cfg.col("pump.col_code"), cfg.col("pump.col_model"),
                                        cfg.col("pump.col_price"))
            for row in wb[sh_pump].iter_rows(max_col=max(p_code, p_model, p_price), values_only=True):
                code, model = s(_at(row, p_code)), s(_at(row, p_model))
                if code and code.upper().startswith("Q") and code not in code_to_model:
                    code_to_model[code] = model
                if model and model not in model_to_price:
                    model_to_price[model] = to_float(_at(row, p_price))

            for code, model in code_to_model.items():
                if code not in by_class:
                    continue
                tbl = model_to_price.get(model, 0.0)
                app = sum(by_class[code])
                if abs(app - tbl) < 1:
                    continue
                why = (f"품목 시트에 {len(by_class[code])}행이 있어 합산됨"
                       if len(by_class[code]) > 1 else "단가가 서로 다름")
                out.append(
                    f"{code}({model}): 사양서 {tbl:,.0f} vs 앱 {app:,.0f} "
                    f"(차이 {app - tbl:+,.0f}) — {why}"
                )
        finally:
            wb.close()
    except Exception as e:
        logger.warning("PUMP 단가 대조 실패 (%s): %s", path, e)
    return out


# ═══════════════════════════════════════════════════════════════
# 읽기 (openpyxl) — 기존과 동일
# ═══════════════════════════════════════════════════════════════

def parse_items_sheet(ws: Worksheet, cfg: Optional[quote_config.QuoteConfig] = None) -> Tuple[
    List[Dict[str, Any]],
    Dict[str, List[Dict[str, Any]]],
    Dict[str, float],
    Dict[str, int],
]:
    """품목 시트 파싱. 반환: (rows, by_class, price_by_spec, spec_order_index)"""
    # iter_rows(values_only=True) 사용: ws.cell(r,c) 임의 접근은 read_only
    # 워크시트에서 매 호출마다 스트림을 처음부터 다시 훑어 O(n²)이 된다
    # (3천 행 기준 100배+ 저하 실측). read_only=True 로 여는 호출부와 짝을 이룬다.
    def _is_header(a: str, b: str, c: str) -> bool:
        h = (a + b + c).replace(" ", "").lower()
        return "분류" in h and ("단가" in h or "가격" in h)

    cfg = cfg or quote_config.QuoteConfig()
    i_cls, i_spec, i_price = (cfg.col("items.col_class"), cfg.col("items.col_spec"),
                              cfg.col("items.col_price"))
    n_col = max(i_cls, i_spec, i_price)

    def _abc(row: tuple) -> tuple:
        return _at(row, i_cls), _at(row, i_spec), _at(row, i_price)

    row_iter = (_abc(r) for r in ws.iter_rows(max_col=n_col, values_only=True))
    try:
        r1a, r1b, r1c = next(row_iter)
    except StopIteration:
        return [], {}, {}, {}
    header = _is_header(s(r1a), s(r1b), s(r1c))
    if not header:
        row_iter = (_abc(r) for r in ws.iter_rows(max_col=n_col, values_only=True))   # 1행부터 다시

    rows: List[Dict[str, Any]] = []
    by_class: Dict[str, List[Dict[str, Any]]] = {}
    price_by_spec: Dict[str, float] = {}
    order_index: Dict[str, int] = {}
    order = 0

    for a_raw, b_raw, c in row_iter:
        a = s(a_raw)
        b = s(b_raw)
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


def _code_map_cols(cfg: quote_config.QuoteConfig) -> Tuple[int, int, int, int, int]:
    return tuple(cfg.col("codemap.col_" + k) for k in ("process", "vendor", "code5d", "acc", "qty"))


def parse_code_map_sheet(ws: Worksheet, cfg: Optional[quote_config.QuoteConfig] = None
                         ) -> Dict[Tuple[str, str, str], List[Tuple[str, float]]]:
    """코드매핑 시트 파싱. 반환: {(공정, 설비사, 5D): [(ACC, 수량), ...]}"""
    def _is_header(a, b, c, d, e) -> bool:
        h = (a + b + c + d + e).replace(" ", "").lower()
        return "공정" in h and "설비" in h and ("acc" in h or "수량" in h or "5d" in h)

    cols = _code_map_cols(cfg or quote_config.QuoteConfig())

    def _five(row: tuple) -> tuple:
        return tuple(_at(row, c) for c in cols)

    row_iter = (_five(r) for r in ws.iter_rows(max_col=max(cols), values_only=True))
    try:
        r1 = next(row_iter)
    except StopIteration:
        return {}
    if not _is_header(*(s(v) for v in r1)):
        row_iter = (_five(r) for r in ws.iter_rows(max_col=max(cols), values_only=True))   # 1행부터 다시

    cmap: Dict[Tuple[str, str, str], List[Tuple[str, float]]] = {}
    for proc_raw, vendor_raw, code5d_raw, acc_raw, qty_raw in row_iter:
        proc   = s(proc_raw)
        vendor = s(vendor_raw)
        code5d = s(code5d_raw)
        acc    = s(acc_raw)
        qty    = to_float(qty_raw)
        if proc and vendor and code5d and acc:
            cmap.setdefault((proc, vendor, code5d), []).append((acc, qty))
    return cmap


_BACKUP_KEEP = 20


def _prune_backups(backup_dir: str, base: str, keep: int = _BACKUP_KEEP) -> None:
    """Backup/ 에 같은 양식의 백업이 끝없이 쌓이지 않도록 최신 keep 개만 남긴다.
    정리 실패는 본 작업에 영향을 주지 않는다."""
    try:
        files = [os.path.join(backup_dir, n) for n in os.listdir(backup_dir)
                 if n.startswith(base + ".backup_")]
        files.sort(key=os.path.getmtime, reverse=True)
        for p in files[keep:]:
            try:
                os.remove(p)
            except OSError:
                pass
    except OSError:
        pass


def append_code_map_rows(
    template_path: str,
    process: str,
    vendor: str,
    code5d: str,
    items: List[Tuple[str, float]],
) -> Tuple[List[Tuple[str, float]], List[Tuple[str, float]]]:
    """
    통합양식(마스터 템플릿) 코드매핑 시트에 (공정,설비사,5D,품목,수량) 행을 추가한다.
    items 중 (공정,설비사,5D,품목,수량)이 완전히 동일한 행이 이미 있으면 그 항목은
    건너뛴다. 새로 추가할 행이 하나도 없으면(전부 중복) 파일을 저장하지 않는다.

    template_path 는 여러 견적서가 공유하는 마스터 파일이라 직접 덮어쓴다.
    openpyxl 저장은 머리글/바닥글 그림(VML)을 읽지도 쓰지도 못해 저장 시
    사라지므로, 저장 전에 템플릿과 같은 폴더의 Backup/ 하위에 만든 백업으로
    restore_header_footer_images 를 호출해 되살린다 — 이 백업은 복원용인 동시에
    만약을 대비한 안전장치이기도 하다. 저장 자체가 실패하면(파일이 다른
    프로그램에서 열려 있는 등) 백업을 지우고 예외를 그대로 올린다.

    반환: (added, skipped) — 실제로 추가된 (품목,수량) 목록 / 중복이라 건너뛴 목록.
    """
    wb = load_workbook(template_path)
    cfg = quote_config.from_workbook(wb)
    ws = wb[cfg["sheet.code_map"]]
    existing = parse_code_map_sheet(ws, cfg)
    key = (process, vendor, code5d)
    seen = set(existing.get(key, []))

    added: List[Tuple[str, float]] = []
    skipped: List[Tuple[str, float]] = []
    for spec, qty in items:
        pair = (spec, qty)
        if pair in seen:
            skipped.append(pair)
            continue
        added.append(pair)
        seen.add(pair)   # 이번 호출 안에서의 자체 중복도 방지

    if not added:
        wb.close()
        return added, skipped

    backup_dir = ensure_dir(os.path.join(os.path.dirname(template_path) or ".", "Backup"))
    base, ext = os.path.splitext(os.path.basename(template_path))
    if not base.startswith("대외비_"):
        base = f"대외비_{base}"
    backup_path = unique_path(
        os.path.join(backup_dir, f"{base}.backup_{datetime.now().strftime('%Y%m%d_%H%M%S')}{ext}")
    )
    shutil.copy2(template_path, backup_path)
    _prune_backups(backup_dir, base)

    row = ws.max_row + 1
    for spec, qty in added:
        qty_val = int(qty) if float(qty).is_integer() else qty
        for col, value in zip(_code_map_cols(cfg), (process, vendor, code5d, spec, qty_val)):
            ws.cell(row, col, value)
        row += 1

    try:
        wb.save(template_path)
    except Exception:
        # 저장이 실패하면(권한 없음·다른 프로그램에서 열림 등) 원본은 그대로이므로
        # 방금 만든 백업도 쓸모가 없다 — 지워서 Backup 폴더에 쌓이지 않게 한다.
        wb.close()
        try:
            os.remove(backup_path)
        except OSError:
            pass
        raise
    wb.close()
    restore_header_footer_images(backup_path, template_path)
    return added, skipped


# 의뢰 행 dict 의 키(예전부터 열 문자로 불러 온 이름) ← 설정의 어느 열에서 읽는지.
# 열 위치는 설정으로 옮겨도 키 이름은 그대로 둬서 이 dict 를 쓰는 곳을 건드리지 않는다.
_REQ_FIELDS = (
    ("D", "req.col_pr"), ("E", "req.col_item"), ("F", "req.col_qcode"), ("G", "req.col_desc"),
    ("H", "req.col_qty"), ("J", "req.col_requester"), ("K", "req.col_line"),
    ("N", "req.col_subproc"), ("R", "req.col_concept"), ("V", "req.col_5d"),
    ("X", "req.col_maker"), ("Y", "req.col_model"), ("Z", "req.col_hogi"), ("AA", "req.col_due"),
)


def parse_request_xlsx(path: str, cfg: Optional[quote_config.QuoteConfig] = None
                       ) -> Tuple[str, List[Dict[str, Any]]]:
    """
    의뢰파일 파싱.
    반환: (active_sheet_name, row_dicts)
    row_dict keys: row_idx, D, E, F, G, H, J, K, N, V, X, Y, Z, AA
    """
    # read_only=True + iter_rows(): ws.cell(r,c) 임의 접근을 read_only 모드에서
    # 쓰면 매 호출이 스트림을 처음부터 다시 훑어 O(n²)이 된다(실측: 3천 행에서
    # 100배+ 저하). 의뢰파일은 수백~수천 행이 흔하므로 반드시 iter_rows.
    wb = load_workbook(path, read_only=True, data_only=True)
    ws = wb.active
    sheet_name = ws.title
    cfg = cfg or quote_config.QuoteConfig()
    fields = [(name, cfg.col(key)) for name, key in _REQ_FIELDS]
    NCOL = max(c for _, c in fields)
    i_hogi = cfg.col("req.col_hogi")

    def _is_header(r1: tuple) -> bool:
        checks = [s(_at(r1, cfg.col(k))).replace(" ", "").lower()
                  for k in ("req.col_hogi", "req.col_line", "req.col_maker", "req.col_5d")]
        return any("설비" in c or "공정" in c or "5d" in c or "코드" in c for c in checks)

    row_iter = ws.iter_rows(max_col=NCOL, values_only=True)
    try:
        first = next(row_iter)
    except StopIteration:
        wb.close()
        return sheet_name, []
    start = 2 if _is_header(first) else 1
    if start == 1:
        row_iter = ws.iter_rows(max_col=NCOL, values_only=True)   # 1행부터 다시

    rows: List[Dict[str, Any]] = []
    for idx, vals in enumerate(row_iter, start=start):
        if s(_at(vals, i_hogi)) == "":      # 호기가 비어 있는 행은 의뢰가 아니다
            continue
        rd: Dict[str, Any] = {"row_idx": idx}
        for name, col in fields:
            rd[name] = _at(vals, col)
        rows.append(rd)
    wb.close()
    return sheet_name, rows


# ═══════════════════════════════════════════════════════════════
# openpyxl 공통 헬퍼
# ═══════════════════════════════════════════════════════════════

def _hide_rows(ws: Worksheet, start: int, end: int, filled_count: int) -> None:
    """start~end 행 중 filled_count 이후 행을 숨김."""
    for i in range(end - start + 1):
        ws.row_dimensions[start + i].hidden = (i >= filled_count)


def _follow_mirror(ws, value: Any, depth: int = 3) -> Any:
    """'=E4' 처럼 같은 시트의 셀 하나만 가리키는 수식이면 그 셀을 따라간다.

    Pump 단가표의 G/X/AA 열이 이런 미러 수식(G=C, X=E, AA=F)이다.
    openpyxl 은 수식을 계산하지 않으므로 캐시된 계산값을 쓸 수도 있지만,
    실측 결과 X열 캐시가 원본과 어긋난 행이 2건 있었다. 엑셀은 파일을 열 때
    재계산(fullCalcOnLoad)하므로 캐시가 아니라 원본을 따라가야 지금과 같은
    값이 나온다.
    """
    for _ in range(depth):
        if not (isinstance(value, str) and value.startswith("=")):
            return value
        m = re.fullmatch(r"=\$?([A-Z]{1,3})\$?(\d+)", value.strip())
        if not m:
            return None
        value = ws[f"{m.group(1)}{m.group(2)}"].value
    return None


def _vlookup_text(v: Any) -> str:
    """VLOOKUP 의 비교 대상/결과를 원본 그대로 다룬다.

    s() 처럼 앞뒤 공백을 털면 안 된다 — 단가표에는 모델명 'DD1055L '(뒤 공백),
    분류 'Gate Valve\\n\\n'(줄바꿈) 같은 값이 실제로 들어 있어서, 털어버리면
    지금 인쇄되는 사양서 내용과 달라진다.
    """
    if v is None:
        return ""
    return v if isinstance(v, str) else str(v)


def _build_price_lookups(wb) -> Dict[str, Dict[str, Any]]:
    """사양서·입고검수확인서가 VLOOKUP 으로 참조하던 단가표 조회를 파이썬으로 옮긴다.

    VLOOKUP(.., 0) 과 똑같이 '처음 일치하는 행'만 쓴다(단가표에 같은 키가
    여러 번 나오는 행이 실제로 있다).

    반환 키 — 각 값은 {조회키: 결과} dict:
      code_model  Q-Code → 모델명            (사양서!A15)
      model_price 모델명 → 단가              (사양서!F15)
      code_spec   Q-Code → 규격(용량外)       (사양서!B15, X열=E열)
      code_dry    Q-Code → DRY/BOOSTER      (입고검수확인서!D17)
      code_cap    Q-Code → 규격              (입고검수확인서!D18)
      code_motor  Q-Code → MOTOR TYPE       (입고검수확인서!D19)
      spec_class  규격   → 분류              (사양서!A17:A41)

    결과는 워크북에 캐시한다. 일괄 생성은 워크북 하나를 여러 건에 재사용하는데,
    첫 건에서 값을 굳히고 나면 단가표를 참조하는 수식이 사라져 그 시트가
    삭제되고, 두 번째 건부터는 조회가 전부 비어 잘못된 견적서가 나간다
    (실측: 2건째부터 분류가 "1", 모델명이 "미등록 Q-Code" 로 떨어짐).
    """
    cached = getattr(wb, "_lot_price_lookups", None)
    if cached is not None:
        return cached

    out = {k: {} for k in ("code_model", "model_price", "code_spec",
                           "code_dry", "code_cap", "code_motor", "spec_class")}
    cfg = quote_config.from_workbook(wb)

    if cfg["sheet.pump_price"] in wb.sheetnames:
        ws = wb[cfg["sheet.pump_price"]]
        for row in ws.iter_rows():
            def val(letter: str) -> str:
                cell = ws[f"{letter}{row[0].row}"]
                return _vlookup_text(_follow_mirror(ws, cell.value))
            def raw(letter: str) -> Any:
                # 조회 '결과'는 글자로 바꾸지 않는다. 숫자를 문자열로 굳히면
                # 셀 서식(예: 입고검수확인서!D18 의 ##,###"L/min")이 먹지 않아
                # "20,000L/min" 이 "20000" 으로 찍힌다.
                v = _follow_mirror(ws, ws[f"{letter}{row[0].row}"].value)
                return "" if v is None else v
            code, model = val(cfg["pump.col_code"]), val(cfg["pump.col_model"])
            if code:
                out["code_model"].setdefault(code, model)
                out["code_spec"].setdefault(code, raw(cfg["pump.col_spec"]))
                out["code_dry"].setdefault(code, raw(cfg["pump.col_dry"]))
                out["code_cap"].setdefault(code, raw(cfg["pump.col_cap"]))
                out["code_motor"].setdefault(code, raw(cfg["pump.col_motor"]))
            if model:
                out["model_price"].setdefault(
                    model, to_float(_follow_mirror(
                        ws, ws[cfg.at("pump.col_price", row[0].row)].value)))

    if cfg["sheet.acc_price"] in wb.sheetnames:
        ws = wb[cfg["sheet.acc_price"]]
        a_spec, a_class = cfg.col("acc.col_spec"), cfg.col("acc.col_class")
        for row in ws.iter_rows(max_col=max(a_spec, a_class), values_only=True):
            spec = _vlookup_text(_at(row, a_spec))
            if spec:
                out["spec_class"].setdefault(spec, _vlookup_text(_at(row, a_class)))

    wb._lot_price_lookups = out
    return out


def _referenced_sheets(wb, names: List[str]) -> set:
    """names 시트들이 수식으로 참조하는 시트를 전이적으로 모아 돌려준다.

    통합양식 ver.2.0 기준으로 정의된이름·데이터유효성·조건부서식에는
    시트 간 참조가 없고 INDIRECT 도 쓰이지 않는 것을 확인했으므로,
    수식 문자열만 훑으면 의존 관계가 전부 잡힌다. 템플릿이 바뀌어
    참조가 늘어나도 실행 시점에 다시 계산하므로 목록을 손댈 필요가 없다.
    """
    sheet_names = list(wb.sheetnames)
    patterns = [(nm, re.compile(r"'?" + re.escape(nm) + r"'?!")) for nm in sheet_names]

    def _direct(title: str) -> set:
        ws = wb[title]
        found = set()
        for row in ws.iter_rows():
            for cell in row:
                v = cell.value
                if not isinstance(v, str) or not v.startswith("="):
                    continue
                for nm, pat in patterns:
                    if nm != title and nm not in found and pat.search(v):
                        found.add(nm)
        return found

    need, stack = set(names), [n for n in names if n in sheet_names]
    while stack:
        for dep in _direct(stack.pop()):
            if dep not in need:
                need.add(dep)
                stack.append(dep)
    return need


def _keep_only_sheets(wb, keep: List[str]) -> None:
    """keep(+ keep 이 수식으로 참조하는 시트)만 남기고 나머지는 '삭제'한다.

    이전에는 veryHidden 으로 숨기기만 했는데, 숨김은 파일 안에 그대로
    남는 것이라 ① 대외비 견적서에 전사 단가표·코드매핑이 실려 나가고
    ② 용량도 전혀 줄지 않았다(견적서 1건 994KB 중 실제 문서 내용은 4%).
    """
    need = _referenced_sheets(wb, keep)
    for name in list(wb.sheetnames):
        if name in need:
            wb[name].sheet_state = "visible" if name in keep else "veryHidden"
        elif len(wb.sheetnames) > 1:
            try:
                del wb[name]
            except Exception as e:
                logger.warning("시트 삭제 실패 (%s): %s", name, e)
                wb[name].sheet_state = "veryHidden"


_VML_RELTYPE = ("http://schemas.openxmlformats.org/officeDocument/2006/"
                "relationships/vmlDrawing")


def restore_header_footer_images(src_template: str, out_path: str) -> int:
    """
    openpyxl 이 버린 머리글/바닥글 이미지(VML)를 원본 템플릿에서 되살린다.

    openpyxl 은 legacyDrawingHF(머리글/바닥글 그림)를 읽지도 쓰지도 못한다.
    머리글 텍스트의 '&G'(그림 삽입) 지시자는 남는데 그림 파트가 사라져
    인쇄물에서 로고가 빠진다. 저장이 끝난 파일에 원본의 VML 파트와
    참조만 되살린다(셀 내용은 건드리지 않는다).

    반환: 복원한 시트 수
    """
    import zipfile
    from xml.etree import ElementTree as ET
    _RELNS = "{http://schemas.openxmlformats.org/package/2006/relationships}Relationship"

    def _rels(z, p):
        try:
            return [(r.get("Id"), r.get("Type"), r.get("Target"))
                    for r in ET.fromstring(z.read(p)).iter(_RELNS)]
        except Exception:
            return []

    def _sheetfiles(z):
        wbrel = {i: t for i, _t, t in _rels(z, "xl/_rels/workbook.xml.rels")}
        return {n: (wbrel.get(r) or "").split("/")[-1] for n, r in
                re.findall(r'<sheet name="([^"]+)"[^>]*r:id="(rId\d+)"',
                           z.read("xl/workbook.xml").decode("utf8", "replace"))}

    try:
        with zipfile.ZipFile(src_template) as zs:
            src_names = set(zs.namelist())
            want = {}
            for name, fn in _sheetfiles(zs).items():
                for _id, ty, tgt in _rels(zs, f"xl/worksheets/_rels/{fn}.rels"):
                    if ty != _VML_RELTYPE:
                        continue
                    vml = tgt.split("/")[-1]
                    parts = {f"xl/drawings/{vml}": zs.read(f"xl/drawings/{vml}")}
                    rp = f"xl/drawings/_rels/{vml}.rels"
                    if rp in src_names:
                        parts[rp] = zs.read(rp)
                        for _i, _t, mt in _rels(zs, rp):
                            m = "xl/" + mt.replace("../", "")
                            if m in src_names:
                                parts[m] = zs.read(m)
                    want[name] = (vml, parts)
        if not want:
            return 0

        with zipfile.ZipFile(out_path) as zo:
            omap = _sheetfiles(zo)
            data = {n: zo.read(n) for n in zo.namelist()}

        added = 0
        for name, (vml, parts) in want.items():
            fn = omap.get(name)
            sheet = f"xl/worksheets/{fn}" if fn else None
            if not sheet or sheet not in data:
                continue           # 삭제된 시트는 건너뛴다
            data.update(parts)

            rp = f"xl/worksheets/_rels/{fn}.rels"
            cur = data.get(
                rp,
                b'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
                b'<Relationships xmlns="http://schemas.openxmlformats.org/'
                b'package/2006/relationships"></Relationships>').decode("utf8")
            used = set(re.findall(r'Id="(rId\d+)"', cur))
            rid = next(f"rId{i}" for i in range(1, 10000) if f"rId{i}" not in used)
            data[rp] = cur.replace(
                "</Relationships>",
                f'<Relationship Id="{rid}" Type="{_VML_RELTYPE}" '
                f'Target="../drawings/{vml}"/></Relationships>').encode("utf8")

            x = data[sheet].decode("utf8")
            if "<legacyDrawingHF" not in x:
                # openpyxl 은 r: 네임스페이스를 선언하지 않는 시트가 있어
                # 그냥 r:id 를 쓰면 'unbound prefix' 로 파일이 깨진다 → 인라인 선언.
                tag = (f'<legacyDrawingHF xmlns:r="http://schemas.openxmlformats.org'
                       f'/officeDocument/2006/relationships" r:id="{rid}"/>')
                # 스키마 순서: drawing → legacyDrawing → legacyDrawingHF
                m = re.search(r'<legacyDrawing\b[^>]*/>|<drawing\b[^>]*/>', x)
                x = (x[:m.end()] + tag + x[m.end():]) if m else \
                    x.replace("</worksheet>", tag + "</worksheet>")
                data[sheet] = x.encode("utf8")
            added += 1

        ct = data["[Content_Types].xml"].decode("utf8")
        for ext, mime in (
            ("vml", "application/vnd.openxmlformats-officedocument.vmlDrawing"),
            ("jpeg", "image/jpeg"), ("jpg", "image/jpeg"), ("png", "image/png"),
            ("gif", "image/gif"), ("emf", "image/x-emf"),
        ):
            if f'Extension="{ext}"' not in ct:
                ct = ct.replace(
                    "</Types>",
                    f'<Default Extension="{ext}" ContentType="{mime}"/></Types>')
        data["[Content_Types].xml"] = ct.encode("utf8")

        tmp = out_path + ".tmp"
        with zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as zw:
            for n, d in data.items():
                zw.writestr(n, d)
        os.replace(tmp, out_path)
        return added
    except Exception as e:
        # 복원은 부가 기능이다. 실패해도 생성물 자체는 유효하므로 로그만 남긴다.
        logger.warning("머리글/바닥글 이미지 복원 실패 (%s): %s", out_path, e)
        try:
            if os.path.exists(out_path + ".tmp"):
                os.remove(out_path + ".tmp")
        except OSError:
            pass
        return 0


def _reset_sheet_selections(wb) -> None:
    """
    모든 시트의 SheetView selection을 A1 단일 셀로 초기화한다.

    [문제]
    openpyxl은 템플릿에 저장된 selection.sqref 값을 그대로 보존한다.
    템플릿이 다중 비연속 범위가 선택된 채로 저장되어 있으면(예: "A1 C3:D5"),
    생성된 xlsx도 동일한 상태를 갖게 된다.
    Excel이 이 파일을 열면 다중 선택 상태가 복원되어:
      - 복사·붙여넣기: "이 작업은 다중선택 범위에서 작동하지 않습니다"
      - 시트 삽입·삭제: 차단
      - 시트 이동/복사: Excel 비정상 종료 (PrintArea 정의명과 충돌)

    [수정]
    저장 직전에 모든 시트의 selection 을 단일 A1 으로 강제 초기화한다.
    """
    from openpyxl.worksheet.views import Selection
    for ws in wb.worksheets:
        sv = ws.sheet_view
        if not sv.selection:
            sv.selection.append(Selection(activeCell="A1", sqref="A1"))
        else:
            first = sv.selection[0]
            first.activeCell = "A1"
            first.sqref = "A1"
            sv.selection[:] = [first]


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

    cfg = quote_config.from_workbook(wb_copy)

    def _copy_row(req_ws) -> None:
        copy_ws = wb_copy[cfg["sheet.req_copy"]]
        dst_row = cfg["copy.row"]
        for c in cfg.cols("copy.cols"):
            copy_ws.cell(dst_row, c).value = req_ws.cell(row_idx, c).value

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

def _pin_images(wb) -> None:
    """
    이미지 바이트를 메모리에 고정해 같은 워크북을 여러 번 저장할 수 있게 한다.

    openpyxl 은 첫 save 때 이미지 원본 파일 핸들을 닫아버리므로, 두 번째
    save 에서 'ValueError: I/O operation on closed file' 로 죽는다.
    일괄 생성에서 워크북을 재사용하려면 반드시 선행해야 한다.
    """
    import io
    for ws in wb.worksheets:
        for img in getattr(ws, "_images", []):
            try:
                data = img._data()
            except Exception:
                continue
            img.ref = io.BytesIO(data)
            img._data = lambda _d=data: _d


def open_template_for_reuse(template_path: str):
    """일괄 생성용 — 템플릿을 1회만 읽어 재사용 가능한 워크북으로 돌려준다."""
    wb = load_workbook(template_path)
    _pin_images(wb)
    return wb


def _split_credits(credits: List[Dict[str, Any]], cfg: quote_config.QuoteConfig
                   ) -> Tuple[List[Dict[str, Any]], List[Dict[str, Any]]]:
    """Credit 을 (Pump Credit, Rack Credit) 으로 나눈다 — 규격 이름에 든 말로 가른다."""
    kw = cfg["credit.pump_keyword"]
    pump = [r for r in credits if kw in r.get("spec", "")]
    rack = [r for r in credits if kw not in r.get("spec", "")]
    return pump, rack


def _fill_domestic(state: QuoteState,
                   rd: Dict[str, Any],
                   items: List[Dict[str, Any]],
                   req_ws_cache: Optional[Any] = None,
                   wb_cache: Optional[Any] = None) -> str:
    """
    국내 견적서: 템플릿을 복사한 뒤 openpyxl 로 데이터 기입.
    반환: 저장된 xlsx 경로.

    wb_cache: open_template_for_reuse() 로 만든 워크북. 주면 건마다 템플릿을
              다시 읽지 않고 재사용한다(일괄 생성에서 60% 이상 단축).
              건별 상태는 아래 ②③④⑤ 에서 전부 덮어쓰므로 누수되지 않지만,
              도중에 예외가 나면 워크북이 더러워지므로 호출자가 다시 열어야 한다.
    """
    cfg = quote_config.from_workbook(wb_cache) if wb_cache is not None \
        else quote_config.load(state.template_path)
    names = {
        "날짜":   datetime.now().strftime("%y%m%d"),
        "PR":     safe_filename(s(rd.get("D"))),
        "항번":   safe_filename(s(rd.get("E"))),
        "라인":   safe_filename(s(rd.get("K"))),
        "의뢰자": safe_filename(s(rd.get("J"))),
        "호기":   safe_filename(s(rd.get("Z"))),
    }
    folder = ensure_dir(
        os.path.join(exe_dir(), cfg["name.out_dir"], cfg.text("name.dom_folder", **names))
    )
    path = unique_path(os.path.join(folder, cfg.text("name.dom_file", **names)))

    # ① 템플릿 준비 — 재사용 워크북이 있으면 복사·로드를 건너뛴다
    if wb_cache is not None:
        wb = wb_cache
        _inflight.path = path      # 저장 중 실패 시 정리 대상으로 등록
    else:
        # 템플릿 전체를 복사 — 이미지·서식·수식 구조 모두 보존
        _copy_template(state.template_path, path)
        wb = load_workbook(path)

    ws_spec = wb[cfg["sheet.spec"]]
    ws_sign = wb[cfg["sheet.sign_spec"]]
    ws_incoming = wb[cfg["sheet.incoming"]]
    spec_start, spec_end = cfg["dom.spec_start"], cfg["dom.spec_end"]
    spec_rows = spec_end - spec_start + 1                 # 사양서 품목 칸 수
    pump_row = cfg["dom.pump_row"]
    c_class, c_spec, c_qty, c_price, c_amt = (
        cfg["dom.col_class"], cfg["dom.col_spec"], cfg["dom.col_qty"],
        cfg["dom.col_price"], cfg["dom.col_amt"])

    # ② 의뢰파일 행 → 견적의뢰복사본 시트 복사
    _write_req_row_openpyxl(wb, state, rd, req_ws_cache)

    # ③ 투자자명
    investor = s(state.investor_name) or cfg["default.investor"]
    ws_spec[cfg["dom.cell_investor"]] = cfg.text("dom.investor_text", 투자자=investor)

    # ③-1 옵션에서 지정한 설비수량/실반입라인/Exh Size — 값이 있을 때만 기입.
    #     0/빈 문자열(=옵션 미설정)이면 템플릿 원본 셀을 건드리지 않는다.
    if state.equip_qty:
        ws_spec[cfg["dom.cell_equip_qty"]] = int(state.equip_qty)
    if state.actual_line:
        ws_sign[cfg["dom.cell_actual_line"]] = state.actual_line
    if state.exh_size:
        ws_incoming[cfg["dom.cell_exh_size"]] = int(state.exh_size)

    rack_items = [r for r in items if r["cat"] != "PUMP" and r["role"] != "CREDIT"]
    credits    = [r for r in items if r["role"] == "CREDIT"]

    # ④ 기존 데이터 클리어
    for rr in range(spec_start, spec_end + 1):
        for col in (c_spec, c_qty, c_price, c_amt):
            ws_spec[f"{col}{rr}"] = None

    # ⑤ 품목 기입
    #
    # [템플릿(ver.1.8) 실제 수식 구조]
    #     사양서!E15 = 견적의뢰복사본!H2                     (수량)
    #     사양서!F15 = VLOOKUP(A15,'Pump 단가표'!F:M,6,0)    (PUMP CH당 단가)
    #     사양서!G15 = E15*F15                              ← PUMP 총액
    #     사양서!G16 = SUM(G17:G41)                         ← Rack system 합계
    #     사양서!F43 = ROUNDDOWN((G16+G15)/견적의뢰복사본!H2,-3)
    #
    # [Pump Credit 처리]
    # Pump Credit 을 G17:G41 에 실금액으로 넣으면 G16(=Rack system 합계)에서
    # 차감되어, PUMP 값이어야 할 금액이 Rack 쪽에서 빠지는 것처럼 보인다.
    # 그래서 PUMP 는 credit 반영된 CH당 단가를 F15 에 직접 기입해 G15 에서
    # 차감되도록 하고, 아래 라인은 표시 전용(G=0)으로 남긴다.
    #   · Rack Credit → 기존대로 G열에 실금액 (G16 에서 차감)
    #   · Pump Credit → F15 하드코딩 + 표시행 G=0 (이중 차감 방지)
    pump_credits, rack_credits = _split_credits(credits, cfg)

    pump_items      = [r for r in items if r["cat"] == "PUMP" and r["role"] != "CREDIT"]
    pump_spec       = pump_items[0]["spec"] if pump_items else None
    pump_amt        = sum(r["amt"] for r in pump_items)
    rack_amt        = sum(r["amt"] for r in rack_items)
    pump_credit_amt = sum(c["amt"] for c in pump_credits)   # 통상 음수
    rack_credit_amt = sum(c["amt"] for c in rack_credits)   # 통상 음수

    h_qty = to_float(rd.get("H"))
    if h_qty <= 0:
        h_qty = 1.0

    # PUMP CH당 단가 — credit 반영 후 정수(원)로 확정한다.
    # 소수점이 남으면 견적서에 13,900,000.333 같은 값이 찍히고,
    # G15=E15*F15 결과도 원 단위로 떨어지지 않는다.
    p_unit_net = round((pump_amt + pump_credit_amt) / h_qty)
    # 표시용 CH당 Credit 단가 (음수)
    _pump_credit_unit = round(pump_credit_amt / h_qty) if pump_credits else 0

    out_list = rack_items[:spec_rows]
    for cr in rack_credits:
        out_list.append({"spec": cr["spec"], "qty": 0.0, "price": 0.0, "amt": cr["amt"]})
    for cr in pump_credits:
        # 표시 전용: CH당 단가는 F열에, 금액(G열)은 0 — G16 에 영향 없음
        out_list.append({
            "spec": cr["spec"], "qty": 0.0,
            "price": _pump_credit_unit, "amt": 0, "display_only": True,
        })
    out_list = out_list[:spec_rows]

    # 템플릿 용량 초과 시 조용히 잘리면 사양서 합계와 갑지 금액이 어긋난다.
    if len(rack_items) + len(credits) > spec_rows:
        raise ValueError(
            f"사양서 품목 칸({spec_rows}행)을 초과했습니다: "
            f"품목 {len(rack_items)}건 + Credit {len(credits)}건. "
            f"품목을 줄이거나 템플릿 행을 늘려주세요."
        )

    def _won(v) -> int:
        """금액은 원 단위 정수로 — 소수점이 찍히지 않게 한다."""
        return int(round(to_float(v)))

    for i, r in enumerate(out_list):
        rr = spec_start + i
        ws_spec[f"{c_spec}{rr}"] = r["spec"]
        if r.get("display_only"):
            # Pump Credit 표시행: CH당 단가만 보여주고 금액은 0
            # (F15 에서 이미 차감했으므로 G16 에 또 넣으면 이중 차감)
            ws_spec[f"{c_qty}{rr}"]   = None
            ws_spec[f"{c_price}{rr}"] = _won(r["price"])
            ws_spec[f"{c_amt}{rr}"]   = 0
        elif r.get("qty") == 0.0 and r.get("price") == 0.0 and "amt" in r:
            ws_spec[f"{c_amt}{rr}"] = _won(r["amt"])
        else:
            ws_spec[f"{c_qty}{rr}"]   = r["qty"]
            ws_spec[f"{c_price}{rr}"]  = _won(r["price"])
            ws_spec[f"{c_amt}{rr}"]    = _won(r["amt"])

    # ⑤-b PUMP CH당 단가를 F15 에 직접 기입 — Pump Credit 을 G15 에서 차감시킨다.
    # 템플릿 F15 는 'Pump 단가표' VLOOKUP 수식이지만, credit 이 있으면 그 값을
    # 그대로 두면 credit 이 반영되지 않는다. 값으로 덮어써 G15=E15*F15 가
    # net 금액이 되게 한다. credit 이 없으면 수식을 그대로 둔다.
    if pump_credits:
        ws_spec[f"{c_price}{pump_row}"] = p_unit_net

    # ⑤-a 견적의뢰복사본 수식 셀 → 계산값 직접 기입 ─────────────────────────
    # openpyxl 저장 시 수식 캐시(cached value)가 소멸된다.
    # generate_cover 에서 data_only=True 로 읽으면 수식 셀(O/P/Q/S/T/U)이 None 반환 →
    # 갑지DATA 해당 열 공백. 계산값을 값(value)으로 직접 써서 이 경로를 막는다.
    #
    # 재현 수식:
    #   S2 = 사양서!F43,  T2 = S2*H2,  U2 = ROUNDUP((P2*H2+Q2)/H2,-3)
    #
    # 아래 값들은 Excel 이 재계산할 값과 반드시 일치해야 한다.
    # 어긋나면 인쇄되는 사양서와 갑지 금액이 서로 다른 견적서가 나간다.
    copy_ws = wb[cfg["sheet.req_copy"]]
    copy_row = cfg["copy.row"]

    # 템플릿 수식 그대로 재현한다. Excel 이 재계산할 값과 어긋나면
    # 인쇄되는 사양서와 갑지 금액이 서로 다른 견적서가 나간다.
    #   G15 = E15*F15  = H2 × F15
    #        · Pump Credit 있음 → F15 는 위에서 기입한 p_unit_net (credit 반영)
    #        · 없음            → F15 는 Pump 단가표 VLOOKUP (= pump 단가)
    #   G16 = SUM(G17:G41) = rack_items + Rack Credit  (Pump Credit 은 0 으로 표시만)
    #   F43 = ROUNDDOWN((G16+G15)/H2, -3)
    #
    # ※ p_unit_net 은 이미 정수로 반올림했으므로 g15 도 그 값에서 직접 만든다.
    #   (A/h)*h 식으로 되돌리면 부동소수점 1 ulp 오차로 floor 가 1,000원
    #   낮게 떨어질 수 있다.
    if pump_credits:
        p_unit = p_unit_net                 # F15 에 실제로 쓴 값
        g15    = p_unit_net * h_qty         # Excel 의 E15*F15 와 정확히 동일
    else:
        p_unit = round(pump_amt / h_qty) if h_qty else 0
        g15    = pump_amt

    g16   = rack_amt + rack_credit_amt      # Pump Credit 은 G열 0 이라 미포함
    q_val = g16                             # Q2 = 사양서!G16

    # S = 사양서!F43 = ROUNDDOWN((G15+G16)/H2, -3)
    s_raw = (g15 + g16) / h_qty
    s_val = math.floor(s_raw / 1000) * 1000   # ROUNDDOWN(..., -3)
    # T = S2 * H2
    t_val = s_val * h_qty
    # U = ROUNDUP((P2*H2+Q2)/H2, -3) — 템플릿 수식 그대로 (Q2 경유로 credit 반영)
    u_raw = (p_unit * h_qty + q_val) / h_qty
    u_val = math.ceil(u_raw / 1000) * 1000    # ROUNDUP(..., -3)

    # 금액은 모두 원 단위 정수로 기입한다 — 소수점이 남으면 갑지·견적서에
    # 13,900,000.333 같은 값이 찍힌다.
    for key, value in (("copy.col_pump_model", pump_spec),     # O: PUMP 메인모듈 (=사양서!A15)
                       ("copy.col_pump_price", int(p_unit)),   # P: PUMP CH당 단가 (=사양서!F15)
                       ("copy.col_rack_sum",   int(q_val)),    # Q: Rack 합계 (=사양서!G16)
                       ("copy.col_unit",       int(s_val)),    # S: 견적단가 (=사양서!F43)
                       ("copy.col_amount",     int(t_val)),    # T: 견적금액 (=S2*H2)
                       ("copy.col_check",      int(u_val))):   # U: 견적단가(Check)
        copy_ws.cell(copy_row, cfg.col(key)).value = value

    # ⑤-b 단가표 VLOOKUP 수식 → 계산값 직접 기입 ─────────────────────────
    # 사양서·입고검수확인서가 'Pump 단가표'/'악세서리 단가표' 를 VLOOKUP 하는
    # 수식 31개를 값으로 바꾼다. 이 수식들 때문에 대외비 견적서에 전사 단가표를
    # 통째로 실어 보내야 했다(빼면 #REF!). 값으로 굳히면 단가표를 지울 수 있다.
    #
    # VLOOKUP 결과와 한 글자도 달라지면 안 되므로, 조회는 단가표 원본에서
    # 그대로 하고 IFERROR 기본값("미등록 Q-Code"/"-"/"0"/"1" 등)도 수식에
    # 적힌 문자열을 그대로 쓴다.
    lk = _build_price_lookups(wb)
    qcode = _vlookup_text(rd.get("F"))                        # 견적의뢰복사본!F2
    a15 = lk["code_model"].get(qcode) or cfg["dom.text_no_model"]
    ws_spec[f"{c_class}{pump_row}"] = a15
    ws_spec[f"{c_spec}{pump_row}"] = lk["code_spec"].get(qcode) or cfg["dom.text_no_spec"]
    # F15 는 credit 이 있으면 위에서 이미 p_unit_net 을 기입했다. 없을 때만
    # 단가표 조회값을 넣는다 — 지금 동작(수식 결과)과 같게 유지하기 위함이다.
    if not pump_credits:
        ws_spec[f"{c_price}{pump_row}"] = _won(lk["model_price"].get(a15, 0))

    for i in range(spec_start, spec_end + 1):                  # A17:A41
        spec = _vlookup_text(ws_spec[f"{c_spec}{i}"].value)
        ws_spec[f"{c_class}{i}"] = lk["spec_class"].get(spec) or cfg["dom.text_no_class"]

    ws_incoming[cfg["dom.cell_dry"]] = lk["code_dry"].get(qcode) or cfg["dom.text_no_dry"]
    ws_incoming[cfg["dom.cell_cap"]] = lk["code_cap"].get(qcode) or cfg["dom.text_no_cap"]
    ws_incoming[cfg["dom.cell_motor"]] = lk["code_motor"].get(qcode) or cfg["dom.text_no_motor"]

    # ⑥ 수식 재계산을 파일 열 때 Excel 에 위임
    wb.calculation.fullCalcOnLoad = True

    # ⑦ 빈 행 숨기기 (SPEC / SIGN_SPEC 모두 같은 개수)
    filled = len(out_list)
    _hide_rows(ws_spec, spec_start, spec_end, filled)
    _hide_rows(ws_sign, cfg["dom.sign_start"], cfg["dom.sign_end"], filled)

    # ⑧ 불필요한 시트 숨기기
    _keep_only_sheets(wb, [
        cfg["sheet.spec"], cfg["sheet.sign_spec"],
        cfg["sheet.incoming"], cfg["sheet.req_copy"],
    ])

    _reset_sheet_selections(wb)
    wb.save(path)
    if wb_cache is None:
        wb.close()          # 재사용 워크북은 호출자가 닫는다
    restore_header_footer_images(state.template_path, path)
    logger.info("국내 견적서 생성: %s", path)
    return path


def _check_capacity(doc: str, section: str, capacity: int,
                    n_items: int, n_credits: int) -> None:
    """양식 칸 수를 넘으면 멈춘다 — 넘친 줄을 조용히 버리면 품목이나 Credit 이
    빠진 견적서가 그대로 나간다(국내 견적서와 같은 기준)."""
    if n_items + n_credits > capacity:
        detail = f"품목 {n_items}건" + (f" + Credit {n_credits}건" if n_credits else "")
        raise ValueError(
            f"{doc} {section} 칸({capacity}행)을 초과했습니다: {detail}. "
            f"품목을 줄이거나 템플릿 행을 늘려주세요."
        )


def _write_lines(ws, cfg: quote_config.QuoteConfig, prefix: str, start: int, end: int,
                 lines: List[Dict[str, Any]]) -> None:
    """start~end 행의 규격·단가·수량 칸을 비우고 lines 를 위에서부터 적는다.
    수량이 0 인 줄(Credit)은 수량 칸을 비워 둔다."""
    c_spec, c_price, c_qty = (cfg[prefix + ".col_spec"], cfg[prefix + ".col_price"],
                              cfg[prefix + ".col_qty"])
    for rr in range(start, end + 1):
        ws[f"{c_spec}{rr}"] = None
        ws[f"{c_price}{rr}"] = None
        ws[f"{c_qty}{rr}"] = None
    for i, r in enumerate(lines[:end - start + 1]):
        rr = start + i
        ws[f"{c_spec}{rr}"] = r["spec"]
        ws[f"{c_price}{rr}"] = r["price"]
        ws[f"{c_qty}{rr}"] = None if float(r.get("qty", 0)) == 0 else r["qty"]


def _credit_line(cr: Dict[str, Any]) -> Dict[str, Any]:
    return {"spec": cr["spec"], "qty": 0.0, "price": cr["amt"]}


def _fill_china(state: QuoteState, items: List[Dict[str, Any]]) -> str:
    """중국 견적서: 템플릿 복사 후 openpyxl 로 데이터 기입."""
    cfg    = quote_config.load(state.template_path)
    info   = state.cn_info
    names  = {"날짜": datetime.now().strftime("%y%m%d"),
              "라인": safe_filename(info.get("line", "")),
              "호기": safe_filename(info.get("tool", ""))}
    folder = ensure_dir(
        os.path.join(exe_dir(), cfg["name.out_dir"], cfg.text("name.cn_folder", **names))
    )
    path = unique_path(os.path.join(folder, cfg.text("name.cn_file", **names)))

    _copy_template(state.template_path, path)
    wb = load_workbook(path)
    ws = wb[cfg["sheet.quote_cn"]]

    ws[cfg["cn.cell_maker"]]   = info.get("maker", "")
    ws[cfg["cn.cell_process"]] = info.get("process", "")
    ws[cfg["cn.cell_tool"]]    = info.get("tool", "")

    pump    = [r for r in items if r["role"] != "CREDIT" and r["cat"] == "PUMP"]
    others  = [r for r in items if r["role"] != "CREDIT" and r["cat"] != "PUMP"]
    credits = [r for r in items if r["role"] == "CREDIT"]
    pump_credits, rack_credits = _split_credits(credits, cfg)

    pump_start, pump_end = cfg["cn.pump_start"], cfg["cn.pump_end"]
    rack_start, rack_end = cfg["cn.rack_start"], cfg["cn.rack_end"]
    _check_capacity("중국 견적서", "PUMP", pump_end - pump_start + 1,
                    len(pump), len(pump_credits))
    _check_capacity("중국 견적서", "RACK", rack_end - rack_start + 1,
                    len(others), len(rack_credits))

    # 펌프 구간 (Pump Credit 포함) / 랙 구간 (Rack Credit 만 포함)
    _write_lines(ws, cfg, "cn", pump_start, pump_end,
                 list(pump) + [_credit_line(cr) for cr in pump_credits])
    out_list = list(others) + [_credit_line(cr) for cr in rack_credits]
    _write_lines(ws, cfg, "cn", rack_start, rack_end, out_list)

    wb.calculation.fullCalcOnLoad = True
    _hide_rows(ws, rack_start, rack_end, len(out_list))
    _keep_only_sheets(wb, [cfg["sheet.quote_cn"]])
    _reset_sheet_selections(wb)
    wb.save(path)
    wb.close()
    restore_header_footer_images(state.template_path, path)
    logger.info("중국 견적서 생성: %s", path)
    return path


def _fill_us(state: QuoteState, items: List[Dict[str, Any]]) -> str:
    """미국 견적서: 템플릿 복사 후 openpyxl 로 데이터 기입."""
    cfg    = quote_config.load(state.template_path)
    info   = state.us_info
    names  = {"날짜": datetime.now().strftime("%y%m%d"),
              "사이트": safe_filename(info.get("site", "")),
              "호기": safe_filename(info.get("tool", ""))}
    folder = ensure_dir(
        os.path.join(exe_dir(), cfg["name.out_dir"], cfg.text("name.us_folder", **names))
    )
    path = unique_path(os.path.join(folder, cfg.text("name.us_file", **names)))

    _copy_template(state.template_path, path)
    wb = load_workbook(path)
    ws = wb[cfg["sheet.quote_us"]]

    ws[cfg["us.cell_maker"]]    = info.get("maker", "")
    ws[cfg["us.cell_process"]]  = info.get("process", "")
    ws[cfg["us.cell_tool"]]     = info.get("tool", "")
    ws[cfg["us.cell_exchange"]] = float(info.get("exchange", 0))
    y, m = info.get("base_ym", (datetime.now().year, datetime.now().month))
    # 글자가 아니라 날짜로 넣는다 — 양식 셀 서식(mmm-yy)이 글자에는 먹지 않아
    # "Nov-26" 이 아니라 "2026-11-1" 로 찍힌다.
    ws[cfg["us.cell_date"]] = datetime(int(y), int(m), 1)

    pump    = [r for r in items if r["role"] != "CREDIT" and r["cat"] == "PUMP"]
    others  = [r for r in items if r["role"] != "CREDIT" and r["cat"] != "PUMP"]
    credits = [r for r in items if r["role"] == "CREDIT"]
    pump_credits, rack_credits = _split_credits(credits, cfg)

    pump_row = cfg["us.pump_row"]
    rack_start, rack_end = cfg["us.rack_start"], cfg["us.rack_end"]
    _check_capacity("미국 견적서", "PUMP", 1, len(pump), 0)
    _check_capacity("미국 견적서", "RACK", rack_end - rack_start + 1,
                    len(others), len(credits))

    # 펌프는 한 줄. 수량은 0 이어도 그대로 적는다(예전 동작).
    for key, field in (("us.col_spec", "spec"), ("us.col_price", "price"), ("us.col_qty", "qty")):
        ws[cfg.at(key, pump_row)] = pump[0][field] if pump else None

    # 랙 구간: Pump Credit → 맨 앞, Rack Credit → 맨 뒤
    out_list = ([_credit_line(cr) for cr in pump_credits] + list(others)
                + [_credit_line(cr) for cr in rack_credits])
    _write_lines(ws, cfg, "us", rack_start, rack_end, out_list)

    wb.calculation.fullCalcOnLoad = True
    _hide_rows(ws, rack_start, rack_end, len(out_list))
    _keep_only_sheets(wb, [cfg["sheet.quote_us"]])
    _reset_sheet_selections(wb)
    wb.save(path)
    wb.close()
    restore_header_footer_images(state.template_path, path)
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
    try:
        if qtype == "중국":
            out = _fill_china(state, items)
        elif qtype == "미국":
            out = _fill_us(state, items)
        else:
            out = _fill_domestic(state, rd, items)
    except BaseException:
        _discard_inflight()   # 미완성 템플릿 사본 제거
        raise
    _clear_inflight()
    return out


def generate_quote_multi(
    state        : QuoteState,
    items        : List[Dict[str, Any]],
    request_rows : List[Dict[str, Any]],
    progress_cb  : Optional[Callable[[int, int, str], None]] = None,
) -> List[Tuple[Dict[str, Any], str, str]]:
    """
    국내 견적서 여러 건을 순서대로 생성.
    Excel 을 전혀 실행하지 않으므로 COM 버전 대비 크게 빠르다.
    반환: [(rd, saved_path, error), ...]
      성공: error="" / 실패: saved_path="" 이고 error 에 사유
    """
    results: List[Tuple[Dict[str, Any], str, str]] = []
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

    # 템플릿도 한 번만 열어 재사용 — 건마다 전체 파싱/직렬화를 반복하던 것이
    # 일괄 생성 시간의 대부분이었다(실측 1.40초/건 → 0.47초/건).
    wb_cache = None
    if state.template_path:
        try:
            wb_cache = open_template_for_reuse(state.template_path)
        except Exception as e:
            logger.warning("템플릿 사전 로드 실패 (건별 로드로 대체): %s", e)
            wb_cache = None

    try:
        for i, rd in enumerate(request_rows):
            try:
                path = _fill_domestic(state, rd, items, req_ws_cache, wb_cache)
                _clear_inflight()
                results.append((rd, path, ""))
                if progress_cb:
                    progress_cb(i + 1, total, os.path.basename(path))
            except Exception as e:
                _discard_inflight()   # 미완성 출력 파일 제거
                logger.error("멀티 생성 실패: %s", e, exc_info=True)
                results.append((rd, "", str(e)))
                if progress_cb:
                    progress_cb(i + 1, total, f"[실패] {e}")
                # 실패 지점에 따라 워크북이 중간 상태로 남을 수 있다.
                # 다음 건이 그 상태를 물려받지 않도록 새로 연다.
                if wb_cache is not None:
                    try:
                        wb_cache.close()
                    except Exception:
                        pass
                    try:
                        wb_cache = open_template_for_reuse(state.template_path)
                    except Exception as e2:
                        logger.warning("템플릿 재로드 실패 (건별 로드로 대체): %s", e2)
                        wb_cache = None
    finally:
        for _wb in (req_wb_cache, wb_cache):
            if _wb is not None:
                try:
                    _wb.close()
                except Exception:
                    pass

    return results


def _cover_formula_cols(cfg: quote_config.QuoteConfig) -> Tuple[int, ...]:
    """견적의뢰복사본에서 계산으로 채워지는 열 (기본 O/P/Q/S/T/U)."""
    return tuple(cfg.col("copy.col_" + k) for k in (
        "pump_model", "pump_price", "rack_sum", "unit", "amount", "check"))


def _needs_excel_values(path: str, cfg: Optional[quote_config.QuoteConfig] = None) -> bool:
    """이 견적서의 견적의뢰복사본 O/P/Q/S/T/U 값을 Excel 로 계산해야만 알 수 있는가.

    지금 버전이 만든 견적서는 이 칸들에 계산값을 직접 써 두므로(_fill_domestic ⑤-a)
    파일만 읽으면 된다. Excel 이 필요한 것은 칸이 수식인데 저장된 계산값이 없는
    경우뿐이다 — 예전 버전이 만들고 Excel 에서 한 번도 저장하지 않은 파일.

    Excel 을 띄웠다 끄는 데 환경에 따라 수십 초가 걸리므로(실측: 종료에만 50초),
    필요 없는 파일 때문에 띄우지 않으려고 미리 가려낸다. 판단할 수 없으면
    True(Excel 사용)로 돌려 예전 동작을 유지한다.
    """
    cfg = cfg or quote_config.QuoteConfig()
    sheet, r, cols = cfg["sheet.req_copy"], cfg["copy.row"], _cover_formula_cols(cfg)
    try:
        wb = load_workbook(path, read_only=True)
        try:
            if sheet not in wb.sheetnames:
                return False            # 어차피 읽을 수 없는 파일 — 뒤에서 누락으로 보고된다
            row = next(wb[sheet].iter_rows(min_row=r, max_row=r, max_col=max(cols),
                                           values_only=True), ())
        finally:
            wb.close()
        formula_cols = [c for c in cols
                        if len(row) >= c and isinstance(row[c - 1], str) and row[c - 1].startswith("=")]
        if not formula_cols:
            return False
        wb = load_workbook(path, read_only=True, data_only=True)
        try:
            cached = next(wb[sheet].iter_rows(min_row=r, max_row=r, max_col=max(cols),
                                              values_only=True), ())
        finally:
            wb.close()
        return any(len(cached) < c or cached[c - 1] is None for c in formula_cols)
    except Exception as e:
        logger.warning("Excel 필요 여부 판단 실패 (%s): %s — Excel 로 읽는다", path, e)
        return True


def _generate_cover_impl(
    template_path : str,
    folder        : str,
    source_files  : List[str],
    investor_name : str = "",
    progress_cb   : Optional[Callable[[int, int, str], None]] = None,
    warranty_years: int = 2,
) -> Tuple[str, List[Tuple[str, str]]]:
    """
    갑지 생성.
    source_files 각각의 견적의뢰복사본 2행 데이터를 openpyxl 로 읽어
    갑지DATA 시트에 쌓은 뒤 저장.

    반환: (저장 경로, 읽기 실패 목록[(파일명, 사유)])
      실패 목록이 비어 있지 않으면 그만큼 갑지에서 누락된 것이므로
      호출자가 반드시 사용자에게 알려야 한다.
    """
    cfg = quote_config.load(template_path)
    investor_name = s(investor_name) or cfg["default.investor"]
    sh_copy, copy_row = cfg["sheet.req_copy"], cfg["copy.row"]
    data_start, last_col = cfg["cover.data_row_start"], cfg.col("cover.data_last_col")
    folder_name = os.path.basename(folder.rstrip("\\/"))
    out_path = unique_path(os.path.join(folder, cfg.text(
        "name.cover_file", 폴더=folder_name, 갑지표시=cfg["name.cover_mark"])))

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

    # ── COM 으로 수식 계산값 미리 수집 ────────────────────────────────────
    # openpyxl 저장 시 수식 캐시(cached value)가 소멸된다.
    # data_only=True 로 열면 O/P/Q/S/T/U 열(15,16,17,19,20,21)이 None 반환.
    # COM 세션을 열어 Excel 이 fullCalcOnLoad 재계산한 값을 직접 읽는다.
    # {abs_path.lower(): {col: value}}
    _FORMULA_COLS = _cover_formula_cols(cfg)
    _need_excel = [p for p in clean if _needs_excel_values(p, cfg)]
    _com_vals: Dict[str, Dict[int, Any]] = {}

    # 구간별 소요시간 — Excel 을 쓰는 경로라 개발 PC 에서 재현할 수 없어,
    # 느리다는 제보가 오면 로그의 이 한 줄로 어디가 느린지 가린다.
    _t0 = time.perf_counter()
    _com_ok = bool(_need_excel) and excel_com._ensure_com()
    _ms_load = (time.perf_counter() - _t0) * 1000
    _ms_start = _ms_open = _ms_wbopen = _ms_read = _ms_close = _ms_quit = 0.0
    if _com_ok:
        _xl = None
        _wb_c = _ws_c = None
        try:
            _t = time.perf_counter()
            excel_com.pythoncom.CoInitialize()
            _xl = excel_com.win32.DispatchEx("Excel.Application")
            _ms_start = (time.perf_counter() - _t) * 1000
            _t = time.perf_counter()
            _xl.Visible = False
            _xl.DisplayAlerts = False
            for _p in _need_excel:
                _pkey = os.path.abspath(_p).lower()
                _row: Dict[int, Any] = {}
                try:
                    _tt = time.perf_counter()
                    _wb_c = _xl.Workbooks.Open(
                        os.path.abspath(_p), ReadOnly=True, UpdateLinks=0, AddToMru=False
                    )
                    _ms_wbopen += (time.perf_counter() - _tt) * 1000
                    try:
                        _tt = time.perf_counter()
                        _ws_c = _wb_c.Worksheets(sh_copy)
                        for _c in _FORMULA_COLS:
                            try:
                                _row[_c] = _ws_c.Cells(copy_row, _c).Value
                            except Exception:
                                pass
                        _ms_read += (time.perf_counter() - _tt) * 1000
                    finally:
                        _ws_c = None          # 시트 프록시 먼저 해제
                        _tt = time.perf_counter()
                        try:
                            _wb_c.Close(False)
                        except Exception:
                            pass
                        _wb_c = None
                        _ms_close += (time.perf_counter() - _tt) * 1000
                except Exception as _e:
                    logger.warning("COM 수식값 읽기 실패 (%s): %s", _p, _e)
                _com_vals[_pkey] = _row
        except Exception as _e:
            logger.warning("COM 초기화 실패 — 수식 셀은 openpyxl 폴백: %s", _e)
        finally:
            # 해제 순서가 중요하다.
            # 살아있는 COM 프록시가 남은 채 Quit() 하면 EXCEL.EXE 가 죽지 않고,
            # CoUninitialize() 를 프록시 해제보다 먼저 부르면 그 뒤의 해제가
            # 이미 정리된 아파트먼트에서 일어나 보이지 않는 좀비 프로세스가 남는다.
            #   프록시 해제 → gc → Quit() → 참조 제거 → gc → CoUninitialize()
            _tt = time.perf_counter()
            _clear_clipboard()
            _ws_c = None
            _wb_c = None
            import gc
            gc.collect()
            if _xl is not None:
                try:
                    _xl.Quit()
                except Exception:
                    pass
                _xl = None
                gc.collect()
            try:
                excel_com.pythoncom.CoUninitialize()
            except Exception:
                pass
            _ms_quit = (time.perf_counter() - _tt) * 1000
            _ms_open = (time.perf_counter() - _t0) * 1000 - _ms_load - _ms_start
    _t_fill = time.perf_counter()

    # 템플릿 복사 후 openpyxl 로 편집
    _copy_template(template_path, out_path)
    wb       = load_workbook(out_path)
    ws_data  = wb[cfg["sheet.cover_data"]]
    ws_cover = wb[cfg["sheet.cover"]]

    # 갑지DATA 초기화 — iter_rows 공개 API 사용 (내부 _cells 직접 접근 금지)
    max_data_row = min(ws_data.max_row or 1, 2000)
    if max_data_row >= data_start:
        for row_cells in ws_data.iter_rows(min_row=data_start, max_row=max_data_row,
                                            min_col=1, max_col=last_col):
            for cell in row_cells:
                if cell.value is not None:
                    cell.value = None

    # 각 source 파일의 견적의뢰복사본 2행 읽기
    write_row = data_start
    total_src = len(clean)
    failed: List[Tuple[str, str]] = []   # (파일명, 사유)
    for idx, path in enumerate(clean):
        if progress_cb:
            progress_cb(idx + 1, total_src, os.path.basename(path))
        src_wb = None
        try:
            src_wb = load_workbook(path, data_only=True)
            if sh_copy not in src_wb.sheetnames:
                raise KeyError(f"'{sh_copy}' 시트 없음")
            src_ws = src_wb[sh_copy]
            for c in range(1, last_col + 1):  # A:AA
                ws_data.cell(write_row, c).value = src_ws.cell(copy_row, c).value
            # COM 계산값으로 수식 셀(O/P/Q/S/T/U) 덮어쓰기
            for _c, _v in _com_vals.get(os.path.abspath(path).lower(), {}).items():
                if _v is not None:
                    ws_data.cell(write_row, _c).value = _v
            write_row += 1
        except Exception as e:
            # 읽기에 실패한 원본은 갑지에서 누락된다. 조용히 넘기면
            # 갑지가 한 건 모자란 채 정상으로 보이므로 호출자에게 돌려준다.
            logger.warning("갑지 원본 읽기 실패 (%s): %s", path, e)
            failed.append((os.path.basename(path), str(e)))
        finally:
            if src_wb is not None:
                try:
                    src_wb.close()
                except Exception:
                    pass

    if not failed and write_row == data_start:
        raise ValueError("갑지에 기록할 원본이 없습니다.")

    # 투자자명
    ws_cover[cfg["cover.cell_investor"]] = cfg.text("cover.investor_text", 투자자=investor_name)

    # 보증기간 (기본값: 2 Year after delivery)
    ws_cover[cfg["cover.cell_warranty"]] = cfg.text("cover.warranty_text", 보증=warranty_years)

    # 갑지DATA A2:A31 을 직접 확인하여 빈 행을 COVER(20~49) + COVER_DATA(2~31) 양쪽에 숨기기
    for i in range(cfg["cover.rows"]):
        a_val = ws_data.cell(data_start + i, 1).value
        blank = a_val is None or str(a_val).strip() == ""
        ws_cover.row_dimensions[cfg["cover.row_start"] + i].hidden = blank
        ws_data.row_dimensions[data_start + i].hidden   = blank

    wb.calculation.fullCalcOnLoad = True
    _keep_only_sheets(wb, [cfg["sheet.cover_data"], cfg["sheet.cover"]])

    # 열릴 때 COVER 시트만 활성 탭으로 — 여러 시트가 tabSelected=True 상태로
    # 저장되면 Excel이 '그룹' 모드로 파일을 열어버린다.
    for _sn in wb.sheetnames:
        wb[_sn].sheet_view.tabSelected = (_sn == cfg["sheet.cover"])
    wb.active = wb[cfg["sheet.cover"]]

    _reset_sheet_selections(wb)
    wb.save(out_path)
    wb.close()
    restore_header_footer_images(template_path, out_path)
    logger.info("갑지 생성: %s (원본 %d건, 누락 %d건)",
                out_path, write_row - data_start, len(failed))
    logger.info("갑지: Excel 로 값을 읽은 원본 %d건 / 전체 %d건", len(_need_excel), len(clean))
    logger.info("갑지 소요: COM 로드 %.0fms / Excel 기동 %.0fms / 원본 열기·종료 %.0fms "
                "(파일 열기 %.0f, 값 읽기 %.0f, 파일 닫기 %.0f, Excel 종료 %.0f) / "
                "갑지 작성·저장 %.0fms", _ms_load, _ms_start, _ms_open,
                _ms_wbopen, _ms_read, _ms_close, _ms_quit,
                (time.perf_counter() - _t_fill) * 1000)
    return out_path, failed


def _export_cover_data_sheet_impl(src_path: str, template_path: Optional[str] = None) -> str:
    """
    갑지 파일에서 갑지DATA 시트만 추출하여 새 xlsx 로 저장.

    파일명 규칙:
      원본 파일명 끝의 '_갑지' 를 제거한다.
      예) 대외비_260512_LOT베큠_P3_CVD_권혁성_갑지.xlsx
        → 대외비_260512_LOT베큠_P3_CVD_권혁성.xlsx

    '_갑지' 가 없는 파일명이면 '_제출용' 을 접미사로 붙여 구분한다.
    동일 경로에 같은 이름이 이미 있으면 unique_path() 로 번호를 붙인다.

    반환: 저장된 파일 경로
    """
    import re

    stem, ext = os.path.splitext(os.path.basename(src_path))
    cfg = quote_config.load(template_path)
    new_stem = re.sub(re.escape(cfg["name.cover_mark"]) + r"(?:_\d+)?$", "", stem)
    if new_stem == stem:
        raise ValueError(f"갑지 파일이 아닙니다: '{os.path.basename(src_path)}'")
    out_path = unique_path(os.path.join(os.path.dirname(src_path), new_stem + ext))

    # 원본을 통째로 복사 → 서식·스타일 완전 보존
    _copy_template(src_path, out_path)

    wb = load_workbook(out_path)
    keep = cfg["sheet.cover_data"]
    if keep not in wb.sheetnames:
        wb.close()
        os.remove(out_path)
        raise ValueError(f"'{keep}' 시트를 찾을 수 없습니다: {os.path.basename(src_path)}")

    for name in list(wb.sheetnames):
        if name != keep:
            del wb[name]

    ws = wb[keep]
    ws.sheet_state = "visible"
    wb.active = ws
    _reset_sheet_selections(wb)
    wb.save(out_path)
    wb.close()
    logger.info("제출용 엑셀 생성: %s", out_path)
    return out_path


# ═══════════════════════════════════════════════════════════════
# COM 헬퍼 — PDF 변환 전용 (변경 없음)
# ═══════════════════════════════════════════════════════════════

def generate_cover(*args, **kwargs) -> Tuple[str, List[Tuple[str, str]]]:
    """generate_cover 래퍼 — 실패 시 미완성 갑지 파일을 지운다."""
    try:
        out = _generate_cover_impl(*args, **kwargs)
    except BaseException:
        _discard_inflight()
        raise
    _clear_inflight()
    return out


def export_cover_data_sheet(*args, **kwargs) -> str:
    """export_cover_data_sheet 래퍼 — 실패 시 미완성 파일을 지운다."""
    try:
        out = _export_cover_data_sheet_impl(*args, **kwargs)
    except BaseException:
        _discard_inflight()
        raise
    _clear_inflight()
    return out
