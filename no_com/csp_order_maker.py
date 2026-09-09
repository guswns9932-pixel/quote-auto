# -*- coding: utf-8 -*-
"""
CSP 주문접수 업로드 파일 생성기  (알파 v0.2)

사용법
  1) 이 파일과 'CSP_주문접수_업로드_통합양식.xlsx' 를 같은 폴더에 둔다
  2) python csp_order_maker.py
  3) 공통값을 채우고 -> 옵션(자재코드/CIP 조건) 입력 -> 품목 라인을 추가
     -> [엑셀 파일 생성]

v0.2: CIP AS-IS FSC 알람 고도화 — 옵션(사업장/DEVICE/대공정/설비사/
      세부공정)을 CIP 시트 기준으로 입력받아, 자재코드 선택/추가 시
      옵션 조건까지 완전히 일치하면 빨강("검토 필요"), 세부공정만
      다르면 주황으로 표시한다.

필요 패키지 : openpyxl
"""

import os
import re
import sys
import json
import zipfile
import subprocess
import datetime as dt
import xml.etree.ElementTree as ET
import tkinter as tk
from tkinter import ttk, filedialog, messagebox

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, PatternFill, Border, Alignment
from openpyxl.utils import get_column_letter, column_index_from_string

APP_TITLE = "CSP 주문접수 업로드 파일 생성기  (alpha v0.2)"
TEMPLATE_NAME = "CSP_주문접수_업로드_통합양식.xlsx"
SETTINGS_NAME = "csp_order_maker_settings.json"
LOG_NAME = "CSP_주문접수_전체로그.xlsx"
UPLOAD_DIR_NAME = "CSP 주문접수 UPLOAD"

# ---------------------------------------------------------------- 양식 정의
# (엑셀열, 헤더명, 필수여부)
COLUMNS = [
    ("A", "판매오더유형", True),
    ("B", "판매처코드", True),
    ("C", "인도처코드", True),
    ("D", "유통경로", True),
    ("E", "제품군", False),
    ("F", "고객PO번호", True),
    ("G", "고객PO일자", False),
    ("H", "인도조건", True),
    ("I", "인도장소", False),
    ("J", "가격결정일", True),
    ("K", "통화", True),
    ("L", "고객라인", False),
    ("M", "대공정", False),
    ("N", "설비MAKER", False),
    ("O", "고객세부공정", False),
    ("P", "고객설비호기", True),
    ("Q", "자재코드", True),
    ("R", "오더수량", True),
    ("S", "단위", False),
    ("T", "납품요청일", True),
    ("U", "출하지점", False),
    ("V", "조건유형", False),
    ("W", "단가", False),
    ("X", "금액", True),
    ("Y", "통신유형", True),
]

# 모든 행이 같은 값을 갖는 항목 (화면 위쪽에서 한 번만 입력)
COMMON_KEYS = ["A", "B", "D", "E", "G", "H", "J", "K",
               "R", "S", "U", "V", "Y"]
# 행마다 달라지는 항목 (아래 표에서 행별 입력)
LINE_KEYS = ["C", "F", "I", "L", "M", "N", "O", "P", "Q", "T", "W", "X"]

HEADER_BY_KEY = {k: h for k, h, _ in COLUMNS}
REQUIRED_KEYS = {k for k, _, r in COLUMNS if r}
COL_INDEX = {k: i for i, (k, _, _) in enumerate(COLUMNS)}


def app_dir():
    if getattr(sys, "frozen", False):
        return os.path.dirname(sys.executable)
    return os.path.dirname(os.path.abspath(__file__))


def upload_dir():
    """생성된 주문접수 파일과 로그를 모아두는 폴더 (실행파일 위치 기준),
    없으면 새로 만든다."""
    path = os.path.join(app_dir(), UPLOAD_DIR_NAME)
    os.makedirs(path, exist_ok=True)
    return path


def open_folder(path):
    """탐색기(또는 각 OS의 파일관리자)로 폴더를 연다."""
    try:
        if sys.platform.startswith("win"):
            os.startfile(path)
        elif sys.platform == "darwin":
            subprocess.run(["open", path], check=False)
        else:
            subprocess.run(["xdg-open", path], check=False)
    except Exception:
        pass


# ---------------------------------------------------------------- CIP 헤더 탐색
# CIP 시트의 열 위치를 코드에 고정(하드코딩)하지 않고, 헤더 텍스트를 찾아서
# 그 아래 실제 데이터가 있는 열을 알아낸다 — 열 순서가 바뀌어도 그대로 동작한다.
_CIP_HEADER_SCAN_ROWS = 6   # 헤더 관련 텍스트는 이 안에 있다고 보고 그 안에서만 찾는다
_CIP_MAX_COL = 40


def _h(v):
    """헤더 텍스트 비교용 정규화: 공백/하이픈/마침표 제거 후 대문자."""
    if v is None:
        return ""
    return re.sub(r"[\s\-\.]+", "", str(v).strip()).upper()


def _find_cip_col(ws, label):
    """헤더 행들 중 어느 셀이든 label과 일치하면 그 열 번호(1-based)를 반환."""
    target = _h(label)
    if not target:
        return None
    for r in range(1, _CIP_HEADER_SCAN_ROWS + 1):
        for c in range(1, _CIP_MAX_COL + 1):
            if _h(ws.cell(row=r, column=c).value) == target:
                return c
    return None


def _find_cip_subheader_col(ws, section_label, sub_label):
    """병합된 섹션 헤더(예: 'AS-IS') 아래에 있는 서브헤더(예: 'FSC') 열을 찾는다.

    read_only 모드로 열면 병합 셀 정보를 읽을 수 없어(맨 왼쪽 셀에만 값이
    있고 나머지는 None) merged_cells를 쓸 수 없다. 대신 한 행을 왼쪽에서
    가장 가까운 값으로 채워 넣어(엑셀 화면에 보이는 대로) 각 열이 어느
    섹션에 속하는지 판단한다.
    """
    section_target = _h(section_label)
    sub_target = _h(sub_label)
    for r in range(1, _CIP_HEADER_SCAN_ROWS + 1):
        filled, last = [], ""
        for c in range(1, _CIP_MAX_COL + 1):
            v = _h(ws.cell(row=r, column=c).value)
            if v:
                last = v
            filled.append(last)
        if section_target not in filled:
            continue
        for r2 in range(r + 1, min(r + 3, _CIP_HEADER_SCAN_ROWS) + 1):
            for c in range(1, _CIP_MAX_COL + 1):
                if (filled[c - 1] == section_target
                        and _h(ws.cell(row=r2, column=c).value) == sub_target):
                    return c
    return None


# ---------------------------------------------------------------- CIP 매치(AS-IS FSC 알람)
# "ALL"(사업장) / "-"(대공정·설비사·세부공정)은 CIP 시트에서 "해당 항목 전체에
# 적용됨"을 뜻하는 값이라 와일드카드로 취급한다. 옵션 칸이 아직 비어 있는
# 경우는(아직 입력 안 함) 와일드카드로 보지 않는다 — 그래야 세부공정만 빈
# 상태에서도 "세부공정 제외 동일"(orange)로 자연스럽게 떨어진다.
_CIP_WILDCARDS = {"ALL", "-"}


def _norm_plain(v):
    if v is None:
        return ""
    if isinstance(v, float) and v.is_integer():
        v = int(v)
    return str(v).strip().upper()


def _norm_subproc(v):
    """세부공정 정규화: 영문/숫자/한글이 아닌 문자는 전부 '-'로 통일하고
    (연속되면 하나로 합침) 대소문자를 구분하지 않는다."""
    raw = _norm_plain(v)
    if raw in _CIP_WILDCARDS or raw == "":
        return raw
    return re.sub(r"[^0-9A-Z가-힣]+", "-", raw).strip("-")


def _cip_field_eq(cip_val, opt_val, normalize=_norm_plain):
    a, b = normalize(cip_val), normalize(opt_val)
    if not a or not b:
        return False
    return a == b or a in _CIP_WILDCARDS or b in _CIP_WILDCARDS


def cip_match_level(cip_rows, site, device, process, vendor, subproc, fsc):
    """옵션(사업장/DEVICE/대공정/설비사/세부공정) + 자재코드(fsc)를 CIP AS-IS
    데이터와 비교한다.

    반환: "red"  — 옵션 5개 필드 + 자재코드가 CIP 한 행과 완전히 일치
                    (해당 조건에서 이 자재코드는 AS-IS로 등록돼 있음 = 검토 필요)
          "orange" — 세부공정을 제외한 나머지(사업장/DEVICE/대공정/설비사)와
                    자재코드가 일치하는 CIP 행이 있지만 세부공정만 다름
          None  — 해당 사항 없음
    """
    fsc_n = _norm_plain(fsc)
    if not fsc_n or not cip_rows:
        return None
    found_orange = False
    for r in cip_rows:
        if _norm_plain(r.get("fsc")) != fsc_n:
            continue
        if not _cip_field_eq(r.get("site"), site):
            continue
        if not _cip_field_eq(r.get("device"), device):
            continue
        if not _cip_field_eq(r.get("process"), process):
            continue
        if not _cip_field_eq(r.get("vendor"), vendor):
            continue
        if _cip_field_eq(r.get("subproc"), subproc, normalize=_norm_subproc):
            return "red"
        found_orange = True
    return "orange" if found_orange else None


# ---------------------------------------------------------------- 마스터 데이터
class MasterData:
    """통합양식 파일의 코드 시트 / FSC 시트를 읽어들인다."""

    def __init__(self, path):
        self.path = path
        self.order_types = []      # [(코드, 내역)]
        self.sold_to = []          # [(코드, 명, 주소)]
        self.ship_to = []
        self.channels = []
        self.inco_terms = []
        self.currencies = []
        self.comm_types = []
        self.fsc = []              # [(FSC, VER, FSC NM, 설명, 상태)]
        self.fsc_filter_note = ""  # 필터가 완화/생략된 경우의 안내 문구
        # CIP 시트: AS-IS FSC 알람용 옵션 드롭다운 + 매치 데이터
        self.cip_rows = []         # [{"site","device","process","vendor","subproc","fsc"}, ...]
        self.cip_sites = []
        self.cip_devices = []
        self.cip_processes = []
        self.cip_vendors = []
        self.cip_subprocs = []     # 정규화(특수문자→'-', 대소문자 무시)된 고유값
        self._load()

    @staticmethod
    def _s(v):
        if v is None:
            return ""
        if isinstance(v, float) and v.is_integer():
            return str(int(v))          # 1000000.0 -> "1000000" (코드값 왜곡 방지)
        return str(v).strip()

    def _code_sheet(self, wb, name, header_rows=2):
        """A=코드, B=내역 형태의 시트를 읽는다."""
        out = []
        if name not in wb.sheetnames:
            return out
        ws = wb[name]
        for row in ws.iter_rows(min_row=header_rows + 1, max_col=2, values_only=True):
            code = self._s(row[0])
            if not code:
                continue
            out.append((code, self._s(row[1] if len(row) > 1 else "")))
        return out

    def _partner_sheet(self, wb, name):
        out = []
        if name not in wb.sheetnames:
            return out
        ws = wb[name]
        for row in ws.iter_rows(min_row=2, max_col=3, values_only=True):
            code = self._s(row[0])
            if not code:
                continue
            out.append((code, self._s(row[1]), self._s(row[2])))
        return out

    def _load(self):
        wb = load_workbook(self.path, read_only=True, data_only=True)
        try:
            self.order_types = self._code_sheet(wb, "판매오더유형")
            self.channels = self._code_sheet(wb, "유통경로")
            self.inco_terms = self._code_sheet(wb, "인도조건")
            self.currencies = self._code_sheet(wb, "통화")
            self.comm_types = self._code_sheet(wb, "통신유형")
            self.sold_to = self._partner_sheet(wb, "판매처코드")
            self.ship_to = self._partner_sheet(wb, "인도처코드")

            if "FSC" in wb.sheetnames:
                ws = wb["FSC"]
                seen = set()
                candidates = []
                for row in ws.iter_rows(min_row=2, max_col=11, values_only=True):
                    code = self._s(row[1])          # B열 : FSC
                    if not code or code in seen:
                        continue
                    seen.add(code)
                    candidates.append((
                        code,
                        self._s(row[2]),            # C열 : VER
                        self._s(row[5]),            # F열 : FSC NM
                        self._s(row[7]).replace("\n", " "),   # H열 : 설명
                        self._s(row[10]),           # K열 : 상태
                    ))

                def is_bom_active(status):
                    # 공백/대소문자 차이를 흡수해서 'BOM활성화', 'BOM 활성화' 등을
                    # 모두 활성으로 인식한다.
                    norm = re.sub(r"\s+", "", status).upper()
                    return norm == "BOM활성화".upper()

                not_d = [f for f in candidates if not f[0].upper().startswith("D")]
                both = [f for f in not_d if is_bom_active(f[4])]

                # 필터를 다 적용했을 때 결과가 하나도 없으면, 실제 파일의 '상태'
                # 표기가 예상('BOM활성화')과 달라서 전부 걸러졌을 가능성이 높다.
                # 검색창이 완전히 비어버리는 것을 막기 위해 단계적으로 필터를
                # 완화해서라도 목록을 보여준다.
                if both:
                    self.fsc, self.fsc_filter_note = both, ""
                elif not_d:
                    self.fsc = not_d
                    self.fsc_filter_note = (
                        "FSC 상태값이 'BOM활성화'와 일치하는 항목이 없어 "
                        "상태 필터 없이 %d건을 표시합니다." % len(not_d))
                elif candidates:
                    self.fsc = candidates
                    self.fsc_filter_note = (
                        "필터 조건과 일치하는 FSC가 없어 전체 %d건을 표시합니다."
                        % len(candidates))
                else:
                    self.fsc, self.fsc_filter_note = [], ""

            if "CIP" in wb.sheetnames:
                ws = wb["CIP"]
                # 열 위치를 하드코딩하지 않고 헤더 텍스트로 찾는다 — 열 순서가
                # 바뀌어도, 열이 추가/삭제돼도 그대로 동작한다.
                col_no      = _find_cip_col(ws, "No.") or _find_cip_col(ws, "No")
                col_site    = _find_cip_col(ws, "사업장")
                col_device  = _find_cip_col(ws, "DEVICE")
                col_process = _find_cip_col(ws, "대공정")
                col_vendor  = _find_cip_col(ws, "설비사")
                col_subproc = _find_cip_col(ws, "세부공정")
                col_fsc     = _find_cip_subheader_col(ws, "AS-IS", "FSC")

                cip_rows = []
                sites, devices, processes, vendors, subprocs = set(), set(), set(), set(), set()
                if col_no and col_fsc:
                    for row in ws.iter_rows(min_row=1, values_only=False):
                        # CIP 시트는 머리글이 여러 줄이라 고정 행번호로 자르는
                        # 대신 No.열이 숫자인 행만 실제 데이터 행으로 본다.
                        no = row[col_no - 1].value if len(row) >= col_no else None
                        if not isinstance(no, (int, float)):
                            continue

                        def _cell(col):
                            return self._s(row[col - 1].value) if col and len(row) >= col else ""

                        site, device = _cell(col_site), _cell(col_device)
                        process, vendor = _cell(col_process), _cell(col_vendor)
                        subproc, fsc = _cell(col_subproc), _cell(col_fsc)
                        cip_rows.append({"site": site, "device": device,
                                         "process": process, "vendor": vendor,
                                         "subproc": subproc, "fsc": fsc})
                        if site: sites.add(site)
                        if device: devices.add(device)
                        if process: processes.add(process)
                        if vendor: vendors.add(vendor)
                        if subproc: subprocs.add(_norm_subproc(subproc))

                self.cip_rows = cip_rows
                self.cip_sites = sorted(sites)
                self.cip_devices = sorted(devices)
                self.cip_processes = sorted(processes)
                self.cip_vendors = sorted(vendors)
                self.cip_subprocs = sorted(subprocs)
        finally:
            wb.close()

    @property
    def fsc_codes(self):
        return {f[0] for f in self.fsc}


# ---------------------------------------------------------------- 값 변환 / 검증
def parse_date(text):
    """YYYYMMDD / YYYY-MM-DD / YYYY.MM.DD / YYYY/MM/DD -> datetime.date"""
    t = str(text).strip().replace(".", "-").replace("/", "-")
    if not t:
        return None
    for fmt in ("%Y-%m-%d", "%Y%m%d"):
        try:
            return dt.datetime.strptime(t, fmt).date()
        except ValueError:
            continue
    return None


def parse_int(text):
    t = str(text).strip().replace(",", "")
    if not t:
        return None
    try:
        return int(float(t))
    except ValueError:
        return None


def combo_code(text):
    """'ZOR1 - 제품 일반주문' 형태에서 코드만 뽑아낸다."""
    return str(text).split(" - ", 1)[0].strip()


_DIGITS_RE = re.compile(r"\D")


def format_date_mask(raw):
    """입력 중인 문자열을 yyyy-mm-dd 형태로 강제 정렬한다."""
    digits = _DIGITS_RE.sub("", raw)[:8]
    if len(digits) <= 4:
        return digits
    if len(digits) <= 6:
        return digits[:4] + "-" + digits[4:]
    return digits[:4] + "-" + digits[4:6] + "-" + digits[6:]


def cursor_after_mask(fixed, digit_count):
    """포맷팅 후 문자열에서, 원래 커서 앞에 있던 숫자 개수(digit_count) 만큼
    지나간 위치를 계산한다. 자동으로 붙는 '-' 뒤로 커서를 옮겨줘서
    숫자를 입력할 때마다 커서가 뒤로 튀는 현상을 막는다."""
    seen = 0
    i = 0
    n = len(fixed)
    while i < n and seen < digit_count:
        if fixed[i].isdigit():
            seen += 1
        i += 1
    if i < n and fixed[i] == "-":
        i += 1
    return i


def due_date_color(d, today=None):
    """납품요청일까지 남은 기간에 따른 경고색을 정한다 (작성일 기준).
    6주 이내: 빨강, 6~8주: 주황, 8주 이상: 파랑."""
    if d is None:
        return None
    today = today or dt.date.today()
    days = (d - today).days
    if days <= 6 * 7:
        return "red"
    if days < 8 * 7:
        return "orange"
    return "blue"


def ship_to_suffix(code):
    """인도처코드의 '-' 뒤 단어를 뽑아낸다. 예: '삼성전자-16L' -> '16L'"""
    code = str(code).strip()
    if "-" not in code:
        return ""
    return code.rsplit("-", 1)[-1].strip()


# ---------------------------------------------------------------- 의뢰파일
# 의뢰파일(견적/발주 의뢰 엑셀)에서 가져올 열 (사용자가 지정한 열 문자 기준)
REQUEST_COLS = {
    "po": "D",          # Purchase Requisition -> 고객PO번호
    "material": "F",    # Material (참고용, 자동입력 없음)
    "desc": "G",        # Material Description -> 'LOT,' 뒤 값으로 자재코드 검색
    "qty": "H",         # 수량 -> 생성수량
    "line": "K",        # 라인 -> '_' 뒤 값으로 대공정
    "subprocess": "N",  # 세부공정 -> 고객세부공정
    "maker": "X",       # 설비Maker -> 설비MAKER
    "equip_no": "Z",    # 설비호기 -> 고객설비호기
    "due": "AA",        # 희망 납품일 -> 납품요청일
}
REQUEST_COL_IDX = {k: column_index_from_string(v) - 1 for k, v in REQUEST_COLS.items()}


def extract_after_underscore(text):
    """'P1F_CVD' -> 'CVD' ('_' 뒤 값을 뽑아낸다. 없으면 원문 그대로)"""
    text = str(text or "").strip()
    if "_" not in text:
        return text
    return text.rsplit("_", 1)[-1].strip()


def load_request_rows(path):
    """의뢰파일에서 D/F/G/H/N/X/Z/AA 열 값을 읽어 dict 리스트로 반환한다."""
    wb = load_workbook(path, read_only=True, data_only=True)
    try:
        ws = wb["Sheet1"] if "Sheet1" in wb.sheetnames else wb.worksheets[0]
        max_idx = max(REQUEST_COL_IDX.values())
        rows = []
        for row in ws.iter_rows(min_row=2, values_only=True):
            if row is None or len(row) <= max_idx:
                continue
            if all(row[i] is None for i in REQUEST_COL_IDX.values()):
                continue
            rows.append({k: row[i] for k, i in REQUEST_COL_IDX.items()})
        return rows
    finally:
        wb.close()


# ---------------------------------------------------------------- 엑셀 출력
def build_output(template_path, rows, out_path):
    """rows : [{열키: 값}] 을 받아 Sheet1 양식의 새 파일을 만든다.
    CIP 경고(빨간 글씨)는 화면 입력 단계에서만 표시하고, 업로드에 쓰이는
    실제 파일에는 서식을 남기지 않는다 (일부 업로드 매크로가 글자색을
    '이 값은 쓰지 말 것'으로 해석해 정상 자재코드를 거부하는 문제가 있었음)."""
    tpl = load_workbook(template_path)
    tws = tpl["Sheet1"]

    wb = Workbook()
    ws = wb.active
    ws.title = "Sheet1"

    # --- 헤더행 : 원본 서식 그대로 복사
    for idx, (key, header, _) in enumerate(COLUMNS, start=1):
        src = tws.cell(row=1, column=idx)
        dst = ws.cell(row=1, column=idx, value=src.value)
        dst.font = Font(name=src.font.name, sz=src.font.sz, b=src.font.b,
                        color=src.font.color)
        if src.fill and src.fill.fill_type:
            dst.fill = PatternFill(fill_type=src.fill.fill_type,
                                   fgColor=src.fill.fgColor,
                                   bgColor=src.fill.bgColor)
        dst.border = Border(left=src.border.left, right=src.border.right,
                            top=src.border.top, bottom=src.border.bottom)
        dst.alignment = Alignment(horizontal=src.alignment.horizontal,
                                  vertical=src.alignment.vertical,
                                  wrap_text=src.alignment.wrap_text)
        letter = get_column_letter(idx)
        if tws.column_dimensions[letter].width:
            ws.column_dimensions[letter].width = tws.column_dimensions[letter].width

    # --- 데이터행
    for r, data in enumerate(rows, start=2):
        for idx, (key, _, _) in enumerate(COLUMNS, start=1):
            cell = ws.cell(row=r, column=idx, value=data.get(key))
            if key in ("G", "J", "T"):     # 텍스트 형식 (업로드 시스템이 날짜형 셀을
                cell.number_format = "@"   # 그대로 인식하지 못하므로 문자열로 고정)

    wb.save(out_path)
    _use_shared_strings(out_path)
    return out_path


_NS_MAIN = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
_NS_CT = "http://schemas.openxmlformats.org/package/2006/content-types"
_NS_REL = "http://schemas.openxmlformats.org/package/2006/relationships"
_NS_XML = "http://www.w3.org/XML/1998/namespace"


def _use_shared_strings(path):
    """openpyxl은 문자열 셀을 항상 인라인 문자열(t="inlineStr")로 저장하는데,
    일부 업로드 프로그램은 이 형식을 인식하지 못하고 엑셀의 표준 공유 문자열
    표(sharedStrings.xml, t="s") 형식만 읽어들인다. 그래서 우리가 만든 파일을
    열었다가 그냥 저장만 해도(엑셀이 공유 문자열로 다시 써주므로) 자재코드가
    갑자기 인식되는 현상이 있었다. 매번 손으로 다시 저장하지 않아도 되도록
    엑셀과 동일한 형식으로 파일을 즉석에서 다시 써준다."""
    ET.register_namespace("", _NS_MAIN)

    with zipfile.ZipFile(path, "r") as zin:
        data = {name: zin.read(name) for name in zin.namelist()}

    sheet_path = "xl/worksheets/sheet1.xml"
    root = ET.fromstring(data[sheet_path])

    strings, index = [], {}

    def sst_index(text):
        if text not in index:
            index[text] = len(strings)
            strings.append(text)
        return index[text]

    is_tag, t_tag, v_tag = (f"{{{_NS_MAIN}}}{n}" for n in ("is", "t", "v"))
    for c in root.iter(f"{{{_NS_MAIN}}}c"):
        if c.get("t") != "inlineStr":
            continue
        is_el = c.find(is_tag)
        if is_el is None:
            continue
        t_el = is_el.find(t_tag)
        text = t_el.text if t_el is not None and t_el.text is not None else ""
        c.remove(is_el)
        c.set("t", "s")
        ET.SubElement(c, v_tag).text = str(sst_index(text))

    data[sheet_path] = ET.tostring(root, encoding="UTF-8", xml_declaration=True)

    sst_root = ET.Element(f"{{{_NS_MAIN}}}sst", {
        "count": str(len(strings)), "uniqueCount": str(len(strings))})
    for s in strings:
        si = ET.SubElement(sst_root, f"{{{_NS_MAIN}}}si")
        t_el = ET.SubElement(si, f"{{{_NS_MAIN}}}t")
        t_el.text = s
        if s != s.strip():
            t_el.set(f"{{{_NS_XML}}}space", "preserve")
    data["xl/sharedStrings.xml"] = ET.tostring(
        sst_root, encoding="UTF-8", xml_declaration=True)

    ct_path = "[Content_Types].xml"
    ct_root = ET.fromstring(data[ct_path])
    if not any(el.get("PartName") == "/xl/sharedStrings.xml" for el in ct_root):
        ET.SubElement(ct_root, f"{{{_NS_CT}}}Override", {
            "PartName": "/xl/sharedStrings.xml",
            "ContentType": "application/vnd.openxmlformats-officedocument."
                           "spreadsheetml.sharedStrings+xml"})
    data[ct_path] = ET.tostring(ct_root, encoding="UTF-8", xml_declaration=True)

    rels_path = "xl/_rels/workbook.xml.rels"
    rels_root = ET.fromstring(data[rels_path])
    if not any(el.get("Target") == "sharedStrings.xml" for el in rels_root):
        existing = {el.get("Id") for el in rels_root}
        n = 1
        while "rId%d" % n in existing:
            n += 1
        ET.SubElement(rels_root, f"{{{_NS_REL}}}Relationship", {
            "Id": "rId%d" % n,
            "Type": "http://schemas.openxmlformats.org/officeDocument/2006/"
                    "relationships/sharedStrings",
            "Target": "sharedStrings.xml"})
    data[rels_path] = ET.tostring(rels_root, encoding="UTF-8", xml_declaration=True)

    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as zout:
        for name, content in data.items():
            zout.writestr(name, content)


# ---------------------------------------------------------------- 전체 로그
def log_path():
    return os.path.join(upload_dir(), LOG_NAME)


def append_log(path, rows, source_name):
    """생성될 때마다 rows 를 통합 로그 파일 뒤에 쌓는다."""
    if os.path.exists(path):
        wb = load_workbook(path)
        ws = wb.active
    else:
        wb = Workbook()
        ws = wb.active
        ws.title = "로그"
        ws.append(["생성일시", "생성파일"] + [h for _, h, _ in COLUMNS])

    ts = dt.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    start_row = ws.max_row + 1
    for row in rows:
        ws.append([ts, source_name] + [row.get(k) for k, _, _ in COLUMNS])
    for r in range(start_row, ws.max_row + 1):
        for key in ("G", "J", "T"):
            ws.cell(row=r, column=3 + COL_INDEX[key]).number_format = "@"

    wb.save(path)


def safe_append_log(rows, source_name):
    """로그 파일에 기록하되, 다른 프로그램(엑셀 등)이 파일을 열어두는 등의
    이유로 쓰기가 충돌하면 실패하는 대신 번호를 붙인 새 로그 파일을 만들어
    기록을 남긴다. 잠금이 풀리면 다음 번에는 다시 원래 로그 파일에 쌓인다."""
    base, ext = os.path.splitext(LOG_NAME)
    path = log_path()
    last_err = None
    for n in range(1, 51):
        try:
            append_log(path, rows, source_name)
            return path
        except Exception as e:
            last_err = e
            path = os.path.join(upload_dir(), "%s_%d%s" % (base, n + 1, ext))
    raise last_err


def _log_file_candidates():
    """upload_dir 안의 로그 파일들을 모두 찾는다 (쓰기 충돌로 번호가 붙어
    따로 생성된 파일들 포함)."""
    base, ext = os.path.splitext(LOG_NAME)
    folder = upload_dir()
    pattern = re.compile(r"^%s(_\d+)?%s$" % (re.escape(base), re.escape(ext)))
    return [os.path.join(folder, name) for name in os.listdir(folder)
            if pattern.match(name)]


def load_price_map(path=None):
    """로그 파일(쓰기 충돌로 나뉜 것 포함)을 모두 읽어 자재코드 -> 가장
    최근 단가 매핑을 만든다. path 를 지정하면 그 파일 하나만 읽는다."""
    paths = [path] if path is not None else _log_file_candidates()
    entries = []   # (생성일시, 자재코드, 단가) - 시간순 정렬 후 뒤에서 덮어써서 최신값을 남김
    for p in paths:
        if not os.path.exists(p):
            continue
        try:
            wb = load_workbook(p, read_only=True, data_only=True)
        except Exception:
            continue
        try:
            ws = wb.active
            q_idx = 2 + COL_INDEX["Q"]
            w_idx = 2 + COL_INDEX["W"]
            for row in ws.iter_rows(min_row=2, values_only=True):
                if len(row) <= max(q_idx, w_idx):
                    continue
                code, price = row[q_idx], row[w_idx]
                if code is None or price is None:
                    continue
                entries.append((str(row[0] or ""), str(code).strip(), price))
        finally:
            wb.close()
    entries.sort(key=lambda t: t[0])
    prices = {}
    for _, code, price in entries:
        prices[code] = price
    return prices


# ---------------------------------------------------------------- 버튼 색상
def _darken(hex_color, amount):
    """hex_color(#RRGGBB)를 amount만큼 어둡게 해 hover 색을 만든다."""
    hex_color = hex_color.lstrip("#")
    r = max(0, int(hex_color[0:2], 16) - amount)
    g = max(0, int(hex_color[2:4], 16) - amount)
    b = max(0, int(hex_color[4:6], 16) - amount)
    return "#%02X%02X%02X" % (r, g, b)


def _colored_button(parent, text, command=None, bg="#E0E0E0", fg="#000000", **kw):
    """색이 잘 안 보이는 기본 ttk.Button 대신 tk.Button으로 배경색을 확실히
    넣는다 (ttk.Button은 Windows 기본 테마(vista 등)에서 배경색 지정이
    무시되는 경우가 많아, 테마를 바꾸지 않고도 항상 보이는 classic
    tk.Button을 색상 버튼 전용으로 쓴다)."""
    hover = _darken(bg, 24)
    return tk.Button(
        parent, text=text, command=command,
        bg=bg, fg=fg, activebackground=hover, activeforeground=fg,
        relief="raised", bd=1, font=("맑은 고딕", 9),
        padx=8, pady=3, cursor="hand2",
        **kw,
    )


# ---------------------------------------------------------------- 검색 팝업
class PickerDialog(tk.Toplevel):
    """검색 + 목록 선택 공용 팝업."""

    def __init__(self, parent, title, columns, widths, rows, key_index=0, initial="",
                 highlight_keys=None, warn_levels=None):
        """
        highlight_keys : 강조 표시할 키 값들의 집합(초록 배경 + 목록 상위 정렬).
          (예: 전체 로그에 이미 등장한 적 있는 자재코드)
        warn_levels    : {키 값: "red"/"orange"}. CIP AS-IS FSC 알람.
          None이 아니면(빈 dict라도) key_index 열 왼쪽에 "검토" 전용 열을
          따로 만든다 — 코드 칸 안에 마커 글자를 붙이면 칸이 좁아 코드
          자체가 가려지는 문제가 있어 열을 분리했다. highlight_keys보다
          우선해 목록 맨 위로 올린다. 실제 반환값(self.result)에는
          섞이지 않도록 원본 행을 iid로 따로 기억해둔다.
        """
        super().__init__(parent)
        self.title(title)
        self.transient(parent)
        self.grab_set()
        self.result = None
        self._key_index = key_index
        self._highlight_keys = {str(k) for k in highlight_keys} if highlight_keys else set()
        self._show_review_col = warn_levels is not None
        self._warn_levels = {str(k): v for k, v in (warn_levels or {}).items()}
        self._iid_to_row = {}

        rows = list(rows)

        def _priority(r):
            key = str(r[key_index])
            level = self._warn_levels.get(key)
            if level == "red":
                return 0
            if level == "orange":
                return 1
            if key in self._highlight_keys:
                return 2
            return 3

        if self._highlight_keys or self._warn_levels:
            # 안정 정렬이므로 각 우선순위 그룹 내 원래 순서는 유지된다.
            rows.sort(key=_priority)
        self._rows = rows

        top = ttk.Frame(self, padding=8)
        top.pack(fill="x")
        ttk.Label(top, text="검색").pack(side="left")
        self.var = tk.StringVar(value=initial)
        ent = ttk.Entry(top, textvariable=self.var, width=40)
        ent.pack(side="left", padx=6)
        ent.focus_set()
        self.var.trace_add("write", lambda *_: self._refresh())
        self.count = ttk.Label(top, text="")
        self.count.pack(side="left", padx=6)
        if self._highlight_keys:
            ttk.Label(top, text="(초록색 = 이전 주문 이력 있음)",
                      foreground="#2e7d32").pack(side="left", padx=(10, 0))
        if self._show_review_col:
            ttk.Label(top, text="(검토 열: 🔴 완전 일치 / 🟠 세부공정만 다름)",
                      foreground="#c00").pack(side="left", padx=(10, 0))

        body = ttk.Frame(self, padding=(8, 0, 8, 8))
        body.pack(fill="both", expand=True)

        if self._show_review_col:
            # 알람 전용 열을 key_index(FSC) 열 바로 왼쪽에 끼워 넣는다.
            self._review_col_pos = key_index
            display_columns = list(columns[:key_index]) + ["검토"] + list(columns[key_index:])
            display_widths = list(widths[:key_index]) + [150] + list(widths[key_index:])
        else:
            self._review_col_pos = None
            display_columns = list(columns)
            display_widths = list(widths)

        self.tree = ttk.Treeview(body, columns=display_columns, show="headings",
                                 height=18, selectmode="browse")
        for c, w in zip(display_columns, display_widths):
            self.tree.heading(c, text=c)
            self.tree.column(c, width=w, anchor="w")
        self.tree.tag_configure("used", background="#C8E6C9")
        vs = ttk.Scrollbar(body, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=vs.set)
        self.tree.pack(side="left", fill="both", expand=True)
        vs.pack(side="left", fill="y")
        self.tree.bind("<Double-1>", lambda e: self._ok())
        self.tree.bind("<Return>", lambda e: self._ok())
        ent.bind("<Return>", lambda e: self._focus_first())
        ent.bind("<Down>", lambda e: self._focus_first())

        btn = ttk.Frame(self, padding=(8, 0, 8, 8))
        btn.pack(fill="x")
        _colored_button(btn, "선택", command=self._ok, bg="#C8E6C9").pack(side="right")
        _colored_button(btn, "취소", command=self.destroy, bg="#ECEFF1").pack(
            side="right", padx=6)

        self._refresh()
        self.geometry("+%d+%d" % (parent.winfo_rootx() + 60, parent.winfo_rooty() + 60))

    def _refresh(self):
        kw = self.var.get().strip().lower()
        self.tree.delete(*self.tree.get_children())
        self._iid_to_row = {}
        shown = 0
        for i, row in enumerate(self._rows):
            if kw and not any(kw in str(v).lower() for v in row):
                continue
            key = str(row[self._key_index])
            level = self._warn_levels.get(key)
            used = key in self._highlight_keys

            if self._show_review_col:
                if level == "red":
                    review_text = "🔴 검토 필요"
                elif level == "orange":
                    review_text = "🟠 세부공정만 다름"
                else:
                    review_text = ""
                pos = self._review_col_pos
                display = list(row[:pos]) + [review_text] + list(row[pos:])
            else:
                display = list(row)

            iid = str(i)
            self._iid_to_row[iid] = row
            # "used"(초록 배경, 이전 주문 이력)와 검토 열의 마커는 서로 다른
            # 채널(배경색 vs 별도 열의 글자)이라 둘 다 해당돼도 항상 같이
            # 보인다 — 이전에는 마커를 FSC 칸 글자에 붙여서 칸이 좁으면
            # 하나가 가려 보이는 문제가 있었다.
            self.tree.insert("", "end", iid=iid, values=display,
                             tags=("used",) if used else ())
            shown += 1
            if shown >= 500:
                break
        self.count.config(text="%d건 표시 (전체 %d건)" % (shown, len(self._rows)))

    def _focus_first(self):
        kids = self.tree.get_children()
        if kids:
            self.tree.selection_set(kids[0])
            self.tree.focus(kids[0])
            self.tree.focus_set()

    def _ok(self):
        sel = self.tree.selection()
        if not sel:
            return
        row = self._iid_to_row.get(sel[0])
        if row is not None:
            self.result = row[self._key_index]
        else:
            self.result = self.tree.item(sel[0], "values")[self._key_index]
        self.destroy()


# ---------------------------------------------------------------- 메인 앱
class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title(APP_TITLE)
        self.minsize(1100, 700)

        try:
            self.option_add("*Font", ("맑은 고딕", 9))
        except tk.TclError:
            pass

        self.master_path = tk.StringVar(value=self._find_template())
        self.md = None
        self.common_vars = {}
        self.line_vars = {}
        # CIP AS-IS FSC 알람용 옵션(사업장/DEVICE/대공정/설비사/세부공정).
        # 품목 라인과 달리 여러 행을 추가하는 동안 값이 유지된다.
        self.option_vars = {}
        self.cip_cbo = {}         # {"site"/"device"/"process"/"vendor": Combobox}
        self._subproc_all_values = []
        self.lines = []          # [{열키: 원시 문자열, "_opt": 추가 당시 옵션 스냅샷}]
        self.price_map = load_price_map()   # 자재코드 -> 최근 단가 (모든 로그 파일 취합)
        self.request_path = tk.StringVar()
        self.request_rows = []    # 의뢰파일에서 읽은 dict 리스트

        self._build_ui()
        self._load_master(initial=True)
        self._load_settings()
        self._fit_window_to_content()

    # ---------- 초기화
    def _find_template(self):
        p = os.path.join(app_dir(), TEMPLATE_NAME)
        return p if os.path.exists(p) else ""

    def _fit_window_to_content(self):
        """최초 실행 시 창이 고정 크기(1280x820)보다 실제 내용이 더 커서
        하단 버튼 등이 화면 밖으로 밀려나 안 보이던 문제를 고친다
        (창 크기를 조절하면 다시 보이는 게 바로 이 증상이었다).
        고정값 대신 update_idletasks() 로 실제 필요한 크기를 계산해
        화면 크기를 넘지 않는 선에서 창을 그 크기로 맞추고 화면 중앙에 놓는다."""
        self.update_idletasks()
        req_w = self.winfo_reqwidth()
        req_h = self.winfo_reqheight()
        screen_w = self.winfo_screenwidth()
        screen_h = self.winfo_screenheight()
        w = min(max(req_w, 1100), screen_w - 80)
        h = min(max(req_h, 700), screen_h - 80)
        x = max(0, (screen_w - w) // 2)
        y = max(0, (screen_h - h) // 2)
        self.geometry("%dx%d+%d+%d" % (w, h, x, y))

    def _load_master(self, initial=False):
        path = self.master_path.get()
        if not path or not os.path.exists(path):
            if not initial:
                messagebox.showerror("오류", "양식 파일을 찾을 수 없습니다.")
            self.status.config(text="양식 파일을 지정해 주세요.")
            self._set_form_locked(self.md is None)
            return
        try:
            self.md = MasterData(path)
        except Exception as e:
            messagebox.showerror("오류", "양식 파일을 읽지 못했습니다.\n\n%s" % e)
            self._set_form_locked(self.md is None)
            return
        self._fill_combos()
        self._fill_cip_combos()
        text = ("양식 로드 완료 · 판매처 %d · 인도처 %d · FSC %d건"
                % (len(self.md.sold_to), len(self.md.ship_to), len(self.md.fsc)))
        if self.md.fsc_filter_note:
            text += " (%s)" % self.md.fsc_filter_note
        self.status.config(text=text)
        self._set_form_locked(self.md is None)

    def _fill_combos(self):
        def items(pairs):
            return ["%s - %s" % (c, d) if d else c for c, d in pairs]

        self.cbo["A"]["values"] = items(self.md.order_types)
        self.cbo["D"]["values"] = items(self.md.channels)
        self.cbo["H"]["values"] = items(self.md.inco_terms)
        self.cbo["K"]["values"] = items(self.md.currencies)
        self.cbo["Y"]["values"] = items(self.md.comm_types)

        # 드롭다운은 최초에 첫 항목이 선택되어 있도록 한다 (이미 값이 있으면 유지)
        for key in ("A", "D", "H", "K", "Y"):
            values = self.cbo[key]["values"]
            if values and not self.common_vars[key].get().strip():
                self.common_vars[key].set(values[0])

    def _fill_cip_combos(self):
        """CIP 시트에서 뽑아낸 고유값으로 옵션 콤보박스 목록을 채운다.
        (공통값 콤보와 달리 기본값을 자동 선택하지 않는다 — '이 조건에 맞는
        값을 직접 고르거나 입력'하는 용도라 임의의 첫 값을 넣으면 오히려
        혼란을 준다.)"""
        self.cip_cbo["site"]["values"] = self.md.cip_sites
        self.cip_cbo["device"]["values"] = self.md.cip_devices
        self.cip_cbo["process"]["values"] = self.md.cip_processes
        self.cip_cbo["vendor"]["values"] = self.md.cip_vendors
        self._subproc_all_values = list(self.md.cip_subprocs)
        self._subproc_cbo["values"] = self._subproc_all_values

    def _on_locked_click(self, event):
        """양식을 불러오기 전에 입력 영역을 클릭하면 안내 문구를 띄운다."""
        if self.md is not None:
            return
        w = event.widget
        locked_boxes = (getattr(self, "_common_box", None),
                       getattr(self, "_option_box", None),
                       getattr(self, "_line_box", None))
        while w is not None:
            if w in locked_boxes:
                messagebox.showinfo("안내", "먼저 통합양식을 업로드 하세요.")
                return
            w = w.master

    def _set_state_recursive(self, widget, disabled):
        flag = "disabled" if disabled else "!disabled"
        for child in widget.winfo_children():
            if isinstance(child, (ttk.Entry, ttk.Combobox, ttk.Button, ttk.Treeview)):
                try:
                    child.state([flag])
                except tk.TclError:
                    pass
            elif isinstance(child, tk.Button):
                # 색상 버튼(_colored_button)은 ttk가 아닌 classic tk.Button이라
                # .state()가 없다 — .config(state=...)로 동일하게 잠근다.
                try:
                    child.config(state=(tk.DISABLED if disabled else tk.NORMAL))
                except tk.TclError:
                    pass
            self._set_state_recursive(child, disabled)

    def _set_form_locked(self, locked):
        """양식을 불러오기 전에는 공통값/옵션/품목 라인 입력 영역을 모두 비활성화한다."""
        if hasattr(self, "_common_box"):
            self._set_state_recursive(self._common_box, locked)
        if hasattr(self, "_option_box"):
            self._set_state_recursive(self._option_box, locked)
        if hasattr(self, "_line_box"):
            self._set_state_recursive(self._line_box, locked)

    # ---------- 화면 구성
    def _build_ui(self):
        root = ttk.Frame(self, padding=8)
        root.pack(fill="both", expand=True)

        # 양식 파일
        bar = ttk.Frame(root)
        bar.pack(fill="x", pady=(0, 6))
        ttk.Label(bar, text="양식 파일").pack(side="left")
        ttk.Entry(bar, textvariable=self.master_path).pack(
            side="left", fill="x", expand=True, padx=6)
        _colored_button(bar, "찾아보기", command=self._pick_master,
                        bg="#BBDEFB").pack(side="left")
        _colored_button(bar, "다시 읽기", command=lambda: self._load_master(),
                        bg="#BBDEFB").pack(side="left", padx=4)

        # 공통값
        common_head = ttk.Frame(root)
        ttk.Label(common_head, text=" 공통값 (모든 행에 동일하게 들어감) ").pack(side="left")
        _colored_button(common_head, "공통값 고정", command=self._save_settings,
                        bg="#E0F2F1").pack(side="left", padx=(6, 0))
        box = ttk.LabelFrame(root, labelwidget=common_head, padding=8)
        box.pack(fill="x")
        self._common_box = box
        self.cbo = {}
        self._common_grid(box)

        # 의뢰파일
        rbox = ttk.LabelFrame(root, text=" 의뢰파일 (더블클릭하면 품목 라인에 자동입력) ",
                              padding=8)
        rbox.pack(fill="x", pady=(8, 0))
        self._request_box(rbox)

        # 옵션 (자재코드 + CIP AS-IS FSC 알람용 조건)
        obox = ttk.LabelFrame(
            root, text=" 옵션 (자재코드 · CIP AS-IS FSC 검토 필요 알람 조건) ", padding=8)
        obox.pack(fill="x", pady=(8, 0))
        self._option_box = obox
        self._build_options_box(obox)

        # 하단 버튼 바 — 품목 라인(표)보다 먼저, side="bottom"으로 붙여서
        # 창이 좁아져도 이 버튼들이 항상 온전히 보이게 한다. (나중에 붙이는
        # expand=True 위젯이 남은 공간을 다 가져가버려 이 버튼들이 찌그러지던
        # 문제가 있었다 — pack()은 먼저 붙인 위젯의 크기부터 확보한다.)
        bottom = ttk.Frame(root)
        bottom.pack(side="bottom", fill="x", pady=(8, 0))
        self.status = ttk.Label(bottom, text="", foreground="#555")
        self.status.pack(side="left")
        _colored_button(bottom, "엑셀 파일 생성", command=self._export,
                        bg="#FFE0B2").pack(side="right")
        _colored_button(bottom, "생성 폴더 열기", command=self._open_upload_dir,
                        bg="#E8EAF6").pack(side="right", padx=(0, 6))
        _colored_button(bottom, "초기화", command=self._reset_all,
                        bg="#FFCDD2").pack(side="right", padx=(0, 6))

        # 품목 라인 입력 — 하단 버튼 바보다 나중에 붙여서, 창이 좁을 때
        # 이 영역(특히 표)이 먼저 줄어들게 한다. 표 자체에 스크롤바가
        # 있으니 버튼처럼 찌그러지는 대신 스크롤로 자연스럽게 대응된다.
        lbox = ttk.LabelFrame(root, text=" 품목 라인 (행마다 달라지는 값) ", padding=8)
        lbox.pack(fill="both", expand=True, pady=(8, 0))
        self._line_box = lbox
        self._line_form(lbox)
        self._line_table(lbox)

        # 양식을 아직 불러오기 전에는 입력칸을 잠그고, 클릭하면 안내 문구를 띄운다.
        self.bind_all("<Button-1>", self._on_locked_click, add="+")

    @staticmethod
    def _common_defaults():
        """공통값 최초 기본값. 초기화 버튼에서도 동일한 값을 써야 하므로
        (재)로딩 시점에 상관없이 항상 오늘 날짜를 기준으로 새로 계산한다."""
        today = dt.date.today().strftime("%Y%m%d")
        return {"E": "10", "G": today, "R": "1", "S": "EA",
                "U": "1100", "V": "PR00"}

    def _common_grid(self, parent):
        """공통값 입력칸을 4열로 배치."""
        specs = [
            ("A", "combo"), ("B", "pick_sold"), ("D", "combo"), ("E", "entry"),
            ("G", "date8"), ("H", "combo"), ("J", "date8"), ("K", "combo"),
            ("R", "fixed"), ("S", "entry"), ("U", "entry"), ("V", "entry"),
            ("Y", "combo"),
        ]
        defaults = self._common_defaults()
        for i, (key, kind) in enumerate(specs):
            r, c = divmod(i, 4)
            cell = ttk.Frame(parent)
            cell.grid(row=r, column=c, sticky="ew", padx=6, pady=3)
            parent.columnconfigure(c, weight=1, minsize=220)

            label = HEADER_BY_KEY[key]
            if key in REQUIRED_KEYS:
                label = "* " + label
            ttk.Label(cell, text="%s (%s)" % (label, key), width=16).pack(side="left")

            if key == "J":
                # 고객PO일자(G)와 가격결정일(J)은 항상 같은 값을 쓰므로
                # 변수를 공유해서 어느 한쪽을 고치면 즉시 서로 같아지게 한다.
                var = self.common_vars["G"]
            else:
                var = tk.StringVar(value=defaults.get(key, ""))
            self.common_vars[key] = var

            if kind == "combo":
                w = ttk.Combobox(cell, textvariable=var, state="readonly", width=22)
                w.pack(side="left", fill="x", expand=True)
                self.cbo[key] = w
            elif kind == "pick_sold":
                # 창이 좁아져도 '찾기' 버튼이 가장 먼저 자리를 확보하도록
                # 오른쪽에 먼저 배치하고, 이름 표시 라벨이 남는 공간을 흡수/축소한다.
                _colored_button(cell, "찾기", width=5, bg="#E8EAF6",
                                command=lambda v=var: self._pick_partner("sold", v)
                                ).pack(side="right")
                ttk.Entry(cell, textvariable=var, width=12).pack(side="left")
                lbl = ttk.Label(cell, text="", foreground="#0a6")
                lbl.pack(side="left", fill="x", expand=True, padx=3)
                self._name_B = lbl
                var.trace_add("write", lambda *_a: self._show_partner_name("B"))
            elif kind == "fixed":
                e = ttk.Entry(cell, textvariable=var, width=22, state="readonly")
                e.pack(side="left", fill="x", expand=True)
            else:
                ttk.Entry(cell, textvariable=var, width=22).pack(
                    side="left", fill="x", expand=True)

    def _request_box(self, parent):
        bar = ttk.Frame(parent)
        bar.pack(fill="x")
        ttk.Entry(bar, textvariable=self.request_path).pack(
            side="left", fill="x", expand=True, padx=(0, 6))
        _colored_button(bar, "불러오기", command=self._pick_request_file,
                        bg="#BBDEFB").pack(side="left")
        self.request_status = ttk.Label(bar, text="", foreground="#555")
        self.request_status.pack(side="left", padx=(10, 0))

        cols = ["po", "material", "desc", "qty", "line", "subprocess", "maker",
                "equip_no", "due"]
        headers = {"po": "고객PO번호(D)", "material": "Material(F)", "desc": "규격(G)",
                  "qty": "수량(H)", "line": "라인(K)", "subprocess": "세부공정(N)",
                  "maker": "설비Maker(X)", "equip_no": "설비호기(Z)",
                  "due": "희망납품일(AA)"}
        widths = {"po": 110, "material": 100, "desc": 220, "qty": 50, "line": 90,
                 "subprocess": 110, "maker": 90, "equip_no": 90, "due": 90}
        wrap = ttk.Frame(parent)
        wrap.pack(fill="x", pady=(6, 0))
        self.request_tree = ttk.Treeview(wrap, columns=cols, show="headings", height=5)
        for c in cols:
            self.request_tree.heading(c, text=headers[c])
            self.request_tree.column(c, width=widths[c], anchor="w")
        vs = ttk.Scrollbar(wrap, orient="vertical", command=self.request_tree.yview)
        self.request_tree.configure(yscrollcommand=vs.set)
        self.request_tree.pack(side="left", fill="both", expand=True)
        vs.pack(side="left", fill="y")
        self.request_tree.bind("<Double-1>", self._on_request_dblclick)

    def _pick_request_file(self):
        p = filedialog.askopenfilename(
            title="의뢰파일 선택",
            filetypes=[("Excel", "*.xlsx *.xlsm"), ("모든 파일", "*.*")])
        if not p:
            return
        self.request_path.set(p)
        self._read_request_file()

    def _read_request_file(self):
        path = self.request_path.get()
        if not path or not os.path.exists(path):
            return
        try:
            self.request_rows = load_request_rows(path)
        except Exception as e:
            messagebox.showerror("오류", "의뢰파일을 읽지 못했습니다.\n\n%s" % e)
            return
        self._refresh_request_tree()
        self.request_status.config(text="%d건 로드 (더블클릭하면 자동입력)"
                                   % len(self.request_rows))

    def _refresh_request_tree(self):
        self.request_tree.delete(*self.request_tree.get_children())
        for i, r in enumerate(self.request_rows):
            due = r.get("due")
            due_text = due.strftime("%Y-%m-%d") if isinstance(due, dt.datetime) else (due or "")
            self.request_tree.insert("", "end", iid=str(i), values=[
                r.get("po") if r.get("po") is not None else "",
                r.get("material") if r.get("material") is not None else "",
                r.get("desc") if r.get("desc") is not None else "",
                r.get("qty") if r.get("qty") is not None else "",
                r.get("line") if r.get("line") is not None else "",
                r.get("subprocess") if r.get("subprocess") is not None else "",
                r.get("maker") if r.get("maker") is not None else "",
                r.get("equip_no") if r.get("equip_no") is not None else "",
                due_text,
            ])

    def _on_request_dblclick(self, event):
        sel = self.request_tree.selection()
        if not sel:
            return
        if self.md is None:
            messagebox.showinfo("안내", "먼저 통합양식을 업로드 하세요.")
            return
        r = self.request_rows[int(sel[0])]

        po = r.get("po")
        if po is not None:
            iv = parse_int(po)
            self.line_vars["F"].set(str(iv) if iv is not None else str(po))

        qty = parse_int(r.get("qty"))
        self.line_qty.set(str(qty) if qty and qty >= 1 else "1")

        if r.get("line") is not None:
            proc = extract_after_underscore(r["line"])
            self.line_vars["M"].set(proc)
            self.option_vars["process"].set(proc)   # 옵션의 대공정도 동일하게
        if r.get("subprocess") is not None:
            self.line_vars["O"].set(str(r["subprocess"]).strip())
        if r.get("maker") is not None:
            maker = str(r["maker"]).strip()
            self.line_vars["N"].set(maker)
            self.option_vars["vendor"].set(maker)   # 옵션의 설비사도 동일하게
        if r.get("equip_no") is not None:
            self.line_vars["P"].set(str(r["equip_no"]).strip())

        due = r.get("due")
        if isinstance(due, dt.datetime):
            self.line_vars["T"].set(due.strftime("%Y-%m-%d"))
        elif due:
            self.line_vars["T"].set(format_date_mask(str(due)))

    # ---------- 옵션 (자재코드 + CIP AS-IS FSC 알람 조건)
    def _build_options_box(self, parent):
        # 옵션(사업장/DEVICE/대공정/설비사/세부공정)을 먼저 고르고, 그
        # 조건으로 자재코드를 찾는 흐름이 더 직관적이라 자재코드 입력을
        # 옵션 아래로 내렸다.
        row1 = ttk.Frame(parent)
        row1.pack(fill="x")

        def _combo_cell(label, key):
            cell = ttk.Frame(row1)
            cell.pack(side="left", padx=(0, 14))
            ttk.Label(cell, text=label).pack(anchor="w")
            var = tk.StringVar()
            self.option_vars[key] = var
            cbo = ttk.Combobox(cell, textvariable=var, width=14)
            cbo.pack()
            self.cip_cbo[key] = cbo
            var.trace_add("write", lambda *_: self._update_cip_status())

        _combo_cell("사업장", "site")
        _combo_cell("DEVICE", "device")
        _combo_cell("대공정", "process")
        _combo_cell("설비사", "vendor")

        # 세부공정: 특수문자를 '-'로 통일해 스펠링만 인식하고, 입력하는
        # 대로 드롭다운 목록을 실시간으로 좁혀 보여준다.
        cell = ttk.Frame(row1)
        cell.pack(side="left")
        ttk.Label(cell, text="세부공정").pack(anchor="w")
        var_sub = tk.StringVar()
        self.option_vars["subproc"] = var_sub
        self._subproc_cbo = ttk.Combobox(cell, textvariable=var_sub, width=14)
        self._subproc_cbo.pack()
        var_sub.trace_add("write", lambda *_: self._on_subproc_input())

        row2 = ttk.Frame(parent)
        row2.pack(fill="x", pady=(8, 0))
        ttk.Label(row2, text="자재코드(Q)", width=16).pack(side="left")
        var_q = tk.StringVar()
        self.line_vars["Q"] = var_q
        self.entry_Q = ttk.Entry(row2, textvariable=var_q, width=18)
        self.entry_Q.pack(side="left")
        _colored_button(row2, "찾기", width=5, bg="#E8EAF6",
                        command=self._pick_fsc).pack(side="left", padx=2)
        self.lbl_cip_status = ttk.Label(row2, text="", foreground="#c00")
        self.lbl_cip_status.pack(side="left", padx=(10, 0))

        # 자재코드(Q) 입력시 로그상 최근 단가 자동입력 (없으면 그대로, 수정 가능)
        var_q.trace_add("write", lambda *_: self._auto_price())
        # 자재코드(Q)가 CIP AS-IS와 (옵션 조건까지 포함해) 일치하면 알람 표시
        var_q.trace_add("write", lambda *_: self._update_cip_status())

    def _on_subproc_input(self):
        self._filter_subproc_combo()
        self._update_cip_status()

    def _filter_subproc_combo(self):
        """세부공정 입력값과 스펠링이 일치하는(특수문자·대소문자 무시) 항목만
        드롭다운 목록에 실시간으로 남긴다."""
        typed = _norm_subproc(self.option_vars["subproc"].get())
        all_values = self._subproc_all_values
        if not typed:
            self._subproc_cbo["values"] = all_values
        else:
            self._subproc_cbo["values"] = [v for v in all_values if typed in v]

    def _current_option_fields(self):
        return {k: self.option_vars[k].get() for k in
                ("site", "device", "process", "vendor", "subproc")}

    def _update_cip_status(self):
        """옵션 5개 필드 + 자재코드를 CIP AS-IS와 비교해 상태 라벨을 갱신한다."""
        if not self.md:
            self.lbl_cip_status.config(text="")
            return
        opt = self._current_option_fields()
        level = cip_match_level(self.md.cip_rows, opt["site"], opt["device"],
                                opt["process"], opt["vendor"], opt["subproc"],
                                self.line_vars["Q"].get())
        if level == "red":
            self.lbl_cip_status.config(
                text="🔴 검토 필요 (현재 조건 AS-IS와 완전히 일치)", foreground="#c00")
        elif level == "orange":
            self.lbl_cip_status.config(
                text="🟠 검토 필요 (세부공정 제외 동일)", foreground="#e65100")
        else:
            self.lbl_cip_status.config(text="")

    def _line_form(self, parent):
        form = ttk.Frame(parent)
        form.pack(fill="x")
        # 고객PO번호는 숫자만 입력되도록 키 입력 단계에서 걸러낸다.
        vcmd_digits = (self.register(lambda p: p == "" or p.isdigit()), "%P")
        # 자재코드(Q)는 옵션 박스로 이동했다 — 이 폼에는 만들지 않는다.
        specs = [("C", 12), ("F", 12), ("I", 8), ("L", 8),
                 ("M", 10), ("N", 12), ("O", 14), ("P", 10),
                 ("T", 12), ("W", 10), ("X", 10)]
        for i, (key, width) in enumerate(specs):
            cell = ttk.Frame(form)
            cell.grid(row=0, column=i, padx=4, sticky="nw")
            label = HEADER_BY_KEY[key]
            if key in REQUIRED_KEYS:
                label = "* " + label
            ttk.Label(cell, text=label).pack(anchor="w")
            var = tk.StringVar()
            self.line_vars[key] = var
            row = ttk.Frame(cell)
            row.pack()
            if key == "F":
                entry = ttk.Entry(row, textvariable=var, width=width,
                                  validate="key", validatecommand=vcmd_digits)
            else:
                entry = ttk.Entry(row, textvariable=var, width=width)
            entry.pack(side="left")
            if key == "T":
                self.entry_T = entry
            if key == "C":
                _colored_button(row, "찾기", width=5, bg="#E8EAF6",
                                command=self._pick_line_ship).pack(side="left", padx=2)
                self._name_C = ttk.Label(cell, text="", foreground="#0a6")
                self._name_C.pack(anchor="w")

        # 금액(X) 옆 : 이 값으로 몇 행을 한번에 만들지 지정 (기본 1)
        qty_cell = ttk.Frame(form)
        qty_cell.grid(row=0, column=len(specs), padx=4, sticky="nw")
        ttk.Label(qty_cell, text="생성수량").pack(anchor="w")
        self.line_qty = tk.StringVar(value="1")
        qty_row = ttk.Frame(qty_cell)
        qty_row.pack()
        ttk.Entry(qty_row, textvariable=self.line_qty, width=6,
                 validate="key", validatecommand=vcmd_digits).pack(side="left")

        # 인도처코드(C) 선택시 이름 표시 + 인도장소/고객라인 자동입력
        self.line_vars["C"].trace_add("write", lambda *_: self._on_line_ship_change())
        # 자재코드(Q)/CIP 매치 관련 트레이스는 옵션 박스(_build_options_box)에서 건다.
        # 단가 -> 금액 자동
        self.line_vars["W"].trace_add("write", lambda *_: self._auto_amount())
        # 납품요청일 입력 형식을 yyyy-mm-dd 로 고정
        self._t_guard = False
        self.line_vars["T"].trace_add("write", lambda *_: self._on_date_input())
        # 납품요청일까지 남은 기간에 따라 입력칸 글자색을 바꾼다 (6주내 빨강/
        # 7주내 주황/8주이상 파랑)
        self.line_vars["T"].trace_add("write", lambda *_: self._update_due_color())

        btns = ttk.Frame(parent)
        btns.pack(fill="x", pady=(6, 6))
        ttk.Label(btns,
                  text="납품요청일은 입력 즉시 yyyy-mm-dd 형식으로 정렬됩니다 · "
                       "여러 행을 체크한 뒤 [선택 행에 반영]을 누르면 아래 입력칸의 "
                       "값이 체크된 모든 행에 그대로 적용됩니다",
                  foreground="#777").pack(side="left")
        self.btn_add = _colored_button(btns, "행 추가", command=self._add_line,
                                        bg="#C8E6C9")
        self.btn_add.pack(side="right")
        _colored_button(btns, "입력칸 비우기", command=self._clear_line_form,
                        bg="#ECEFF1").pack(side="right", padx=6)

    def _line_table(self, parent):
        # 행 조작 버튼 바를 표보다 먼저 side="bottom"으로 붙인다 — 창이
        # 좁아졌을 때 표 대신 이 버튼들이 찌그러지는 일이 없게 하기 위함
        # (표는 자체 스크롤바가 있어 공간이 부족하면 스크롤로 대응된다).
        tb = ttk.Frame(parent)
        tb.pack(side="bottom", fill="x", pady=(6, 0))
        _colored_button(tb, "전체 선택", command=self._select_all_lines,
                        bg="#BBDEFB").pack(side="left")
        _colored_button(tb, "선택 행에 반영", command=self._apply_to_selected,
                        bg="#FFF9C4").pack(side="left", padx=6)
        _colored_button(tb, "선택 행 복제", command=self._dup_line,
                        bg="#E8EAF6").pack(side="left")
        _colored_button(tb, "선택 행 삭제", command=self._del_line,
                        bg="#FFCDD2").pack(side="left", padx=6)
        _colored_button(tb, "전체 삭제", command=self._clear_lines,
                        bg="#EF9A9A").pack(side="left")
        self.line_count = ttk.Label(tb, text="0 행")
        self.line_count.pack(side="right")

        wrap = ttk.Frame(parent)
        wrap.pack(fill="both", expand=True)
        cols = ["선택", "No"] + LINE_KEYS
        # 옵션 박스가 추가되며 창 전체 높이가 늘어나, 노트북 화면(1366x768
        # 등)에서 창이 눌려 아래 버튼이 잘 안 보이던 문제가 있었다. 표
        # 자체엔 스크롤바가 있으니 기본 표시 행 수를 줄여 여유를 둔다.
        self.tree = ttk.Treeview(wrap, columns=cols, show="headings",
                                 height=8, selectmode="extended")
        self.tree.heading("선택", text="선택")
        self.tree.column("선택", width=40, anchor="center")
        self.tree.heading("No", text="No")
        self.tree.column("No", width=40, anchor="center")
        widths = {"C": 110, "F": 100, "I": 80, "L": 80, "M": 90, "N": 110,
                  "O": 120, "P": 100, "Q": 210, "T": 100, "W": 100, "X": 100}
        for k in LINE_KEYS:
            self.tree.heading(k, text=HEADER_BY_KEY[k])
            self.tree.column(k, width=widths[k], anchor="w")
        vs = ttk.Scrollbar(wrap, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=vs.set)
        self.tree.pack(side="left", fill="both", expand=True)
        vs.pack(side="left", fill="y")
        # CIP AS-IS 알람(빨강/주황)은 자재코드 앞 마커 텍스트로 표시하므로
        # (색이 아니라 텍스트라 아래 글자색 태그와 절대 충돌하지 않는다)
        # 여기서는 납품임박색만 태그로 관리한다.
        self.tree.tag_configure("due_red", foreground="red")
        self.tree.tag_configure("due_orange", foreground="#E65100")
        self.tree.tag_configure("due_blue", foreground="blue")
        self.tree.bind("<Double-1>", lambda e: self._load_selected())
        self.tree.bind("<Button-1>", self._on_tree_click)
        self.tree.bind("<<TreeviewSelect>>", self._refresh_checks)

    # ---------- 동작
    def _pick_master(self):
        p = filedialog.askopenfilename(
            title="통합양식 파일 선택",
            filetypes=[("Excel", "*.xlsx *.xlsm"), ("모든 파일", "*.*")])
        if p:
            self.master_path.set(p)
            self._load_master()

    def _pick_partner(self, which, var):
        if not self.md:
            return
        if which == "sold":
            rows, title = self.md.sold_to, "판매처 선택"
        else:
            rows, title = self.md.ship_to, "인도처 선택"
        dlg = PickerDialog(self, title, ("코드", "이름", "주소"),
                           (90, 220, 220), rows)
        self.wait_window(dlg)
        if dlg.result:
            var.set(dlg.result)

    def _pick_line_ship(self):
        self._pick_partner("ship", self.line_vars["C"])

    def _on_line_ship_change(self):
        code = self.line_vars["C"].get().strip()
        name = ""
        if self.md:
            name = next((r[1] for r in self.md.ship_to if r[0] == code), "")
        if hasattr(self, "_name_C"):
            self._name_C.config(text=name if name else ("코드 없음" if code else ""),
                                foreground="#0a6" if name else "#c00")
        # 인도장소(I)/고객라인(L)은 인도처코드의 '-' 뒤 단어를 최초값으로 사용한다.
        # (동일한 값으로 채워지되, 이후 각각 자유롭게 수정 가능)
        # 인도처코드 자체는 순수 숫자(엑셀 숫자로 저장되어야 함)라 '-'가 없는
        # 경우가 많으므로, 코드에 없으면 이름(예: '삼성전자-16L')에서도 찾는다.
        suffix = ship_to_suffix(code) or ship_to_suffix(name)
        if suffix:
            if not self.line_vars["I"].get().strip():
                self.line_vars["I"].set(suffix)
            if not self.line_vars["L"].get().strip():
                self.line_vars["L"].set(suffix)

    def _show_partner_name(self, key):
        if not self.md:
            return
        code = self.common_vars[key].get().strip()
        rows = self.md.sold_to if key == "B" else self.md.ship_to
        name = next((r[1] for r in rows if r[0] == code), "")
        lbl = getattr(self, "_name_%s" % key, None)
        if lbl is not None:
            lbl.config(text=name if name else ("코드 없음" if code else ""),
                       foreground="#0a6" if name else "#c00")

    def _pick_fsc(self, initial_search=""):
        if not self.md:
            return
        # 전체 로그(price_map)에 등장한 적 있는 자재코드는 강조 표시하고
        # 목록 맨 위로 올려서, 이전에 실제로 주문했던 FSC를 빠르게 찾을 수 있게 한다.
        # 현재 옵션(사업장/DEVICE/대공정/설비사/세부공정) 조건에서 CIP AS-IS와
        # 일치하는 후보는 그보다 더 우선해 빨강/주황으로 표시한다.
        opt = self._current_option_fields()
        warn_levels = {}
        for f in self.md.fsc:
            code = f[0]
            level = cip_match_level(self.md.cip_rows, opt["site"], opt["device"],
                                    opt["process"], opt["vendor"], opt["subproc"], code)
            if level:
                warn_levels[code] = level
        # "검토" 열이 왼쪽에 따로 추가되는 만큼 나머지 열 너비를 조금씩
        # 줄여 창이 과하게 넓어지지 않게 균형을 맞췄다.
        dlg = PickerDialog(self, "자재코드(FSC) 선택",
                           ("FSC", "VER", "모델명", "설명", "상태"),
                           (130, 45, 100, 260, 80), self.md.fsc,
                           initial=initial_search,
                           highlight_keys=set(self.price_map.keys()),
                           warn_levels=warn_levels)
        self.wait_window(dlg)
        if dlg.result:
            self.line_vars["Q"].set(dlg.result)

    def _auto_amount(self):
        w = parse_int(self.line_vars["W"].get())
        if w is not None:
            self.line_vars["X"].set(str(w))

    def _auto_price(self):
        code = self.line_vars["Q"].get().strip()
        if not code:
            return
        price = self.price_map.get(code)
        if price is not None and not self.line_vars["W"].get().strip():
            self.line_vars["W"].set(str(price))

    def _on_date_input(self):
        if self._t_guard:
            return
        raw = self.line_vars["T"].get()
        fixed = format_date_mask(raw)
        if fixed == raw:
            return
        entry = getattr(self, "entry_T", None)
        try:
            cursor = entry.index("insert") if entry is not None else len(raw)
        except tk.TclError:
            cursor = len(raw)
        digit_count = len(_DIGITS_RE.sub("", raw[:cursor]))
        self._t_guard = True
        self.line_vars["T"].set(fixed)

        def _fix_cursor():
            # Entry 위젯 자체의 삽입 후처리가 이 트레이스보다 나중에 커서를
            # 재배치하므로, 이벤트 루프가 한 번 돈 뒤(after_idle)에 다시
            # 올바른 위치로 옮겨야 덮어써지지 않는다.
            if entry is not None:
                try:
                    entry.icursor(cursor_after_mask(fixed, digit_count))
                except tk.TclError:
                    pass
            self._t_guard = False

        if entry is not None:
            entry.after_idle(_fix_cursor)
        else:
            self._t_guard = False

    def _update_due_color(self):
        d = parse_date(self.line_vars["T"].get())
        color = due_date_color(d) or "black"
        entry = getattr(self, "entry_T", None)
        if entry is not None:
            entry.configure(foreground=color)

    def _clear_line_form(self, reset_qty=True):
        for k in LINE_KEYS:
            self.line_vars[k].set("")
        if hasattr(self, "_name_C"):
            self._name_C.config(text="")
        if reset_qty and hasattr(self, "line_qty"):
            self.line_qty.set("1")

    def _validate_line(self, data):
        errs = []
        for k in LINE_KEYS:
            if k in REQUIRED_KEYS and not data[k].strip():
                errs.append("%s(%s) 은(는) 필수입니다." % (HEADER_BY_KEY[k], k))
        c = data["C"].strip()
        if c and self.md and c not in {r[0] for r in self.md.ship_to}:
            errs.append("인도처코드 '%s' 은(는) 목록에 없습니다." % c)
        q = data["Q"].strip()
        if q and self.md and q not in self.md.fsc_codes:
            errs.append("자재코드 '%s' 은(는) FSC 목록에 없습니다." % q)
        if data["T"].strip() and parse_date(data["T"]) is None:
            errs.append("납품요청일 형식이 올바르지 않습니다. (예: 2026-10-26)")
        for k in ("W", "X"):
            if data[k].strip() and parse_int(data[k]) is None:
                errs.append("%s 은(는) 숫자여야 합니다." % HEADER_BY_KEY[k])
        return errs

    def _add_line(self):
        data = {k: self.line_vars[k].get().strip() for k in LINE_KEYS}
        errs = self._validate_line(data)
        if errs:
            messagebox.showwarning("확인 필요", "\n".join(errs))
            return
        qty = parse_int(self.line_qty.get())
        if qty is None or qty < 1:
            qty = 1
        # 옵션(사업장/DEVICE/대공정/설비사/세부공정)은 여러 행을 추가하는 동안
        # 계속 바뀔 수 있으므로, 나중에 CIP 알람을 다시 계산할 때 "지금
        # 옵션이 뭔지"가 아니라 "이 행을 추가할 당시 옵션이 뭐였는지"를 써야
        # 한다. 행마다 그 시점의 옵션 값을 그대로 저장해둔다.
        opt_snapshot = self._current_option_fields()
        for _ in range(qty):
            row = dict(data)
            row["_opt"] = dict(opt_snapshot)
            self.lines.append(row)
        self._refresh_tree()
        # 다음 행 입력 편의를 위해 유지 (자재코드/단가/금액만 새로 입력)
        # 생성수량도 여기서 "1"로 되돌리지 않는다 — 의뢰파일 더블클릭으로
        # 채워진 값(H열 수량)이 방금 몇 행 생성했는지 그대로 남아 있어야
        # 방금 생성된 수량과 화면이 어긋나 보이지 않는다.
        keep = {k: data[k] for k in ("C", "F", "I", "L", "M", "N", "O", "P", "T")}
        self._clear_line_form(reset_qty=False)
        for k, v in keep.items():
            self.line_vars[k].set(v)

    def _selected_indices(self):
        return sorted(self.tree.index(iid) for iid in self.tree.selection())

    def _select_all_lines(self):
        self.tree.selection_set(self.tree.get_children())

    def _selected_index(self):
        idxs = self._selected_indices()
        return idxs[0] if idxs else None

    def _on_tree_click(self, event):
        """'선택' 열을 클릭하면 다중 선택을 켜고 끈다 (체크박스처럼 동작)."""
        region = self.tree.identify("region", event.x, event.y)
        if region != "cell":
            return
        col = self.tree.identify_column(event.x)
        row = self.tree.identify_row(event.y)
        if not row or col != "#1":
            return
        if row in self.tree.selection():
            self.tree.selection_remove(row)
        else:
            self.tree.selection_add(row)
        return "break"

    def _refresh_checks(self, *_):
        sel = set(self.tree.selection())
        for iid in self.tree.get_children():
            vals = list(self.tree.item(iid, "values"))
            vals[0] = "☑" if iid in sel else "☐"
            self.tree.item(iid, values=vals)

    def _load_selected(self):
        """체크(선택)된 행 중 첫 번째 행의 값을 입력칸으로 불러온다."""
        i = self._selected_index()
        if i is None:
            return
        for k in LINE_KEYS:
            self.line_vars[k].set(self.lines[i][k])

    def _apply_to_selected(self):
        """입력칸의 값을 지금 체크되어 있는 모든 행에 그대로 반영한다.
        체크 표시가 곧 적용 대상이므로 둘이 어긋날 일이 없다."""
        idxs = self._selected_indices()
        if not idxs:
            messagebox.showwarning("확인 필요", "반영할 행을 먼저 체크하세요.")
            return
        data = {k: self.line_vars[k].get().strip() for k in LINE_KEYS}
        errs = self._validate_line(data)
        if errs:
            messagebox.showwarning("확인 필요", "\n".join(errs))
            return
        opt_snapshot = self._current_option_fields()
        for i in idxs:
            row = dict(data)
            row["_opt"] = dict(opt_snapshot)
            self.lines[i] = row
        self._refresh_tree()
        kids = self.tree.get_children()
        self.tree.selection_set([kids[i] for i in idxs])

    def _dup_line(self):
        """선택된 행(여러 행 가능)을 각각 바로 아래에 복제한다."""
        idxs = self._selected_indices()
        if not idxs:
            return
        for i in sorted(idxs, reverse=True):
            self.lines.insert(i + 1, dict(self.lines[i]))
        self._refresh_tree()

    def _del_line(self):
        idxs = self._selected_indices()
        if not idxs:
            return
        if not messagebox.askyesno("확인", "선택한 %d개 행을 삭제할까요?" % len(idxs)):
            return
        for i in sorted(idxs, reverse=True):
            del self.lines[i]
        self._clear_line_form()
        self._refresh_tree()

    def _clear_lines(self):
        if self.lines and messagebox.askyesno("확인", "품목 라인을 모두 지울까요?"):
            self.lines = []
            self._clear_line_form()
            self._refresh_tree()

    def _reset_all(self):
        """공통값·의뢰파일·옵션·품목 라인을 전부 초기 상태로 되돌린다.
        통합양식(마스터) 파일 선택은 그대로 둔다 — 다시 읽을 필요가
        없고, 매번 파일을 다시 고르게 하면 오히려 불편하다."""
        if not messagebox.askyesno(
                "초기화 확인",
                "공통값·의뢰파일·옵션·품목 라인이 모두 초기화됩니다. 계속할까요?"):
            return

        defaults = self._common_defaults()
        for k, var in self.common_vars.items():
            if k == "J":     # G와 변수를 공유하므로 G에서 이미 반영됨
                continue
            var.set(defaults.get(k, ""))
        if self.md:          # 콤보박스 첫 항목 재적용 (A/D/H/K/Y)
            self._fill_combos()

        self.request_path.set("")
        self.request_rows = []
        self._refresh_request_tree()
        self.request_status.config(text="")

        # 옵션(사업장/DEVICE/대공정/설비사/세부공정)은 여러 행 추가 동안
        # 일부러 유지시키는 값이라 [행 추가]/[입력칸 비우기]로는 안 지워진다
        # — 완전 초기화는 여기서만 비운다.
        for var in self.option_vars.values():
            var.set("")
        self._filter_subproc_combo()

        self.lines = []
        self._clear_line_form()
        self._refresh_tree()
        self._update_cip_status()

        self.status.config(text="초기화했습니다.")

    def _refresh_tree(self):
        self.tree.delete(*self.tree.get_children())
        cip_rows = self.md.cip_rows if self.md else []
        q_idx = LINE_KEYS.index("Q")
        for n, d in enumerate(self.lines, start=1):
            # CIP AS-IS 알람은 배경색이 아니라 자재코드 앞 마커 텍스트로
            # 표시한다 — 납품임박색(due_*, 글자색)과 같은 채널을 쓰지
            # 않으므로 우선순위 없이 항상 둘 다 눈에 보인다. 이 행을
            # 추가할 당시의 옵션 값(_opt)을 써야 나중에 옵션을 바꿔도
            # 예전 행이 엉뚱하게 다시 칠해지지 않는다.
            opt = d.get("_opt", {})
            level = cip_match_level(
                cip_rows, opt.get("site", ""), opt.get("device", ""),
                opt.get("process", ""), opt.get("vendor", ""),
                opt.get("subproc", ""), d["Q"])
            values = ["☐", n] + [d[k] for k in LINE_KEYS]
            if level == "red":
                values[2 + q_idx] = "🔴검토필요 " + d["Q"]
            elif level == "orange":
                values[2 + q_idx] = "🟠검토필요(세부공정↓) " + d["Q"]

            tags = []
            color = due_date_color(parse_date(d["T"]))
            if color:
                tags.append("due_%s" % color)
            self.tree.insert("", "end", values=values, tags=tuple(tags))
        self.line_count.config(text="%d 행" % len(self.lines))

    # ---------- 출력
    def _collect_common(self):
        out, errs = {}, []
        for k in COMMON_KEYS:
            raw = self.common_vars[k].get().strip()
            if k in ("A", "D", "H", "K", "Y"):
                raw = combo_code(raw)
            if k in REQUIRED_KEYS and not raw:
                errs.append("%s(%s) 은(는) 필수입니다." % (HEADER_BY_KEY[k], k))
            out[k] = raw
        if out.get("B") and self.md and out["B"] not in {r[0] for r in self.md.sold_to}:
            errs.append("판매처코드 '%s' 은(는) 목록에 없습니다." % out["B"])
        for k in ("G", "J"):
            if out[k]:
                d = parse_date(out[k])
                if d is None:
                    errs.append("%s 형식이 올바르지 않습니다." % HEADER_BY_KEY[k])
                else:
                    out[k] = d.strftime("%Y%m%d")
        return out, errs

    def _build_rows(self, common):
        # 판매처코드/인도처코드/유통경로/제품군/출하지점/고객PO번호는 업로드
        # 시스템이 반드시 엑셀 숫자 형식으로 인식해야 하므로 정수로 변환해
        # 저장한다. (숫자로 변환되지 않는 값은 원래 문자열을 그대로 둔다)
        rows = []
        for d in self.lines:
            row = {}
            for k in COMMON_KEYS:
                v = common[k]
                if k in ("B", "D", "E", "U"):
                    iv = parse_int(v)
                    v = iv if iv is not None else (v or None)
                elif k == "R":
                    v = 1                       # 오더수량은 항상 1
                row[k] = v if v != "" else None
            for k in LINE_KEYS:
                v = d[k]
                if k in ("C", "F"):
                    iv = parse_int(v)
                    v = iv if iv is not None else (v or None)
                elif k == "T":
                    dv = parse_date(v)
                    v = dv.strftime("%Y-%m-%d") if dv else None
                elif k in ("W", "X"):
                    iv = parse_int(v)
                    v = iv if iv is not None else (v or None)
                row[k] = v if v != "" else None
            rows.append(row)
        return rows

    def _open_upload_dir(self):
        open_folder(upload_dir())

    def _export(self):
        if not self.md:
            messagebox.showerror("오류", "먼저 양식 파일을 읽어주세요.")
            return
        if not self.lines:
            messagebox.showwarning("확인 필요", "품목 라인이 없습니다.")
            return
        common, errs = self._collect_common()
        for n, d in enumerate(self.lines, start=1):
            for e in self._validate_line(d):
                errs.append("%d행: %s" % (n, e))
        if errs:
            messagebox.showwarning("확인 필요", "\n".join(errs[:15]))
            return

        out_dir = upload_dir()
        stamp = dt.datetime.now().strftime("%Y%m%d_%H%M%S")
        out = os.path.join(out_dir, "CSP_주문접수_%s.xlsx" % stamp)
        n = 1
        while os.path.exists(out):        # 같은 초에 두 번 생성되는 경우 대비
            n += 1
            out = os.path.join(out_dir, "CSP_주문접수_%s_%d.xlsx" % (stamp, n))
        rows = self._build_rows(common)
        try:
            build_output(self.master_path.get(), rows, out)
        except Exception as e:
            messagebox.showerror("오류", "파일 생성에 실패했습니다.\n\n%s" % e)
            return
        try:
            safe_append_log(rows, os.path.basename(out))
        except Exception as e:
            # 주문 파일 자체는 이미 만들어졌으니 실패로 처리하지 않고 알리기만 한다.
            messagebox.showwarning(
                "안내", "주문 파일은 생성되었지만 로그 기록에는 실패했습니다.\n\n%s" % e)
        for row in rows:                   # 다음 입력을 위해 최근 단가를 갱신
            q, w = row.get("Q"), row.get("W")
            if q and w is not None:
                self.price_map[str(q)] = w
        # 상태 라벨에 전체 경로를 넣으면 길이 때문에 하단 버튼이 화면 밖으로
        # 밀려 사라지는 문제가 있어(가로 한 줄 배치), 경로 없이 완료 사실만
        # 짧게 표시한다. 실제 저장 위치는 완료 팝업에서 확인할 수 있다.
        self.status.config(text="생성 완료 (%d행)" % len(self.lines))
        messagebox.showinfo("완료", "%d행이 생성되었습니다.\n전체 로그에 누적 저장되었습니다.\n\n%s"
                             % (len(self.lines), out))

    # ---------- 설정 저장 / 복원
    def _settings_path(self):
        return os.path.join(app_dir(), SETTINGS_NAME)

    def _save_settings(self, silent=False):
        # 고객PO일자/가격결정일(G, J)은 항상 '작성 당일'이어야 하므로
        # 공통값 고정 여부와 상관없이 저장/복원 대상에서 제외한다.
        data = {"master": self.master_path.get(),
                "common": {k: v.get() for k, v in self.common_vars.items()
                          if k not in ("G", "J")}}
        try:
            with open(self._settings_path(), "w", encoding="utf-8") as f:
                json.dump(data, f, ensure_ascii=False, indent=2)
            if not silent:
                self.status.config(text="공통값을 고정했습니다.")
        except Exception:
            pass

    def _load_settings(self):
        try:
            with open(self._settings_path(), encoding="utf-8") as f:
                data = json.load(f)
        except Exception:
            return
        for k, v in data.get("common", {}).items():
            if k in ("G", "J"):
                continue
            if k in self.common_vars:
                self.common_vars[k].set(v)


if __name__ == "__main__":
    App().mainloop()
