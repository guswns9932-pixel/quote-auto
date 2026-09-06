# -*- coding: utf-8 -*-
"""
CSP 주문접수 업로드 파일 생성기  (알파 v0.1)

사용법
  1) 이 파일과 'CSP_주문접수_업로드_통합양식.xlsx' 를 같은 폴더에 둔다
  2) python csp_order_maker.py
  3) 공통값을 채우고 -> 품목 라인을 추가 -> [엑셀 파일 생성]

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

APP_TITLE = "CSP 주문접수 업로드 파일 생성기  (alpha v0.1)"
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
        self.cip_fsc = set()       # CIP 시트 J열(AS-IS FSC)에 등장하는 값들
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
                cip_fsc = set()
                for row in ws.iter_rows(min_row=2, values_only=True):
                    # CIP 시트는 머리글이 여러 줄이라 고정 행번호로 자르는 대신
                    # No. 열(B, 데이터행에서만 숫자)로 실제 데이터행을 가려낸다.
                    no = row[1] if len(row) > 1 else None
                    if not isinstance(no, (int, float)):
                        continue
                    j_val = self._s(row[9]) if len(row) > 9 else ""   # J열 : AS-IS FSC
                    if j_val:
                        cip_fsc.add(j_val)
                self.cip_fsc = cip_fsc
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


def extract_after_lot(text):
    """'DRY_PUMP;EQ,LOT,HD4500PW' -> 'HD4500PW' (LOT, 뒤 값을 뽑아낸다)"""
    text = str(text or "")
    marker = "LOT,"
    idx = text.find(marker)
    if idx == -1:
        return ""
    return text[idx + len(marker):].strip()


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


# ---------------------------------------------------------------- 검색 팝업
class PickerDialog(tk.Toplevel):
    """검색 + 목록 선택 공용 팝업."""

    def __init__(self, parent, title, columns, widths, rows, key_index=0, initial=""):
        super().__init__(parent)
        self.title(title)
        self.transient(parent)
        self.grab_set()
        self.result = None
        self._rows = rows
        self._key_index = key_index

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

        body = ttk.Frame(self, padding=(8, 0, 8, 8))
        body.pack(fill="both", expand=True)
        self.tree = ttk.Treeview(body, columns=columns, show="headings",
                                 height=18, selectmode="browse")
        for c, w in zip(columns, widths):
            self.tree.heading(c, text=c)
            self.tree.column(c, width=w, anchor="w")
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
        ttk.Button(btn, text="선택", command=self._ok).pack(side="right")
        ttk.Button(btn, text="취소", command=self.destroy).pack(side="right", padx=6)

        self._refresh()
        self.geometry("+%d+%d" % (parent.winfo_rootx() + 60, parent.winfo_rooty() + 60))

    def _refresh(self):
        kw = self.var.get().strip().lower()
        self.tree.delete(*self.tree.get_children())
        shown = 0
        for row in self._rows:
            if kw and not any(kw in str(v).lower() for v in row):
                continue
            self.tree.insert("", "end", values=row)
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
        self.lines = []          # [{열키: 원시 문자열}]
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

    def _on_locked_click(self, event):
        """양식을 불러오기 전에 입력 영역을 클릭하면 안내 문구를 띄운다."""
        if self.md is not None:
            return
        w = event.widget
        while w is not None:
            if w in (getattr(self, "_common_box", None), getattr(self, "_line_box", None)):
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
            self._set_state_recursive(child, disabled)

    def _set_form_locked(self, locked):
        """양식을 불러오기 전에는 공통값/품목 라인 입력 영역을 모두 비활성화한다."""
        if hasattr(self, "_common_box"):
            self._set_state_recursive(self._common_box, locked)
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
        ttk.Button(bar, text="찾아보기", command=self._pick_master).pack(side="left")
        ttk.Button(bar, text="다시 읽기",
                   command=lambda: self._load_master()).pack(side="left", padx=4)

        # 공통값
        common_head = ttk.Frame(root)
        ttk.Label(common_head, text=" 공통값 (모든 행에 동일하게 들어감) ").pack(side="left")
        ttk.Button(common_head, text="공통값 고정",
                   command=self._save_settings).pack(side="left", padx=(6, 0))
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

        # 품목 라인 입력
        lbox = ttk.LabelFrame(root, text=" 품목 라인 (행마다 달라지는 값) ", padding=8)
        lbox.pack(fill="both", expand=True, pady=(8, 0))
        self._line_box = lbox
        self._line_form(lbox)
        self._line_table(lbox)

        # 하단
        bottom = ttk.Frame(root)
        bottom.pack(fill="x", pady=(8, 0))
        self.status = ttk.Label(bottom, text="", foreground="#555")
        self.status.pack(side="left")
        ttk.Button(bottom, text="엑셀 파일 생성",
                   command=self._export).pack(side="right")
        ttk.Button(bottom, text="생성 폴더 열기",
                   command=self._open_upload_dir).pack(side="right", padx=(0, 6))

        # 양식을 아직 불러오기 전에는 입력칸을 잠그고, 클릭하면 안내 문구를 띄운다.
        self.bind_all("<Button-1>", self._on_locked_click, add="+")

    def _common_grid(self, parent):
        """공통값 입력칸을 4열로 배치."""
        specs = [
            ("A", "combo"), ("B", "pick_sold"), ("D", "combo"), ("E", "entry"),
            ("G", "date8"), ("H", "combo"), ("J", "date8"), ("K", "combo"),
            ("R", "fixed"), ("S", "entry"), ("U", "entry"), ("V", "entry"),
            ("Y", "combo"),
        ]
        today = dt.date.today().strftime("%Y%m%d")
        defaults = {"E": "10", "G": today, "R": "1", "S": "EA",
                    "U": "1100", "V": "PR00"}
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
                ttk.Button(cell, text="찾기", width=5,
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
        ttk.Button(bar, text="불러오기", command=self._pick_request_file).pack(side="left")
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
            self.line_vars["M"].set(extract_after_underscore(r["line"]))
        if r.get("subprocess") is not None:
            self.line_vars["O"].set(str(r["subprocess"]).strip())
        if r.get("maker") is not None:
            self.line_vars["N"].set(str(r["maker"]).strip())
        if r.get("equip_no") is not None:
            self.line_vars["P"].set(str(r["equip_no"]).strip())

        due = r.get("due")
        if isinstance(due, dt.datetime):
            self.line_vars["T"].set(due.strftime("%Y-%m-%d"))
        elif due:
            self.line_vars["T"].set(format_date_mask(str(due)))

        keyword = extract_after_lot(r.get("desc"))
        self._pick_fsc(initial_search=keyword)

    def _line_form(self, parent):
        form = ttk.Frame(parent)
        form.pack(fill="x")
        # 고객PO번호는 숫자만 입력되도록 키 입력 단계에서 걸러낸다.
        vcmd_digits = (self.register(lambda p: p == "" or p.isdigit()), "%P")
        specs = [("C", 12), ("F", 12), ("I", 8), ("L", 8),
                 ("M", 10), ("N", 12), ("O", 14), ("P", 10),
                 ("Q", 14), ("T", 12), ("W", 10), ("X", 10)]
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
            elif key == "Q":
                self.entry_Q = entry
            if key == "C":
                ttk.Button(row, text="찾기", width=5,
                           command=self._pick_line_ship).pack(side="left", padx=2)
                self._name_C = ttk.Label(cell, text="", foreground="#0a6")
                self._name_C.pack(anchor="w")
            elif key == "Q":
                ttk.Button(row, text="찾기", width=5,
                           command=self._pick_fsc).pack(side="left", padx=2)

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
        # 자재코드(Q) 입력시 로그상 최근 단가 자동입력 (없으면 그대로, 수정 가능)
        self.line_vars["Q"].trace_add("write", lambda *_: self._auto_price())
        # 자재코드(Q)가 CIP 시트의 AS-IS FSC와 일치하면 빨간 글씨로 경고
        self.line_vars["Q"].trace_add("write", lambda *_: self._check_cip_warning())
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
        self.btn_add = ttk.Button(btns, text="행 추가", command=self._add_line)
        self.btn_add.pack(side="right")
        ttk.Button(btns, text="입력칸 비우기",
                   command=self._clear_line_form).pack(side="right", padx=6)

    def _line_table(self, parent):
        wrap = ttk.Frame(parent)
        wrap.pack(fill="both", expand=True)
        cols = ["선택", "No"] + LINE_KEYS
        self.tree = ttk.Treeview(wrap, columns=cols, show="headings",
                                 height=12, selectmode="extended")
        self.tree.heading("선택", text="선택")
        self.tree.column("선택", width=40, anchor="center")
        self.tree.heading("No", text="No")
        self.tree.column("No", width=40, anchor="center")
        widths = {"C": 110, "F": 100, "I": 80, "L": 80, "M": 90, "N": 110,
                  "O": 120, "P": 100, "Q": 120, "T": 100, "W": 100, "X": 100}
        for k in LINE_KEYS:
            self.tree.heading(k, text=HEADER_BY_KEY[k])
            self.tree.column(k, width=widths[k], anchor="w")
        vs = ttk.Scrollbar(wrap, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=vs.set)
        self.tree.pack(side="left", fill="both", expand=True)
        vs.pack(side="left", fill="y")
        self.tree.tag_configure("cip_warn", foreground="red")
        self.tree.tag_configure("due_red", foreground="red")
        self.tree.tag_configure("due_orange", foreground="orange")
        self.tree.tag_configure("due_blue", foreground="blue")
        self.tree.bind("<Double-1>", lambda e: self._load_selected())
        self.tree.bind("<Button-1>", self._on_tree_click)
        self.tree.bind("<<TreeviewSelect>>", self._refresh_checks)

        tb = ttk.Frame(parent)
        tb.pack(fill="x", pady=(6, 0))
        ttk.Button(tb, text="전체 선택", command=self._select_all_lines).pack(side="left")
        ttk.Button(tb, text="선택 행에 반영", command=self._apply_to_selected).pack(
            side="left", padx=6)
        ttk.Button(tb, text="선택 행 복제", command=self._dup_line).pack(side="left")
        ttk.Button(tb, text="선택 행 삭제", command=self._del_line).pack(side="left", padx=6)
        ttk.Button(tb, text="전체 삭제", command=self._clear_lines).pack(side="left")
        self.line_count = ttk.Label(tb, text="0 행")
        self.line_count.pack(side="right")

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
        dlg = PickerDialog(self, "자재코드(FSC) 선택",
                           ("FSC", "VER", "모델명", "설명", "상태"),
                           (120, 45, 110, 300, 90), self.md.fsc,
                           initial=initial_search)
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

    def _check_cip_warning(self):
        """자재코드가 CIP 시트의 AS-IS FSC(J열)와 완전히 같으면 빨간 글씨로 표시."""
        code = self.line_vars["Q"].get().strip()
        is_cip = bool(self.md and code and code in self.md.cip_fsc)
        entry = getattr(self, "entry_Q", None)
        if entry is not None:
            entry.configure(foreground="red" if is_cip else "black")

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
        for _ in range(qty):
            self.lines.append(dict(data))
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
        for i in idxs:
            self.lines[i] = dict(data)
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

    def _refresh_tree(self):
        self.tree.delete(*self.tree.get_children())
        cip = self.md.cip_fsc if self.md else set()
        for n, d in enumerate(self.lines, start=1):
            # 자재코드 경고(cip_warn)와 납품요청일 임박색은 ttk.Treeview가
            # 행 하나에 글자색을 하나만 줄 수 있어 동시에 표시하지 못한다.
            # 자재코드 문제가 더 치명적이므로 그쪽을 우선한다.
            if d["Q"].strip() in cip:
                tags = ("cip_warn",)
            else:
                color = due_date_color(parse_date(d["T"]))
                tags = ("due_%s" % color,) if color else ()
            self.tree.insert("", "end", values=["☐", n] + [d[k] for k in LINE_KEYS],
                             tags=tags)
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
        self.status.config(text="생성 완료 : %s" % out)
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
