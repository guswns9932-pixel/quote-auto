# -*- coding: utf-8 -*-
"""
csp_core.py
===========
CSP 주문접수 업로드 파일 생성기의 화면 없는 부분 — 양식 읽기, CIP/FSC 판정,
값 변환, 출력 파일·로그 쓰기. tkinter 에 의존하지 않아 화면 없이 테스트할 수 있다.
화면은 csp_order_maker.py 에 있다.
"""

import os
import re
import sys
import json
import shutil
import zipfile
import subprocess
import datetime as dt
import xml.etree.ElementTree as ET

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, PatternFill, Border, Alignment
from openpyxl.utils import get_column_letter, column_index_from_string

APP_TITLE = "CSP 주문접수 업로드 파일 생성기  (alpha v0.2)"
TEMPLATE_NAME = "CSP_주문접수_업로드_통합양식.xlsx"
SETTINGS_NAME = "csp_order_maker_settings.json"
LOG_NAME = "대외비_CSP_주문접수_전체로그.xlsx"
# 대외비_ 접두 도입 전에 만들어진 로그 — 단가 이력을 이어 읽기 위해 계속 후보에 넣는다.
LEGACY_LOG_NAME = "CSP_주문접수_전체로그.xlsx"
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
    """헤더 텍스트 비교용 정규화: 공백/하이픈/마침표/별표 제거 후 대문자.
    (별표는 FSC매핑 시트가 필수 열 표시로 "*사업장"처럼 쓰기 때문에 포함)"""
    if v is None:
        return ""
    return re.sub(r"[\s\-\.\*]+", "", str(v).strip()).upper()


def _header_grid(ws, rows=_CIP_HEADER_SCAN_ROWS, cols=_CIP_MAX_COL):
    """헤더 영역을 한 번만 읽어 정규화된 2차원 리스트로 돌려준다.

    read_only 워크시트에서 ws.cell(r, c) 임의 접근은 호출마다 XML 스트림을
    처음부터 다시 훑어 O(n²)이 된다. 헤더 한 칸 찾자고 시트를 수백 번 다시
    읽던 것이 통합양식 로드 1.28초 중 0.65초를 차지했다(실측). 아래 찾기
    함수들은 전부 이 격자를 받아서 메모리에서만 뒤진다.
    """
    grid = [["" for _ in range(cols)] for _ in range(rows)]
    for r, row in enumerate(ws.iter_rows(min_row=1, max_row=rows,
                                         max_col=cols, values_only=True)):
        if r >= rows:
            break
        for c, v in enumerate(row):
            if c < cols:
                grid[r][c] = _h(v)
    return grid


def _find_header_col(grid, header_row, label):
    """단일 헤더 행에서 label과 일치하는 열 번호(1-based)를 찾는다.
    (FSC매핑 시트처럼 헤더가 한 줄뿐인 단순한 표에 쓴다 — CIP 시트처럼
    병합된 여러 줄 헤더가 필요하면 _find_cip_col/_find_cip_subheader_col
    을 쓴다.)"""
    target = _h(label)
    if not target or header_row > len(grid):
        return None
    for c, v in enumerate(grid[header_row - 1], start=1):
        if v == target:
            return c
    return None


def _find_cip_col(grid, label):
    """헤더 행들 중 어느 셀이든 label과 일치하면 그 열 번호(1-based)를 반환."""
    target = _h(label)
    if not target:
        return None
    for row in grid:
        for c, v in enumerate(row, start=1):
            if v == target:
                return c
    return None


def _find_cip_subheader_col(grid, section_label, sub_label):
    """병합된 섹션 헤더(예: 'AS-IS') 아래에 있는 서브헤더(예: 'FSC') 열을 찾는다.

    read_only 모드로 열면 병합 셀 정보를 읽을 수 없어(맨 왼쪽 셀에만 값이
    있고 나머지는 None) merged_cells를 쓸 수 없다. 대신 한 행을 왼쪽에서
    가장 가까운 값으로 채워 넣어(엑셀 화면에 보이는 대로) 각 열이 어느
    섹션에 속하는지 판단한다.
    """
    section_target = _h(section_label)
    sub_target = _h(sub_label)
    for r, row in enumerate(grid):
        filled, last = [], ""
        for v in row:
            if v:
                last = v
            filled.append(last)
        if section_target not in filled:
            continue
        for r2 in range(r + 1, min(r + 3, len(grid) - 1) + 1):
            for c, v in enumerate(filled, start=1):
                if v == section_target and grid[r2][c - 1] == sub_target:
                    return c
    return None


# ---------------------------------------------------------------- CIP 매치(AS-IS FSC 알람)
# "-"(대공정·설비사·세부공정)은 CIP 시트에서 "해당 항목 전체에 적용됨"을
# 뜻하는 값이라 와일드카드로 취급한다. 옵션 칸이 아직 비어 있는 경우는
# (아직 입력 안 함) 와일드카드로 보지 않는다 — 그래야 세부공정만 빈
# 상태에서도 "세부공정 제외 동일"(orange)로 자연스럽게 떨어진다.
#
# 사업장의 "ALL"은 와일드카드로 두지 않는다 — 실사용자 확인 결과 "ALL"은
# 다른 사업장(SCS/TAYLOR 등)을 포괄하는 개념이 아니라 그 자체로 하나의
# 개별 사업장이다. 그래서 사업장은 다른 필드와 똑같이 정확히 일치할
# 때만 매치되고("ALL"을 직접 선택했을 때만 CIP의 "ALL" 행과 매치),
# 별도 취급이 필요 없다.
_CIP_WILDCARDS = {"-"}


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


def _split_fsc_codes(v):
    """CIP AS-IS/TO-BE 칸의 '코드1 + 코드2'처럼 묶인 복합 자재코드를
    개별 코드 리스트로 나눈다(양쪽 공백 트림, 빈 조각은 버림)."""
    return [p.strip() for p in str(v or "").split("+") if p.strip()]


def _cip_fsc_match_index(row_fsc, fsc_n):
    """CIP 행의 AS-IS FSC(단일 또는 '+'로 묶인 복합값) 중 fsc_n과 일치하는
    조각의 인덱스를 반환한다. 일치하는 조각이 없으면 -1."""
    for i, p in enumerate(_split_fsc_codes(row_fsc)):
        if _norm_plain(p) == fsc_n:
            return i
    return -1


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
        if _cip_fsc_match_index(r.get("fsc"), fsc_n) < 0:
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


def cip_tobe_recommend(cip_rows, site, device, process, vendor, subproc, fsc):
    """AS-IS로 감지된 자재코드(fsc)에 대해 CIP TO-BE(L열) 교체 추천 코드를
    찾는다. AS-IS가 "코드1 + 코드2"처럼 여러 코드가 묶인 세트면, 일치한
    코드와 같은 위치의 TO-BE 코드만 짚어서 추천한다(세트 전체가 아니라
    실제로 바꿔야 할 그 코드 하나).

    반환: {"level": "red"/"orange", "tobe": "<코드>"} 또는 일치/추천할
    TO-BE 값이 없으면 None. level은 cip_match_level과 같은 기준
    (세부공정까지 일치=red, 나머지만 일치=orange)이며, red가 하나라도
    있으면 그걸 우선한다.
    """
    fsc_n = _norm_plain(fsc)
    if not fsc_n or not cip_rows:
        return None
    best = None
    for r in cip_rows:
        idx = _cip_fsc_match_index(r.get("fsc"), fsc_n)
        if idx < 0:
            continue
        if not _cip_field_eq(r.get("site"), site):
            continue
        if not _cip_field_eq(r.get("device"), device):
            continue
        if not _cip_field_eq(r.get("process"), process):
            continue
        if not _cip_field_eq(r.get("vendor"), vendor):
            continue

        tobe_parts = _split_fsc_codes(r.get("tobe"))
        if not tobe_parts:
            continue
        tobe = tobe_parts[idx] if idx < len(tobe_parts) else tobe_parts[0]

        if _cip_field_eq(r.get("subproc"), subproc, normalize=_norm_subproc):
            return {"level": "red", "tobe": tobe}
        if best is None:
            best = {"level": "orange", "tobe": tobe}
    return best


def _fsc_map_key(site, device, process, vendor, subproc, qcode):
    return (_norm_plain(site), _norm_plain(device), _norm_plain(process),
            _norm_plain(vendor), _norm_subproc(subproc), _norm_plain(qcode))


def _fsc_map_key5(site, device, process, vendor, qcode):
    return (_norm_plain(site), _norm_plain(device), _norm_plain(process),
            _norm_plain(vendor), _norm_plain(qcode))


def fsc_recommend(fsc_map, fsc_map_5key, site, device, process, vendor, subproc, qcode):
    """옵션(사업장/DEVICE/대공정/설비사/세부공정/Q-code) 조합으로 FSC매핑
    시트를 조회해 추천 자재코드를 찾는다.

    세부공정 표기가 워낙 다양해서(오타/약어 등) 6개 조건을 다 요구하면
    추천이 거의 안 뜬다 — 그래서 세부공정을 제외한 5개 조건(사업장/
    DEVICE/대공정/설비사/Q-code)이 일치하면 "후보" 추천을 먼저 주고,
    세부공정까지 정확히 일치하면 "확정" 추천으로 격상한다.

    반환: {"fsc": 추천 자재코드, "confirmed": bool, "prev": 이전 추천 코드
    또는 None} 또는 일치하는 조합이 전혀 없으면 None. confirmed=True는
    6개 조건 전부 일치, False는 세부공정만 다른 5개 조건 일치. prev는
    같은 조합(6키)의 이력에서 현재 추천 바로 앞에 있던 값 — 동일조건에
    새 FSC가 등록되면서 추천이 바뀐 경우, 방금까지 추천되던 코드를
    가리킨다(6키가 정확히 일치할 때만 의미가 있다).
    """
    entry6 = fsc_map.get(_fsc_map_key(site, device, process, vendor, subproc, qcode))
    if entry6 and entry6.get("history"):
        history = entry6["history"]
        prev = history[-2] if len(history) >= 2 else None
        return {"fsc": history[-1], "confirmed": True, "prev": prev}
    entry5 = fsc_map_5key.get(_fsc_map_key5(site, device, process, vendor, qcode))
    if entry5:
        return {"fsc": entry5["fsc"], "confirmed": False, "prev": None}
    return None


_NS_R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"


def _xlsx_sheet_part(path, sheet_name):
    """워크북 안에서 시트 이름에 해당하는 실제 zip 내부 경로
    (예: xl/worksheets/sheet11.xml)를 workbook.xml + 관계 파일로 찾는다.
    시트를 못 찾으면 None."""
    with zipfile.ZipFile(path, "r") as z:
        wb_root = ET.fromstring(z.read("xl/workbook.xml"))
        rels_root = ET.fromstring(z.read("xl/_rels/workbook.xml.rels"))
    rid = None
    for sheet_el in wb_root.iter(f"{{{_NS_MAIN}}}sheet"):
        if sheet_el.get("name") == sheet_name:
            rid = sheet_el.get(f"{{{_NS_R}}}id")
            break
    if not rid:
        return None
    for rel_el in rels_root:
        if rel_el.get("Id") == rid:
            target = rel_el.get("Target", "")
            return target if target.startswith("xl/") else "xl/" + target
    return None


_BACKUP_KEEP = 20


def _prune_backups(backup_dir, base, keep=_BACKUP_KEEP):
    """Backup/ 에 같은 양식의 백업이 끝없이 쌓이지 않도록 최신 keep 개만 남긴다."""
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


def _sync_fsc_map_file(template_path, new_rows, history_updates):
    """FSC매핑 시트를 갱신한다 — 둘 다 한 번의 백업/저장으로 함께 적용한다.

    new_rows        : 새로 추가할 조합 행. {"site","device","process",
                       "vendor","subproc","qcode","des","fsc"} 키를 가진
                       dict 목록 — 이미 FSC매핑에 있는 조합인지는 호출부
                       에서 걸러서 보낸다(여기서는 그냥 추가만 한다).
    history_updates  : 이미 있는 행에 새 FSC 이력을 추가할 목록.
                       [{"row": 엑셀 행번호, "col": 열 번호, "fsc": 값}, ...]

    ★ openpyxl로 전체를 다시 읽고 저장하지 않는다 ★ — 실제 FSC매핑
    시트의 사업장/DEVICE 열은 외부(현재 연결 안 된) 워크북을 참조하는
    VLOOKUP 수식이고, 그 계산된 값이 수식과 함께 셀 XML에 캐시돼 있다.
    openpyxl은 수식을 계산하지 않으므로, 이 시트를 openpyxl로 열었다가
    그대로 다시 저장하기만 해도 그 캐시된 값이 사라져(수식은 남지만
    "결과 없음" 상태가 됨) 기존 행들의 사업장/DEVICE가 전부 빈 칸이
    되는 것을 실제로 재현해서 확인했다. 그래서 zip 안의 해당 시트 XML
    텍스트만 직접 파싱해 필요한 <row>/<c>만 추가·수정하고, 그 외의
    모든 내용(다른 시트, 이 시트의 기존 셀·수식 캐시값, 스타일 등)은
    완전히 그대로 둔다.

    template_path는 여러 사용자가 공유하는 마스터 파일이라 직접
    덮어쓴다. 저장 전 같은 폴더의 Backup/ 밑에 타임스탬프를 붙여
    원본을 복사해두고, 저장 자체가 실패하면(다른 프로그램에서 파일을
    열어둔 경우 등) 방금 만든 백업을 지우고 예외를 그대로 올린다
    (원본이 그대로 남아 있으니 백업이 따로 필요 없다).

    반환: (추가한 새 행 수, 갱신한 이력 칸 수).
    """
    if not new_rows and not history_updates:
        return 0, 0

    # 열 위치는 안전한 read_only 모드로만 확인한다(저장하지 않으므로
    # 기존 내용에 전혀 영향이 없다).
    wb_ro = load_workbook(template_path, read_only=True, data_only=True)
    try:
        if "FSC매핑" not in wb_ro.sheetnames:
            return 0, 0
        ws_ro = wb_ro["FSC매핑"]
        hdr = _header_grid(ws_ro, rows=1, cols=30)
        col_map = {
            "site": _find_header_col(hdr, 1, "사업장"),
            "device": _find_header_col(hdr, 1, "DEVICE"),
            "process": _find_header_col(hdr, 1, "대공정"),
            "vendor": _find_header_col(hdr, 1, "설비사"),
            "subproc": _find_header_col(hdr, 1, "세부공정"),
            "qcode": (_find_header_col(hdr, 1, "Q-code")
                     or _find_header_col(hdr, 1, "Qcode")),
            "des": _find_header_col(hdr, 1, "DES"),
            "fsc": _find_header_col(hdr, 1, "FSC"),
        }
        max_row = ws_ro.max_row
        max_col = max(ws_ro.max_column, max(c for c in col_map.values() if c) or 0,
                      max((u["col"] for u in history_updates), default=0))
    finally:
        wb_ro.close()

    required = ("site", "device", "process", "vendor", "qcode", "fsc")
    if new_rows and not all(col_map.get(k) for k in required):
        new_rows = []
    if not new_rows and not history_updates:
        return 0, 0

    sheet_part = _xlsx_sheet_part(template_path, "FSC매핑")
    if not sheet_part:
        return 0, 0

    backup_dir = os.path.join(os.path.dirname(template_path) or ".", "Backup")
    os.makedirs(backup_dir, exist_ok=True)
    base, ext = os.path.splitext(os.path.basename(template_path))
    if not base.startswith("대외비_"):
        base = "대외비_" + base
    stamp = dt.datetime.now().strftime("%Y%m%d_%H%M%S")
    backup_path = os.path.join(backup_dir, "%s.backup_%s%s" % (base, stamp, ext))
    n = 1
    while os.path.exists(backup_path):
        n += 1
        backup_path = os.path.join(backup_dir, "%s.backup_%s_%d%s" % (base, stamp, n, ext))
    shutil.copy2(template_path, backup_path)
    _prune_backups(backup_dir, base)

    try:
        _xlsx_update_fsc_map_xml(template_path, sheet_part, max_row, max_col,
                                 new_rows, col_map, history_updates)
    except Exception:
        try:
            os.remove(backup_path)
        except OSError:
            pass
        raise
    return len(new_rows), len(history_updates)


def _xml_escape_text(s):
    return s.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")


def _inline_str_cell_xml(col, row_num, text):
    ref = "%s%d" % (get_column_letter(col), row_num)
    space_attr = ' xml:space="preserve"' if text != text.strip() else ""
    return '<c r="%s" t="inlineStr"><is><t%s>%s</t></is></c>' % (
        ref, space_attr, _xml_escape_text(text))


def _xlsx_set_cell_text(xml, row_num, col, text):
    """시트 XML 텍스트(문자열) 안에서 특정 행(row_num)의 특정 열(col)
    칸을 텍스트 값으로 설정한다. 이미 그 칸에 <c>가 있으면(예: 이력이
    없음을 나타내는 잔여 0 값) 그 <c> 하나만 통째로 교체하고, 없으면
    열 순서를 지켜 새로 끼워 넣는다. 그 행/그 칸 밖의 XML은 문자 하나
    바뀌지 않는다."""
    row_pat = re.compile(r'(<row r="%d"[^>]*>)(.*?)(</row>)' % row_num, re.DOTALL)
    m = row_pat.search(xml)
    if not m:
        return xml
    open_tag, inner, close_tag = m.group(1), m.group(2), m.group(3)

    ref = "%s%d" % (get_column_letter(col), row_num)
    new_cell = _inline_str_cell_xml(col, row_num, text)
    cell_pat = re.compile(r'<c r="%s"(?:[^>]*/>|[^>]*>.*?</c>)' % re.escape(ref), re.DOTALL)
    cm = cell_pat.search(inner)
    if cm:
        new_inner = inner[:cm.start()] + new_cell + inner[cm.end():]
    else:
        insert_pos = len(inner)
        for pm in re.finditer(r'<c r="([A-Z]+)%d"' % row_num, inner):
            if column_index_from_string(pm.group(1)) > col:
                insert_pos = pm.start()
                break
        new_inner = inner[:insert_pos] + new_cell + inner[insert_pos:]

    return xml[:m.start()] + open_tag + new_inner + close_tag + xml[m.end():]


def _xlsx_update_fsc_map_xml(path, sheet_part, start_row, max_col, new_rows, col_map,
                             history_updates):
    """xlsx zip 안의 시트 XML에 새 <row>를 추가하고/또는 기존 <row>의
    이력 칸을 갱신한다. col_map은 {필드명: 열 번호(1-based) 또는 None}
    — 값이 없는 필드는 그 칸을 아예 비워 둔다.

    ★ ElementTree로 전체 문서를 파싱해 다시 직렬화하지 않는다 ★ —
    실제로 그렇게 했다가 파일이 열리지 않는 문제를 겪었다: 이 시트의
    루트 태그에는 mc:Ignorable="x14ac xr xr2 xr3" 처럼 접두어 이름을
    "문자열 값"으로 나열하는 속성이 있는데, x14ac/xr2/xr3 네임스페이스는
    실제 태그·속성에는 전혀 쓰이지 않고 그 문자열 안에서만 언급된다.
    ElementTree는 이런 문자열까지 이해하지 못해서, 다시 저장할 때 실제로
    "쓰인" 네임스페이스만 남기고 나머지 선언을 통째로 지워버리며(mc→ns1,
    xr→ns2처럼 이름도 바뀜), 그 결과 mc:Ignorable 값이 가리키는 접두어와
    실제 선언이 어긋나 엑셀이 파일을 손상된 것으로 인식했다(첨부해주신
    파일로 재현·확인함). 그래서 시트 XML을 순수 텍스트로만 다루고, 건드
    리는 셀/행 밖의 내용은 글자 하나도 바꾸지 않는다.
    """
    with zipfile.ZipFile(path, "r") as zin:
        data = {name: zin.read(name) for name in zin.namelist()}

    xml = data[sheet_part].decode("utf-8")

    # 1) 이미 있는 행에 새 FSC 이력 칸을 채운다.
    for upd in history_updates:
        xml = _xlsx_set_cell_text(xml, upd["row"], upd["col"], str(upd["fsc"]))

    # 2) 새로운 조합을 새 행으로 추가한다.
    row_num = start_row
    cells_order = sorted(((c, f) for f, c in col_map.items() if c), key=lambda x: x[0])
    new_row_chunks = []
    for values in new_rows:
        row_num += 1
        cell_xml = []
        for col, field in cells_order:
            text = str(values.get(field, "") or "")
            if not text:
                continue
            cell_xml.append(_inline_str_cell_xml(col, row_num, text))
        new_row_chunks.append('<row r="%d" spans="1:%d">%s</row>' %
                              (row_num, max_col, "".join(cell_xml)))

    if new_row_chunks:
        marker = "</sheetData>"
        idx = xml.rindex(marker)
        xml = xml[:idx] + "".join(new_row_chunks) + xml[idx:]

    xml = re.sub(
        r'<dimension ref="([^":]+):[^"]+"\s*/>',
        lambda dm: '<dimension ref="%s:%s%d"/>' % (
            dm.group(1), get_column_letter(max_col), max(row_num, start_row)),
        xml, count=1)

    data[sheet_part] = xml.encode("utf-8")

    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as zout:
        for name, content in data.items():
            zout.writestr(name, content)


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
        # CIP 시트: AS-IS FSC 알람 매치 데이터 (옵션 드롭다운 목록은 FSC매핑
        # 시트에서 뽑는다 — 그래야 같은 조건으로 FSC 추천과 CIP 알람을
        # 함께 확인할 수 있다).
        self.cip_rows = []         # [{"site","device","process","vendor","subproc","fsc"}, ...]

        # FSC매핑 시트: 옵션(사업장/DEVICE/대공정/설비사/세부공정/Q-code)
        # 조합별 추천 자재코드 + 이력, 그리고 옵션 드롭다운 목록.
        self.fsc_map = {}          # {정규화된 6키 튜플: {"row": 엑셀행번호, "history": [FSC, ...]}}
        # 세부공정을 제외한 5키(사업장/DEVICE/대공정/설비사/Q-code) 기준
        # 후보 추천용 — 같은 5키에 여러 세부공정 변형이 있으면 엑셀 행번호가
        # 가장 큰(가장 나중에 추가된) 것을 대표값으로 쓴다.
        self.fsc_map_5key = {}     # {정규화된 5키 튜플: {"row": 엑셀행번호, "fsc": 최근 FSC}}
        self.fsc_map_qcodes = []   # 옵션 Q-code 콤보박스용 고유값
        self.fsc_map_fsc_col = 0  # FSC(첫 이력) 열 번호 — 새 이력을 쓸 때 기준
        # DES(H열, 규격) <-> Q-code — 거의 1:1 대응이라, 규격을 검색해
        # 고르면 Q-code를 자동으로 채워줄 수 있다(그 반대도 마찬가지).
        self.opt_des_values = []          # 옵션 규격 콤보박스용 고유값(원문 그대로)
        self.fsc_map_des_to_qcode = {}    # {정규화된 DES: Q-code}
        self.fsc_map_qcode_to_des = {}    # {정규화된 Q-code: DES(원문)}
        self.opt_sites = []
        self.opt_devices = []
        self.opt_processes = []
        self.opt_vendors = []
        self.opt_subprocs = []     # 정규화(특수문자→'-', 대소문자 무시)된 고유값
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
                hdr         = _header_grid(ws)
                col_no      = _find_cip_col(hdr, "No.") or _find_cip_col(hdr, "No")
                col_site    = _find_cip_col(hdr, "사업장")
                col_device  = _find_cip_col(hdr, "DEVICE")
                col_process = _find_cip_col(hdr, "대공정")
                col_vendor  = _find_cip_col(hdr, "설비사")
                col_subproc = _find_cip_col(hdr, "세부공정")
                col_fsc     = _find_cip_subheader_col(hdr, "AS-IS", "FSC")
                col_tobe    = _find_cip_subheader_col(hdr, "TO-BE", "FSC")

                cip_rows = []
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
                        tobe = _cell(col_tobe) if col_tobe else ""
                        cip_rows.append({"site": site, "device": device,
                                         "process": process, "vendor": vendor,
                                         "subproc": subproc, "fsc": fsc, "tobe": tobe})

                self.cip_rows = cip_rows

            if "FSC매핑" in wb.sheetnames:
                ws = wb["FSC매핑"]
                hdr         = _header_grid(ws, rows=1, cols=30)
                col_site    = _find_header_col(hdr, 1, "사업장")
                col_device  = _find_header_col(hdr, 1, "DEVICE")
                col_process = _find_header_col(hdr, 1, "대공정")
                col_vendor  = _find_header_col(hdr, 1, "설비사")
                col_subproc = _find_header_col(hdr, 1, "세부공정")
                col_qcode   = _find_header_col(hdr, 1, "Q-code") or _find_header_col(hdr, 1, "Qcode")
                col_des     = _find_header_col(hdr, 1, "DES")
                col_fsc     = _find_header_col(hdr, 1, "FSC")

                def _hist_val(v):
                    """FSC 이력 칸이 '진짜' 값인지 판단한다. 일부 행에 남아있는
                    잔여 숫자 0을 실제 FSC로 착각하지 않기 위함."""
                    s = self._s(v)
                    return s if s and s != "0" else ""

                fsc_map = {}
                fsc_map_5key = {}
                fsc_map_des_to_qcode = {}
                fsc_map_qcode_to_des = {}
                qcodes = set()
                des_values = set()
                sites, devices, processes, vendors, subprocs = set(), set(), set(), set(), set()
                if col_site and col_device and col_process and col_vendor and col_qcode and col_fsc:
                    # read_only 모드에서는 ws.cell(row, col) 랜덤 접근이 매우
                    # 느려(행마다 스트림을 다시 훑음) 대량 행에서는 사실상
                    # 멈춘 것처럼 보인다 — CIP 파싱과 같이 iter_rows로 한 번만
                    # 순차로 훑는다.
                    for r_idx, row in enumerate(ws.iter_rows(min_row=2, values_only=True), start=2):
                        def _cell(col):
                            return self._s(row[col - 1]) if col and len(row) >= col else ""
                        site, device = _cell(col_site), _cell(col_device)
                        process, vendor = _cell(col_process), _cell(col_vendor)
                        subproc = _cell(col_subproc) if col_subproc else ""
                        qcode = _cell(col_qcode)
                        if not (site and device and process and vendor and qcode):
                            continue

                        # 옵션 드롭다운 목록: FSC 이력 유무와 상관없이 FSC매핑에
                        # 있는 조합이면 선택 가능한 값으로 취급한다.
                        sites.add(site); devices.add(device)
                        processes.add(process); vendors.add(vendor)
                        if subproc: subprocs.add(_norm_subproc(subproc))
                        qcodes.add(qcode)

                        # 규격(DES) <-> Q-code: 처음 나온 값을 대표값으로 쓴다
                        # (실제 데이터는 대소문자 차이 정도만 있고 사실상 1:1).
                        des = _cell(col_des) if col_des else ""
                        if des:
                            des_values.add(des)
                            fsc_map_des_to_qcode.setdefault(_norm_plain(des), qcode)
                            fsc_map_qcode_to_des.setdefault(_norm_plain(qcode), des)

                        # FSC(첫 이력 칸)부터 오른쪽으로 값이 있는 동안만 이력에 담는다.
                        # (이력 중간에 빈 칸이 나오면 그 뒤는 아직 안 쓴 것으로 본다)
                        history = []
                        c = col_fsc
                        while c <= len(row):
                            v = _hist_val(row[c - 1])
                            if not v:
                                break
                            history.append(v)
                            c += 1
                        if not history:
                            continue

                        key = (_norm_plain(site), _norm_plain(device), _norm_plain(process),
                               _norm_plain(vendor), _norm_subproc(subproc), _norm_plain(qcode))
                        fsc_map[key] = {"row": r_idx, "history": history}
                        # 세부공정을 뺀 5키는 행번호 순으로 훑으므로 마지막에
                        # 덮어쓴 값이 자연히 "가장 나중 행"이 된다.
                        key5 = _fsc_map_key5(site, device, process, vendor, qcode)
                        fsc_map_5key[key5] = {"row": r_idx, "fsc": history[-1]}

                self.fsc_map = fsc_map
                self.fsc_map_5key = fsc_map_5key
                self.fsc_map_qcodes = sorted(qcodes)
                self.fsc_map_fsc_col = col_fsc or 0
                self.opt_des_values = sorted(des_values)
                self.fsc_map_des_to_qcode = fsc_map_des_to_qcode
                self.fsc_map_qcode_to_des = fsc_map_qcode_to_des
                self.opt_sites = sorted(sites)
                self.opt_devices = sorted(devices)
                self.opt_processes = sorted(processes)
                self.opt_vendors = sorted(vendors)
                self.opt_subprocs = sorted(subprocs)
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
    "material": "F",    # Material -> 옵션 Q-code
    "desc": "G",        # Material Description -> 'LOT,' 뒤 값으로 자재코드 검색
    "qty": "H",         # 수량 -> 생성수량
    "line": "K",        # 라인 -> '_' 뒤 값으로 대공정
    "subprocess": "N",  # 세부공정 -> 고객세부공정 + 옵션 세부공정
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


_MODEL_KEYWORD_RE = re.compile(r"^[A-Za-z]+[0-9]+")


def extract_model_keyword(text):
    """'DRY_PUMP;EQ,LOT,HD4500PW' -> 'HD4500'
    ('LOT,' 뒤 값에서 스펠링+숫자까지만 추출하고, 그 뒤에 붙는 서브모델
    스펠링(PW 등)은 잘라낸다). 패턴과 안 맞으면 'LOT,' 뒤 값 전체를
    그대로 돌려준다."""
    text = str(text or "")
    marker = "LOT,"
    idx = text.find(marker)
    if idx == -1:
        return ""
    after = text[idx + len(marker):].strip()
    m = _MODEL_KEYWORD_RE.match(after)
    return m.group(0) if m else after


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
# 화면 이름과 양식 헤더 이름이 다른 열 (키 -> 양식 1행의 헤더 글자)
_OUTPUT_HEADER_ALIAS = {
    "M": "고객공정",    # 화면: 대공정
    "N": "고객설비",    # 화면: 설비MAKER
}


def _norm_out_header(text):
    """양식 헤더 비교용: 필수 표시(*)와 공백 제거."""
    return re.sub(r"[\s*]+", "", str(text or ""))


def _map_output_columns(tpl_ws):
    """COLUMNS 의 각 키가 양식 Sheet1 의 몇 번째 열(1-based)인지 헤더 글자로 찾는다.
    찾지 못하는 항목이 있으면 엉뚱한 열에 쓰지 않도록 오류로 멈춘다."""
    by_header = {}
    # 양식에 열이 추가되어 뒤로 밀린 헤더도 찾도록 양식의 마지막 열까지 본다.
    for idx in range(1, max(tpl_ws.max_column, len(COLUMNS)) + 1):
        name = _norm_out_header(tpl_ws.cell(row=1, column=idx).value)
        if name and name not in by_header:
            by_header[name] = idx
    col_of, missing = {}, []
    for key, header, _ in COLUMNS:
        want = _OUTPUT_HEADER_ALIAS.get(key, header)
        idx = by_header.get(_norm_out_header(want))
        if idx is None:
            missing.append("%s(%s)" % (want, key))
        else:
            col_of[key] = idx
    if missing:
        raise ValueError(
            "통합양식 Sheet1 1행에서 다음 열을 찾지 못했습니다: %s\n"
            "양식의 헤더 이름이 바뀌었는지 확인하세요." % ", ".join(missing))
    return col_of


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

    # --- 데이터를 넣을 열: 위치가 아니라 양식 1행의 헤더 글자로 찾는다.
    # (양식의 열 이름이 화면 이름과 다르거나 순서가 달라도 값이 엉뚱한 열에
    #  들어가지 않도록 — 예전에는 열 위치를 코드에 고정해서 양식의
    #  '고객설비/고객공정' 열에 대공정/설비MAKER 값이 서로 바뀌어 들어갔다.)
    col_of = _map_output_columns(tws)

    # --- 헤더행 : 원본 서식 그대로 복사
    for idx in range(1, max(tws.max_column, len(COLUMNS)) + 1):
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
        for key, _, _ in COLUMNS:
            cell = ws.cell(row=r, column=col_of[key], value=data.get(key))
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
    folder = upload_dir()
    names = []
    for log_name in (LOG_NAME, LEGACY_LOG_NAME):
        base, ext = os.path.splitext(log_name)
        names.append(r"%s(_\d+)?%s" % (re.escape(base), re.escape(ext)))
    pattern = re.compile(r"^(?:%s)$" % "|".join(names))
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
