"""
esign_engine.py
===============
전자서명 엔진 — 화면(Qt) 없이 동작한다.

흐름
----
1. export_workbook()   Excel 이 대상 시트를 시트별 PDF 로 저장 → 바이트로 읽고 즉시 삭제
2. merge_sheet_pdfs()  시트별 PDF 를 파일 하나의 PDF 로 합치고 "몇 쪽이 어느 시트인지" 기록
3. render_page()       화면 표시용으로 한 쪽을 그림으로 변환
4. build_signed_pdf()  원본 PDF 쪽 위에 서명 그림을 얹어 제출용 PDF 저장

예전에는 시트를 클립보드로 복사해 그림으로 받았다(화면 모습 캡처). 그 방식은
① 캡처 중 사용자의 클립보드를 건드리고 ② 화면 배율에 따라 결과가 달라지며
③ 최종 PDF 가 저화질 그림이 되었다. 지금은 Excel 의 "PDF 로 저장"을 그대로 써서
인쇄 결과와 같은 PDF 를 얻고, 그 위에 서명만 얹는다.

좌표는 모두 PDF 포인트(1pt = 1/72 inch), 쪽의 왼쪽 위가 원점이다.

흔적 파일
--------
Excel 은 PDF 를 파일로만 내보낼 수 있어 임시 PDF 가 잠깐 생긴다. 엑셀과 같은
폴더에 TMP_PREFIX 로 시작하는 이름으로 만들고, 읽자마자 지운다(이 환경은 로컬
디스크 쓰기가 막혀 있어 로컬 임시 폴더를 쓸 수 없다). 비정상 종료로 남은 것은
다음에 같은 폴더를 열 때 cleanup_stale_temp() 가 지운다.
"""

from __future__ import annotations

import logging
import os
import time
import uuid
from dataclasses import dataclass, field
from typing import Callable, Dict, List, Optional, Sequence, Tuple

logger = logging.getLogger("QuoteApp")

TMP_PREFIX = "~esign_tmp_"

# 서명 한 개의 기본 크기 — "엑셀 시트를 100% 로 봤을 때"의 포인트.
# 실제 PDF 위의 크기는 여기에 그 시트의 인쇄 배율(PageRef.scale)을 곱한다.
# 시트마다 인쇄 배율이 달라서(갑지 73%, 사양서 62%, 입고검수확인서 95% …)
# 쪽 크기 기준 고정값으로 두면 축소 인쇄되는 시트에서 서명이 칸을 넘친다.
# 값은 실제 사인 칸에 맞춰 사용자가 확인한 크기다(가로세로 비율은 예전 170:40).
# 화면의 "서명 크기 %" 로 조절할 수 있다.
SIGN_W_PT = 129.6
SIGN_H_PT = 30.5


@dataclass(frozen=True)
class PageRef:
    """합쳐진 PDF 의 한 쪽이 어느 시트의 몇 번째 쪽인지."""
    sheet: str
    index_in_sheet: int
    pages_in_sheet: int
    scale: float = 1.0          # 이 시트의 인쇄 배율 (0.62 = 62% 로 축소 인쇄)


@dataclass
class SignFile:
    """불러온 엑셀 한 개 = 합쳐진 PDF 한 개."""
    path: str
    pdf: bytes = b""
    pages: List[PageRef] = field(default_factory=list)
    error: str = ""

    @property
    def name(self) -> str:
        return os.path.basename(self.path)


@dataclass(frozen=True)
class Placement:
    """한 쪽 위에 놓인 서명 하나. (x, y) 는 왼쪽 위 모서리, 단위 pt."""
    sign: int
    x: float
    y: float
    w: float = SIGN_W_PT
    h: float = SIGN_H_PT


def sign_size(page_scale: float, pct: float = 100.0) -> Tuple[float, float]:
    """인쇄 배율 page_scale 인 쪽에 놓을 서명의 크기(pt). pct = 사용자가 고른 크기 %."""
    k = page_scale * pct / 100.0
    return SIGN_W_PT * k, SIGN_H_PT * k


PageKey = Tuple[int, int]            # (파일 순번, 쪽 번호)


# ──────────────────────────────────────────────
# 임시 파일
# ──────────────────────────────────────────────
def _remove_quietly(path: str, tries: int = 5) -> bool:
    """지운다. 백신 검사 등으로 잠깐 잠겨 있을 수 있어 몇 번 다시 시도한다."""
    for i in range(tries):
        try:
            os.remove(path)
            return True
        except FileNotFoundError:
            return True
        except OSError:
            time.sleep(0.1 * (i + 1))
    logger.warning("임시 PDF 를 지우지 못했습니다(다음 실행 때 다시 시도): %s", path)
    return False


def cleanup_stale_temp(folder: str) -> int:
    """folder 에 남아 있는 예전 임시 PDF 를 지운다. 반환: 지운 개수."""
    n = 0
    try:
        names = os.listdir(folder)
    except OSError:
        return 0
    for name in names:
        if name.startswith(TMP_PREFIX) and name.lower().endswith(".pdf"):
            if _remove_quietly(os.path.join(folder, name), tries=1):
                n += 1
    return n


# ──────────────────────────────────────────────
# 1. Excel → 시트별 PDF (Excel COM)
# ──────────────────────────────────────────────
def export_workbook(xl_app, xlsx_path: str, target_sheets: Sequence[str],
                    progress_cb: Optional[Callable[[int, int, str], None]] = None,
                    should_cancel: Optional[Callable[[], bool]] = None,
                    ) -> List[Tuple[str, bytes, dict]]:
    """xlsx 의 대상 시트를 시트별 PDF 로 내보내 [(시트이름, PDF bytes, 인쇄설정)] 로 돌려준다.

    target_sheets 중 보이는 시트만 내보낸다. 하나도 없으면 보이는 시트 전부
    (예전 캡처 방식과 같은 규칙). 시트별로 따로 내보내는 이유는 "몇 쪽이 어느
    시트인지"를 정확히 알기 위해서다 — 서명 위치를 시트 기준으로 기억한다.
    """
    folder = os.path.dirname(os.path.abspath(xlsx_path))
    tag = "%d_%s" % (os.getpid(), uuid.uuid4().hex[:8])
    parts: List[Tuple[str, bytes, dict]] = []

    t0 = time.perf_counter()
    wb = xl_app.Workbooks.Open(xlsx_path, ReadOnly=True, UpdateLinks=0, AddToMru=False)
    ms_open = (time.perf_counter() - t0) * 1000
    ms_export = 0.0
    try:
        targets = []                       # [(이름, 시트)]
        for name in target_sheets:
            try:
                ws = wb.Worksheets(name)
                if int(ws.Visible) == -1:          # xlSheetVisible
                    targets.append((name, ws))
            except Exception:
                continue
        if not targets:
            for ws in wb.Worksheets:
                try:
                    if int(ws.Visible) == -1:
                        targets.append((ws.Name, ws))
                except Exception:
                    continue

        total = len(targets)
        for idx, (name, ws) in enumerate(targets):
            if should_cancel is not None and should_cancel():
                break
            tmp = os.path.join(folder, "%s%s_%02d.pdf" % (TMP_PREFIX, tag, idx))
            t = time.perf_counter()
            try:
                ws.ExportAsFixedFormat(0, tmp)             # 0 = xlTypePDF
                with open(tmp, "rb") as f:
                    parts.append((name, f.read(), _page_setup_info(ws)))
            except Exception as e:
                # 인쇄할 내용이 없는 시트 등 — 그 시트만 건너뛴다
                logger.warning("시트 PDF 저장 실패 (%s / %s): %s", xlsx_path, name, e)
            finally:
                _remove_quietly(tmp)
            ms_export += (time.perf_counter() - t) * 1000
            if progress_cb is not None:
                try:
                    progress_cb(idx + 1, total, name)
                except Exception:
                    pass
    finally:
        try:
            wb.Close(False)                # 디스크의 엑셀은 건드리지 않는다
        except Exception:
            pass
    logger.info("PDF 변환 %s: 열기 %.0fms / 저장·읽기 %.0fms, 시트 %d개, %.1fMB",
                os.path.basename(xlsx_path), ms_open, ms_export, len(parts),
                sum(len(p[1]) for p in parts) / 1024 / 1024)
    return parts


def _page_setup_info(ws) -> dict:
    """인쇄 배율 계산에 필요한 시트의 인쇄 설정을 Excel 에서 읽는다(단위 pt).
    못 읽으면 빈 dict — 그러면 배율은 100% 로 본다."""
    try:
        ps = ws.PageSetup
        info = {"zoom": ps.Zoom, "wide": ps.FitToPagesWide, "tall": ps.FitToPagesTall,
                "left": float(ps.LeftMargin), "right": float(ps.RightMargin),
                "top": float(ps.TopMargin), "bottom": float(ps.BottomMargin)}
        area = ps.PrintArea
        rng = ws.Range(area) if area else ws.UsedRange
        info["area_w"] = float(rng.Width)       # 숨긴 행·열은 0 으로 계산된다
        info["area_h"] = float(rng.Height)
        return info
    except Exception as e:
        logger.warning("인쇄 설정을 읽지 못했습니다 (%s): %s", getattr(ws, "Name", "?"), e)
        return {}


def print_scale(info: dict, page_w: float, page_h: float) -> float:
    """시트의 인쇄 배율(1.0 = 100%)을 구한다.

    배율을 직접 지정한 시트는 그 값을 쓴다. "한 쪽에 맞춤"인 시트는 Excel 이
    배율을 알려주지 않으므로 같은 방식으로 계산한다: 여백을 뺀 종이에
    인쇄영역이 들어가도록 줄이되(늘리지는 않음) 1% 단위로 내림.
    """
    def num(v):
        return float(v) if isinstance(v, (int, float)) and not isinstance(v, bool) and v > 0 else 0.0
    if not info:
        return 1.0
    zoom = num(info.get("zoom"))
    if zoom:
        return zoom / 100.0
    cands = [1.0]
    wide, tall = num(info.get("wide")), num(info.get("tall"))
    aw, ah = num(info.get("area_w")), num(info.get("area_h"))
    if wide and aw:
        cands.append((page_w - info.get("left", 0.0) - info.get("right", 0.0)) * wide / aw)
    if tall and ah:
        cands.append((page_h - info.get("top", 0.0) - info.get("bottom", 0.0)) * tall / ah)
    scale = int(min(cands) * 100) / 100.0
    return scale if scale >= 0.1 else 1.0


# ──────────────────────────────────────────────
# 2. 시트별 PDF → 파일 하나의 PDF
# ──────────────────────────────────────────────
def merge_sheet_pdfs(parts: Sequence[tuple]) -> Tuple[bytes, List[PageRef]]:
    """[(시트이름, PDF bytes[, 인쇄설정])] 를 하나로 합친다. 반환: (PDF bytes, 쪽별 PageRef)."""
    import fitz
    out = fitz.open()
    pages: List[PageRef] = []
    try:
        for part in parts:
            sheet, data = part[0], part[1]
            info = part[2] if len(part) > 2 else {}
            try:
                src = fitz.open(stream=data, filetype="pdf")
            except Exception as e:
                logger.warning("시트 PDF 를 읽지 못했습니다 (%s): %s", sheet, e)
                continue
            try:
                n = src.page_count
                scale = print_scale(info, src[0].rect.width, src[0].rect.height) if n else 1.0
                out.insert_pdf(src)
                pages.extend(PageRef(sheet, i, n, scale) for i in range(n))
            finally:
                src.close()
        if not pages:
            return b"", []
        return out.tobytes(deflate=True, garbage=3), pages
    finally:
        out.close()


def load_workbook_pdf(xl_app, xlsx_path: str, target_sheets: Sequence[str],
                      progress_cb=None, should_cancel=None) -> SignFile:
    """엑셀 한 개를 SignFile 로. 실패해도 예외를 던지지 않고 error 에 사유를 담는다."""
    sf = SignFile(path=xlsx_path)
    try:
        parts = export_workbook(xl_app, xlsx_path, target_sheets, progress_cb, should_cancel)
        sf.pdf, sf.pages = merge_sheet_pdfs(parts)
        if not sf.pages:
            sf.error = "표시할 시트가 없습니다"
    except Exception as e:
        logger.error("PDF 변환 실패 (%s): %s", xlsx_path, e, exc_info=True)
        sf.error = str(e)
    return sf


# ──────────────────────────────────────────────
# 3. 화면 표시용 그림
# ──────────────────────────────────────────────
def open_pdf(data: bytes):
    """PDF bytes 를 연다(호출한 쪽이 close 한다). 한 문서는 한 스레드에서만 쓴다."""
    import fitz
    return fitz.open(stream=data, filetype="pdf")


def page_size(doc, page_no: int) -> Tuple[float, float]:
    r = doc[page_no].rect
    return float(r.width), float(r.height)


def render_page(doc, page_no: int, zoom: float = 2.0) -> Tuple[int, int, int, bytes]:
    """한 쪽을 RGB 그림으로. 반환: (너비px, 높이px, 한 줄 바이트 수, 픽셀 bytes)."""
    import fitz
    pix = doc[page_no].get_pixmap(matrix=fitz.Matrix(zoom, zoom), alpha=False)
    return pix.width, pix.height, pix.stride, bytes(pix.samples)


# ──────────────────────────────────────────────
# 4. 서명을 얹어 제출용 PDF 저장
# ──────────────────────────────────────────────
def build_signed_pdf(files: Sequence[SignFile],
                     placements: Dict[PageKey, List[Placement]],
                     sign_images: Sequence[bytes],
                     out_path: str,
                     progress_cb: Optional[Callable[[int, int], None]] = None) -> int:
    """모든 파일의 쪽을 순서대로 이어 붙이고 서명을 얹어 out_path 에 저장한다.

    원본 PDF 쪽을 그대로 옮기므로 글자·선이 그림으로 바뀌지 않는다.
    쓰다 실패하면 반쯤 쓰인 파일을 남기지 않는다. 반환: 저장한 쪽 수.
    """
    import fitz
    total = sum(len(f.pages) for f in files)
    out = fitz.open()
    done = 0
    try:
        for fi, f in enumerate(files):
            if not f.pages:
                continue
            src = fitz.open(stream=f.pdf, filetype="pdf")
            try:
                base = out.page_count
                out.insert_pdf(src)
            finally:
                src.close()
            for pno in range(len(f.pages)):
                page = out[base + pno]
                for pl in placements.get((fi, pno), []):
                    if not (0 <= pl.sign < len(sign_images)):
                        continue
                    rect = fitz.Rect(pl.x, pl.y, pl.x + pl.w, pl.y + pl.h)
                    # 쪽이 회전돼 있으면 화면에서 본 좌표를 쪽 좌표로 되돌린다
                    rect = rect * page.derotation_matrix
                    page.insert_image(rect, stream=sign_images[pl.sign],
                                      keep_proportion=False, rotate=page.rotation)
                done += 1
                if progress_cb is not None:
                    progress_cb(done, total)
        if out.page_count == 0:
            raise ValueError("저장할 쪽이 없습니다.")
        try:
            out.save(out_path, deflate=True, garbage=3)
        except BaseException:
            _remove_quietly(out_path, tries=1)
            raise
        return out.page_count
    finally:
        out.close()
