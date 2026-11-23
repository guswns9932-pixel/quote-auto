"""
esign_capture.py
================
전자서명용 Excel 시트 → 이미지/PDF 변환 (Excel COM).
견적서 생성 엔진(excel_io, openpyxl)에 의존하지 않는다.
"""

from __future__ import annotations

import io
import logging
import os
import time
from typing import List

from core import SheetName, unique_path
from excel_com import ExcelCOM, _clear_clipboard

logger = logging.getLogger("QuoteApp")


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
                                  xl_app=None, progress_cb=None,
                                  should_cancel=None, blob_sink=None) -> List[str]:
    """
    xlsx ESIGN_TARGET 가시 시트를 클립보드로 캡처 → PNG 저장.
    ExportAsFixedFormat/PrintOut 미사용 → RenameFile 없음.

    xl_app       : 공유 Excel.Application COM 객체. None이면 자체 ExcelCOM 사용.
    progress_cb  : (완료 시트수, 전체 시트수, 시트명) 콜백. 시트 캡처가 끝날 때마다 호출.
    should_cancel: 인자 없이 bool 반환하는 콜백. True 면 남은 시트를 건너뛰고 지금까지
                   캡처한 것만 반환한다 (파일 단위보다 촘촘한 취소 체크).
    blob_sink    : dict 를 넘기면 PNG 를 파일로 쓰지 않고 이 dict 에 bytes 로 담는다
                   ({키: PNG bytes}). 반환 리스트에는 경로 대신 그 키가 들어간다.
                   캡처 결과는 이미 메모리(PIL Image)에 있는데 굳이 네트워크
                   드라이브에 썼다가 다시 읽을 이유가 없어서 만든 경로다
                   — 실측상 저장 구간이 전체의 20%(장당 ~90ms)였다.
                   None 이면 기존처럼 tmp_dir 에 파일로 저장한다.
    반환: PNG 경로(또는 blob_sink 키) 리스트 (순서 = ESIGN_TARGET 순서)
    """
    try:
        from PIL import ImageGrab
    except ImportError:
        logger.error("Pillow(PIL) 없음 — pip install pillow")
        return []

    png_paths: List[str] = []

    def _process(app) -> List[str]:
        # 구간별 시간을 로그에 남긴다. 이 경로는 전부 Excel COM 왕복이라
        # 개발 PC 에서 재현·측정할 수가 없어서, 실제 사용 환경의
        # quote_app.log 를 보고 어디가 느린지 판단하기 위한 계측이다.
        t_open = time.perf_counter()
        wb = app.Workbooks.Open(xlsx_path, ReadOnly=True, UpdateLinks=0, AddToMru=False)
        ms_open = (time.perf_counter() - t_open) * 1000
        ms_prep = ms_cap = ms_save = 0.0
        bytes_saved = 0
        try:
            # ESIGN_TARGET 시트 우선, 없으면 보이는 시트 전체 캡처.
            # 이름 조회 + Visible 확인을 한 번만 하고 ws 프록시를 그대로 들고 있는다
            # (기존엔 이 목록을 만들 때 한 번, 캡처 루프에서 또 한 번 조회했다).
            target_sheets = []   # [(name, ws), ...]
            for name in SheetName.ESIGN_TARGET:
                try:
                    ws = wb.Worksheets(name)
                    if int(ws.Visible) == -1:
                        target_sheets.append((name, ws))
                except Exception:
                    continue
            if not target_sheets:
                for ws in wb.Worksheets:
                    try:
                        if int(ws.Visible) == -1:
                            target_sheets.append((ws.Name, ws))
                    except Exception:
                        continue

            total = len(target_sheets)
            for idx, (name, ws) in enumerate(target_sheets):
                if should_cancel is not None:
                    try:
                        if should_cancel():
                            break
                    except Exception:
                        pass
                safe_name = "".join(c if c not in r'\/:*?"<>|' else "_" for c in name)
                png_path = os.path.join(
                    tmp_dir, f"cap_{file_index:03d}_{idx:02d}_{safe_name}.png")
                try:
                    _t = time.perf_counter()
                    ws.Activate()
                    # 페이지 나누기 미리보기 → 기본 보기로 전환 후 캡처.
                    # [주의] View 는 창 속성이지만, 시트마다 마지막 저장 시점의
                    # 보기 모드를 따로 기억하고 있어 Activate() 로 다른 시트로
                    # 넘어가면 그 시트의 저장된 모드(페이지 나누기 미리보기)가
                    # 되살아날 수 있다. 워크북당 1회로 줄였다가 "1 페이지"
                    # 워터마크가 다시 나타나는 회귀가 있었다 — 반드시 시트마다
                    # 매번 설정해야 한다.
                    try:
                        app.ActiveWindow.View = 1  # xlNormalView
                    except Exception:
                        pass
                    rng = _print_area_range(ws)
                    ms_prep += (time.perf_counter() - _t) * 1000
                    _t = time.perf_counter()
                    rng.CopyPicture(Appearance=1, Format=2)  # xlScreen, xlBitmap
                    img = ImageGrab.grabclipboard()
                    # ── 클립보드 즉시 해제 ──────────────────────────────────────
                    # CopyPicture 후 Windows 클립보드에 CF_ENHMETAFILE(EMF) 핸들이
                    # 남아 있는 상태에서 COM 세션(xl.Quit)이 종료되면 stale 핸들이
                    # 발생한다. 이후 사용자가 별도 Excel에서 Ctrl+C를 시도하면
                    # Windows가 이 핸들을 해제하려다 액세스 위반 → Excel 전체 종료.
                    # 이미지를 읽은 직후 클립보드를 비워 이 경로를 차단한다.
                    _clear_clipboard()
                    ms_cap += (time.perf_counter() - _t) * 1000
                    _t = time.perf_counter()
                    if img is not None:
                        # compress_level=1: 잠깐 쓰고 버릴 이미지라 압축률보다
                        # 인코딩 속도가 낫다(기본값 6 대비 체감 저하 없음).
                        if blob_sink is not None:
                            key = f"mem://{os.path.basename(png_path)}"
                            buf = io.BytesIO()
                            img.save(buf, "PNG", compress_level=1)
                            blob_sink[key] = buf.getvalue()
                            png_paths.append(key)
                            bytes_saved += len(blob_sink[key])
                        else:
                            # 이전 세션 잠금 파일이 남아 있으면 삭제 시도 후 저장
                            _dst = png_path
                            if os.path.exists(_dst):
                                try:
                                    os.remove(_dst)
                                except OSError:
                                    _dst = unique_path(_dst)
                            img.save(_dst, "PNG", compress_level=1)
                            png_paths.append(_dst)
                            try:
                                bytes_saved += os.path.getsize(_dst)
                            except OSError:
                                pass
                    else:
                        logger.warning("클립보드 캡처 실패 (%s / %s)", xlsx_path, name)
                    ms_save += (time.perf_counter() - _t) * 1000
                except Exception as e:
                    logger.warning("시트 캡처 실패 (%s / %s): %s", xlsx_path, name, e)
                if progress_cb is not None:
                    try:
                        progress_cb(idx + 1, total, name)
                    except Exception:
                        pass
        finally:
            _clear_clipboard()   # 워크북 닫기 전 최종 클립보드 해제
            try:
                wb.Close(False)
            except Exception:
                pass
            logger.info(
                "시트 캡처 %s: 총 %.0fms (열기 %.0f / 준비 %.0f / 캡처 %.0f / 인코딩 %.0f, "
                "시트 %d개, PNG %.1fMB)",
                os.path.basename(xlsx_path), ms_open + ms_prep + ms_cap + ms_save,
                ms_open, ms_prep, ms_cap, ms_save, len(png_paths),
                bytes_saved / 1024 / 1024)
        return png_paths

    if xl_app is not None:
        return _process(xl_app)

    with ExcelCOM() as xl:
        return _process(xl.app)
