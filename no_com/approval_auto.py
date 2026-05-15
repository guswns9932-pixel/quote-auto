"""
approval_auto.py
================
Selenium 기반 전자결재 자동화 모듈.
구매요청서(NPN) 결재상신 워크플로를 자동으로 진행한다.

연결 방식
---------
앱 전용 Chrome 프로필을 별도로 실행한다(기존 Chrome 창에 영향 없음).
ID/PW를 받아 자동 로그인 후 결재상신을 진행한다.

요구 패키지:
    pip install selenium webdriver-manager
"""
from __future__ import annotations

import os
import sys
import time
import logging
from typing import Callable, Dict, Optional

logger = logging.getLogger("QuoteApp.approval")

APPROVAL_URL = "https://groupware.lotvacuum.com/app/approval"
FORM_NAME    = "구매요청서(NPN)"


# ──────────────────────────────────────────────────────────────────────────────
# 내부 유틸
# ──────────────────────────────────────────────────────────────────────────────
def _profile_dir() -> str:
    """앱 전용 Chrome 프로필 경로 (AppData\\Roaming — 쓰기 권한 보장)."""
    appdata = os.environ.get("APPDATA") or os.path.expanduser("~")
    return os.path.join(appdata, "quote-auto", "approval_profile")


def _find_chrome_binary() -> Optional[str]:
    candidates = [
        os.path.join(os.environ.get("PROGRAMFILES", ""),
                     "Google", "Chrome", "Application", "chrome.exe"),
        os.path.join(os.environ.get("PROGRAMFILES(X86)", ""),
                     "Google", "Chrome", "Application", "chrome.exe"),
        os.path.join(os.environ.get("LOCALAPPDATA", ""),
                     "Google", "Chrome", "Application", "chrome.exe"),
    ]
    for p in candidates:
        if os.path.isfile(p):
            return p
    return None


# ──────────────────────────────────────────────────────────────────────────────
# WebDriver 생성
# ──────────────────────────────────────────────────────────────────────────────
def _patch_chrome_shortcut() -> None:
    """바탕화면 Chrome 바로가기에 --remote-debugging-port=9222 를 자동으로 추가한다."""
    import glob
    try:
        import win32com.client
    except ImportError:
        logger.warning("pywin32 없음 - 바로가기 자동 수정 건너뜀 (pip install pywin32)")
        return

    shell = win32com.client.Dispatch("WScript.Shell")
    desktops = [
        shell.SpecialFolders("Desktop"),
        os.path.join(os.environ.get("PUBLIC", r"C:\Users\Public"), "Desktop"),
    ]

    for desktop in desktops:
        for lnk in glob.glob(os.path.join(desktop, "*.lnk")):
            try:
                sc = shell.CreateShortcut(lnk)
                if "chrome" not in (sc.TargetPath or "").lower():
                    continue
                args = sc.Arguments or ""
                if "--remote-debugging-port=9222" in args:
                    continue
                sc.Arguments = (args + " --remote-debugging-port=9222").strip()
                sc.Save()
                logger.info("Chrome 바로가기 수정 완료: %s", lnk)
            except Exception as e:
                logger.warning("바로가기 수정 실패 (%s): %s", lnk, e)


def _create_driver():
    """
    원격 디버깅 포트로 실행 중인 Chrome에 연결하거나,
    없으면 subprocess로 Chrome을 직접 실행 후 연결한다.

    사전 준비 (최초 1회):
        Chrome 바탕화면 바로가기 대상에 --remote-debugging-port=9222 추가
        예) "C:\\...\\chrome.exe" --remote-debugging-port=9222
    """
    _patch_chrome_shortcut()  # 바로가기에 디버깅 포트 자동 추가

    try:
        from selenium import webdriver
        from selenium.webdriver.chrome.options import Options
        from selenium.webdriver.chrome.service import Service
    except ImportError:
        raise RuntimeError(
            "selenium 패키지가 없습니다.\n"
            "pip install selenium webdriver-manager"
        )

    try:
        from webdriver_manager.chrome import ChromeDriverManager
        service = Service(ChromeDriverManager().install())
    except Exception:
        service = Service()

    debug_port = 9222

    # 이미 디버깅 포트로 실행 중인 Chrome에 연결 시도
    try:
        opts = Options()
        opts.debugger_address = f"localhost:{debug_port}"
        driver = webdriver.Chrome(service=service, options=opts)
        driver.set_window_size(1400, 950)
        return driver
    except Exception:
        pass

    # 연결 실패 → 바로가기를 이미 패치했으므로 재시작 안내
    raise RuntimeError(
        "Chrome에 연결할 수 없습니다.\n\n"
        "바탕화면 Chrome 바로가기가 자동으로 업데이트되었습니다.\n"
        "Chrome을 완전히 닫은 후 바탕화면 바로가기로 다시 열고\n"
        "결재상신을 다시 시도해 주세요."
    )


# ──────────────────────────────────────────────────────────────────────────────
# 로그인
# ──────────────────────────────────────────────────────────────────────────────
def _is_login_page(driver) -> bool:
    try:
        url = driver.current_url.lower()
        if any(kw in url for kw in ("login", "signin", "auth", "sso")):
            return True
        src = driver.page_source.lower()
        return 'type="password"' in src or 'id="password"' in src
    except Exception:
        return False


def _do_login(
    driver,
    wait,
    username: str,
    password: str,
    progress_cb: Optional[Callable[[str], None]] = None,
) -> None:
    """로그인 페이지에서 계정/비밀번호를 자동 입력하고 로그인한다."""
    from selenium.webdriver.common.by import By
    from selenium.webdriver.support import expected_conditions as EC

    def _log(msg: str) -> None:
        logger.info(msg)
        if progress_cb:
            progress_cb(msg)

    _log("로그인 정보 입력 중…")

    # 계정 입력
    account_field = wait.until(EC.presence_of_element_located((By.ID, "username")))
    account_field.click()
    account_field.clear()
    account_field.send_keys(username)

    # 비밀번호 입력
    pw_field = driver.find_element(By.ID, "password")
    pw_field.click()
    pw_field.clear()
    pw_field.send_keys(password)

    # 로그인 버튼 클릭
    login_btn = wait.until(EC.element_to_be_clickable((By.ID, "login_submit")))
    driver.execute_script("arguments[0].click();", login_btn)
    _log("로그인 버튼 클릭…")

    # 로그인 완료 대기 (최대 30초)
    for _ in range(30):
        time.sleep(1)
        if not _is_login_page(driver):
            _log("로그인 완료!")
            time.sleep(1)
            return

    raise RuntimeError(
        "로그인 실패. 계정/비밀번호를 확인하세요.\n"
        "로그인 페이지에서 벗어나지 못했습니다."
    )


# ──────────────────────────────────────────────────────────────────────────────
# 메인 자동화 함수
# ──────────────────────────────────────────────────────────────────────────────
def run_approval(
    xlsx_path: str,
    title: str,
    pr_no: str,
    sales_person: str,
    line: str,
    process: str,
    equipment: str,
    equip_model: str,
    remark: str,
    po_no: str = "",
    username: str = "",
    password: str = "",
    progress_cb: Optional[Callable[[str], None]] = None,
) -> None:
    """
    전자결재 구매요청서(NPN) 결재상신 자동화.

    Parameters
    ----------
    username / password : 그룹웨어 로그인 계정 (없으면 수동 로그인 대기)
    """
    from selenium.webdriver.common.by import By
    from selenium.webdriver.support.ui import WebDriverWait
    from selenium.webdriver.support import expected_conditions as EC

    def _log(msg: str) -> None:
        logger.info(msg)
        if progress_cb:
            progress_cb(msg)

    def _fill_by_id(field_id: str, value: str) -> None:
        if not value:
            return
        try:
            el = driver.find_element(By.ID, field_id)
            driver.execute_script("arguments[0].scrollIntoView(true);", el)
            el.click()
            el.clear()
            el.send_keys(value)
        except Exception as e:
            logger.warning("필드 입력 실패 (id=%s): %s", field_id, e)

    driver = _create_driver()
    wait   = WebDriverWait(driver, 30)

    try:
        # ── STEP 1: 결재 홈 이동 ─────────────────────────────────────────────
        _log("① 전자결재 페이지 이동 중…")
        driver.get(APPROVAL_URL)
        wait.until(EC.presence_of_element_located((By.TAG_NAME, "body")))
        time.sleep(2)

        # ── 로그인 처리 ───────────────────────────────────────────────────────
        if _is_login_page(driver):
            if username and password:
                _do_login(driver, wait, username, password, progress_cb)
                if APPROVAL_URL not in driver.current_url:
                    driver.get(APPROVAL_URL)
                    wait.until(EC.presence_of_element_located((By.TAG_NAME, "body")))
                    time.sleep(2)
            else:
                _log("⚠️ 로그인이 필요합니다. Chrome에서 로그인해 주세요…")
                for _ in range(300):
                    time.sleep(1)
                    if not _is_login_page(driver):
                        _log("로그인 완료!")
                        break
                else:
                    raise TimeoutError("로그인 대기 시간 초과 (5분)")
                if APPROVAL_URL not in driver.current_url:
                    driver.get(APPROVAL_URL)
                    wait.until(EC.presence_of_element_located((By.TAG_NAME, "body")))
                    time.sleep(2)

        # ── STEP 2: 새 결재 진행 버튼 클릭 ──────────────────────────────────
        _log("② 새 결재 진행 버튼 클릭…")
        btn_new = wait.until(EC.element_to_be_clickable((By.ID, "writeBtn")))
        driver.execute_script("arguments[0].click();", btn_new)
        time.sleep(1)

        # ── STEP 3: 구매요청서(NPN) 직접 선택 ──────────────────────────────
        _log(f"③ '{FORM_NAME}' 선택…")
        # gpopupLayer 모달이 완전히 로드될 때까지 대기
        wait.until(EC.visibility_of_element_located((By.ID, "gpopupLayer")))
        time.sleep(1)

        # Full XPath로 직접 클릭 (id 방식 폴백 포함)
        FORM_FULL_XPATH = (
            "/html/body/div[3]/div/div[2]/div[1]/div/div[1]"
            "/ul/li[1]/ul/li[7]/ul/li[10]/a"
        )
        try:
            item_el = wait.until(EC.element_to_be_clickable((By.XPATH, FORM_FULL_XPATH)))
            driver.execute_script("arguments[0].scrollIntoView(true);", item_el)
            driver.execute_script("arguments[0].click();", item_el)
        except Exception:
            # 폴백: id 속성으로 재시도
            item_el = wait.until(EC.element_to_be_clickable((By.ID, "FORM_8637")))
            driver.execute_script("arguments[0].scrollIntoView(true);", item_el)
            driver.execute_script("arguments[0].click();", item_el)
        time.sleep(0.5)

        # ── STEP 4: 확인 버튼 클릭 ──────────────────────────────────────────
        _log("④ 확인 버튼 클릭…")
        confirm_btn = wait.until(EC.element_to_be_clickable(
            (By.XPATH, "//*[@id='gpopupLayer']/footer/a[1]")
        ))
        driver.execute_script("arguments[0].click();", confirm_btn)

        # ── STEP 5: 팝업 창 전환 ────────────────────────────────────────────
        _log("⑤ 팝업 창 대기 중…")
        main_window = driver.current_window_handle
        wait.until(lambda d: len(d.window_handles) > 1)
        popup_handle = next(
            h for h in driver.window_handles if h != main_window
        )
        driver.switch_to.window(popup_handle)
        wait.until(EC.presence_of_element_located((By.TAG_NAME, "body")))
        time.sleep(2)

        # ── STEP 6: 제목 입력 ────────────────────────────────────────────────
        _log("⑥ 제목 입력…")
        title_input = wait.until(EC.presence_of_element_located((By.ID, "subject")))
        title_input.click()
        title_input.clear()
        title_input.send_keys(title)

        # ── STEP 7: 구매요청 양식 데이터 입력 ───────────────────────────────
        _log("⑦ 구매요청 양식 데이터 입력…")
        _fill_by_id("addContentTable_1_3",  pr_no)
        _fill_by_id("addContentTable_1_4",  po_no)
        _fill_by_id("addContentTable_1_5",  sales_person)
        _fill_by_id("addContentTable_1_6",  line)
        _fill_by_id("addContentTable_1_7",  process)
        _fill_by_id("addContentTable_1_8",  equipment)
        _fill_by_id("addContentTable_1_9",  equip_model)
        _fill_by_id("addContentTable_1_10", remark)

        # ── STEP 8: 파일 첨부 ────────────────────────────────────────────────
        _log("⑧ 파일 첨부 중…")
        try:
            drop_zone = driver.find_element(By.ID, "dropZone")
            file_input = drop_zone.find_element(By.XPATH, ".//input[@type='file']")
            driver.execute_script(
                "arguments[0].style.display='block';"
                "arguments[0].style.visibility='visible';"
                "arguments[0].style.opacity='1';",
                file_input,
            )
            file_input.send_keys(xlsx_path)
            time.sleep(2)
            _log("   파일 첨부 완료.")
        except Exception as e:
            logger.warning("파일 첨부 중 오류: %s", e)

        _log("⑨ 구매요청서 작성 완료. 검토 후 결재 하세요.")

    except Exception as e:
        import traceback
        # 오류 발생 시에만 브라우저 종료
        try:
            driver.quit()
        except Exception:
            pass
        raise RuntimeError(
            f"결재상신 자동화 중 오류:\n{e}\n\n{traceback.format_exc()}"
        ) from e
    # finally 에서 driver.quit() 제거 → 브라우저를 열어둬서 사용자가 직접 결재요청


# ──────────────────────────────────────────────────────────────────────────────
# Excel 데이터 읽기 (openpyxl)
# ──────────────────────────────────────────────────────────────────────────────
def read_rack_row2(xlsx_path: str) -> Dict[str, str]:
    """
    RACK발주양식 시트 2행에서 결재상신에 필요한 컬럼 값을 반환.
    반환 키: pr_no, sales_person, line, process, equipment, equip_model, remark
    """
    from openpyxl import load_workbook
    wb = load_workbook(xlsx_path, data_only=True)
    ws = wb["RACK발주양식"] if "RACK발주양식" in wb.sheetnames else wb.active
    row2 = ws[2]

    def _cell(col_idx: int) -> str:
        v = row2[col_idx - 1].value if col_idx <= len(row2) else None
        return str(v).strip() if v is not None else ""

    # Z열(26번째) 전체에서 비어있지 않은 셀 수 - 1 (헤더 제외)
    z_count = sum(
        1 for r in ws.iter_rows(min_col=26, max_col=26, values_only=True)
        if r[0] is not None and str(r[0]).strip() != ""
    ) - 1

    result = {
        "pr_no":        _cell(7),    # G열
        "sales_person": _cell(9),    # I열
        "line":         _cell(23),   # W열
        "process":      _cell(21),   # U열
        "equipment":    _cell(24),   # X열
        "equip_model":  _cell(25),   # Y열
        "remark":       f"세부List 유첨 (총 {max(z_count, 0)}건)",
    }
    wb.close()
    return result
