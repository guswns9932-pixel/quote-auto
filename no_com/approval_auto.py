"""
approval_auto.py
================
Selenium 기반 전자결재 자동화 모듈.
구매요청서(NPN) 결재상신 워크플로를 자동으로 진행한다.

연결 방식
---------
현재 실행 중인 Chrome에 Remote Debugging Protocol(CDP)로 연결한다.
Chrome이 디버그 포트 없이 실행 중이면 종료 후 재시작(실제 프로필 유지).
기존 Chrome 창/탭은 그대로 유지하고 새 탭만 열어 작업한다.

요구 패키지:
    pip install selenium webdriver-manager
"""
from __future__ import annotations

import os
import sys
import time
import socket
import logging
import subprocess
from typing import Callable, Dict, Optional, Set

logger = logging.getLogger("QuoteApp.approval")

APPROVAL_URL = "https://groupware.lotvacuum.com/app/approval"
FORM_NAME    = "구매요청서(NPN)"
DEBUG_PORT   = 9222


# ──────────────────────────────────────────────────────────────────────────────
# Chrome 프로세스 / 디버그 포트 유틸
# ──────────────────────────────────────────────────────────────────────────────
def is_debug_port_open() -> bool:
    """Chrome 디버그 포트(9222)가 열려있으면 True."""
    try:
        s = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
        s.settimeout(1)
        result = s.connect_ex(("127.0.0.1", DEBUG_PORT))
        s.close()
        return result == 0
    except Exception:
        return False


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


def _real_user_data_dir() -> str:
    local_app = os.environ.get("LOCALAPPDATA", "")
    return os.path.join(local_app, "Google", "Chrome", "User Data")


def kill_chrome() -> None:
    """실행 중인 모든 Chrome 프로세스를 강제 종료한다."""
    try:
        subprocess.run(
            ["taskkill", "/F", "/IM", "chrome.exe"],
            capture_output=True, timeout=15,
        )
        time.sleep(2)
    except Exception as e:
        logger.warning("Chrome 종료 실패: %s", e)


def launch_chrome_with_debug() -> bool:
    """
    실제 Chrome 프로필로 디버그 포트를 열어 Chrome을 실행한다.
    Chrome이 이미 실행 중이면 kill_chrome()을 먼저 호출해야 한다.
    성공하면 True.
    """
    chrome_bin = _find_chrome_binary()
    if not chrome_bin:
        logger.error("Chrome 실행 파일을 찾을 수 없습니다.")
        return False

    user_data = _real_user_data_dir()
    args = [
        chrome_bin,
        f"--remote-debugging-port={DEBUG_PORT}",
        f"--user-data-dir={user_data}",
        "--profile-directory=Default",
        "--no-first-run",
        "--no-default-browser-check",
        "--disable-extensions",
    ]
    try:
        flags = 0x00000008 if sys.platform == "win32" else 0   # DETACHED_PROCESS
        subprocess.Popen(args, creationflags=flags)
        logger.info("Chrome 디버그 모드로 실행: port=%d", DEBUG_PORT)
        return True
    except Exception as e:
        logger.error("Chrome 실행 실패: %s", e)
        return False


def ensure_chrome_debug(progress_cb: Optional[Callable[[str], None]] = None) -> None:
    """
    Chrome이 디버그 포트로 실행 중인지 확인한다.
    아니면 kill → relaunch 후 포트가 열릴 때까지 최대 30초 대기.
    """
    def _log(msg: str) -> None:
        logger.info(msg)
        if progress_cb:
            progress_cb(msg)

    if is_debug_port_open():
        _log("기존 Chrome 디버그 포트에 연결합니다…")
        return

    _log("Chrome을 디버그 모드로 재시작합니다…")
    kill_chrome()

    if not launch_chrome_with_debug():
        raise RuntimeError(
            "Chrome을 찾을 수 없습니다. Google Chrome이 설치되어 있는지 확인하세요."
        )

    for _ in range(30):
        time.sleep(1)
        if is_debug_port_open():
            _log("Chrome 디버그 포트 연결 완료.")
            time.sleep(1)  # 추가 안정화
            return

    raise RuntimeError(
        "Chrome 디버그 포트 대기 시간 초과(30초). "
        "Chrome을 수동으로 종료 후 다시 시도하세요."
    )


# ──────────────────────────────────────────────────────────────────────────────
# WebDriver 생성 (기존 Chrome에 연결)
# ──────────────────────────────────────────────────────────────────────────────
def _create_driver(progress_cb: Optional[Callable[[str], None]] = None):
    """
    현재 실행 중인 Chrome에 CDP로 연결한 WebDriver를 반환한다.
    Chrome이 디버그 포트 없이 실행 중이면 재시작한다.
    """
    try:
        from selenium import webdriver
        from selenium.webdriver.chrome.options import Options
        from selenium.webdriver.chrome.service import Service
    except ImportError:
        raise RuntimeError(
            "selenium 패키지가 없습니다.\n"
            "pip install selenium webdriver-manager"
        )

    ensure_chrome_debug(progress_cb)

    options = Options()
    options.add_experimental_option("debuggerAddress", f"127.0.0.1:{DEBUG_PORT}")

    try:
        from webdriver_manager.chrome import ChromeDriverManager
        service = Service(ChromeDriverManager().install())
    except Exception:
        service = Service()

    driver = webdriver.Chrome(service=service, options=options)
    return driver


# ──────────────────────────────────────────────────────────────────────────────
# 로그인 감지 및 대기
# ──────────────────────────────────────────────────────────────────────────────
def _wait_for_login(
    driver,
    progress_cb: Optional[Callable[[str], None]] = None,
    timeout: int = 300,
) -> None:
    """
    현재 탭이 로그인 페이지면 사용자가 로그인할 때까지 최대 timeout초 대기.
    실제 Chrome 프로필 사용 시 이미 로그인되어 있으면 즉시 반환한다.
    """
    def _log(msg: str) -> None:
        logger.info(msg)
        if progress_cb:
            progress_cb(msg)

    def _is_login_page() -> bool:
        try:
            url = driver.current_url.lower()
            if any(kw in url for kw in ("login", "signin", "auth", "sso", "account")):
                return True
            src = driver.page_source.lower()
            return 'type="password"' in src or 'id="password"' in src
        except Exception:
            return False

    if not _is_login_page():
        return

    _log("⚠️ 로그인이 필요합니다. Chrome 브라우저에서 로그인해 주세요… (최대 5분 대기)")

    for _ in range(timeout):
        time.sleep(1)
        if not _is_login_page():
            _log("로그인 완료! 계속 진행합니다…")
            return

    raise TimeoutError(
        f"로그인 대기 시간 초과 ({timeout // 60}분). "
        "Chrome에서 로그인 후 다시 시도하세요."
    )


# ──────────────────────────────────────────────────────────────────────────────
# 테이블 입력 헬퍼
# ──────────────────────────────────────────────────────────────────────────────
def _find_col_index(driver, table_elem, header_text: str) -> int:
    from selenium.webdriver.common.by import By
    headers = table_elem.find_elements(By.XPATH, ".//th | .//thead//td")
    for i, th in enumerate(headers):
        if header_text.strip() in th.text.strip():
            return i
    return -1


def _fill_table_cell(driver, table_elem, header_text: str, value: str) -> bool:
    from selenium.webdriver.common.by import By

    col_idx = _find_col_index(driver, table_elem, header_text)
    if col_idx < 0:
        logger.warning("열 헤더를 찾지 못했습니다: %s", header_text)
        return False

    try:
        data_rows = table_elem.find_elements(By.XPATH, ".//tbody/tr")
        if not data_rows:
            data_rows = table_elem.find_elements(
                By.XPATH, ".//tr[not(ancestor::thead) and position()>1]"
            )
        if not data_rows:
            return False

        cells = data_rows[0].find_elements(By.XPATH, "./td")
        if col_idx >= len(cells):
            return False

        inp_list = cells[col_idx].find_elements(By.XPATH, ".//textarea | .//input")
        if not inp_list:
            return False

        inp = inp_list[0]
        driver.execute_script("arguments[0].scrollIntoView(true);", inp)
        inp.click()
        inp.clear()
        inp.send_keys(value)
        return True
    except Exception as e:
        logger.warning("셀 입력 실패 (%s): %s", header_text, e)
        return False


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
    progress_cb: Optional[Callable[[str], None]] = None,
) -> None:
    """
    전자결재 구매요청서(NPN) 결재상신 자동화.

    기존 Chrome에 새 탭을 열어 작업한다.
    완료 후 자동화 탭만 닫고 Chrome은 그대로 유지한다.
    """
    from selenium.webdriver.common.by import By
    from selenium.webdriver.support.ui import WebDriverWait
    from selenium.webdriver.support import expected_conditions as EC

    def _log(msg: str) -> None:
        logger.info(msg)
        if progress_cb:
            progress_cb(msg)

    driver = _create_driver(progress_cb)
    wait   = WebDriverWait(driver, 30)

    # 자동화 전 존재하던 탭 핸들 기록 (작업 후 해당 탭만 닫기 위함)
    initial_handles: Set[str] = set(driver.window_handles)

    # 새 탭 열기
    driver.execute_script("window.open('about:blank', '_blank');")
    time.sleep(0.5)
    new_tab = next(h for h in driver.window_handles if h not in initial_handles)
    driver.switch_to.window(new_tab)

    auto_handles: Set[str] = {new_tab}   # 자동화 중 열린 탭 추적

    try:
        # ── STEP 1: 결재 홈 이동 ─────────────────────────────────────────────
        _log("① 전자결재 페이지 이동 중…")
        driver.get(APPROVAL_URL)
        wait.until(EC.presence_of_element_located((By.TAG_NAME, "body")))
        time.sleep(2)

        # 로그인 여부 확인 — 실제 프로필이므로 대부분 이미 로그인 상태
        _wait_for_login(driver, progress_cb=progress_cb)

        # 로그인 후 결재 페이지가 아니면 재이동
        if APPROVAL_URL not in driver.current_url:
            driver.get(APPROVAL_URL)
            wait.until(EC.presence_of_element_located((By.TAG_NAME, "body")))
            time.sleep(2)

        # ── STEP 2: 새 결재 진행 버튼 클릭 ──────────────────────────────────
        _log("② 새 결재 진행 버튼 클릭…")
        btn_new = wait.until(EC.element_to_be_clickable(
            (By.XPATH,
             "//*[contains(text(),'새 결재 진행') or contains(text(),'새결재진행')]"
             "[self::button or self::a or self::span or self::div]")
        ))
        btn_new.click()
        time.sleep(1)

        # ── STEP 3: 결재양식 선택 모달 - 검색 ───────────────────────────────
        _log("③ 결재양식 검색 중…")
        search_input = wait.until(EC.visibility_of_element_located(
            (By.XPATH,
             "//input[contains(@placeholder,'검색') or contains(@placeholder,'제목')]"
             "[ancestor::*[contains(@class,'modal') or contains(@class,'dialog')"
             " or contains(@class,'popup') or contains(@class,'layer')]]"
             " | //input[contains(@placeholder,'검색') or contains(@placeholder,'제목')]")
        ))
        search_input.click()
        search_input.clear()
        search_input.send_keys(FORM_NAME)
        time.sleep(0.5)

        try:
            search_btn = driver.find_element(
                By.XPATH,
                "//button[.//i[contains(@class,'search') or contains(@class,'magnif')]"
                " or contains(@class,'search') or contains(@class,'btn-search')]"
                "[not(ancestor::input)]"
            )
            search_btn.click()
        except Exception:
            from selenium.webdriver.common.keys import Keys
            search_input.send_keys(Keys.RETURN)

        time.sleep(1.5)

        # ── STEP 4: 구매요청서(NPN) 항목 클릭 ──────────────────────────────
        _log(f"④ '{FORM_NAME}' 항목 선택…")
        item_el = wait.until(EC.element_to_be_clickable(
            (By.XPATH,
             f"//*[normalize-space(text())='{FORM_NAME}'"
             f" or contains(text(),'{FORM_NAME}')]"
             f"[self::span or self::li or self::td or self::a or self::div]")
        ))
        item_el.click()
        time.sleep(0.5)

        # ── STEP 5: 확인 버튼 클릭 ──────────────────────────────────────────
        _log("⑤ 확인 버튼 클릭…")
        confirm_btn = wait.until(EC.element_to_be_clickable(
            (By.XPATH,
             "//button[normalize-space(text())='확인' or normalize-space(text())='확 인']"
             "[not(contains(@class,'cancel')) and not(contains(@class,'close'))]")
        ))
        confirm_btn.click()

        # ── STEP 6: 팝업 창 전환 ────────────────────────────────────────────
        _log("⑥ 팝업 창 대기 중…")
        wait.until(lambda d: len(d.window_handles) > len(auto_handles) + len(initial_handles) - 1)
        popup_handle = next(
            h for h in driver.window_handles
            if h not in initial_handles and h != new_tab
        )
        auto_handles.add(popup_handle)
        driver.switch_to.window(popup_handle)
        wait.until(EC.presence_of_element_located((By.TAG_NAME, "body")))
        time.sleep(2)

        # ── STEP 7: 제목 입력 ────────────────────────────────────────────────
        _log("⑦ 제목 입력…")
        title_input = wait.until(EC.presence_of_element_located(
            (By.XPATH,
             "//tr[.//td[normalize-space(text())='제목' or normalize-space(text())='제 목']]//input"
             " | //td[normalize-space(text())='제목' or normalize-space(text())='제 목']"
             "/following-sibling::td[1]//input"
             " | //label[contains(text(),'제목')]/following::input[1]")
        ))
        driver.execute_script("arguments[0].scrollIntoView(true);", title_input)
        title_input.click()
        title_input.clear()
        title_input.send_keys(title)

        # ── STEP 8: NPN 신규투자 테이블 필드 입력 ───────────────────────────
        _log("⑧ 구매요청 양식 데이터 입력…")
        tables = driver.find_elements(By.XPATH, "//table")
        npn_table = None
        for tbl in tables:
            if "PR NO" in tbl.text and "설비" in tbl.text:
                npn_table = tbl
                break

        if npn_table is None:
            logger.warning("NPN 테이블을 찾지 못했습니다. 텍스트 기반 폴백 사용.")

        def _fill(header: str, value: str) -> None:
            if not value:
                return
            if npn_table is not None and _fill_table_cell(driver, npn_table, header, value):
                return
            try:
                xp = (f"//td[contains(text(),'{header}')]/following-sibling::td[1]//input"
                      f" | //th[contains(text(),'{header}')]/following-sibling::th[1]//input"
                      f" | //td[contains(text(),'{header}')]/following::input[1]")
                inp = driver.find_element(By.XPATH, xp)
                inp.click(); inp.clear(); inp.send_keys(value)
            except Exception as e:
                logger.warning("필드 '%s' 입력 실패: %s", header, e)

        _fill("PR NO",     pr_no)
        _fill("영업",       sales_person)
        _fill("라인",       line)
        _fill("공정",       process)
        _fill("설비",       equipment)
        _fill("설비MODEL",  equip_model)
        _fill("비 고",      remark)

        # ── STEP 9: 파일 첨부 ────────────────────────────────────────────────
        _log("⑨ 파일 첨부 중…")
        try:
            file_inputs = driver.find_elements(By.XPATH, "//input[@type='file']")
            if file_inputs:
                inp_file = file_inputs[0]
                driver.execute_script(
                    "arguments[0].style.display='block';"
                    "arguments[0].style.visibility='visible';",
                    inp_file,
                )
                inp_file.send_keys(xlsx_path)
                time.sleep(2)
                _log("   파일 첨부 완료.")
            else:
                logger.warning("파일 input 요소를 찾지 못했습니다.")
        except Exception as e:
            logger.warning("파일 첨부 중 오류: %s", e)

        # ── STEP 10: 결재요청 버튼 클릭 ─────────────────────────────────────
        _log("⑩ 결재요청 버튼 클릭…")
        submit_btn = wait.until(EC.element_to_be_clickable(
            (By.XPATH,
             "//button[contains(text(),'결재요청') or contains(text(),'결 재요청')]"
             "[not(contains(@class,'cancel'))]")
        ))
        driver.execute_script("arguments[0].scrollIntoView(true);", submit_btn)
        submit_btn.click()
        time.sleep(2)

        # ── STEP 11: 확인 팝업(alert 또는 새 창) 처리 ─────────────────────
        _log("⑪ 결재 확인 처리…")
        try:
            driver.switch_to.alert.accept()
            _log("   alert 확인.")
            time.sleep(1)
        except Exception:
            pass

        extra_handles = [
            h for h in driver.window_handles
            if h not in initial_handles and h not in auto_handles
        ]
        if extra_handles:
            auto_handles.add(extra_handles[0])
            driver.switch_to.window(extra_handles[0])
            wait.until(EC.presence_of_element_located((By.TAG_NAME, "body")))
            time.sleep(1)
            try:
                confirm2 = wait.until(EC.element_to_be_clickable(
                    (By.XPATH,
                     "//button[contains(text(),'결재요청') or contains(text(),'확인')]")
                ))
                confirm2.click()
                _log("   추가 확인 창 처리 완료.")
                time.sleep(1.5)
            except Exception as e:
                logger.warning("추가 확인 창 처리 실패: %s", e)

        _log("✅ 결재상신 완료!")

    except Exception as e:
        import traceback
        raise RuntimeError(
            f"결재상신 자동화 중 오류:\n{e}\n\n{traceback.format_exc()}"
        ) from e
    finally:
        time.sleep(2)
        # 자동화 탭만 닫기 — 기존 Chrome 창/탭은 유지
        for h in list(auto_handles):
            try:
                if h in driver.window_handles:
                    driver.switch_to.window(h)
                    driver.close()
            except Exception:
                pass
        # 원래 탭으로 포커스 복귀 (있는 경우)
        try:
            remaining = [h for h in driver.window_handles if h in initial_handles]
            if remaining:
                driver.switch_to.window(remaining[0])
        except Exception:
            pass


# ──────────────────────────────────────────────────────────────────────────────
# Excel 데이터 읽기 (openpyxl)
# ──────────────────────────────────────────────────────────────────────────────
def read_rack_row2(xlsx_path: str) -> Dict[str, str]:
    """
    RACK발주양식 시트 2행에서 결재상신에 필요한 컬럼 값을 반환.
    반환 키: pr_no, sales_person, process, line, equipment, equip_model
    """
    from openpyxl import load_workbook
    wb = load_workbook(xlsx_path, data_only=True)
    ws = wb["RACK발주양식"] if "RACK발주양식" in wb.sheetnames else wb.active
    row = ws[2]

    def _cell(col_idx: int) -> str:
        v = row[col_idx - 1].value if col_idx <= len(row) else None
        return str(v).strip() if v is not None else ""

    result = {
        "pr_no":        _cell(9),    # I열
        "sales_person": _cell(11),   # K열
        "process":      _cell(23),   # W열
        "line":         _cell(25),   # Y열
        "equipment":    _cell(26),   # Z열
        "equip_model":  _cell(27),   # AA열
    }
    wb.close()
    return result
