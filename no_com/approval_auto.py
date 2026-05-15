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
def _create_driver():
    """앱 전용 프로필로 Chrome WebDriver를 생성한다 (기존 Chrome과 독립)."""
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

    profile = _profile_dir()
    os.makedirs(profile, exist_ok=True)

    options = Options()
    options.add_argument(f"--user-data-dir={profile}")
    options.add_argument("--no-first-run")
    options.add_argument("--no-default-browser-check")
    options.add_argument("--disable-notifications")
    options.add_argument("--no-sandbox")
    options.add_argument("--disable-dev-shm-usage")
    options.add_argument("--disable-gpu")
    options.add_argument("--remote-allow-origins=*")
    options.add_experimental_option("excludeSwitches", ["enable-automation"])
    options.add_experimental_option("useAutomationExtension", False)

    chrome_bin = _find_chrome_binary()
    if chrome_bin:
        options.binary_location = chrome_bin

    from selenium import webdriver
    driver = webdriver.Chrome(service=service, options=options)
    driver.set_window_size(1400, 950)
    return driver


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
            import pyperclip
            import pyautogui

            # 파일 선택 버튼 클릭 → Windows 파일 열기 다이얼로그 오픈
            file_btn = driver.find_element(
                By.XPATH, '//*[@id="dropZone"]/div[1]/span[2]/span[1]'
            )
            driver.execute_script("arguments[0].scrollIntoView(true);", file_btn)
            driver.execute_script("arguments[0].click();", file_btn)
            time.sleep(2)  # 다이얼로그 로딩 대기

            # 경로를 클립보드에 복사 후 붙여넣기 (한글/특수문자 안전)
            pyperclip.copy(xlsx_path)
            pyautogui.hotkey("ctrl", "a")   # 파일이름 입력란 전체 선택
            pyautogui.hotkey("ctrl", "v")   # 경로 붙여넣기
            time.sleep(0.5)
            pyautogui.press("enter")        # 열기 버튼
            time.sleep(2)
            _log("   파일 첨부 완료.")
        except Exception as e:
            logger.warning("파일 첨부 중 오류: %s", e)

        # ── STEP 9: 결재요청 버튼 클릭 ─────────────────────────────────────
        _log("⑨ 결재요청 버튼 클릭…")
        submit_btn = wait.until(EC.element_to_be_clickable((By.ID, "act_draft")))
        driver.execute_script("arguments[0].scrollIntoView(true);", submit_btn)
        driver.execute_script("arguments[0].click();", submit_btn)
        time.sleep(2)

        # ── STEP 10: 확인 팝업(alert 또는 새 창) 처리 ─────────────────────
        _log("⑩ 결재 확인 처리…")
        try:
            driver.switch_to.alert.accept()
            _log("   alert 확인.")
            time.sleep(1)
        except Exception:
            pass

        extra_handles = [
            h for h in driver.window_handles
            if h not in (main_window, popup_handle)
        ]
        if extra_handles:
            driver.switch_to.window(extra_handles[0])
            wait.until(EC.presence_of_element_located((By.TAG_NAME, "body")))
            time.sleep(1)
            try:
                confirm2 = wait.until(EC.element_to_be_clickable(
                    (By.XPATH,
                     "//button[contains(text(),'결재요청') or contains(text(),'확인')]")
                ))
                driver.execute_script("arguments[0].click();", confirm2)
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
        try:
            driver.quit()
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
