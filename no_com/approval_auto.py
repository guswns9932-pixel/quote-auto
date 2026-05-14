"""
approval_auto.py
================
Selenium 기반 전자결재 자동화 모듈.
구매요청서(NPN) 결재상신 워크플로를 자동으로 진행한다.

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
def _exe_dir() -> str:
    if getattr(sys, "frozen", False):
        return os.path.dirname(sys.executable)
    return os.path.dirname(os.path.abspath(__file__))


def _profile_dir() -> str:
    """앱 전용 Chrome 프로필 경로.
    AppData\\Roaming 아래에 생성 — exe 디렉터리와 달리 항상 쓰기 권한이 있음."""
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


def _create_driver():
    """Chrome WebDriver 생성. 앱 전용 프로필로 세션 유지."""
    try:
        from selenium import webdriver
        from selenium.webdriver.chrome.options import Options
        from selenium.webdriver.chrome.service import Service
    except ImportError:
        raise RuntimeError(
            "selenium 패키지가 없습니다.\n"
            "pip install selenium webdriver-manager"
        )

    # chromedriver 확보
    try:
        from webdriver_manager.chrome import ChromeDriverManager
        service = Service(ChromeDriverManager().install())
    except Exception:
        service = Service()   # PATH에 있는 chromedriver 사용

    profile = _profile_dir()
    os.makedirs(profile, exist_ok=True)

    options = Options()
    options.add_argument(f"--user-data-dir={profile}")
    # --profile-directory 는 제거: Default 이외 이름의 서브디렉터리가 있으면
    # prefs 쓰기 실패가 발생하므로 Chrome이 알아서 선택하도록 함
    options.add_argument("--no-first-run")
    options.add_argument("--no-default-browser-check")
    options.add_argument("--disable-notifications")
    options.add_argument("--no-sandbox")               # 쓰기 권한 문제 완화
    options.add_argument("--disable-dev-shm-usage")    # 공유 메모리 오류 방지
    options.add_argument("--disable-gpu")
    options.add_argument("--disable-software-rasterizer")
    options.add_argument("--remote-allow-origins=*")   # session 생성 오류 방지
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
# 테이블 열 인덱스 기반 입력 헬퍼
# ──────────────────────────────────────────────────────────────────────────────
def _find_col_index(driver, table_elem, header_text: str) -> int:
    """table 요소에서 header_text를 포함하는 헤더 셀의 0-based 열 인덱스 반환. 없으면 -1."""
    from selenium.webdriver.common.by import By
    headers = table_elem.find_elements(By.XPATH, ".//th | .//thead//td")
    for i, th in enumerate(headers):
        if header_text.strip() in th.text.strip():
            return i
    return -1


def _fill_table_cell(driver, table_elem, header_text: str, value: str) -> bool:
    """
    테이블에서 header_text 열의 첫 번째 데이터 행 입력 필드에 value를 입력.
    성공하면 True.
    """
    from selenium.webdriver.common.by import By

    col_idx = _find_col_index(driver, table_elem, header_text)
    if col_idx < 0:
        logger.warning("열 헤더를 찾지 못했습니다: %s", header_text)
        return False

    try:
        # thead를 제외한 tbody의 첫 번째 tr
        data_rows = table_elem.find_elements(By.XPATH, ".//tbody/tr")
        if not data_rows:
            data_rows = table_elem.find_elements(By.XPATH,
                ".//tr[not(ancestor::thead) and position()>1]")
        if not data_rows:
            return False
        row = data_rows[0]
        cells = row.find_elements(By.XPATH, "./td")
        if col_idx >= len(cells):
            return False

        cell = cells[col_idx]
        # textarea 우선, 없으면 input
        inputs = cell.find_elements(By.XPATH, ".//textarea | .//input")
        if not inputs:
            return False

        inp = inputs[0]
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

    Parameters
    ----------
    xlsx_path    : 첨부할 RACK 구매요청서 xlsx 경로
    title        : 제목 (파일명 기반)
    pr_no        : PR NO. (I열)
    sales_person : 영업 담당자 (K열)
    line         : 라인 (Y열)
    process      : 공정 (W열)
    equipment    : 설비 (Z열)
    equip_model  : 설비MODEL (AA열)
    remark       : 비고 (예: "세부List 유첨 (총 1건)")
    progress_cb  : 진행 상황 콜백 (str → None)
    """
    from selenium.webdriver.common.by import By
    from selenium.webdriver.support.ui import WebDriverWait
    from selenium.webdriver.support import expected_conditions as EC

    def _log(msg: str) -> None:
        logger.info(msg)
        if progress_cb:
            progress_cb(msg)

    driver = _create_driver()
    wait   = WebDriverWait(driver, 30)

    try:
        # ── STEP 1: 결재 홈 이동 ─────────────────────────────────────────────
        _log("① 전자결재 페이지 이동 중…")
        driver.get(APPROVAL_URL)
        wait.until(EC.presence_of_element_located((By.TAG_NAME, "body")))
        time.sleep(2)

        # ── STEP 2: 새 결재 진행 버튼 클릭 ──────────────────────────────────
        _log("② 새 결재 진행 버튼 클릭…")
        btn_new = wait.until(EC.element_to_be_clickable(
            (By.XPATH, "//*[contains(text(),'새 결재 진행') or contains(text(),'새결재진행')]"
             "[self::button or self::a or self::span or self::div]")
        ))
        btn_new.click()
        time.sleep(1)

        # ── STEP 3: 결재양식 선택 모달 - 검색 ───────────────────────────────
        _log("③ 결재양식 검색 중…")
        search_input = wait.until(EC.visibility_of_element_located(
            (By.XPATH,
             "//input[contains(@placeholder,'검색') or contains(@placeholder,'제목')]"
             "[ancestor::*[contains(@class,'modal') or contains(@class,'dialog') "
             " or contains(@class,'popup') or contains(@class,'layer')]]"
             " | //input[contains(@placeholder,'검색') or contains(@placeholder,'제목')]")
        ))
        search_input.click()
        search_input.clear()
        search_input.send_keys(FORM_NAME)
        time.sleep(0.5)

        # 검색 버튼 클릭 (돋보기 버튼)
        try:
            search_btn = driver.find_element(
                By.XPATH,
                "//button[.//i[contains(@class,'search') or contains(@class,'magnif')] "
                " or contains(@class,'search') or contains(@class,'btn-search')]"
                "[not(ancestor::input)]"
            )
            search_btn.click()
        except Exception:
            # 엔터로 대체
            from selenium.webdriver.common.keys import Keys
            search_input.send_keys(Keys.RETURN)

        time.sleep(1.5)

        # ── STEP 4: 구매요청서(NPN) 항목 클릭 ──────────────────────────────
        _log(f"④ '{FORM_NAME}' 항목 선택…")
        item_el = wait.until(EC.element_to_be_clickable(
            (By.XPATH,
             f"//*[normalize-space(text())='{FORM_NAME}' "
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
        main_window = driver.current_window_handle
        wait.until(lambda d: len(d.window_handles) > 1)
        popup_handle = next(
            h for h in driver.window_handles if h != main_window
        )
        driver.switch_to.window(popup_handle)
        wait.until(EC.presence_of_element_located((By.TAG_NAME, "body")))
        time.sleep(2)  # 폼 완전 로드 대기

        # ── STEP 7: 제목 입력 ────────────────────────────────────────────────
        _log("⑦ 제목 입력…")
        title_input = wait.until(EC.presence_of_element_located(
            (By.XPATH,
             # '제목' 레이블 바로 옆 또는 다음 입력 필드
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

        # NPN 신규투자 구매요청서 테이블을 찾음 (여러 테이블 중 헤더로 구분)
        tables = driver.find_elements(By.XPATH, "//table")
        npn_table = None
        for tbl in tables:
            headers_text = tbl.text
            if "PR NO" in headers_text and "설비" in headers_text:
                npn_table = tbl
                break

        if npn_table is None:
            logger.warning("NPN 테이블을 찾지 못했습니다. 텍스트 기반 폴백 사용.")

        def _fill(header: str, value: str):
            if not value:
                return
            if npn_table is not None:
                if _fill_table_cell(driver, npn_table, header, value):
                    return
            # 폴백: 전체 문서에서 label→input 탐색
            try:
                xp = (f"//td[contains(text(),'{header}')]/following-sibling::td[1]//input"
                      f" | //th[contains(text(),'{header}')]/following-sibling::th[1]//input"
                      f" | //td[contains(text(),'{header}')]/following::input[1]")
                inp = driver.find_element(By.XPATH, xp)
                inp.click(); inp.clear(); inp.send_keys(value)
            except Exception as e:
                logger.warning("필드 '%s' 입력 실패: %s", header, e)

        _fill("PR NO",      pr_no)
        _fill("영업",        sales_person)   # 영업 담당자
        _fill("라인",        line)
        _fill("공정",        process)
        _fill("설비",        equipment)
        _fill("설비MODEL",   equip_model)
        _fill("비 고",       remark)

        # ── STEP 9: 파일 첨부 ────────────────────────────────────────────────
        _log("⑨ 파일 첨부 중…")
        try:
            # type="file" 입력 요소 탐색 (숨겨진 경우도 포함)
            file_inputs = driver.find_elements(By.XPATH, "//input[@type='file']")
            if file_inputs:
                inp_file = file_inputs[0]
                # 숨겨진 파일 input을 보이게 해서 send_keys 허용
                driver.execute_script(
                    "arguments[0].style.display='block';"
                    "arguments[0].style.visibility='visible';",
                    inp_file
                )
                inp_file.send_keys(xlsx_path)
                time.sleep(2)  # 업로드 완료 대기
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
        # alert 처리
        try:
            alert = driver.switch_to.alert
            alert.accept()
            _log("   alert 확인.")
            time.sleep(1)
        except Exception:
            pass

        # 새 확인 창이 열린 경우
        all_handles = driver.window_handles
        new_handles = [h for h in all_handles
                       if h not in (main_window, popup_handle)]
        if new_handles:
            driver.switch_to.window(new_handles[0])
            wait.until(EC.presence_of_element_located((By.TAG_NAME, "body")))
            time.sleep(1)
            try:
                confirm2 = wait.until(EC.element_to_be_clickable(
                    (By.XPATH,
                     "//button[contains(text(),'결재요청') or contains(text(),'확인')]")
                ))
                confirm2.click()
                _log("   추가 확인 창에서 결재요청 클릭.")
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
    row = ws[2]  # 2번째 행

    def _cell(col_idx: int) -> str:
        """1-based 열 인덱스에서 문자열 값 반환."""
        v = row[col_idx - 1].value if col_idx <= len(row) else None
        return str(v).strip() if v is not None else ""

    result = {
        "pr_no":        _cell(9),   # I열
        "sales_person": _cell(11),  # K열
        "process":      _cell(23),  # W열
        "line":         _cell(25),  # Y열
        "equipment":    _cell(26),  # Z열
        "equip_model":  _cell(27),  # AA열
    }
    wb.close()
    return result
