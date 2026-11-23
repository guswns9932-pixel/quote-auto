# _archive — 지금은 쓰지 않는 파일 보관

어떤 앱도 불러오지 않는 파일을 모아 둔 곳입니다. 빌드에도 들어가지 않습니다.
다시 필요해지면 `no_com/` 으로 되돌리면 됩니다.

| 파일 | 내용 | 안 쓰는 이유 |
|---|---|---|
| `approval_auto.py` | 결재상신 자동화 (Selenium) | 화면 어디에서도 호출하지 않음 |
| `pages.py` | 옛 이름 호환용 재내보내기 | `page_quote` / `page_esign` 을 직접 불러오므로 불필요 |
| `test_printout.py` | 인쇄영역 캡처 수동 시험 스크립트 | 특정 PC 경로가 박힌 일회용 스크립트 |
| `esign_capture.py` | 전자서명: 시트를 클립보드로 캡처해 그림으로 받던 방식 | PDF 저장 방식(`esign_engine.py`)으로 교체 |
| `content_search/`, `run_search.py` | 키워드 검색기 앱 | 사용 중지 — 런처·빌드에서 뺌 |
| `csp_order_maker.py`, `csp_core.py`, `run_csp.py`, `test_csp.py` | CSP UPLOAD 자동화 앱 (+ 테스트) | 사용 중지 — 런처·빌드에서 뺌 |

### 키워드 검색기·CSP UPLOAD 되살리기

1. 소스를 `no_com/` 으로, `test_csp.py` 를 `tests/` 로 되돌린다.
2. `no_com/main.py` 의 `_APPS`·런처 버튼 목록·창 높이(앱당 +78px)·`_APP_TAB_LABELS` 에 다시 넣는다.
3. `no_com/app_main.py` 의 `_MODES` 에 `--search`·`--csp` 를 넣는다.
4. `no_com/build_onedir.py` 의 `APP_KEYS`·`_PROFILES` 에 search·csp 프로필을 되살리고,
   합본 빌드 제외 목록에서 `_NO_TK` 를 뺀다(두 앱은 Tk 앱).
5. `tests/conftest.py` 에 `csp_template` 픽스처를, `tests/test_build_spec.py` 의 앱 목록을 되돌린다.

이 작업 직전 커밋(git log 로 "키워드 검색기·CSP UPLOAD 사용 중지" 의 부모)을 보면 원래 코드가 그대로 있다.
