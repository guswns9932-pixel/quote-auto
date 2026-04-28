# 견적/전자서명 통합 시스템

## 실행
```
pip install PySide6 openpyxl PyMuPDF pywin32
python main.py
```

## 파일 구조 (총 2,869줄 / 20클래스 / 169메서드)

| 파일 | 줄수 | 역할 | 의존 |
|------|------|------|------|
| `core.py`     |  439 | 데이터 모델·DB·유틸 | 없음 (단독 테스트 가능) |
| `excel_io.py` |  646 | Excel 파싱(openpyxl) + COM 생성 | core |
| `widgets.py`  |  308 | 커스텀 Qt 위젯 | PySide6 |
| `pages.py`    | 1331 | 3개 페이지 + 옵션 다이얼로그 | core, excel_io, widgets |
| `main.py`     |  145 | 진입점·MainWindow·LeftNav | pages, core |

## 의존 방향 (단방향)
```
main.py → pages.py → excel_io.py → core.py
                   └─ widgets.py ──→ core.py
```

## 기능
- **견적서작성**: STEP1 통합양식 LOAD → STEP2 의뢰파일 LOAD → STEP3 조건입력 →
  STEP4 품목 드래그/더블클릭 → STEP5 견적 테이블(Credit/TOTAL) → STEP6 국내/중국/미국 생성 → STEP7 결과목록
- **전자서명**: 승인코드+서명이미지 LOAD → 엑셀 LOAD(PDF변환) → PDF뷰어 더블클릭 서명삽입 → PDF저장
- **견적LOG**: 검색/필터 → 체크 선택 → 품목 합산 → Excel Export
