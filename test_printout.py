import win32com.client
import os
import traceback

XLSX = r"N:\01.기본문서함\사업1팀.영업공용\☆ 공유 폴더\99. 내부자료\07. 업무자동화\01. 견적작성자동화\견적서\260430_LOT베큠_P34D_ETCH_최동섭\7269457619-20_260430_LOT베큠_P34D_ETCH_최동섭_ELP3655.xlsx"
OUT  = os.path.join(os.path.dirname(os.path.abspath(__file__)), "test_printout_out.pdf")

print("테스트 시작...")
print(f"입력: {XLSX}")
print(f"출력: {OUT}")

try:
    xl = win32com.client.Dispatch("Excel.Application")
    xl.Visible = False
    xl.DisplayAlerts = False

    try:
        wb = xl.Workbooks.Open(XLSX, ReadOnly=True, UpdateLinks=0, AddToMru=False)
        try:
            ws = wb.Worksheets(1)
            print(f"시트명: {ws.Name}")
            print(f"인쇄영역: {ws.PageSetup.PrintArea!r}")

            ws.PrintOut(
                Copies=1,
                PrintToFile=True,
                ActivePrinter="Microsoft Print to PDF",
                PrToFileName=OUT,
            )

            if os.path.exists(OUT):
                print(f"\n성공: PDF 생성됨 ({os.path.getsize(OUT):,} bytes)")
            else:
                print("\n실패: PDF 파일이 생성되지 않음")
        finally:
            wb.Close(False)
    finally:
        xl.Quit()

except Exception:
    print("\n=== 오류 발생 ===")
    traceback.print_exc()

print("\n완료 — 보안 소프트웨어 경고 여부를 확인하세요.")
input("\nEnter 키를 누르면 창이 닫힙니다...")
