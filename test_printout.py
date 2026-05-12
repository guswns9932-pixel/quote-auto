import win32com.client
import win32print
import os
import traceback

XLSX = r"N:\01.기본문서함\사업1팀.영업공용\☆ 공유 폴더\99. 내부자료\07. 업무자동화\01. 견적작성자동화\견적서\260430_LOT베큠_P34D_ETCH_최동섭\7269457619-20_260430_LOT베큠_P34D_ETCH_최동섭_ELP3655.xlsx"
OUT  = os.path.join(os.path.dirname(os.path.abspath(__file__)), "test_printout_out.pdf")

print("=== 설치된 프린터 목록 ===")
printers = win32print.EnumPrinters(
    win32print.PRINTER_ENUM_LOCAL | win32print.PRINTER_ENUM_CONNECTIONS)
pdf_name = None
for flags, desc, name, comment in printers:
    print(f"  {name}")
    if "pdf" in name.lower():
        pdf_name = name

print()
if pdf_name:
    print(f"PDF 프린터 발견: {pdf_name}")
else:
    print("PDF 프린터를 찾지 못했습니다.")
    input("\nEnter 키를 누르면 창이 닫힙니다...")
    raise SystemExit

print(f"\n출력: {OUT}")

try:
    xl = win32com.client.Dispatch("Excel.Application")
    xl.Visible = False
    xl.DisplayAlerts = False

    # Excel 에서 인식하는 프린터 이름(포트 포함) 확인
    orig_printer = xl.ActivePrinter
    print(f"현재 Excel 기본 프린터: {orig_printer!r}")

    # Excel ActivePrinter 형식은 "프린터이름 on 포트:" 이므로
    # 실제 포트명을 얻기 위해 win32print 사용
    info = win32print.OpenPrinter(pdf_name)
    pinfo = win32print.GetPrinter(info, 2)
    port = pinfo["pPortName"]
    win32print.ClosePrinter(info)
    xl_printer = f"{pdf_name} on {port}:"
    print(f"Excel 프린터 문자열: {xl_printer!r}")

    try:
        wb = xl.Workbooks.Open(XLSX, ReadOnly=True, UpdateLinks=0, AddToMru=False)
        try:
            ws = wb.Worksheets(1)
            print(f"시트명: {ws.Name}")

            xl.ActivePrinter = xl_printer
            ws.PrintOut(
                Copies=1,
                PrintToFile=True,
                PrToFileName=OUT,
            )

            if os.path.exists(OUT):
                print(f"\n성공: PDF 생성됨 ({os.path.getsize(OUT):,} bytes)")
            else:
                print("\n실패: PDF 파일이 생성되지 않음")
        finally:
            wb.Close(False)
    finally:
        xl.ActivePrinter = orig_printer
        xl.Quit()

except Exception:
    print("\n=== 오류 발생 ===")
    traceback.print_exc()

print("\n완료 — 보안 소프트웨어 경고 여부를 확인하세요.")
input("\nEnter 키를 누르면 창이 닫힙니다...")
