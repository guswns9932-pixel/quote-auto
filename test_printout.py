import win32com.client
import os
import traceback

try:
    from PIL import ImageGrab
except ImportError:
    print("pillow 없음: pip install pillow")
    input("Enter...")
    raise SystemExit

XLSX = r"N:\01.기본문서함\사업1팀.영업공용\☆ 공유 폴더\99. 내부자료\07. 업무자동화\01. 견적작성자동화\견적서\260430_LOT베큠_P34D_ETCH_최동섭\7269457619-20_260430_LOT베큠_P34D_ETCH_최동섭_ELP3655.xlsx"
BASE = os.path.dirname(os.path.abspath(__file__))

try:
    xl = win32com.client.Dispatch("Excel.Application")
    xl.Visible = False
    xl.DisplayAlerts = False

    wb = xl.Workbooks.Open(XLSX, ReadOnly=True, UpdateLinks=0, AddToMru=False)
    try:
        ws = wb.Worksheets(1)
        print(f"시트명: {ws.Name}")

        # 인쇄 영역 사용 (없으면 UsedRange)
        pa = ws.PageSetup.PrintArea
        rng = ws.Range(pa) if pa else ws.UsedRange
        print(f"캡처 범위: {rng.Address}")

        for label, appearance in [("screen_quality", 1), ("printer_quality", 2)]:
            rng.CopyPicture(Appearance=appearance, Format=2)  # Format=2: xlBitmap
            img = ImageGrab.grabclipboard()
            if img is None:
                print(f"[{label}] 클립보드 캡처 실패")
                continue
            out = os.path.join(BASE, f"test_{label}.png")
            img.save(out, "PNG")
            print(f"[{label}] 저장: {out}  크기: {img.size}  ({os.path.getsize(out):,} bytes)")

    finally:
        wb.Close(False)
    xl.Quit()

except Exception:
    print("\n=== 오류 발생 ===")
    traceback.print_exc()

input("\nEnter 키를 누르면 창이 닫힙니다...")
