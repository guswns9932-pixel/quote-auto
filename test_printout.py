import win32com.client
import os
import traceback

try:
    from PIL import ImageGrab
except ImportError:
    print("pillow 없음: pip install pillow")
    input("Enter..."); raise SystemExit

XLSX = r"N:\01.기본문서함\사업1팀.영업공용\☆ 공유 폴더\99. 내부자료\07. 업무자동화\01. 견적작성자동화\견적서\260430_LOT베큠_P34D_ETCH_최동섭\7269457619-20_260430_LOT베큠_P34D_ETCH_최동섭_ELP3655.xlsx"
BASE = os.path.dirname(os.path.abspath(__file__))

ESIGN_TARGET = ["견적서 갑지", "사양서", "현업 사인용 사양서", "입고검수확인서"]

# Visible 코드: -1=보임, 0=숨김, 2=매우숨김
VISIBLE_LABEL = {-1: "보임", 0: "숨김", 2: "매우숨김"}

try:
    xl = win32com.client.Dispatch("Excel.Application")
    xl.Visible = False
    xl.DisplayAlerts = False

    wb = xl.Workbooks.Open(XLSX, ReadOnly=True, UpdateLinks=0, AddToMru=False)
    try:
        print("=== 전체 시트 목록 ===")
        for ws in wb.Worksheets:
            v = int(ws.Visible)
            label = VISIBLE_LABEL.get(v, str(v))
            target = "★" if ws.Name in ESIGN_TARGET else " "
            print(f"  {target} [{label:6s}]  {ws.Name!r}")

        print("\n=== ESIGN_TARGET 캡처 시도 ===")
        for name in ESIGN_TARGET:
            try:
                ws = wb.Worksheets(name)
            except Exception:
                print(f"  {name!r}  → 시트 없음 (이름 불일치)")
                continue

            v = int(ws.Visible)
            if v != -1:
                print(f"  {name!r}  → 건너뜀 ({VISIBLE_LABEL.get(v, v)})")
                continue

            ws.Activate()
            pa_raw = ws.PageSetup.PrintArea
            pa = pa_raw
            if pa and "!" in pa:
                pa = pa.split("!")[-1]
            print(f"  {name!r}  PrintArea raw={pa_raw!r}  → 사용={pa!r}", flush=True)
            try:
                rng = ws.Range(pa) if pa else ws.UsedRange
            except Exception:
                rng = ws.UsedRange
            print(f"           캡처 범위: {rng.Address}", end="", flush=True)
            try:
                rng.CopyPicture(Appearance=1, Format=2)
                img = ImageGrab.grabclipboard()
                if img is not None:
                    safe = name.replace(" ", "_")
                    out = os.path.join(BASE, f"test_{safe}.png")
                    img.save(out, "PNG")
                    print(f"  → 저장 {img.size}")
                else:
                    print("  → 클립보드 비어있음")
            except Exception as e:
                print(f"  → 오류: {e}")

    finally:
        wb.Close(False)
    xl.Quit()

except Exception:
    print("\n=== 오류 ===")
    traceback.print_exc()

input("\nEnter 키를 누르면 창이 닫힙니다...")
