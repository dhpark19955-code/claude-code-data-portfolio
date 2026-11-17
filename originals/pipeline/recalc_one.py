# -*- coding: utf-8 -*-
"""xlsx 하나를 Excel COM으로 완전 재계산하고 수식 오류 수를 센다.

사용: python recalc_one.py <파일명 패턴>
  한글 경로를 argv로 넘기면 cp949로 깨지므로 패턴만 받아 Downloads 아래를 glob 한다(최신 수정본 1개).
  여러 파일을 재계산할 때는 파일마다 이 스크립트를 따로 실행한다 —
  한 Excel 인스턴스에서 3개 이상 열면 COM이 끊긴다(-2147417848).
openpyxl은 수식 캐시를 쓰지 않으므로 저장 후 이걸 돌려야 data_only 읽기·미리보기에 값이 보인다.
"""
import sys, os, glob

sys.stdout.reconfigure(encoding="utf-8", errors="replace")
import win32com.client as win32

ROOT = os.path.join(os.path.expanduser("~"), "Downloads")
pat = sys.argv[1] if len(sys.argv) > 1 else "DATASET_IB"
cands = [p for p in glob.glob(os.path.join(ROOT, "**", "*.xlsx"), recursive=True)
         if pat in os.path.basename(p) and not os.path.basename(p).startswith("~$")]
if not cands:
    sys.exit(f"'{pat}' 에 맞는 xlsx 없음")
XL = max(cands, key=os.path.getmtime)
print("대상:", os.path.relpath(XL, ROOT))

xlFormulas, xlErrors = -4123, 16
app = win32.gencache.EnsureDispatch("Excel.Application")
app.Visible = False
app.DisplayAlerts = False
wb = app.Workbooks.Open(XL)
try:
    app.CalculateFullRebuild()
    errs, cells = 0, 0
    for ws in wb.Worksheets:
        try:
            cells += ws.UsedRange.SpecialCells(xlFormulas).Count
        except Exception:
            continue
        try:
            for c in ws.UsedRange.SpecialCells(xlFormulas, xlErrors):
                errs += 1
                print(f"  ERR {ws.Name}!{c.Address} = {c.Formula} -> {c.Text}")
        except Exception:
            pass
    wb.Save()
    print(f"시트 {wb.Worksheets.Count} · 수식 {cells} · errors_found = {errs}")
finally:
    wb.Close(SaveChanges=False)
    app.Quit()
