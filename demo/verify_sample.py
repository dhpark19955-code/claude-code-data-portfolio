# -*- coding: utf-8 -*-
"""샘플 2종 검증 (Windows + Excel 필요). Excel 이 없으면 verify_sample_lo.py (LibreOffice).

  1) Excel 로 재계산하고 오류 셀을 센다 (0 이어야 함)
  2) 각 파일의 Master check 와 개별 점검을 읽는다. 행 번호는 B열 라벨로 찾는다
  3) 모델은 케이스 스위치를 1, 2, 3 으로 바꿔 가며 DCF·SOTP 주당가치를 읽고,
     엑셀 수식과 무관하게 파이썬으로 다시 계산한 값(independent.py)과 대조한다 (Base 만 보고 통과시키지 않음)
  4) 화면 캡처를 docs/img 에 저장한다

한글 경로에서 Excel 자동화가 실패하는 경우가 있어 임시 폴더(영문 경로)로 복사해 작업한 뒤 되돌린다.
실행:  python -I -X utf8 verify_sample.py
"""
import os
import shutil
import sys
import tempfile
import time

import win32com.client as win32

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)  # -I 옵션은 스크립트 폴더를 검색 경로에 넣지 않는다
from independent import CASES, COVER, SHOTS, independent_per_share, independent_sotp, render_report  # noqa: E402

OUT = os.path.join(HERE, "output")
IMG = os.path.join(os.path.dirname(HERE), "docs", "img")
XL_ERRORS = 16  # xlCellTypeFormulas 의 xlErrors


def count_errors(wb):
    n = 0
    for ws in wb.Worksheets:
        try:
            n += ws.UsedRange.SpecialCells(-4123, XL_ERRORS).Count  # xlCellTypeFormulas
        except Exception:
            pass  # 오류 셀이 없으면 예외가 난다
    return n


def find_row(ws, label, col=2):
    """col 열에서 값이 label 과 같은 첫 행. 없으면 예외."""
    used = ws.UsedRange
    for r in range(1, used.Row + used.Rows.Count):
        if ws.Cells(r, col).Value == label:
            return r
    raise KeyError("%s: '%s' 라벨 없음" % (ws.Name, label))


def read_checks(ws, result_col):
    """머리글('점검')과 'Master check' 사이의 B열이 채워진 행을 전부 점검으로 본다.
    결과가 PASS 가 아니면(FAIL, 오류값 포함) 실패."""
    head, master = find_row(ws, "점검"), find_row(ws, "Master check")
    rows = [r for r in range(head + 1, master) if ws.Cells(r, 2).Value not in (None, "")]
    fails = [ws.Cells(r, 2).Value for r in rows if ws.Cells(r, result_col).Value != "PASS"]
    return ws.Cells(master, result_col).Value, len(rows), fails


def snap(xl, ws, rng, path):
    """범위를 그림으로 복사해 PNG 로 저장."""
    ws.Activate()
    r = ws.Range(rng)
    for _ in range(3):
        try:
            r.CopyPicture(1, 2)  # xlScreen, xlBitmap
            co = ws.ChartObjects().Add(0, 0, r.Width, r.Height)
            co.Activate()
            co.Chart.Paste()
            co.Chart.Export(path)
            co.Delete()
            return True
        except Exception:
            time.sleep(1)
    return False


def main():
    os.makedirs(IMG, exist_ok=True)
    tmp = tempfile.mkdtemp(prefix="ibsample_")
    ok, model_rows = True, []
    xl = win32.DispatchEx("Excel.Application")
    xl.Visible, xl.DisplayAlerts = False, False
    try:
        engine = "Microsoft Excel %s (verify_sample.py)" % xl.Version
        # ---- 데이터셋
        name = "sample_dataset_ib.xlsx"
        p = shutil.copy(os.path.join(OUT, name), os.path.join(tmp, name))
        wb = xl.Workbooks.Open(p)
        xl.CalculateFull()
        err = count_errors(wb)
        master, n_checks, fails = read_checks(wb.Worksheets("99_Check"), 5)
        ds = {"sheets": wb.Worksheets.Count, "errors": err, "n_checks": n_checks, "fails": fails, "master": master}
        ok &= err == 0 and master == "PASS" and not fails
        done = [f for s, r, f in SHOTS["dataset"] if snap(xl, wb.Worksheets(s), r, os.path.join(IMG, f))]
        wb.Save(); wb.Close(False)
        shutil.copy(p, os.path.join(OUT, name))

        # ---- 모델
        name = "sample_model_dcf.xlsx"
        p = shutil.copy(os.path.join(OUT, name), os.path.join(tmp, name))
        wb = xl.Workbooks.Open(p)
        cv = wb.Worksheets("Cover")
        row = {k: find_row(cv, v) for k, v in COVER.items()}
        sw = cv.Cells(row["switch"], 3)
        for i, case in enumerate(CASES):
            sw.Value = i + 1
            xl.CalculateFull()
            x, y = float(cv.Cells(row["per_share"], 3).Value), independent_per_share(i)
            sx, sy = float(cv.Cells(row["sotp"], 3).Value), independent_sotp(i)
            master, err = cv.Cells(row["check"], 3).Value, count_errors(wb)
            c_master, _, c_fails = read_checks(wb.Worksheets("Checks"), 4)
            model_rows.append((case, x, y, sx, sy, master, err))
            ok &= abs(x - y) < 0.01 and abs(sx - sy) < 0.01 and master == "PASS" and c_master == "PASS" \
                and not c_fails and err == 0
        sw.Value = 1
        xl.CalculateFull()
        done += [f for s, r, f in SHOTS["model"] if snap(xl, wb.Worksheets(s), r, os.path.join(IMG, f))]
        wb.Save(); wb.Close(False)
        shutil.copy(p, os.path.join(OUT, name))
    finally:
        xl.Quit()
        shutil.rmtree(tmp, ignore_errors=True)

    ok &= len(done) == len(SHOTS["dataset"]) + len(SHOTS["model"])
    lines = render_report(engine, ds, model_rows, len(done), ok)
    with open(os.path.join(OUT, "verify_report.md"), "w", encoding="utf-8") as f:
        f.write("\n".join(lines) + "\n")
    print("ok=%s shots=%d" % (ok, len(done)))
    return 0 if ok else 1


if __name__ == "__main__":
    raise SystemExit(main())
