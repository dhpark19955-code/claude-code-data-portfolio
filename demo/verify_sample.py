# -*- coding: utf-8 -*-
"""샘플 2종 검증 (Windows + Excel 필요).

  1) Excel 로 재계산하고 오류 셀을 센다 (0 이어야 함)
  2) 각 파일의 Master check 를 읽는다
  3) 모델은 케이스 스위치를 1, 2, 3 으로 바꿔 가며 주당가치를 읽고,
     엑셀 수식과 무관하게 파이썬으로 다시 계산한 값과 대조한다 (Base 만 보고 통과시키지 않음)
  4) 화면 캡처를 docs/img 에 저장한다

한글 경로에서 Excel 자동화가 실패하는 경우가 있어 임시 폴더(영문 경로)로 복사해 작업한 뒤 되돌린다.
실행:  python -I -X utf8 verify_sample.py
"""
import os
import shutil
import tempfile
import time

import sys

import win32com.client as win32

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)  # -I 옵션은 스크립트 폴더를 검색 경로에 넣지 않는다
from build_ib_sample import INPUTS  # noqa: E402

OUT = os.path.join(HERE, "output")
IMG = os.path.join(os.path.dirname(HERE), "docs", "img")
XL_ERRORS = 16  # xlCellTypeFormulas 의 xlErrors


def independent_per_share(case):
    """엑셀을 거치지 않는 독립 재계산."""
    I = INPUTS
    g, m = I["growth"][case], I["margin"][case]
    rev, pv, fcf = I["rev0"], 0.0, 0.0
    for t in range(1, 6):
        new = rev * (1 + g)
        ebit = new * m
        fcf = ebit * (1 - I["tax"]) + new * I["da"] - new * I["capex"] - (new - rev) * I["nwc"]
        pv += fcf / (1 + I["wacc"]) ** t
        rev = new
    tv = fcf * (1 + I["g"]) / (I["wacc"] - I["g"]) / (1 + I["wacc"]) ** 5
    return (pv + tv - I["net_debt"]) * 1e6 / I["shares"]


def count_errors(wb):
    n = 0
    for ws in wb.Worksheets:
        try:
            n += ws.UsedRange.SpecialCells(-4123, XL_ERRORS).Count  # xlCellTypeFormulas
        except Exception:
            pass  # 오류 셀이 없으면 예외가 난다
    return n


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
    lines, ok = ["# 검증 리포트 (가상 데이터 샘플)", ""], True
    xl = win32.DispatchEx("Excel.Application")
    xl.Visible, xl.DisplayAlerts = False, False
    try:
        # ---- 데이터셋
        name = "sample_dataset_ib.xlsx"
        p = shutil.copy(os.path.join(OUT, name), os.path.join(tmp, name))
        wb = xl.Workbooks.Open(p)
        xl.CalculateFull()
        err = count_errors(wb)
        master = wb.Worksheets("99_Check").Range("E15").Value
        fails = [wb.Worksheets("99_Check").Range("B%d" % r).Value for r in range(6, 14)
                 if wb.Worksheets("99_Check").Range("E%d" % r).Value != "PASS"]
        lines += ["## 데이터셋", "", "- 시트 %d개, 오류 셀 %d개" % (wb.Worksheets.Count, err),
                  "- 항등식·분기합·배수 역산 8개 점검: 실패 %d개, Master check = %s" % (len(fails), master), ""]
        ok &= err == 0 and master == "PASS" and not fails
        shots = [("01_KPI", "B1:E12", "dataset_kpi.png"), ("10_FS_연간", "B1:E26", "dataset_fs.png"),
                 ("20_밸류에이션", "B1:E16", "dataset_valuation.png"), ("99_Check", "B1:E15", "dataset_check.png")]
        done = [f for s, r, f in shots if snap(xl, wb.Worksheets(s), r, os.path.join(IMG, f))]
        wb.Save(); wb.Close(False)
        shutil.copy(p, os.path.join(OUT, name))

        # ---- 모델
        name = "sample_model_dcf.xlsx"
        p = shutil.copy(os.path.join(OUT, name), os.path.join(tmp, name))
        wb = xl.Workbooks.Open(p)
        cv = wb.Worksheets("Cover")
        lines += ["## 모델 (3케이스 전부)", "", "| 케이스 | 엑셀 주당가치 | 독립 재계산 | 차이 | Master check | 오류 셀 |", "|---|---|---|---|---|---|"]
        for i, case in enumerate(("Base", "Bull", "Bear")):
            cv.Range("C6").Value = i + 1
            xl.CalculateFull()
            x, y = float(cv.Range("C8").Value), independent_per_share(i)
            master, err = cv.Range("C10").Value, count_errors(wb)
            lines.append("| %s | %s | %s | %.4f | %s | %d |" % (case, format(round(x), ","), format(round(y), ","), x - y, master, err))
            ok &= abs(x - y) < 0.01 and master == "PASS" and err == 0
        cv.Range("C6").Value = 1
        xl.CalculateFull()
        shots = [("Cover", "B1:D10", "model_cover.png"), ("Inputs", "B1:F20", "model_inputs.png"),
                 ("Projection", "B1:H14", "model_projection.png"), ("DCF", "B1:G27", "model_dcf.png")]
        done += [f for s, r, f in shots if snap(xl, wb.Worksheets(s), r, os.path.join(IMG, f))]
        wb.Save(); wb.Close(False)
        shutil.copy(p, os.path.join(OUT, name))
    finally:
        xl.Quit()
        shutil.rmtree(tmp, ignore_errors=True)

    lines += ["", "화면 캡처 %d개 저장" % len(done), "", "종합: %s" % ("PASS" if ok else "FAIL")]
    with open(os.path.join(OUT, "verify_report.md"), "w", encoding="utf-8") as f:
        f.write("\n".join(lines) + "\n")
    print("ok=%s shots=%d" % (ok, len(done)))
    return 0 if ok else 1


if __name__ == "__main__":
    raise SystemExit(main())
