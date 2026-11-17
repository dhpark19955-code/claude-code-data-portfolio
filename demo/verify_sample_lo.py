# -*- coding: utf-8 -*-
"""샘플 2종 검증 (LibreOffice headless. Linux/macOS/Windows).

verify_sample.py (Excel) 와 같은 검증을 Excel 없이 한다.
  1) 임시 사본을 soffice --convert-to xlsx 로 재계산하고 openpyxl(data_only) 로 값을 읽는다
  2) 오류 셀 (#REF!, #DIV/0!, Err:502 등) 과 값이 비어 있는 수식 셀을 센다
  3) 99_Check, Checks 의 개별 점검과 Master check 를 읽는다 (행 번호는 B열 라벨로 찾음)
  4) 모델은 사본의 케이스 스위치를 1, 2, 3 으로 바꿔 DCF·SOTP 주당가치를 independent.py 와 대조
  5) 화면 캡처: 대상 시트만 보이게 한 사본 -> PDF -> pdftoppm -> ImageMagick 으로 여백 정리

LibreOffice 가 다시 저장한 xlsx 는 서식이 달라지므로 값 계산과 캡처에만 쓰고 배포 파일은 건드리지 않는다.
필요: soffice (PATH), openpyxl, pdftoppm (poppler), ImageMagick (magick 또는 convert)
실행: python -I -X utf8 verify_sample_lo.py
"""
import os
import shutil
import subprocess
import sys
import tempfile
from pathlib import Path

from openpyxl import load_workbook
from openpyxl.worksheet.page import PageMargins

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)  # -I 옵션은 스크립트 폴더를 검색 경로에 넣지 않는다
from independent import CASES, COVER, SHOTS, independent_per_share, independent_sotp, render_report  # noqa: E402

OUT = os.path.join(HERE, "output")
IMG = os.path.join(os.path.dirname(HERE), "docs", "img")
SOFFICE = shutil.which("soffice") or shutil.which("libreoffice")
MAGICK = [shutil.which("magick")] if shutil.which("magick") else [shutil.which("convert")]


class LO(object):
    """별도 프로필 폴더를 쓰는 soffice 변환기 (사용 중인 LibreOffice 프로필 잠금을 피함)."""

    def __init__(self, tmp):
        self.profile = Path(tmp, "lo_profile").as_uri()

    def convert(self, src, fmt, outdir):
        os.makedirs(outdir, exist_ok=True)
        cmd = [SOFFICE, "-env:UserInstallation=" + self.profile, "--headless", "--norestore",
               "--calc", "--convert-to", fmt, "--outdir", outdir, src]
        subprocess.run(cmd, check=True, stdout=subprocess.PIPE, stderr=subprocess.PIPE, timeout=300)
        out = os.path.join(outdir, os.path.splitext(os.path.basename(src))[0] + "." + fmt.split(":")[0])
        if not os.path.exists(out):
            raise RuntimeError("soffice 변환 실패: %s" % src)
        return out


def version():
    r = subprocess.run([SOFFICE, "--version"], stdout=subprocess.PIPE, stderr=subprocess.PIPE, timeout=120)
    words = r.stdout.decode("utf-8", "replace").split()
    return " ".join(words[:2]) if words else "LibreOffice"


def recalc(lo, src, workdir, edit=None):
    """사본 (필요하면 openpyxl 로 수정) -> soffice 재계산 -> (수식 wb, 값 wb)."""
    os.makedirs(workdir, exist_ok=True)
    copy = os.path.join(workdir, os.path.basename(src))
    if edit:
        wb = load_workbook(src)
        edit(wb)
        wb.save(copy)
    else:
        shutil.copy(src, copy)
    out = lo.convert(copy, "xlsx", os.path.join(workdir, "calc"))
    return load_workbook(copy), load_workbook(out, data_only=True)


def count_errors(fwb, vwb):
    """오류 문자열 셀 + 값이 비어 있는 수식 셀."""
    n = 0
    for ws in fwb.worksheets:
        vs = vwb[ws.title]
        for row in ws.iter_rows():
            for cell in row:
                v = vs[cell.coordinate].value
                if isinstance(v, str) and (v.startswith("#") or v.startswith("Err:")):
                    n += 1
                elif cell.data_type == "f" and v is None:
                    n += 1
    return n


def find_row(ws, label, col=2):
    for r in range(1, ws.max_row + 1):
        if ws.cell(r, col).value == label:
            return r
    raise KeyError("%s: '%s' 라벨 없음" % (ws.title, label))


def read_checks(ws, result_col):
    """머리글('점검')과 'Master check' 사이 B열이 채워진 행 = 점검. PASS 가 아니면 실패."""
    head, master = find_row(ws, "점검"), find_row(ws, "Master check")
    rows = [r for r in range(head + 1, master) if ws.cell(r, 2).value not in (None, "")]
    fails = [ws.cell(r, 2).value for r in rows if ws.cell(r, result_col).value != "PASS"]
    return ws.cell(master, result_col).value, len(rows), fails


def snap(lo, src, sheet, rng, png, workdir):
    """대상 시트만 보이는 사본에 인쇄 영역을 걸고 PDF -> PNG."""
    os.makedirs(workdir, exist_ok=True)
    wb = load_workbook(src)
    target = wb[sheet]
    wb.active = wb.worksheets.index(target)
    for ws in wb.worksheets:
        ws.sheet_view.tabSelected = ws is target
        ws.sheet_state = "visible" if ws is target else "hidden"
    target.print_area = rng
    target.page_setup.orientation = "landscape"
    target.page_setup.paperSize = target.PAPERSIZE_A4
    target.page_setup.fitToWidth, target.page_setup.fitToHeight = 1, 1
    target.sheet_properties.pageSetUpPr.fitToPage = True
    target.print_options.gridLines = target.print_options.headings = False
    target.page_margins = PageMargins(left=0.3, right=0.3, top=0.3, bottom=0.3, header=0, footer=0)
    for hf in (target.oddHeader, target.oddFooter, target.evenHeader, target.evenFooter,
               target.firstHeader, target.firstFooter):
        hf.left.text = hf.center.text = hf.right.text = None
    base = os.path.splitext(os.path.basename(png))[0]
    xlsx = os.path.join(workdir, base + ".xlsx")
    wb.save(xlsx)
    pdf = lo.convert(xlsx, "pdf", workdir)
    raw = os.path.join(workdir, base)
    subprocess.run(["pdftoppm", "-r", "150", "-png", "-f", "1", "-l", "1", "-singlefile", pdf, raw],
                   check=True, timeout=120)
    subprocess.run(MAGICK + [raw + ".png", "-trim", "+repage", "-bordercolor", "white", "-border", "12",
                             "-bordercolor", "#BFBFBF", "-border", "1", png], check=True, timeout=120)
    return os.path.exists(png)


def main():
    if not SOFFICE:
        print("soffice 를 PATH 에서 찾지 못함")
        return 1
    os.makedirs(IMG, exist_ok=True)
    tmp = tempfile.mkdtemp(prefix="ibsample_lo_")
    lo, ok, done, model_rows = LO(tmp), True, [], []
    try:
        engine = "%s headless (verify_sample_lo.py)" % version()
        # ---- 데이터셋
        src = os.path.join(OUT, "sample_dataset_ib.xlsx")
        fwb, vwb = recalc(lo, src, os.path.join(tmp, "dataset"))
        err = count_errors(fwb, vwb)
        master, n_checks, fails = read_checks(vwb["99_Check"], 5)
        ds = {"sheets": len(vwb.worksheets), "errors": err, "n_checks": n_checks, "fails": fails, "master": master}
        ok &= err == 0 and master == "PASS" and not fails
        for s, r, f in SHOTS["dataset"]:
            if snap(lo, src, s, r, os.path.join(IMG, f), os.path.join(tmp, "shots")):
                done.append(f)

        # ---- 모델: 케이스마다 사본의 스위치를 바꿔 재계산
        src = os.path.join(OUT, "sample_model_dcf.xlsx")
        cover = load_workbook(src)["Cover"]
        row = {k: find_row(cover, v) for k, v in COVER.items()}
        for i, case in enumerate(CASES):
            def edit(wb, v=i + 1):
                wb["Cover"].cell(row["switch"], 3).value = v
            fwb, vwb = recalc(lo, src, os.path.join(tmp, "model_%d" % (i + 1)), edit)
            cv = vwb["Cover"]
            x, y = float(cv.cell(row["per_share"], 3).value), independent_per_share(i)
            sx, sy = float(cv.cell(row["sotp"], 3).value), independent_sotp(i)
            master, err = cv.cell(row["check"], 3).value, count_errors(fwb, vwb)
            c_master, _, c_fails = read_checks(vwb["Checks"], 4)
            model_rows.append((case, x, y, sx, sy, master, err))
            ok &= abs(x - y) < 0.01 and abs(sx - sy) < 0.01 and master == "PASS" and c_master == "PASS" \
                and not c_fails and err == 0
        for s, r, f in SHOTS["model"]:  # 배포 파일 그대로 = Base 케이스
            if snap(lo, src, s, r, os.path.join(IMG, f), os.path.join(tmp, "shots")):
                done.append(f)
    finally:
        shutil.rmtree(tmp, ignore_errors=True)

    ok &= len(done) == len(SHOTS["dataset"]) + len(SHOTS["model"])
    lines = render_report(engine, ds, model_rows, len(done), ok)
    with open(os.path.join(OUT, "verify_report.md"), "w", encoding="utf-8") as f:
        f.write("\n".join(lines) + "\n")
    print("\n".join(lines))
    return 0 if ok else 1


if __name__ == "__main__":
    raise SystemExit(main())
