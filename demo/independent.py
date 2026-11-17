# -*- coding: utf-8 -*-
"""독립 재계산과 검증 공통 상수 (verify_sample.py, verify_sample_lo.py 가 함께 쓴다).

엑셀/LibreOffice 수식을 거치지 않고 INPUTS 만으로 주당가치를 다시 계산한다.
win32com 같은 플랫폼 의존 모듈을 import 하지 않는다.
"""
import os
import sys

HERE = os.path.dirname(os.path.abspath(__file__))
if HERE not in sys.path:
    sys.path.insert(0, HERE)  # -I 옵션은 스크립트 폴더를 검색 경로에 넣지 않는다
from build_ib_sample import INPUTS  # noqa: E402

CASES = ("Base", "Bull", "Bear")

# Cover 시트에서 값을 찾을 라벨 (B열). 행 번호를 하드코딩하지 않는다
COVER = {"switch": "케이스 스위치", "per_share": "주당가치 (원)", "check": "Master check",
         "sotp": "SOTP 주당가치 (원, 참고)"}

# 화면 캡처 (시트, 범위, 파일명). 범위는 각 시트의 사용 영역에 맞춤 (10_FS_연간 은 1~26행만)
SHOTS = {
    "dataset": [("01_KPI", "B1:E13", "dataset_kpi.png"), ("10_FS_연간", "B1:E26", "dataset_fs.png"),
                ("20_밸류에이션", "B1:E27", "dataset_valuation.png"), ("30_매출구성", "B1:E44", "dataset_revenue.png"),
                ("31_FCF", "B1:E24", "dataset_fcf.png"), ("32_순차입금", "B1:E23", "dataset_netdebt.png"),
                ("99_Check", "B1:E24", "dataset_check.png")],
    "model": [("Cover", "B1:E11", "model_cover.png"), ("Inputs", "B1:F31", "model_inputs.png"),
              ("Projection", "B1:H14", "model_projection.png"), ("DCF", "B1:G28", "model_dcf.png"),
              ("SOTP", "B1:G23", "model_sotp.png")],
}


def independent_per_share(case):
    """DCF 주당가치 (원). case = 0 Base, 1 Bull, 2 Bear."""
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


def independent_sotp(case):
    """SOTP 주당가치 (원). 2026E 매출을 부문으로 나눠 배수를 곱하고 브리지를 적용."""
    I = INPUTS
    rev26 = I["rev0"] * (1 + I["growth"][case])
    plat, sol = rev26 * I["plat_share"], rev26 * (1 - I["plat_share"])
    ev = plat * I["plat_ebitda_margin"] * I["plat_ev_ebitda"] + sol * I["sol_ev_sales"]
    non_op = I["assoc_book"] * (1 - I["holdco_disc"])
    equity = ev + non_op - I["net_debt"] - I["nci"]
    return equity * 1e6 / I["shares"]


def render_report(engine, ds, model, n_shots, ok):
    """검증 리포트 (마크다운 줄 목록).
    ds: dict(sheets, errors, n_checks, fails, master)
    model: [(case, dcf_calc, dcf_ind, sotp_calc, sotp_ind, master, errors), ...]
    """
    lines = ["# 검증 리포트 (가상 데이터 샘플)", "", "검증 엔진: %s" % engine, "",
             "## 데이터셋", "",
             "- 시트 %d개, 오류 셀 %d개" % (ds["sheets"], ds["errors"]),
             "- 항등식·분기합·부문합·FCF·EV 브리지·배수 역산 %d개 점검: 실패 %d개, Master check = %s"
             % (ds["n_checks"], len(ds["fails"]), ds["master"])]
    lines += ["  - 실패: %s" % x for x in ds["fails"]]
    lines += ["", "## 모델 (3케이스 전부)", "",
              "| 케이스 | DCF 주당가치 | DCF 독립 재계산 | 차이 | SOTP 주당가치 | SOTP 독립 재계산 | 차이 | Master check | 오류 셀 |",
              "|---|---|---|---|---|---|---|---|---|"]
    for case, x, y, sx, sy, master, err in model:
        lines.append("| %s | %s | %s | %.4f | %s | %s | %.4f | %s | %d |" % (
            case, format(round(x), ","), format(round(y), ","), x - y,
            format(round(sx), ","), format(round(sy), ","), sx - sy, master, err))
    lines += ["", "화면 캡처 %d개 저장" % n_shots, "", "종합: %s" % ("PASS" if ok else "FAIL")]
    return lines
