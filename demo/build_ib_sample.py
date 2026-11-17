# -*- coding: utf-8 -*-
"""IB 양식 샘플 빌더 (가상 회사 '가온테크', 이 저장소용으로 새로 작성).

산출:
  output/sample_dataset_ib.xlsx  - 수치 전용 데이터셋 8시트
  output/sample_model_dcf.xlsx   - 3케이스 스위치형 DCF 모델 5시트

양식 규칙 (docs/06_ib_format_spec.md 와 같은 내용):
  파랑 = 하드코딩 입력 / 검정 = 같은 시트 수식 / 초록 = 다른 시트 링크
  음수는 괄호, 0 은 '-', 퍼센트는 분수로 저장, 배수는 0.0x
  A열은 여백, B열은 라벨, C열부터 데이터, 격자선 없음, 병합 대신 centerContinuous
  1행 제목, 2행 단위, 3행 출처, 4행 구획 띠, 5행 머리글, 6행부터 데이터

실행:  python -I -X utf8 build_ib_sample.py   (openpyxl 필요)
수식 값은 Excel 이 계산한다. 재계산과 검증은 verify_sample.py.
"""
import os

from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter as L

HERE = os.path.dirname(os.path.abspath(__file__))
OUT = os.path.join(HERE, "output")

FONT = "맑은 고딕"
BLUE, BLACK, GREEN, WHITE, GREY = "FF0000FF", "FF000000", "FF007D00", "FFFFFFFF", "FF595959"
BAND, HEAD, KEY = "FF0018A8", "FFA8BCC2", "FFFFF2CC"
NUM = "#,##0_);(#,##0);-"
NUM1 = "#,##0.0_);(#,##0.0);-"
PCT = "0.0%_);(0.0%);-"
MULT = "0.0x"
COLOR = {"input": BLUE, "formula": BLACK, "link": GREEN, "text": BLACK}


def font(bold=False, color=BLACK, size=9):
    return Font(name=FONT, size=size, bold=bold, color=color)


class Sheet(object):
    """한 시트의 공통 골격과 행 쓰기. self.r 은 다음에 쓸 행 번호."""

    def __init__(self, wb, name, title, unit, source, label_w=28, data_w=11, ncols=8):
        self.ws = ws = wb.create_sheet(name)
        self.name, self.ncols = name, ncols
        ws.sheet_view.showGridLines = False
        ws.column_dimensions["A"].width = 0.9
        ws.column_dimensions["B"].width = label_w
        for i in range(3, 3 + ncols):
            ws.column_dimensions[L(i)].width = data_w
        ws["B1"].value, ws["B1"].font = title, font(True, size=11)
        ws["B2"].value, ws["B2"].font = unit, font()
        ws["B3"].value, ws["B3"].font = source, font(color=GREY)
        self.r = 4

    def band(self, text, width):
        for c in range(2, 2 + width):
            cell = self.ws.cell(self.r, c)
            cell.fill = PatternFill("solid", fgColor=BAND)
            cell.font = font(True, WHITE)
            cell.alignment = Alignment(horizontal="centerContinuous")
        self.ws.cell(self.r, 2).value = text
        self.r += 1

    def header(self, labels):
        for i, text in enumerate(labels):
            cell = self.ws.cell(self.r, 2 + i, text)
            cell.fill = PatternFill("solid", fgColor=HEAD)
            cell.font = font(True)
            cell.alignment = Alignment(horizontal="left" if i == 0 else "center")
        self.r += 1

    def row(self, label, values, kind="input", fmt=NUM, bold=False, start=3, key=False):
        """라벨 1칸 + 값 여러 칸. values 의 None 은 빈칸. 쓴 행 번호를 돌려준다."""
        self.ws.cell(self.r, 2, label).font = font(bold)
        for i, v in enumerate(values):
            if v is None:
                continue
            cell = self.ws.cell(self.r, start + i, v)
            cell.font = font(bold, COLOR[kind])
            cell.alignment = Alignment(horizontal="left" if kind == "text" else "right")
            if kind != "text":
                cell.number_format = fmt
            if key:
                cell.fill = PatternFill("solid", fgColor=KEY)
        self.r += 1
        return self.r - 1

    def note(self, text):
        self.ws.cell(self.r, 2, text).font = font(color=GREY)
        self.r += 1

    def gap(self):
        self.r += 1

    def freeze(self, cell="C6"):
        self.ws.freeze_panes = cell


def ref(sheet, col, row):
    return "'%s'!%s%d" % (sheet, col, row)


# ------------------------------------------------------------------ 데이터셋
def build_dataset(path):
    wb = Workbook()
    wb.remove(wb.active)
    Y, C3 = ["2023/12", "2024/12", "2025/12"], ["C", "D", "E"]
    SRC = "출처: 가상 데이터 (실제 회사와 무관)"

    # 10 연간 재무제표: 원자료는 전부 파랑
    fs = Sheet(wb, "10_FS_연간", "Financial Statements - 연결 · 연간", "(단위: 백만원)", SRC)
    fs.band("손익계산서", 4); fs.header(["항목"] + Y)
    R = {}
    for k, v, b in [
        ("매출액", [4100, 4700, 5300], True), ("매출원가", [2460, 2770, 3070], False),
        ("매출총이익", [1640, 1930, 2230], False), ("판매비와관리비", [1150, 1320, 1540], False),
        ("영업이익", [490, 610, 690], True), ("감가상각비", [210, 240, 270], False),
        ("금융손익", [-40, -45, -50], False), ("법인세비용", [99, 124, 141], False),
        ("당기순이익", [351, 441, 499], True)]:
        R[k] = fs.row(k, [x * 100 for x in v], bold=b)
    fs.gap(); fs.band("재무상태표", 4); fs.header(["항목"] + Y)
    for k, v, b in [
        ("현금및현금성자산", [800, 950, 1100], False), ("기타유동자산", [2300, 2450, 2600], False),
        ("비유동자산", [5600, 5800, 6000], False), ("자산총계", [8700, 9200, 9700], True),
        ("단기차입금", [1000, 950, 900], False), ("장기차입금", [1700, 1600, 1500], False),
        ("기타부채", [1150, 1450, 1700], False), ("부채총계", [3850, 4000, 4100], True),
        ("자본총계", [4850, 5200, 5600], True)]:
        R[k] = fs.row(k, [x * 100 for x in v], bold=b)
    fs.freeze()

    # 11 분기 손익: 3개월치
    Q = ["24.1Q", "24.2Q", "24.3Q", "24.4Q", "25.1Q", "25.2Q", "25.3Q", "25.4Q"]
    qs = Sheet(wb, "11_IS_분기", "Financial Statements - 연결 · 분기", "(단위: 백만원, 3개월치)", SRC)
    qs.band("손익계산서", 9); qs.header(["항목"] + Q)
    QR = {"매출액": qs.row("매출액", [x * 100 for x in (1100, 1150, 1180, 1270, 1200, 1350, 1280, 1470)], bold=True),
          "영업이익": qs.row("영업이익", [x * 100 for x in (130, 150, 155, 175, 150, 190, 160, 190)], bold=True)}
    qs.row("  % margin", ["=%s%d/%s%d" % (L(c), QR["영업이익"], L(c), QR["매출액"]) for c in range(3, 11)],
           "formula", PCT)
    qs.freeze()

    # 01 KPI: 다른 시트 링크(초록) + 같은 시트 수식(검정)
    kp = Sheet(wb, "01_KPI", "Key Performance Indicators", "(단위: 억원, %)", "출처: 10_FS_연간 링크 + 동일 시트 수식")
    kp.band("핵심 KPI (연결, 억원)", 4); kp.header(["항목", "2023A", "2024A", "2025A"])
    K = {}
    for k, b in [("매출액", True), ("영업이익", True), ("당기순이익", False)]:
        K[k] = kp.row(k, ["=%s/100" % ref("10_FS_연간", c, R[k]) for c in C3], "link", NUM, b)
        if k != "당기순이익":
            kp.row("  % growth", [None] + ["=IFERROR(%s%d/%s%d-1,\"\")" % (C3[i], K[k], C3[i - 1], K[k])
                                           for i in (1, 2)], "formula", PCT)
    kp.row("  영업이익률", ["=%s%d/%s%d" % (c, K["영업이익"], c, K["매출액"]) for c in C3], "formula", PCT)
    kp.freeze()

    # 15 재무비율: 전부 10번 시트 참조
    rt = Sheet(wb, "15_재무비율", "Financial Ratios", "(단위: %, 배)", "출처: 10_FS_연간 참조 수식")
    rt.band("수익성 · 안정성", 4); rt.header(["항목", "2023A", "2024A", "2025A"])
    f = lambda k, c: ref("10_FS_연간", c, R[k])
    rt.row("매출총이익률", ["=%s/%s" % (f("매출총이익", c), f("매출액", c)) for c in C3], "link", PCT)
    rt.row("영업이익률", ["=%s/%s" % (f("영업이익", c), f("매출액", c)) for c in C3], "link", PCT)
    rt.row("ROE", ["=%s/%s" % (f("당기순이익", c), f("자본총계", c)) for c in C3], "link", PCT)
    rt.row("부채비율", ["=%s/%s" % (f("부채총계", c), f("자본총계", c)) for c in C3], "link", PCT)
    rt.freeze()

    # 32 순차입금
    nd = Sheet(wb, "32_순차입금", "Net Debt", "(단위: 백만원)", "출처: 10_FS_연간 링크. 리스부채 제외 기준")
    nd.band("순차입금", 4); nd.header(["항목", "2023A", "2024A", "2025A"])
    a = nd.row("단기차입금", ["=" + f("단기차입금", c) for c in C3], "link")
    b = nd.row("장기차입금", ["=" + f("장기차입금", c) for c in C3], "link")
    t = nd.row("총차입금", ["=%s%d+%s%d" % (c, a, c, b) for c in C3], "formula", bold=True)
    c_ = nd.row("현금및현금성자산", ["=" + f("현금및현금성자산", c) for c in C3], "link")
    N = nd.row("순차입금", ["=%s%d-%s%d" % (c, t, c, c_) for c in C3], "formula", bold=True)
    nd.freeze()

    # 20 밸류에이션 배수: 금액 x 배수가 같은 시총으로 돌아와야 한다
    va = Sheet(wb, "20_밸류에이션", "Valuation Multiples", "(단위: 백만원, 원, 배)",
               "출처: 가상 시세 (기준일 2026-11-13 종가), 10_FS_연간 · 32_순차입금 링크", data_w=13)
    va.band("시가총액", 2); va.header(["항목", "값"])
    p = va.row("종가 (원)", [2000], key=True)
    s = va.row("발행주식수 (주)", [500000000])
    ts = va.row("자기주식 (주)", [20000000])
    m = va.row("시가총액", ["=C%d*(C%d-C%d)/1000000" % (p, s, ts)], "formula", bold=True)
    ev = va.row("EV", ["=C%d+%s" % (m, ref("32_순차입금", "E", N))], "link", bold=True)
    va.gap(); va.band("배수 (2025A)", 4); va.header(["항목", "금액", "배수", "역산 시총"])
    r1 = va.row("PER", ["=" + f("당기순이익", "E")], "link")
    va.ws.cell(r1, 4, "=C%d/C%d" % (m, r1)); va.ws.cell(r1, 5, "=C%d*D%d" % (r1, r1))
    r2 = va.row("PBR", ["=" + f("자본총계", "E")], "link")
    va.ws.cell(r2, 4, "=C%d/C%d" % (m, r2)); va.ws.cell(r2, 5, "=C%d*D%d" % (r2, r2))
    r3 = va.row("EV/EBITDA", ["=%s+%s" % (f("영업이익", "E"), f("감가상각비", "E"))], "link")
    va.ws.cell(r3, 4, "=C%d/C%d" % (ev, r3))
    va.ws.cell(r3, 5, "=C%d*D%d-%s" % (r3, r3, ref("32_순차입금", "E", N)))
    for r in (r1, r2, r3):
        for col, fmt in ((4, MULT), (5, NUM)):
            cell = va.ws.cell(r, col)
            cell.font, cell.number_format = font(color=GREEN if (r == r3 and col == 5) else BLACK), fmt
            cell.alignment = Alignment(horizontal="right")
    VAL = {"mcap": m, "rows": (r1, r2, r3)}

    # 99 Check
    ck = Sheet(wb, "99_Check", "Checks", "(차이 0 이면 PASS)", "출처: 각 시트 링크", label_w=36, data_w=12)
    ck.band("기계 감사", 4); ck.header(["점검", "좌변", "우변", "결과"])
    checks = [("자산 = 부채 + 자본 (%s)" % y, "=" + f("자산총계", c), "=%s+%s" % (f("부채총계", c), f("자본총계", c)))
              for y, c in zip(Y, C3)]
    checks += [("분기합 = 연간 매출 (2025)", "=SUM('11_IS_분기'!G%d:J%d)" % (QR["매출액"], QR["매출액"]), "=" + f("매출액", "E")),
               ("분기합 = 연간 영업이익 (2025)", "=SUM('11_IS_분기'!G%d:J%d)" % (QR["영업이익"], QR["영업이익"]), "=" + f("영업이익", "E"))]
    checks += [("역산 시총 = 시총 (%s)" % n, "='20_밸류에이션'!E%d" % r, "='20_밸류에이션'!C%d" % VAL["mcap"])
               for n, r in zip(("PER", "PBR", "EV/EBITDA"), VAL["rows"])]
    first = ck.r
    for name, lhs, rhs in checks:
        r = ck.row(name, [lhs, rhs], "link")
        cell = ck.ws.cell(r, 5, '=IF(ABS(C%d-D%d)<0.5,"PASS","FAIL")' % (r, r))
        cell.font, cell.alignment = font(True), Alignment(horizontal="center")
    ck.gap()
    r = ck.row("Master check", [None, None], "formula", bold=True)
    cell = ck.ws.cell(r, 5, '=IF(COUNTIF(E%d:E%d,"FAIL")=0,"PASS","FAIL")' % (first, r - 2))
    cell.font, cell.alignment = font(True), Alignment(horizontal="center")

    # 00 README: 데이터 사전, 색 규칙, 제외한 것
    rd = Sheet(wb, "00_README", "가온테크 (가상) - DATASET", "수치 전용 데이터셋 · IB 서식 · 기준일 2026-11-13",
               "가상 데이터. 실제 회사, 실제 시세와 무관", label_w=22, data_w=30, ncols=3)
    rd.note("색 규칙: 파랑 = 하드코딩 원자료 / 검정 = 같은 시트 수식 / 초록 = 다른 시트 링크")
    rd.note("숫자 규칙: 음수는 괄호, 0 은 '-', 퍼센트는 분수 저장, 배수는 0.0x")
    rd.gap(); rd.band("제외한 것", 4)
    rd.row("줄글", ["연혁, 사업모델, 논지는 싣지 않음"], "text")
    rd.row("밸류에이션 모델링", ["DCF 는 별도 파일 (sample_model_dcf.xlsx)"], "text")
    rd.gap(); rd.band("데이터 사전", 4); rd.header(["시트", "내용", "단위", "기간"])
    for row in [("01_KPI", "핵심지표 (링크 + 수식)", "억원, %", "2023A~2025A"),
                ("10_FS_연간", "손익계산서, 재무상태표 원자료", "백만원", "2023~2025"),
                ("11_IS_분기", "분기 손익 (3개월치)", "백만원", "24.1Q~25.4Q"),
                ("15_재무비율", "수익성, 안정성 비율", "%", "2023A~2025A"),
                ("20_밸류에이션", "시총, EV, 배수와 역산 검증", "백만원, 배", "기준일"),
                ("32_순차입금", "차입금과 현금", "백만원", "2023A~2025A"),
                ("99_Check", "항등식, 분기합, 배수 역산", "-", "-")]:
        rd.row(row[0], list(row[1:]), "text")

    wb._sheets.sort(key=lambda w: w.title)
    wb.save(path)


# ---------------------------------------------------------------------- 모델
INPUTS = {  # verify_sample.py 의 독립 재계산이 같은 값을 읽는다
    "growth": [0.08, 0.12, 0.04], "margin": [0.135, 0.150, 0.115],
    "tax": 0.22, "da": 0.051, "capex": 0.060, "nwc": 0.10, "wacc": 0.09, "g": 0.02,
    "rev0": 530000, "net_debt": 130000, "shares": 480000000, "price": 2000,
}


def build_model(path):
    wb = Workbook()
    wb.remove(wb.active)
    I = INPUTS
    SW = "Cover!$C$6"

    cv = Sheet(wb, "Cover", "가온테크 (가상) - DCF Model", "(단위: 백만원, 원)", "가상 데이터. 입력은 Inputs 시트에만 있음",
               label_w=30, data_w=14, ncols=3)
    cv.band("Control", 3); cv.header(["항목", "값", "설명"])
    cv.row("케이스 스위치", [1], key=True)
    cv.ws.cell(cv.r - 1, 4, "1 = Base, 2 = Bull, 3 = Bear").font = font(color=GREY)
    r_case = cv.row("선택된 케이스", ['=CHOOSE(C6,"Base","Bull","Bear")'], "formula")
    cv.ws.cell(r_case, 3).alignment = Alignment(horizontal="right")
    cv_rows = {"per_share": cv.r, "upside": cv.r + 1, "check": cv.r + 2}

    ip = Sheet(wb, "Inputs", "Inputs", "(단위: 백만원, %, 주, 원)", "가상 가정. 파랑 = 입력, 노랑 바탕 = 판단값 (사람 확인 대상)",
               label_w=30, data_w=11, ncols=5)
    ip.band("케이스별 가정", 5); ip.header(["항목", "Base", "Bull", "Bear", "선택값"])
    P = {}
    for key, label in (("growth", "매출 성장률"), ("margin", "영업이익률")):
        r = ip.row(label, I[key], fmt=PCT)
        cell = ip.ws.cell(r, 6, "=CHOOSE(%s,C%d,D%d,E%d)" % (SW, r, r, r))
        cell.font, cell.number_format, cell.alignment = font(color=GREEN), PCT, Alignment(horizontal="right")
        P[key] = "Inputs!$F$%d" % r
    ip.gap(); ip.band("공통 가정", 5); ip.header(["항목", "값", "근거"])
    for key, label, fmt, judge, why in [
        ("tax", "법인세율", PCT, False, "가상"), ("da", "감가상각비 / 매출", PCT, False, "2025A 실적 비율"),
        ("capex", "CAPEX / 매출", PCT, False, "가상"), ("nwc", "운전자본 / 매출 증분", PCT, False, "가상"),
        ("wacc", "WACC", PCT, True, "판단값"), ("g", "영구성장률", PCT, True, "판단값"),
        ("rev0", "기준 매출 (2025A)", NUM, False, "데이터셋 10_FS_연간"),
        ("net_debt", "순차입금", NUM, False, "데이터셋 32_순차입금"),
        ("shares", "유통주식수 (주)", NUM, False, "데이터셋 20_밸류에이션"),
        ("price", "현재가 (원)", NUM, False, "기준일 종가")]:
        r = ip.row(label, [I[key]], fmt=fmt, key=judge)
        ip.ws.cell(r, 4, why).font = font(color=GREY)
        P[key] = "Inputs!$C$%d" % r

    YRS = ["2025A", "2026E", "2027E", "2028E", "2029E", "2030E"]
    pj = Sheet(wb, "Projection", "Projection", "(단위: 백만원)", "출처: Inputs 링크 + 동일 시트 수식. 수식 안에 숫자 없음")
    pj.band("실적", 2)
    for c in range(4, 9):
        cell = pj.ws.cell(pj.r - 1, c)
        cell.fill, cell.font = PatternFill("solid", fgColor="FFB0DAF6"), font(True)
        cell.alignment = Alignment(horizontal="centerContinuous")
    pj.ws.cell(pj.r - 1, 4).value = "추정"
    pj.header(["항목"] + YRS)
    E = [L(c) for c in range(4, 9)]  # D..H
    prev = lambda c: L(ord(c) - 64 - 1)
    rev = pj.row("매출액", ["=" + P["rev0"]] + ["=%s%d*(1+%s)" % (prev(c), pj.r, P["growth"]) for c in E], "link", bold=True)
    for c in E:
        pj.ws[c + str(rev)].font = font(True, GREEN)
    pj.row("  % growth", [None] + ["=%s%d/%s%d-1" % (c, rev, prev(c), rev) for c in E], "formula", PCT)
    ebit = pj.row("영업이익", [None] + ["=%s%d*%s" % (c, rev, P["margin"]) for c in E], "link", bold=True)
    tax = pj.row("법인세", [None] + ["=%s%d*%s" % (c, ebit, P["tax"]) for c in E], "link")
    nop = pj.row("NOPAT", [None] + ["=%s%d-%s%d" % (c, ebit, c, tax) for c in E], "formula")
    da = pj.row("감가상각비", [None] + ["=%s%d*%s" % (c, rev, P["da"]) for c in E], "link")
    cpx = pj.row("CAPEX", [None] + ["=%s%d*%s" % (c, rev, P["capex"]) for c in E], "link")
    nwc = pj.row("운전자본 증가", [None] + ["=(%s%d-%s%d)*%s" % (c, rev, prev(c), rev, P["nwc"]) for c in E], "link")
    fcf = pj.row("FCF", [None] + ["=%s%d+%s%d-%s%d-%s%d" % (c, nop, c, da, c, cpx, c, nwc) for c in E], "formula", bold=True)
    pj.freeze()

    dc = Sheet(wb, "DCF", "DCF Valuation", "(단위: 백만원, 원)", "출처: Projection · Inputs 링크. 기말 할인", data_w=12)
    dc.band("현금흐름 할인", 6); dc.header(["항목"] + YRS[1:])
    D = [L(c) for c in range(3, 8)]  # C..G
    t = dc.row("기간 (년)", ["=1"] + ["=%s%d+1" % (L(c - 1), dc.r) for c in range(4, 8)], "formula")
    f_ = dc.row("FCF", ["=Projection!%s%d" % (E[i], fcf) for i in range(5)], "link")
    df = dc.row("할인계수", ["=1/(1+%s)^%s%d" % (P["wacc"], c, t) for c in D], "link", "0.000")
    pv = dc.row("현재가치", ["=%s%d*%s%d" % (c, f_, c, df) for c in D], "formula", bold=True)
    dc.gap(); dc.band("기업가치에서 주당가치까지", 2); dc.header(["항목", "값"])
    s_pv = dc.row("추정기간 현재가치 합", ["=SUM(C%d:G%d)" % (pv, pv)], "formula")
    tv = dc.row("영구가치 (말기)", ["=G%d*(1+%s)/(%s-%s)" % (f_, P["g"], P["wacc"], P["g"])], "link")
    pv_tv = dc.row("영구가치 현재가치", ["=C%d*G%d" % (tv, df)], "formula")
    ev = dc.row("기업가치 (EV)", ["=C%d+C%d" % (s_pv, pv_tv)], "formula", bold=True)
    nd = dc.row("(-) 순차입금", ["=" + P["net_debt"]], "link")
    eq = dc.row("주주가치", ["=C%d-C%d" % (ev, nd)], "formula", bold=True)
    ps = dc.row("주당가치 (원)", ["=C%d*1000000/%s" % (eq, P["shares"])], "link", bold=True)
    up = dc.row("현재가 대비", ["=C%d/%s-1" % (ps, P["price"])], "link", PCT)
    tvs = dc.row("영구가치 비중", ["=C%d/C%d" % (pv_tv, ev)], "formula", PCT)
    dc.gap(); dc.band("민감도: 주당가치 (원)", 4); dc.header(["WACC \\ g", "-1.0%p", "기준", "+1.0%p"])
    sens = {}
    for dw in (-0.01, 0.0, 0.01):
        r = dc.r
        dc.ws.cell(r, 2, "=%s+(%s)" % (P["wacc"], dw))
        dc.ws.cell(r, 2).number_format, dc.ws.cell(r, 2).font = PCT, font(True, GREEN)
        dc.ws.cell(r, 2).alignment = Alignment(horizontal="left")
        for j, dg in enumerate((-0.01, 0.0, 0.01)):
            w, g = "$B%d" % r, "(%s+(%s))" % (P["g"], dg)
            fml = ("=(SUMPRODUCT($C$%d:$G$%d/(1+%s)^$C$%d:$G$%d)+$G$%d*(1+%s)/(%s-%s)/(1+%s)^$G$%d-$C$%d)*1000000/%s"
                   % (f_, f_, w, t, t, f_, g, w, g, w, t, nd, P["shares"]))
            cell = dc.ws.cell(r, 3 + j, fml)
            cell.font, cell.number_format, cell.alignment = font(color=GREEN), NUM, Alignment(horizontal="right")
            sens[(dw, dg)] = (r, 3 + j)
        dc.r += 1
    dc.note("민감도는 증감폭만 수식에 둔 재계산 표. 가운데 칸은 위 주당가치와 같아야 한다 (Checks 에서 대조)")

    ck = Sheet(wb, "Checks", "Checks", "(FAIL 이 하나라도 있으면 납품하지 않음)", "출처: 각 시트 링크", label_w=40, data_w=12, ncols=3)
    ck.band("모델 점검", 3); ck.header(["점검", "값", "결과"])
    mid = sens[(0.0, 0.0)]
    items = [("영구성장률 < WACC", "=%s-%s" % (P["wacc"], P["g"]), "C{r}>0", PCT),
             ("영구가치 비중 < 85%", "=DCF!C%d" % tvs, "C{r}<0.85", PCT),
             ("케이스 스위치가 1~3", "=" + SW, "AND(C{r}>=1,C{r}<=3)", NUM),
             ("민감도 가운데 칸 = 주당가치 (차이)", "=DCF!%s%d-DCF!C%d" % (L(mid[1]), mid[0], ps), "ABS(C{r})<0.5", NUM1),
             ("EV = 추정기간 PV + 영구가치 PV (차이)", "=DCF!C%d-DCF!C%d-DCF!C%d" % (ev, s_pv, pv_tv), "ABS(C{r})<0.5", NUM1),
             ("FCF 가 전 기간 양수 (최솟값)", "=MIN(DCF!C%d:G%d)" % (f_, f_), "C{r}>0", NUM)]
    first = ck.r
    for name, val, cond, fmt in items:
        r = ck.row(name, [val], "link", fmt)
        cell = ck.ws.cell(r, 4, '=IF(%s,"PASS","FAIL")' % cond.format(r=r))
        cell.font, cell.alignment = font(True), Alignment(horizontal="center")
    last = ck.r - 1
    ck.gap()
    r = ck.row("Master check", [None], "formula", bold=True)
    cell = ck.ws.cell(r, 4, '=IF(COUNTIF(D%d:D%d,"FAIL")=0,"PASS","FAIL")' % (first, last))
    cell.font, cell.alignment = font(True), Alignment(horizontal="center")

    cv.row("주당가치 (원)", ["=DCF!C%d" % ps], "link", NUM, True)
    cv.row("현재가 대비", ["=DCF!C%d" % up], "link", PCT)
    cv.row("Master check", ["=Checks!D%d" % r], "link")
    cv.ws.cell(cv.r - 1, 3).alignment = Alignment(horizontal="right")
    assert cv_rows["per_share"] == cv.r - 3

    wb.save(path)
    return {"per_share": "C%d" % cv_rows["per_share"], "check": "C%d" % cv_rows["check"], "switch": "C6",
            "ev": "C%d" % ev}


def main():
    os.makedirs(OUT, exist_ok=True)
    build_dataset(os.path.join(OUT, "sample_dataset_ib.xlsx"))
    cells = build_model(os.path.join(OUT, "sample_model_dcf.xlsx"))
    print("built 2 files", cells)


if __name__ == "__main__":
    main()
