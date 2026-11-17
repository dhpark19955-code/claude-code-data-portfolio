# -*- coding: utf-8 -*-
"""IB 양식 샘플 빌더 (가상 회사 '가온테크', 이 저장소용으로 새로 작성).

산출:
  output/sample_dataset_ib.xlsx  - 수치 전용 데이터셋 10시트
                                   (재무제표, 매출 구성, FCF, 순차입금·순현금, EV 브리지·배수, 점검)
  output/sample_model_dcf.xlsx   - 3케이스 스위치형 DCF + SOTP 모델 6시트

양식 규칙 (docs/06_ib_format_spec.md 와 같은 내용):
  파랑 = 하드코딩 입력 / 검정 = 같은 시트 수식 / 초록 = 다른 시트 링크
  음수는 괄호, 0 은 '-', 퍼센트는 분수로 저장, 배수는 0.0x
  A열은 여백, B열은 라벨, C열부터 데이터, 격자선 없음, 병합 대신 centerContinuous
  1행 제목, 2행 단위, 3행 출처, 4행 구획 띠, 5행 머리글, 6행부터 데이터

실행:  python -I -X utf8 build_ib_sample.py   (openpyxl 필요)
수식 값은 저장하지 않는다 (fullCalcOnLoad). 계산과 검증은 verify_sample.py (Excel)
또는 verify_sample_lo.py (LibreOffice).
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

    def put(self, r, c, value, kind="formula", fmt=NUM, bold=False, align="right"):
        """이미 쓴 행의 한 칸을 따로 채운다 (한 행 안에 링크와 수식이 섞일 때)."""
        cell = self.ws.cell(r, c, value)
        cell.font, cell.alignment = font(bold, COLOR[kind]), Alignment(horizontal=align)
        if kind != "text":
            cell.number_format = fmt
        return cell

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
    H3 = ["항목", "2023A", "2024A", "2025A"]
    SRC = "출처: 가상 데이터 (실제 회사와 무관)"

    # 10 연간 재무제표: 원자료는 전부 파랑. 1~26행은 손익·재무상태표, 그 아래에 주석·현금흐름표
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
    fs.gap(); fs.band("자본 구성 · 주석", 4); fs.header(["항목"] + Y)
    for k, label, v in [
        ("지배주주지분", "지배주주지분", [455000, 488000, 525000]),
        ("비지배지분", "비지배지분", [30000, 32000, 35000]),
        ("지배주주순이익", "지배주주순이익", [33500, 42100, 47600]),
        ("관계기업투자", "(비유동자산 중) 관계기업투자", [60000, 65000, 70000]),
        ("리스부채", "(기타부채 중) 리스부채", [20000, 22000, 25000])]:
        R[k] = fs.row(label, v)
    fs.note("지배주주지분 + 비지배지분 = 자본총계 (99_Check 에서 대조)")
    fs.gap(); fs.band("현금흐름표", 4); fs.header(["항목"] + Y)
    for k, v, b in [
        ("영업활동현금흐름", [62000, 75000, 86000], True), ("유형자산의 취득", [-32000, -38000, -41000], False),
        ("무형자산의 취득", [-5000, -6000, -7000], False), ("배당금의 지급", [-10000, -12000, -14000], False)]:
        R[k] = fs.row(k, v, bold=b)
    fs.note("유출(취득, 배당)은 음수로 저장. abs() 쓰지 않음")
    fs.freeze()
    f = lambda k, c: ref("10_FS_연간", c, R[k])

    # 11 분기 손익: 3개월치
    Q = ["24.1Q", "24.2Q", "24.3Q", "24.4Q", "25.1Q", "25.2Q", "25.3Q", "25.4Q"]
    qs = Sheet(wb, "11_IS_분기", "Financial Statements - 연결 · 분기", "(단위: 백만원, 3개월치)", SRC)
    qs.band("손익계산서", 9); qs.header(["항목"] + Q)
    QR = {"매출액": qs.row("매출액", [x * 100 for x in (1100, 1150, 1180, 1270, 1200, 1350, 1280, 1470)], bold=True),
          "영업이익": qs.row("영업이익", [x * 100 for x in (130, 150, 155, 175, 150, 190, 160, 190)], bold=True)}
    qs.row("  % margin", ["=%s%d/%s%d" % (L(c), QR["영업이익"], L(c), QR["매출액"]) for c in range(3, 11)],
           "formula", PCT)
    qs.freeze()

    # 30 매출 구성: 부문·지역 원자료는 파랑, 연결 수치는 10번 링크, 차이는 0 이어야 한다
    rv = Sheet(wb, "30_매출구성", "Revenue Breakdown - 부문 · 지역", "(단위: 백만원, %)",
               "출처: 사업보고서 'Ⅱ.사업의 내용' 매출 표 + 연결주석 영업부문 (가상)")
    RV = {}
    rv.band("부문별 매출", 4); rv.header(H3)
    SEG = {k: rv.row(k, v) for k, v in [
        ("플랫폼", [230000, 270000, 310000]), ("솔루션", [175000, 196000, 220000]),
        ("기타", [9000, 9000, 6000]), ("내부거래 제거", [-4000, -5000, -6000])]}
    RV["seg_sum"] = rv.row("부문 합계", ["=SUM(%s%d:%s%d)" % (c, SEG["플랫폼"], c, SEG["내부거래 제거"]) for c in C3],
                           "formula", bold=True)
    con = rv.row("연결 매출", ["=" + f("매출액", c) for c in C3], "link")
    rv.row("차이", ["=%s%d-%s%d" % (c, RV["seg_sum"], c, con) for c in C3], "formula")
    rv.gap(); rv.band("부문 비중 · 성장률", 4); rv.header(H3)
    for k in ("플랫폼", "솔루션", "기타"):
        rv.row("%s 비중" % k, ["=%s%d/%s%d" % (c, SEG[k], c, con) for c in C3], "formula", PCT)
    for k in ("플랫폼", "솔루션", "기타"):
        rv.row("%s 성장률" % k, [None] + ["=%s%d/%s%d-1" % (C3[i], SEG[k], C3[i - 1], SEG[k]) for i in (1, 2)],
               "formula", PCT)
    rv.note("비중 분모 = 연결 매출 (내부거래 제거 후). 세 부문 비중 합은 내부거래만큼 100% 를 넘는다")
    rv.gap(); rv.band("부문별 영업이익", 4); rv.header(H3)
    OP = {k: rv.row(k, v) for k, v in [
        ("플랫폼", [38000, 47000, 52000]), ("솔루션", [13000, 16000, 19000]),
        ("기타", [-1000, -1000, -1000]), ("본사 미배분 · 조정", [-1000, -1000, -1000])]}
    RV["op_sum"] = rv.row("부문 합계", ["=SUM(%s%d:%s%d)" % (c, OP["플랫폼"], c, OP["본사 미배분 · 조정"]) for c in C3],
                          "formula", bold=True)
    op_con = rv.row("연결 영업이익", ["=" + f("영업이익", c) for c in C3], "link")
    rv.row("차이", ["=%s%d-%s%d" % (c, RV["op_sum"], c, op_con) for c in C3], "formula")
    for k in ("플랫폼", "솔루션", "기타"):
        rv.row("%s 영업이익률" % k, ["=%s%d/%s%d" % (c, OP[k], c, SEG[k]) for c in C3], "formula", PCT)
    rv.gap(); rv.band("지역별 매출", 4); rv.header(H3)
    dom = rv.row("국내", [330000, 365000, 400000])
    ovs = rv.row("해외", [80000, 105000, 130000])
    RV["geo_sum"] = rv.row("지역 합계", ["=%s%d+%s%d" % (c, dom, c, ovs) for c in C3], "formula", bold=True)
    geo_con = rv.row("연결 매출", ["=" + f("매출액", c) for c in C3], "link")
    rv.row("차이", ["=%s%d-%s%d" % (c, RV["geo_sum"], c, geo_con) for c in C3], "formula")
    rv.row("해외 비중", ["=%s%d/%s%d" % (c, ovs, c, geo_con) for c in C3], "formula", PCT)
    rv.freeze()

    # 32 순차입금: 1~10행은 기존 그대로, 아래에 리스 포함 · 순현금 · 레버리지
    nd = Sheet(wb, "32_순차입금", "Net Debt / Net Cash", "(단위: 백만원, 배)",
               "출처: 10_FS_연간 링크. 순차입금은 리스부채 제외 기준 (리스 포함은 별도 행)")
    nd.band("순차입금", 4); nd.header(H3)
    a = nd.row("단기차입금", ["=" + f("단기차입금", c) for c in C3], "link")
    b = nd.row("장기차입금", ["=" + f("장기차입금", c) for c in C3], "link")
    t = nd.row("총차입금", ["=%s%d+%s%d" % (c, a, c, b) for c in C3], "formula", bold=True)
    c_ = nd.row("현금및현금성자산", ["=" + f("현금및현금성자산", c) for c in C3], "link")
    N = nd.row("순차입금", ["=%s%d-%s%d" % (c, t, c, c_) for c in C3], "formula", bold=True)
    ls = nd.row("리스부채", ["=" + f("리스부채", c) for c in C3], "link")
    nd.row("순차입금 (리스 포함)", ["=%s%d+%s%d" % (c, N, c, ls) for c in C3], "formula", bold=True)
    nd.gap(); nd.band("순현금 (Net cash)", 4); nd.header(H3)
    nd.row("순현금", ["=%s%d-%s%d" % (c, c_, c, t) for c in C3], "formula", bold=True)
    st = nd.row("상태", ['=IF(%s%d<0,"순현금","순차입금")' % (c, N) for c in C3], "text")
    for c in C3:
        nd.ws[c + str(st)].alignment = Alignment(horizontal="right")
    eb = nd.row("EBITDA", ["=%s+%s" % (f("영업이익", c), f("감가상각비", c)) for c in C3], "link")
    nd.row("순차입금 / EBITDA (배)", ["=%s%d/%s%d" % (c, N, c, eb) for c in C3], "formula", MULT)
    nd.note("부호: 순차입금 = 총차입금 - 현금. 양수 = 순차입, 음수 = 순현금")
    nd.note("순현금 = 현금 - 총차입금 = -순차입금")
    nd.note("리스부채(IFRS 16)는 기본 순차입금에서 제외, '리스 포함' 행에만 더함")
    nd.note("EBITDA = 영업이익 + 감가상각비")
    nd.freeze()

    # 20 밸류에이션: 시총 -> EV 브리지 -> 배수. 금액 x 배수가 같은 시총으로 돌아와야 한다
    va = Sheet(wb, "20_밸류에이션", "Valuation - EV Bridge · Multiples", "(단위: 백만원, 원, 배)",
               "출처: 가상 시세 (기준일 2026-11-13 종가), 10_FS_연간 · 32_순차입금 링크", data_w=13)
    va.band("시가총액", 2); va.header(["항목", "값"])
    p = va.row("종가 (원)", [2000], key=True)
    s = va.row("발행주식수 (주)", [500000000])
    ts = va.row("자기주식 (주)", [20000000])
    m = va.row("시가총액", ["=C%d*(C%d-C%d)/1000000" % (p, s, ts)], "formula", bold=True)
    va.gap(); va.band("EV 산출 (2025A 말 기준)", 2); va.header(["항목", "값"])
    b0 = va.row("시가총액", ["=C%d" % m], "formula")
    b1 = va.row("(+) 순차입금", ["=" + ref("32_순차입금", "E", N)], "link")
    b2 = va.row("(+) 비지배지분", ["=" + f("비지배지분", "E")], "link")
    b3 = va.row("(-) 비영업자산: 관계기업투자", ["=" + f("관계기업투자", "E")], "link")
    ev = va.row("EV", ["=C%d+C%d+C%d-C%d" % (b0, b1, b2, b3)], "formula", bold=True)
    va.note("시총은 기준일, 나머지는 2025A 말 장부. 비영업자산은 양수로 두고 뺀다")
    va.gap(); va.band("배수 (2025A)", 4); va.header(["항목", "금액", "배수", "역산 시총"])
    mult_rows = []
    for name, amount, base in [
            ("PER (지배주주순이익)", "=" + f("지배주주순이익", "E"), m),
            ("PBR (지배주주지분)", "=" + f("지배주주지분", "E"), m),
            ("EV/EBITDA", "=%s+%s" % (f("영업이익", "E"), f("감가상각비", "E")), ev),
            ("EV/Sales", "=" + f("매출액", "E"), ev)]:
        r = va.row(name, [amount], "link")
        va.put(r, 4, "=C%d/C%d" % (base, r), fmt=MULT)
        back = "=C%d*D%d" % (r, r) if base == m else "=C%d*D%d-(C%d-C%d)" % (r, r, ev, m)
        va.put(r, 5, back)
        mult_rows.append(r)
    va.note("EBITDA = 영업이익 + 감가상각비. PER · PBR 은 지배주주 기준")
    va.note("역산 시총: 금액 x 배수 (EV 배수는 여기서 EV - 시총을 뺀다)")
    VAL = {"mcap": m, "ev": ev, "rows": mult_rows}

    # 31 FCF: 현금흐름표 링크, 부호 그대로 합산
    fc = Sheet(wb, "31_FCF", "Free Cash Flow", "(단위: 백만원, %)",
               "출처: 10_FS_연간 현금흐름표 링크. FCF = 영업활동현금흐름 − 유형·무형자산 취득")
    fc.band("잉여현금흐름 (FCF)", 4); fc.header(H3)
    o = fc.row("영업활동현금흐름", ["=" + f("영업활동현금흐름", c) for c in C3], "link")
    k1 = fc.row("유형자산의 취득", ["=" + f("유형자산의 취득", c) for c in C3], "link")
    k2 = fc.row("무형자산의 취득", ["=" + f("무형자산의 취득", c) for c in C3], "link")
    FCF = fc.row("FCF", ["=%s%d+%s%d+%s%d" % (c, o, c, k1, c, k2) for c in C3], "formula", bold=True)
    fc.note("CAPEX 는 유출이라 음수, abs() 쓰지 않음. FCF = OCF + 유형 + 무형")
    fc.gap(); fc.band("마진 · 현금전환율", 4); fc.header(H3)
    sl = fc.row("매출액", ["=" + f("매출액", c) for c in C3], "link")
    fc.row("FCF 마진", ["=%s%d/%s%d" % (c, FCF, c, sl) for c in C3], "formula", PCT)
    ni = fc.row("지배주주순이익", ["=" + f("지배주주순이익", c) for c in C3], "link")
    fc.row("현금전환율 (FCF / 지배순이익)", ["=%s%d/%s%d" % (c, FCF, c, ni) for c in C3], "formula", PCT)
    fc.gap(); fc.band("배당 · FCF yield", 4); fc.header(H3)
    dv = fc.row("배당금의 지급", ["=" + f("배당금의 지급", c) for c in C3], "link")
    fc.row("배당 후 FCF", ["=%s%d+%s%d" % (c, FCF, c, dv) for c in C3], "formula", bold=True)
    mc = fc.row("시가총액 (기준일)", [None, None, "=" + ref("20_밸류에이션", "C", VAL["mcap"])], "link")
    fc.row("FCF yield (2025A)", [None, None, "=E%d/E%d" % (FCF, mc)], "formula", PCT)
    fc.freeze()

    # 01 KPI: 다른 시트 링크(초록) + 같은 시트 수식(검정)
    kp = Sheet(wb, "01_KPI", "Key Performance Indicators", "(단위: 억원, %)", "출처: 10_FS_연간 · 31_FCF · 32_순차입금 링크 + 동일 시트 수식")
    kp.band("핵심 KPI (연결, 억원)", 4); kp.header(H3)
    K = {}
    for k, b in [("매출액", True), ("영업이익", True), ("당기순이익", False)]:
        K[k] = kp.row(k, ["=%s/100" % ref("10_FS_연간", c, R[k]) for c in C3], "link", NUM, b)
        if k != "당기순이익":
            kp.row("  % growth", [None] + ["=IFERROR(%s%d/%s%d-1,\"\")" % (C3[i], K[k], C3[i - 1], K[k])
                                           for i in (1, 2)], "formula", PCT)
    kp.row("  영업이익률", ["=%s%d/%s%d" % (c, K["영업이익"], c, K["매출액"]) for c in C3], "formula", PCT)
    kp.row("FCF", ["=%s/100" % ref("31_FCF", c, FCF) for c in C3], "link", NUM)
    kp.row("순차입금", ["=%s/100" % ref("32_순차입금", c, N) for c in C3], "link", NUM)
    kp.freeze()

    # 15 재무비율: 전부 10번 시트 참조
    rt = Sheet(wb, "15_재무비율", "Financial Ratios", "(단위: %, 배)", "출처: 10_FS_연간 참조 수식")
    rt.band("수익성 · 안정성", 4); rt.header(H3)
    rt.row("매출총이익률", ["=%s/%s" % (f("매출총이익", c), f("매출액", c)) for c in C3], "link", PCT)
    rt.row("영업이익률", ["=%s/%s" % (f("영업이익", c), f("매출액", c)) for c in C3], "link", PCT)
    rt.row("ROE", ["=%s/%s" % (f("당기순이익", c), f("자본총계", c)) for c in C3], "link", PCT)
    rt.row("부채비율", ["=%s/%s" % (f("부채총계", c), f("자본총계", c)) for c in C3], "link", PCT)
    rt.freeze()

    # 99 Check: 좌변은 결과 시트, 우변은 원자료에서 다시 계산
    ck = Sheet(wb, "99_Check", "Checks", "(차이 0 이면 PASS)", "출처: 각 시트 링크", label_w=46, data_w=12)
    ck.band("기계 감사", 4); ck.header(["점검", "좌변", "우변", "결과"])
    checks = [("자산 = 부채 + 자본 (%s)" % y, "=" + f("자산총계", c), "=%s+%s" % (f("부채총계", c), f("자본총계", c)))
              for y, c in zip(Y, C3)]
    checks += [("분기합 = 연간 매출 (2025)", "=SUM('11_IS_분기'!G%d:J%d)" % (QR["매출액"], QR["매출액"]), "=" + f("매출액", "E")),
               ("분기합 = 연간 영업이익 (2025)", "=SUM('11_IS_분기'!G%d:J%d)" % (QR["영업이익"], QR["영업이익"]), "=" + f("영업이익", "E"))]
    checks += [("역산 시총 = 시총 (%s)" % n, "=" + ref("20_밸류에이션", "E", r), "=" + ref("20_밸류에이션", "C", VAL["mcap"]))
               for n, r in zip(("PER", "PBR", "EV/EBITDA", "EV/Sales"), VAL["rows"])]
    checks += [("자본 = 지배 + 비지배 (2025)", "=" + f("자본총계", "E"), "=%s+%s" % (f("지배주주지분", "E"), f("비지배지분", "E")))]
    checks += [("부문 매출 합 = 연결 매출 (%s)" % y[:4], "=" + ref("30_매출구성", c, RV["seg_sum"]), "=" + f("매출액", c))
               for y, c in zip(Y, C3)]
    checks += [("부문 영업이익 합 = 연결 영업이익 (2025)", "=" + ref("30_매출구성", "E", RV["op_sum"]), "=" + f("영업이익", "E")),
               ("지역 매출 합 = 연결 매출 (2025)", "=" + ref("30_매출구성", "E", RV["geo_sum"]), "=" + f("매출액", "E")),
               ("FCF = OCF + CAPEX (2025)", "=" + ref("31_FCF", "E", FCF),
                "=%s+%s+%s" % (f("영업활동현금흐름", "E"), f("유형자산의 취득", "E"), f("무형자산의 취득", "E"))),
               ("EV = 시총 + 순차입금 + 비지배 - 비영업 (2025)", "=" + ref("20_밸류에이션", "C", VAL["ev"]),
                "=%s+%s+%s-%s" % (ref("20_밸류에이션", "C", VAL["mcap"]), ref("32_순차입금", "E", N),
                                  f("비지배지분", "E"), f("관계기업투자", "E")))]
    first = ck.r
    for name, lhs, rhs in checks:
        r = ck.row(name, [lhs, rhs], "link")
        cell = ck.ws.cell(r, 5, '=IF(ABS(C%d-D%d)<0.5,"PASS","FAIL")' % (r, r))
        cell.font, cell.alignment = font(True), Alignment(horizontal="center")
    last = ck.r - 1
    ck.gap()
    r = ck.row("Master check", [None, None], "formula", bold=True)
    cell = ck.ws.cell(r, 5, '=IF(COUNTIF(E%d:E%d,"FAIL")=0,"PASS","FAIL")' % (first, last))
    cell.font, cell.alignment = font(True), Alignment(horizontal="center")

    # 00 README: 데이터 사전, 정의, 색 규칙, 제외한 것
    rd = Sheet(wb, "00_README", "가온테크 (가상) - DATASET", "수치 전용 데이터셋 · IB 서식 · 기준일 2026-11-13",
               "가상 데이터. 실제 회사, 실제 시세와 무관", label_w=22, data_w=30, ncols=3)
    rd.note("색 규칙: 파랑 = 하드코딩 원자료 / 검정 = 같은 시트 수식 / 초록 = 다른 시트 링크")
    rd.note("숫자 규칙: 음수는 괄호, 0 은 '-', 퍼센트는 분수 저장, 배수는 0.0x")
    rd.gap(); rd.band("제외한 것", 4)
    rd.row("줄글", ["연혁, 사업모델, 논지는 싣지 않음"], "text")
    rd.row("밸류에이션 모델링", ["DCF · SOTP 는 별도 파일 (sample_model_dcf.xlsx)"], "text")
    rd.gap(); rd.band("데이터 사전", 4); rd.header(["시트", "내용", "단위", "기간"])
    for row in [("01_KPI", "핵심지표 (매출, 이익, FCF, 순차입금)", "억원, %", "2023A~2025A"),
                ("10_FS_연간", "손익, 재무상태, 자본 구성, 현금흐름표", "백만원", "2023~2025"),
                ("11_IS_분기", "분기 손익 (3개월치)", "백만원", "24.1Q~25.4Q"),
                ("15_재무비율", "수익성, 안정성 비율", "%", "2023A~2025A"),
                ("20_밸류에이션", "시총, EV 브리지, 배수와 역산 검증", "백만원, 배", "기준일 · 2025A"),
                ("30_매출구성", "부문별 매출·영업이익, 지역별 매출", "백만원, %", "2023A~2025A"),
                ("31_FCF", "FCF, 마진, 현금전환율, FCF yield", "백만원, %", "2023A~2025A"),
                ("32_순차입금", "순차입금 (리스 제외/포함), 순현금", "백만원, 배", "2023A~2025A"),
                ("99_Check", "항등식, 분기합, 부문합, FCF, EV, 역산", "-", "-")]:
        rd.row(row[0], list(row[1:]), "text")
    rd.gap(); rd.band("정의", 4)
    for k, v in [("FCF", "영업활동현금흐름 + 유형자산의 취득 + 무형자산의 취득 (취득은 음수 저장)"),
                 ("순차입금", "단기차입금 + 장기차입금 - 현금및현금성자산 (리스부채 제외)"),
                 ("순차입금 (리스 포함)", "순차입금 + 리스부채. 32_순차입금 에 별도 행"),
                 ("순현금", "현금 - 총차입금 = -순차입금 (양수면 순현금)"),
                 ("EV", "시가총액 + 순차입금 + 비지배지분 - 비영업자산(관계기업투자)"),
                 ("PER / PBR", "지배주주순이익 / 지배주주지분 기준 (비지배 제외)"),
                 ("EBITDA", "영업이익 + 감가상각비")]:
        rd.row(k, [v], "text")

    wb._sheets.sort(key=lambda w: w.title)
    wb.calculation.fullCalcOnLoad = True
    wb.save(path)


# ---------------------------------------------------------------------- 모델
INPUTS = {  # independent.py 의 독립 재계산이 같은 값을 읽는다
    "growth": [0.08, 0.12, 0.04], "margin": [0.135, 0.150, 0.115],
    "tax": 0.22, "da": 0.051, "capex": 0.060, "nwc": 0.10, "wacc": 0.09, "g": 0.02,
    "rev0": 530000, "net_debt": 130000, "shares": 480000000, "price": 2000,
    # SOTP
    "plat_share": 0.585, "plat_ebitda_margin": 0.22, "plat_ev_ebitda": 9.0, "sol_ev_sales": 1.2,
    "assoc_book": 70000, "holdco_disc": 0.30, "nci": 35000,
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
    ip.gap(); ip.band("SOTP 가정", 5); ip.header(["항목", "값", "근거"])
    for key, label, fmt, judge, why in [
        ("plat_share", "플랫폼 매출 비중", PCT, False, "데이터셋 30_매출구성 2025A 비중"),
        ("plat_ebitda_margin", "플랫폼 EBITDA 마진", PCT, False, "가상 (부문 영업이익률 + 상각)"),
        ("plat_ev_ebitda", "플랫폼 EV/EBITDA (배)", MULT, True, "판단값"),
        ("sol_ev_sales", "솔루션 EV/Sales (배)", MULT, True, "판단값"),
        ("assoc_book", "관계기업투자 장부가", NUM, False, "데이터셋 10_FS_연간 주석"),
        ("holdco_disc", "지주 할인 (비영업자산)", PCT, True, "판단값"),
        ("nci", "비지배지분", NUM, False, "데이터셋 10_FS_연간 비지배지분")]:
        r = ip.row(label, [I[key]], fmt=fmt, key=judge)
        ip.ws.cell(r, 4, why).font = font(color=GREY)
        P[key] = "Inputs!$C$%d" % r
        if key == "plat_share":
            r2 = ip.row("솔루션 매출 비중", ["=1-C%d" % r], "formula", PCT)
            ip.ws.cell(r2, 4, "1 - 플랫폼 비중").font = font(color=GREY)
            P["sol_share"] = "Inputs!$C$%d" % r2

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

    # SOTP: 부문별 배수 합산. 2026E 매출이 케이스 스위치를 따라 바뀐다
    so = Sheet(wb, "SOTP", "SOTP Valuation", "(단위: 백만원, 원, 배)",
               "출처: Projection · Inputs · DCF 링크. 배수와 지주 할인은 판단값 (Inputs 노랑)",
               label_w=32, data_w=12, ncols=5)
    so.band("부문별 가치 (2026E 기준)", 6)
    so.header(["부문", "2026E 매출", "기준 지표", "지표 값", "배수", "부문 가치(EV)"])
    rev26 = "Projection!$D$%d" % rev
    sp = so.row("플랫폼", ["=%s*%s" % (rev26, P["plat_share"])], "link")
    so.put(sp, 4, "EBITDA", "text", align="center")
    so.put(sp, 5, "=C%d*%s" % (sp, P["plat_ebitda_margin"]), "link")
    so.put(sp, 6, "=" + P["plat_ev_ebitda"], "link", MULT)
    so.put(sp, 7, "=E%d*F%d" % (sp, sp))
    ss = so.row("솔루션", ["=%s*%s" % (rev26, P["sol_share"])], "link")
    so.put(ss, 4, "매출", "text", align="center")
    so.put(ss, 5, "=C%d" % ss)
    so.put(ss, 6, "=" + P["sol_ev_sales"], "link", MULT)
    so.put(ss, 7, "=E%d*F%d" % (ss, ss))
    st = so.row("합계", ["=SUM(C%d:C%d)" % (sp, ss)], "formula", bold=True)
    so.put(st, 7, "=SUM(G%d:G%d)" % (sp, ss), bold=True)
    so.gap(); so.band("EV에서 주당가치까지", 2); so.header(["항목", "값"])
    s_ev = so.row("부문 가치 합계 (영업 EV)", ["=G%d" % st], "formula", bold=True)
    s_no = so.row("(+) 비영업자산 (관계기업, 할인 후)", ["=%s*(1-%s)" % (P["assoc_book"], P["holdco_disc"])], "link")
    s_nd = so.row("(-) 순차입금", ["=" + P["net_debt"]], "link")
    s_mi = so.row("(-) 비지배지분", ["=" + P["nci"]], "link")
    s_eq = so.row("주주가치", ["=C%d+C%d-C%d-C%d" % (s_ev, s_no, s_nd, s_mi)], "formula", bold=True)
    s_ps = so.row("주당가치 (원)", ["=C%d*1000000/%s" % (s_eq, P["shares"])], "link", bold=True)
    so.row("현재가 대비", ["=C%d/%s-1" % (s_ps, P["price"])], "link", PCT)
    s_dcf = so.row("DCF 주당가치 (원)", ["=DCF!C%d" % ps], "link")
    so.row("SOTP - DCF 차이", ["=C%d/C%d-1" % (s_ps, s_dcf)], "formula", PCT)
    so.row("암묵적 EV/EBITDA (2026E)", ["=C%d/(Projection!$D$%d+Projection!$D$%d)" % (s_ev, ebit, da)], "link", MULT)
    so.note("플랫폼은 EV/EBITDA, 솔루션은 EV/Sales. 부문 비중은 데이터셋 2025A 기준")
    so.note("DCF 브리지는 순차입금만 뺀다. SOTP 는 비영업자산과 비지배지분까지 반영")

    ck = Sheet(wb, "Checks", "Checks", "(FAIL 이 하나라도 있으면 납품하지 않음)", "출처: 각 시트 링크", label_w=56, data_w=12, ncols=3)
    ck.band("모델 점검", 3); ck.header(["점검", "값", "결과"])
    mid = sens[(0.0, 0.0)]
    items = [("영구성장률 < WACC", "=%s-%s" % (P["wacc"], P["g"]), "C{r}>0", PCT),
             ("영구가치 비중 < 85%", "=DCF!C%d" % tvs, "C{r}<0.85", PCT),
             ("케이스 스위치가 1~3", "=" + SW, "AND(C{r}>=1,C{r}<=3)", NUM),
             ("민감도 가운데 칸 = 주당가치 (차이)", "=DCF!%s%d-DCF!C%d" % (L(mid[1]), mid[0], ps), "ABS(C{r})<0.5", NUM1),
             ("EV = 추정기간 PV + 영구가치 PV (차이)", "=DCF!C%d-DCF!C%d-DCF!C%d" % (ev, s_pv, pv_tv), "ABS(C{r})<0.5", NUM1),
             ("FCF 가 전 기간 양수 (최솟값)", "=MIN(DCF!C%d:G%d)" % (f_, f_), "C{r}>0", NUM),
             ("부문 비중 합 = 100% (차이)", "=%s+%s-1" % (P["plat_share"], P["sol_share"]), "ABS(C{r})<0.000001", PCT),
             ("SOTP 주주가치 = 영업 EV + 비영업 - 순차입금 - 비지배 (차이)",
              "=SOTP!C%d-(SOTP!G%d+%s*(1-%s)-%s-%s)" % (s_eq, st, P["assoc_book"], P["holdco_disc"], P["net_debt"], P["nci"]),
              "ABS(C{r})<0.5", NUM1),
             ("SOTP 배수 > 0 (최솟값)", "=MIN(SOTP!F%d:F%d)" % (sp, ss), "C{r}>0", MULT)]
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
    r_sotp = cv.row("SOTP 주당가치 (원, 참고)", ["=SOTP!C%d" % s_ps], "link", NUM)
    cv.ws.cell(r_sotp, 4, "부문별 배수 합산 (SOTP 시트)").font = font(color=GREY)

    wb.calculation.fullCalcOnLoad = True
    wb.save(path)
    return {"per_share": "C%d" % cv_rows["per_share"], "check": "C%d" % cv_rows["check"], "switch": "C6",
            "ev": "C%d" % ev, "sotp_per_share": "C%d" % r_sotp}


def main():
    os.makedirs(OUT, exist_ok=True)
    build_dataset(os.path.join(OUT, "sample_dataset_ib.xlsx"))
    cells = build_model(os.path.join(OUT, "sample_model_dcf.xlsx"))
    print("built 2 files", cells)


if __name__ == "__main__":
    main()
