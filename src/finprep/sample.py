"""가상 회사 원자료 생성기. 공시 API(fnlttSinglAcntAll) 응답과 같은 모양으로 만든다.

회사와 숫자는 전부 가상이다. 대신 실제 공시에서 반복해서 만난 '지저분함'을 회사마다 심었다.

  가온테크     기준선. 표준계정코드 사용, 유출 계정을 양수로 표기
  누리소프트   차입금 계정명 변형, 2025년부터 '영업이익(손실)'로 표기 변경, 일부 코드 미사용
  다온바이오   적자, 유출 계정을 음수로 표기(반대 관행), 유동·비유동 모두 '차입금' (동명 계정)
  라온리테일   손익계산서와 포괄손익계산서 분리, 로마숫자 머리표, '자산 총계' 띄어쓰기
  마루에너지   2024년 반기보고서 누락, 별칭 표에 없는 차입금 계정명

같은 seed 면 항상 같은 파일이 나온다.
"""

from __future__ import annotations

import random
from dataclasses import dataclass, field

import pandas as pd

from finprep.schema import NO_STD

YEARS = (2023, 2024, 2025)
REPORTS = ("11013", "11012", "11014", "11011")
SEASONALITY = (0.22, 0.25, 0.24, 0.29)


@dataclass(frozen=True)
class Profile:
    corp_code: str  # 가상 코드 (9 로 시작, 실제 고유번호와 겹치지 않게)
    corp_name: str
    stock_code: str
    revenue: float  # 2023년 연 매출 (원)
    growth: float
    cogs: float
    sga: float
    leverage: float
    nci: float = 0.0
    outflow_positive: bool = True
    separate_is: bool = False
    roman_prefix: bool = False
    bond: bool = False
    labels: dict[str, str] = field(default_factory=dict)  # 표준 표기 -> 이 회사 표기
    labels_from: dict[int, dict[str, str]] = field(default_factory=dict)  # 연도부터 바뀌는 표기
    no_std_id: tuple[str, ...] = ()
    skip_reports: tuple[tuple[int, str], ...] = ()
    split_short_debt: str | None = None  # 단기차입금 일부를 이 이름의 계정으로 따로 표기
    segments: tuple[tuple[str, float], ...] = ()
    segment_unit: str = "억원"


PROFILES: tuple[Profile, ...] = (
    Profile(
        "90000001",
        "가온테크",
        "900010",
        212e9,
        1.11,
        0.62,
        0.22,
        0.38,
        nci=0.04,
        segments=(("플랫폼", 0.58), ("솔루션", 0.42)),
        segment_unit="억원",
    ),
    Profile(
        "90000002",
        "누리소프트",
        "900020",
        96e9,
        1.18,
        0.41,
        0.43,
        0.31,
        labels={
            "단기차입금": "유동 차입금(사채 포함)",
            "장기차입금": "비유동 차입금(사채 포함)",
            "매출액": "영업수익",
            "매출원가": "영업비용",
        },
        labels_from={2025: {"영업이익": "영업이익(손실)", "당기순이익": "당기순이익(손실)"}},
        no_std_id=("단기차입금", "장기차입금", "판매비와관리비"),
        segments=(("구독", 0.71), ("용역", 0.29)),
        segment_unit="십억원",
    ),
    Profile(
        "90000003",
        "다온바이오",
        "900030",
        38e9,
        1.24,
        0.55,
        0.61,
        0.44,
        outflow_positive=False,
        labels={
            "단기차입금": "차입금",
            "장기차입금": "차입금",
            "영업이익": "영업이익(손실)",
            "당기순이익": "당기순이익(손실)",
        },
        no_std_id=("단기차입금", "장기차입금"),
    ),
    Profile(
        "90000004",
        "라온리테일",
        "900040",
        540e9,
        1.04,
        0.74,
        0.21,
        0.52,
        nci=0.09,
        separate_is=True,
        roman_prefix=True,
        bond=True,
        labels={
            "자산총계": "자산 총계",
            "부채총계": "부채 총계",
            "자본총계": "자본 총계",
            "지배주주지분": "지배기업의 소유주에게 귀속되는 자본",
        },
        segments=(("오프라인", 0.64), ("온라인", 0.27), ("기타", 0.09)),
        segment_unit="백만원",
    ),
    Profile(
        "90000005",
        "마루에너지",
        "900050",
        318e9,
        1.07,
        0.81,
        0.09,
        0.47,
        skip_reports=((2024, "11012"),),
        split_short_debt="유동성 차입금 및 사채",
        labels={
            "영업활동현금흐름": "영업활동으로 인한 현금흐름",
            "투자활동현금흐름": "투자활동으로 인한 현금흐름",
            "재무활동현금흐름": "재무활동으로 인한 현금흐름",
        },
    ),
)

ACCOUNT_IDS = {
    "유동자산": "ifrs-full_CurrentAssets",
    "현금및현금성자산": "ifrs-full_CashAndCashEquivalents",
    "매출채권및기타채권": "ifrs-full_TradeAndOtherCurrentReceivables",
    "재고자산": "ifrs-full_Inventories",
    "비유동자산": "ifrs-full_NoncurrentAssets",
    "유형자산": "ifrs-full_PropertyPlantAndEquipment",
    "무형자산": "ifrs-full_IntangibleAssetsOtherThanGoodwill",
    "자산총계": "ifrs-full_Assets",
    "유동부채": "ifrs-full_CurrentLiabilities",
    "단기차입금": "ifrs-full_ShorttermBorrowings",
    "기타유동부채": "dart_OtherCurrentLiabilities",
    "비유동부채": "ifrs-full_NoncurrentLiabilities",
    "사채": "dart_BondsIssued",
    "장기차입금": "dart_LongTermBorrowingsGross",
    "기타비유동부채": "dart_OtherNonCurrentLiabilities",
    "부채총계": "ifrs-full_Liabilities",
    "지배주주지분": "ifrs-full_EquityAttributableToOwnersOfParent",
    "비지배지분": "ifrs-full_NoncontrollingInterests",
    "자본총계": "ifrs-full_Equity",
    "매출액": "ifrs-full_Revenue",
    "매출원가": "ifrs-full_CostOfSales",
    "매출총이익": "ifrs-full_GrossProfit",
    "판매비와관리비": "dart_TotalSellingGeneralAdministrativeExpenses",
    "영업이익": "dart_OperatingIncomeLoss",
    "금융수익": "ifrs-full_FinanceIncome",
    "금융비용": "ifrs-full_FinanceCosts",
    "법인세비용차감전순이익": "ifrs-full_ProfitLossBeforeTax",
    "법인세비용": "ifrs-full_IncomeTaxExpenseContinuingOperations",
    "당기순이익": "ifrs-full_ProfitLoss",
    "지배기업의 소유주에게 귀속되는 당기순이익": "ifrs-full_ProfitLossAttributableToOwnersOfParent",
    "비지배지분순이익": "ifrs-full_ProfitLossAttributableToNoncontrollingInterests",
    "기본주당이익": "ifrs-full_BasicEarningsLossPerShare",
    "기타포괄손익": "ifrs-full_OtherComprehensiveIncome",
    "총포괄손익": "ifrs-full_ComprehensiveIncome",
    "영업활동현금흐름": "ifrs-full_CashFlowsFromUsedInOperatingActivities",
    "감가상각비": "ifrs-full_AdjustmentsForDepreciationExpense",
    "대손상각비(환입)": "dart_AdjustmentsForBadDebtExpenses",
    "이자의 지급": "ifrs-full_InterestPaidClassifiedAsOperatingActivities",
    "투자활동현금흐름": "ifrs-full_CashFlowsFromUsedInInvestingActivities",
    "유형자산의 취득": "ifrs-full_PurchaseOfPropertyPlantAndEquipmentClassifiedAsInvestingActivities",
    "무형자산의 취득": "ifrs-full_PurchaseOfIntangibleAssetsClassifiedAsInvestingActivities",
    "유형자산의 처분": "ifrs-full_ProceedsFromSalesOfPropertyPlantAndEquipmentClassifiedAsInvestingActivities",
    "재무활동현금흐름": "ifrs-full_CashFlowsFromUsedInFinancingActivities",
    "단기차입금의 차입": "dart_ProceedsFromShortTermBorrowings",
    "장기차입금의 상환": "dart_RepaymentsOfLongTermBorrowings",
    "배당금의 지급": "ifrs-full_DividendsPaidClassifiedAsFinancingActivities",
}

SJ_NAMES = {"BS": "재무상태표", "IS": "손익계산서", "CIS": "포괄손익계산서", "CF": "현금흐름표"}
ROMAN = {"유동자산": "Ⅰ. ", "비유동자산": "Ⅱ. ", "유동부채": "Ⅰ. ", "비유동부채": "Ⅱ. "}
OUTFLOWS = {
    "이자의 지급",
    "유형자산의 취득",
    "무형자산의 취득",
    "장기차입금의 상환",
    "배당금의 지급",
}


def _won(value: float) -> int:
    return int(round(value / 1000.0)) * 1000  # 천원 단위로 끊어 실제 공시처럼 보이게


def _quarter_flows(rng: random.Random, p: Profile, year: int, q: int) -> dict[str, int]:
    """한 분기의 3개월 손익·현금흐름 (참값, 부호는 경제적 의미 그대로)."""
    scale = p.revenue * p.growth ** (year - YEARS[0]) * SEASONALITY[q]
    revenue = _won(scale * rng.uniform(0.96, 1.04))
    cogs = _won(revenue * p.cogs * rng.uniform(0.98, 1.02))
    sga = _won(revenue * p.sga * rng.uniform(0.97, 1.03))
    operating = revenue - cogs - sga
    fin_income = _won(revenue * 0.004)
    fin_cost = _won(revenue * 0.02 * p.leverage)
    pretax = operating + fin_income - fin_cost
    tax = _won(max(pretax, 0) * 0.22)
    net = pretax - tax
    nci = _won(net * p.nci)
    oci = _won(revenue * rng.uniform(-0.003, 0.003))

    depreciation = _won(revenue * 0.035)
    bad_debt = _won(revenue * rng.uniform(-0.002, 0.003))  # 음수면 환입
    interest_paid = fin_cost
    capex = _won(revenue * rng.uniform(0.04, 0.07))
    intangibles = _won(revenue * 0.008)
    disposal = _won(revenue * rng.uniform(0.0, 0.004))
    borrow = _won(revenue * rng.uniform(0.0, 0.03))
    repay = _won(revenue * rng.uniform(0.0, 0.02))
    dividend = _won(max(net, 0) * 0.3) if q == 1 else 0
    return {
        "매출액": revenue,
        "매출원가": cogs,
        "매출총이익": revenue - cogs,
        "판매비와관리비": sga,
        "영업이익": operating,
        "금융수익": fin_income,
        "금융비용": fin_cost,
        "법인세비용차감전순이익": pretax,
        "법인세비용": tax,
        "당기순이익": net,
        "지배기업의 소유주에게 귀속되는 당기순이익": net - nci,
        "비지배지분순이익": nci,
        "기본주당이익": int(round((net - nci) / 20_000_000)),
        "기타포괄손익": oci,
        "총포괄손익": net + oci,
        "감가상각비": depreciation,
        "대손상각비(환입)": bad_debt,
        "이자의 지급": -interest_paid,
        "영업활동현금흐름": net + depreciation + bad_debt - interest_paid,
        "유형자산의 취득": -capex,
        "무형자산의 취득": -intangibles,
        "유형자산의 처분": disposal,
        "투자활동현금흐름": disposal - capex - intangibles,
        "단기차입금의 차입": borrow,
        "장기차입금의 상환": -repay,
        "배당금의 지급": -dividend,
        "재무활동현금흐름": borrow - repay - dividend,
    }


def _balance(rng: random.Random, p: Profile, revenue_q: int) -> dict[str, int]:
    """분기말 재무상태표. 합계는 구성 항목에서 계산해 항등식이 원 단위까지 맞는다."""
    cash = _won(revenue_q * rng.uniform(0.5, 0.9))
    receivables = _won(revenue_q * 0.62)
    inventory = _won(revenue_q * 0.35)
    ppe = _won(revenue_q * rng.uniform(2.2, 2.6))
    intangible = _won(revenue_q * 0.3)
    assets = cash + receivables + inventory + ppe + intangible
    liabilities = _won(assets * p.leverage * rng.uniform(0.97, 1.03))
    short_debt = _won(liabilities * 0.24)
    bond = _won(liabilities * 0.12) if p.bond else 0
    long_debt = _won(liabilities * 0.28)
    other_current = _won(liabilities * 0.26)
    other_noncurrent = liabilities - short_debt - bond - long_debt - other_current
    equity = assets - liabilities
    nci = _won(equity * p.nci)
    return {
        "유동자산": cash + receivables + inventory,
        "현금및현금성자산": cash,
        "매출채권및기타채권": receivables,
        "재고자산": inventory,
        "비유동자산": ppe + intangible,
        "유형자산": ppe,
        "무형자산": intangible,
        "자산총계": assets,
        "유동부채": short_debt + other_current,
        "단기차입금": short_debt,
        "기타유동부채": other_current,
        "비유동부채": bond + long_debt + other_noncurrent,
        "사채": bond,
        "장기차입금": long_debt,
        "기타비유동부채": other_noncurrent,
        "부채총계": liabilities,
        "지배주주지분": equity - nci,
        "비지배지분": nci,
        "자본총계": equity,
    }


BS_ORDER = [
    "유동자산",
    "현금및현금성자산",
    "매출채권및기타채권",
    "재고자산",
    "비유동자산",
    "유형자산",
    "무형자산",
    "자산총계",
    "유동부채",
    "단기차입금",
    "기타유동부채",
    "비유동부채",
    "사채",
    "장기차입금",
    "기타비유동부채",
    "부채총계",
    "지배주주지분",
    "비지배지분",
    "자본총계",
]
PL_ORDER = [
    "매출액",
    "매출원가",
    "매출총이익",
    "판매비와관리비",
    "영업이익",
    "금융수익",
    "금융비용",
    "법인세비용차감전순이익",
    "법인세비용",
    "당기순이익",
    "지배기업의 소유주에게 귀속되는 당기순이익",
    "비지배지분순이익",
    "기본주당이익",
]
OCI_ORDER = ["기타포괄손익", "총포괄손익"]
CF_ORDER = [
    "영업활동현금흐름",
    "당기순이익",
    "감가상각비",
    "대손상각비(환입)",
    "이자의 지급",
    "투자활동현금흐름",
    "유형자산의 취득",
    "무형자산의 취득",
    "유형자산의 처분",
    "재무활동현금흐름",
    "단기차입금의 차입",
    "장기차입금의 상환",
    "배당금의 지급",
]


def _label(p: Profile, account: str, year: int) -> str:
    label = p.labels.get(account, account)
    for start, changes in p.labels_from.items():
        if year >= start:
            label = changes.get(account, label)
    if account == "비지배지분순이익":
        label = "비지배지분"
    if p.roman_prefix and account in ROMAN:
        label = ROMAN[account] + label
    return label


def _fmt(value: int | None) -> str:
    return "" if value is None else str(value)


def _company_rows(rng: random.Random, p: Profile) -> list[dict]:
    rows: list[dict] = []
    for year in YEARS:
        flows = [_quarter_flows(rng, p, year, q) for q in range(4)]
        for q, reprt in enumerate(REPORTS):
            cumulative = {k: sum(f[k] for f in flows[: q + 1]) for k in flows[0]}
            balance = _balance(rng, p, flows[q]["매출액"])  # 누락 보고서도 난수는 소비한다
            if (year, reprt) in p.skip_reports:
                continue
            annual = reprt == "11011"

            def add(sj, account, amount, add_amount=None, order=0, *, year=year, reprt=reprt):
                std_id = ACCOUNT_IDS.get(account, NO_STD)
                rows.append(
                    {
                        "corp_code": p.corp_code,
                        "corp_name": p.corp_name,
                        "stock_code": p.stock_code,
                        "bsns_year": year,
                        "reprt_code": reprt,
                        "fs_div": "CFS",
                        "sj_div": sj,
                        "sj_nm": SJ_NAMES[sj],
                        "account_id": NO_STD if account in p.no_std_id else std_id,
                        "account_nm": _label(p, account, year),
                        "thstrm_amount": _fmt(amount),
                        "thstrm_add_amount": _fmt(add_amount),
                        "ord": order,
                        "currency": "KRW",
                    }
                )

            order = 0
            for account in BS_ORDER:
                amount = balance[account]
                if account in ("사채", "비지배지분") and amount == 0:
                    continue
                if account == "단기차입금" and p.split_short_debt:
                    part = _won(amount * 0.4)
                    order += 1
                    add("BS", account, amount - part, order=order)
                    order += 1
                    rows.append(
                        {
                            **rows[-1],
                            "account_id": NO_STD,
                            "account_nm": p.split_short_debt,
                            "thstrm_amount": _fmt(part),
                            "ord": order,
                        }
                    )
                    continue
                order += 1
                add("BS", account, amount, order=order)

            pl_sj = "IS" if p.separate_is else "CIS"
            order = 0
            for account in PL_ORDER:
                if account == "비지배지분순이익" and p.nci == 0:
                    continue
                order += 1
                if annual:
                    add(pl_sj, account, cumulative[account], order=order)
                else:
                    add(pl_sj, account, flows[q][account], cumulative[account], order=order)
            if p.separate_is:  # 포괄손익계산서는 당기순이익에서 다시 시작한다
                order = 0
                oci_accounts = ["당기순이익", *OCI_ORDER]
            else:
                oci_accounts = OCI_ORDER
            for account in oci_accounts:
                order += 1
                if annual:
                    add("CIS", account, cumulative[account], order=order)
                else:
                    add("CIS", account, flows[q][account], cumulative[account], order=order)

            for order, account in enumerate(CF_ORDER, start=1):
                amount = cumulative[account]
                if account in OUTFLOWS and p.outflow_positive:
                    amount = -amount  # 유출을 양수로 적는 관행
                if account == "배당금의 지급" and amount == 0:
                    add("CF", account, None, order=order)  # 빈 문자열로 오는 칸
                    continue
                add("CF", account, amount, order=order)
    return rows


def generate(seed: int = 20261117) -> pd.DataFrame:
    """가상 5개사, 3개년, 분기·사업보고서 원자료 (long). 금액은 API 처럼 문자열이다."""
    rng = random.Random(seed)
    rows: list[dict] = []
    for profile in PROFILES:
        rows += _company_rows(rng, profile)
    return pd.DataFrame(rows)


def generate_segments(raw: pd.DataFrame | None = None, seed: int = 20261117) -> pd.DataFrame:
    """IR 자료에서 옮긴 것 같은 부문 매출. 회사마다 단위가 다르다 (억원, 십억원, 백만원)."""
    from finprep.schema import UNIT_TO_WON

    raw = generate(seed) if raw is None else raw
    annual = raw[(raw["reprt_code"] == "11011") & (raw["account_id"] == ACCOUNT_IDS["매출액"])]
    revenue = annual.set_index(["corp_code", "bsns_year"])["thstrm_amount"].astype(float)
    rows = []
    for p in PROFILES:
        for year in YEARS:
            for name, weight in p.segments:
                amount = revenue[(p.corp_code, year)] * weight / UNIT_TO_WON[p.segment_unit]
                rows.append(
                    {
                        "corp_code": p.corp_code,
                        "corp_name": p.corp_name,
                        "bsns_year": year,
                        "segment": name,
                        "revenue": round(amount, 1),
                        "unit": p.segment_unit,
                    }
                )
    return pd.DataFrame(rows)
