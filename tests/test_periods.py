import math

import pandas as pd
import pytest

from finprep.clean.parse import add_account_keys, parse_amounts
from finprep.clean.periods import GAP_FLAG, to_annual, to_quarterly

from .conftest import flow_year, frame, row


def _prepare(df: pd.DataFrame) -> pd.DataFrame:
    return add_account_keys(parse_amounts(df))


def _values(df: pd.DataFrame, name: str) -> list[float]:
    sub = df[df["account_nm"] == name].sort_values("qn")
    return list(sub["value"])


def test_income_q4_is_annual_minus_three_quarters():
    raw = frame(*flow_year("CIS", "매출액", [100, 120, 110, 170], ytd_thstrm=False))
    q = to_quarterly(_prepare(raw))
    assert _values(q, "매출액") == [100, 120, 110, 170]


def test_cash_flow_ytd_is_differenced():
    raw = frame(*flow_year("CF", "영업활동현금흐름", [30, 50, -10, 40], ytd_thstrm=True))
    q = to_quarterly(_prepare(raw))
    assert _values(q, "영업활동현금흐름") == [30, 50, -10, 40]


def test_balance_sheet_is_not_differenced():
    raw = frame(
        *[
            row("BS", "자산총계", v, reprt=r)
            for v, r in zip(
                [900, 950, 990, 1000], ["11013", "11012", "11014", "11011"], strict=True
            )
        ]
    )
    q = to_quarterly(_prepare(raw))
    assert _values(q, "자산총계") == [900, 950, 990, 1000]


def test_difference_does_not_cross_fiscal_years():
    raw = frame(
        *flow_year("CF", "영업활동현금흐름", [10, 10, 10, 10], ytd_thstrm=True, year=2024),
        *flow_year("CF", "영업활동현금흐름", [20, 20, 20, 20], ytd_thstrm=True, year=2025),
    )
    q = to_quarterly(_prepare(raw))
    q25 = q[q["bsns_year"] == 2025]
    assert _values(q25, "영업활동현금흐름") == [20, 20, 20, 20]


def test_missing_previous_report_leaves_blank_with_flag():
    rows = flow_year("CF", "영업활동현금흐름", [30, 50, -10, 40], ytd_thstrm=True)
    raw = frame(*[r for r in rows if r["reprt_code"] != "11012"])  # 반기 누락
    q = to_quarterly(_prepare(raw)).sort_values("qn")
    q3 = q[q["qn"] == 3].iloc[0]
    # 3분기 누적 70 에서 1분기 30 을 빼면 두 분기치 40 이 된다. 틀린 값 대신 빈 값을 낸다.
    assert math.isnan(q3["value"])
    assert q3["flag"] == GAP_FLAG
    assert q[q["qn"] == 4].iloc[0]["value"] == 40


def test_annual_takes_business_report_only():
    raw = frame(*flow_year("CIS", "매출액", [100, 120, 110, 170], ytd_thstrm=False))
    annual = to_annual(_prepare(raw))
    assert list(annual["value"]) == [500]
    assert list(annual["period"]) == ["2025"]


def test_unknown_report_code_raises():
    raw = frame(row("CIS", "매출액", 1, reprt="99999"))
    with pytest.raises(ValueError, match="보고서 코드"):
        to_quarterly(_prepare(raw))
