"""누적 금액을 분기(3개월) 값으로 바꾼다.

공시 API 의 기간 기준은 재무제표마다 다르다.
  - 재무상태표: 기말 잔액. 그대로 쓴다.
  - 손익: thstrm_amount 는 3개월치, thstrm_add_amount 는 누적. 사업보고서는 연간값만 온다.
  - 현금흐름: thstrm_amount 자체가 누적이다.
그래서 '누적값'을 하나로 정한 뒤(add 가 있으면 add, 없으면 thstrm) 같은 회계연도 안에서
직전 분기 누적을 빼 3개월 값을 만든다. 4분기는 '연간 - 3분기 누적'이 된다.

직전 분기 보고서가 없으면 차분이 두 분기치가 되므로 값을 비우고 flag 를 남긴다.
"""

from __future__ import annotations

import numpy as np
import pandas as pd

from finprep.schema import ANNUAL_REPRT, ENTITY, REPRT_TO_QN, STATEMENT_KIND

GAP_FLAG = "직전 분기 보고서 누락"
_SERIES = [*ENTITY, "sj_div", "account_key", "bsns_year"]


def _with_period(df: pd.DataFrame) -> pd.DataFrame:
    out = df.copy()
    out["reprt_code"] = out["reprt_code"].astype(str)
    out["qn"] = out["reprt_code"].map(REPRT_TO_QN)
    if out["qn"].isna().any():
        bad = sorted(out.loc[out["qn"].isna(), "reprt_code"].unique())
        raise ValueError(f"알 수 없는 보고서 코드: {bad}")
    out["qn"] = out["qn"].astype(int)
    out["bsns_year"] = out["bsns_year"].astype(int)
    out["kind"] = out["sj_div"].map(STATEMENT_KIND)
    if out["kind"].isna().any():
        bad = sorted(out.loc[out["kind"].isna(), "sj_div"].unique())
        raise ValueError(f"알 수 없는 재무제표 구분: {bad}")
    return out


def cumulative(df: pd.DataFrame) -> pd.Series:
    """기초 대비 누적액. 누적 필드가 있으면 그것을, 없으면 당기 금액을 쓴다."""
    return df["amount_cum"].where(df["amount_cum"].notna(), df["amount"])


def to_quarterly(df: pd.DataFrame) -> pd.DataFrame:
    """분기 값 컬럼 value 와 period('2025.3Q')를 붙인다. 입력은 parse·key 단계를 거친 것."""
    out = _with_period(df)
    out = out.sort_values([*_SERIES, "qn"], kind="stable").reset_index(drop=True)
    flow = out["kind"] == "flow"

    out["cum"] = np.where(flow, cumulative(out), np.nan)
    grouped = out.groupby(_SERIES, sort=False)
    prev_cum = grouped["cum"].shift()
    prev_qn = grouped["qn"].shift()

    first = out["qn"] == 1
    consecutive = prev_qn == out["qn"] - 1
    diffed = np.where(first, out["cum"], np.where(consecutive, out["cum"] - prev_cum, np.nan))
    out["value"] = np.where(flow, diffed, out["amount"])

    gap = flow & ~first & ~consecutive & out["cum"].notna()
    if "flag" not in out.columns:
        out["flag"] = ""
    out.loc[gap, "flag"] = GAP_FLAG
    out["period"] = out["bsns_year"].astype(str) + "." + out["qn"].astype(str) + "Q"
    return out


def to_annual(df: pd.DataFrame) -> pd.DataFrame:
    """사업보고서 행만 골라 연간 값으로. 손익·현금흐름은 연간 합계, 재무상태표는 연말 잔액."""
    out = _with_period(df)
    out = out[out["reprt_code"] == ANNUAL_REPRT].copy()
    out["value"] = out["amount"]
    out["period"] = out["bsns_year"].astype(str)
    if "flag" not in out.columns:
        out["flag"] = ""
    return out.reset_index(drop=True)
