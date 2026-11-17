"""현금흐름 유출 계정의 부호 통일.

공시 현금흐름표는 '유형자산의 취득' 같은 유출 항목을 양수로 적는 회사와 음수로 적는 회사가
섞여 있다. 그대로 합치면 CAPEX 가 현금 유입으로 계산된다.

규칙
  1) 계정명으로 유출 성격(취득, 지급, 상환, 납부)을 판단한다.
  2) 회사별로 유출 계정이 주로 양수인지 본다 (표기 관행 감지).
  3) 양수 관행 회사만 유출 계정에 -1 을 곱한다.

abs() 를 쓰지 않는 이유: -abs(x) 로 밀어 버리면 환입·환급처럼 원래 반대 부호인 값의 방향이
사라진다. -1 을 곱하면 관행만 뒤집고 방향 정보는 남는다.
"""

from __future__ import annotations

import re

import pandas as pd

from finprep.schema import ENTITY

OUTFLOW_PATTERN = re.compile(r"(취득|지급|상환|납부)$")


def is_outflow_account(label: object) -> bool:
    base = re.sub(r"#\d+$", "", str(label))
    return bool(OUTFLOW_PATTERN.search(base))


def detect_outflow_convention(df: pd.DataFrame) -> pd.Series:
    """회사(corp_code, fs_div)별 유출 계정 표기 관행: 'positive' 또는 'negative'."""
    cf = df[(df["sj_div"] == "CF") & df["account_label"].map(is_outflow_account)]
    cf = cf[cf["amount"].notna() & (cf["amount"] != 0)]
    if cf.empty:
        return pd.Series(dtype=object, name="outflow_convention")
    share_positive = (cf["amount"] > 0).groupby([cf[c] for c in ENTITY]).mean()
    return share_positive.map(lambda s: "positive" if s > 0.5 else "negative").rename(
        "outflow_convention"
    )


def apply_cf_signs(df: pd.DataFrame) -> pd.DataFrame:
    """양수 관행 회사의 유출 계정에 -1 을 곱한다. 바꾼 행은 sign_flipped=True."""
    out = df.copy()
    convention = detect_outflow_convention(out)
    out = out.join(convention, on=ENTITY)
    flip = (
        (out["sj_div"] == "CF")
        & out["account_label"].map(is_outflow_account)
        & (out["outflow_convention"] == "positive")
    )
    out.loc[flip, ["amount", "amount_cum"]] *= -1
    out["sign_flipped"] = flip
    return out
