"""단위 환산. 주당 금액과 주식수는 금액 단위로 나누지 않는다."""

from __future__ import annotations

import pandas as pd

from finprep.schema import UNIT_DIV, UNIT_TO_WON

PER_SHARE_PATTERN = r"주당|주식수|^EPS$"


def is_per_share(df: pd.DataFrame) -> pd.Series:
    label = df["account_label"].str.contains(PER_SHARE_PATTERN, na=False)
    if "std_account" in df.columns:
        label |= df["std_account"].fillna("").str.contains(PER_SHARE_PATTERN)
    return label


def scale_values(df: pd.DataFrame, unit: str, column: str = "value") -> pd.DataFrame:
    """원 단위 column 을 unit 으로 환산하고 행마다 단위를 적는다."""
    if unit not in UNIT_DIV:
        raise ValueError(f"지원하지 않는 단위: {unit!r} (가능: {', '.join(UNIT_DIV)})")
    out = df.copy()
    keep = is_per_share(out)
    out[column] = out[column].astype("float64")
    out.loc[~keep, column] = out.loc[~keep, column] / UNIT_DIV[unit]
    out["unit"] = unit
    out.loc[keep, "unit"] = "원"
    return out


def convert(amount: float, from_unit: str, to_unit: str) -> float:
    """출처마다 다른 단위(억원, 십억원 등)를 맞춘다. 모르는 단위는 추측하지 않고 멈춘다."""
    for unit in (from_unit, to_unit):
        if unit not in UNIT_TO_WON:
            raise ValueError(f"지원하지 않는 단위: {unit!r}")
    return amount * UNIT_TO_WON[from_unit] / UNIT_TO_WON[to_unit]
