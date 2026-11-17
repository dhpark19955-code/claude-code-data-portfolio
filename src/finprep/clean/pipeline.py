"""원자료 -> 분기 tidy 데이터. 규칙을 정해진 순서로 건다.

    parse -> 계정 키 -> 표준 계정 -> 현금흐름 부호 -> 누적 차분 -> 단위

단계마다 '몇 행을 바꿨고 몇 행을 못 바꿨는지'를 quality 에 남긴다. 못 한 것은 warnings 로 따로 낸다.
"""

from __future__ import annotations

from dataclasses import dataclass, field

import pandas as pd

from finprep.clean.aliases import required_coverage, standardize_accounts, unmapped_borrowings
from finprep.clean.parse import add_account_keys, parse_amounts
from finprep.clean.periods import GAP_FLAG, to_annual, to_quarterly
from finprep.clean.signs import apply_cf_signs
from finprep.clean.units import scale_values
from finprep.schema import ENTITY, RAW_REQUIRED, REPORT

TIDY_COLUMNS = [
    "corp_code",
    "corp_name",
    "fs_div",
    "sj_div",
    "account_key",
    "account_label",
    "std_account",
    "bsns_year",
    "qn",
    "period",
    "value",
    "unit",
    "reported_3m",
    "sign_flipped",
    "flag",
]


@dataclass
class CleanResult:
    quarterly: pd.DataFrame
    annual: pd.DataFrame
    unit: str
    quality: dict = field(default_factory=dict)
    warnings: list[str] = field(default_factory=list)


def validate_raw(raw: pd.DataFrame) -> None:
    missing = [c for c in RAW_REQUIRED if c not in raw.columns]
    if missing:
        raise ValueError(f"원자료에 필요한 컬럼이 없습니다: {missing}")
    if raw.empty:
        raise ValueError("원자료가 비어 있습니다")


def _finish(df: pd.DataFrame, unit: str) -> pd.DataFrame:
    out = scale_values(df, unit, "value")
    out = scale_values(out, unit, "reported_3m")
    return out[TIDY_COLUMNS].reset_index(drop=True)


def clean(raw: pd.DataFrame, unit: str = "백만원") -> CleanResult:
    """공시 API 원자료(long)를 분기·연간 tidy 데이터로 정리한다."""
    validate_raw(raw)
    df = raw.copy()
    for column in ("corp_code", "reprt_code", "fs_div", "sj_div"):
        df[column] = df[column].astype(str)
    if "corp_name" not in df.columns:
        df["corp_name"] = df["corp_code"]

    df = parse_amounts(df)
    df = add_account_keys(df)
    df = standardize_accounts(df)
    df = apply_cf_signs(df)

    quarterly = to_quarterly(df)
    # 손익의 3개월치 필드는 누적 차분 결과와 교차검증하는 데 쓴다 (누적 필드가 함께 온 행만)
    has_both = quarterly["sj_div"].isin(["IS", "CIS"]) & quarterly["amount_cum"].notna()
    quarterly["reported_3m"] = quarterly["amount"].where(has_both)
    annual = to_annual(df)
    annual["reported_3m"] = float("nan")

    quality, warnings = _quality(raw, df, quarterly)
    return CleanResult(
        quarterly=_finish(quarterly, unit),
        annual=_finish(annual, unit),
        unit=unit,
        quality=quality,
        warnings=warnings,
    )


def _quality(
    raw: pd.DataFrame, df: pd.DataFrame, quarterly: pd.DataFrame
) -> tuple[dict, list[str]]:
    coverage = required_coverage(df)
    unmapped = unmapped_borrowings(df)
    gaps = quarterly[quarterly["flag"] == GAP_FLAG]
    flow = quarterly["kind"] == "flow"
    names = df.drop_duplicates("corp_code").set_index("corp_code")["corp_name"]
    convention = (
        df.drop_duplicates(ENTITY).set_index("corp_code")["outflow_convention"].fillna("해당 없음")
    )

    quality = {
        "raw_rows": len(raw),
        "companies": df["corp_code"].nunique(),
        "reports": len(df[REPORT].drop_duplicates()),
        "amount_missing": int(df["amount"].isna().sum()),
        "no_std_id_rows": int((~df["has_std_id"]).sum()),
        "duplicate_candidates": int(df["flag"].str.startswith("표준계정 중복").sum()),
        "std_mapped_rows": int(df["std_account"].notna().sum()),
        "required_found": int(coverage["found"].sum()),
        "required_total": int(coverage["required"].sum()),
        "sign_flipped_rows": int(df["sign_flipped"].sum()),
        "outflow_convention": {names[c]: v for c, v in convention.items()},
        "diffed_rows": int((flow & (quarterly["qn"] > 1) & quarterly["value"].notna()).sum()),
        "gap_rows": len(gaps),
        "quarterly_rows": len(quarterly),
        "quarterly_values": int(quarterly["value"].notna().sum()),
    }

    warnings = []
    for (corp, period), n in gaps.groupby(["corp_code", "period"]).size().items():
        warnings.append(f"{names[corp]} {period}: 직전 분기 보고서가 없어 분기값 {n}행을 비워 둠")
    for corp, labels in unmapped.groupby("corp_code")["account_label"]:
        joined = ", ".join(sorted(labels))
        warnings.append(f"{names[corp]}: 차입금 별칭 누락 의심 - {joined}")
    short = coverage[coverage["found"] < coverage["required"]]
    for corp, n in short.groupby("corp_code").size().items():
        warnings.append(f"{names[corp]}: 필수 계정을 다 찾지 못한 보고서 {n}건")
    return quality, warnings
