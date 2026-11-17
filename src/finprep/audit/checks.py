"""기계 감사. 숫자가 서로 맞는지를 회사별로 점검한다.

점검은 '통과 / 실패 / 경고' 셋 중 하나를 낸다. 실패는 데이터나 전처리가 틀렸다는 뜻이고,
경고는 자동으로 판단할 수 없어 사람이 봐야 한다는 뜻이다.
원본 `originals/pipeline/dataset_audit.py` 의 항등식·범위 점검을 tidy 데이터용으로 옮겼다.
"""

from __future__ import annotations

from dataclasses import dataclass

import pandas as pd

from finprep.clean.aliases import BORROWING_ACCOUNTS
from finprep.clean.periods import GAP_FLAG
from finprep.clean.pipeline import CleanResult
from finprep.clean.units import convert

PASS, FAIL, WARN = "통과", "실패", "경고"

RATIO_BOUNDS = {"영업이익률": (-3.0, 1.0), "부채비율": (0.0, 10.0), "현금/자산": (0.0, 1.0)}


@dataclass(frozen=True)
class CheckResult:
    check: str
    corp_name: str
    status: str
    n_checked: int
    detail: str = ""


def _wide(df: pd.DataFrame) -> pd.DataFrame:
    """표준 계정을 열로 펼친다. 행은 (회사, 기간)."""
    std = df[df["std_account"].notna()]
    return std.pivot_table(
        index=["corp_name", "period"], columns="std_account", values="value", aggfunc="first"
    )


def _col(wide: pd.DataFrame, name: str) -> pd.Series:
    if name in wide.columns:
        return wide[name]
    return pd.Series(float("nan"), index=wide.index)


def _by_corp(check: str, bad: pd.Series, checked: pd.Series, fmt) -> list[CheckResult]:
    """bad: 위반 여부, checked: 점검 대상 여부. 둘 다 (corp_name, period) 인덱스."""
    out = []
    for corp in checked.index.get_level_values("corp_name").unique():
        n = int(checked.xs(corp, level="corp_name").sum())
        if n == 0:
            continue
        hits = bad.xs(corp, level="corp_name")
        periods = list(hits[hits].index)
        if periods:
            shown = ", ".join(fmt(corp, p) for p in periods[:4])
            more = f" 외 {len(periods) - 4}건" if len(periods) > 4 else ""
            out.append(CheckResult(check, corp, FAIL, n, shown + more))
        else:
            out.append(CheckResult(check, corp, PASS, n))
    return out


def _tolerance(base: pd.Series, relative: float) -> pd.Series:
    return (base.abs() * relative).clip(lower=1.0)


def check_identity(
    wide: pd.DataFrame, total: str, parts: tuple[str, ...], check: str, relative: float = 0.001
) -> list[CheckResult]:
    """합계 = 부분의 합. 예) 자산총계 = 부채총계 + 자본총계."""
    t = _col(wide, total)
    p = [_col(wide, name) for name in parts]
    checked = t.notna()
    for series in p:
        checked &= series.notna()
    gap = t - sum(p)
    bad = checked & (gap.abs() > _tolerance(t, relative))
    return _by_corp(check, bad, checked, lambda c, per: f"{per} 차이 {gap[(c, per)]:,.0f}")


def check_three_month(quarterly: pd.DataFrame) -> list[CheckResult]:
    """손익: API 가 준 3개월치 필드 = 누적 차분으로 만든 값.

    두 값은 서로 다른 필드에서 나오므로, 어긋나면 누적 기준이 섞였거나
    정정 전후 보고서가 섞인 것이다.
    """
    rows = quarterly[quarterly["reported_3m"].notna() & quarterly["value"].notna()].copy()
    rows["bad"] = (rows["value"] - rows["reported_3m"]).abs() > _tolerance(
        rows["reported_3m"], 0.001
    )
    out = []
    for corp, grp in rows.groupby("corp_name"):
        bad = grp[grp["bad"]]
        if bad.empty:
            out.append(CheckResult("분기 3개월치 = 누적 차분", corp, PASS, len(grp)))
            continue
        shown = ", ".join(
            f"{r.period} {r.account_label} {r.reported_3m:,.0f} vs {r.value:,.0f}"
            for r in bad.head(3).itertuples()
        )
        out.append(CheckResult("분기 3개월치 = 누적 차분", corp, FAIL, len(grp), shown))
    return out


def check_capex_sign(wide: pd.DataFrame) -> list[CheckResult]:
    """유형자산의 취득은 모든 분기에서 유출(0 이하)이어야 한다."""
    capex = _col(wide, "유형자산의 취득")
    checked = capex.notna()
    bad = checked & (capex > 0)
    return _by_corp(
        "CAPEX 는 유출(음수)", bad, checked, lambda c, per: f"{per} {capex[(c, per)]:,.0f}"
    )


def check_borrowings(wide: pd.DataFrame, result: CleanResult) -> list[CheckResult]:
    """차입금이 0 으로 새지 않았는지. 별칭 누락이 의심되면 경고, 0 인데 부채가 많으면 실패."""
    present = [name for name in BORROWING_ACCOUNTS if name in wide.columns]
    debt = wide[present].fillna(0).sum(axis=1) if present else pd.Series(0.0, index=wide.index)
    leverage = _col(wide, "부채총계") / _col(wide, "자본총계")
    suspicious = {w.split(":")[0]: w for w in result.warnings if "차입금 별칭 누락" in w}

    out = []
    for corp in wide.index.get_level_values("corp_name").unique():
        d, lev = debt.xs(corp, level="corp_name"), leverage.xs(corp, level="corp_name")
        zero_but_levered = (d == 0) & (lev > 0.5)
        n = int(lev.notna().sum())
        if zero_but_levered.any():
            periods = ", ".join(zero_but_levered[zero_but_levered].index[:4])
            out.append(
                CheckResult("차입금 누수", corp, FAIL, n, f"차입금 0, 부채비율 50% 초과: {periods}")
            )
        elif corp in suspicious:
            out.append(CheckResult("차입금 누수", corp, WARN, n, suspicious[corp].split(": ")[1]))
        else:
            out.append(CheckResult("차입금 누수", corp, PASS, n))
    return out


def check_ratio_ranges(annual_wide: pd.DataFrame) -> list[CheckResult]:
    """연간 비율이 상식 범위 안인지. 범위 밖이면 단위나 행 매칭이 틀렸을 가능성이 높다."""
    ratios = pd.DataFrame(
        {
            "영업이익률": _col(annual_wide, "영업이익") / _col(annual_wide, "매출액"),
            "부채비율": _col(annual_wide, "부채총계") / _col(annual_wide, "자본총계"),
            "현금/자산": _col(annual_wide, "현금및현금성자산") / _col(annual_wide, "자산총계"),
        }
    )
    out = []
    for corp, grp in ratios.groupby(level="corp_name"):
        problems, n = [], 0
        for name, (low, high) in RATIO_BOUNDS.items():
            series = grp[name].dropna()
            n += len(series)
            for (_, period), value in series[(series < low) | (series > high)].items():
                problems.append(f"{period} {name} {value:.2f}")
        status = FAIL if problems else PASS
        out.append(CheckResult("비율 범위", corp, status, n, ", ".join(problems[:4])))
    return out


def check_segments(
    annual_wide: pd.DataFrame, segments: pd.DataFrame, unit: str, relative: float = 0.01
) -> list[CheckResult]:
    """IR 자료의 부문 매출 합 = 연결 매출. 출처마다 단위가 달라 환산 후 비교한다."""
    seg = segments.copy()
    seg["value"] = [convert(v, u, unit) for v, u in zip(seg["revenue"], seg["unit"], strict=True)]
    total = seg.groupby(["corp_name", "bsns_year"])["value"].sum()
    out = []
    for corp, grp in total.groupby(level="corp_name"):
        problems = []
        for (_, year), seg_sum in grp.items():
            key = (corp, str(year))
            revenue = _col(annual_wide, "매출액").get(key)
            if revenue is None or pd.isna(revenue):
                problems.append(f"{year} 연결 매출 없음")
            elif abs(seg_sum - revenue) > max(1.0, abs(revenue) * relative):
                problems.append(f"{year} 부문합 {seg_sum:,.0f} vs 연결 {revenue:,.0f}")
        status = FAIL if problems else PASS
        out.append(
            CheckResult("부문합 = 연결 매출", corp, status, len(grp), ", ".join(problems[:4]))
        )
    return out


def check_gaps(quarterly: pd.DataFrame) -> list[CheckResult]:
    """분기 보고서가 빠져 값을 못 만든 기간. 틀린 값 대신 빈 값을 냈다는 사실을 알린다."""
    out = []
    for corp, grp in quarterly.groupby("corp_name"):
        gaps = sorted(grp.loc[grp["flag"] == GAP_FLAG, "period"].unique())
        n = grp["period"].nunique()
        if gaps:
            out.append(CheckResult("분기 연속성", corp, WARN, n, "값 비움: " + ", ".join(gaps)))
        else:
            out.append(CheckResult("분기 연속성", corp, PASS, n))
    return out


def run_audit(result: CleanResult, segments: pd.DataFrame | None = None) -> list[CheckResult]:
    quarterly, annual = _wide(result.quarterly), _wide(result.annual)
    checks = [
        *check_identity(quarterly, "자산총계", ("부채총계", "자본총계"), "자산 = 부채 + 자본"),
        *check_identity(
            quarterly, "자본총계", ("지배주주지분", "비지배지분"), "자본 = 지배 + 비지배"
        ),
        *check_identity(
            quarterly,
            "당기순이익",
            ("지배주주순이익", "비지배순이익"),
            "순이익 = 지배 + 비지배",
            0.01,
        ),
        *check_three_month(result.quarterly),
        *check_capex_sign(quarterly),
        *check_borrowings(quarterly, result),
        *check_ratio_ranges(annual),
        *check_gaps(result.quarterly),
    ]
    if segments is not None and not segments.empty:
        checks += check_segments(annual, segments, result.unit)
    return checks


def summarize(checks: list[CheckResult]) -> dict[str, int]:
    return {status: sum(c.status == status for c in checks) for status in (PASS, FAIL, WARN)}
