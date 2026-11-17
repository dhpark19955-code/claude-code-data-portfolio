from finprep.clean.aliases import (
    build_lookup,
    normalize_label,
    required_coverage,
    standardize_accounts,
    unmapped_borrowings,
)
from finprep.clean.parse import add_account_keys, parse_amounts

from .conftest import frame, row


def _std(*rows):
    return standardize_accounts(add_account_keys(parse_amounts(frame(*rows))))


def test_spacing_is_ignored():
    assert normalize_label("자산 총계") == normalize_label("자산총계")
    df = _std(row("BS", "자산 총계", 100))
    assert df["std_account"].iloc[0] == "자산총계"


def test_variants_map_to_standard_name():
    df = _std(
        row("BS", "유동 차입금(사채 포함)", 10, order=1),
        row("BS", "비유동 차입금(사채 포함)", 20, order=2),
        row("CIS", "영업수익", 30, order=1),
        row("CIS", "영업이익(손실)", 5, order=2),
    )
    assert list(df["std_account"]) == ["단기차입금", "장기차입금", "매출액", "영업이익"]


def test_same_label_twice_maps_to_short_then_long_debt():
    df = _std(row("BS", "차입금", 10, order=1), row("BS", "차입금", 20, order=2))
    assert list(df["std_account"]) == ["단기차입금", "장기차입금"]


def test_income_statement_preferred_over_comprehensive():
    df = _std(
        row("IS", "당기순이익", 70, order=5),
        row("CIS", "당기순이익", 70, order=1),
    )
    chosen = df[df["std_account"] == "당기순이익"]
    assert list(chosen["sj_div"]) == ["IS"]
    assert df[df["sj_div"] == "CIS"]["flag"].iloc[0].startswith("표준계정 중복")


def test_lookup_keys_are_grouped_by_statement():
    lookup = build_lookup()
    # '비지배지분' 은 재무상태표에서는 자본 항목, 손익에서는 비지배 순이익이다
    assert lookup[("BS", "비지배지분")] == "비지배지분"
    assert lookup[("IS", "비지배지분")] == "비지배순이익"


def test_unknown_borrowing_label_is_reported():
    df = _std(
        row("BS", "단기차입금", 10, order=1),
        row("BS", "유동성 차입금 및 사채", 5, order=2),
    )
    found = unmapped_borrowings(df)
    assert list(found["account_label"]) == ["유동성 차입금 및 사채"]


def test_required_coverage_counts_per_report():
    df = _std(row("BS", "자산총계", 100), row("BS", "부채총계", 40, order=2))
    cov = required_coverage(df)
    assert cov["found"].iloc[0] == 2
    assert cov["required"].iloc[0] == 7
