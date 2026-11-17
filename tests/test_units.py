import pandas as pd
import pytest

from finprep.clean.units import convert, scale_values


def test_scale_keeps_per_share_rows_in_won():
    df = pd.DataFrame(
        {
            "account_label": ["매출액", "기본주당이익", "발행주식수"],
            "std_account": ["매출액", "EPS", None],
            "value": [5_000_000_000.0, 1234.0, 20_000_000.0],
        }
    )
    out = scale_values(df, "백만원")
    assert list(out["value"]) == [5000.0, 1234.0, 20_000_000.0]
    assert list(out["unit"]) == ["백만원", "원", "원"]


def test_unknown_output_unit_raises():
    df = pd.DataFrame({"account_label": ["매출액"], "value": [1.0]})
    with pytest.raises(ValueError, match="단위"):
        scale_values(df, "만원")


@pytest.mark.parametrize(
    ("amount", "src", "dst", "expected"),
    [
        (31.0, "억원", "백만원", 3100.0),
        (2.2, "십억원", "백만원", 2200.0),
        (5300.0, "백만원", "억원", 53.0),
    ],
)
def test_convert_between_source_units(amount, src, dst, expected):
    assert convert(amount, src, dst) == pytest.approx(expected)


def test_convert_refuses_to_guess_unit():
    with pytest.raises(ValueError):
        convert(1.0, "만원", "백만원")
