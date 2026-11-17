import math

import pytest

from finprep.clean.parse import add_account_keys, clean_name, dedupe, parse_amounts, to_num

from .conftest import frame, row


@pytest.mark.parametrize(
    ("text", "expected"),
    [
        ("1,234,000", 1_234_000.0),
        ("(1,234)", -1234.0),
        ("-5,000", -5000.0),
        ("0", 0.0),
        (" 12 ", 12.0),
        (12.5, 12.5),
    ],
)
def test_to_num_reads_reported_formats(text, expected):
    assert to_num(text) == expected


@pytest.mark.parametrize("text", ["", "-", "－", None, "nan", "N/A"])
def test_to_num_keeps_blank_as_none_not_zero(text):
    # 빈 칸을 0 으로 바꾸면 '없음'과 '0' 이 구분되지 않는다
    assert to_num(text) is None


@pytest.mark.parametrize(
    ("raw", "expected"),
    [
        ("Ⅰ. 유동자산", "유동자산"),
        ("1. 매출액", "매출액"),
        ("매출채권(주석 5, 7)", "매출채권"),
        ("재고자산(주5)", "재고자산"),
        ("자산\u3000총계", "자산 총계"),
        ("  영업이익   (손실) ", "영업이익 (손실)"),
    ],
)
def test_clean_name_strips_prefixes_and_note_refs(raw, expected):
    assert clean_name(raw) == expected


def test_dedupe_numbers_repeats_in_order():
    assert dedupe(["차입금", "사채", "차입금", "차입금"]) == [
        "차입금",
        "사채",
        "차입금#2",
        "차입금#3",
    ]


def test_parse_amounts_uses_nan_for_missing_cum():
    df = parse_amounts(frame(row("BS", "자산총계", "1,000")))
    assert df.loc[0, "amount"] == 1000.0
    assert math.isnan(df.loc[0, "amount_cum"])


def test_account_key_prefers_standard_id_over_label():
    # 같은 표준코드면 연도별로 계정명이 달라도 같은 키가 된다
    df = frame(
        row("CIS", "영업이익", 10, year=2024, account_id="dart_OperatingIncomeLoss"),
        row("CIS", "영업이익(손실)", 12, year=2025, account_id="dart_OperatingIncomeLoss"),
    )
    keyed = add_account_keys(parse_amounts(df))
    assert keyed["account_key"].nunique() == 1


def test_duplicate_labels_in_one_table_get_suffix():
    df = frame(
        row("BS", "차입금", 100, order=1),
        row("BS", "차입금", 300, order=2),
    )
    keyed = add_account_keys(parse_amounts(df))
    assert list(keyed["account_label"]) == ["차입금", "차입금#2"]
    assert list(keyed["account_key"]) == ["nm:차입금", "nm:차입금#2"]
