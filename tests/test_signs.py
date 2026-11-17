from finprep.clean.parse import add_account_keys, parse_amounts
from finprep.clean.signs import apply_cf_signs, detect_outflow_convention, is_outflow_account

from .conftest import frame, row


def _signed(*rows):
    return apply_cf_signs(add_account_keys(parse_amounts(frame(*rows))))


def test_outflow_account_detection():
    assert is_outflow_account("유형자산의 취득")
    assert is_outflow_account("배당금의 지급#2")
    assert is_outflow_account("장기차입금의 상환")
    assert not is_outflow_account("단기차입금의 차입")
    assert not is_outflow_account("유형자산의 처분")


def test_positive_convention_is_flipped():
    df = _signed(
        row("CF", "유형자산의 취득", 300, order=1),
        row("CF", "배당금의 지급", 50, order=2),
    )
    assert list(df["amount"]) == [-300, -50]
    assert df["sign_flipped"].all()


def test_negative_convention_is_left_alone():
    df = _signed(
        row("CF", "유형자산의 취득", -300, order=1),
        row("CF", "배당금의 지급", -50, order=2),
    )
    assert list(df["amount"]) == [-300, -50]
    assert not df["sign_flipped"].any()


def test_reversal_keeps_direction_unlike_abs():
    # 양수 관행 회사에서 '취득' 계정에 환급성 음수가 섞여 있으면,
    # -abs() 는 방향을 지우지만 -1 곱셈은 방향을 보존한다.
    df = _signed(
        row("CF", "유형자산의 취득", 300, order=1),
        row("CF", "무형자산의 취득", 80, order=2),
        row("CF", "이자의 지급", -5, order=3),  # 이자 환급으로 순유입
    )
    interest = df[df["account_nm"] == "이자의 지급"]["amount"].iloc[0]
    assert interest == 5
    assert -abs(-5) == -5  # abs 로 처리했다면 유출로 둔갑


def test_convention_is_detected_per_company():
    df = parse_amounts(
        frame(
            row("CF", "유형자산의 취득", 300, corp="A"),
            row("CF", "유형자산의 취득", -300, corp="B"),
        )
    )
    conv = detect_outflow_convention(add_account_keys(df))
    assert conv[("A", "CFS")] == "positive"
    assert conv[("B", "CFS")] == "negative"


def test_non_cf_rows_untouched():
    df = _signed(row("BS", "배당금의 지급", 50), row("CF", "유형자산의 취득", 10, order=2))
    assert df[df["sj_div"] == "BS"]["amount"].iloc[0] == 50
