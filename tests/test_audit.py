"""감사가 오류를 실제로 잡는지 본다.

정상 데이터에 오류를 하나씩 심고 해당 점검이 실패하는지 확인한다.
"""

import pandas as pd
import pytest

from finprep import clean, run_audit
from finprep.audit.checks import FAIL, PASS, WARN, summarize


def _status(checks, name, corp):
    return next(c.status for c in checks if c.check == name and c.corp_name == corp)


def _set(raw: pd.DataFrame, mask, **values) -> pd.DataFrame:
    out = raw.copy()
    for column, value in values.items():
        out.loc[mask, column] = value
    return out


def test_clean_sample_has_no_failures(sample_raw, sample_segments):
    checks = run_audit(clean(sample_raw), sample_segments)
    counts = summarize(checks)
    assert counts[FAIL] == 0
    # 심어 둔 함정 2개 (보고서 누락, 별칭 없는 차입금) 는 경고로 나와야 한다
    warned = {(c.check, c.corp_name) for c in checks if c.status == WARN}
    assert warned == {("분기 연속성", "마루에너지"), ("차입금 누수", "마루에너지")}


def test_identity_break_is_caught(sample_raw):
    mask = (
        (sample_raw["corp_name"] == "가온테크")
        & (sample_raw["account_nm"] == "자산총계")
        & (sample_raw["bsns_year"] == 2024)
        & (sample_raw["reprt_code"] == "11014")
    )
    broken = _set(sample_raw, mask, thstrm_amount="1")
    checks = run_audit(clean(broken))
    assert _status(checks, "자산 = 부채 + 자본", "가온테크") == FAIL
    assert _status(checks, "자산 = 부채 + 자본", "누리소프트") == PASS


def test_mixed_cumulative_field_is_caught(sample_raw):
    # 3분기 보고서의 누적 필드가 반기 값으로 잘못 들어온 상황 (정정 전후 보고서 혼재)
    mask = (
        (sample_raw["corp_name"] == "누리소프트")
        & (sample_raw["account_nm"] == "영업수익")
        & (sample_raw["bsns_year"] == 2025)
    )
    h1 = sample_raw.loc[mask & (sample_raw["reprt_code"] == "11012"), "thstrm_add_amount"].iloc[0]
    broken = _set(sample_raw, mask & (sample_raw["reprt_code"] == "11014"), thstrm_add_amount=h1)
    checks = run_audit(clean(broken))
    assert _status(checks, "분기 3개월치 = 누적 차분", "누리소프트") == FAIL


def test_abs_on_cash_flow_would_be_caught(sample_raw):
    # 음수 관행 회사의 유출 계정 중 한 해만 양수로 섞여 들어온 경우
    mask = (
        (sample_raw["corp_name"] == "다온바이오")
        & (sample_raw["account_nm"] == "유형자산의 취득")
        & (sample_raw["bsns_year"] == 2023)
    )
    flipped = sample_raw.loc[mask, "thstrm_amount"].str.lstrip("-")
    broken = sample_raw.copy()
    broken.loc[mask, "thstrm_amount"] = flipped
    checks = run_audit(clean(broken))
    assert _status(checks, "CAPEX 는 유출(음수)", "다온바이오") == FAIL


def test_borrowing_leak_is_caught(sample_raw):
    # 차입금 계정이 전부 별칭 표에 없는 이름으로 바뀌면 총차입금이 0 이 된다
    debt = sample_raw["account_nm"].isin(["단기차입금", "장기차입금"]) & (
        sample_raw["corp_name"] == "가온테크"
    )
    broken = _set(
        sample_raw, debt, account_nm="리스 및 금융부채", account_id="-표준계정코드 미사용-"
    )
    broken = broken.drop_duplicates(
        ["corp_code", "bsns_year", "reprt_code", "sj_div", "account_nm"]
    )
    checks = run_audit(clean(broken))
    assert _status(checks, "차입금 누수", "가온테크") == FAIL


def test_segment_unit_error_is_caught(sample_raw, sample_segments):
    # IR 자료가 실제로는 백만원인데 억원으로 적힌 경우 (100배 오차)
    seg = sample_segments.copy()
    mask = seg["corp_name"] == "가온테크"
    seg.loc[mask, "revenue"] = seg.loc[mask, "revenue"] * 100
    checks = run_audit(clean(sample_raw), seg)
    assert _status(checks, "부문합 = 연결 매출", "가온테크") == FAIL
    assert _status(checks, "부문합 = 연결 매출", "누리소프트") == PASS


@pytest.mark.parametrize("unit", ["원", "천원", "백만원", "억원"])
def test_audit_result_does_not_depend_on_output_unit(sample_raw, sample_segments, unit):
    counts = summarize(run_audit(clean(sample_raw, unit=unit), sample_segments))
    assert counts[FAIL] == 0
