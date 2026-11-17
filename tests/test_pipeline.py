import numpy as np
import pandas as pd
import pytest

from finprep import clean, sample
from finprep.clean.pipeline import TIDY_COLUMNS


def test_sample_is_deterministic():
    pd.testing.assert_frame_equal(sample.generate(7), sample.generate(7))
    assert not sample.generate(7).equals(sample.generate(8))


def test_tidy_has_fixed_schema(sample_raw):
    result = clean(sample_raw)
    assert list(result.quarterly.columns) == TIDY_COLUMNS
    assert list(result.annual.columns) == TIDY_COLUMNS


def test_quarters_sum_to_annual(sample_raw):
    result = clean(sample_raw)
    keys = ["corp_name", "std_account", "bsns_year"]
    flows = ["매출액", "영업이익", "영업활동현금흐름", "유형자산의 취득"]
    q = result.quarterly[result.quarterly["std_account"].isin(flows)]
    full = q.groupby(keys)["value"].agg(["sum", "count"])
    full = full[full["count"] == 4]["sum"]  # 네 분기가 다 있는 해만
    a = result.annual[result.annual["std_account"].isin(flows)].set_index(keys)["value"]
    joined = pd.concat([full.rename("q"), a.rename("a")], axis=1, join="inner")
    assert len(joined) > 40
    np.testing.assert_allclose(joined["q"], joined["a"], rtol=1e-9)


def test_label_change_does_not_split_series(sample_raw):
    # 누리소프트는 2025년부터 '영업이익(손실)' 로 표기한다. 한 계열로 이어져야 한다.
    q = clean(sample_raw).quarterly
    op = q[(q["corp_name"] == "누리소프트") & (q["std_account"] == "영업이익")]
    assert op["period"].nunique() == 12
    assert op["value"].notna().all()


def test_quality_counts_are_consistent(sample_raw):
    qa = clean(sample_raw).quality
    assert qa["raw_rows"] == len(sample_raw)
    assert qa["required_found"] == qa["required_total"]
    assert qa["outflow_convention"]["다온바이오"] == "negative"


def test_missing_columns_are_reported():
    with pytest.raises(ValueError, match="컬럼"):
        clean(pd.DataFrame({"corp_code": ["1"]}))
