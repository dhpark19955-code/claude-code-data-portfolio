"""작은 손 데이터 원자료를 만드는 도우미. 실제 API 응답처럼 금액은 문자열로 둔다."""

from __future__ import annotations

import pandas as pd
import pytest

from finprep import sample

NO_STD = "-표준계정코드 미사용-"


def row(
    sj: str,
    name: str,
    amount,
    *,
    year: int = 2025,
    reprt: str = "11011",
    cum=None,
    account_id: str = NO_STD,
    corp: str = "A",
    order: int = 1,
) -> dict:
    return {
        "corp_code": corp,
        "corp_name": f"회사{corp}",
        "bsns_year": year,
        "reprt_code": reprt,
        "fs_div": "CFS",
        "sj_div": sj,
        "account_id": account_id,
        "account_nm": name,
        "thstrm_amount": "" if amount is None else str(amount),
        "thstrm_add_amount": "" if cum is None else str(cum),
        "ord": order,
    }


def frame(*rows: dict) -> pd.DataFrame:
    return pd.DataFrame(list(rows))


QUARTERS = ("11013", "11012", "11014", "11011")


def flow_year(sj: str, name: str, quarters: list[float], *, ytd_thstrm: bool, **kw) -> list[dict]:
    """한 계정의 한 해치 보고서 4건.

    ytd_thstrm=False: 손익처럼 thstrm 은 3개월치, add 는 누적, 사업보고서는 연간만.
    ytd_thstrm=True : 현금흐름처럼 thstrm 자체가 누적.
    """
    out, cum = [], 0.0
    for i, (reprt, q) in enumerate(zip(QUARTERS, quarters, strict=True)):
        cum += q
        if ytd_thstrm or reprt == "11011":
            out.append(row(sj, name, cum, reprt=reprt, order=i + 1, **kw))
        else:
            out.append(row(sj, name, q, reprt=reprt, cum=cum, order=i + 1, **kw))
    return out


@pytest.fixture(scope="session")
def sample_raw() -> pd.DataFrame:
    return sample.generate()


@pytest.fixture(scope="session")
def sample_segments(sample_raw) -> pd.DataFrame:
    return sample.generate_segments(sample_raw)
