"""계정명 별칭 표. 회사·연도마다 다른 계정명을 표준 계정으로 모은다.

별칭은 실제 공시에서 만난 표기를 `originals/pipeline/fs_summary.py` 의 SPEC 에서 가져왔다.
매칭은 공백을 무시한다 ('자산 총계' == '자산총계').
"""

from __future__ import annotations

import re
from dataclasses import dataclass

import pandas as pd

from finprep.schema import REPORT, STATEMENT_GROUP


@dataclass(frozen=True)
class StandardAccount:
    name: str
    group: str  # "IS" | "BS" | "CF"
    variants: tuple[str, ...]


def _acc(name: str, group: str, *variants: str) -> StandardAccount:
    return StandardAccount(name, group, (name, *variants))


STANDARD_ACCOUNTS: tuple[StandardAccount, ...] = (
    # 손익
    _acc("매출액", "IS", "영업수익", "수익(매출액)", "수익", "매출"),
    _acc("매출원가", "IS", "영업비용"),
    _acc("매출총이익", "IS"),
    _acc("판매비와관리비", "IS"),
    _acc("영업이익", "IS", "영업이익(손실)"),
    _acc("금융수익", "IS"),
    _acc("금융비용", "IS", "금융원가"),
    _acc(
        "세전이익",
        "IS",
        "법인세비용차감전이익",
        "법인세비용차감전순이익",
        "법인세비용차감전순이익(손실)",
        "법인세차감전순이익",
    ),
    _acc("법인세비용", "IS", "법인세비용(수익)", "계속영업법인세비용"),
    _acc(
        "당기순이익",
        "IS",
        "당기순이익(손실)",
        "반기순이익",
        "반기순이익(손실)",
        "분기순이익",
        "분기순이익(손실)",
    ),
    _acc(
        "지배주주순이익",
        "IS",
        "지배기업소유지분",
        "지배기업의 소유주",
        "지배기업소유주",
        "지배기업 소유주지분",
        "지배기업의 소유주지분",
        "지배기업의 소유주에게 귀속되는 당기순이익",
        "지배기업의 소유주에게 귀속되는 당기순이익(손실)",
        "지배지분순이익",
        "지배주주당기순이익",
    ),
    _acc(
        "비지배순이익",
        "IS",
        "비지배지분",
        "비지배지분순이익",
        "비지배지분에 귀속되는 당기순이익(손실)",
    ),
    _acc(
        "EPS",
        "IS",
        "기본주당이익",
        "기본주당이익(손실)",
        "기본주당순이익",
        "기본주당보통주순이익",
        "기본및희석주당이익",
    ),
    # 재무상태
    _acc("자산총계", "BS"),
    _acc("현금및현금성자산", "BS"),
    _acc("매출채권및기타채권", "BS", "매출채권 및 기타채권", "유동 매출채권"),
    _acc("재고자산", "BS"),
    _acc("유형자산", "BS"),
    _acc("부채총계", "BS", "총부채"),
    _acc(
        "단기차입금",
        "BS",
        "유동 차입금(사채 포함)",
        "단기차입금및사채",
        "유동차입금",
        "단기차입부채",
        "차입금",
        "단기차입금(사채 제외)",
    ),
    _acc(
        "유동성장기차입금",
        "BS",
        "유동성장기부채",
        "비유동차입금(사채 포함)의 유동성 대체 부분",
        "유동성장기차입부채",
    ),
    _acc("사채", "BS", "유동성사채"),
    _acc(
        "장기차입금",
        "BS",
        "장기차입금(사채 포함), 총액",
        "비유동 차입금(사채 포함)",
        "장기차입부채",
        "비유동차입금",
        "차입금#2",
    ),
    _acc("자본총계", "BS"),
    _acc(
        "지배주주지분",
        "BS",
        "지배기업주주지분",
        "지배기업의 소유주에게 귀속되는 지분",
        "지배기업의 소유주에게 귀속되는 자본",
        "지배기업 소유주지분",
        "지배기업소유지분",
    ),
    _acc("비지배지분", "BS"),
    # 현금흐름
    _acc(
        "영업활동현금흐름",
        "CF",
        "영업활동 현금흐름",
        "영업활동으로 인한 현금흐름",
        "영업활동순현금흐름",
    ),
    _acc(
        "투자활동현금흐름",
        "CF",
        "투자활동 현금흐름",
        "투자활동으로 인한 현금흐름",
        "투자활동순현금흐름",
    ),
    _acc(
        "재무활동현금흐름",
        "CF",
        "재무활동 현금흐름",
        "재무활동으로 인한 현금흐름",
        "재무활동순현금흐름",
    ),
    _acc("유형자산의 취득", "CF"),
    _acc("배당금의 지급", "CF", "배당금지급", "배당금의지급"),
)

BORROWING_ACCOUNTS = ("단기차입금", "유동성장기차입금", "사채", "장기차입금")
REQUIRED_ACCOUNTS = (
    "매출액",
    "영업이익",
    "당기순이익",
    "자산총계",
    "부채총계",
    "자본총계",
    "영업활동현금흐름",
)

_BORROWING_HINT = re.compile(r"차입|사채")
_SJ_PRIORITY = {"IS": 0, "CIS": 1, "BS": 0, "CF": 0}


def normalize_label(label: object) -> str:
    return str(label).replace(" ", "").replace("\u3000", "")


def build_lookup(
    accounts: tuple[StandardAccount, ...] = STANDARD_ACCOUNTS,
) -> dict[tuple[str, str], str]:
    lookup: dict[tuple[str, str], str] = {}
    for account in accounts:
        for variant in account.variants:
            lookup.setdefault((account.group, normalize_label(variant)), account.name)
    return lookup


def standardize_accounts(
    df: pd.DataFrame, accounts: tuple[StandardAccount, ...] = STANDARD_ACCOUNTS
) -> pd.DataFrame:
    """account_label 을 표준 계정(std_account)에 대응시킨다. 못 찾으면 비워 둔다.

    한 보고서 안에서 같은 표준 계정 후보가 여럿이면 손익계산서(IS)를 포괄손익(CIS)보다,
    원문 앞쪽 행을 뒤쪽보다 우선한다. 밀린 후보는 std_account 를 비우고 flag 에 남긴다.
    """
    lookup = build_lookup(accounts)
    out = df.copy()
    group = out["sj_div"].map(STATEMENT_GROUP)
    labels = out["account_label"].map(normalize_label)
    out["std_account"] = [lookup.get(pair) for pair in zip(group, labels, strict=True)]

    out["_rank"] = out["sj_div"].map(_SJ_PRIORITY).fillna(9)
    mapped = out[out["std_account"].notna()].sort_values(["_rank", "ord"], kind="stable")
    loser = mapped.duplicated([*REPORT, "std_account"], keep="first")
    losers = mapped.index[loser]
    out["flag"] = ""
    out.loc[losers, "flag"] = "표준계정 중복 후보: " + out.loc[losers, "std_account"]
    out.loc[losers, "std_account"] = None
    return out.drop(columns="_rank")


def unmapped_borrowings(df: pd.DataFrame) -> pd.DataFrame:
    """차입금·사채로 보이는데 별칭 표에 없는 재무상태표 계정.

    이 계정들은 총차입금 집계에서 빠져 순차입금이 작게 나온다. 0 으로 새기 전에 경고한다.
    """
    bs = df[(df["sj_div"] == "BS") & df["std_account"].isna()]
    hit = bs[bs["account_label"].str.contains(_BORROWING_HINT, na=False)]
    hit = hit[~hit["flag"].str.startswith("표준계정 중복")]
    return hit[["corp_code", "fs_div", "account_label"]].drop_duplicates().reset_index(drop=True)


def required_coverage(df: pd.DataFrame) -> pd.DataFrame:
    """보고서마다 필수 계정 7개 중 몇 개를 찾았는지."""
    found = (
        df[df["std_account"].isin(REQUIRED_ACCOUNTS)]
        .groupby(REPORT)["std_account"]
        .nunique()
        .rename("found")
    )
    reports = df[REPORT].drop_duplicates().set_index(REPORT)
    out = reports.join(found).fillna({"found": 0}).reset_index()
    out["required"] = len(REQUIRED_ACCOUNTS)
    return out
