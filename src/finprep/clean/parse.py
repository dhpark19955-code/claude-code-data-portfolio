"""문자열 금액과 계정명 정리.

원본 `originals/pipeline/dart_listed_fs.py` 의 to_num, clean_name, dedupe, account_key 를
옮기고 DataFrame 단위로 쓸 수 있게 했다.
"""

from __future__ import annotations

import re

import pandas as pd

from finprep.schema import NO_STD, REPORT

_EMPTY = {"", "-", "－", "nan", "None"}
_ROMAN_PREFIX = re.compile(r"^[ⅠⅡⅢⅣⅤⅥⅦⅧⅨⅩIVXivx]+[.\s]+")
_NUMBER_PREFIX = re.compile(r"^[0-9]+[.)]\s*")
# 원본은 (주5) 만 지웠다. (주석 5, 7) 표기도 지우도록 넓혔다 (테스트로 발견).
_NOTE_REF = re.compile(r"\(주(?:석)?\s*\d+[^)]*\)")


def to_num(value: object) -> float | None:
    """'1,234', '(1,234)', '-', '' 같은 표기를 숫자로. 못 읽으면 None (0 으로 바꾸지 않는다)."""
    if value is None:
        return None
    s = str(value).replace(",", "").strip()
    if s in _EMPTY:
        return None
    negative = s.startswith("(") and s.endswith(")")
    if negative:
        s = s[1:-1]
    try:
        number = float(s)
    except ValueError:
        return None
    return -number if negative else number


def clean_name(name: object) -> str:
    """로마숫자·번호 머리표, 주석 참조, 전각 공백을 떼고 공백을 하나로 줄인다."""
    s = str(name or "").replace("\u3000", " ")
    s = _ROMAN_PREFIX.sub("", s)
    s = _NUMBER_PREFIX.sub("", s)
    s = _NOTE_REF.sub("", s)
    return re.sub(r"\s+", " ", s).strip()


def dedupe(names: list[str]) -> list[str]:
    """같은 표 안의 동명 항목에 #2, #3 을 붙인다 (원문 순서 유지)."""
    seen: dict[str, int] = {}
    out = []
    for name in names:
        seen[name] = seen.get(name, 0) + 1
        out.append(name if seen[name] == 1 else f"{name}#{seen[name]}")
    return out


def parse_amounts(raw: pd.DataFrame) -> pd.DataFrame:
    """thstrm_amount, thstrm_add_amount 를 숫자 컬럼 amount, amount_cum 으로 만든다."""
    df = raw.copy()
    df["amount"] = df["thstrm_amount"].map(to_num).astype("float64")
    if "thstrm_add_amount" in df.columns:
        df["amount_cum"] = df["thstrm_add_amount"].map(to_num).astype("float64")
    else:
        df["amount_cum"] = float("nan")
    df["ord"] = pd.to_numeric(df["ord"], errors="coerce")
    return df


def add_account_keys(df: pd.DataFrame) -> pd.DataFrame:
    """행을 기간 간에 이어 붙일 키(account_key)와 표시용 이름(account_label)을 만든다.

    키는 표준계정코드가 있으면 그것을 쓴다. 계정명은 기간마다 표기가 흔들리기 때문이다
    (영업이익 / 영업이익(손실)). 코드가 없으면 정리한 계정명을 쓴다.
    한 표 안에 같은 키나 이름이 두 번 나오면 원문 순서대로 #2, #3 을 붙인다.
    """
    out = df.sort_values([*REPORT, "sj_div", "ord"], kind="stable").reset_index(drop=True)
    label = out["account_nm"].map(clean_name)
    account_id = out["account_id"].fillna("").astype(str).str.strip()
    has_id = (account_id != "") & (account_id != NO_STD)
    key = account_id.where(has_id, "nm:" + label)

    table = [*REPORT, "sj_div"]
    for column, base in (("account_key", key), ("account_label", label)):
        nth = out.assign(_base=base).groupby([*table, "_base"], sort=False).cumcount() + 1
        out[column] = base.where(nth == 1, base + "#" + nth.astype(str))
    out["has_std_id"] = has_id
    return out
