"""원자료 스키마와 상수. 컬럼명은 OpenDART fnlttSinglAcntAll 응답을 따른다."""

from __future__ import annotations

# 보고서 코드 -> 분기 번호 (1분기, 반기, 3분기, 사업보고서)
REPRT_TO_QN: dict[str, int] = {"11013": 1, "11012": 2, "11014": 3, "11011": 4}
ANNUAL_REPRT = "11011"

# 재무제표 구분 -> 잔액(stock)인지 기간값(flow)인지
STATEMENT_KIND: dict[str, str] = {"BS": "stock", "IS": "flow", "CIS": "flow", "CF": "flow"}

# 표준 계정을 찾을 때 묶는 그룹. 손익은 IS 와 CIS 어느 쪽에든 올 수 있다.
STATEMENT_GROUP: dict[str, str] = {"BS": "BS", "IS": "IS", "CIS": "IS", "CF": "CF"}

# 원 단위 금액을 나누는 값
UNIT_DIV: dict[str, int] = {"원": 1, "천원": 1_000, "백만원": 1_000_000, "억원": 100_000_000}

# 외부 출처(IR 자료 등) 단위 -> 원
UNIT_TO_WON: dict[str, int] = {**UNIT_DIV, "십억원": 1_000_000_000, "조원": 1_000_000_000_000}

# 표준계정코드가 없을 때 API 가 주는 값
NO_STD = "-표준계정코드 미사용-"

RAW_REQUIRED = [
    "corp_code",
    "bsns_year",
    "reprt_code",
    "fs_div",
    "sj_div",
    "account_id",
    "account_nm",
    "thstrm_amount",
    "ord",
]

ENTITY = ["corp_code", "fs_div"]
REPORT = [*ENTITY, "bsns_year", "reprt_code"]
