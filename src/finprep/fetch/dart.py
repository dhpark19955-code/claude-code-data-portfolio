"""OpenDART 단일회사 전체 재무제표(fnlttSinglAcntAll) 수집.

`originals/pipeline/dart_listed_fs.py` 의 수집부만 떼어 냈다. 결과는 finprep.clean 이 받는
long 형식 그대로이고, 금액은 API 가 준 문자열 그대로 둔다 (정리는 clean 단계의 일).

API 키는 환경변수 DART_API_KEY 로만 받는다. 키를 출력하거나 파일에 남기지 않는다.
필요 패키지: pip install "finprep[dart]"
"""

from __future__ import annotations

import io
import os
import re
import time
import xml.etree.ElementTree as ET
import zipfile
from datetime import datetime, timedelta
from pathlib import Path

import pandas as pd

from finprep.schema import REPRT_TO_QN

BASE = "https://opendart.fss.or.kr/api"
CACHE_DIR = Path.home() / ".cache" / "finprep"
CACHE_MAX_AGE = timedelta(days=7)
NO_DATA = "013"
REQUEST_INTERVAL = 0.4  # 초당 요청 한도 보호


class DartError(RuntimeError):
    pass


def _requests():
    try:
        import requests
    except ImportError as exc:  # pragma: no cover - 선택 의존성
        raise DartError('수집에는 requests 가 필요합니다: pip install "finprep[dart]"') from exc
    return requests


def api_key() -> str:
    key = os.environ.get("DART_API_KEY", "").strip()
    if not key:
        raise DartError(
            "환경변수 DART_API_KEY 가 없습니다. 키 발급: https://opendart.fss.or.kr "
            "(키 없이 돌려 보려면: python -m finprep demo)"
        )
    return key


def corp_codes(refresh: bool = False) -> pd.DataFrame:
    """상장사 고유번호 표. 7일간 캐시한다."""
    CACHE_DIR.mkdir(parents=True, exist_ok=True)
    cache = CACHE_DIR / "corp_codes.csv"
    if cache.exists() and not refresh:
        age = datetime.now() - datetime.fromtimestamp(cache.stat().st_mtime)
        if age < CACHE_MAX_AGE:
            return pd.read_csv(cache, dtype=str).fillna("")

    response = _requests().get(f"{BASE}/corpCode.xml", params={"crtfc_key": api_key()}, timeout=60)
    response.raise_for_status()
    xml_bytes = zipfile.ZipFile(io.BytesIO(response.content)).read("CORPCODE.xml")
    root = ET.fromstring(xml_bytes)
    rows = [
        {
            "corp_code": node.findtext("corp_code"),
            "corp_name": node.findtext("corp_name"),
            "stock_code": (node.findtext("stock_code") or "").strip(),
        }
        for node in root.findall("list")
    ]
    table = pd.DataFrame(rows)
    table = table[table["stock_code"] != ""]
    table.to_csv(cache, index=False, encoding="utf-8-sig")
    return table


def resolve(query: str, table: pd.DataFrame) -> tuple[str, str]:
    """종목코드(6자리) 또는 회사명 -> (corp_code, corp_name)."""
    q = query.strip()
    if re.fullmatch(r"[0-9A-Z]{6}", q):
        hit = table[table["stock_code"] == q]
    else:

        def norm(s: object) -> str:
            return re.sub(r"[\s\-_.·(),주식회사㈜]", "", str(s)).lower()

        names = table["corp_name"].map(norm)
        hit = table[names == norm(q)]
        if hit.empty:
            hit = table[names.str.startswith(norm(q))]
    if hit.empty:
        raise DartError(f"'{query}' 에 해당하는 상장사를 찾지 못했습니다")
    row = hit.iloc[0]
    return row["corp_code"], row["corp_name"]


def fetch_report(corp_code: str, year: int, reprt_code: str, fs_div: str) -> pd.DataFrame | None:
    response = _requests().get(
        f"{BASE}/fnlttSinglAcntAll.json",
        params={
            "crtfc_key": api_key(),
            "corp_code": corp_code,
            "bsns_year": year,
            "reprt_code": reprt_code,
            "fs_div": fs_div,
        },
        timeout=30,
    )
    response.raise_for_status()
    data = response.json()
    time.sleep(REQUEST_INTERVAL)
    status = data.get("status")
    if status == NO_DATA:
        return None
    if status != "000":
        raise DartError(f"{year} {reprt_code} {fs_div}: {status} {data.get('message')}")
    frame = pd.DataFrame(data["list"])
    frame["bsns_year"] = year
    frame["reprt_code"] = reprt_code
    frame["fs_div"] = fs_div
    return frame


def fetch_company(query: str, years: int = 3, fs_div: str = "CFS") -> pd.DataFrame:
    """최근 years 개 회계연도의 분기·사업보고서를 모은다. 없는 보고서는 건너뛴다."""
    api_key()  # 네트워크를 쓰기 전에 키부터 확인한다
    _requests()
    corp_code, corp_name = resolve(query, corp_codes())
    this_year = datetime.now().year
    frames = []
    for year in range(this_year - years, this_year + 1):
        for reprt_code in REPRT_TO_QN:
            frame = fetch_report(corp_code, year, reprt_code, fs_div)
            if frame is not None:
                frames.append(frame)
    if not frames:
        raise DartError(f"{corp_name}: 받은 보고서가 없습니다")
    raw = pd.concat(frames, ignore_index=True)
    raw["corp_code"] = corp_code
    raw["corp_name"] = corp_name
    return raw
