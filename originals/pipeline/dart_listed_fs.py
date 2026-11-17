"""
OpenDART 상장사 재무제표 추출기 (XBRL 표준계정 기반)
====================================================
종목코드 또는 회사명을 넣으면 사업보고서/분기보고서의 4대 재무제표를
연도별 · 분기별 피벗 시트로 정리해 Excel 한 개로 저장합니다.

Usage:
    python dart_listed_fs.py 214180                  # 종목코드
    python dart_listed_fs.py 헥토이노베이션            # 회사명
    python dart_listed_fs.py 214180 --years 6 --quarters 10
    python dart_listed_fs.py 214180 --fs_div OFS     # 별도재무제표
    python dart_listed_fs.py 214180 --unit 원        # 표시 단위 변경

API 키는 환경변수 DART_API_KEY 로 주입합니다(코드 하드코딩 금지).
없으면 아래 ENV_CANDIDATES 의 .env 파일에서 읽습니다.

필요 패키지: pip install requests pandas openpyxl lxml
"""

import io
import os
import re
import sys
import time
import zipfile
import argparse
import xml.etree.ElementTree as ET
from datetime import datetime, timedelta
from pathlib import Path

import requests
import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter

BASE = "https://opendart.fss.or.kr/api"
CACHE_DIR = Path.home() / ".cache" / "dart_listed_fs"
CACHE_MAX_AGE = timedelta(days=7)

# 환경변수가 없을 때 훑어볼 .env 후보 (레포 내 기존 봇들이 쓰던 위치)
ENV_CANDIDATES = [
    Path(__file__).resolve().parent / ".env",
]

# 보고서 코드 → 분기 라벨
REPRT = {"11013": "1Q", "11012": "2Q", "11014": "3Q", "11011": "4Q"}
QUARTER_SEQ = ["11013", "11012", "11014", "11011"]

SJ_SHEETS = {           # sj_div → (시트 접두어, 스톡/플로우)
    "BS": ("재무상태표", "stock"),
    "IS": ("손익계산서", "flow"),
    "CIS": ("포괄손익계산서", "flow"),
    "CF": ("현금흐름표", "flow"),
}

UNIT_DIV = {"원": 1, "천원": 1_000, "백만원": 1_000_000, "억원": 100_000_000}


# ── API 키 ────────────────────────────────────────────
def load_api_key() -> str:
    key = os.environ.get("DART_API_KEY", "").strip()
    if key:
        return key
    for env in ENV_CANDIDATES:
        if not env.exists():
            continue
        for line in env.read_text(encoding="utf-8", errors="ignore").splitlines():
            if line.startswith("DART_API_KEY="):
                key = line.split("=", 1)[1].strip().strip('"').strip("'")
                if key:
                    print(f"  → API 키: {env} 에서 로드")
                    return key
    raise SystemExit(
        "DART_API_KEY 가 없습니다. 환경변수로 설정하세요:\n"
        "  PowerShell) $env:DART_API_KEY = '발급받은키'\n"
        "  키 발급: https://opendart.fss.or.kr"
    )


# ── corp_code 매핑 ────────────────────────────────────
def load_corp_codes(key: str, refresh: bool = False) -> pd.DataFrame:
    CACHE_DIR.mkdir(parents=True, exist_ok=True)
    cache = CACHE_DIR / "corp_codes.csv"
    fresh = cache.exists() and datetime.now() - datetime.fromtimestamp(cache.stat().st_mtime) < CACHE_MAX_AGE
    if fresh and not refresh:
        return pd.read_csv(cache, dtype=str).fillna("")

    print("  → corpCode.xml 갱신 중...")
    r = requests.get(f"{BASE}/corpCode.xml", params={"crtfc_key": key}, timeout=60)
    r.raise_for_status()
    xml_bytes = zipfile.ZipFile(io.BytesIO(r.content)).read("CORPCODE.xml")
    root = ET.fromstring(xml_bytes)
    rows = [
        {
            "corp_code": n.findtext("corp_code"),
            "corp_name": n.findtext("corp_name"),
            "stock_code": (n.findtext("stock_code") or "").strip(),
        }
        for n in root.findall("list")
    ]
    df = pd.DataFrame(rows)
    df = df[df["stock_code"] != ""]
    df.to_csv(cache, index=False, encoding="utf-8-sig")
    print(f"  → 상장사 {len(df)}개 캐시 완료")
    return df


def resolve_corp(query: str, corp_df: pd.DataFrame) -> tuple[str, str, str]:
    q = query.strip()
    if re.fullmatch(r"[0-9A-Z]{6}", q):   # 2026 신규 종목코드는 영문 포함(예: 0117P0)
        hit = corp_df[corp_df["stock_code"] == q]
        if hit.empty:
            raise ValueError(f"종목코드 {q} 를 DART 상장사 목록에서 찾을 수 없습니다.")
    else:
        norm = lambda s: re.sub(r"[\s\-_.·(),주식회사㈜]", "", str(s)).lower()
        nq = norm(q)
        names = corp_df["corp_name"].map(norm)
        hit = corp_df[names == nq]
        if hit.empty:
            hit = corp_df[names.str.startswith(nq)]
        if hit.empty:
            hit = corp_df[names.str.contains(re.escape(nq), na=False)]
        if hit.empty:
            raise ValueError(f"'{query}' 에 해당하는 상장사를 찾을 수 없습니다.")
        if len(hit) > 1:
            print(f"  → 후보 {len(hit)}건: " + ", ".join(hit["corp_name"].head(5)))
    row = hit.iloc[0]
    return row["corp_code"], row["corp_name"], row["stock_code"]


# ── 재무제표 조회 ─────────────────────────────────────
def fetch_fs(key: str, corp_code: str, year: int, reprt_code: str, fs_div: str) -> pd.DataFrame | None:
    r = requests.get(
        f"{BASE}/fnlttSinglAcntAll.json",
        params={
            "crtfc_key": key,
            "corp_code": corp_code,
            "bsns_year": year,
            "reprt_code": reprt_code,
            "fs_div": fs_div,
        },
        timeout=30,
    )
    r.raise_for_status()
    data = r.json()
    time.sleep(0.4)

    if data.get("status") in ("013",):     # 조회된 데이터 없음
        return None
    if data.get("status") != "000":
        print(f"     · [{year} {REPRT.get(reprt_code, reprt_code)} {fs_div}] "
              f"{data.get('status')} {data.get('message')}")
        return None
    df = pd.DataFrame(data["list"])
    df["bsns_year"] = year
    df["reprt_code"] = reprt_code
    return df


def to_num(s):
    if s is None:
        return None
    s = str(s).replace(",", "").strip()
    if s in ("", "-", "－", "－", "nan"):
        return None
    neg = s.startswith("(") and s.endswith(")")
    if neg:
        s = s[1:-1]
    try:
        v = float(s)
    except ValueError:
        return None
    return -v if neg else v


def clean_name(s: str) -> str:
    """로마숫자/번호 prefix, 주석 참조, 전각공백 제거."""
    s = str(s or "")
    s = s.replace("　", " ")
    s = re.sub(r"^[ⅠⅡⅢⅣⅤⅥⅦⅧⅨⅩIVXivx]+[.\s]+", "", s)
    s = re.sub(r"^[0-9]+[.)]\s*", "", s)
    s = re.sub(r"\(주\s*\d+[^)]*\)", "", s)
    return re.sub(r"\s+", " ", s).strip()


def dedupe(names: list[str]) -> list[str]:
    """같은 표 안의 동명 계정에 #2, #3 접미 (원문 순서 유지)."""
    seen, out = {}, []
    for n in names:
        seen[n] = seen.get(n, 0) + 1
        out.append(n if seen[n] == 1 else f"{n}#{seen[n]}")
    return out


NO_STD = "-표준계정코드 미사용-"


def account_key(row) -> str:
    """
    행 매칭 키. 표준계정코드(account_id)가 있으면 그것을 쓴다.
    기간마다 표기가 흔들리는 계정명(영업이익 ↔ 영업이익(손실))을 한 행으로 묶기 위함.
    """
    acc_id = str(row.get("account_id") or "").strip()
    if acc_id and acc_id != NO_STD:
        return acc_id
    return "nm:" + clean_name(row.get("account_nm"))


def _period_series(df: pd.DataFrame, sj_div: str, col: str) -> tuple[pd.Series, dict[str, str]]:
    """한 기간·한 재무제표의 (키→값) 시리즈와 (키→표시명) 매핑."""
    sub = df[df["sj_div"] == sj_div].copy()
    if sub.empty:
        return pd.Series(dtype=float), {}
    sub["ord"] = pd.to_numeric(sub["ord"], errors="coerce")
    sub = sub.sort_values("ord")

    keys = dedupe([account_key(r) for _, r in sub.iterrows()])
    labels = dedupe([clean_name(n) for n in sub["account_nm"]])
    vals = [to_num(v) for v in sub[col]] if col in sub.columns else [None] * len(sub)
    return pd.Series(vals, index=keys), dict(zip(keys, labels))


def cumulative_col(df: pd.DataFrame) -> pd.Series:
    """
    기초 대비 누적액. 손익(CIS/IS)은 thstrm_add_amount 가 누적,
    현금흐름(CF)은 thstrm_amount 자체가 누적으로 오는 것이 DART 관행이라
    add 가 있으면 add, 없으면 thstrm 을 누적으로 본다. 1Q·연간은 둘이 같다.
    """
    thstrm = df["thstrm_amount"] if "thstrm_amount" in df.columns else pd.Series(dtype=object)
    if "thstrm_add_amount" not in df.columns:
        return thstrm
    add = df["thstrm_add_amount"]
    return add.where(add.notna() & (add.astype(str).str.strip() != ""), thstrm)


def pivot_period(frames: dict[str, pd.DataFrame], sj_div: str, kind: str, quarterly: bool) -> pd.DataFrame:
    """
    {기간라벨: raw df} → 계정(행) × 기간(열) 피벗.

    · stock(재무상태표)  : 기말잔액 그대로
    · flow(손익·현금흐름) : 연간 시트는 당해연도 금액,
                           분기 시트는 누적액을 직전 분기와 차분해 3개월 값으로 환산
    행 순서는 가장 최근 기간의 DART 원문 ord 를 따르고, 과거에만 있던 계정은 뒤에 붙인다.
    """
    cols, labels, order = {}, {}, []
    for label, df in frames.items():
        if kind == "stock" or not quarterly:
            src = df.assign(_v=df["thstrm_amount"])
        else:
            src = df.assign(_v=cumulative_col(df))
        ser, lab = _period_series(src, sj_div, "_v")
        if ser.empty:
            continue
        cols[label] = ser
        labels.update(lab)
        for k in ser.index:
            if k not in order:
                order.append(k)

    if not cols:
        return pd.DataFrame()

    out = pd.DataFrame(cols)

    # 분기 플로우: 같은 회계연도 안에서 누적액을 차분 → 3개월 값
    if kind == "flow" and quarterly:
        diffed, prev_year, prev_col = {}, None, None
        for label in out.columns:
            year, q = label.split(".")
            cur = out[label]
            diffed[label] = cur if (q == "1Q" or year != prev_year) else cur - out[prev_col]
            prev_year, prev_col = year, label
        out = pd.DataFrame(diffed)

    latest = list(cols[list(cols)[-1]].index)
    out = out.reindex(latest + [k for k in order if k not in latest])
    # 기간별로 붙은 #n 접미는 떼고 전체 기준으로 다시 매긴다 (표기 흔들림에 따른 라벨 충돌 방지)
    out.index = dedupe([re.sub(r"#\d+$", "", labels.get(k, k)) for k in out.index])
    out.index.name = "계정과목"
    return out


def scale(df: pd.DataFrame, div: int) -> pd.DataFrame:
    """단위 환산. 주당 금액·주식수 행은 원 단위 그대로 둔다."""
    if div == 1:
        return df.round(0)
    keep = df.index.to_series().str.contains("주당|주식수|배당금", na=False)
    out = df.copy()
    out.loc[~keep.values] = out.loc[~keep.values] / div
    return out.round(0)


def check_bs(sheets: dict[str, pd.DataFrame], unit: str):
    """회계 항등식(자산총계 = 부채총계 + 자본총계) 검증 — 잘못된 행 매칭 조기 발견용."""
    for name in ("재무상태표(연간)", "재무상태표(분기)"):
        df = sheets.get(name)
        if df is None or df.empty:
            continue
        try:
            a, l, e = df.loc["자산총계"], df.loc["부채총계"], df.loc["자본총계"]
        except KeyError:
            print(f"     · [검증] {name}: 자산/부채/자본 총계 행을 찾지 못함")
            continue
        gap = (a - (l + e)).abs()
        bad = gap[(gap > a.abs() * 0.001) & a.notna()]
        if bad.empty:
            print(f"     · [검증] {name}: 전 기간 항등식 통과 ({len(a.dropna())}개 기간)")
        else:
            print(f"     · [검증] {name}: 항등식 불일치 {list(bad.index)} (단위 {unit})")


# ── Excel 저장 ────────────────────────────────────────
HDR_FILL = PatternFill("solid", fgColor="1F3864")
SUB_FILL = PatternFill("solid", fgColor="D9E1F2")


def write_excel(path: Path, meta: dict, sheets: dict[str, pd.DataFrame], raw: pd.DataFrame, unit: str):
    with pd.ExcelWriter(path, engine="openpyxl") as xw:
        toc = pd.DataFrame(
            [{"구분": k.split("(")[1].rstrip(")"), "시트": k, "행×열": f"{v.shape[0]} × {v.shape[1]}"}
             for k, v in sheets.items() if not v.empty]
        )
        toc.to_excel(xw, sheet_name="목차", index=False, startrow=5)
        for name, df in sheets.items():
            if df.empty:
                continue
            df.to_excel(xw, sheet_name=name[:31], startrow=1, startcol=1)
        raw.to_excel(xw, sheet_name="Raw(원본)", index=False)

    wb = load_workbook(path)

    ws = wb["목차"]
    ws["A1"] = f"{meta['corp_name']} ({meta['stock_code']}) 재무제표 — OpenDART"
    ws["A1"].font = Font(bold=True, size=14)
    ws["A2"] = f"corp_code {meta['corp_code']} · 기준 {meta['fs_label']} · 단위 {unit}"
    ws["A3"] = f"출처: OpenDART fnlttSinglAcntAll · 추출 {meta['stamp']}"
    ws["A4"] = "분기 손익/현금흐름은 각 보고서의 '당기 3개월' 값입니다(누적 아님)."
    for c in ws[6]:
        if c.value:
            c.font, c.fill = Font(bold=True, color="FFFFFF"), HDR_FILL
    ws.column_dimensions["A"].width = 18
    ws.column_dimensions["B"].width = 34
    ws.column_dimensions["C"].width = 14

    for name in sheets:
        if name[:31] not in wb.sheetnames:
            continue
        ws = wb[name[:31]]
        ws["A1"] = f"{meta['corp_name']} — {name} (단위: {unit})"
        ws["A1"].font = Font(bold=True, size=12)
        for c in ws[2]:
            if c.value:
                c.font, c.fill = Font(bold=True), SUB_FILL
                c.alignment = Alignment(horizontal="center")
        ws.column_dimensions["B"].width = 40
        for col in range(3, ws.max_column + 1):
            ws.column_dimensions[get_column_letter(col)].width = 16
            for row in range(3, ws.max_row + 1):
                ws.cell(row, col).number_format = "#,##0"
        ws.freeze_panes = "C3"
        ws.auto_filter.ref = f"B2:{get_column_letter(ws.max_column)}{ws.max_row}"

    wb.save(path)


# ── 메인 ──────────────────────────────────────────────
def main():
    ap = argparse.ArgumentParser(description="OpenDART 상장사 재무제표 추출기")
    ap.add_argument("company", help="종목코드(6자리) 또는 회사명")
    ap.add_argument("--years", type=int, default=5, help="연간 최근 N개 회계연도 (기본 5)")
    ap.add_argument("--quarters", type=int, default=8, help="분기 최근 N개 분기 (기본 8)")
    ap.add_argument("--fs_div", default="CFS", choices=["CFS", "OFS"], help="CFS=연결(기본), OFS=별도")
    ap.add_argument("--unit", default="백만원", choices=list(UNIT_DIV), help="표시 단위 (기본 백만원)")
    ap.add_argument("--output", default="./output", help="저장 경로")
    ap.add_argument("--refresh", action="store_true", help="corp_code 캐시 강제 갱신")
    args = ap.parse_args()

    print("=" * 60)
    print("OpenDART 상장사 재무제표 추출기")
    print("=" * 60)

    key = load_api_key()
    corp_df = load_corp_codes(key, args.refresh)
    corp_code, corp_name, stock_code = resolve_corp(args.company, corp_df)
    print(f"[1/4] 대상: {corp_name} ({stock_code}) · corp_code {corp_code}")

    this_year = datetime.now().year
    fs_used = set()
    ofs_fallback = []  # 연결 요청했으나 별도로 대체된 기간

    # ── 연간 ──
    print(f"[2/4] 연간 재무제표 {args.years}개년 조회...")
    annual = {}
    for year in range(this_year, this_year - args.years - 2, -1):
        if len(annual) >= args.years:
            break
        df = fetch_fs(key, corp_code, year, "11011", args.fs_div)
        if df is None and args.fs_div == "CFS":
            df = fetch_fs(key, corp_code, year, "11011", "OFS")
            if df is not None:
                fs_used.add("OFS")
                ofs_fallback.append(f"{year}/12")
        elif df is not None:
            fs_used.add(args.fs_div)
        if df is None:
            continue
        annual[f"{year}/12"] = df
        print(f"     · {year} 사업보고서 {len(df)}행")
    annual = dict(reversed(list(annual.items())))

    # ── 분기 ──
    print(f"[3/4] 분기 재무제표 {args.quarters}개 분기 조회...")
    quarterly = {}
    for year in range(this_year, this_year - (args.quarters // 4) - 3, -1):
        for rc in reversed(QUARTER_SEQ):
            if len(quarterly) >= args.quarters:
                break
            df = fetch_fs(key, corp_code, year, rc, args.fs_div)
            if df is None and args.fs_div == "CFS":
                df = fetch_fs(key, corp_code, year, rc, "OFS")
                if df is not None:
                    ofs_fallback.append(f"{str(year)[2:]}.{REPRT[rc]}")
            if df is None:
                continue
            quarterly[f"{str(year)[2:]}.{REPRT[rc]}"] = df
            print(f"     · {year} {REPRT[rc]} {len(df)}행")
        if len(quarterly) >= args.quarters:
            break
    quarterly = dict(reversed(list(quarterly.items())))

    if not annual and not quarterly:
        raise SystemExit("조회된 재무제표가 없습니다. 종목코드/기간을 확인하세요.")

    # ── 피벗 ──
    print("[4/4] 피벗 및 Excel 저장...")
    div = UNIT_DIV[args.unit]
    sheets: dict[str, pd.DataFrame] = {}
    for sj, (label, kind) in SJ_SHEETS.items():
        for period, frames in (("연간", annual), ("분기", quarterly)):
            if not frames:
                continue
            piv = pivot_period(frames, sj, kind, quarterly=(period == "분기"))
            if piv.empty:
                continue
            sheets[f"{label}({period})"] = scale(piv, div)

    check_bs(sheets, args.unit)

    raw = pd.concat(
        [df.assign(기간=lbl) for lbl, df in {**annual, **quarterly}.items()],
        ignore_index=True,
    )

    outdir = Path(args.output)
    outdir.mkdir(parents=True, exist_ok=True)
    # 파일명은 요청 기준(연결/별도)으로 고정 — 하위 파이프라인이 *_연결.xlsx 를 찾는다.
    # 연결이 없어 별도로 대체된 기간은 경고로 남긴다('/' 가 든 라벨은 경로로 해석돼 저장 실패).
    fs_label = "별도" if args.fs_div == "OFS" else "연결"
    if ofs_fallback:
        print(f"  ⚠ 연결 미제출로 별도 대체된 기간: {', '.join(ofs_fallback)}")
    path = outdir / f"{corp_name}_{stock_code}_DART재무제표_{fs_label}.xlsx"
    write_excel(
        path,
        {
            "corp_name": corp_name,
            "stock_code": stock_code,
            "corp_code": corp_code,
            "fs_label": fs_label,
            "stamp": datetime.now().strftime("%Y-%m-%d %H:%M"),
        },
        sheets,
        raw,
        args.unit,
    )
    print(f"  → 저장 완료: {path}")
    print(f"  → 시트: {', '.join(sheets)}")


if __name__ == "__main__":
    main()
