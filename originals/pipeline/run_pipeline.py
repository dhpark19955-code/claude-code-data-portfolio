# -*- coding: utf-8 -*-
r"""국내 상장사 IB 데이터셋 원클릭 파이프라인.

    python run_pipeline.py 매커스            # 종목명
    python run_pipeline.py 093520           # 종목코드
    python run_pipeline.py 매커스 --skip-fetch   # 수집 생략(이미 받은 원자료로 재빌드)
    python run_pipeline.py 매커스 --no-docs      # 공시 원문 XML 다운로드 생략(빠름)

단계
  0 종목 해석      종목명/코드 → 종목코드·corp_code·DART 정식명 (corpCode 마스터)
  1 원자료 수집    DART 재무제표(연결·별도) · 네이버 시세/일별/수급/재무요약 · DART 메타(주식총수·배당·최대주주)
                   · CAPEX · 5%/임원 보고 · 분기별 주식수 · (옵션) 공시 원문
  2 재무요약       fs_summary.py 4시트 + 분기합=연간 검산
  3 표준 데이터셋  build_base_dataset_ib.py (README·KPI·재무·컨센·비율·배당·주식수·밸류·FCF·순차입금·지분·5%)
  4 파생 시트      Historical PER/PBR(23) · 수급 250D(70)
  5 재계산·감사    recalc_one.py(Excel COM) · dataset_audit.py
  6 배포           Downloads\<YYMMDD>_<종목>_Data\<YYMMDD>_<종목>_Data.xlsx
  7 점검 리포트    자동화 안 되는 회사별 작업(부문매출·오버행·업종 KPI·IR) 체크리스트 + 경고

키는 env_keys 로 .env 에서 자동 로드(사용자에게 묻지 않는다). 한글 인자는 subprocess 로 넘기므로 cp949 문제 없음.
"""
import os, sys, json, time, glob, shutil, argparse, subprocess, urllib.request, urllib.parse
from datetime import datetime

sys.stdout.reconfigure(encoding="utf-8", errors="replace")
HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)
from env_keys import need  # noqa: E402
BASE = os.path.dirname(HERE)
DL = os.path.join(os.path.expanduser("~"), "Downloads")
ENV = {**os.environ, "PYTHONIOENCODING": "utf-8"}
LOG = []


def say(msg):
    print(msg, flush=True); LOG.append(msg)


def run(step, args, timeout=900, tail=4, must=True):
    t0 = time.time()
    p = subprocess.run([sys.executable] + args, cwd=HERE, env=ENV, capture_output=True, text=True,
                       encoding="utf-8", errors="replace", timeout=timeout)
    out = (p.stdout or "") + (p.stderr or "")
    lines = [l for l in out.splitlines() if l.strip()]
    ok = p.returncode == 0
    say(f"  {'✓' if ok else '✗'} {step} ({time.time()-t0:.0f}s)")
    for l in lines[-tail:]:
        say(f"      {l[:160]}")
    if not ok and must:
        raise SystemExit(f"[중단] {step} 실패 — 위 로그 확인")
    return ok, out


def dart(ep, **kw):
    u = f"https://opendart.fss.or.kr/api/{ep}.json?" + urllib.parse.urlencode({"crtfc_key": K, **kw})
    try:
        return json.loads(urllib.request.urlopen(u, timeout=60).read())
    except Exception:
        return {}


# ─────────────── 0. 종목 해석 ───────────────
ap = argparse.ArgumentParser()
ap.add_argument("query", help="종목명 또는 6자리 종목코드")
ap.add_argument("--skip-fetch", action="store_true")
ap.add_argument("--no-docs", action="store_true")
ap.add_argument("--date", default=datetime.now().strftime("%y%m%d"))
a = ap.parse_args()
YMD = a.date
K = need("DART_API_KEY")
import dart_reports as dr  # noqa: E402
import re  # noqa: E402
cm = dr.corp_map(K)                                   # {종목코드: (corp_code, 정식명)}
q = a.query.strip()
norm = lambda s: re.sub(r"[\s\-_.·(),주식회사㈜]", "", str(s)).lower()
if re.fullmatch(r"[0-9A-Z]{6}", q):
    hits = [(c, v) for c, v in cm.items() if c == q]
else:
    nq = norm(q)
    hits = [(c, v) for c, v in cm.items() if norm(v[1]) == nq] or \
           [(c, v) for c, v in cm.items() if norm(v[1]).startswith(nq)] or \
           [(c, v) for c, v in cm.items() if nq in norm(v[1])]
if not hits:
    raise SystemExit(f"'{q}' 상장사를 찾지 못했다")
if len(hits) > 1:
    say(f"  후보 {len(hits)}건: " + ", ".join(f"{v[1]}({c})" for c, v in hits[:8]) + " — 정확한 이름이나 코드로 다시 실행할 것")
    raise SystemExit(1)
CODE, (CORP, NAME) = hits[0]
say(f"■ 대상: {NAME} ({CODE}) · corp_code {CORP} · 기준일 {YMD}")
D = os.path.join(BASE, "data", f"{NAME}_{CODE}")
META = os.path.join(D, "04_DART메타")
for sub in ("01_보고서원문", "02_DART재무제표", "03_네이버증권", "04_DART메타"):
    os.makedirs(os.path.join(D, sub), exist_ok=True)

# ─────────────── 1. 원자료 수집 ───────────────
if not a.skip_fetch:
    say("■ 1. 원자료 수집")
    fsdir = os.path.join(D, "02_DART재무제표")
    run("DART 재무제표 연결", ["dart_listed_fs.py", CODE, "--years", "6", "--quarters", "10", "--output", fsdir], tail=2)
    run("DART 재무제표 별도", ["dart_listed_fs.py", CODE, "--years", "6", "--quarters", "10", "--fs_div", "OFS", "--output", fsdir], tail=2, must=False)
    run("네이버 시세·일별5년·수급250D·재무요약", ["naver_market.py", CODE, "--out", os.path.join(D, "03_네이버증권"), "--years", "5"], tail=4)
    run("DART 메타(주식총수·배당·최대주주)", ["fetch_fnguide_dart_meta.py", NAME, CODE, CORP], tail=2, must=False)
    run("CAPEX(유·무형 취득)", ["fetch_capex.py", NAME, CODE, CORP, "--years", f"{datetime.now().year-6}-{datetime.now().year-1}",
                                "--quarters", f"{str(datetime.now().year-2)[2:]}.3Q-{str(datetime.now().year)[2:]}.2Q"], tail=2, must=False)
    # 5%·임원 보고 + 분기별 주식수
    hold = {ep: dart(ep, corp_code=CORP) for ep in ("majorstock", "elestock")}
    json.dump(hold, open(os.path.join(META, "holders.json"), "w", encoding="utf-8"), ensure_ascii=False, indent=1)
    say(f"  ✓ 5%/임원 보고 majorstock {len(hold['majorstock'].get('list', []))}건 · elestock {len(hold['elestock'].get('list', []))}건")
    sh = {}
    for y in range(datetime.now().year - 6, datetime.now().year + 1):
        for qq, rp in (("1Q", "11013"), ("2Q", "11012"), ("3Q", "11014"), ("4Q", "11011")):
            d = dart("stockTotqySttus", corp_code=CORP, bsns_year=str(y), reprt_code=rp)
            for r_ in d.get("list", []):
                if "보통" in str(r_.get("se", "")):
                    n = lambda x: int(str(x).replace(",", "")) if str(x).replace(",", "").strip().lstrip("-").isdigit() else None
                    sh[f"{y}{qq}"] = {"발행": n(r_.get("istc_totqy")), "자기주식": n(r_.get("tesstk_co")), "유통": n(r_.get("distb_stock_co"))}
            time.sleep(0.12)
    json.dump(sh, open(os.path.join(META, "shares_by_period.json"), "w", encoding="utf-8"), ensure_ascii=False, indent=1)
    say(f"  ✓ 분기별 주식수 {len(sh)}개 기간" + (f" (최근 {sorted(sh)[-1]} 발행 {sh[sorted(sh)[-1]]['발행']:,})" if sh else " — 없음"))
    if not a.no_docs:
        run("공시 원문 XML", ["dart_fetch_docs.py", "--corp", CORP, "--out", os.path.join(D, "01_보고서원문"),
                             "--bgn", f"{datetime.now().year-3}0101", "--end", datetime.now().strftime("%Y%m%d"),
                             "--keywords", "사업보고서,반기보고서,분기보고서,대량보유,자기주식,전환사채,신주인수권,주식매수선택권,공정공시,증권신고서,투자설명서"],
            timeout=1800, tail=1, must=False)

# 주식수 이상치 정리(--skip-fetch로 옛 json을 쓸 때도 적용)
_shp = os.path.join(META, "shares_by_period.json")
if os.path.exists(_shp):
    from share_utils import clean_shares
    _sh, _fix = clean_shares(json.load(open(_shp, encoding="utf-8")))
    if _fix:
        for m in _fix:
            say(f"  ! DART 주식수 이상치 — {m}")
        json.dump(_sh, open(_shp, "w", encoding="utf-8"), ensure_ascii=False, indent=1)

# ─────────────── 2. 재무요약 ───────────────
say("■ 2. 재무요약 4시트")
import openpyxl  # noqa: E402
fs_c = glob.glob(os.path.join(D, "02_DART재무제표", f"*_{CODE}_DART재무제표_연결.xlsx"))
if not fs_c:
    raise SystemExit("[중단] DART 연결 재무제표 파일 없음")
prefix = os.path.basename(fs_c[0]).replace("_연결.xlsx", "_{}.xlsx")
wb = openpyxl.load_workbook(fs_c[0], read_only=True)
def hdr(sheet):
    if sheet not in wb.sheetnames:
        return []
    for row in wb[sheet].iter_rows(min_row=1, max_row=4, values_only=True):
        cols = [str(c) for c in row if isinstance(c, str) and re.match(r"^(\d{4}/\d{2}|\d{2}\.\dQ)$", str(c))]
        if cols:
            return cols
    return []
ANN, QTR = hdr("포괄손익계산서(연간)")[-6:], hdr("포괄손익계산서(분기)")[-8:]
if not os.path.exists(os.path.join(D, "02_DART재무제표", prefix.format("별도"))):
    shutil.copy(fs_c[0], os.path.join(D, "02_DART재무제표", prefix.format("별도")))   # 별도 없으면 연결로 대체(표기)
    say("  ! 별도 재무제표 없음 — 연결로 대체(12·13 시트는 연결과 동일)")
if len(ANN) < 2:
    say(f"  ! 연간 재무 {len(ANN)}개 — 신규상장 등으로 사업보고서가 없다. 감사보고서 파싱(kido_fs_summary.py 패턴)으로 연간을 보강해야 한다. 여기서 중단.")
    raise SystemExit(2)
FIN = os.path.join(BASE, "분석", f"{NAME}_재무요약_5개년8분기.xlsx")
ok, out = run("fs_summary", ["fs_summary.py", "--file-pattern", os.path.join(D, "02_DART재무제표", prefix), "--name", NAME,
                              "--years", ",".join(ANN), "--quarters", ",".join(QTR), "--out", FIN], tail=0)
checks = [l.strip() for l in out.splitlines() if ("OK" in l or "차이" in l or "산출불가" in l) and ("연결" in l or "별도" in l)]
bad = [c for c in checks if "OK" not in c]
say(f"    분기합=연간 검산: OK {len(checks)-len(bad)} / 이상 {len(bad)}")
for c in bad[:6]:
    say(f"      ! {c}")

# ─────────────── 3. 표준 데이터셋 ───────────────
say("■ 3. 표준 데이터셋 빌드")
run("build_base_dataset_ib", ["build_base_dataset_ib.py", NAME, CODE, CORP, "--date", YMD], tail=1)
PAT = f"{NAME}_{CODE}_DATASET_IB_{YMD}"

# ─────────────── 4. 파생 시트 ───────────────
say("■ 4. 파생 시트")
if run("Historical PER/PBR 산출", ["hist_valuation.py", NAME, CODE], tail=0, must=False)[0]:
    run("Historical PER/PBR 주입(23)", ["add_hist_valuation_sheet.py", NAME, CODE, PAT, "23"], tail=1, must=False)
run("수급 250D 주입(70)", ["add_flow_sheet.py", NAME, CODE, PAT, "70"], tail=1, must=False)

# ─────────────── 5. 재계산·감사 ───────────────
say("■ 5. 재계산·감사")
run("Excel COM 재계산", ["recalc_one.py", PAT], timeout=600, tail=1)
ok, aud = run("dataset_audit", ["dataset_audit.py", PAT], tail=0, must=False)
flags = [l.strip() for l in aud.splitlines() if l.strip().startswith(("★", "?"))]
say(f"    감사: ★ {sum(f.startswith('★') for f in flags)}건 · ? {sum(f.startswith('?') for f in flags)}건")
for f in flags[:8]:
    say(f"      {f}")

# ─────────────── 6. 배포 ───────────────
src = os.path.join(BASE, "마스터", f"{PAT}.xlsx")
outdir = os.path.join(DL, f"{YMD}_{NAME}_Data"); os.makedirs(outdir, exist_ok=True)  # 사용자 지정(2026-10-12): 회사별 날짜폴더
dst = os.path.join(outdir, f"{YMD}_{NAME}_Data.xlsx")  # 사용자 지정 파일명(2026-10-12): 오늘자_회사명_Data
shutil.copy(src, dst)
say(f"■ 6. 배포 → {dst}")

# ─────────────── 7. 점검 리포트 ───────────────
say("■ 7. 회사별 추가 작업 체크리스트 (자동화 대상 아님 — 원문을 읽고 판단)")
warns = []
try:
    import pandas as pd
    A = pd.read_excel(FIN, sheet_name="연결_연간", index_col=0); A.index = [str(i).strip() for i in A.index]
    last = A.columns[-1]
    debt_rows = [r_ for r_ in ("단기차입금", "유동성장기차입금", "사채(유동)", "장기차입금") if r_ in A.index]
    debt = sum(float(A.loc[r_, last]) for r_ in debt_rows if pd.notna(A.loc[r_, last]))
    lev = float(A.loc["부채총계", last]) / float(A.loc["자본총계", last]) if "부채총계" in A.index else 0
    if debt == 0 and lev > 0.8:
        warns.append(f"차입금 0인데 부채비율 {lev:.0%} — fs_summary.py SPEC 차입금 별칭 누락 의심(원문 BS 라벨 확인)")
    if "매출액(영업수익)" in A.index and A.loc["매출액(영업수익)"].isna().all():
        warns.append("매출 행이 비었다 — 매출 라벨 별칭 추가 필요('수익'·'영업수익' 등)")
    blank = lambda lab: lab not in A.index or A.loc[lab].isna().all()
    if blank("지배주주순이익") and not blank("당기순이익"):
        if blank("비지배지분"):
            say("  · 비지배지분 없음 → 지배주주 행 미제공이 정상. PER=당기순이익·PBR=자본총계 기준으로 대체됨")
        else:
            warns.append("지배주주순이익 비었는데 비지배지분은 있음 — fs_summary.py 라벨 별칭 추가 필요")
except Exception as e:
    warns.append(f"점검 실패: {e}")
cap = json.load(open(os.path.join(META, "capex.json"), encoding="utf-8")) if os.path.exists(os.path.join(META, "capex.json")) else {}
vals = [v["capex"] for v in cap.get("annual", {}).values() if v]
if len(vals) >= 3 and max(vals) > 5 * (sorted(vals)[len(vals)//2] or 1):
    warns.append("CAPEX가 특정 연도에 급증 — 사옥·토지 등 비영업 투자인지 원문 유형자산 변동표 확인 후 31_FCF 주석 추가")
for w in warns:
    say(f"  ⚠ {w}")
for t in ["부문·제품·지역별 매출 — 사업보고서 'Ⅱ.사업의 내용' 표 + 연결주석 '영업부문'·'주요 고객' (30번대 시트)",
          "오버행 — 자기주식 소각계획/락업 스케줄/CB·BW·RSU (33번대 시트, 신규상장은 투자설명서·증권발행실적보고서)",
          "업종 KPI — 이 회사를 설명하는 핵심 지표(관광객·설계사·고객집중·가동률 등, 60번대)",
          "IR 자료 — 홈페이지·DART 기업설명회·공정공시 확인, 없으면 IR급 공시(주주환원계획·배당정책·RSU 조건)로 대체",
          "README에 추가 시트·빈티지·미수집 항목 기재 → recalc_one.py → dataset_audit.py → 배포 재복사"]:
    say(f"  □ {t}")
json.dump(LOG, open(os.path.join(D, f"pipeline_log_{YMD}.json"), "w", encoding="utf-8"), ensure_ascii=False, indent=1)
say(f"■ 완료 — 로그 {os.path.relpath(os.path.join(D, f'pipeline_log_{YMD}.json'), BASE)}")
