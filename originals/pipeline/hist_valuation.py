# -*- coding: utf-8 -*-
r"""Historical PER / PBR 산정 — 일별 주가 × 트레일링 EPS·BPS (종목 범용).

입력
  data\<종목>_<코드>\03_네이버증권\<코드>_네이버_시세_투자지표.xlsx  (일별시세: naver_market.py)
  분석\<종목>_재무요약_5개년8분기.xlsx                             (연결_연간·연결_분기: DART, 백만원)
  data\<종목>_<코드>\04_DART메타\shares_by_period.json              (분기별 발행·자기주식, stockTotqySttus)
방법
  TTM EPS(q) = 최근 4개 분기 지배주주순이익 합 ÷ 유통주식수(발행−자기주식, 그 시점)   ← 8분기 창 안
  연간만 있는 구간은 FY 지배주주순이익 ÷ 유통주식수.  BPS(q) = 지배주주지분 ÷ 유통주식수.
  각 수치는 공시 시점부터 적용(분기 +45일, 결산 +80일) → 계단함수로 일별 주가에 대응.
  PER = 종가 ÷ TTM EPS (EPS≤0 이면 공란=적자), PBR = 종가 ÷ BPS.
출력  분석\<종목>_Historical_PERPBR.xlsx  (일별 · 월말 · 통계 · 펀더멘털포인트)
사용  python hist_valuation.py <종목명> <코드> [--shares-override 20260821:5860707]
"""
import os, sys, json, argparse, glob
from datetime import datetime, timedelta
import numpy as np, pandas as pd, openpyxl

sys.stdout.reconfigure(encoding="utf-8", errors="replace")
BASE = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))


def q_end(y, q):
    return datetime(y, [3, 6, 9, 12][q - 1], [31, 30, 30, 31][q - 1])


def parse_col(c):
    """'2024/12' -> (FY, 2024-12-31, is_annual) / '24.3Q' -> (Q, 2024-09-30) / '26.1H(누적)' -> None(제외)"""
    s = str(c).strip()
    if "/" in s and s[:4].isdigit():
        return ("A", datetime(int(s[:4]), 12, 31))
    if "Q" in s and "." in s and "누적" not in s and "H" not in s:
        yy, q = s.split("."); return ("Q", q_end(2000 + int(yy), int(q[0])))
    return None


def load_fs(name):
    p = os.path.join(BASE, "분석", f"{name}_재무요약_5개년8분기.xlsx")
    out = {}
    for sh in ("연결_연간", "연결_분기"):
        df = pd.read_excel(p, sheet_name=sh, index_col=0)
        df.index = [str(i).strip() for i in df.index]
        out[sh] = df
    return out


def shares_lookup(name, code, override):
    p = os.path.join(BASE, "data", f"{name}_{code}", "04_DART메타", "shares_by_period.json")
    pts = []
    if os.path.exists(p):
        from share_utils import clean_shares
        raw, fixlog = clean_shares(json.load(open(p, encoding="utf-8")))
        for m in fixlog:
            print(f"  ! 주식수 이상치 보정: {m}")
        for k, v in raw.items():
            y, q = int(k[:4]), int(k[4])
            if v.get("발행"):
                pts.append((q_end(y, q), v["발행"] - (v.get("자기주식") or 0)))
    for k, v in (override or {}).items():
        pts.append((datetime.strptime(k, "%Y%m%d"), int(v)))
    pts.sort()
    # 액면분할·병합·무상증자: 네이버 일별시세는 수정주가라 과거 주식수도 같은 배수로 맞춰야 한다.
    # data\<종목>_DART메타\splits.json = {"YYYYMMDD(효력일)": 배수, ...} (분할 1:5 → 5, 병합 10:1 → 0.1)
    sp = os.path.join(BASE, "data", f"{name}_{code}", "04_DART메타", "splits.json")
    if os.path.exists(sp):
        for k, fac in json.load(open(sp, encoding="utf-8")).items():
            d0 = datetime.strptime(k, "%Y%m%d")
            pts = [(t, round(s * float(fac)) if t < d0 else s) for t, s in pts]
            print(f"  · 주식수 분할보정: {k} 이전 ×{fac}")
    for (t0, s0), (t1, s1) in zip(pts, pts[1:]):
        if s0 and s1 and not (0.6 < s1 / s0 < 1.6):
            print(f"  ⚠ 유통주식수 급변 {t0:%Y-%m-%d} {s0:,} → {t1:%Y-%m-%d} {s1:,} (×{s1/s0:.2f}) — 액면분할·병합·무상증자면 splits.json 필요")
    def f(d):
        cand = [s for t, s in pts if t <= d]
        return cand[-1] if cand else (pts[0][1] if pts else None)
    return f, pts


def fundamentals(fs, shares_at):
    """(effective_date, ttm_eps, bps, label) 포인트 리스트."""
    A, Q = fs["연결_연간"], fs["연결_분기"]
    def g(df, row, col):
        for r in (row, "  " + row):
            if r in df.index:
                v = df.loc[r, col]; return None if pd.isna(v) else float(v)
        return None
    def eq_parent(df, col):
        # 지배주주지분 행이 없으면 자본총계 − 비지배지분 (라벨 별칭 누락 대비; 비지배도 없으면 자본총계)
        tot, nci = g(df, "자본총계", col), g(df, "비지배지분", col)
        return None if tot is None else tot - (nci or 0)
    pts = []
    qcols = [(c, parse_col(c)) for c in Q.columns if parse_col(c) and parse_col(c)[0] == "Q"]
    qcols.sort(key=lambda x: x[1][1])
    ni_q = {}
    for c, (_, d) in qcols:
        ni_q[d] = g(Q, "지배주주순이익", c) if g(Q, "지배주주순이익", c) is not None else g(Q, "당기순이익", c)
    qd = sorted(ni_q)
    for i, d in enumerate(qd):
        if i >= 3 and all(ni_q[qd[j]] is not None for j in range(i - 3, i + 1)):
            ttm = sum(ni_q[qd[j]] for j in range(i - 3, i + 1)) * 1e6
            eq = g(Q, "지배주주지분", qcols[i][0]) or eq_parent(Q, qcols[i][0])
            sh = shares_at(d)
            if sh:
                lag = 80 if d.month == 12 else 45
                pts.append((d + timedelta(days=lag), ttm / sh, (eq * 1e6 / sh) if eq else None, f"{d:%y}.{(d.month-1)//3+1}Q TTM"))
    for c in A.columns:
        pc = parse_col(c)
        if not pc: continue
        d = pc[1]
        if any(abs((p_d - timedelta(days=80) - d).days) < 5 for p_d, *_ in pts):
            continue  # 분기 TTM 으로 이미 커버
        ni = g(A, "지배주주순이익", c) if g(A, "지배주주순이익", c) is not None else g(A, "당기순이익", c)
        eq = g(A, "지배주주지분", c) or eq_parent(A, c)
        sh = shares_at(d)
        if ni is not None and sh:
            pts.append((d + timedelta(days=80), ni * 1e6 / sh, (eq * 1e6 / sh) if eq else None, f"FY{d.year}"))
    pts.sort(key=lambda x: x[0])
    return pts


def build(name, code, override=None):
    px_path = glob.glob(os.path.join(BASE, "data", f"{name}_{code}", "03_네이버증권", f"{code}_네이버_시세_투자지표.xlsx"))[0]
    px = pd.read_excel(px_path, sheet_name="일별시세"); px["날짜"] = pd.to_datetime(px["날짜"])
    shares_at, share_pts = shares_lookup(name, code, override)
    pts = fundamentals(load_fs(name), shares_at)
    if not pts:
        raise SystemExit("펀더멘털 포인트 없음")
    eff = [p[0] for p in pts]
    def at(d):
        i = np.searchsorted(eff, d, side="right") - 1
        return pts[i] if i >= 0 else None
    rows = []
    for _, r in px.iterrows():
        p = at(r["날짜"])
        if not p: continue
        eps, bps = p[1], p[2]
        rows.append({"날짜": r["날짜"], "종가": r["종가"], "TTM EPS(원)": round(eps, 1), "BPS(원)": round(bps, 1) if bps else None,
                     "PER(x)": round(r["종가"] / eps, 2) if eps and eps > 0 else None,
                     "PBR(x)": round(r["종가"] / bps, 3) if bps and bps > 0 else None, "적용 실적": p[3]})
    daily = pd.DataFrame(rows)
    monthly = daily.groupby(daily["날짜"].dt.to_period("M")).tail(1).reset_index(drop=True)

    def stats(s, label):
        s = s.dropna()
        if s.empty: return {"지표": label}
        cur = s.iloc[-1]
        return {"지표": label, "관측수": int(s.size), "현재": round(cur, 2), "평균": round(s.mean(), 2), "중앙값": round(s.median(), 2),
                "표준편차": round(s.std(), 2), "최소": round(s.min(), 2), "최대": round(s.max(), 2),
                "25%": round(s.quantile(.25), 2), "75%": round(s.quantile(.75), 2),
                "현재 백분위(%)": round((s < cur).mean() * 100, 1), "Z": round((cur - s.mean()) / s.std(), 2) if s.std() else None}
    st = []
    for col in ("PER(x)", "PBR(x)"):
        st.append(stats(daily[col], f"{col} 전기간"))
        st.append(stats(daily[daily["날짜"] >= daily["날짜"].max() - pd.Timedelta(days=365 * 3)][col], f"{col} 최근3년"))
        st.append(stats(daily[daily["날짜"] >= daily["날짜"].max() - pd.Timedelta(days=365)][col], f"{col} 최근1년"))
    stats_df = pd.DataFrame(st)
    fund_df = pd.DataFrame([{"적용시작일": d.strftime("%Y-%m-%d"), "실적": lab, "TTM EPS(원)": round(e, 1), "BPS(원)": round(b, 1) if b else None,
                             "유통주식수(발행−자기주식)": shares_at(d)} for d, e, b, lab in pts])
    sh_df = pd.DataFrame([{"기준일": d.strftime("%Y-%m-%d"), "유통주식수": s} for d, s in share_pts])
    out = os.path.join(BASE, "분석", f"{name}_Historical_PERPBR.xlsx")
    with pd.ExcelWriter(out, engine="openpyxl") as w:
        d2 = daily.copy(); d2["날짜"] = d2["날짜"].dt.strftime("%Y-%m-%d"); d2.to_excel(w, sheet_name="일별", index=False)
        m2 = monthly.copy(); m2["날짜"] = m2["날짜"].dt.strftime("%Y-%m-%d"); m2.to_excel(w, sheet_name="월말", index=False)
        stats_df.to_excel(w, sheet_name="통계", index=False); fund_df.to_excel(w, sheet_name="펀더멘털포인트", index=False); sh_df.to_excel(w, sheet_name="주식수", index=False)
    print(f"[{name}] 일별 {len(daily)}행 ({daily['날짜'].min():%Y-%m-%d}~{daily['날짜'].max():%Y-%m-%d}) · 펀더멘털 포인트 {len(pts)}개")
    print(stats_df.to_string(index=False)); print("->", os.path.relpath(out, BASE))
    return daily, monthly, stats_df, fund_df


if __name__ == "__main__":
    ap = argparse.ArgumentParser(); ap.add_argument("name"); ap.add_argument("code")
    ap.add_argument("--shares-override", default="", help="YYYYMMDD:주식수,YYYYMMDD:주식수 (신규상장 등)")
    a = ap.parse_args()
    ov = dict(x.split(":") for x in a.shares_override.split(",") if x)
    build(a.name, a.code, ov)
