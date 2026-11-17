# -*- coding: utf-8 -*-
r"""네이버증권 시세/투자지표/수급 수집기 — JSON API 기반 (2026-09 개편 대응 재작성)

finance.naver.com HTML(no_today·#_per·sise_day 페이징)은 2026-09 개편으로 전부 깨졌다.
아래 네 JSON 엔드포인트는 인증 없이 동작한다(2026-09-22 확인, UA·Referer 헤더만 필요):
  1) 현재가·거래일       polling.finance.naver.com/api/realtime/domestic/stock/{code}
  2) 시총·PER·PBR·EPS·BPS m.stock.naver.com/api/stock/{code}/integration   (totalInfos)
  3) 일별 OHLCV          api.stock.naver.com/chart/domestic/item/{code}/day?startDateTime=&endDateTime=
                          (폴백: fchart.stock.naver.com/siseJson.naver — 외국인소진율 포함)
  4) 투자자별 순매매       m.stock.naver.com/api/stock/{code}/trend?pageSize=&page=   (외국인·기관·개인, 일별)
  5) 재무요약(컨센 포함)   m.stock.naver.com/api/stock/{code}/finance/annual|quarter

공개 함수 시그니처는 구버전과 같다: snapshot(code) -> dict, sise_day(code, years) -> DataFrame.
추가: investor_trend(code, days), fin_summary(code, period).
CLI:  python naver_market.py 093520 --out ../data/매커스_093520/03_네이버증권 --years 3
"""
import sys, time, argparse, json, re
from pathlib import Path
from datetime import datetime, timedelta

import requests
import pandas as pd

sys.stdout.reconfigure(encoding="utf-8", errors="replace")
UA = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64)",
      "Referer": "https://m.stock.naver.com/", "Accept": "application/json"}


def num(s):
    if s is None:
        return None
    if isinstance(s, (int, float)):
        return float(s)
    s = str(s).replace(",", "").replace("%", "").replace("원", "").replace("배", "").replace("+", "").strip()
    if s in ("", "-", "N/A", "None"):
        return None
    try:
        return float(s)
    except ValueError:
        return None


def parse_market_sum(raw):
    """'30조 386억' / '3,014억' -> 억원 정수"""
    if raw is None:
        return None
    txt = str(raw).replace(",", "").replace("억", "").strip()
    m = re.match(r"(?:(\d+)\s*조)?\s*(\d+)?", txt)
    if not m:
        return None
    jo = int(m.group(1)) if m.group(1) else 0
    eok = int(m.group(2)) if m.group(2) else 0
    return (jo * 10000 + eok) or None


def _get_json(url, retries=2):
    for i in range(retries + 1):
        try:
            r = requests.get(url, headers=UA, timeout=30)
            if r.status_code == 200:
                return r.json()
        except (requests.RequestException, ValueError):
            pass
        time.sleep(0.6 * (i + 1))
    return None


def snapshot(code):
    """현재가·시총·PER·PBR·EPS·BPS·배당수익률·추정PER/EPS·52주 고저·거래일."""
    out = {"종목코드": code, "조회시각": datetime.now().strftime("%Y-%m-%d %H:%M:%S")}
    pol = _get_json(f"https://polling.finance.naver.com/api/realtime/domestic/stock/{code}")
    d0 = (pol or {}).get("datas", [{}])[0] if pol else {}
    out["종목명"] = d0.get("stockName", "")
    out["현재가"] = num(d0.get("closePrice"))
    out["전일대비"] = num(d0.get("compareToPreviousClosePrice"))
    out["거래일"] = str(d0.get("localTradedAt", ""))[:10]
    out["시장"] = (d0.get("stockExchangeType") or {}).get("name") or (d0.get("stockExchangeType") or {}).get("code")
    integ = _get_json(f"https://m.stock.naver.com/api/stock/{code}/integration") or {}
    ti = {x.get("code"): x.get("value") for x in integ.get("totalInfos", [])}
    out["전일종가"] = num(ti.get("lastClosePrice"))
    out["시가총액(원문)"] = ti.get("marketValue")
    out["시가총액(억원)"] = parse_market_sum(ti.get("marketValue"))
    out["PER"] = num(ti.get("per")); out["EPS"] = num(ti.get("eps"))
    out["PBR"] = num(ti.get("pbr")); out["BPS"] = num(ti.get("bps"))
    out["배당수익률(%)"] = num(ti.get("dividendRatio") or ti.get("dividendYieldRatio"))
    out["추정PER"] = num(ti.get("estimatePer")); out["추정EPS"] = num(ti.get("estimateEps"))
    out["52주최고"] = num(ti.get("highPriceOf52Weeks")); out["52주최저"] = num(ti.get("lowPriceOf52Weeks"))
    out["외국인소진율(%)"] = num(ti.get("foreignRate") or ti.get("foreignerHoldRatio"))
    out["거래량"] = num(ti.get("accumulatedTradingVolume"))
    # 상장주식수 역산(시총 ÷ 현재가). 정확한 수치는 DART stockTotqySttus 를 쓸 것.
    if out["시가총액(억원)"] and out["현재가"]:
        out["상장주식수(역산)"] = round(out["시가총액(억원)"] * 1e8 / out["현재가"])
    return out


def sise_day(code, years=3):
    """일별 OHLCV (+외국인소진율). 컬럼: 날짜·시가·고가·저가·종가·거래량·외국인소진율"""
    end = datetime.now(); start = end - timedelta(days=365 * years)
    url = (f"https://api.stock.naver.com/chart/domestic/item/{code}/day?"
           f"startDateTime={start:%Y%m%d}0000&endDateTime={end:%Y%m%d}2359")
    rows = _get_json(url)
    if rows:
        df = pd.DataFrame(rows).rename(columns={
            "localDate": "날짜", "openPrice": "시가", "highPrice": "고가", "lowPrice": "저가",
            "closePrice": "종가", "accumulatedTradingVolume": "거래량", "foreignRetentionRate": "외국인소진율"})
    else:  # 폴백: fchart
        r = requests.get(f"https://fchart.stock.naver.com/siseJson.naver?symbol={code}&requestType=1"
                         f"&startTime={start:%Y%m%d}&endTime={end:%Y%m%d}&timeframe=day", headers=UA, timeout=30)
        arr = json.loads(re.sub(r"\s+", "", r.text).replace("'", '"'))
        df = pd.DataFrame(arr[1:], columns=arr[0])
    if df.empty:
        return df
    df["날짜"] = pd.to_datetime(df["날짜"].astype(str), format="%Y%m%d", errors="coerce")
    df = df.dropna(subset=["날짜"]).sort_values("날짜").reset_index(drop=True)
    df["날짜"] = df["날짜"].dt.strftime("%Y-%m-%d")
    cols = [c for c in ["날짜", "시가", "고가", "저가", "종가", "거래량", "외국인소진율"] if c in df.columns]
    return df[cols]


def investor_trend(code, days=250):
    """투자자별 일별 순매매(주): 외국인·기관·개인 + 외국인보유비율·종가·거래량."""
    # page 파라미터는 무시되고 pageSize>60 은 JSON이 아닌 응답을 준다.
    # 대신 bizdate=<YYYYMMDD> 를 주면 그 날짜 이전 60영업일을 돌려주므로, 마지막 날짜-1일로 거슬러 간다.
    rows, bizdate, seen = [], None, set()
    for _ in range(40):
        url = f"https://m.stock.naver.com/api/stock/{code}/trend?pageSize=60" + (f"&bizdate={bizdate}" if bizdate else "")
        d = _get_json(url)
        if not d:
            break
        new = [x for x in d if x["bizdate"] not in seen]
        if not new:
            break
        rows += new; seen.update(x["bizdate"] for x in new)
        if len(rows) >= days or len(d) < 60:
            break
        last = datetime.strptime(min(x["bizdate"] for x in d), "%Y%m%d") - timedelta(days=1)
        bizdate = last.strftime("%Y%m%d"); time.sleep(0.3)
    if not rows:
        return pd.DataFrame()
    df = pd.DataFrame([{
        "날짜": pd.to_datetime(x["bizdate"], format="%Y%m%d").strftime("%Y-%m-%d"),
        "종가": num(x.get("closePrice")), "전일대비": num(x.get("compareToPreviousClosePrice")),
        "거래량": num(x.get("accumulatedTradingVolume")),
        "외국인순매수(주)": num(x.get("foreignerPureBuyQuant")), "기관순매수(주)": num(x.get("organPureBuyQuant")),
        "개인순매수(주)": num(x.get("individualPureBuyQuant")), "외국인보유비율(%)": num(x.get("foreignerHoldRatio")),
    } for x in rows]).drop_duplicates("날짜").sort_values("날짜").reset_index(drop=True)
    return df.tail(days).reset_index(drop=True)


def fin_summary(code, period="annual"):
    """네이버 재무요약(FnGuide 항목 16개, 컨센서스 열은 '(E)' 접미). period: annual|quarter"""
    d = _get_json(f"https://m.stock.naver.com/api/stock/{code}/finance/{period}")
    if not d:
        return pd.DataFrame()
    info = d["financeInfo"]
    cols = [(t["key"], t["title"].rstrip(".") + ("(E)" if t.get("isConsensus") == "Y" else "")) for t in info["trTitleList"]]
    out = []
    for r in info["rowList"]:
        row = {"항목": r["title"]}
        for k, title in cols:
            v = (r.get("columns") or {}).get(k)
            row[title] = v.get("value") if isinstance(v, dict) else v
        out.append(row)
    return pd.DataFrame(out)


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("code"); ap.add_argument("--out", required=True)
    ap.add_argument("--years", type=int, default=3); ap.add_argument("--flow-days", type=int, default=250)
    a = ap.parse_args()
    snap = snapshot(a.code)
    print(f"  스냅샷: {snap.get('종목명')} {snap.get('거래일')} 현재가 {snap.get('현재가'):,} 시총 {snap.get('시가총액(억원)')}억 "
          f"PER {snap.get('PER')} PBR {snap.get('PBR')} EPS {snap.get('EPS')} BPS {snap.get('BPS')}")
    day = sise_day(a.code, a.years)
    print(f"  일별시세: {len(day)}행 ({day['날짜'].iloc[0] if len(day) else '-'} ~ {day['날짜'].iloc[-1] if len(day) else '-'})")
    flow = investor_trend(a.code, a.flow_days)
    print(f"  수급: {len(flow)}행 ({flow['날짜'].iloc[0] if len(flow) else '-'} ~ {flow['날짜'].iloc[-1] if len(flow) else '-'})")
    fa, fq = fin_summary(a.code, "annual"), fin_summary(a.code, "quarter")
    print(f"  재무요약: 연간 {fa.shape} · 분기 {fq.shape}")
    outp = Path(a.out) / f"{a.code}_네이버_시세_투자지표.xlsx"
    outp.parent.mkdir(parents=True, exist_ok=True)
    with pd.ExcelWriter(outp, engine="openpyxl") as w:
        pd.DataFrame([{"항목": k, "값": v} for k, v in snap.items()]).to_excel(w, sheet_name="스냅샷", index=False)
        if len(day): day.to_excel(w, sheet_name="일별시세", index=False)
        if len(flow): flow.to_excel(w, sheet_name="투자자별수급", index=False)
        if len(fa): fa.to_excel(w, sheet_name="재무요약_연간", index=False)
        if len(fq): fq.to_excel(w, sheet_name="재무요약_분기", index=False)
    print(f"  → {outp}")


if __name__ == "__main__":
    main()
