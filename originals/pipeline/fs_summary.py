# -*- coding: utf-8 -*-
"""종목 범용 5개년 + 8분기 재무요약 (연결/별도) + 분기합-연간 검산"""
import sys, math, argparse
from pathlib import Path
import pandas as pd
sys.stdout.reconfigure(encoding="utf-8")
pd.set_option("display.width", 320)

SPEC = [
    ("매출액(영업수익)",  "IS", ["매출액", "영업수익", "수익(매출액)", "수익", "매출"]),
    ("매출원가(영업비용)", "IS", ["매출원가", "영업비용"]),
    ("매출총이익",       "IS", ["매출총이익"]),
    ("판매비와관리비",   "IS", ["판매비와관리비"]),
    ("영업이익",         "IS", ["영업이익(손실)", "영업이익"]),
    ("지분법손익",       "IS", ["지분법이익(손실)", "지분법손익", "관계기업투자손익", "지분법이익",
                                "지분법투자당기손익지분"]),  # 한국알콜 2025~ (공백 무시 매칭)
    ("금융수익",         "IS", ["금융수익"]),
    ("금융비용",         "IS", ["금융비용", "금융원가"]),
    ("기타수익",         "IS", ["기타수익", "기타영업외수익", "기타이익"]),
    ("기타비용",         "IS", ["기타비용", "기타영업외비용", "기타손실"]),
    ("세전이익",         "IS", ["법인세비용차감전이익", "법인세비용차감전순이익", "법인세비용차감전순이익(손실)", "법인세차감전순이익"]),
    ("법인세비용",       "IS", ["법인세비용", "법인세비용(수익)", "계속영업법인세비용"]),
    ("당기순이익",       "IS", ["당기순이익(손실)", "당기순이익", "반기순이익(손실)", "반기순이익",
                                "분기순이익(손실)", "분기순이익"]),
    ("  지배주주순이익", "IS", ["지배기업소유지분", "지배기업의 소유주", "지배기업소유주", "지배기업 소유주지분",
                                "지배기업의 소유주지분",
                                "지배기업의 소유주에게 귀속되는 당기순이익(손실)", "지배기업의소유주지분",
                                "지배기업의 소유주에게 귀속되는 당기순이익",
                                "지배기업의 소유주에 귀속될 당기순이익", "지배기업의 소유주에게 귀속될 당기순이익",
                                "지배기업의 소유주에 귀속되는 당기순이익", "지배기업의 소유주에 귀속될 당기순이익(손실)",
                                "지배지분순이익", "지배기업 소유주 지분", "지배주주 당기순이익", "지배주주당기순이익"]),
    ("총포괄손익",       "IS", ["총포괄이익", "당기총포괄손익", "반기총포괄이익(손실)", "총포괄손익"]),
    ("EPS(원)",          "IS", ["기본주당이익", "기본주당보통주순이익", "기본주당보통주순이익(손실)",
                              "기본주당이익(손실)", "보통주 기본주당손익", "기본주당순이익",
                              # 녹십자웰빙 2022~2024: 기본·희석이 같아 한 줄로 공시 (2026-11-02)
                              "기본및희석주당순이익", "기본및희석주당이익", "기본및희석주당이익(손실)"]),
    ("__1__", None, None),
    ("자산총계",         "BS", ["자산총계"]),
    ("  현금및현금성자산", "BS", ["현금및현금성자산"]),
    ("  단기금융상품",   "BS", ["단기금융상품"]),
    ("  기타유동금융자산", "BS", ["기타유동금융자산", "기타 유동 금융자산"]),
    ("  당기손익-공정가치금융자산", "BS", ["당기손익-공정가치측정금융자산", "당기손익-공정가치 측정 금융자산"]),
    ("  기타포괄손익-공정가치금융자산", "BS", ["기타포괄손익-공정가치 측정 비유동금융자산",
                                            "기타포괄손익-공정가치 측정 금융자산"]),
    ("  장기금융상품",     "BS", ["장기금융상품"]),
    ("  매출채권및기타채권", "BS", ["매출채권 및 기타채권", "매출채권및기타채권", "유동 매출채권"]),
    ("  재고자산",       "BS", ["재고자산"]),
    ("  관계기업투자",   "BS", ["관계기업 및 공동기업투자", "관계기업투자", "종속기업,관계기업 및 공동기업투자",
                              "종속ㆍ관계기업투자주식", "지분법적용 투자자산", "관계기업주식",
                              "관계기업 및 공동기업 투자"]),
    ("  기타비유동금융자산", "BS", ["기타비유동금융자산", "기타 비유동 금융자산"]),
    ("  투자부동산",     "BS", ["투자부동산"]),
    ("  유형자산",       "BS", ["유형자산"]),
    ("부채총계",         "BS", ["부채총계", "총부채"]),
    # 회사마다 차입금 라벨이 다르다(매커스: '유동 차입금(사채 포함)' / '비유동차입금(사채 포함)의 유동성 대체 부분' / '장기차입금(사채 포함), 총액')
    ("  단기차입금",     "BS", ["단기차입금", "유동 차입금(사채 포함)", "단기차입금및사채", "유동차입금", "단기차입부채", "차입금", "단기차입금(사채 제외)"]),  # 덴티움: 유동 '차입금' / 비유동 '차입금#2'
    ("  유동성장기차입금", "BS", ["유동성장기차입금", "유동성장기부채", "비유동차입금(사채 포함)의 유동성 대체 부분", "유동성장기차입부채", "유동성장기차입금(사채 제외)"]),
    ("  사채(유동)",     "BS", ["사채", "유동성사채"]),
    ("  장기차입금",     "BS", ["장기차입금", "장기차입금(사채 포함), 총액", "비유동 차입금(사채 포함)", "장기차입부채", "비유동차입금", "차입금#2"]),
    ("자본총계",         "BS", ["자본총계"]),
    ("  지배주주지분",   "BS", ["지배기업주주지분", "지배기업의 소유주에게 귀속되는 지분",
                              "지배기업의 소유주에게 귀속되는 자본",
                              "지배기업 소유주지분", "지배기업의 소유주지분", "지배기업소유주지분",
                              "지배기업소유지분"]),  # 코미코 BS (2026-10-01)
    ("  비지배지분",     "BS", ["비지배지분"]),
    ("__2__", None, None),
    ("영업활동현금흐름", "CF", ["영업활동 현금흐름", "영업활동현금흐름", "영업활동으로 인한 현금흐름",
                         "영업활동순현금흐름"]),  # 코리안리(보험) XBRL (2026-10-12)
    ("투자활동현금흐름", "CF", ["투자활동 현금흐름", "투자활동현금흐름", "투자활동으로 인한 현금흐름", "투자활동순현금흐름"]),
    ("재무활동현금흐름", "CF", ["재무활동 현금흐름", "재무활동현금흐름", "재무활동으로 인한 현금흐름", "재무활동순현금흐름"]),
]

def _norm(x):
    # '자산 총계'·'지배기업 소유주지분'처럼 공백만 다른 표기를 같은 계정으로 본다(한국알콜 2025~, 2026-10-01)
    return str(x).replace(" ", "").replace("　", "")


def pick(df, names):
    acc = df["계정과목"].map(_norm)
    for n in names:
        hit = df[acc == _norm(n)]
        if len(hit):
            return hit.iloc[0]
    return None


def pick_all(df, names):
    """이름 자체 + '#2','#3' 접미사 변형까지 모두 수집.
    같은 계정명이 손익귀속/포괄손익귀속에 중복 등장할 때 후보를 다 얻기 위함."""
    out = []
    acc = df["계정과목"].map(_norm)
    seen = set()
    for n in names:
        for suf in ["", "#2", "#3", "#4"]:
            hit = df[acc == _norm(n) + suf]
            for i in hit.index:  # IS를 손익+포괄 두 시트로 합치면 같은 라벨이 2행 — 전부 후보로
                if i not in seen:
                    seen.add(i); out.append(hit.loc[i])
    return out


def resolve_by_identity(df, cand_names, total_row, other_row, cols):
    """지배 + 비지배 = 당기순이익 항등식을 가장 잘 만족하는 후보를 고른다.
    (같은 계정명이 '당기순이익의 귀속'과 '총포괄손익의 귀속'에 중복 등장하는 문제 해결)"""
    if total_row is None:
        return None
    cands = pick_all(df, cand_names)
    if not cands:
        return None
    others = pick_all(df, other_row) if other_row else []
    best, best_err = None, None
    for c in cands:
        for o in (others or [None]):
            err = 0.0; n = 0
            for col in cols:
                try:
                    t = float(total_row[col]); v = float(c[col])
                    w = float(o[col]) if o is not None else 0.0
                except (KeyError, TypeError, ValueError):
                    continue
                # NaN이 하나라도 섞이면 그 컬럼은 건너뛴다(평균오차가 NaN이 되어 비교가 죽는 것 방지)
                if any(math.isnan(z) for z in (t, v, w)):
                    continue
                err += abs(v + w - t); n += 1
            if n and (best_err is None or err / n < best_err):
                best, best_err = c, err / n
    return best

def _read_is(path, is_sh):
    """포괄손익 시트 + (있으면) 별도 손익계산서 시트. 한국알콜은 2024년까지 IS(손익계산서)·CIS 분리 공시,
    2025년부터 단일 CIS — 포괄 시트만 읽으면 2024 이전 매출·영업이익이 통째로 빈다(2026-10-01)."""
    df = pd.read_excel(path, sheet_name=is_sh, header=1)
    alt = is_sh.replace("포괄손익계산서", "손익계산서")
    try:
        df2 = pd.read_excel(path, sheet_name=alt, header=1)
        df = pd.concat([df, df2], ignore_index=True)
    except ValueError:
        pass
    return df


def build(path, is_sh, bs_sh, cf_sh, cols):
    src = {"IS": _read_is(path, is_sh),
           "BS": pd.read_excel(path, sheet_name=bs_sh, header=1),
           "CF": pd.read_excel(path, sheet_name=cf_sh, header=1)}
    rows = {}
    for label, kind, names in SPEC:
        if kind is None:
            rows[label] = {c: None for c in cols}
            continue
        s = pick(src[kind], names)
        rows[label] = {c: (s[c] if s is not None and c in s.index else None) for c in cols}

    # 같은 계정명이 '#2','#3'으로 중복 등장할 때, 비어 있는 연도만 중복행에서 채운다
    # (안랩 연결 IS: 2025년은 '영업비용', 2024년은 '영업비용#2'에 값이 들어 있음)
    SKIP_FILL = {"  지배주주순이익", "  비지배지분", "  지배주주지분"}
    for label, kind, names in SPEC:
        if kind is None or label in SKIP_FILL:
            continue
        dups = pick_all(src[kind], names)
        if len(dups) < 2:
            continue
        for c in cols:
            if pd.isna(rows[label].get(c)):
                for d in dups:
                    if c in d.index and not pd.isna(d[c]):
                        rows[label][c] = d[c]
                        break

    # 지배주주순이익: '지배 + 비지배 = 당기순이익'을 만족하는 후보로 교정
    ni = pick(src["IS"], ["당기순이익(손실)", "당기순이익", "반기순이익(손실)", "반기순이익",
                          "분기순이익(손실)", "분기순이익"])
    fixed = resolve_by_identity(
        src["IS"],
        ["지배기업소유지분", "지배기업의 소유주", "지배기업소유주", "지배기업 소유주지분",
         "지배기업의 소유주지분", "지배지분순이익", "지배기업 소유주 지분",
         "지배주주 당기순이익", "지배주주당기순이익"],
        ni, ["비지배지분", "비지배분", "비지배지분순이익"], cols)
    if fixed is not None:
        rows["  지배주주순이익"] = {c: (fixed[c] if c in fixed.index else None) for c in cols}
    # 기간마다 라벨이 바뀌면(한국알콜 '지배기업의 소유주에게 귀속되는 당기순이익(손실)'→'지배지분순이익')
    # 한 행만으로는 일부 기간이 빈다 → 빈 열만 '지배+비지배=당기순이익'을 만족하는 다른 후보로 채운다
    if ni is not None:
        alias = dict((l, n) for l, k, n in SPEC if k)["  지배주주순이익"]
        cands = pick_all(src["IS"], alias)
        nci_c = pick_all(src["IS"], ["비지배지분", "비지배지분순이익", "비지배지분에 귀속되는 당기순이익(손실)",
                                     "비지배기업의 소유주에게 귀속되는 당기순이익(손실)"])
        for c in cols:
            if not pd.isna(rows["  지배주주순이익"].get(c)):
                continue
            try:
                t = float(rows["당기순이익"].get(c))  # ni 원행이 아닌 기간보정된 행(라벨 바뀐 기간 포함)
            except (KeyError, TypeError, ValueError):
                continue
            if math.isnan(t):
                continue
            for cd in cands:
                v = cd.get(c)
                if v is None or pd.isna(v):
                    continue
                if any((w := o.get(c)) is not None and not pd.isna(w) and abs(float(v) + float(w) - t) <= 2 for o in nci_c):
                    rows["  지배주주순이익"][c] = v; break

    # 별칭에 없는 변형 표기(예: 월덱스 '지배기업의 소유주에 귀속될 당기순이익') → 정규식 폴백
    if all(pd.isna(v) for v in rows["  지배주주순이익"].values()):
        acc = src["IS"]["계정과목"].astype(str).str.strip()
        m = src["IS"][acc.str.match(r"^지배기업.*(소유주|주주).*(순이익|순손익)") & ~acc.str.contains("포괄")]
        if len(m):
            s = m.iloc[0]
            rows["  지배주주순이익"] = {c: (s[c] if c in s.index else None) for c in cols}
            print(f"  · 지배주주순이익 정규식 폴백: '{s['계정과목']}'")

    return pd.DataFrame(rows).T[cols]

def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--file-pattern", required=True, help="'...{}.xlsx' 형태, {}에 연결/별도")
    ap.add_argument("--name", required=True)
    ap.add_argument("--years", required=True, help="콤마구분 연간 컬럼명")
    ap.add_argument("--quarters", required=True, help="콤마구분 분기 컬럼명(최근 8개)")
    ap.add_argument("--out", required=True)
    a = ap.parse_args()
    ANN = a.years.split(","); QTR = a.quarters.split(",")

    res = {}
    # 회사별 추가 별칭: data/<종목>_<코드>/04_DART메타/fs_alias.json {"  단기금융상품": ["기타금융자산"], ...}
    # (범용 SPEC에 넣으면 다른 회사에서 오분류될 수 있는 라벨만 여기에)
    _fa = Path(a.file_pattern).parent.parent / "04_DART메타" / "fs_alias.json"
    if _fa.exists():
        import json
        extra = json.load(open(_fa, encoding="utf-8"))
        for i, (l, k, n) in enumerate(SPEC):
            if k and l in extra:
                SPEC[i] = (l, k, list(n) + list(extra[l]))
        print(f"  · 회사별 별칭 적용: {_fa.name} {list(extra)}")
    for fs in ["연결", "별도"]:
        p = a.file_pattern.format(fs)
        res[(fs, "연간")] = build(p, "포괄손익계산서(연간)", "재무상태표(연간)", "현금흐름표(연간)", ANN)
        res[(fs, "분기")] = build(p, "포괄손익계산서(분기)", "재무상태표(분기)", "현금흐름표(분기)", QTR)

    # 비지배지분이 없는 연결(100% 자회사만 보유)은 DART에 귀속 행 자체가 없다(현대이지웰 2022~).
    # 비지배지분 공란/0 인 기간에 한해 지배주주순이익=당기순이익, 지배주주지분=자본총계로 채운다.
    for per in ["연간", "분기"]:
        df = res[("연결", per)]
        filled = []
        for c in df.columns:
            nci = df.at["  비지배지분", c]
            if not (pd.isna(nci) or float(nci) == 0):
                continue
            if pd.isna(df.at["  지배주주순이익", c]) and not pd.isna(df.at["당기순이익", c]):
                df.at["  지배주주순이익", c] = df.at["당기순이익", c]; filled.append(c)
            if pd.isna(df.at["  지배주주지분", c]) and not pd.isna(df.at["자본총계", c]):
                df.at["  지배주주지분", c] = df.at["자본총계", c]
                df.at["  비지배지분", c] = 0
        if filled:
            print(f"  · 연결 {per}: 비지배지분 없음 → 지배주주순이익=당기순이익 채움 {filled}")

    # 4Q 공란 보정: IS→CIS 결합공시 전환(한국알콜 2025)처럼 3Q누적과 FY가 다른 표에 있으면
    # dart_listed_fs 가 4Q=FY−3Q누적을 못 만든다 → FY − (1Q+2Q+3Q)로 채운다(손익 흐름 항목만, 2026-10-01)
    FLOW = ["매출액(영업수익)", "매출원가(영업비용)", "매출총이익", "판매비와관리비", "영업이익", "지분법손익",
            "금융수익", "금융비용", "기타수익", "기타비용", "세전이익", "법인세비용", "당기순이익", "  지배주주순이익"]
    for fs in ["연결", "별도"]:
        q, ann = res[(fs, "분기")], res[(fs, "연간")]
        done = set()
        for y in ann.columns:
            yy = str(y)[2:4]; qs = [f"{yy}.{i}Q" for i in range(1, 5)]
            if not all(c in q.columns for c in qs):
                continue
            for item in FLOW:
                if item not in q.index or not pd.isna(q.at[item, qs[3]]):
                    continue
                vals = [q.at[item, c] for c in qs[:3]] + [ann.at[item, y]]
                if any(pd.isna(v) for v in vals):
                    continue
                q.at[item, qs[3]] = float(vals[3]) - sum(float(v) for v in vals[:3]); done.add(qs[3])
        if done:
            print(f"  · {fs} 4Q=FY−(1Q+2Q+3Q) 보정: {sorted(done)}")

    for k, v in res.items():
        print(f"\n{'='*140}\n### {a.name} {k[0]} {k[1]} (백만원)\n{'='*140}")
        print(v.to_string(float_format=lambda x: f"{x:,.0f}"))

    print(f"\n[검산] 분기합 vs 연간 — 직전 회계연도")
    for fs in ["연결", "별도"]:
        q, ann = res[(fs, "분기")], res[(fs, "연간")]
        yr = ANN[-1]; qs = [c for c in QTR if c.startswith(yr[2:4])]
        if len(qs) != 4:
            print(f"  {fs}: {yr} 분기 4개가 안 잡힘({qs}) — 검산 생략"); continue
        for item in ["매출액(영업수익)", "영업이익", "당기순이익", "영업활동현금흐름"]:
            try:
                qsum = float(q.loc[item, qs].astype(float).sum()); an = float(ann.loc[item, yr])
                d = qsum - an
                print(f"  {fs} {item:<14} 분기합={qsum:>13,.0f} 연간={an:>13,.0f} "
                      f"{'OK' if abs(d) <= 2 else f'차이 {d:,.0f}'}")
            except (KeyError, TypeError, ValueError):
                print(f"  {fs} {item:<14} 산출불가")

    Path(a.out).parent.mkdir(parents=True, exist_ok=True)
    with pd.ExcelWriter(a.out, engine="openpyxl") as w:
        for (fs, per), df in res.items():
            df.to_excel(w, sheet_name=f"{fs}_{per}")
    print(f"\n-> {a.out}")

if __name__ == "__main__":
    main()
