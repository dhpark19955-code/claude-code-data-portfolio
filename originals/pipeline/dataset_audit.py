# -*- coding: utf-8 -*-
"""IB DATASET 내부정합성 감사 — 회계 항등식 · 비율 범위 · 배수 시총 역산 · 지배+비지배.

사용: python dataset_audit.py <xlsx 파일명 패턴> [패턴2 ...]
  한글 경로를 argv로 넘기면 cp949로 깨지므로, 파일명 '패턴'만 받아 Downloads 아래를 glob 한다.
  예) python dataset_audit.py DATASET_IB_260919
출력: 이상 항목만. "(이상 없으면 무출력)" 이 뜨면 통과.
"""
import sys, os, glob, openpyxl

sys.stdout.reconfigure(encoding="utf-8", errors="replace")
ROOT = os.path.join(os.path.expanduser("~"), "Downloads")


def find(pat):
    hits = [p for p in glob.glob(os.path.join(ROOT, "**", "*.xlsx"), recursive=True)
            if pat in os.path.basename(p) and not os.path.basename(p).startswith("~$")]
    return sorted(hits, key=os.path.getmtime, reverse=True)


def num(v):
    return float(v) if isinstance(v, (int, float)) and not isinstance(v, bool) else None


def label_map(ws):
    """{B열 라벨: [C열~ 값]}"""
    out = {}
    for row in ws.iter_rows(values_only=True):
        if len(row) > 1 and isinstance(row[1], str) and row[1].strip():
            out.setdefault(row[1].strip(), list(row[2:]))
    return out


def check_identity(wb, sheet):
    if sheet not in wb.sheetnames:
        return
    m = label_map(wb[sheet])
    g = lambda k: [num(x) for x in m.get(k, [])]
    a, li, cap = g("자산총계"), g("부채총계"), g("자본총계")
    ctrl, nci = g("지배주주지분"), g("비지배지분")
    ni, nic = g("당기순이익"), g("지배주주순이익")
    print(f"  [{sheet}] 항등식  (이상 없으면 무출력)")
    for i in range(min(map(len, (a, li, cap))) if a and li and cap else 0):
        if None in (a[i], li[i], cap[i]):
            continue
        d = a[i] - li[i] - cap[i]
        if abs(d) > max(1.0, abs(a[i]) * 1e-3):
            print(f"    ★ 자산≠부채+자본 [열{i}] Δ{d:,.0f}")
    for i in range(min(map(len, (cap, ctrl, nci))) if cap and ctrl and nci else 0):
        if None in (cap[i], ctrl[i], nci[i]):
            continue
        d = cap[i] - ctrl[i] - nci[i]
        if abs(d) > max(1.0, abs(cap[i]) * 1e-3):
            print(f"    ★ 자본≠지배+비지배 [열{i}] Δ{d:,.0f}")


def check_ratio(wb):
    if "15_재무비율" not in wb.sheetnames:
        return
    bounds = {"영업이익률": (-3, 1), "순이익률": (-3, 2), "ROE": (-2, 2), "ROA": (-2, 2),
              "부채비율": (0, 10), "현금/자산": (0, 1), "성장률": (-1, 6)}
    print("  [15_재무비율] 범위  (OCF/영업이익은 원래 요동치므로 제외)")
    for lab, vals in label_map(wb["15_재무비율"]).items():
        for key, (lo, hi) in bounds.items():
            if key in lab:
                for i, v in enumerate(vals):
                    n = num(v)
                    if n is not None and not (lo <= n <= hi):
                        print(f"    ? {lab[:26]} [열{i}] = {n:.3f}  (범위 {lo}~{hi} 밖)")


def check_multiples(wb):
    """20_밸류에이션: 각 행 금액×배수 = 시총 이 서로 같아야 한다 → 10배 오류 탐지."""
    if "20_밸류에이션_배수" not in wb.sheetnames:
        return
    implied = []
    for row in wb["20_밸류에이션_배수"].iter_rows(values_only=True):
        lab = row[1] if len(row) > 1 else None
        nums = [num(c) for c in row[2:] if num(c) is not None]
        if isinstance(lab, str) and len(nums) >= 2 and 0.05 < nums[-1] < 100 and nums[0]:
            implied.append((lab.strip(), nums[0] * nums[-1], nums[0], nums[-1]))
    if len(implied) < 3:
        return
    mid = sorted(x[1] for x in implied)[len(implied) // 2]
    print(f"  [20_밸류에이션_배수] 시총 역산 일관성 (중앙값 {mid:,.0f})")
    for lab, mc, base, mult in implied:
        if abs(mc - mid) / mid > 0.15:
            print(f"    ? {lab[:28]:30s} {base:,.1f}×{mult:.2f}={mc:,.0f}  (중앙값의 {mc/mid:.2f}배)")


def check_ni_split(wb):
    for sheet in ("14_투자지표", "14_투자지표_컨센서스", "14_컨센서스"):
        if sheet not in wb.sheetnames:
            continue
        m = label_map(wb[sheet])
        tot, c, n = ([num(x) for x in m.get(k, [])] for k in
                     ("당기순이익", "당기순이익(지배)", "당기순이익(비지배)"))
        if tot and c and n:
            print(f"  [{sheet}] 당기순이익 = 지배 + 비지배")
            for i in range(min(map(len, (tot, c, n)))):
                if None in (tot[i], c[i], n[i]):
                    continue
                if abs(tot[i] - c[i] - n[i]) > max(1.0, abs(tot[i]) * 0.01):
                    print(f"    ★ [열{i}] {tot[i]:,.0f} ≠ {c[i]:,.0f}+{n[i]:,.0f}")


def check_hist(wb):
    """Historical PER/PBR 밴드 통계 — 평균이 중앙값의 3배 이상이거나 최대가 중앙값의 20배 이상이면 원천(주식수·이익) 이상치.
    실제 사례: 월덱스 DART 주식수 ×1000 오기 → PER 최대 15,710배·평균 703배(중앙값 9.8배)."""
    for sheet in wb.sheetnames:
        if "Historical" not in sheet:
            continue
        print(f"  [{sheet}] 밴드 통계 이상치  (평균/중앙값 >3 또는 최대/중앙값 >20)")
        for lab, v in label_map(wb[sheet]).items():
            if not lab.startswith(("PER(x)", "PBR(x)")):
                continue
            mean, med, mx = (num(v[i]) if len(v) > i else None for i in (2, 3, 6))
            if not (mean and med and mx) or med <= 0:
                continue
            if mean / med > 3 or mx / med > 20:
                print(f"    ★ {lab:14s} 평균 {mean:,.1f} · 중앙값 {med:,.1f} · 최대 {mx:,.1f} — 주식수(shares_by_period)·TTM 이익 원천 확인")


if __name__ == "__main__":
    pats = sys.argv[1:] or ["DATASET_IB"]
    for pat in pats:
        for p in find(pat):
            print("=" * 72, f"\n{os.path.relpath(p, ROOT)}")
            wb = openpyxl.load_workbook(p, data_only=True)
            check_identity(wb, "10_IS_연간_연결")
            check_identity(wb, "12_IS_연간_별도")
            check_ratio(wb)
            check_multiples(wb)
            check_ni_split(wb)
            check_hist(wb)
