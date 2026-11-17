# -*- coding: utf-8 -*-
r"""DART 주식수(stockTotqySttus) 이상치 정리.

실제 사례(2026-09-23 월덱스): 2025년 1Q·반기 보고서의 발행주식총수가 16,510,993,000주로
정상치(16,510,993주)의 1000배로 기재돼 있었다 → 그 기간 EPS·BPS가 1/1000, Historical PER이 12,000배로 튀었다.
원천이 틀리므로 수집 단계에서 걸러야 한다.

규칙: 전체 기간 발행주식수 중앙값 대비 50배 이상 차이 나는 값은 이상치.
      ÷1000(또는 ×1000) 했을 때 중앙값 ±20% 안이면 '단위 오기'로 보정, 아니면 제외.
"""
import statistics


def clean_shares(sh: dict):
    """{기간: {'발행','자기주식','유통'}} → (정리된 dict, 조치 로그 list)."""
    vals = [v["발행"] for v in sh.values() if v.get("발행")]
    if len(vals) < 3:
        return sh, []
    med = statistics.median(vals)
    out, log = {}, []
    for k, v in sh.items():
        n = v.get("발행")
        if not n:
            out[k] = v; continue
        ratio = n / med
        if 1 / 50 < ratio < 50:
            out[k] = v; continue
        fixed = None
        for f in (1000, 1 / 1000, 100, 1 / 100):
            if abs(n / f - med) / med < 0.2:
                fixed = f; break
        if fixed:
            nv = {kk: (int(round(x / fixed)) if isinstance(x, (int, float)) and x else x) for kk, x in v.items()}
            out[k] = nv
            log.append(f"{k}: 발행 {n:,} → {nv['발행']:,} (÷{fixed:g} 단위 오기 보정, 중앙값 {med:,.0f})")
        else:
            log.append(f"{k}: 발행 {n:,} 제외 (중앙값 {med:,.0f} 대비 {ratio:.1f}배, 보정 불가)")
    return out, log
