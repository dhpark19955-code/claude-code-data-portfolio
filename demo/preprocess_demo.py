# -*- coding: utf-8 -*-
"""전처리 규칙 데모 (가상 데이터, 표준 라이브러리만 사용).

실제 공시 API 응답에서 반복해서 만난 문제 5가지를 가상 회사 '가온테크' 데이터로 재현하고,
규칙으로 정리한 뒤 회계 항등식으로 감사한다. 키, 네트워크, 외부 패키지가 필요 없다.

  1) 분기 손익: 1~3분기는 3개월치, 4분기는 '연간 - 3개 분기 합'으로 만든다
  2) 현금흐름: 누적(YTD)으로 오므로 같은 계정 안에서만 차분한다
  3) 현금흐름 부호: 유출 계정이 양수로 온다. abs() 가 아니라 계정 성격으로 뒤집는다
  4) 단위 혼재: 억원 출처와 백만원 출처를 백만원으로 통일한다
  5) 계정명 변형: 별칭 표로 표준 계정에 모은다 (못 모으면 0 으로 새지 않게 경고)

실행:  python -I -X utf8 preprocess_demo.py
산출:  output/tidy.csv, output/audit_log.md
"""
import csv
import os

HERE = os.path.dirname(os.path.abspath(__file__))
OUT = os.path.join(HERE, "output")

# ---------------------------------------------------------------- 가상 원자료
# (출처, 재무제표, 계정명, 기간, 금액, 단위)
RAW = [
    # 손익: 1~3분기는 3개월치, FY 는 연간
    ("api", "IS", "매출액", "1Q", 1200, "백만원"),
    ("api", "IS", "매출액", "2Q", 1350, "백만원"),
    ("api", "IS", "매출액", "3Q", 1280, "백만원"),
    ("api", "IS", "매출액", "FY", 5300, "백만원"),
    ("api", "IS", "영업이익", "1Q", 150, "백만원"),
    ("api", "IS", "영업이익", "2Q", 190, "백만원"),
    ("api", "IS", "영업이익", "3Q", 160, "백만원"),
    ("api", "IS", "영업이익", "FY", 690, "백만원"),
    # 현금흐름: 누적(YTD), 유출 계정이 양수로 옴
    ("api", "CF", "영업활동현금흐름", "1Q", 180, "백만원"),
    ("api", "CF", "영업활동현금흐름", "2Q", 420, "백만원"),
    ("api", "CF", "영업활동현금흐름", "3Q", 610, "백만원"),
    ("api", "CF", "영업활동현금흐름", "FY", 860, "백만원"),
    ("api", "CF", "유형자산의 취득", "1Q", 90, "백만원"),
    ("api", "CF", "유형자산의 취득", "2Q", 210, "백만원"),
    ("api", "CF", "유형자산의 취득", "3Q", 300, "백만원"),
    ("api", "CF", "유형자산의 취득", "FY", 410, "백만원"),
    # 환입성 항목: 음수로 온다. abs() 를 쓰면 방향이 뒤집힌다
    ("api", "CF", "대손상각비", "FY", -6, "백만원"),
    # 재무상태표: 연말 잔액. 차입금 계정명이 회사마다 다르다
    ("api", "BS", "자산총계", "FY", 9800, "백만원"),
    ("api", "BS", "부채총계", "FY", 4100, "백만원"),
    ("api", "BS", "자본총계", "FY", 5700, "백만원"),
    ("api", "BS", "유동 차입금(사채 포함)", "FY", 900, "백만원"),
    ("api", "BS", "비유동차입금", "FY", 1500, "백만원"),
    ("api", "BS", "현금및현금성자산", "FY", 1100, "백만원"),
    # IR 자료: 억원 단위로 온 부문 매출 (합계가 연결 매출과 맞아야 함)
    ("ir", "SEG", "부문매출_플랫폼", "FY", 31.0, "억원"),
    ("ir", "SEG", "부문매출_솔루션", "FY", 22.0, "억원"),
]

UNIT_TO_MN = {"백만원": 1.0, "억원": 100.0, "십억원": 1000.0}

ALIAS = {
    "유동 차입금(사채 포함)": "단기차입금",
    "단기차입금": "단기차입금",
    "비유동차입금": "장기차입금",
    "장기차입금": "장기차입금",
}

# 현금흐름에서 '양수로 오지만 실제로는 유출'인 계정. 부호는 계정 성격으로 정한다.
CF_OUTFLOW = {"유형자산의 취득", "무형자산의 취득", "배당금의 지급", "이자의 지급"}


def normalize(raw):
    """단위 통일 + 계정명 표준화. 반환: {(재무제표, 계정, 기간): 백만원}"""
    data, warnings = {}, []
    for src, stmt, account, period, amount, unit in raw:
        if unit not in UNIT_TO_MN:
            warnings.append("단위 미상: %s %s" % (account, unit))
            continue
        std = ALIAS.get(account, account)
        if stmt == "BS" and "차입금" in account and account not in ALIAS:
            warnings.append("차입금 별칭 누락 의심: %s" % account)
        data[(stmt, std, period)] = amount * UNIT_TO_MN[unit]
    return data, warnings


def quarterly_is(data, account):
    """손익 분기값: 1~3분기는 그대로, 4분기 = 연간 - 3개 분기 합."""
    q = [data[("IS", account, p)] for p in ("1Q", "2Q", "3Q")]
    return q + [data[("IS", account, "FY")] - sum(q)]


def quarterly_cf(data, account):
    """현금흐름 분기값: 누적값을 같은 계정 안에서 차분하고, 유출 계정은 부호를 뒤집는다."""
    ytd = [data[("CF", account, p)] for p in ("1Q", "2Q", "3Q", "FY")]
    q = [ytd[0]] + [ytd[i] - ytd[i - 1] for i in range(1, 4)]
    sign = -1.0 if account in CF_OUTFLOW else 1.0
    return [sign * v for v in q]


def audit(data, tidy):
    """기계적 감사. 반환: [(항목, 통과 여부, 근거)]"""
    rows = []
    a, l, e = (data[("BS", k, "FY")] for k in ("자산총계", "부채총계", "자본총계"))
    rows.append(("자산 = 부채 + 자본", abs(a - (l + e)) < 0.5, "%.0f vs %.0f" % (a, l + e)))

    for acc in ("매출액", "영업이익"):
        s, fy = sum(tidy[("IS", acc)]), data[("IS", acc, "FY")]
        rows.append(("분기합 = 연간 (%s)" % acc, abs(s - fy) < 0.5, "%.0f vs %.0f" % (s, fy)))

    capex = tidy[("CF", "유형자산의 취득")]
    rows.append(("CAPEX 는 전 분기 유출(음수)", all(v <= 0 for v in capex), str([int(v) for v in capex])))

    seg = sum(v for (s, k, p), v in data.items() if s == "SEG")
    rev = data[("IS", "매출액", "FY")]
    rows.append(("부문합 = 연결 매출 (단위 통일 후)", abs(seg - rev) < 0.5, "%.0f vs %.0f" % (seg, rev)))

    debt = data.get(("BS", "단기차입금", "FY"), 0) + data.get(("BS", "장기차입금", "FY"), 0)
    rows.append(("차입금이 0 으로 새지 않음", debt > 0, "총차입금 %.0f" % debt))

    # abs() 로 처리했다면 틀렸을 항목을 일부러 대조해 보여준다
    v = data[("CF", "대손상각비", "FY")]
    rows.append(("환입성 항목 부호 보존 (abs 금지)", v < 0, "원값 %.0f, abs 처리 시 %.0f" % (v, abs(v))))
    return rows


def main():
    os.makedirs(OUT, exist_ok=True)
    data, warnings = normalize(RAW)

    tidy = {}
    for acc in ("매출액", "영업이익"):
        tidy[("IS", acc)] = quarterly_is(data, acc)
    for acc in ("영업활동현금흐름", "유형자산의 취득"):
        tidy[("CF", acc)] = quarterly_cf(data, acc)

    with open(os.path.join(OUT, "tidy.csv"), "w", newline="", encoding="utf-8-sig") as f:
        w = csv.writer(f)
        w.writerow(["재무제표", "계정", "1Q", "2Q", "3Q", "4Q", "단위"])
        for (stmt, acc), q in tidy.items():
            w.writerow([stmt, acc] + [int(v) for v in q] + ["백만원"])

    results = audit(data, tidy)
    lines = ["# 감사 로그 (가상 데이터)", "", "| 점검 | 결과 | 근거 |", "|---|---|---|"]
    for name, ok, why in results:
        lines.append("| %s | %s | %s |" % (name, "통과" if ok else "실패", why))
    lines += ["", "경고: %s" % ("; ".join(warnings) if warnings else "없음")]
    with open(os.path.join(OUT, "audit_log.md"), "w", encoding="utf-8") as f:
        f.write("\n".join(lines) + "\n")

    failed = [n for n, ok, _ in results if not ok]
    print("checks=%d failed=%d warnings=%d" % (len(results), len(failed), len(warnings)))
    return 1 if failed else 0


if __name__ == "__main__":
    raise SystemExit(main())
