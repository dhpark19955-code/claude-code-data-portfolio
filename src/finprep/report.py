"""정리 결과와 감사 결과를 파일로 쓴다."""

from __future__ import annotations

from pathlib import Path

import pandas as pd

from finprep.audit.checks import FAIL, PASS, WARN, CheckResult, summarize
from finprep.clean.pipeline import CleanResult

SUMMARY_ACCOUNTS = ["매출액", "영업이익", "당기순이익", "영업활동현금흐름", "유형자산의 취득"]


def summary_table(result: CleanResult) -> pd.DataFrame:
    """회사 x 표준 계정 x 분기 요약표 (사람이 훑어보기 좋은 wide 형식)."""
    q = result.quarterly
    q = q[q["std_account"].isin(SUMMARY_ACCOUNTS)]
    table = q.pivot_table(
        index=["corp_name", "std_account"], columns="period", values="value", aggfunc="first"
    )
    order = {name: i for i, name in enumerate(SUMMARY_ACCOUNTS)}
    return table.sort_index(key=lambda idx: idx.map(lambda v: order.get(v, v)))


def _md_table(header: list[str], rows: list[list[object]]) -> list[str]:
    lines = ["| " + " | ".join(header) + " |", "|" + "---|" * len(header)]
    lines += ["| " + " | ".join(str(c) for c in row) + " |" for row in rows]
    return lines


def quality_markdown(result: CleanResult, checks: list[CheckResult], title: str) -> str:
    qa = result.quality
    counts = summarize(checks)
    lines = [f"# {title}", ""]
    lines += [
        "## 감사 요약",
        "",
        f"통과 {counts[PASS]}, 실패 {counts[FAIL]}, 경고 {counts[WARN]} (점검 {len(checks)}건)",
        "",
    ]
    lines += _md_table(
        ["점검", "회사", "결과", "점검 수", "근거"],
        [[c.check, c.corp_name, c.status, c.n_checked, c.detail or "-"] for c in checks],
    )
    lines += ["", "## 전처리 단계별 처리량", ""]
    lines += _md_table(
        ["단계", "처리", "비고"],
        [
            ["입력", f"{qa['raw_rows']:,}행", f"{qa['companies']}개사, 보고서 {qa['reports']}건"],
            ["금액 파싱", f"빈 값 {qa['amount_missing']}행", "0 으로 채우지 않고 빈 값 유지"],
            [
                "계정 키",
                f"표준코드 없는 행 {qa['no_std_id_rows']}개",
                "정리한 계정명으로 대체 키 생성",
            ],
            [
                "표준 계정",
                f"{qa['std_mapped_rows']:,}행 매핑",
                f"필수 계정 {qa['required_found']}/{qa['required_total']}, "
                f"중복 후보 {qa['duplicate_candidates']}행은 우선순위로 정리",
            ],
            [
                "현금흐름 부호",
                f"{qa['sign_flipped_rows']}행 반전",
                ", ".join(f"{k} {v}" for k, v in qa["outflow_convention"].items()),
            ],
            [
                "누적 차분",
                f"{qa['diffed_rows']:,}행",
                f"직전 분기 누락으로 비운 행 {qa['gap_rows']}개",
            ],
            [
                "출력",
                f"분기 {qa['quarterly_rows']:,}행",
                f"값 있음 {qa['quarterly_values']:,}행, 단위 {result.unit}",
            ],
        ],
    )
    lines += ["", "## 사람이 확인할 것", ""]
    lines += [f"- {w}" for w in result.warnings] or ["- 없음"]
    return "\n".join(lines) + "\n"


def write_outputs(
    result: CleanResult, checks: list[CheckResult], out_dir: Path, title: str
) -> list[Path]:
    out_dir.mkdir(parents=True, exist_ok=True)
    paths = {
        "tidy_quarterly.csv": result.quarterly,
        "tidy_annual.csv": result.annual,
    }
    written = []
    for name, frame in paths.items():
        path = out_dir / name
        frame.to_csv(
            path, index=False, encoding="utf-8-sig", float_format="%.1f", lineterminator="\n"
        )
        written.append(path)

    path = out_dir / "summary_quarterly.csv"
    summary_table(result).to_csv(
        path, encoding="utf-8-sig", float_format="%.1f", lineterminator="\n"
    )
    written.append(path)

    path = out_dir / "quality_report.md"
    path.write_text(quality_markdown(result, checks, title), encoding="utf-8", newline="\n")
    written.append(path)
    return written
