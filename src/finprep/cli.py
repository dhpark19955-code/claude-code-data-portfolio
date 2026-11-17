"""명령줄 진입점.

    python -m finprep demo                     # 가상 원자료 생성 -> 정리 -> 감사 (키 불필요)
    python -m finprep run raw.csv -o out/      # 내 원자료 CSV 정리·감사
    python -m finprep fetch 005930 -o out/     # OpenDART 수집 후 정리·감사 (DART_API_KEY 필요)

실패한 감사가 있으면 종료 코드 1 을 낸다 (CI 에서 그대로 쓸 수 있게).
"""

from __future__ import annotations

import argparse
import sys
from pathlib import Path

import pandas as pd

from finprep import clean, run_audit, sample
from finprep.audit.checks import FAIL, summarize
from finprep.report import write_outputs
from finprep.schema import UNIT_DIV

ROOT = Path(__file__).resolve().parents[2]


def _finish(raw: pd.DataFrame, segments, out_dir: Path, unit: str, title: str) -> int:
    result = clean(raw, unit=unit)
    checks = run_audit(result, segments)
    written = write_outputs(result, checks, out_dir, title)
    counts = summarize(checks)
    print(
        f"rows={len(raw)} companies={result.quality['companies']} "
        f"checks={len(checks)} pass={counts['통과']} fail={counts[FAIL]} warn={counts['경고']}"
    )
    for warning in result.warnings:
        print(f"  확인 필요: {warning}")
    for path in written:
        print(f"  -> {path}")
    return 1 if counts[FAIL] else 0


def cmd_demo(args: argparse.Namespace) -> int:
    raw = sample.generate(args.seed)
    segments = sample.generate_segments(raw, args.seed)
    data_dir = Path(args.data_dir)
    data_dir.mkdir(parents=True, exist_ok=True)
    raw.to_csv(data_dir / "raw_fnltt.csv", index=False, encoding="utf-8-sig", lineterminator="\n")
    segments.to_csv(
        data_dir / "ir_segments.csv", index=False, encoding="utf-8-sig", lineterminator="\n"
    )
    return _finish(raw, segments, Path(args.out), args.unit, "품질 리포트 (가상 데이터 5개사)")


def cmd_run(args: argparse.Namespace) -> int:
    raw = pd.read_csv(args.raw, dtype=str, keep_default_na=False)
    segments = pd.read_csv(args.segments) if args.segments else None
    return _finish(raw, segments, Path(args.out), args.unit, f"품질 리포트 ({Path(args.raw).name})")


def cmd_fetch(args: argparse.Namespace) -> int:
    from finprep.fetch.dart import DartError, fetch_company

    try:
        raw = fetch_company(args.query, years=args.years, fs_div=args.fs_div)
    except DartError as exc:
        print(f"수집 실패: {exc}", file=sys.stderr)
        return 2
    out_dir = Path(args.out)
    out_dir.mkdir(parents=True, exist_ok=True)
    raw.to_csv(out_dir / "raw_fnltt.csv", index=False, encoding="utf-8-sig")
    name = raw["corp_name"].iloc[0]
    return _finish(raw, None, out_dir, args.unit, f"품질 리포트 ({name})")


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(prog="finprep", description=__doc__.split("\n")[0])
    sub = parser.add_subparsers(dest="command", required=True)

    def common(p: argparse.ArgumentParser, out: str) -> None:
        p.add_argument("-o", "--out", default=out, help="출력 폴더")
        p.add_argument("--unit", default="백만원", choices=list(UNIT_DIV), help="출력 단위")

    demo = sub.add_parser("demo", help="가상 원자료로 전체 흐름 실행 (키 불필요)")
    common(demo, str(ROOT / "reports" / "sample"))
    demo.add_argument("--data-dir", default=str(ROOT / "data" / "sample"), help="원자료 저장 폴더")
    demo.add_argument("--seed", type=int, default=20261117)
    demo.set_defaults(func=cmd_demo)

    run = sub.add_parser("run", help="원자료 CSV 정리·감사")
    run.add_argument("raw", help="fnlttSinglAcntAll 형식 long CSV")
    run.add_argument("--segments", help="부문 매출 CSV (corp_name, bsns_year, revenue, unit)")
    common(run, "out")
    run.set_defaults(func=cmd_run)

    fetch = sub.add_parser("fetch", help="OpenDART 수집 후 정리·감사 (DART_API_KEY 필요)")
    fetch.add_argument("query", help="종목코드 6자리 또는 회사명")
    fetch.add_argument("--years", type=int, default=3)
    fetch.add_argument("--fs-div", default="CFS", choices=["CFS", "OFS"], help="연결/별도")
    common(fetch, "out")
    fetch.set_defaults(func=cmd_fetch)
    return parser


def main(argv: list[str] | None = None) -> int:
    if hasattr(sys.stdout, "reconfigure"):
        sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    args = build_parser().parse_args(argv)
    return args.func(args)
