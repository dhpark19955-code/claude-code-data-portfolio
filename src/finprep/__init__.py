"""finprep - 공시 API 재무제표 원자료 전처리와 기계 감사.

raw = finprep.sample.generate()          # 또는 finprep.fetch.dart.fetch_company(...)
result = finprep.clean(raw)              # 분기 tidy, 연간, 품질 로그
checks = finprep.run_audit(result)       # 항등식, 교차검증, 부호, 범위
"""

from finprep.audit.checks import run_audit
from finprep.clean.pipeline import CleanResult, clean

__version__ = "0.1.0"
__all__ = ["CleanResult", "clean", "run_audit", "__version__"]
