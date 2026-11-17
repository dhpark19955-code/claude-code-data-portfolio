# 실제 사용 중인 파이프라인 코드 (원본)

데이터셋 파이프라인의 핵심 모듈 10개입니다. 공개 공시 API와 공개 시세만 사용합니다.

| 파일 | 역할 |
|---|---|
| `pipeline/run_pipeline.py` | 종목명 하나로 수집부터 배포까지 8단계 실행 |
| `pipeline/dart_listed_fs.py` | 공시 API에서 연결·별도, 연간·분기 재무제표 수집 |
| `pipeline/fs_summary.py` | 재무 요약 4시트 생성, 분기합과 연간 대조, 계정명 별칭 |
| `pipeline/naver_market.py` | 시세, 일별 주가, 투자자별 수급, 재무 요약 |
| `pipeline/share_utils.py` | 주식수 이상치 정리 |
| `pipeline/hist_valuation.py` | 일별 Historical PER/PBR |
| `pipeline/ib_style.py` | IB 서식 상수와 셀 쓰기 함수 |
| `pipeline/ib_dataset.py` | 시트 선언만으로 데이터셋을 만드는 엔진 |
| `pipeline/dataset_audit.py` | 항등식, 비율 범위, 배수 역산 감사 |
| `pipeline/recalc_one.py` | Excel 재계산과 오류 셀 수 확인 |

## 실행에 필요한 것

- 공시 API 키는 환경변수 `DART_API_KEY` 로 넣습니다. 키 로더(`env_keys.py`)와 회사별 빌드 스크립트는 이 저장소에 없어서 `run_pipeline.py` 는 그대로는 끝까지 돌지 않습니다.
- 서식 엔진과 감사 도구는 단독으로 쓸 수 있습니다. 같은 양식의 실행 가능한 예시는 `scripts/build_ib_sample.py` 입니다.

## 원본과 다른 점

`dart_listed_fs.py` 에서 키 파일 탐색 후보 경로 3줄(개인 PC의 폴더 이름)을 지웠습니다. 그 밖에는 원본 그대로입니다.

---

# 실제 사용 중인 Claude Code 설정 (원본)

| 파일 | 내용 |
|---|---|
| `claude/CLAUDE.md` | 전역 지침 8개 절. 작업 유형별 라우팅 표, 일하는 방식, 보안, 환경 함정 |
| `claude/skills/company-dataset-ib/SKILL.md` | 데이터셋 수집·정리·검증·갱신·배포 절차 |
| `claude/skills/company-dataset-ib/references/pitfalls.md` | Excel 파일 손상, 공시·통계·시세 수집 함정 (증상, 원인, 해법) |
| `claude/skills/fin-data-fetch/SKILL.md` | 국내·해외 원자료 수집 진입 파일과 반드시 지킬 규칙 |
| `claude/skills/model-builder/SKILL.md` | 모델 제작 표준 흐름, 3관문, 하지 않는 것 |
| `claude/agents/worker.md` `helper.md` `checker.md` | 서브에이전트 3종 정의 |

## 원본과 다른 점

- 사용자 폴더 경로, 드라이브 폴더 이름, 키 파일 위치를 `<패키지>`, `<프로젝트>`, `<키 파일>` 같은 자리표시자로 바꿨습니다.
- 스킬 이름 앞의 접두어와 특정 회사·프로젝트를 가리키는 표현을 지웠습니다.
- 스킬이 가리키는 정본 폴더(엔진, 린터, 품질 기준 문서)는 이 저장소에 없습니다. 구조 설명은 `docs/05_setup_detail.md` 에 있습니다.
- 대시보드 작업기록 스킬은 싣지 않았습니다.
