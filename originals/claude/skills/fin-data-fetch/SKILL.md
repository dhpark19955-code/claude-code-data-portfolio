---
name: fin-data-fetch
description: 국내(DART·네이버·KOSIS·ECOS·관세청)·해외(SEC EDGAR companyfacts·XBRL 인스턴스·6-K·Form 4·13F·FINRA·Yahoo·FMP·HKEX) 상장사 재무·시세·지분·컨센서스 원자료를 어디서 어떻게 받는지 정한 수집 표준과 단독 실행형 수집 템플릿. 사용자가 "재무데이터 받아줘", "재무제표 뽑아줘/긁어줘", "DART/SEC/EDGAR에서", "해외 종목 재무", "미국 주식 10-K/10-Q 숫자", "주가·시총·PER 받아", "컨센서스 수집", "13F/공매도/내부자 매매", "홍콩 공시", "이 데이터 어디서 받아?", "API 막혔어" 등 숫자 원자료 수집·소스 선택을 말하면 명시적으로 부르지 않아도 이 스킬을 따른다. 받은 숫자로 IB 데이터셋 xlsx 를 만드는 단계는 company-dataset-ib 스킬이 맡는다(수집 단계에선 이 스킬의 라우팅·함정을 먼저 본다).
---

# 재무데이터 수집 표준 (국내·해외) — 진입 파일

**정본은 Google Drive(H:) 의 이식 패키지에 있다** (2026-10-19 정본 이전, 같은 날 사용자가 Downloads → Google Drive 로 옮김 — 옮긴 자리가 정본):

```
<패키지>   = <DATASET_지침_패키지>\
<정본폴더> = <패키지>\skills\fin-data-fetch\
  references\재무데이터_수집로직.md   ← 먼저 읽는다 (공통 원칙·국내/해외 라우팅 표·DART/SEC 규칙·막힌 경로·체크리스트)
  templates\kr_fetch_template.py · us_fetch_template.py · inputs\*.json · data_README_template.md · env.example
  scripts\preflight.py
<패키지>\01_데이터셋_작업지침_통합.md   ← 사용자 지정 규칙 전부(저장 위치·파일명·IB 서식·감사)
```

⚠ **유실분**: `eu_fetch_template.py`(ESEF) · `kr_unlisted_audit_template.py`(비상장 외감) · `검증샘플_261014\` — 구 정본 `<개인폴더>\<수집로직>\` 과 함께 PC에서 사라짐(10-19 확인). 요청이 오면 없다고 알리고 `<정본폴더>\SKILL.md` 의 대체 방법(ESEF는 filings.xbrl.org·LEI, 비상장은 DART 감사보고서 원문)으로 다시 만든다.
패키지 폴더가 없으면 사용자가 옮긴 것이다(10-19 Downloads → Google Drive `<Drive>\` 로 이동됨) → `<Drive>\`·`Downloads` 아래에서 `*DATASET_지침_이식패키지*` 폴더를 찾아 쓰고, 찾으면 이 파일의 경로를 갱신한다.
이 스킬 폴더(`~\.claude\skills\fin-data-fetch\`)의 `references\`·`templates\`·`scripts\` 는 옛 사본이라 **읽지 않는다**.

## 0. 시작 (매 세션)

1. 인터프리터: `python` 이 WindowsApps 스텁이면 exit 49 로 조용히 실패 → `<실제 python 경로>` (또는 `py`).
2. 점검: `"<위 python>" "<정본폴더>\scripts\preflight.py" --net` — 패키지·키 존재·정본 스크립트 폴더·엔드포인트(키 값은 출력 안 함).
3. 키는 공용 `.env`(`<키 파일>`)에 이미 있다 — DART·KOSIS·ECOS·DATA_GO_KR·KRX·SEC_USER_AGENT, FMP 는 `~\mcp\fmp\.env`.
   **사용자에게 키를 다시 묻지 않는다.** 키를 출력·로그·산출물에 남기지 않고, 채팅에 붙인 키는 먼저 "노출됐으니 폐기·재발급"을 알린다.

## 1. 진입점

| 요청 | 진입점 |
|---|---|
| 국내 종목 IB 데이터셋까지 | `<프로젝트>\scripts\run_pipeline.py <종목|코드>` → 이후 `company-dataset-ib` |
| 국내 원자료만 / 정본 스크립트 폴더 없음 | `<정본폴더>\templates\kr_fetch_template.py <코드> [--fs both] --out <Downloads 하위 폴더>` |
| 미국(SEC 제출사) 원자료 | `<정본폴더>\templates\us_fetch_template.py <TICKER|CIK> [--short] --out <Downloads 하위 폴더>` → 디멘션·Form4·13F 는 정본 `sec_instance_facts.py`·`sec_form4.py`·`sec_13f_holders.py` |
| 유럽 상장사(ESEF: NL·DK·FR·IT·ES·SE·FI·BE·GB) | `<정본폴더>\templates\eu_fetch_template.py <LEI> --yahoo <본국심볼>` — 독일·스위스·아일랜드는 미수록 → SEC 20-F 또는 Yahoo |
| 국내 비상장 외감법인 | `<정본폴더>\templates\kr_unlisted_audit_template.py --corp <8자리>` (감사보고서 원문 본표 파싱·재작성 대조) |
| 4개 공시체계 예시(포트폴리오용) | `Downloads\<수집예시>\`(⚠2026-10-30 확인: 휴지통에 있음, 복원은 사용자 판단) — 삼성전자·NVIDIA·Novo·우아한형제들 워크북·빌더 |
| 20-F 외국기업·유럽/홍콩 peer·매크로·IR/컨콜·막힌 경로 | 통합 문서 §3.1 · §2.1 · §4 |

## 2. 반드시 지킬 규칙 (요약 — 근거는 통합 문서 §1)

- 분기 손익 = 3개월치. 4Q = FY − 3Q누적(해외 FY − 9M). 현금흐름은 YTD 차분. **차분은 같은 계정/태그 안에서만**, EPS·주식수는 차분 금지.
- 같은 기간 여러 값이면 **최신 정기보고서**(SEC: DEF 14A·8-K 제외). 한 행 = 한 태그, 보충한 칸은 표시.
- 계정은 ID(account_id / us-gaap 태그)로, 이름은 폴백. **포괄손익 귀속 행을 순이익으로 집지 않는다.**
- 0 ≠ None, "없음(확인)" ≠ "미수집". 단위·통화·결산월·시세 시각(장중≠종가)을 데이터와 함께 저장.
- 수집 직후 자산=부채+자본·지배+비지배·분기합=연간 검산 → 경고는 원문으로 확인(값을 맞추려고 고치지 않는다).
- 출처등급 1차/2차를 남기고 2차 수치는 1차와 대조.

## 3. 저장·보고

- **모든 산출물은 Downloads 아래**: 새 주제 `Downloads\<YYMMDD>_<주제>\`, 데이터셋 `Downloads\<YYMMDD>_<회사명>_Data\`, 사용자가 옮긴 폴더는 그 자리. 스크래치패드엔 진짜 임시물만.
- 보고는 "확인한 것 / 못 한 것 / 정정한 것"으로.
- 새 함정·막힌 경로는 정본 통합 문서 해당 표에 날짜·사례와 함께 추가, 템플릿을 고치면 실제 종목 1개로 재검증(`검증샘플_261014\` 방식).
