# 함정 모음 — 증상 → 원인 → 해법

전부 2026-09 GTF·NAVER·에이플러스에셋 데이터셋 작업에서 실제로 밟은 것들이다.

## Excel이 "파일이 손상되었습니다"로 열기를 거부

| 증상 | 원인 | 해법 |
|---|---|---|
| openpyxl로는 열리는데 Excel COM `Workbooks.Open`이 실패 | `freeze_panes`를 걸었다가 `"A1"`로 되돌림 → `<pane>`은 지워지고 `<selection pane="bottomRight">`만 남음 | 해제하지 말고 애초에 안 건다(`if ws.title != "00_README": ws.freeze_panes = ...`) |
| 열리긴 하는데 목차 하이퍼링크가 조용히 사라짐 | `cell.hyperlink = "#'시트'!A1"` → 외부 관계(rels)로 기록 | `Hyperlink(ref=cell.coordinate, location="'시트'!A1")` |
| 라벨/헤더 셀이 비어 보임 | 문자열이 `=`로 시작해 수식으로 해석, 캐시값 없음 | 선행 `=` 제거 |
| 수식이 든 시트를 넣은 뒤 COM Open 실패 | 데이터가 구간보다 짧아 `SUM(F-228:F22)` 같은 **음수/0 행 참조** 수식이 생김(신규상장 22일에 250일 누적) | 시작 행을 `max(첫 데이터행, 끝행-N+1)`로 클램프 |

**진단 순서**: ① COM Open 시도 → ② `CorruptLoad=2`로 열어 SaveAs 후 원본과 zip 파트 비교(사라진 파트가 범인) → ③ 시트 하나씩 남기고 저장·열기 → ④ 그 시트에서 fill/font/alignment/hyperlink/값 제거 반복. `<sheetViews>` XML을 정상 시트와 직접 대조하면 가장 빠르다.

## 수식 값이 None / 0으로 보임
openpyxl은 수식 캐시를 쓰지 않는다. 저장 후 Excel COM으로 `CalculateFullRebuild` → `Save`. `data_only=True`로 읽으려면 그 뒤에.
**한 프로세스에서 3개 이상 열면 `-2147417848 (개체 연결 끊김)`** → 종목마다 프로세스(인스턴스)를 분리해 돌린다(`recalc_one.py`).

## 한글 경로·파일명
- bash argv로 한글 경로 → cp949 깨짐(`FileNotFoundError`에 `���` 포함). 스크립트가 **패턴만 받아 내부 glob**.
- 프로젝트 하위 파일명이 NFD(분해형) → `Path("…한글.xlsx")` 직접 열기 실패. `os.walk` + `unicodedata.normalize("NFC", fn)`로 매칭.
- 콘솔 출력이 깨지는 건 표시 문제일 뿐(`sys.stdout.reconfigure(encoding="utf-8")`). 파일은 정상.

## 원본 xlsx의 raw 문자열
`r"data\폴더\"` 처럼 raw string이 `\`로 끝나면 SyntaxError. 끝 백슬래시를 뺀다.

## DART
- 엔드포인트 URL에 `.json`이 빠지면 status 101을 **조용히** 반환(빈 결과). 반드시 붙인다.
- `tesstkAcqsDspsSttus`: 활동 없는 연도를 빈 리스트가 아니라 **필드 전부 `-`인 더미 1행**으로 준다. `acqs_mth3=="총계"` 필터가 안 걸려 회사가 통째로 빠짐 → 연도별 0행을 명시적으로 남긴다.
- 분기 IS는 `thstrm_amount`(3개월치), 4Q = FY − 3Q누적. CF는 누적이므로 차분. CAPEX = 유형자산의 취득 + 무형자산의 취득(FnGuide CAPEX는 유형만 → 값이 다름).
- 부문정보 단순합은 매출은 연결과 맞지만 영업이익·자산은 내부거래·본사미배분으로 다르다. '연결조정' 행을 넣어야 tie.

## 프로젝트 루트 하드코딩
`glob(r"<HOME>\Downloads\260908_*")` 같은 절대경로는 사용자가 폴더를 `Downloads\Work\` 아래로 옮기는 순간 `IndexError`로 죽는다(2026-09-22 실제 발생).
`ib_dataset.project_base()`는 `__file__` 기준(scripts/의 상위)으로 잡고, 못 찾으면 `~/Downloads/**/260908_*` 재귀 glob으로 폴백한다. 새 스크립트도 같은 방식으로.

## 네이버 재무요약(FnGuide 항목)·컨센서스
- **되는 것**: `https://m.stock.naver.com/api/stock/{code}/finance/annual` · `/finance/quarter` (JSON). `trTitleList[].isConsensus=="Y"` 열이 추정치(E).
  항목 16개(매출액·영업이익·당기순이익·지배/비지배·이익률·ROE·부채비율·당좌비율·유보율·EPS·PER·BPS·PBR·주당배당금). `scripts\fetch_fnguide_dart_meta.py`가 이걸로 `financial_summary.xlsx`를 만든다.
- **안 되는 것**: wisereport `ajax/cF1001.aspx`는 `encparam`+`id` 토큰이 필요하고 토큰이 있어도 단일 값 더미 표를 돌려준다. 시간 쓰지 말 것.
- 신규상장·미커버 종목은 E열이 전부 `-`/None — 누락이 아니라 컨센 부재. README에 그렇게 쓴다.

## 계정명이 회사마다 다르다 (2026-09-22 매커스에서 밟음)
- **차입금**: '단기차입금' 대신 '유동 차입금(사채 포함)', '비유동차입금(사채 포함)의 유동성 대체 부분', '장기차입금(사채 포함), 총액'으로 쓰는 회사가 있다. `fs_summary.py` SPEC 별칭에 없으면 재무요약에 빈칸 → 순차입금 시트가 **차입금 0·순현금**으로 잘못 나온다. 부채비율이 높은데 차입금이 0이면 이 함정을 의심할 것.
- **CAPEX**: '유형자산의 취득' 한 줄이 아니라 '건설중인자산의 취득·기계장치의 취득·사무용비품의 취득·컴퓨터소프트웨어의 취득·영업권 이외의 무형자산의 취득'으로 쪼개는 회사가 있다. `fetch_capex.py`는 '…취득' 줄을 정규식으로 유형/무형/투자부동산 분류해 합산한다(금융·대여·자기주식 취득은 제외).
- **매출**: 분기 IS 라벨이 '수익'만 있는 회사(매커스). SPEC에 '수익'·'매출' 별칭 추가돼 있음.

## DART 원문 표 파싱
- 사업보고서의 '주식의 총수'·'대량보유 보유비율' 같은 표는 셀이 중첩·병합돼 `pandas.read_html`이나 단순 td 순회로는 값이 빈 행으로 나온다. `<tr>`마다 `td/th` 텍스트를 모아 행 단위로 읽고, 그래도 안 잡히면 `soup.get_text()`를 공백 정규화한 뒤 키워드 주변 300자를 정규식으로 뽑는 것이 빠르다(scratch `parse_docs2/3.py` 패턴).
- 5%·임원 보유 변동은 원문 파싱보다 **`majorstock.json`·`elestock.json` API**가 구조화돼 있어 훨씬 빠르다(보고자·주식수·비율·증감·사유).
- 파일명에 `[기재정정]` 대괄호가 있으면 `glob`이 문자클래스로 해석해 못 찾는다 → `os.listdir` 부분일치로 찾을 것.
- IR 자료(설명회 PPT)는 회사 홈페이지에 없는 경우가 많다(매커스·기도산업, 2026-09-22 확인). 그럴 땐 사업보고서 '사업의 내용'·투자설명서로 대체하고 README에 "IR 덱 미게시"라 쓴다.

## IR 자료 탐색에서 배운 것
- 회사 사이트가 `https`가 없거나(매커스는 http만), WebFetch가 `ENOTFOUND/ECONNRESET`을 내면 `curl -L -A Mozilla`로 받아 `href`를 정규식으로 뽑는 게 빠르다. IR 메뉴가 "주가정보·공시정보(DART 미러)·공고(주총소집)"만인 경우가 대부분.
- 38커뮤니케이션 IPO 페이지를 WebFetch 요약으로 읽으면 소형 모델이 숫자·업종을 잘못 요약한다(기도산업을 '자동차부품·2026/06 상장'으로 답함). IPO 수치는 증권발행실적보고서·투자설명서 원문에서만 가져올 것.
- IR 덱이 없어도 **IR급 문서**는 있다: 공정공시(주주환원계획·소각계획), 반기·사업보고서 III.6 배당정책 서술(정량 목표 문장), RSU/스톡옵션 부여 표와 **지급조건**(회사가 스스로 쓴 3개년 매출·OPM 목표 = 사실상 가이던스), 원재료 매입·판매경로 표. 이것들을 시트로 옮기고 원문 rcept_no를 출처에 적는다.
- 매커스 주주환원율 정의는 '배당 + 자기주식 매입 후 소각 ÷ 별도 순이익'이라 매입 연도·소각 연도가 어긋난다. 정책 비율을 재현할 땐 회사 정의를 그대로 쓰고 근사치임을 밝힌다.

## 신규상장사(사업보고서 없음)
DART `fnlttSinglAcntAll`·`alotMatter`·`hyslrSttus`가 status 013으로 빈다. 연간 재무는 감사보고서(증권신고서 첨부) 파싱으로, 지분·주식수는 증권신고서/반기보고서로 채운다(기도산업 사례 `kido_fs_summary.py`).
반기보고서 IS/CF 열은 3개월치인지 누적인지 값 크기로 판별해 열 이름에 `(누적)`을 붙인다.

## 네이버 금융 — 2026-09 개편 후 확정된 JSON 5종 (`scripts\naver_market.py`가 전부 구현)
- HTML(`finance.naver.com/item/main·sise_day·frgn`)은 UTF-8 전환·마크업 변경으로 `read_html`이 표를 못 찾고 `#_per` 류 셀렉터가 전부 None. 쓰지 말 것.
  - 현재가·거래일·종목명: `polling.finance.naver.com/api/realtime/domestic/stock/{code}` → `datas[0].closePrice/localTradedAt/stockName`
  - 시총·PER·PBR·EPS·BPS·52주·외국인소진율: `m.stock.naver.com/api/stock/{code}/integration` → `totalInfos[].code` (`marketValue`는 `lastClosePrice`가 아니라 폴링 종가 기준일 때가 있어 상장주식수×종가로 재계산)
  - 일별 OHLCV: `api.stock.naver.com/chart/domestic/item/{code}/day?startDateTime=YYYYMMDD0000&endDateTime=YYYYMMDD2359` (폴백 `fchart.stock.naver.com/siseJson.naver`, 외국인소진율 포함)
  - 투자자별 순매매(외국인·기관·개인): `m.stock.naver.com/api/stock/{code}/trend?pageSize=60[&bizdate=YYYYMMDD]` — **`page`는 무시되고 pageSize>60은 JSON이 아닌 응답**. 마지막 bizdate−1일을 `bizdate`로 넘겨 60일씩 거슬러 간다.
  - 재무요약(FnGuide 16항목, 컨센 열 `isConsensus=Y`): `m.stock.naver.com/api/stock/{code}/finance/annual|quarter`
  - 헤더는 `User-Agent`+`Referer: https://m.stock.naver.com/`+`Accept: application/json`. `api.stock.naver.com/stock/{code}/basic`은 409.
- `lastClosePrice`는 전일, 폴링의 `closePrice`가 최신 종가. 시총은 **상장주식수×종가**로 직접 계산하고 기준일을 노트에 적는다.
- 커버리지 없는 종목은 `estimatePer`가 None — 컨센서스 기반 밸류가 성립하지 않는다고 명시.

## KOSIS
- 키는 base64라 끝에 `=`가 있다. 빠지면 `err 11 유효하지않은 인증KEY`.
- 테이블마다 파라미터 구조가 다르다: 관광공사 입국(`DT_TRD_TGT_ENT_AGG_MONTH`)은 `itmId=목적코드, objL1=ALL(국적)`; 내국인출국(`DT_SEX_DEP_AGG_MONTH`)은 `itmId=측정항목 13103314008T01, objL1=성별코드`. `getMeta type=ITM/OBJ/PRD`로 먼저 본다. 법무부(`DT_091_111_…`) 표는 파라미터 API로 안 열린다(objL 누락 오류 반복) — 시간 쓰지 말 것.
- 통계청 국제이동(`DT_1B28022`)은 90일 체류 기준 인구이동이라 월 2.7만 명 규모 — 관광 입국(월 200만)과 정의가 다르다. 관광 분석에 쓰면 안 된다.
- 외국인 **출국** 월별은 KOSIS API로 안정 추출이 안 된다. 입국↔출국 1개월 시차 동행을 근거로 입국을 프록시로 쓰고 그렇게 적는다.

## Historical PER/PBR · 컨센서스
- PER 분모를 뭘로 잡았는지에 따라 2배 가까이 다르다: `hist_valuation.py`는 **유통주식수(발행−자기주식)**, 네이버·FnGuide EPS는 가중평균 **발행주식수(자기주식 포함)**. 자기주식 29%인 매커스는 8.3x vs 13.7x. 데이터셋 시트와 네이버 표기를 나란히 둘 때 반드시 기준을 적는다.
- TTM은 8분기 창(재무요약 분기 시트) 안에서만 4개 분기 합이 나온다. 그 이전은 FY 순이익 단일값이라 연도 경계에서 계단이 크다. 실적 반영 시점(분기 +45일·결산 +80일)은 근사치.
- 적자(TTM EPS≤0)면 PER 공란 → 관측수가 줄어 통계가 얇아진다(GTF 2021~23). 그럴 땐 PBR 밴드가 기준.
- 신규상장사는 DART `stockTotqySttus`에 상장 전 주식수만 있어 `--shares-override 상장일:상장후주식수`가 필요하다(기도산업 5,860,707).
- 네이버 `finance/annual`의 컨센서스는 **다음 1개 회계연도(E)만**, 추정기관 수·변화율 없음. 정식 컨센(다년·기관수·괴리율)은 FnGuide DataGuide/S&P Capital IQ 등 유료. 커버리지 없는 종목(에이플러스에셋·기도산업)은 E열 전부 `-`.

## 단위·빈티지
- 같은 회사 소스 xlsx 안에서 시트 제목 단위가 값과 다른 사례(억원 값에 '십억원'). 대표값 하나를 DART와 대조해 라벨을 확정한다.
- IR 덱은 같은 과거 KPI를 덱마다 다르게 재기재한다(설계사 재적·신규위촉·업계평균). 덱 빈티지별 열로 나눠 싣고 하나로 합치지 않는다.
- 밸류에이션 시트는 대개 최신 분기 BS + FY/TTM 손익을 섞는다. 어느 분기 잔액인지 노트에 적는다.
- 프로젝트 폴더가 이동하기도 한다(`Downloads\260908_…` → `Downloads\Work\260908_…`, 작업폴더 → `Downloads\<작업폴더>\…`). 절대경로 하드코딩 대신 `glob("Downloads\\**\\260908_*")` 류로 잡는다.

## DART 주식수 ×1000 오기 (2026-09-23 월덱스)
- `stockTotqySttus` 가 2025 1Q·반기 보고서에서 발행주식총수를 16,510,993,000주(정상 16,510,993주의 1000배)로 줬다. 원천 오류라 그대로 쓰면 그 기간 EPS·BPS가 1/1000 → Historical PER 최대 15,710배·평균 703배(중앙값 9.8배), 밴드가격이 수천만원으로 튄다. **재계산 errors=0, 기존 감사 ★0으로 통과했다** — 수식 오류가 아니라 입력 오류라서.
- 방어: `scripts\share_utils.clean_shares` — 전 기간 중앙값 대비 50배 이상 벗어난 값을 ÷1000·÷100 해서 중앙값 ±20%면 보정, 아니면 제외하고 로그. `run_pipeline`(json 저장 전), `hist_valuation.shares_lookup`, `build_base`(시총) 세 곳에서 모두 거친다.
- 감지: `dataset_audit.check_hist` — Historical 시트 PER/PBR 통계에서 평균/중앙값 >3 또는 최대/중앙값 >20 이면 ★. 기존 6개 데이터셋은 오탐 0.
- 교훈: 통계 시트는 **평균과 중앙값을 나란히 보고** 벌어져 있으면 원천부터 의심한다.

## 지배주주순이익 라벨 변형 (2026-09-23 월덱스)
- '지배기업의 소유주에 **귀속될** 당기순이익'(‘에게 귀속되는’ 아님)이라 SPEC 별칭에 안 걸려 행이 통째로 비었다. 별칭 추가 + `fs_summary.build` 에 정규식 폴백(`^지배기업.*(소유주|주주).*(순이익|순손익)`, '포괄' 제외)을 넣었다.
- 비지배지분이 정말 없는 회사는 DART가 지배 귀속 행을 안 주는 게 정상 → `build_base` 가 PER=당기순이익·PBR=자본총계로 대체하고 20번 시트 주석에 적는다. 구분은 “비지배지분 행이 비었는가”로 한다.


## 보험·금융업 (2026-10-12 코리안리)
- **증상** 2023 2Q 이전 재무가 전부 비고 fnlttSinglAcntAll 이 status 013. **원인** 금융업은 DART XBRL 구조화 데이터가 2023 3Q 부터만 있다. **해법** 정기보고서 원문(document.xml) 연결 포괄손익계산서·재무상태표 표를 직접 파싱(`koreanre_hist_parse.py` 패턴), 겹치는 2023 3Q 를 XBRL 과 대조해 검증.
- 2018 이전 원문은 utf-8 선언이지만 실제 cp949 → utf-8 디코드 실패 시 cp949. 셀 태그는 TD/TE/TU 혼용.
- 보험사는 표준 시트의 매출·매출원가·차입금 행이 비는 게 정상(파이프라인 ⚠ 2건 무시 가능) — 보험손익·투자손익 회사별 시트로 대체. CF 라벨 '영업활동순현금흐름' 별칭은 fs_summary 에 추가됨.
- IFRS17 원문 서식의 '보험영업수익'(보험수익만)과 XBRL '보험영업수익'(보험수익+재보험수익)은 다르다.
- openpyxl 로 `T.DIST.2T` 등 2010 이후 함수를 쓰면 `_xlfn.` 접두 없이는 #NAME? → `TDIST(x,df,2)` 같은 구 함수로 쓴다.
