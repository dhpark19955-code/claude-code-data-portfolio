# 데이터 원천. 어떤 API에서 무엇을 받는가

데이터셋 파이프라인이 실제로 부르는 API와, 그 숫자로 매출 구성, FCF, 순차입금, EV를 어떻게 구하는지 정리했습니다. 원천은 모두 공개 자료입니다. 키는 환경변수(`DART_API_KEY` 등)나 키 파일에서 읽고 코드에 적지 않습니다.

표의 "코드 위치"는 두 가지로 적었습니다. [originals/pipeline/](../originals/pipeline/) 에 있는 파일은 파일명과 함수를, 스킬 문서에만 나오고 코드가 이 저장소에 없는 것은 "스킬 문서 기준 (코드 미수록: 스크립트명)"으로 적었습니다. 문서에 이름이 없는 엔드포인트와 파라미터는 적지 않았습니다.

## 1. 국내 원천과 API

### 공시 API (OpenDART)

호출 주소는 `https://opendart.fss.or.kr/api/<엔드포인트>` 이고 키는 `crtfc_key` 파라미터로 붙습니다. 보고서 코드는 `11013` 1분기, `11012` 반기, `11014` 3분기, `11011` 사업보고서입니다.

| 원천 | 엔드포인트·방식 | 받는 것 | 데이터셋 시트 | 코드 위치 | 키 |
|---|---|---|---|---|---|
| 회사 목록 | `corpCode.xml` (zip 안의 CORPCODE.xml) | 공시 고유번호, 회사명, 종목코드. 상장사만 남겨 7일 캐시 | - (종목 해석) | `dart_listed_fs.py` `load_corp_codes`. `run_pipeline.py` 0단계는 미수록 모듈 `dart_reports` 사용 | DART |
| 재무제표 | `fnlttSinglAcntAll.json` (`corp_code`, `bsns_year`, `reprt_code`, `fs_div`=CFS/OFS) | 재무상태표, 손익, 포괄손익, 현금흐름 전 계정. 표준계정 ID로 행을 맞춤. 연결이 없는 기간은 별도로 대체하고 경고 | 10~13 재무제표 | `dart_listed_fs.py` `fetch_fs` | DART |
| 주식 총수 | `stockTotqySttus.json` (`corp_code`, `bsns_year`, `reprt_code`) | 보통주 발행주식수, 자기주식수, 유통주식수. 최근 7개 연도 x 보고서 4종 | 16~2x 주식수, 23 Historical PER/PBR 분모 | `run_pipeline.py` 1단계 `dart()`. 이상치는 `share_utils.py` `clean_shares` | DART |
| 대량보유 보고 | `majorstock.json` (`corp_code`) | 5% 보고: 보고자, 주식수, 비율, 증감, 사유 | 50번대 지분, 33번대 오버행 | `run_pipeline.py` 1단계 | DART |
| 임원·주요주주 보고 | `elestock.json` (`corp_code`) | 임원·10% 주주 보유 변동 | 50번대 지분, 33번대 오버행 | `run_pipeline.py` 1단계 | DART |
| 배당, 최대주주 | `alotMatter`, `hyslrSttus` | 배당, 최대주주. 신규상장사는 status 013으로 빔 | 16~2x 배당, 50번대 지분 | 스킬 문서 기준 (코드 미수록: fetch_fnguide_dart_meta.py) | DART |
| 자기주식 취득·처분 | `tesstkAcqsDspsSttus` | 연도별 취득·처분. 무활동 연도는 `-` 더미 1행으로 와서 0행을 명시적으로 남김 | 16~2x 주식수, 33번대 오버행 | 스킬 문서 기준 (코드 미수록) | DART |
| CAPEX | 현금흐름표의 '…의 취득' 줄 | 유형자산, 무형자산, 투자부동산 취득 | 31_FCF | `run_pipeline.py` 1단계가 호출 (코드 미수록: fetch_capex.py) | DART |
| 공시 원문 | 원문 XML 일괄 수집. 기간(최근 3년)과 보고서 키워드를 넘김 | 사업·반기·분기보고서, 대량보유, 자기주식, CB·BW, 주식매수선택권, 공정공시, 증권신고서, 투자설명서 | 30번대 매출 구성, 33번대 오버행 | `run_pipeline.py` 1단계가 호출 (코드 미수록: dart_fetch_docs.py) | DART |
| 정기보고서 원문 표 | `document.xml` | XBRL 수치가 없는 기간(금융업 2023년 3분기 이전)의 재무제표 표 | 10~13 보강 | 스킬 문서 기준 (코드 미수록: 회사별 파서) | DART |
| 비상장 외감법인 | 감사보고서 원문 본표 파싱, 재작성 대조 | 비상장사 재무제표 | 10~13 | 스킬 문서 기준 (코드 미수록: kr_unlisted_audit_template.py, 원본 유실) | DART |

### 시세 (네이버 증권 JSON)

인증 없이 동작하고, `User-Agent`, `Referer: https://m.stock.naver.com/`, `Accept: application/json` 헤더만 붙입니다.

| 원천 | 엔드포인트·방식 | 받는 것 | 데이터셋 시트 | 코드 위치 | 키 |
|---|---|---|---|---|---|
| 현재가 | `polling.finance.naver.com/api/realtime/domestic/stock/{code}` | 최신 종가, 거래일, 종목명 | 20 밸류에이션 기준 종가 | `naver_market.py` `snapshot` | 없음 |
| 투자지표 | `m.stock.naver.com/api/stock/{code}/integration` | `totalInfos`: 시총(`marketValue`), PER, PBR, EPS, BPS, 52주 고저, 외국인소진율, 추정 PER·EPS | 14 투자지표, 시총 대조 | `naver_market.py` `snapshot` | 없음 |
| 일별 시세 | `api.stock.naver.com/chart/domestic/item/{code}/day` (`startDateTime`, `endDateTime`) | 일별 시가·고가·저가·종가·거래량 (수정주가) | 23 Historical PER/PBR | `naver_market.py` `sise_day` | 없음 |
| 일별 시세 폴백 | `fchart.stock.naver.com/siseJson.naver` (`symbol`, `requestType=1`, `startTime`, `endTime`, `timeframe=day`) | 위와 같음, 외국인소진율 포함 | 23 Historical PER/PBR | `naver_market.py` `sise_day` | 없음 |
| 투자자별 수급 | `m.stock.naver.com/api/stock/{code}/trend` (`pageSize=60`, `bizdate`) | 외국인·기관·개인 일별 순매수, 외국인보유비율 | 70 수급 250D | `naver_market.py` `investor_trend` | 없음 |
| 재무요약·컨센서스 | `m.stock.naver.com/api/stock/{code}/finance/annual`, `/finance/quarter` | FnGuide 항목 16개. `isConsensus=Y` 열이 추정치(다음 1개 회계연도만) | 14 투자지표·컨센서스 | `naver_market.py` `fin_summary`. 미수록 fetch_fnguide_dart_meta.py 도 같은 JSON 사용 | 없음 |

### 통계, IR, 기타

| 원천 | 엔드포인트·방식 | 받는 것 | 데이터셋 시트 | 코드 위치 | 키 |
|---|---|---|---|---|---|
| KOSIS | `statisticsParameterData.do`, 구조 확인용 `getMeta` (type=ITM/OBJ/PRD) | 산업 통계. 표마다 itmId·objL 구조가 다름 | 60번대 업종 KPI | 스킬 문서 기준 (코드 미수록: 회사별 스크립트) | KOSIS |
| IR 자료 | ① 회사 홈페이지 IR 메뉴(`curl` 로 href 목록) ② 공시 목록의 기업설명회(IR)개최·공정공시 ③ IRGO `m.irgo.co.kr/IR-COMP/<코드>` ④ 38커뮤니케이션(IPO) | IR 덱 KPI. 덱이 없으면 IR급 공시(주주환원계획, 배당정책, RSU 지급조건) | 60번대 업종 KPI | 스킬 문서 기준 (수기 정리, 7단계 체크리스트 항목). 전역 지침의 원문·IR 툴킷 dart_ir_fetch.py 는 미수록 | 없음 |
| ECOS, 관세청, 공공데이터포털, KRX | 스킬 설명과 키 목록에 이름만 있음. 엔드포인트는 정본 수집 문서(미수록)에 있음 | - | - | 스킬 문서 기준 (코드 미수록) | ECOS, DATA_GO_KR, KRX |

파이프라인 폴더 없이 국내 원자료만 받을 때의 단독 실행형 진입점은 `kr_fetch_template.py <코드> [--fs both] --out <폴더>` 입니다(코드 미수록).

## 2. 해외 원천

해외 원천은 fin-data-fetch 스킬에 정리돼 있고, 수집 스크립트는 이 저장소에 없습니다. 스킬에 적힌 원천 이름과 진입점만 옮겼습니다. 엔드포인트 주소는 스킬에 없어서 적지 않았습니다.

| 원천 | 엔드포인트·방식 | 받는 것 | 데이터셋 시트 | 코드 위치 | 키 |
|---|---|---|---|---|---|
| SEC EDGAR | companyfacts. 진입점 `us_fetch_template.py <TICKER 또는 CIK> [--short]` | 미국 SEC 제출사의 10-K·10-Q 수치. 한 행 = 한 태그, 같은 기간 여러 값이면 최신 정기보고서(DEF 14A·8-K 제외) | 10~13 (같은 양식) | 스킬 문서 기준 (코드 미수록: us_fetch_template.py) | SEC_USER_AGENT |
| SEC XBRL 인스턴스 | 인스턴스 문서의 사실값 | 디멘션이 붙은 수치 | - | 스킬 문서 기준 (코드 미수록: sec_instance_facts.py) | SEC_USER_AGENT |
| SEC 20-F, 6-K | 외국 발행사 제출 서류 | 외국기업 재무. 해외 기업 분기 실적은 실적발표 HTML 표를 파싱([사례 2](02_unstructured_parsing.md)) | 10~13 (같은 양식) | 스킬 문서 기준 (코드 미수록) | SEC_USER_AGENT |
| SEC Form 4 | 내부자 매매 보고 | 내부자 매매 | 50번대 지분 | 스킬 문서 기준 (코드 미수록: sec_form4.py) | SEC_USER_AGENT |
| SEC 13F | 기관 보유 보고 | 기관 보유 내역 | 50번대 지분 | 스킬 문서 기준 (코드 미수록: sec_13f_holders.py) | SEC_USER_AGENT |
| ESEF | `filings.xbrl.org`, LEI로 조회. 진입점 `eu_fetch_template.py <LEI> --yahoo <본국심볼>` | 유럽 상장사(NL·DK·FR·IT·ES·SE·FI·BE·GB) 재무 공시. 독일·스위스·아일랜드는 미수록이라 SEC 20-F 또는 Yahoo | 10~13 (같은 양식) | 스킬 문서 기준 (코드 미수록: eu_fetch_template.py, 원본 유실) | - |
| Yahoo | 위 템플릿의 `--yahoo <본국심볼>` | 본국 상장 심볼 기준 데이터 | - | 스킬 문서 기준 (코드 미수록) | 없음 |
| HKEX | 공시 검색 화면에서 PDF 링크 확보. 기존 검색 주소는 폐지([사례 2](02_unstructured_parsing.md)) | 홍콩 상장사 공시 | - | 스킬 문서 기준 (코드 미수록) | - |
| FINRA, FMP | 스킬 설명에 원천 이름만 있음 | - | - | 스킬 문서 기준 (코드 미수록) | FMP는 별도 키 파일 |

## 3. 데이터셋이 구하는 것

데이터셋은 숫자만 싣고 밸류에이션 모델링은 싣지 않습니다. 그래서 SOTP와 DCF는 아래 표에서 "데이터셋 밖"으로 표시했습니다.

| 항목 | 원천 | 방법 | 자동화 수준 | 샘플 시트 |
|---|---|---|---|---|
| 부문·제품별 매출 | 사업보고서 원문 'Ⅱ.사업의 내용' 표, 연결주석 '영업부문에 대한 공시'·'주요 고객' | 원문 XML 표를 `<tr>` 단위 셀 텍스트로 풀어 읽음. 부문합을 연결 매출과 대조하고 차이는 '연결조정' 행으로 맞춤. '사업부문 구분 없음' 회사는 품목별 표와 고객 집중도로 대신 | 반자동. 원문 수집은 자동, 표 해석은 회사별 스크립트. `run_pipeline.py` 7단계가 체크리스트로 출력 | 30_매출구성 |
| 지역별 매출 | 위와 같은 원문 표, 연결주석 영업부문 | 위와 같음 | 반자동. 7단계 체크리스트의 "부문·제품·지역별 매출" 항목 | 30_매출구성 |
| FCF (OCF − 유형·무형 CAPEX) | 공시 API 현금흐름표: 영업활동현금흐름, '…의 취득' 줄 | '…의 취득' 줄을 정규식으로 유형·무형·투자부동산으로 분류해 합산(금융·대여·자기주식 취득 제외). 분기는 누적값을 같은 계정 안에서 차분. 유출이 양수로 와서 abs() 대신 계정 성격으로 부호를 정함 | 자동 (fetch_capex.py, 미수록). CAPEX가 특정 연도에 급증하면 7단계가 경고하고, 사옥·토지 같은 비영업 투자인지 원문에서 확인해 주석 | 31_FCF |
| 순차입금 / 순현금 | 재무상태표 → `fs_summary.py` 재무요약 | 차입금 4행 − 현금성 3행 (정의는 4절). 회사마다 다른 차입금 라벨을 `fs_summary.py` 별칭 표와 회사별 `fs_alias.json` 으로 표준 행에 모음. 리스부채 포함 여부를 구분해 표기 | 자동 + 경고. 연결 연간 마지막 연도의 차입금 합이 0인데 부채비율(부채총계 ÷ 자본총계)이 80%를 넘으면 `run_pipeline.py` 7단계가 별칭 누락을 경고 | 32_순차입금 |
| EV 산출 (브리지 항목) | 시가총액, 32_순차입금, 재무요약의 비지배지분·관계기업투자 행 | 시가총액에 순차입금과 비지배지분을 더하고 관계기업투자를 뺌. 어느 분기 BS 잔액을 썼는지 노트에 적음 | 자동 (시트 수식). 실제 빌드 스크립트 build_base_dataset_ib.py 는 미수록이라 항목 구성은 샘플 기준 | 20_밸류에이션 |
| 밸류에이션 배수와 역산 검증 | 10번대 재무, 시가총액, 네이버 `integration`(대조용), 일별 시세 | PER·PBR·EV/EBITDA·EV/Sales 행마다 "금액 × 배수"를 다시 계산해 같은 시총으로 돌아오는지 봄. `dataset_audit.py` `check_multiples` 가 중앙값에서 15% 넘게 벗어난 행을 표시. Historical PER/PBR은 일별 종가 ÷ TTM EPS·BPS | 자동 (`dataset_audit.py`, `hist_valuation.py`) | 20_밸류에이션, 99_Check |
| SOTP | 데이터셋 밖 (별도 모델 파일) | 데이터셋 README의 "제외한 것"에 밸류에이션 모델링(DCF·SOTP·민감도)이 있음. 모델 워크플로우(model-builder 스킬)에서 입력 파일 → 빌드 → 3관문으로 만듦. 전역 지침상 연결·별도 기준, EBITDA 정의, 순차입금의 리스부채 포함 여부, 지분가치 기준을 명시 | 데이터셋 자동화 대상 아님 | 모델 샘플의 SOTP 시트 |

## 4. 정의와 기준

- **FCF** = 영업활동현금흐름 − CAPEX. CAPEX = 유형자산의 취득 + 무형자산의 취득. FnGuide CAPEX는 유형만이라 값이 다릅니다. 투자부동산 취득은 따로 분류하고, 사옥 같은 비영업 CAPEX는 주석으로 분리합니다.
- **순차입금** = (단기차입금 + 유동성장기차입금 + 사채 + 장기차입금) − (현금및현금성자산 + 단기금융상품 + 기타유동금융자산). 기본은 리스부채 제외입니다. 리스부채를 넣은 광의 순차입금은 별도 행으로 구분합니다. 샘플 워크북의 32_순차입금은 이를 단순화해 단기차입금 + 장기차입금 − 현금및현금성자산만 씁니다.
- **순현금** = −순차입금. 순차입금이 음수면 순현금이고, 순차입금 행에서는 괄호로 보입니다.
- **EV 브리지** = 시가총액 + 순차입금 + 비지배지분 − 관계기업투자. 밸류에이션 시트는 최신 분기 BS와 FY 또는 TTM 손익을 섞으므로 어느 분기 잔액인지 노트에 적습니다.
- **PER, PBR** = 지배주주순이익, 지배주주지분 기준. 비지배지분이 없는 회사는 공시에 지배 귀속 행이 없는 것이 정상이라 당기순이익, 자본총계로 대신하고 주석에 적습니다. Historical PER/PBR의 분모는 유통주식수(발행 − 자기주식)라서 가중평균 발행주식수를 쓰는 네이버·FnGuide 값과 다릅니다.
- **시가총액**: 스킬 원문은 "시총은 상장주식수×종가로 직접 계산해 기준일을 적고, 네이버 marketValue는 대조용"입니다. `marketValue` 가 전일 종가(`lastClosePrice`)가 아니라 실시간 폴링 가격 기준일 때가 있어서입니다. 샘플 20_밸류에이션은 발행주식수에서 자기주식을 뺀 주식수에 기준일 종가를 곱합니다.

## 5. 막힌 경로와 대체

- **네이버 금융 HTML**: `finance.naver.com` 의 main·sise_day·frgn 페이지는 2026-09 개편 후 표와 셀렉터가 잡히지 않습니다. 1절의 JSON 엔드포인트로 다시 썼습니다. `api.stock.naver.com/stock/{code}/basic` 은 409를 돌려줍니다.
- **수급 `trend` API**: `page` 를 무시하고 `pageSize` 가 60을 넘으면 JSON이 아닌 응답을 줍니다. 마지막 날짜의 하루 전을 `bizdate` 로 넘겨 60거래일씩 거슬러 갑니다.
- **컨센서스**: wisereport `ajax/cF1001.aspx` 는 토큰이 필요하고, 토큰이 있어도 단일 값 더미 표를 돌려줍니다. 네이버 `finance/annual|quarter` JSON으로 대체했습니다.
- **KOSIS**: 키가 base64라 끝의 `=` 가 빠지면 err 11이 납니다. 키 로더가 빠진 `=` 를 붙입니다. 법무부 표는 파라미터 API로 열리지 않아 쓰지 않습니다.
- **공시 API**: URL에 `.json` 이 빠지면 status 101을 조용히 돌려줍니다. 신규상장사는 `fnlttSinglAcntAll`·`alotMatter`·`hyslrSttus` 가 status 013으로 비어 감사보고서 파싱으로 보강하고, 금융업은 XBRL이 2023년 3분기부터라 그 이전은 원문 `document.xml` 표를 파싱합니다.

## 6. 관련 파일

| 구분 | 파일 |
|---|---|
| 파이프라인 | [run_pipeline.py](../originals/pipeline/run_pipeline.py), [dart_listed_fs.py](../originals/pipeline/dart_listed_fs.py), [naver_market.py](../originals/pipeline/naver_market.py), [fs_summary.py](../originals/pipeline/fs_summary.py) |
| 파생·감사 | [hist_valuation.py](../originals/pipeline/hist_valuation.py), [share_utils.py](../originals/pipeline/share_utils.py), [dataset_audit.py](../originals/pipeline/dataset_audit.py) |
| 스킬 원문 | [company-dataset-ib](../originals/claude/skills/company-dataset-ib/SKILL.md), [함정 모음](../originals/claude/skills/company-dataset-ib/references/pitfalls.md), [fin-data-fetch](../originals/claude/skills/fin-data-fetch/SKILL.md) |
| 샘플 | [scripts/build_ib_sample.py](../scripts/build_ib_sample.py) |
| 관련 문서 | [IB 양식 규격](../docs/06_ib_format_spec.md), [모델 빌드](../docs/07_model_build.md) |
