# g2b_costdb — 나라장터 발주공고 → LIMAC 공사비 DB 파이프라인

표2 시설용도(미술관·문화예술회관·수장고·공공주택·경기장·주차장·화장시설 등)의 공사 발주공고를
나라장터 Open API로 전량 수집하고, 첨부 공고문(HWP/PDF)을 파싱·LLM 추출하여 Excel DB로 정리한다.
설계 배경과 단계별 상세는 `docs/나라장터_공사비DB_자동화_울트라플랜.md` 참조. 세션별 붙여넣기용 지시문은 `PROMPTS.md`.

## 1. 준비 (한 번만)
```powershell
git clone https://github.com/seio0701/g2b_costdb.git C:\DB_WORK\g2b_costdb   # 처음 한 번
cd C:\DB_WORK\g2b_costdb                   # ★ 모든 명령은 이 폴더에서 실행 (config.yaml 이 있는 곳)
pip install -r requirements.txt
pip install pyhwp                          # 선택: HWP 폴백(hwp5txt)
pip install formulas                       # 선택: Excel 수식 오류 점검(tests/check_excel.py). 없으면 Excel 에서 직접 열어 확인
# 공공데이터포털 → 「조달청_나라장터 입찰공고정보서비스」 활용신청(자동승인) → 마이페이지에서 '일반 인증키(Decoding)' 복사
#   선택: 「조달청_나라장터 낙찰정보서비스」도 활용신청하면(같은 키) 낙찰금액·낙찰률 참고 컬럼을 함께 수집 → config.yaml api.use_awards: true
setx G2B_SERVICE_KEY "발급받은키"           # 영구 설정(새 터미널·새 Claude Code 세션부터 적용). 키를 채팅·파일에 붙여 넣지 말 것
$env:G2B_SERVICE_KEY="발급받은키"           # 지금 열려 있는 터미널에서 바로 쓰려면 이것도 실행
setx ANTHROPIC_API_KEY "..."               # S6 LLM 추출 단계에서만 필요
python -m g2b_costdb.pipeline doctor       # 파이썬·패키지·키 존재 여부(값은 안 보임)·네트워크 점검
```
`config.yaml`: 수집 기간(`period`), 경로(프로젝트 폴더 기준), 모델명, 임계값. `keywords.yaml`: 표2 검색어·변형어·제외어·분류 규칙.

## 2. 실행 순서
| 단계 | 명령 | 산출물 | API 호출 |
|---|---|---|---|
| 환경 점검 | `python -m g2b_costdb.pipeline doctor` | 콘솔 보고 | 1회(키 없이 연결 확인) |
| 필드 확인 | `python -m g2b_costdb.pipeline probe --ym 2026-08` | 응답 필드명·표본 + `config.yaml fields` 매핑 대조 결과(없는 항목 표시), 면허제한 표본, (`use_awards` 켜면) 낙찰 표본과 `[award_*]` 매핑 상태 | 3~6회 |
| S1 전량수집 | `python -m g2b_costdb.pipeline collect` | `data/raw/*.jsonl`, `data/notices_all.parquet`, `data/bsis_all.parquet`, (`use_awards`) `data/award_all.parquet` | 월×페이지(2026-08 기준 공고 약 8,800건=9페이지, 기초금액 약 6,200건=7페이지 → 93개월 약 1,500회). 일일예산 초과 시 `[미완료]` 표시와 남은 월을 보여주며 정상 종료 → 다음날 같은 명령 |
| (선택) 용역 공고 | `python -m g2b_costdb.pipeline probe --servc --ym 2026-08` 로 응답·건수 확인 → `python -m g2b_costdb.pipeline collect --servc` | `data/raw/servc_YYYYMM.jsonl`, `data/servc_all.parquet` (폐기물처리 용역·건설사업관리·설계용역만 보존. 사업의 폐기물 비용·공고 완료 보조지표·공사개요 보완용) | 월 15~26회, 전 기간 약 1,300~2,400회(일일 예산 안에서 여러 날 이어받기) |
| S2 후보탐색 | `python -m g2b_costdb.pipeline discover` | `output/facility_candidates.xlsx` (노란 셀 검수) | 0 |
| **검수** | Excel에서 `검수_포함여부 / 검수_시설명 / 검수_별칭 / 검수_수요기관` 수정 후 저장·닫기. `추정가격_합계` 내림차순으로 '포함'·'검토필요'부터 본다. 지자체명만 붙은 공동주택('구례군 국민임대아파트')에 단지가 여럿 섞여 있으면 `검수_시설명`에 단지·지구명을 넣거나 별칭으로 나눈다. 같은 시설이 두 행(수요기관이 다름)이면 `시설ID` 를 같게 적어 합친다 | — | — |
| S3 재검색 | `python -m g2b_costdb.pipeline research` | `data/notices_research.parquet` (전 공종 공고), 시설별 공종 분포 표. 정확한 이름으로 없으면 어절 간격을 허용한 예비 검색. 검수 행 여러 개에 **같은 시설ID** 를 적으면 한 시설로 병합(지자체·도시공사로 나뉜 같은 시설) | 0 (기본). `config.yaml api.use_license_limit: true` 로 바꾸면 공고별 면허제한 조회(공고 수만큼) |
| S4 중복정리 | `python -m g2b_costdb.pipeline dedup` | `data/notices_hist.parquet`, `data/notices_latest.parquet`, `data/logs_dedup.parquet` | 0 |
| S5 첨부수집 | `python -m g2b_costdb.pipeline attach` | `data/files/`, `data/text/`, `data/texts.json`, `data/notes_attach.parquet` | 0 (파일 다운로드만, 중단 후 재실행 시 이어서. 텍스트가 없는 공고는 매번 다시 시도하고, HWP 백엔드를 나중에 설치했으면 `attach --retry-failed` 로 실패 파일이 있는 공고를 일괄 재처리. `attach --report` 는 처리 없이 요약과 텍스트 없는 공고의 원인(`data/notes_notext.parquet`)만 출력. 대표 공고가 정정·변경 차수여도 같은 공고번호의 다른 차수 첨부(원 공고문·현장설명서)를 합쳐 처리) |
| S6 LLM추출 | `python -m g2b_costdb.pipeline extract` → 견적 확인 → 아래 세 방식 중 택일 (`--limit N` 을 붙이면 건축 공종·추정가격 큰 순 N건만: 시범용. `--redo-low` 는 신뢰도 low 이거나 건축인데 연면적이 없는 기존 결과를 다시 대상에 포함) | `data/llm_docs.json`, `data/logs_verify.parquet` | Claude API(`--yes` 없이는 비용 견적만 출력) |
| S7 Excel | `python -m g2b_costdb.pipeline excel` | `output/공사비DB.xlsx` | 0 |

- **S6 추출 방식 세 가지** (뒤 단계는 동일하게 이어짐, 섞어 써도 됨 — 이미 추출된 공고는 건너뜀):
  | 방식 | 명령 | 특징 |
  |---|---|---|
  | 즉시 호출 | `extract --yes` | 한 건씩 바로 처리. 정가 |
  | 배치(권장) | `extract --batch --yes` 로 제출 → 나중에 `extract --batch` 로 수거 | Message Batches, **50% 할인**, 대개 1시간·최대 24시간. 제출 상태는 `data/llm_batch.json` |
  | 2차 추출(공종별 금액) | `extract --batch --yes --trades` 로 1차 재제출과 함께 제출 → `extract --batch --trades` 로 수거 (단독: `extract-trades [--yes|--batch [--yes]|--refetch|--limit N]`) | 건축·통합발주 공고의 첨부 내역서 집계표에서 공종별 금액을 읽어 04b 시트와 05 의 문서배분_*(다른 공종 공고 안에 묶인 금액의 참고 배분) 열에 반영. 입력은 집계표 후보 행만(≤20,000자)이라 건당 비용이 1차의 1/3 수준. 내역서 Excel 도 공종별 금액 표도 없는 공고는 "집계표 없음"으로 건너뜀(과금 없음). 캐시 `data/llm_trades.json`, 배치 상태 `data/llm_trades_batch.json`(청크마다 즉시 기록) |
  | 보강 추출(기타금액·면세 문구·공사개요 세부) | `extract-extra` 로 견적 확인 → `extract-extra --batch --yes` 제출 → `extract-extra --batch` 수거 (즉시: `--yes`, 시범: `--limit 30`, 공종 한정: `--only 전기,건축`). 1차 결과는 건드리지 않고 `data/llm_extra.json` 에 저장 → 01 공사개요(조경면적·건폐율·용적률·높이·주차대수·설계특이사항)와 04 기타금액ⓒ·면세 문구에 반영 | 대표 공고 전체 기준 1차 추출과 비슷한 규모(견적 출력) |
  | 파일 인수인계 | `extract --export` → `data/llm_in/<공고키>.txt` 를 Cowork·사람이 읽고 `data/llm_out/<공고키>.json` 작성 → `extract --import`(반영한 txt·json 은 `data/llm_done/` 으로 이동) | API 비용 0. 시범 검증·예외 처리용. 지시문은 `data/llm_in/README_지시문.md` |
- `discover` 를 다시 실행해도 시설ID와 **사용자가 고친** 검수 값은 (시설키, 수요기관)이 같은 행에 그대로 이어진다(기본값 그대로였던 칸은 새 규칙으로 다시 계산). 검수를 시작하기 전에 규칙을 바꿨다면 `discover --fresh` 로 검수 칸을 전부 초기화(시설ID는 유지)할 수 있다.
- 오늘이 속한 달은 완료 표시를 하지 않고 매번 다시 받는다. 특정 달을 다시 받으려면 `data/raw/notices_done.txt`(또는 `bsis_done.txt`)에서 해당 월을 지우고 `collect` 를 실행한다(파일이 있으므로 캐시를 우회해 새로 받음).

## 3. 검증(합성 데이터, 네트워크·키 불필요)
```powershell
python -m tests.test_dedup_and_excel     # 분류·중복정리·Excel 수식 점검 → output/공사비DB_샘플.xlsx(저장소의 sample_output 사본은 --sample 을 줬을 때만 갱신), 이어서 회귀 테스트 실행
python -m tests.test_regressions         # 회귀 테스트만
python -m tests.test_e2e_mock            # 가짜 나라장터 서버로 probe→collect→discover→research→dedup→attach→extract→excel 전 단계 통합 실행(약 20초)
python -m tests.check_excel output/공사비DB.xlsx   # 생성된 Excel 의 수식 오류 점검(LibreOffice 또는 formulas 패키지)
```
`sample_output/공사비DB_샘플.xlsx`는 **가상 공고**로 만든 스키마 예시다(실데이터 아님). 통합 테스트의 가짜 서버(`tests/mock_g2b.py`)는
포털 오류 응답(키 오류 XML, 데이터 없음 03), 면허제한 필드명 차이, 별도 End Point 의 낙찰정보서비스(재개찰 2행), HTML 응답 첨부, cp949 텍스트 등을 재현한다.

## 4. 주요 설계 포인트
- **전량 수집 후 로컬 검색**: 키워드마다 API를 반복 호출하지 않고 월 단위로 공사 공고를 전량 내려받아 pandas에서 검색·재검색(호출량 최소화, 재분류 무제한).
- **시설 후보 단위 = (표2 분류, 검색어, 시설키, 수요기관)**: 같은 이름의 시설이 여러 지자체에 있어도 섞이지 않는다. 재검색도 `검수_수요기관`이 있으면 그 기관의 공고만 찾는다. 시설명이 검색어와 일반어뿐이면('영구임대아파트', '공공주택지구', '종합운동장') 수요기관의 지자체명을 앞에 붙여 '구례군 국민임대아파트'처럼 만들고, 재검색은 접두어를 뗀 패턴으로 그 수요기관 공고를 찾는다.
- **사업유형**: 신축·증축·리모델링 모두 DB 대상(유지보수만 제외). 시설 후보의 사업유형은 공고 중 하나라도 신축·증축·리모델링이면 그 유형(우선순위 신축 > 증축·리모델링 > 증축 > 리모델링)으로 두어, 유지보수 공고가 많은 시설의 본공사가 묻히지 않게 한다. 같은 시설의 신축과 리모델링은 프로젝트ID(시설ID-N/E/R/ER)로 분리 집계.
- **재발주 처리**: 프로젝트키(프로젝트ID×공종×단계토큰)별로 취소공고 제외 후 최신 공고 1건을 대표로 채택. 구공고는 `대체됨` 표시로 이력 보존, 금액 ±30% 변동은 경고. 공고명이 같은 재공고는 `대체됨(이후 재공고/재발주)`, **공고명이 다른 이후 공고에 밀린 경우는 `대체됨(공고명 상이 — 별개 공사 가능)`** 으로 구분하고 06 에 '대표 교체(공고명 상이)' 경고, 05 의 `대체된_다른이름_공고수·추정가격합` 에 프로젝트별로 집계한다(0 이 아니면 여러 사업이 한 시설 후보로 묶였을 수 있음). 같은 프로젝트·공종에서 추정가격이 최대치의 30% 미만인 공고(무대기계·승강기 설치 등 부대공사)는 대표가 되지 않고 경고로 남긴다.
- **공종 분류 우선순위**: 면허제한 업종명(선택) → 주공종명·부공종명(목록 응답의 업종명, 예: '기계설비ㆍ가스공사업') → 공고명 규칙('토목건축공사업'은 건축). 유지보수(방수·도색·교체 등) 공고는 기본 제외.
- **금액 구성·대표값(04·05 시트)**: 04 는 요청 양식 순서 `추정가격 | 부가세 | 도급자관급비 | 추정금액 계ⓐ | 관급자관급비ⓑ | 기타금액ⓒ | 금액합계` 로 채택값 수식을 두고, 금액합계는 **추정가격 기준**(추정가격+부가세+관급+기타)과 **기초금액 기준**(기초금액+관급+기타) 두 열을 함께 제시한다(05 의 총공사비합계·㎡당 공사비도 두 기준). 부가세는 API 부가세 → 문서 부가세 → 기초금액−추정가격 → 10% 가정 순(`부가세_출처`), `과세구분` 으로 면세·부분면세(조세특례제한법 제106조) 표시. 공종은 공고명 꼬리표 `(전기)`·`(통신)`·`(소방)`·`(통신설비)`·`(전기/2차분)` 도 인식하고(괄호 앞이 '소방'·'건축' 이면 그 공종: '소방(전기)공사' → 소방), 미분류 공고는 `기타` 로 두되 `공종_비고` 에 사유를 적는다. 05 `건축공종계` = 건축+기계설비+토목+조경.
- **금액 기준**: 추정가격·기초금액·관급자재는 모두 API 값이 1순위(2026-09 실제 응답에서 `presmptPrce`, `bssamt`, `govsplyAmt`, 도급자/관급자 설치 관급액 `contrctrcnstrtnGovsplyMtrlAmt`/`govcnstrtnGovsplyMtrlAmt`, `VAT`, `bdgtAmt` 확인). 문서 추출값은 API 값이 없을 때의 대체·교차검증용. 금액합계(기초금액 기준) = 기초금액 + 도급자관급 + 관급자관급 + 기타금액ⓒ, 추정가격 기준은 위 항목 참조(Excel 수식). 문서의 기초금액은 추정가격의 100~110.5% 범위일 때만, 10% 에 못 미치는 부가세라도 기초금액이 추정가격×1.1 과 맞으면(±1%) 채택하지 않는다(문서의 0·단위 오독 방어). LLM에는 금액을 힌트로 주지 않아 교차검증이 독립적이다.
- **낙찰 정보(선택, 참고 컬럼)**: 「낙찰정보서비스」 공사 낙찰 목록을 개찰일 기준 월 단위로 받아 공고번호+차수(없으면 공고번호)로 붙인다 → `02/03/04` 시트의 `낙찰금액_API`, `낙찰률_API`, `낙찰자`, `참가업체수`, `04` 의 `낙찰률(수식)`(=낙찰금액/기초금액, API 낙찰률과 대조). 같은 공고에 낙찰 행이 여럿(재개찰)이면 개찰일시가 최신인 행. **총공사비·㎡당 공사비에는 쓰지 않는다**(DB 기준은 추정가격·기초금액).
- **재공고 연결**: API 의 이전공고번호(`befBidBbancNo`)로 재공고↔원공고 관계를 확정하고, 이름 규칙으로 못 묶은 경우도 대체 처리한다.
- **교차검증**: 문서 추출 추정가격·기초금액 vs API 값(0.5%), 기초금액≈추정가격×1.1, 연면적·층수 범위, 총공사금액 구성 정합성 → `06_검증로그`.

## 5. 알려진 제약
- 응답 필드명은 2026-09-08 `probe` 로 확인됨(공고목록·기초금액 `[매핑 OK]`, 공고차수는 `000` 처럼 3자리). 코드는 미존재 필드를 공란 처리한다.
- 면허제한 조회는 기본 꺼져 있음(`api.use_license_limit: false`). 필요하면 `probe` 가 시험한 조회구분 결과를 보고 `license_query` 를 맞춘 뒤 켠다.
- 낙찰정보서비스는 기본 꺼져 있음(`api.use_awards: false`). 활용신청 뒤 `true` 로 켜고 `probe` 로 End Point(`api.award_base_url`)·기간 조회 파라미터(`api.award_query`)·필드 매핑(`fields.award_*`)을 확인한다. 실제 응답 필드명은 아직 미확인(2026-09-09 기준 추정값, 후보 이름을 자동 탐색). 수집 실패(20/30) 시 공고·기초금액 수집은 유지되고 안내만 출력된다.
- 첨부 URL 은 `https://www.g2b.go.kr/pn/pnp/pnpe/UntyAtchFile/downloadFile.do?...` 형식이며, 로그인 페이지가 돌아오면 `07_추출노트`에 'HTML 응답' 실패로 기록된다(수동 다운로드 후 `data/files/<공고번호>/` 에 넣고 `attach` 재실행).
- HWP 파서: HWP 5.0 규격(olefile+zlib 레코드) 기반이며 실제 공고문으로 첫 실행 시 확인 필요. 실패 시 `hwp5txt`(pyhwp) → Windows 한컴 COM 순으로 폴백. 배포용(DRM)·암호 문서는 수동 처리.
- 규모: 2019-01~2026-09 전량은 공고 약 129만 건·기초금액 약 103만 건(2026-09-10 수집). `discover` 는 필요한 컬럼만 읽어 처리하며 합성 데이터 기준 수 분 소요(메모리 약 3~4 GB). `notices_std.parquet` 는 전량 표준화본이므로 수백 MB 가 될 수 있다.
- 스캔 PDF는 `pytesseract`+`pdf2image` 설치 시 OCR. 연면적이 공고문에 없으면 현장설명서·설계설명서를 함께 파싱해도 공란일 수 있음(경고 기록).
- LH 자체 조달(ebid.lh.or.kr) 발주분은 나라장터 API 범위 밖.
