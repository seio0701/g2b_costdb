# g2b_costdb — 나라장터 발주공고 → LIMAC 공사비 DB 파이프라인

표2 시설용도(미술관·문화예술회관·수장고·공공주택·경기장·주차장·화장시설 등)의 공사 발주공고를
나라장터 Open API로 전량 수집하고, 첨부 공고문(HWP/PDF)을 파싱·LLM 추출하여 Excel DB로 정리한다.
설계 배경과 단계별 상세는 `docs/나라장터_공사비DB_자동화_울트라플랜.md` 참조. 세션별 붙여넣기용 지시문은 `PROMPTS.md`.

## 1. 준비 (한 번만)
```powershell
cd <저장소>\g2b_costdb                     # ★ 모든 명령은 이 폴더에서 실행
pip install -r requirements.txt
pip install pyhwp                          # 선택: HWP 폴백(hwp5txt)
pip install formulas                       # 선택: Excel 수식 오류 점검(tests/check_excel.py). 없으면 Excel 에서 직접 열어 확인
# 공공데이터포털 → 「조달청_나라장터 입찰공고정보서비스」 활용신청(자동승인) → 마이페이지에서 '일반 인증키(Decoding)' 복사
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
| 필드 확인 | `python -m g2b_costdb.pipeline probe --ym 2026-08` | 응답 필드명·표본 + `config.yaml fields` 매핑 대조 결과(없는 항목 표시), 면허제한 표본 | 3회 |
| S1 전량수집 | `python -m g2b_costdb.pipeline collect` | `data/raw/*.jsonl`, `data/notices_all.parquet`, `data/bsis_all.parquet` | 월×페이지(약 800~1,100회). 일일예산 초과 시 `[미완료]` 표시와 남은 월을 보여주며 정상 종료 → 다음날 같은 명령 |
| S2 후보탐색 | `python -m g2b_costdb.pipeline discover` | `output/facility_candidates.xlsx` (노란 셀 검수) | 0 |
| **검수** | Excel에서 `검수_포함여부 / 검수_시설명 / 검수_별칭 / 검수_수요기관` 수정 후 저장·닫기 | — | — |
| S3 재검색 | `python -m g2b_costdb.pipeline research` | `data/notices_research.parquet` (전 공종 공고 + 면허제한 업종), 시설별 공종 분포 표 | 재검색 공고 수만큼(캐시됨) |
| S4 중복정리 | `python -m g2b_costdb.pipeline dedup` | `data/notices_hist.parquet`, `data/notices_latest.parquet`, `data/logs_dedup.parquet` | 0 |
| S5 첨부수집 | `python -m g2b_costdb.pipeline attach` | `data/files/`, `data/text/`, `data/texts.json`, `data/notes_attach.parquet` | 0 (파일 다운로드만, 중단 후 재실행 시 이어서) |
| S6 LLM추출 | `python -m g2b_costdb.pipeline extract` → 견적 확인 → `extract --yes` | `data/llm_docs.json`, `data/logs_verify.parquet` | Claude API(`--yes` 없이는 비용 견적만 출력) |
| S7 Excel | `python -m g2b_costdb.pipeline excel` | `output/공사비DB.xlsx` | 0 |

- `discover` 를 다시 실행해도 기존 검수 내용과 시설ID는 (시설키, 수요기관)이 같은 행에 그대로 이어진다.
- 오늘이 속한 달은 완료 표시를 하지 않고 매번 다시 받는다. 특정 달을 다시 받으려면 `data/raw/notices_done.txt`(또는 `bsis_done.txt`)에서 해당 월을 지우고 `collect` 를 실행한다(파일이 있으므로 캐시를 우회해 새로 받음).

## 3. 검증(합성 데이터, 네트워크·키 불필요)
```powershell
python -m tests.test_dedup_and_excel     # 분류·중복정리·Excel 수식 점검 → sample_output/공사비DB_샘플.xlsx, 이어서 회귀 테스트 실행
python -m tests.test_regressions         # 회귀 테스트만
python -m tests.test_e2e_mock            # 가짜 나라장터 서버로 probe→collect→discover→research→dedup→attach→extract→excel 전 단계 통합 실행(약 20초)
python -m tests.check_excel output/공사비DB.xlsx   # 생성된 Excel 의 수식 오류 점검(LibreOffice 또는 formulas 패키지)
```
`sample_output/공사비DB_샘플.xlsx`는 **가상 공고**로 만든 스키마 예시다(실데이터 아님). 통합 테스트의 가짜 서버(`tests/mock_g2b.py`)는
포털 오류 응답(키 오류 XML, 데이터 없음 03), 면허제한 필드명 차이, HTML 응답 첨부, cp949 텍스트 등을 재현한다.

## 4. 주요 설계 포인트
- **전량 수집 후 로컬 검색**: 키워드마다 API를 반복 호출하지 않고 월 단위로 공사 공고를 전량 내려받아 pandas에서 검색·재검색(호출량 최소화, 재분류 무제한).
- **시설 후보 단위 = (표2 분류, 검색어, 시설키, 수요기관)**: 같은 이름의 시설이 여러 지자체에 있어도 섞이지 않는다. 재검색도 `검수_수요기관`이 있으면 그 기관의 공고만 찾는다.
- **사업유형**: 신축·증축·리모델링 모두 DB 대상(유지보수만 제외). 같은 시설의 신축과 리모델링은 프로젝트ID(시설ID-N/E/R/ER)로 분리 집계.
- **재발주 처리**: 프로젝트키(프로젝트ID×공종×단계토큰)별로 취소공고 제외 후 최신 공고 1건을 대표로 채택. 구공고는 `대체됨` 표시로 이력 보존, 금액 ±30% 변동은 경고. 같은 프로젝트·공종에서 추정가격이 최대치의 30% 미만인 공고(무대기계·승강기 설치 등 부대공사)는 대표가 되지 않고 경고로 남긴다.
- **공종 분류 우선순위**: 면허제한 업종명 → 주공종명 → 공고명 규칙('토목건축공사업'은 건축). 유지보수(방수·도색·교체 등) 공고는 기본 제외.
- **금액 기준**: 추정가격(API)·기초금액(API)·관급자재(문서). 총공사비 = 기초금액 + 도급자관급액 + 관급자관급액(Excel 수식). LLM에는 금액을 힌트로 주지 않아 교차검증이 독립적이다.
- **교차검증**: 문서 추출 추정가격·기초금액 vs API 값(0.5%), 기초금액≈추정가격×1.1, 연면적·층수 범위, 총공사금액 구성 정합성 → `06_검증로그`.

## 5. 알려진 제약
- 응답 필드명은 활용가이드(1.2)와 `probe` 결과로 확정할 것. 코드는 미존재 필드를 공란 처리하며, 면허제한 업종명 필드는 후보(`lcnsLmtNm` 등)를 자동 탐색한다.
- 면허제한 조회 파라미터(`config.yaml api.license_query.inqryDiv`)는 활용가이드 버전에 따라 다를 수 있음 — `probe` 가 오류 10/11 을 보이면 조정.
- HWP 파서: HWP 5.0 규격(olefile+zlib 레코드) 기반이며 실제 공고문으로 첫 실행 시 확인 필요. 실패 시 `hwp5txt`(pyhwp) → Windows 한컴 COM 순으로 폴백. 배포용(DRM)·암호 문서는 수동 처리.
- 스캔 PDF는 `pytesseract`+`pdf2image` 설치 시 OCR. 연면적이 공고문에 없으면 현장설명서·설계설명서를 함께 파싱해도 공란일 수 있음(경고 기록).
- LH 자체 조달(ebid.lh.or.kr) 발주분은 나라장터 API 범위 밖.
