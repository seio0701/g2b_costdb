# CLAUDE.md — g2b_costdb (나라장터 발주공고 → LIMAC 공사비 DB)

이 저장소는 나라장터 Open API로 공사 발주공고를 수집하고, 첨부 공고문(HWP/PDF)을 파싱·LLM 추출하여
Excel 공사비 DB(`output/공사비DB.xlsx`)를 만드는 파이프라인이다. 사용자는 코딩 비전공자(지방투자사업 타당성조사 연구자)이다.
설계 근거는 `docs/나라장터_공사비DB_자동화_울트라플랜.md`, 실행 순서는 `README.md`, 단계별 프롬프트는 `PROMPTS.md`를 본다.

## 사용자와 소통하는 방식
- 사용자는 명령어·코드를 직접 다루지 않는다. **모든 명령 실행·파일 수정은 Claude Code가 수행**하고, 결과는 한국어로 짧게 요약한다(어느 파일이 생겼고, 몇 건이 처리됐고, 다음에 무엇을 할지).
- 사용자에게 판단을 요청할 때는 선택지를 1~3개로 좁혀서 묻는다. 기술 용어는 한 줄 설명을 붙인다.
- 오류가 나면 원인 → 조치 → 재실행 결과를 보고한다. 같은 오류로 3회 이상 실패하면 멈추고 사용자에게 상황을 설명한다.

## 절대 규칙
1. **API 키를 채팅·파일·로그에 출력하거나 저장하지 않는다.** 키는 환경변수 `G2B_SERVICE_KEY`, `ANTHROPIC_API_KEY`로만 읽는다. 키가 없으면 사용자에게 터미널에서 직접 `setx`로 설정하라고 안내한다(키를 채팅에 붙이지 말라고 함께 안내).
2. **단계 순서를 지킨다**: `doctor → probe → collect → discover → (사용자 검수) → research → dedup → attach → extract(견적) → extract --yes → excel`. 검수(`output/facility_candidates.xlsx`의 노란 셀)는 사용자가 직접 하며, Claude Code가 대신 채우지 않는다.
3. `config.yaml`의 `fields` 매핑은 `probe` 출력(실제 응답 필드명)을 확인한 뒤에만 수정한다. 추측으로 바꾸지 않는다.
4. API 일일 호출 예산(`daily_call_budget`)에 도달해 `DailyBudgetExceeded`/"예산 도달"이 나오면 정상 상황이다. 우회하지 말고 "내일 같은 단계를 다시 실행" 안내로 끝낸다.
5. 나라장터 첨부파일 다운로드 외에 외부 사이트를 크롤링하거나, 파이프라인에 없는 대량 호출 코드를 새로 만들지 않는다.
6. Excel 산출물은 전달 전에 수식 오류가 없는지 확인한다(LibreOffice가 있으면 재계산, 없으면 openpyxl로 수식 셀 존재·참조 시트명 점검). 하드코딩 값으로 수식을 대체하지 않는다.
7. 금액 기준은 **추정가격·기초금액(예정가격 산정 기준)** 이며 낙찰가가 아니다. 관급자재는 별도 컬럼으로 유지한다. 이 원칙을 바꾸는 요청은 사용자에게 재확인한다.
   DB 대상 사업유형은 신축·증축·리모델링(유지보수만 제외)이며, 같은 시설의 신축과 리모델링은 별도 프로젝트ID로 분리한다.
8. `data/raw/*.jsonl`(원문 응답)은 삭제·덮어쓰기하지 않는다. 재수집이 필요하면 `data/raw/*_done.txt`에서 해당 월만 제거한다(그 달은 캐시를 우회해 새로 받고 파일이 교체된다). 오늘이 속한 달은 완료 표시가 되지 않으며 매번 다시 받는다.

## 매 세션 시작·종료
- 시작: `HANDOFF.md`가 있으면 먼저 읽고 현재 단계와 미해결 이슈를 파악한다. 없으면 `README.md`를 읽는다.
- 종료: `HANDOFF.md`에 (완료 단계 / 생성 파일과 건수 / 발생한 오류와 조치 / 다음 할 일 / 사용자 결정 필요 사항)을 개조식으로 갱신한다.

## 실행 환경
- Windows PowerShell 기준. Python 3.10+. 의존성은 `pip install -r requirements.txt` (+ `pip install pyhwp`). 환경 점검은 `python -m g2b_costdb.pipeline doctor`.
- 모든 단계는 **`g2b_costdb` 폴더**(config.yaml 이 있는 곳)에서 `python -m g2b_costdb.pipeline <stage>` 로 실행한다. `config.yaml paths` 의 상대경로는 이 폴더 기준으로 해석된다.
- 중간 산출물은 `data/`(parquet·jsonl·json), 결과물은 `output/`, 첨부 원본은 `data/files/`, 추출 텍스트는 `data/text/`.

## 단계별 성공 기준
| 단계 | 성공 기준 | 실패 시 |
|---|---|---|
| doctor | 파이썬 3.10+, 필수 패키지 OK, `G2B_SERVICE_KEY` 설정됨, 네트워크 연결 OK | 없는 패키지 설치, 키 없으면 setx 안내 후 중단 |
| probe | 공고·기초금액·면허제한 표본 출력, `[매핑 OK]` 또는 `[매핑 확인 필요]` 목록 | 존재하지 않는 매핑만 실제 필드명으로 수정, 변경 내역을 사용자에게 표로 보고. 면허제한 오류 10/11 이면 `api.license_query.inqryDiv` 조정 |
| collect | `data/notices_all.parquet`, `data/bsis_all.parquet` 생성, 월별 건수 로그, 끝에 `[완료]` 또는 `[미완료] … 남은 월` | 오류코드 30/20 → 키·승인 문제 안내(코드가 한국어로 설명) / 22 또는 예산 도달 → 내일 재실행 |
| discover | `output/facility_candidates.xlsx` 생성, 시설 수·대분류별·사업유형별 건수 보고, 상위 20개 시설명 미리보기. 재실행 시 기존 검수 내용은 유지됨 | 후보 0건이면 keywords.yaml include 확장 제안(사용자 확인 후 적용). `[경고] 수집 미완료 월`이 보이면 사용자에게 알림 |
| research | `data/notices_research.parquet`, 시설당 공종 분포(건축/전기/정보통신/소방) 보고 | 검수 Excel 미저장·열림 상태 확인 |
| dedup | "전체 N건 → 대표 M건" 보고, 경고 건 요약 | — |
| attach | `07` 추출노트 기준 파일 성공률 보고. HWP 실패율 20% 초과 시 pyhwp/한컴 COM 폴백 점검 | 배포용 HWP는 목록만 정리해 사용자에게 수동 처리 안내 |
| extract | `extract`(플래그 없음)가 출력한 건수·토큰·예상 비용을 사용자에게 보고하고 확인을 받은 뒤 실행. 방식은 사용자가 고른다: `extract --yes`(즉시) / `extract --batch --yes` 제출 후 `extract --batch` 수거(50% 할인) / `extract --export` → Cowork·사람이 `data/llm_out` 에 JSON 작성 → `extract --import`. `data/llm_docs.json` 건수, 검증로그의 정상/경고/오류 건수 보고 | 인증 실패면 키 안내 후 중단, 429면 잠시 후 재실행(완료 건은 건너뜀). `--import` 의 "문제 파일" 목록은 사용자에게 보여 주고 대신 채우지 않는다 |
| excel | `output/공사비DB.xlsx` 생성, 시트별 행수, 수식 오류 0건, 06_검증로그 경고·오류 건수 요약 | 오류 셀 위치와 원인 보고 |

## 코드 수정 원칙
- 버그 수정은 최소 범위로, 수정한 파일·함수·이유를 보고한다. 새 의존성 추가 시 `requirements.txt`에 반영한다.
- 실제 API 응답으로 필드명·구조가 다르게 확인되면 `config.yaml`(매핑) → `discover.standardize`(표준화) 순으로 고친다. 원문 JSONL은 항상 보존되므로 재수집 없이 재표준화가 가능하다.
- `python -m tests.test_dedup_and_excel`(회귀 테스트 `tests/test_regressions.py` 포함)과 `python -m tests.test_e2e_mock`(가짜 서버 통합 실행)이 항상 통과해야 한다. 파이프라인 코드를 고쳤으면 둘 다 실행해서 확인한다. 실제 API 응답으로 필드명이 다르게 확인되면 `tests/mock_g2b.py` 의 가짜 데이터도 같은 필드명으로 맞춘다.
