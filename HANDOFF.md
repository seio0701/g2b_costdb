# HANDOFF.md — g2b_costdb 진행 상황

> 매 세션 시작 시 이 파일을 먼저 읽는다. 갱신일: 2026-09-05

## 1. 현재 단계
- **세션 0(코드 정비) 완료 / 세션 1(환경 점검 + probe)의 사용자 PC 실행 대기**
- 이번 작업은 네트워크가 막힌 샌드박스(Claude Code 웹)에서 진행되어 **나라장터 API 호출(probe/collect)은 실행하지 못했다.**
  API 키(`G2B_SERVICE_KEY`)도 설정되어 있지 않았다. 따라서 `probe` 이후 단계는 모두 사용자 PC(Windows)에서 진행해야 한다.

## 2. 이번 세션에서 한 일
- 파이프라인(`g2b_costdb/`)을 저장소에 편입하고 설계 문서를 `docs/나라장터_공사비DB_자동화_울트라플랜.md` 로 이동.
- 합성 데이터 테스트(`python -m tests.test_dedup_and_excel`) 통과 확인 → `sample_output/공사비DB_샘플.xlsx` 생성(수식 오류 0건, `formulas` 패키지로 평가).
- 실데이터 첫 실행에서 터질 결함을 모듈별 적대적 검토(7개 관점 × 재현 검증)로 찾아 수정. 주요 항목:
  | 영역 | 수정 |
  |---|---|
  | API | 포털 오류 XML(키 미등록 30, 트래픽 초과 22 등)을 해석해 재시도 없이 한국어 안내로 중단, 22는 "내일 재개"로 처리, NODATA(03)는 빈 결과. Encoding 키(`%2B…`) 자동 변환. totalCount 누락 시에도 끝까지 페이지 조회 |
  | 수집 | 오늘이 속한 달은 완료 표시 없이 매번 재수집(캐시 우회), 미래 달 건너뜀, done.txt 에서 지운 달은 새로 받음, 월별 실패는 건너뛰고 마지막에 `[미완료]` 보고. 월별 타입이 달라도 parquet 저장 실패 없음 |
  | 탐색·검수 | 시설 후보 키에 수요기관 포함(동명 시설 분리), `discover` 재실행 시 검수 내용·시설ID 유지, `검수_수요기관` 열 추가(재검색 범위 제한), 빈 검수_시설명 처리, 한 공고가 두 분류에 걸리면 구체적 검색어 1건만 채택, '제주시' 접두어 보존, 띄어쓴 검색어('온실 건립') 처리 |
  | 분류 | '토목건축공사업'은 건축, '소방서' 건물은 소방 아님, '(취소 후 재공고)'를 취소로 오인하지 않음, 매입임대·복합화 규칙 정리 |
  | 중복정리 | 같은 공고가 두 시설에 걸려도 탈락 없음, 괄호 안 (1단계)/(2공구) 분할발주 토큰 인식, 'N차 재공고'는 회차로 처리, 날짜 없는 공고를 최신으로 오인하지 않음, **추정가격이 최대치의 30% 미만인 공고(무대기계·승강기 설치 등)는 본공사 대표를 대체하지 못함**(경고 기록) |
  | 첨부 | 다운로드 중단 파일 캐시 방지(.part), HTML(로그인) 응답 감지, 기존 파일은 요청 없이 재사용, 확장자 없으면 매직바이트로 보완, HWPX 탭 뒤 텍스트 보존, cp949 파일명·텍스트, 공고 단위 오류 격리와 25건마다 저장(중단 후 이어서) |
  | LLM 추출 | JSON 스키마 강제(구조화 출력, 미지원 시 폴백), 금액 문자열·NaN 안전 처리, 실패 건은 다음 실행에서 재시도, 인증 실패 즉시 중단, **`extract` 는 비용 견적만 출력하고 `extract --yes` 로만 실제 호출**, API 금액을 힌트로 주지 않아 교차검증 독립성 확보 |
  | Excel | API 금액이 없을 때 문서 추출 금액으로 총공사비 산출, 공종누락경고를 '누락'(공고 없음)/'금액없음'으로 구분, 05 연면적은 01 시트(노란 셀) 참조, 검증로그 숫자 컬럼 통일 |
  | CLI | `doctor` 단계(파이썬·패키지·키 존재·네트워크 점검), `config.yaml paths` 를 프로젝트 폴더 기준으로 해석(어느 폴더에서 실행해도 동일), Excel 파일이 열려 있을 때 안내, cp949 콘솔 보호 |
- 회귀 테스트 `tests/test_regressions.py` 추가(위 수정사항 대부분을 검증), `tests/check_excel.py`(수식 오류 점검) 추가.
- 문서 갱신: README.md(실행 순서·doctor·extract --yes), CLAUDE.md(단계·성공기준), PROMPTS.md(세션 1·6 지시문), config.yaml(새 설정 `api.license_query`, `dedup.minor_notice_ratio`, `llm.structured_output`, `llm.usd_krw`, `max_tokens 8000`).

## 3. 생성·변경된 파일
- 코드: `g2b_costdb/api_client.py`, `collect.py`, `classify.py`, `discover.py`, `dedup.py`, `attachments.py`, `extract_llm.py`, `build_excel.py`, `pipeline.py`
- 설정: `config.yaml`, `keywords.yaml` · 테스트: `tests/test_regressions.py`, `tests/check_excel.py` · 문서: `README.md`, `CLAUDE.md`, `PROMPTS.md`, `docs/…울트라플랜.md`
- 데이터 산출물(`data/`, `output/`)은 아직 없음(API 미실행). `sample_output/공사비DB_샘플.xlsx` 는 가상 데이터 예시.

## 4. 다음 할 일 (사용자 PC, PROMPTS.md 세션 1)
1. PowerShell 에서 `cd <저장소>\g2b_costdb` → `pip install -r requirements.txt` → `pip install pyhwp`
2. 공공데이터포털에서 「조달청_나라장터 입찰공고정보서비스」 활용신청 후 **일반 인증키(Decoding)** 를 `setx G2B_SERVICE_KEY "키"` 로 설정(새 터미널 열기)
3. `python -m g2b_costdb.pipeline doctor` → 모두 OK 인지 확인
4. `python -m tests.test_dedup_and_excel` → `OK` 두 줄 확인
5. `python -m g2b_costdb.pipeline probe --ym 2026-08` → `[매핑 확인 필요]` 항목이 있으면 `config.yaml fields` 수정. 면허제한 표본에서 업종명 필드 안내(`lcnsLmtNm` 등)가 나오면 `fields.license_name` 수정, 오류 10/11 이면 `api.license_query.inqryDiv` 조정
6. 이후 PROMPTS.md 세션 2(3개월 시험 수집: `period` 를 2026-06~2026-08 로) → 세션 3(본 수집) 순서

## 5. 사용자 결정 필요 사항
- LLM 추출 모델: 기본 `claude-sonnet-5`(저비용). 정확도를 우선하면 `config.yaml llm.model` 을 `claude-opus-5` 로(비용 약 2.5배). `extract`(플래그 없음)가 견적을 보여준다.
- `dedup.minor_notice_ratio`(기본 0.30): 같은 프로젝트·공종에서 추정가격이 최대치의 30% 미만인 공고를 부대공사로 보고 대표에서 제외한다. 소규모 시설(온실·재난물자창고)에서 본공사가 잘못 제외되면 06_검증로그의 '소액 공고 분리' 항목을 보고 값을 낮춘다.

## 6. 미해결·주의
- 실제 API 응답 필드명·면허제한 조회 파라미터는 `probe` 로 확정 전까지 추정값이다(코드는 미존재 필드를 공란 처리하고 업종명 필드 후보를 자동 탐색).
- HWP 파서는 규격 기반 합성 데이터로만 검증했다. 첫 `attach` 실행 후 `07_추출노트`의 HWP 실패율을 확인할 것(20% 초과 시 pyhwp/한컴 COM 폴백 점검).
- 나라장터 첨부 URL이 로그인/세션을 요구하면 다운로드가 'HTML 응답' 실패로 기록된다. 이 경우 수동 다운로드 후 `data/files/<공고번호>/` 에 넣고 `attach` 재실행.
