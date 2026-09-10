# PROMPTS.md — Claude Code에 붙여 넣을 세션별 지시문

각 세션은 프로젝트 폴더(`C:\DB_WORK\g2b_costdb`, config.yaml 이 있는 곳)에서 `claude`를 실행한 뒤 아래 글을 그대로 붙여 넣는다.
한 세션에 한 단계씩. 세션이 끝나면 Claude Code가 `HANDOFF.md`를 갱신하므로 다음 세션은 이어서 진행된다.
API 키는 채팅에 붙이지 말고 PowerShell에서 `setx`로 직접 설정한다(README 0단계).

---

## 세션 1 — 환경 점검 + probe (약 20분)
```
CLAUDE.md와 README.md를 읽고 시작해줘. 나는 코딩을 잘 못하니 결과는 한국어로 짧게 요약해줘.
1) requirements.txt와 pyhwp 설치 후 python -m g2b_costdb.pipeline doctor 를 실행해 결과를 요약해줘(키 값은 절대 출력하지 말 것).
   G2B_SERVICE_KEY가 없으면 내가 setx로 넣는 방법을 알려주고 멈춰.
2) python -m tests.test_dedup_and_excel 를 실행해 통과하는지 확인.
3) python -m g2b_costdb.pipeline probe --ym 2026-08 를 실행하고, 출력의 [매핑 확인 필요] 항목이 있으면 실제 필드명으로 config.yaml을 고쳐줘.
   면허제한 표본에서 업종명 필드 안내가 나오면 fields.license_name 도 맞춰줘. 무엇을 왜 바꿨는지 표로 보고해줘.
5) 끝나면 HANDOFF.md를 만들어 진행상황을 정리해줘.
```

## 세션 2 — 3개월 시험 수집 (약 30분)
```
HANDOFF.md를 읽고 이어서 해줘.
config.yaml의 period를 2026-06 ~ 2026-08 로 바꾼 뒤 collect → discover 를 실행해줘.
output/facility_candidates.xlsx가 만들어지면 시설 수와 상위 20개 시설명·사업유형을 보여줘.
표2 검색어(keywords.yaml)로 잡힌 결과가 상식적으로 맞는지 네 의견도 짧게 덧붙여줘. HANDOFF.md 갱신.
```

## 세션 2-보강 — 낙찰정보서비스 연결 (활용신청 승인 뒤, 약 10분)
```
HANDOFF.md를 읽고 이어서 해줘. 공공데이터포털에서 「조달청_나라장터 낙찰정보서비스」 활용신청이 승인됐어.
config.yaml의 api.use_awards 를 true 로 바꾸고 python -m g2b_costdb.pipeline probe --ym 2026-08 을 실행해서
낙찰 덩어리의 End Point·조회 파라미터·[award_*] 매핑 결과를 보고, 필요한 부분만 config.yaml 을 고쳐줘(무엇을 왜 바꿨는지 표로).
그 다음 collect 를 실행해 낙찰 목록이 몇 건 받아졌는지, discover 를 다시 실행해 낙찰금액이 붙은 공고가 몇 건인지 알려줘.
낙찰금액·낙찰률은 참고 컬럼이고 총공사비 기준은 그대로 추정가격·기초금액이야. HANDOFF.md 갱신.
```

## 세션 3 — 본 수집 (하루 1회, 완료까지 2~3일)
```
HANDOFF.md를 읽고 이어서 해줘.
config.yaml의 period를 2019-01 ~ 2026-09 로 되돌리고 collect 를 실행해줘.
"예산 도달"로 멈추면 우회하지 말고, 오늘까지 완료된 월과 남은 월 수를 알려주고 끝내줘. HANDOFF.md 갱신.
```
(다음 날에도 같은 글을 붙여 넣는다. 끝난 월은 자동으로 건너뛴다.)

## 세션 4 — 후보 시설 생성 (약 10분) → 이후 내가 Excel 검수
```
HANDOFF.md를 읽고 이어서 해줘. collect가 전 기간 끝났는지 확인한 뒤 discover 를 실행해줘.
output/facility_candidates.xlsx의 시설 수, 표2 대분류별 건수, 사업유형별 건수를 표로 보여줘.
내가 노란 셀(검수_포함여부/검수_시설명/검수_별칭)을 직접 채울 거니까 그 파일은 수정하지 마. HANDOFF.md 갱신.
```
> 여기서 Excel을 열어 노란 셀을 검수한다. 신축·증축·리모델링은 `포함`, 방수·도색·교체 같은 유지보수만 `제외`, `검토필요`는 판단해서 둘 중 하나로. 정확한 시설명, 별칭은 `;`로 구분, `검수_수요기관`은 그대로 두면 그 기관 공고만 찾는다(이름만으로 찾으려면 비움). 저장 후 Excel을 닫는다.

## 세션 5 — 재검색·정리·첨부 수집 (약 1~2시간, 대부분 대기)
```
HANDOFF.md를 읽고 이어서 해줘. 검수한 output/facility_candidates.xlsx를 기준으로
research → dedup → attach 를 순서대로 실행해줘.
- research 후: 시설별로 건축/전기/정보통신/소방 공고가 몇 건 잡혔는지 표로.
- dedup 후: 전체 N건 → 대표 M건, 경고 목록.
- attach 후: 첨부파일 다운로드·텍스트 추출 성공률, 실패 파일 유형별 건수. HWP 실패가 많으면 원인과 폴백(pyhwp/한컴 COM) 적용 결과.
data/text 폴더의 txt 3개를 열어 공사개요·추정가격 문구가 보이는지 확인해서 알려줘. HANDOFF.md 갱신.
```

## 세션 6 — LLM 추출 (비용 발생, 먼저 견적)
```
HANDOFF.md를 읽고 이어서 해줘.
1) 환경변수 ANTHROPIC_API_KEY 존재 여부만 확인(값 출력 금지).
2) python -m g2b_costdb.pipeline extract (플래그 없이) 를 실행해 출력된 대상 건수·예상 토큰·예상 비용을 보고하고 내 확인을 기다려줘.
3) 내가 "배치 진행"이라고 하면 extract --batch --yes 로 제출하고, 1시간 뒤(또는 다음 세션에서) extract --batch 로 결과를 수거해줘.
   "바로 진행"이라고 하면 extract --yes 로 즉시 처리해줘. 끝나면 추출 성공 건수와 06_검증로그 기준 정상/경고/오류 건수, 경고 상위 10건을 보여줘. HANDOFF.md 갱신.
```

## 세션 6-대안 — Cowork(구독)로 직접 추출하기 (API 비용 0, 시범·예외 처리용)
```
HANDOFF.md를 읽고 이어서 해줘. python -m g2b_costdb.pipeline extract --export --limit 40 을 실행해 data/llm_in 에 작업지시 파일을 만들어줘(시범 40건. 전체는 --limit 없이).
그다음 data/llm_in/README_지시문.md 의 규칙대로 llm_in 의 txt 를 20개씩 읽어 llm_out 에 같은 이름의 .json 으로 저장해줘.
전부 끝나면 python -m g2b_costdb.pipeline extract --import 를 실행해 반영하고, 문제 파일 목록과 미추출 건수를 알려줘. HANDOFF.md 갱신.
(반영된 txt·json 은 data/llm_done 으로 옮겨지므로 llm_in·llm_out 에는 남은 일만 보인다. 연면적이 빠진 건을 다시 하려면 extract --export --redo-low)
```

## 세션 7 — Excel DB 생성·점검
```
HANDOFF.md를 읽고 이어서 해줘. excel 을 실행해서 output/공사비DB.xlsx를 만들고,
시트별 행 수, 수식 오류 여부, 05_공사비DB_시설합산의 ㎡당 공사비 상위·하위 5개 시설, 공종누락경고가 있는 시설 목록을 보여줘.
이상치(㎡당 공사비가 다른 시설의 3배 이상이거나 1/3 이하)는 원인 후보(연면적 오독, 공종 누락, 관급 누락)와 함께 정리해줘. HANDOFF.md 갱신.
```

## 문제가 생겼을 때
```
방금 오류가 났어. 원인을 한 줄로 설명하고, 고친 뒤 같은 단계를 다시 실행해줘. 같은 오류가 3번 반복되면 멈추고 상황을 알려줘.
```
```
지금까지 뭘 했고 다음에 뭘 해야 하는지 HANDOFF.md 기준으로 5줄로 요약해줘.
```
