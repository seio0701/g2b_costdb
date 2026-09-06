# PC 설치 및 첫 실행 가이드 (Windows, 코딩 경험 없어도 따라 할 수 있게)

이 문서는 `HANDOFF.md`의 "다음 할 일"을 처음부터 끝까지 풀어 쓴 것이다. 모든 명령은 **PowerShell** 창에 한 줄씩 붙여 넣고 Enter 를 누른다.
PowerShell 여는 법: 시작 버튼 → `powershell` 입력 → "Windows PowerShell" 클릭. (관리자 권한은 필요 없다)

> 표기: `…` 로 시작하는 회색 상자는 그대로 입력하는 명령이고, "기대 결과"는 정상일 때 화면에 보이는 문구다.

---

## 0단계. 준비물 확인 (10분)

### 0-1. 파이썬이 있는지 확인
```powershell
python --version
```
- 기대 결과: `Python 3.10.x` ~ `3.13.x` 처럼 3.10 이상.
- `'python' 용어가 인식되지 않습니다` 또는 Microsoft Store 가 열리면 → https://www.python.org/downloads/windows/ 에서 최신 3.x 설치.
  설치 화면 **첫 페이지 맨 아래 "Add python.exe to PATH" 체크**를 반드시 켠 뒤 Install Now. 설치 후 PowerShell 을 닫았다 다시 열고 위 명령을 다시 실행.
- `python` 대신 `py -3 --version` 이 되는 PC 도 있다. 그 경우 이 문서의 `python` 을 모두 `py -3` 으로 바꿔 입력하면 된다.

### 0-2. Git 이 있는지 확인 (저장소를 내려받기 위해)
```powershell
git --version
```
- 없으면 https://git-scm.com/download/win 에서 설치(기본 옵션으로 계속 Next). 설치 후 PowerShell 을 다시 연다.
- Git 을 설치하고 싶지 않으면 1단계의 "ZIP 으로 받기"를 쓴다.

---

## 1단계. 저장소 받기 → 프로젝트 폴더로 이동 (5분)

### 방법 A. Git 으로 받기 (권장)
```powershell
cd $HOME\Documents
git clone -b claude/step-by-step-task-q2xkba https://github.com/seio0701/research-eval-web.git
cd research-eval-web\g2b_costdb
dir
```
- 기대 결과: `dir` 목록에 `config.yaml`, `keywords.yaml`, `README.md`, `HANDOFF.md`, `g2b_costdb`, `tests` 가 보인다.
- 비공개 저장소라 GitHub 로그인 창이 뜨면 본인 계정으로 로그인한다.

### 방법 B. ZIP 으로 받기
1. 브라우저에서 https://github.com/seio0701/research-eval-web 접속 → 왼쪽 위 브랜치 선택(`main`)을 눌러 `claude/step-by-step-task-q2xkba` 선택
2. 초록색 **Code** 버튼 → **Download ZIP** → 다운로드 폴더의 ZIP 을 "압축 풀기"
3. PowerShell 에서 압축 푼 폴더 안의 `g2b_costdb` 로 이동:
```powershell
cd $HOME\Downloads\research-eval-web-claude-step-by-step-task-q2xkba\g2b_costdb
dir
```

> ★ 이후 모든 명령은 **이 `g2b_costdb` 폴더 안에서** 실행한다. PowerShell 을 새로 열 때마다 `cd <위 경로>` 를 먼저 한다.
> 현재 폴더는 명령줄 맨 앞(`PS C:\Users\...\g2b_costdb>`)에서 확인할 수 있다.

---

## 2단계. 파이썬 패키지 설치 (5~10분, 인터넷 필요)
```powershell
python -m pip install --upgrade pip
python -m pip install -r requirements.txt
python -m pip install pyhwp
```
- 기대 결과: 마지막 줄 근처에 `Successfully installed …` 가 보인다. 노란색 WARNING 은 무시해도 된다.
- `pyhwp` 설치가 실패해도 진행할 수 있다(HWP 파일 해석의 예비 수단일 뿐).
- 선택: Excel 수식 점검 도구가 필요하면 `python -m pip install formulas` (없어도 Excel 에서 파일을 열어 `#` 오류가 없는지 보면 된다).
- 회사 네트워크에서 `SSL` 또는 `proxy` 오류가 나면 IT 담당자에게 pip 프록시 설정을 문의한다.

---

## 3단계. 공공데이터포털에서 API 인증키 받기 (15분 + 승인 대기)

1. https://www.data.go.kr 접속 → 회원가입/로그인
2. 검색창에 **`조달청_나라장터 입찰공고정보서비스`** 입력 → 검색 결과에서 해당 서비스 클릭
3. 오른쪽의 **활용신청** 버튼 → 활용목적: "연구(학술)" 등 선택, 상세기능은 **전부 체크**, 라이선스 표시 동의 → 신청
   - 이 서비스는 **자동승인**이다. 다만 승인 직후 1~2시간은 키가 "등록되지 않은 키(30)"로 나올 수 있으니 잠시 기다린다.
4. 상단 **마이페이지 → 오픈API → 활용신청 현황** → 방금 신청한 서비스 클릭
5. 인증키가 두 줄 보인다: **일반 인증키(Encoding)** 와 **일반 인증키(Decoding)**. → **Decoding** 옆의 복사 버튼을 누른다.
   - Encoding 키를 넣어도 프로그램이 자동 변환하지만, Decoding 키를 쓰는 것이 원칙이다.
   - 키는 비밀번호와 같다. 채팅·문서·메일에 붙여 넣지 않는다.
6. 개발계정은 하루 **1,000회** 호출 제한이 있다. 프로그램은 950회에서 스스로 멈추고 "내일 재개"를 안내한다.
   (나중에 마이페이지에서 "운영계정 전환/트래픽 증설"을 신청하면 하루 한도가 늘어난다)

---

## 4단계. 인증키를 환경변수로 저장 (3분)

PowerShell 에 아래를 입력한다. 따옴표 안에 3단계에서 복사한 키를 붙여 넣는다(앞뒤 공백 없이).
```powershell
setx G2B_SERVICE_KEY "여기에_복사한_Decoding_키"
```
- 기대 결과: `성공: 지정한 값을 저장했습니다.`
- `setx` 는 **새로 여는 창부터** 적용된다. 지금 창에서 바로 쓰려면 아래도 한 번 실행한다:
```powershell
$env:G2B_SERVICE_KEY="여기에_복사한_Decoding_키"
```
- 저장됐는지 확인(키 값은 화면에 나오지 않는다):
```powershell
if ($env:G2B_SERVICE_KEY) { "설정됨" } else { "없음" }
```
- Claude 추출 단계(세션 6)에서 쓰는 `ANTHROPIC_API_KEY` 는 지금은 필요 없다. 그때 같은 방법으로 `setx ANTHROPIC_API_KEY "키"` 로 넣는다.

---

## 5단계. 환경 점검 `doctor` (1분)
```powershell
python -m g2b_costdb.pipeline doctor
```
기대 결과(요약):
```
Python 3.x.x ...
  파이썬 3.10 이상: OK
필수 패키지:
  requests     OK ...  pandas OK ...  pyarrow OK ...  openpyxl OK ...  PyYAML OK ...  olefile OK ...  pdfplumber OK ...  python-docx OK ...  anthropic OK
선택 패키지(폴백용): ... (없음이어도 됨)
환경변수(값은 출력하지 않음):
  G2B_SERVICE_KEY: 설정됨
  ANTHROPIC_API_KEY: 없음          ← 지금은 정상
경로: ... (쓰기 가능)
수집 기간: 2019-01 ~ 2026-09 (93개월)
네트워크: https://apis.data.go.kr/... 연결 OK (HTTP 200, 포털 응답 코드 30 — 키 없이 호출했으므로 30/20 이면 정상)
```
문제가 있을 때:
| 화면 | 원인 | 조치 |
|---|---|---|
| `xxx 없음 → pip install xxx` | 패키지 미설치 | 2단계 명령 다시 실행 |
| `G2B_SERVICE_KEY: 없음` | setx 후 창을 새로 열지 않음 | PowerShell 닫고 다시 열기 → `cd` → 다시 실행 (또는 4단계의 `$env:` 명령) |
| `네트워크: ... 연결 실패` | 인터넷/방화벽/프록시 | 다른 네트워크(휴대폰 핫스팟)로 시험, IT 담당자 문의 |

---

## 6단계. 자체 테스트 (2분, 인터넷·키 불필요)
```powershell
python -m tests.test_dedup_and_excel
python -m tests.test_e2e_mock
```
- 기대 결과 ①: `OK: 전체 11건 / 대표 8건 / 시설 3개 / 경고 0건 → ...공사비DB_샘플.xlsx` 와 `OK: 회귀 테스트 통과`
- 기대 결과 ②: 중간에 `INFO`/`WARNING` 줄이 많이 지나간 뒤 맨 끝에 `OK: 통합 실행 통과 (임시 폴더 ..., API 호출 19회)`
  - `다운로드 실패(HTML 응답 …)` WARNING 한 줄은 **의도된 시험**이므로 정상이다.
  - `(formulas 패키지 없음 → 05 값 검증 생략)` 도 정상이다.
- 빨간 `Traceback` 으로 끝나면 그 화면 전체를 복사해 두었다가 Claude Code 에 붙여 넣고 원인을 묻는다.

---

## 7단계. 실제 API 표본 확인 `probe` (2분, 호출 3회)
```powershell
python -m g2b_costdb.pipeline probe --ym 2026-08
```
2026년 8월 공고 5건만 받아 **실제 필드명**을 보여 주고 설정 파일과 대조한다. 화면은 세 덩어리로 나온다.

```
=== getBidPblancListInfoCnstwk (2026-08) totalCount=1234 표본 5건 ===
필드: ['bidNtceDt', 'bidNtceNm', 'bidNtceNo', ...]
표본: { ... 공고 1건의 원문 ... }
[매핑 OK] config.yaml fields 의 관련 항목이 모두 응답에 존재합니다.        ← 이 줄이 목표
=== getBidPblancListInfoCnstwkBsisAmount (2026-08) ... ===
...
[매핑 OK] ...
=== getBidPblancListInfoLicenseLimit (공고번호 ..., params={...}) ... ===
필드: [...]
[업종명 필드] config 매핑 indstrytyNm 존재        ← 또는 "후보 lcnsLmtNm 발견 → config.yaml fields.license_name 을 이 값으로 수정"
오늘 API 호출 3회 (예산 950)
```

### 7-1. `[매핑 확인 필요] {...}` 가 나오면
예: `[매핑 확인 필요] config.yaml fields 중 응답에 없는 항목: {'bsis_amount': 'bssamt'}`
1. 같은 덩어리의 `필드:` 목록에서 비슷한 이름을 찾는다(예: `bssAmt`).
2. `config.yaml` 을 메모장으로 연다: `notepad config.yaml`
3. `fields:` 아래에서 해당 줄을 찾아 **오른쪽 따옴표 안의 값만** 실제 이름으로 바꾼다.
   `bsis_amount: "bssamt"` → `bsis_amount: "bssAmt"`
4. 저장(인코딩은 UTF-8 유지) 후 probe 를 다시 실행해 `[매핑 OK]` 를 확인한다.

### 7-2. `[업종명 필드] 후보 lcnsLmtNm 발견 → ...` 이 나오면
`config.yaml` 의 `license_name: "indstrytyNm"` 을 `license_name: "lcnsLmtNm"` 으로 바꾼다. (안 바꿔도 프로그램이 후보를 자동으로 쓰지만, 바꿔 두면 안내 문구가 사라진다)

### 7-3. 면허제한 덩어리가 `실패: API 오류 10` 또는 `11` 이면
면허제한 조회의 "조회구분(inqryDiv)" 값이 활용가이드와 다른 경우다.
1. data.go.kr 서비스 페이지 하단 **참고문서**에서 활용가이드(docx/hwp)를 내려받아 `getBidPblancListInfoLicenseLimit` 항목의 `inqryDiv` 설명(예: "1:공고게시일시, 2:개찰일시, 3:공고번호")을 확인한다.
2. `config.yaml` 의 `license_query:` 아래 `inqryDiv: "2"` 를 그 값(예: `"3"`)으로 바꾸고 probe 를 다시 실행한다.
3. 잘 모르겠으면 probe 출력을 통째로 Claude Code 에 붙여 넣고 "config.yaml 을 고쳐 달라"고 하면 된다.

### 7-4. `API 오류 30` / `API 오류 20`
- 30: 키가 틀렸거나 아직 반영 전. 4단계 확인(공백·따옴표 없이 Decoding 키), 활용신청 후 1~2시간 대기.
- 20: 활용신청이 승인되지 않은 서비스. 3단계에서 신청한 서비스가 "조달청_나라장터 입찰공고정보서비스"가 맞는지 확인.

---

## 8단계. 3개월 시험 수집 (PROMPTS.md 세션 2, 약 30분)

### 8-1. 기간을 3개월로 줄이기
`notepad config.yaml` → `period:` 부분을 아래처럼 바꾸고 저장:
```yaml
period:
  start: "2026-06"
  end: "2026-08"
```

### 8-2. 수집
```powershell
python -m g2b_costdb.pipeline collect
```
- 기대 결과: 월별로 `getBidPblancListInfoCnstwk 202606: 1,234건` 같은 줄이 지나가고, 끝에 `[완료] ... 3개월 수집 완료` 와 `누적 공사 공고 N건 → notices_all.parquet`, 기초금액도 같은 형식.
- 3개월이면 호출은 수십 회 수준이다. `[미완료]` 와 "예산 도달"이 나오면 **다음 날 같은 명령**을 다시 실행한다(끝난 달은 건너뛴다).

### 8-3. 후보 시설 만들기
```powershell
python -m g2b_costdb.pipeline discover
```
- 기대 결과: `공고 N건 중 후보 시설 M개 → ...\output\facility_candidates.xlsx 에서 검수(노란 셀) 후 research 실행` 과 대분류별·사업유형별 건수.
- `output\facility_candidates.xlsx` 를 Excel 로 열어 노란 셀을 검수한다(범례 시트 참고). 신축·증축·리모델링은 `포함`, 방수·도색·교체 같은 유지보수만 `제외`. 저장 후 **Excel 을 닫는다**(열려 있으면 다음 단계가 파일을 읽지 못한다).
- 3개월치는 결과가 상식에 맞는지 보는 것이 목적이다. 검색어가 엉뚱한 공고를 많이 잡으면 `keywords.yaml` 의 `exclude` 에 단어를 추가하고 `discover` 를 다시 실행한다(검수 내용은 유지된다).

### 8-4. 다음 세션들
- 세션 3(본 수집): `period` 를 `2019-01` ~ `2026-09` 로 되돌리고 매일 `collect` 를 실행해 `[완료]` 가 나올 때까지 반복(약 800~1,100회 호출 → 하루 950회 예산으로 1~2일).
- 세션 4~7: `discover` → 검수 → `research` → `dedup` → `attach` → `extract`(견적) → `extract --yes` → `excel`. 각 세션의 붙여넣기용 지시문은 `PROMPTS.md` 에 있다.

---

## Claude Code 로 대신 시키는 방법
직접 명령을 치는 대신 Claude Code 에 시킬 수도 있다. `g2b_costdb` 폴더에서 `claude` 를 실행한 뒤 `PROMPTS.md` 의 "세션 1" 글을 그대로 붙여 넣으면 2·5·6·7단계를 Claude Code 가 수행하고 결과를 요약한다. 단, **3단계(키 발급)와 4단계(setx)는 사용자가 직접** 해야 하며, 키를 채팅에 붙여 넣지 않는다.

---

## 자주 나는 오류 모음
| 화면에 보이는 문구 | 뜻 | 해결 |
|---|---|---|
| `'python' 용어가 인식되지 않습니다` | 파이썬이 없거나 PATH 미등록 | 0-1 단계, 또는 `py -3` 사용 |
| `No module named g2b_costdb` / `No module named tests` | 다른 폴더에서 실행함 | `cd ...\g2b_costdb` 후 재실행 (`dir` 에 config.yaml 이 보여야 함) |
| `No module named pandas` (등) | 패키지가 다른 파이썬에 설치됨 | `python -m pip install -r requirements.txt` 를 **같은 `python`** 으로 실행 |
| `환경변수 G2B_SERVICE_KEY 에 ... 설정하세요` | 키 미설정/새 창 아님 | 4단계 |
| `API 오류 30: ... 등록되지 않은 서비스키` | 키 오타/승인 대기/Encoding 키 | 4단계·3단계 확인, 1~2시간 후 재시도 |
| `[미완료] ... (내일 같은 명령으로 재개)` / `예산 950 도달` | 하루 호출 한도 | 정상. 다음 날 같은 명령 |
| `[중단] '...xlsx' 파일을 열 수 없습니다` | Excel 에서 파일이 열려 있음 | Excel 닫고 재실행 |
| 한글이 `?` 로 보임 | PowerShell 글꼴/코드페이지 | 실행에는 영향 없음. 보기 싫으면 `chcp 65001` 후 재실행 |
| `Traceback (most recent call last)` 로 끝남 | 프로그램 오류 | 화면 전체를 복사해 Claude Code 에 붙여 넣고 원인·수정 요청 |
