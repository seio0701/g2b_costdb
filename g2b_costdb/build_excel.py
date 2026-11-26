"""S7 Excel DB 생성 (openpyxl, 수식 기반).

시트: README / 01_시설마스터 / 02_공고목록_전체 / 03_공고목록_최신 / 04_공사비DB_공종별 / 05_공사비DB_시설합산 / 06_검증로그 / 07_추출노트
      (+ 05 뒤에 05-요약_실제누락_유형별 / 05-0_실제누락0개 … 05-4_실제누락4개: 실제 누락 공종 수별 프로젝트 목록과 유형별 요약 — 05 를 읽는 수식)
      (+ 08_용역공고(참고): link-servc 결과가 있을 때만 — 05 에 용역 참고 열이 함께 생긴다)
색상: 파란 글자 = 원천값(API/문서), 검은 글자 = 수식, 노란 배경 = 사용자 입력·검수 셀
"""
from __future__ import annotations

import datetime as dt
import json
import math
import re
from typing import Dict, List, Optional

import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

from .dedup import _parse_dt


FONT = Font(name="Arial", size=10)
BLUE = Font(name="Arial", size=10, color="0000FF")
BOLD = Font(name="Arial", size=10, bold=True)
HEAD = PatternFill("solid", fgColor="DDEBF7")
YELLOW = PatternFill("solid", fgColor="FFFF00")
MONEY = "#,##0;(#,##0);-"
TRADES = ["건축", "전기", "정보통신", "소방", "조경", "기계설비", "토목", "기타"]

# 01 시설마스터의 공사개요 열(조달청 '공사비 정보광장' 공사개요 항목 순서)
FAC_OVERVIEW_COLS = ["공사명", "공사규모(수식)", "구조", "지하층수", "지상층수", "연면적_m2", "대지면적_m2", "건축면적_m2", "조경면적_m2",
                     "건폐율_pct", "건폐율_계산(수식)", "용적률_pct", "기준층층고_m", "최고높이_m", "주차대수", "세대수", "현장위치", "발주년월", "발주년월_기준",
                     "공사기간_일", "건축법상용도(수식)", "건축법상용도_문서", "용도", "설계특이사항", "개요_출처"]
FAC_INPUT_COLS = ["구조", "지하층수", "지상층수", "연면적_m2", "대지면적_m2", "건축면적_m2", "조경면적_m2", "건폐율_pct", "용적률_pct", "기준층층고_m",
                  "최고높이_m", "주차대수", "세대수", "공사기간_일", "건축법상용도_문서", "용도", "설계특이사항"]     # 노란 셀(문서 추출값, 사용자 수정 가능)

S_TRADE = "04_공사비DB_공종별"
S_TRADE_DETAIL = "04b_공종세부내역(문서)"
S_FAC = "05_공사비DB_시설합산"
S_SERVC = "08_용역공고(참고)"
SERVC_KINDS = ["설계", "건설사업관리·감리", "폐기물처리"]                       # servc.KINDS 와 같다(05 의 용역 참고 열)
MISS_REAL_COL = "공종누락_통합발주제외(수식)"       # 05: 공종누락경고의 '누락' 가운데 통합발주_포함공종에 적힌(다른 공고에 묶인) 공종을 뺀 것
MISS_TRADES = ("건축", "전기", "정보통신", "소방")     # 공종 누락을 보는 주요 공종
MISS_CNT_COL = "실제누락_공종수(수식)"              # 05: MISS_REAL_COL 에 남은 공종의 수(0~4)
S_MISS_SUM, S_MISS_N = "05-요약_실제누락_유형별", "05-{k}_실제누락{k}개"    # 05 의 실제누락_공종수별 프로젝트 목록(05-0 … 05-4)과 유형별 요약. 값은 모두 05 를 읽는 수식
MISS_ROW_COL = "05시트_행(수식)"                    # 05-k: 그 프로젝트가 05 시트에서 몇째 자료 행인지(MATCH) — 같은 행의 다른 칸이 이 번호로 05 를 읽는다
SERVC_DONE_COL, SERVC_DONE_KIND = "발주완료_폐기물용역(수식)", "폐기물처리"      # 05: 이 프로젝트의 폐기물처리 용역(합산여부 Y)이 08 에 있으면 '발주완료'(보조지표)
SERVC5_COLS = [f"용역_{k}_추정가격(수식)" for k in SERVC_KINDS] + ["용역_공고수(수식)", SERVC_DONE_COL]
SERVC_COLS = ["프로젝트ID", "시설ID", "시설명", "용역종류", "용역세부", "합산여부", "합산제외_까닭", "연결강도", "연결근거", "공고번호", "공고차수", "공고명", "공고일시",
              "공고기관", "수요기관", "추정가격", "배정예산", "상태", "공고종류", "용역구분", "조달분류", "프로젝트_후보", "공사_최초공고일", "공사대비_개월",
              "공사비대비_비율", "상세URL"]
# 08: 공사 공고 수집 종료일(= 워크북에 싣는 공사 공고의 가장 늦은 공고일) − 용역 공고일(날수. 음수 = 그 뒤에 난 용역 공고). 워크북을 만들 때 계산해 넣는다(연결 표에는 없는 열).
# 05: 이 프로젝트의 폐기물처리 용역(Y)이 모두 SERVC_RECENT_DAYS 일(6개월) 안쪽이거나 그 뒤면 '발주완료' 대신 SERVC_DONE_RECENT — 아직 발주 중일 수 있는 사업
SERVC_AGE_COL, SERVC_RECENT_DAYS, SERVC_DONE_RECENT = "공사수집종료까지_일수", 183, "발주완료(최근)"      # 183일 = 6개월
SERVC8_COLS = [x for c in SERVC_COLS for x in ((c, SERVC_AGE_COL) if c == "공사대비_개월" else (c,))]      # 08 시트의 열 = 연결 표(SERVC_COLS) + 일수 열
DETAIL_COLS = ["프로젝트ID", "시설ID", "시설명", "사업유형", "상위공종", "공고번호", "공고차수", "공고키", "세부공종", "항목명_원문",
               "금액_원_문서", "비율(수식)", "배분금액(수식)", "배분적용", "금액기준", "관급자재_포함", "단위", "집계표합계_원_문서", "출처파일", "신뢰도", "근거문구"]
DETAIL_EXCLUDED = ["공통가설", "간접비", "관급자재", "부가세", "합계", "배분제외"]     # 비율 배분의 분자·분모에서 빼는 행('합계' = 여러 공종을 합친 도급비·총공사비 행, '배분제외' = 이설비·분담금·용역비 등 도급 공사비 밖 항목)
DETAIL_ALLOC_TRADES = ["전기", "정보통신", "소방", "조경", "기계설비", "토목", "철거", "기타"]   # 05 문서배분 열
TOTAL5_COLS = ["총공사비합계_기초금액기준(수식)", "총공사비합계_추정가격기준(수식)", "㎡당공사비_기초금액기준_원(수식)", "㎡당공사비_추정가격기준_원(수식)"]   # 05 대표값 2종
ADJ5_COLS = ["보정_㎡당공사비_기초금액기준_원(수식)", "보정_㎡당공사비_추정가격기준_원(수식)"]


_ILLEGAL = re.compile("[\x00-\x08\x0b\x0c\x0e-\x1f\ud800-\udfff]")   # openpyxl 이 거부하는 제어문자·서로게이트(API 업종명·LLM 출력에 섞여 들어옴)


def _clean(v):
    if v is None or v is pd.NA or v is pd.NaT:
        return None
    if isinstance(v, str):
        v = _ILLEGAL.sub("", v)
    if isinstance(v, float) and (math.isnan(v) or math.isinf(v)):
        return None
    if isinstance(v, (pd.Timestamp, dt.datetime)):
        return v.to_pydatetime() if isinstance(v, pd.Timestamp) else v
    if hasattr(v, "item"):
        try:
            return v.item()
        except (ValueError, TypeError):
            return str(v)
    if isinstance(v, str) and v in ("nan", "None", "<NA>", "NaT"):
        return ""
    if isinstance(v, (dict, list, tuple, set)):
        return json.dumps(v, ensure_ascii=False, default=str)
    return v


def _write_df(ws, df: pd.DataFrame, money_cols: Optional[List[str]] = None, blue_cols: Optional[List[str]] = None,
              yellow_cols: Optional[List[str]] = None, widths: Optional[Dict[str, int]] = None) -> None:
    cols = list(df.columns)
    ws.append(cols)
    for c in range(1, len(cols) + 1):
        cell = ws.cell(row=1, column=c)
        cell.font, cell.fill = BOLD, HEAD
        cell.alignment = Alignment(wrap_text=True, vertical="center")
    for row in df.itertuples(index=False):
        ws.append([_clean(v) for v in row])
    money_idx = [cols.index(c) + 1 for c in (money_cols or []) if c in cols]
    blue_idx = [cols.index(c) + 1 for c in (blue_cols or []) if c in cols]
    yellow_idx = [cols.index(c) + 1 for c in (yellow_cols or []) if c in cols]
    for r in range(2, ws.max_row + 1):
        for c in range(1, len(cols) + 1):
            cell = ws.cell(row=r, column=c)
            cell.font = BLUE if c in blue_idx else FONT
            if c in money_idx:
                cell.number_format = MONEY
            if c in yellow_idx:
                cell.fill = YELLOW
    for i, c in enumerate(cols, 1):
        w = (widths or {}).get(c, min(40, max(10, len(str(c)) * 1.6 + 4)))
        ws.column_dimensions[get_column_letter(i)].width = w
    ws.freeze_panes = "B2"
    if ws.max_row > 1:
        ws.auto_filter.ref = ws.dimensions


NOTICE_COLS = ["공고번호", "공고차수", "공고키", "공고명", "공고종류", "재공고여부", "공고일시", "입찰마감일시", "개찰일시", "공고기관",
               "수요기관", "공사현장지역", "추정가격", "기초금액", "주공종명", "부공종명", "면허제한업종", "공종", "공종근거", "사업유형", "사업유형_원분류",
               "낙찰금액_API", "낙찰률_API", "낙찰자", "참가업체수", "낙찰하한율", "이전공고번호",
               "시설ID", "시설명", "프로젝트ID", "세부사업명", "프로젝트키", "최신여부", "대체공고번호", "대표선정사유", "사전규격번호", "첨부파일수", "상세URL"]


def bundled_trades_by_project(trade: pd.DataFrame) -> Dict[str, str]:
    """프로젝트별 통합발주 표기: 04 행 중 포함공종(종합, 없으면 API)이 배정 공종 외에 더 있는 공고만 '건축(+토목·조경); 기계설비(+가스)' 로."""
    out: Dict[str, List[str]] = {}
    col = "포함공종(종합)" if trade is not None and not trade.empty and "포함공종(종합)" in trade.columns else "포함공종_API"
    if trade is None or trade.empty or col not in trade.columns:
        return {}
    for pid, tr, inc in zip(trade["프로젝트ID"], trade["공종"], trade[col]):
        parts = [p for p in str(inc or "").split("·") if p]
        extra = [p for p in parts if p != str(tr)]
        if not extra:
            continue
        item = f"{tr}(+{'·'.join(extra)})"
        if item not in out.setdefault(str(pid), []):
            out[str(pid)].append(item)
    return {k: "; ".join(v) for k, v in out.items()}


def build_workbook(path: str, facility: pd.DataFrame, hist: pd.DataFrame, latest: pd.DataFrame,
                   trade: pd.DataFrame, logs: pd.DataFrame, notes: pd.DataFrame, meta: Optional[Dict] = None,
                   trade_detail: Optional[pd.DataFrame] = None, servc: Optional[pd.DataFrame] = None) -> None:
    """servc = link-servc 의 연결 표(없으면 None). None 이면 워크북은 예전과 똑같이 만들어진다(08 시트·05 의 용역 열 없음).
    hist(02 시트에 싣는 공사 공고)의 가장 늦은 공고일이 '공사 공고 수집 종료일' 이다 — 08 의 '공사수집종료까지_일수' 와 05 의 '발주완료(최근)' 이 이 날을 기준으로 한다.
    hist 에서 공고일을 하나도 읽지 못하면 그 칸은 비고 05 는 '발주완료' 로만 적는다."""
    wb = Workbook()
    # ── README ──
    ws = wb.active
    ws.title = "README"
    meta = meta or {}
    # 공사 공고 수집 종료일 = 이 워크북에 실리는 공사 공고(02 시트) 가운데 가장 늦은 공고일. config 의 수집 기간(period.end)이 아니라 실린 자료에서 읽는다 —
    # 수집 기간만 늘렸거나 collect 뒤에 discover~dedup 을 아직 다시 돌리지 않았을 때 기준일이 자료보다 앞서가면 '발주완료(최근)' 이 '발주완료' 로 바뀌기 때문이다.
    # 읽을 공고일이 없으면 NaT → 08 의 일수 칸을 비우고 '최근' 을 가리지 않는다
    cn_end = _parse_dt(hist["공고일시"].reset_index(drop=True)).max() if (hist is not None and len(hist) and "공고일시" in hist.columns) else pd.NaT
    cn_end_s = "" if pd.isna(cn_end) else cn_end.date().isoformat()
    cn_end_def = "공사 공고 수집 종료일 = 이 워크북에 실린 공사 공고 가운데 가장 늦은 공고일(02 시트)"
    lines = [
        ["LIMAC 공사비 DB — 나라장터 발주공고 기반"],
        ["생성일", dt.date.today().isoformat()],
        ["수집기간", meta.get("period", "")],
        ["출처", "조달청 나라장터 입찰공고정보서비스(공공데이터포털) + 공고 첨부문서(공고문·현장설명서 등) LLM 추출"],
        ["금액 기준", "낙찰가 아님. 대표값을 두 기준으로 함께 제시: ① 기초금액 기준 = 기초금액 + 도급자관급 + 관급자관급 + 기타금액, ② 추정가격 기준 = 추정가격 + 부가세 + 도급자관급 + 관급자관급 + 기타금액. 조달청이 공고한 공사는 기초금액이 조사금액 기준이라 ①이 ②보다 2~3% 낮을 수 있음. 낙찰금액·낙찰률은 참고 컬럼(금액 합계 수식에 사용하지 않음)."],
        ["금액 구성(04)", "요청 양식 순서: 추정가격 | 부가세 | 도급자관급비 | 추정금액 계ⓐ | 관급자관급비ⓑ | 기타금액ⓒ | 금액 합계(ⓐ+ⓑ+ⓒ). 관급 금액은 부가세 포함 금액. 부가세 = API 부가세 → 문서 부가세 → 기초금액−추정가격 → 10% 가정 순('부가세_출처'에 표시. 10% 에 못 미치는 부가세라도 기초금액이 추정가격×1.1 과 맞으면 채택하지 않음 — 문서의 0·단위 오독 방어). 기초금액(문서)은 추정가격의 100~110.5% 범위일 때만 사용(추정가격이 API·문서 어디에도 없는 행은 문서 기초금액이 그대로 합계). 관급은 API 값이 있으면 API, 없으면 문서 값 — 단, 문서의 관급 금액이 그 공고의 기초금액·추정가격(관급자관급액은 부가세도)·문서의 총공사금액과 같은 금액이거나 문서의 총공사금액이 두 관급액의 합과 같으면 도급액을 관급 칸으로 읽은 것으로 보아 싣지 않는다(06 '도급자관급액'·'관급자관급액' 행에 경고와 까닭). 같은 금액이라도 빼면 문서의 총공사금액에서 오히려 더 멀어지는 값은 그대로 싣고 같은 행에 '그대로 실음' 경고를 남긴다. 기타금액ⓒ = 한전불입금·분담금 등 문서에 별도로 적힌 금액(보강 추출 extract-extra 결과, '기타금액_판정'이 가산일 때만 합계에 포함 — 관급 칸과 같은 값이면 이중계산 방지로 제외)."],
        ["부가세·면세", "과세구분: 부가세율(부가세÷추정가격)이 10% 면 과세, 0 이면 면세, 그 사이면 부분면세(국민주택규모 이하 주택 건설용역은 조세특례제한법 제106조로 면세 — 공고에 적힌 부가세액을 그대로 싣고 재계산하지 않음). 출처가 '기초금액−추정가격' 인데 10% 미만이면 면세 또는 조달청 조사금액 차이일 수 있어 '확인필요'. API 의 부가세 0 은 '미입력'과 구분되지 않아 면세 근거로 쓰지 않는다(면세는 기초금액=추정가격 이거나 문서 부가세로만 판정)."],
        ["공사개요(01)", "노란 셀은 문서 추출값이며 직접 고칠 수 있다(연면적은 05 ㎡당 공사비에 반영). 연면적은 본 건축 공고의 문서 값 → 여러 공고가 같게 적은 값 → 다른 공종 공고의 값 순으로 고르며, 철거·목업·부속 건물 같은 부분 공사 공고(건축·기타 공종이면서 추정가격이 프로젝트 최대 공고의 20% 미만)의 면적은 연면적·건축면적에 쓰지 않는다(그 값밖에 없으면 빈칸 — 까닭과 그 면적은 '개요_출처'. 최대 공고의 20% 이상인 부분 공사 공고의 면적은 그대로 실리므로 ㎡당 공사비가 유난히 큰 행은 01 의 연면적을 확인할 것. 층수·구조·공사기간은 값이 있는 첫 공고의 것이라 부분 공사 공고에서 올 수 있다). 건폐율_pct·용적률_pct 는 문서에 적힌 값(%), 건폐율_계산 = 건축면적 ÷ 대지면적. 용적률은 용적률 산정용 연면적(지하층 등 제외)이 필요해 계산하지 않는다. 발주년월은 건축 대표 공고의 최초 공고월('발주년월_기준' 참조). 건축법상용도(수식)는 표2 분류(검색어 기준)이고 문서에 적힌 용도는 '건축법상용도_문서'·'용도'. 01·04 시트는 표 아래에 주석 행을 두지 않는다(행 = 프로젝트·대표 공고)."],
        ["공종", "건축·전기·정보통신·소방·조경·기계설비·토목·기타. 05 '건축공종계' = 건축+기계설비+토목+조경(요청 양식의 건축공종). '기타' = 주공종명·공고명으로 공종을 판별하지 못한 공고로, 04 '공종_비고'에 미분류 사유와 공사 성격을 적음. 공고명 끝의 '(전기)'·'(통신)'·'(소방)' 꼬리표는 그 공종으로 분류."],
        ["대표 공고(합산 대상)", "한 시설 후보(시설ID×사업유형)의 공고를 세부사업(한 건물·한 단지·한 현장)으로 나누고(나뉘면 프로젝트ID 끝에 -01, -02 …), 세부사업 안에서 재공고·재발주로 대체된 공고와 취소공고·구차수를 뺀 모든 공고를 합산한다. 같은 공종 공고가 여럿이면(기계소방/전기소방, 차수별 발주, 부대공사) 모두 합산 — 06 '세부사업 분리'·'같은 공종 합산' 에서 확인. 공고명의 차수 표기가 다르면('조성사업' / '조성사업(2차분)') 재공고가 아니라 차수별 발주로 보아 둘 다 합산하고, 공법·자재 선정용 제안서 제출 안내 공고(공사 발주 아님)는 합산하지 않는다(06 '제안 안내 공고 제외'). 이력은 02시트 보존(대표선정사유에 제외 사유)."],
        ["프로젝트ID", "시설ID-사업유형코드(N 신축 / E 증축 / R 리모델링 / ER 증축·리모델링). 같은 시설의 신축과 리모델링은 별도 프로젝트로 집계되며, ㎡당 공사비 비교는 반드시 사업유형이 같은 프로젝트끼리 할 것. 한 시설 후보가 여러 세부사업으로 나뉘면 끝에 -01, -02 … 가 붙는다(예: F0123-N-02). 같은 공고가 여러 시설 후보에 걸려 같은 사업이 두 프로젝트로 집계되면 05 '시설간_중복'·'중복_제외권장' 에 표시된다."],
        ["색상", "파란 글자=원천값(API/문서), 검은 글자=수식, 노란 배경=사용자 입력·검수 셀 (01 시트의 연면적·층수 등은 문서 추출값이며 직접 고치면 05 ㎡당 공사비에 반영됨)"],
        ["단위", "금액: 원 / 면적: ㎡ / 기간: 일"],
        ["포함공종", "04 '포함공종(종합)' = 그 공고(계약) 금액에 묶여 있는 공종(예: 건축·토목·기계설비·조경). 출처별로 '포함공종_API'(주공종명·부공종명)와 '포함공종_문서'(공고문·현장설명서의 공사금액 표 구분란. 출처 LLM=추출 결과, 규칙=텍스트에서 자동 탐지)를 따로 둠. 금액은 공종별로 나뉘지 않으므로 05 의 '건축(수식)' 은 그 안의 토목·조경 등을 포함한 값. 05 '통합발주_포함공종' 은 프로젝트의 통합발주 공고만 모아 표기. '분리발주_언급공종_문서' 는 공고문이 별도 발주라고 언급한 공종(참고)."],
        [],
        ["시트", "내용"],
        ["01_시설마스터", "시설(프로젝트) 단위 공사개요 — 조달청 '공사비 정보광장' 항목 틀(공사명·공사규모·면적[연면적·대지면적·건축면적·조경면적·건폐율·용적률]·높이[기준층층고·최고높이]·주차대수·현장위치·발주년월·공사기간·건축법상용도·설계특이사항). 값은 공고 첨부문서 LLM 추출(나라장터 API 에는 면적 항목이 없음), 건축 공고 문서 우선이며 출처 공종을 '개요_출처'에 표시. 조경면적·건폐율·용적률·높이·주차대수·설계특이사항은 보강 추출(extract-extra) 뒤에 채워짐. 공종 커버리지 포함"],
        ["02_공고목록_전체", "수집된 모든 공고(차수·재공고·취소 이력 포함)"],
        ["03_공고목록_최신", "대표 공고(합산 대상): 세부사업·공종별로 재공고에 대체된 공고를 뺀 모든 공고"],
        ["04_공사비DB_공종별", "대표 공고 기준 금액 구성(API+문서): 원천값(파란 글자) → 채택값 수식(추정가격·부가세·도급자관급비·추정금액 계ⓐ·관급자관급비ⓑ·기타금액ⓒ) → 금액합계 2종(추정가격 기준·기초금액 기준). 관급자재는 API 값 우선, 없으면 문서 추출값"],
        ["04b_공종세부내역(문서)", "건축·통합발주 공고의 첨부 내역서 집계표(또는 공고문 공사금액 표)에서 읽은 공종별 금액(2차 추출, extract-trades). 비율 = 금액 / 공종 행 합(공통가설·간접비·관급자재·부가세·합계·배분제외 행 제외), 배분금액 = 비율 × 04 금액합계(기초금액 기준). '합계' = 여러 공종을 합친 도급비·총공사비 행, '배분제외' = 이설비·분담금·용역비 등 도급 공사비 밖 항목(둘 다 배분 안 함). 배분적용 'N(사유)' 공고(공종 행 합이 집계표의 30% 미만 등 일부 행만 읽힌 의심)는 비율을 비운다. 참고용이며 05 총공사비합계에는 더하지 않는다"],
        ["05_공사비DB_시설합산", "시설 단위 공종 합산(SUMIFS, 공종별 열은 기초금액 기준), 건축공종계, 총공사비합계 2종(기초금액 기준·추정가격 기준)과 각각의 ㎡당 공사비, 보정계수 입력, 공종 누락 경고(공종누락경고 = 그 공종 공고가 없음, 공종누락_통합발주제외 = 그 가운데 통합발주로 다른 공고에 묶인 공종을 뺀 것). 행 단위 = 프로젝트(시설ID×사업유형, 여러 사업이 묶인 시설 후보는 세부사업별로 나눈 행 — 세부사업명·세부사업수). 문서배분_*(수식) = 다른 공종 공고 안에 묶인 그 공종의 배분금액 합(해당 공종 공고 자신의 몫은 제외, 참고), 건축_순건축추정 = 건축(수식) − 건축 공고 안의 비건축 배분, 문서배분_중복경고 = 건축 집계표에 전기·정보통신·소방 금액이 있는데 분리발주 공고도 있음. 시설간_중복 = 다른 시설 후보의 프로젝트와 대표 공고가 겹치는 경우 그 관계, 중복_제외권장 = Y 인 행을 빼면 같은 공고가 두 번 합산되지 않는다(합계·평균을 낼 때는 Y 행을 필터로 제외. '일부 겹침' 은 그것만으로는 Y 가 붙지 않는다 — Y 행을 뺀 뒤에도 두 번 남는 공고가 있는지는 06_검증로그 '시설 간 중복 프로젝트' 의 남은 건수로 확인하고, 남아 있을 때만 검수에서 판단)"],
        [S_MISS_SUM, f"실제 누락 공종 수(05 '{MISS_CNT_COL}' — 건축·전기·정보통신·소방 가운데 공고도 없고 통합발주로 묶였다는 기록도 없는 공종의 수)별 프로젝트 수를 "
                     "사업유형·시설 유형(표2 분류)별로 센 표. 왼쪽 = 전체, 오른쪽 = 중복_제외권장 Y 를 뺀 것. 모든 칸이 05 를 세는 수식"],
        [f"{S_MISS_N.format(k=0)} ~ {S_MISS_N.format(k=len(MISS_TRADES))}",
         "실제 누락 공종 수별 프로젝트 목록(시트 하나 = 누락 0개 … 4개). 공종별 금액·총공사비합계·㎡당 공사비·연면적 등은 05 의 그 프로젝트 행을 읽는 수식이라 05 와 늘 같은 값이다. "
         "어느 프로젝트가 어느 시트에 오는지는 워크북을 만들 때 정해진다(04 를 고쳐 누락 수가 달라졌으면 excel 단계를 다시 실행)"],
        ["06_검증로그", "API-문서 교차검증, 재발주 금액변동, 정합성 점검 결과"],
        ["07_추출노트", "첨부파일별 다운로드·텍스트 추출 결과"],
    ]
    if servc is not None:
        lines += [
            [S_SERVC, "용역 입찰공고(설계·건설사업관리/감리·폐기물 처리) 가운데 이 DB 의 시설에 이어진 공고(link-servc). 한 줄 = 용역 공고 × 시설. "
                      "공사비가 아니라 참고 자료다 — 05 의 총공사비합계·㎡당 공사비에는 들어가지 않는다."],
            ["용역(08·05)", "연결 조건: 용역 공고명에 시설 이름이 있고, 수요기관이 그 시설 공사 공고의 수요기관과 같은 발주 주체(같은 지자체·교육청·기관. 광역 이름이 바뀐 경우 포함). "
                          "어느 프로젝트의 용역인지는 공고 시점이 맞는 프로젝트 가운데 공고명의 사업 이름·사업유형·공사와의 간격으로 정하고, 가릴 수 없으면 프로젝트ID 를 비워 둔다. "
                          "05 의 '용역_○○_추정가격' = 08 에서 그 프로젝트·그 종류이면서 합산여부가 Y 인 줄의 추정가격 합(부가세 제외. 합산여부 Y 인 공고가 없으면 빈칸 — 0 이 아니다. "
                          "N 으로만 이어진 공고는 08 에 있을 수 있다). "
                          "05 의 '발주완료_폐기물용역' = 그 프로젝트에 합산여부 Y 인 폐기물처리 용역 공고가 하나라도 있으면 '발주완료'(공사 발주가 끝났다고 보는 보조지표. "
                          "빈칸은 '미완료' 가 아니라 '판단할 근거 없음' — 폐기물이 적은 공사·수의계약은 용역 공고가 없다. "
                          + (f"그 폐기물 용역 공고가 모두 공사 공고 수집 종료일({cn_end_s})에서 6개월({SERVC_RECENT_DAYS}일) 안쪽이거나 그 뒤면 '{SERVC_DONE_RECENT}' 으로 따로 적는다"
                             f"(08 의 '{SERVC_AGE_COL}' 가 모두 {SERVC_RECENT_DAYS} 미만. {cn_end_def}) — 전기·정보통신·소방 공고가 더 나올 수 있다. 발주가 끝난 사업을 고를 때는 '발주완료' 만 고를 것"
                             "('발주완료' 인 행도 그 사업의 공사 공고가 최근이면 — 01 의 '최종공고일' — 아직 발주 중일 수 있다). "
                             if cn_end_s else "폐기물 용역 공고가 최근인 사업은 전기·정보통신·소방 공고가 더 나올 수 있다. ") +
                          "'발주완료' 여도 그 행에 공사비가 모두 들어 있다는 뜻은 아니므로 공종누락경고를 함께 볼 것). "
                          "합산여부 N 의 까닭은 '합산제외_까닭' 열: 취소·재공고로 대체된 공고, 여러 시설을 묶은 공고, 연간·단가 계약, 부대 용역(타당성·VE·공모 관리 등), "
                          "유지보수 성격의 공사에 딸린 용역, 같은 계열의 다른 기관(다른 소속기관·지사·학교)이 낸 공고, 다른 시설에도 이어진 공고, 연결이 추정인 줄, 프로젝트를 정하지 못한 줄. "
                          "노란 칸(합산여부·프로젝트ID)은 고칠 수 있고 05 가 따라 바뀐다 — 프로젝트ID 가 빈 줄은 프로젝트ID 도 적어야 더해지고, 종류가 '기타' 인 줄은 금액 열에 들어가지 않는다"
                          "(excel 을 다시 만들면 처음 값으로 돌아간다). "
                          "이어진 공고가 없다고 용역이 없었던 것은 아니다: 설계 공고는 공사보다 1~3년 먼저 나므로 수집 기간의 첫 2~3년에 공고된 공사의 설계 공고는 용역 수집 기간보다 앞일 수 있고, "
                          "소액 용역은 공고 없이 수의계약으로 집행되며, 2024-04 이전 LH 공공주택의 설계·건설사업관리는 LH 자체 시스템으로 발주됐다. "
                          "연결은 공고명과 발주 기관만으로 한 것이므로 금액을 쓰기 전에 08 의 공고명을 확인할 것."],
        ]
    for l in lines:
        ws.append(l)
    ws["A1"].font = Font(name="Arial", size=12, bold=True)
    for r in range(2, ws.max_row + 1):
        for c in (1, 2):
            ws.cell(row=r, column=c).font = FONT
    ws.column_dimensions["A"].width, ws.column_dimensions["B"].width = 22, 110

    # ── 01 시설마스터 — 공사개요는 조달청 '공사비 정보광장' 항목 순서 ──
    ws = wb.create_sheet("01_시설마스터")
    fac = facility.copy()
    for c in FAC_OVERVIEW_COLS + ["검수_포함여부", "메모"]:
        if c not in fac.columns:
            fac[c] = "포함" if c == "검수_포함여부" else None
    head = [c for c in ["프로젝트ID", "시설ID", "시설명", "세부사업명", "세부사업수", "사업유형", "표2_대분류", "표2_중분류", "검색어", "수요기관"] if c in fac.columns]
    tail = ["검수_포함여부", "메모"]
    rest = [c for c in fac.columns if c not in head + FAC_OVERVIEW_COLS + tail]
    fac = fac[head + FAC_OVERVIEW_COLS + rest + tail]
    area_cols = ["연면적_m2", "대지면적_m2", "건축면적_m2", "조경면적_m2"]
    _write_df(ws, fac, money_cols=area_cols,
              blue_cols=[c for c in fac.columns if c not in ("검수_포함여부", "메모") and "(수식)" not in c],
              yellow_cols=FAC_INPUT_COLS + ["검수_포함여부", "메모"],
              widths={"시설명": 34, "세부사업명": 40, "공사명": 46, "공사규모(수식)": 28, "현장위치": 30, "건축법상용도(수식)": 26, "설계특이사항": 44, "용도": 24, "수요기관": 24, "개요_출처": 18})
    f1 = {c: get_column_letter(i + 1) for i, c in enumerate(fac.columns)}
    for r in range(2, ws.max_row + 1):
        for c in area_cols:
            ws[f"{f1[c]}{r}"].number_format = "#,##0.0"
        st, ub, ug = f"{f1['구조']}{r}", f"{f1['지하층수']}{r}", f"{f1['지상층수']}{r}"
        ws[f"{f1['공사규모(수식)']}{r}"] = (f'=TRIM({st}&" "&IF(N({ub})>0,"지하 "&{ub}&"층","")&IF(AND(N({ub})>0,N({ug})>0),"~","")'
                                         f'&IF(N({ug})>0,"지상 "&{ug}&"층",""))')
        site, bld = f"{f1['대지면적_m2']}{r}", f"{f1['건축면적_m2']}{r}"
        ws[f"{f1['건폐율_계산(수식)']}{r}"] = f'=IF(AND(ISNUMBER({site}),{site}>0,ISNUMBER({bld})),{bld}/{site},"")'
        ws[f"{f1['건폐율_계산(수식)']}{r}"].number_format = "0.00%"
        if "표2_대분류" in f1 and "표2_중분류" in f1:
            big, mid = f"{f1['표2_대분류']}{r}", f"{f1['표2_중분류']}{r}"
            ws[f"{f1['건축법상용도(수식)']}{r}"] = f'={big}&IF(OR({mid}="",{mid}="-"),"","("&{mid}&")")'
        for c in ("공사규모(수식)", "건폐율_계산(수식)", "건축법상용도(수식)"):
            ws[f"{f1[c]}{r}"].font = FONT

    # ── 02 / 03 공고목록 ──
    for title, df in (("02_공고목록_전체", hist), ("03_공고목록_최신", latest)):
        ws = wb.create_sheet(title)
        cols = [c for c in NOTICE_COLS if c in df.columns]
        _write_df(ws, df[cols], money_cols=["추정가격", "기초금액", "낙찰금액_API"], blue_cols=cols,
                  widths={"공고명": 60, "수요기관": 24, "공고기관": 24, "상세URL": 40, "프로젝트키": 26, "세부사업명": 36})

    # ── 04 공종별 — 원천값(API·문서) → 채택값(요청 양식 순서: 추정가격 | 부가세 | 도급자관급비 | 추정금액 계ⓐ | 관급자관급비ⓑ | 기타금액ⓒ | 금액합계) ──
    ws = wb.create_sheet(S_TRADE)
    tcols = ["프로젝트ID", "시설ID", "시설명", "세부사업명", "사업유형", "공종", "공종_비고", "포함공종(종합)", "포함공종_API", "포함공종_문서", "포함공종_출처",
             "분리발주_언급공종_문서", "공고번호", "공고차수", "공고키", "공고명", "공고일시", "수요기관",
             "추정가격_API", "기초금액_API", "부가세_API", "관급자재_API", "도급자관급액_API", "관급자관급액_API",
             "추정가격_문서", "기초금액_문서", "부가세_문서", "도급자관급액_문서", "관급자관급액_문서", "기타금액_문서", "기타금액_항목_문서", "기타금액_판정",
             "추정가격(수식)", "부가세(수식)", "도급자관급비(수식)", "추정금액_계ⓐ(수식)", "관급자관급비ⓑ(수식)", "기타금액ⓒ(수식)",
             "금액합계_추정가격기준(수식)", "금액합계_기초금액기준(수식)", "부가세_출처(수식)", "부가세율(수식)", "과세구분(수식)", "면세_문구_문서",
             "예산금액_API", "낙찰금액_API", "낙찰률_API", "낙찰률(수식)", "낙찰자", "낙찰하한율",
             "공사기간_일_문서", "추정가격_차이율(수식)", "신뢰도", "근거문구", "출처파일", "상세URL"]
    t = trade.copy()
    for c in tcols:
        if c not in t.columns:
            t[c] = None
    t = t[tcols]
    amount_formula_cols = ["추정가격(수식)", "부가세(수식)", "도급자관급비(수식)", "추정금액_계ⓐ(수식)", "관급자관급비ⓑ(수식)", "기타금액ⓒ(수식)",
                           "금액합계_추정가격기준(수식)", "금액합계_기초금액기준(수식)"]
    _write_df(ws, t, money_cols=["추정가격_API", "기초금액_API", "부가세_API", "관급자재_API", "도급자관급액_API", "관급자관급액_API",
                                 "추정가격_문서", "기초금액_문서", "부가세_문서", "도급자관급액_문서", "관급자관급액_문서", "기타금액_문서",
                                 "예산금액_API", "낙찰금액_API"] + amount_formula_cols,
              blue_cols=[c for c in tcols if "(수식)" not in c],
              widths={"시설명": 30, "세부사업명": 36, "공고명": 55, "공종_비고": 46, "포함공종(종합)": 24, "포함공종_API": 20, "포함공종_문서": 22, "분리발주_언급공종_문서": 22,
                      "기타금액_항목_문서": 30, "기타금액_판정": 26, "과세구분(수식)": 22, "면세_문구_문서": 30, "근거문구": 50, "출처파일": 30, "상세URL": 40})
    col = {c: get_column_letter(i + 1) for i, c in enumerate(tcols)}
    for r in range(2, ws.max_row + 1):
        P, B, V = f"{col['추정가격_API']}{r}", f"{col['기초금액_API']}{r}", f"{col['부가세_API']}{r}"
        Ga, G1a, G2a = f"{col['관급자재_API']}{r}", f"{col['도급자관급액_API']}{r}", f"{col['관급자관급액_API']}{r}"
        G1, G2 = f"{col['도급자관급액_문서']}{r}", f"{col['관급자관급액_문서']}{r}"
        Pd, Bd, Vd = f"{col['추정가격_문서']}{r}", f"{col['기초금액_문서']}{r}", f"{col['부가세_문서']}{r}"
        E, EJ = f"{col['기타금액_문서']}{r}", f"{col['기타금액_판정']}{r}"
        PX, VX, D1 = f"{col['추정가격(수식)']}{r}", f"{col['부가세(수식)']}{r}", f"{col['도급자관급비(수식)']}{r}"
        TA, D2, EC = f"{col['추정금액_계ⓐ(수식)']}{r}", f"{col['관급자관급비ⓑ(수식)']}{r}", f"{col['기타금액ⓒ(수식)']}{r}"
        VS, VR = f"{col['부가세_출처(수식)']}{r}", f"{col['부가세율(수식)']}{r}"
        ws[PX] = f'=IF(AND(ISNUMBER({P}),{P}>0),{P},IF(AND(ISNUMBER({Pd}),{Pd}>0),{Pd},""))'      # 0 은 값으로 보지 않는다(문서 오독 방지)
        # 부가세 출처: API 부가세 → 문서 부가세 → 기초금액−추정가격(0~10.5% 일 때만) → 10% 가정. 범위를 벗어난 값(오입력)은 건너뛴다.
        # 단, 9.9% 에 못 미치는 부가세(면세·부분면세로 읽힘)인데 기초금액이 추정가격×1.1 과 ±1% 로 맞으면 그 부가세는 채택하지 않는다
        # (문서의 0·천원 오독·다른 범위의 부가세가 '면세'를 만들고 추정가격 기준 합계를 10% 낮추던 문제). 기초금액과 어긋나지 않는 공고 부가세는 적힌 대로 싣는다
        ok = lambda v: f'AND(ISNUMBER({v}),{v}>=0,{v}<={PX}*0.105)'
        okd = lambda b: f'AND(ISNUMBER({b}),{b}-{PX}>=0,{b}-{PX}<={PX}*0.105)'
        bb = f'IF(ISNUMBER({B}),{B},IF(ISNUMBER({Bd}),{Bd},""))'
        okv = lambda v: f'AND({ok(v)},NOT(AND({v}<{PX}*0.099,ISNUMBER({bb}),ABS(N({bb})-{PX}*1.1)<={PX}*0.011)))'
        ws[VS] = (f'=IF(NOT(ISNUMBER({PX})),"",IF({okv(V)},"API",IF({okv(Vd)},"문서",IF({okd(B)},"기초금액-추정가격",'
                  f'IF({okd(Bd)},"기초금액(문서)-추정가격","10% 가정")))))')
        ws[VX] = (f'=IF({VS}="","",IF({VS}="API",{V},IF({VS}="문서",{Vd},IF({VS}="기초금액-추정가격",{B}-{PX},'
                  f'IF({VS}="기초금액(문서)-추정가격",{Bd}-{PX},ROUND({PX}*0.1,0))))))')
        # 관급자재: API 도급자+관급자 → API 합계(관급자 설치분으로 봄) → 문서 도급자·관급자
        ws[D1] = f'=IF(N({G1a})+N({G2a})>0,N({G1a}),IF(N({Ga})>0,0,N({G1})))'
        ws[D2] = f'=IF(N({G1a})+N({G2a})>0,N({G2a}),IF(N({Ga})>0,N({Ga}),N({G2})))'
        ws[TA] = f'=IF(ISNUMBER({PX}),{PX}+N({VX})+{D1},"")'
        ws[EC] = f'=IF(LEFT({EJ}&"",2)="가산",N({E}),0)'
        # 금액합계 2종. 기초금액 기준 = [기초금액(API) → 기초금액(문서: 추정가격의 100~110.5% 일 때만) → 추정가격+부가세] + 관급 + 기타
        #              추정가격 기준 = [추정가격+부가세 → 기초금액] + 관급 + 기타
        # (AND/OR 는 뒤 인수까지 계산하므로 추정가격이 없는 행에서 오류가 나지 않게 IF 를 겹쳐 쓴다)
        base_b = (f'IF(AND(ISNUMBER({B}),{B}>0),{B},IF(ISNUMBER({PX}),IF({okd(Bd)},{Bd},{PX}+N({VX})),'
                  f'IF(AND(ISNUMBER({Bd}),{Bd}>0),{Bd},0)))')
        ws[f"{col['금액합계_기초금액기준(수식)']}{r}"] = f"={base_b}+{D1}+{D2}+{EC}"
        ws[f"{col['금액합계_추정가격기준(수식)']}{r}"] = f"=IF(ISNUMBER({PX}),{PX}+N({VX}),{base_b})+{D1}+{D2}+{EC}"
        ws[VR] = f'=IF(AND(ISNUMBER({PX}),{PX}>0,ISNUMBER({VX})),{VX}/{PX},"")'
        ws[VR].number_format = "0.00%"
        ws[f"{col['과세구분(수식)']}{r}"] = (f'=IF({VS}="","",IF({VS}="10% 가정","미확인(10% 가정)",IF({VR}<0.001,"면세",IF({VR}<0.099,'
                                          f'IF(OR({VS}="API",{VS}="문서"),"부분면세","확인필요(면세 또는 조달청 조사금액 차이)"),IF({VR}<=0.101,"과세","확인필요")))))')
        ws[f"{col['추정가격_차이율(수식)']}{r}"] = f'=IF(AND(ISNUMBER({P}),ISNUMBER({Pd}),{P}<>0),({Pd}-{P})/{P},"")'
        ws[f"{col['추정가격_차이율(수식)']}{r}"].number_format = "0.00%"
        # 낙찰률(수식) = 낙찰금액 / 기초금액 (참고. API 낙찰률과 대조용)
        W = f"{col['낙찰금액_API']}{r}"
        ws[f"{col['낙찰률(수식)']}{r}"] = f'=IF(AND(ISNUMBER({W}),ISNUMBER({B}),{B}<>0),{W}/{B},"")'
        ws[f"{col['낙찰률(수식)']}{r}"].number_format = "0.00%"
        for c in tcols:
            if "(수식)" in c:
                ws[f"{col[c]}{r}"].font = FONT

    # ── 04b 공종세부내역(문서) — 내역서 집계표의 공종별 금액(2차 추출). 금액은 04 와 별도 시트에 두어 05 SUMIFS 에 섞이지 않게 한다 ──
    ws = wb.create_sheet(S_TRADE_DETAIL)
    det = (trade_detail if trade_detail is not None else pd.DataFrame()).copy()
    for c in DETAIL_COLS:
        if c not in det.columns:
            det[c] = None
    det = det[DETAIL_COLS]
    _write_df(ws, det, money_cols=["금액_원_문서", "배분금액(수식)", "집계표합계_원_문서"],
              blue_cols=[c for c in DETAIL_COLS if "(수식)" not in c],
              widths={"시설명": 30, "항목명_원문": 28, "출처파일": 36, "근거문구": 40})
    dcol = {c: get_column_letter(i + 1) for i, c in enumerate(DETAIL_COLS)}
    n_d = max(len(det), 1) + 1
    D = {c: f"'{S_TRADE_DETAIL}'!${dcol[c]}$2:${dcol[c]}${n_d}" for c in ("금액_원_문서", "배분금액(수식)", "공고키", "프로젝트ID", "세부공종", "상위공종")}
    t_tot = f"'{S_TRADE}'!${col['금액합계_기초금액기준(수식)']}:${col['금액합계_기초금액기준(수식)']}"
    t_key, t_pid = f"'{S_TRADE}'!${col['공고키']}:${col['공고키']}", f"'{S_TRADE}'!${col['프로젝트ID']}:${col['프로젝트ID']}"
    kA, kK, kL = dcol["프로젝트ID"], dcol["공고키"], dcol["세부공종"]
    for r in range(2, len(det) + 2):
        base_sum = f"SUMIFS({D['금액_원_문서']},{D['공고키']},${kK}{r},{D['프로젝트ID']},${kA}{r}"
        denom = "(" + base_sum + ")" + "".join(f'-{base_sum},{D["세부공종"]},"{x}")' for x in DETAIL_EXCLUDED) + ")"   # 바깥 괄호 닫기
        excl = ",".join(f'${kL}{r}="{x}"' for x in DETAIL_EXCLUDED)
        amt = f"${dcol['금액_원_문서']}{r}"
        ratio = f"{dcol['비율(수식)']}{r}"
        flag = f"${dcol['배분적용']}{r}"                      # 'Y' 또는 'N(사유)'(비면 Y 로 본다)
        ws[ratio] = f'=IF(OR({excl},LEFT({flag}&"Y",1)="N"),"",IF(AND(ISNUMBER({amt}),{denom}>0),{amt}/{denom},""))'
        ws[ratio].number_format = "0.00%"
        ws[f"{dcol['배분금액(수식)']}{r}"] = f'=IF(ISNUMBER({ratio}),{ratio}*SUMIFS({t_tot},{t_key},${kK}{r},{t_pid},${kA}{r}),"")'
        for c in ("비율(수식)", "배분금액(수식)"):
            ws[f"{dcol[c]}{r}"].font = FONT
    nd = len(det) + 3
    ws[f"A{nd}"] = ("주: 금액_원_문서는 첨부 내역서 집계표(또는 공고문 공사금액 표)에서 읽은 값으로, 금액기준(부가세 포함/제외/직접공사비)이 공고마다 달라 절대금액을 그대로 쓰지 않는다. "
                    "비율 = 금액 / 같은 공고의 공종 행 합(공통가설·간접비·관급자재·부가세·합계·배분제외 제외), 배분금액 = 비율 × 04 금액합계_기초금액기준(수식). "
                    "배분적용이 'N(…)' 인 공고(공종 행 합이 집계표의 30% 미만·125% 초과, 집계표 합계가 API 금액과 불일치 등 모델이 일부 행만 읽은 의심)는 비율·배분금액을 비운다. "
                    "05 의 문서배분_* 는 다른 공종 공고(주로 건축) 안에 묶인 그 공종의 배분금액을 프로젝트별로 합한 참고값(해당 공종 공고 자신의 몫은 제외)이며 총공사비합계에는 더하지 않는다. "
                    "분리발주 공고가 따로 있는 공종(전기·정보통신·소방)이 건축 집계표에도 있으면 05 '문서배분_중복경고' 에 표시된다.")
    ws[f"A{nd}"].font = FONT

    # ── 05 시설합산 ──
    ws = wb.create_sheet(S_FAC)
    fcols = ["프로젝트ID", "시설ID", "시설명", "세부사업명", "세부사업수", "사업유형", "표2_대분류", "표2_중분류", "수요기관", "연면적_m2", "지상층수", "지하층수", "구조",
             "공고연도"] + [f"{tr}(수식)" for tr in TRADES] + ["건축공종계(수식)"] + TOTAL5_COLS + ["건설공사비지수_보정계수(입력)"] + ADJ5_COLS \
        + ["공종누락경고(수식)", "통합발주_포함공종", MISS_REAL_COL, MISS_CNT_COL] \
        + [f"문서배분_{tr}(수식)" for tr in DETAIL_ALLOC_TRADES] + ["건축_순건축추정(수식)", "문서배분_중복경고(수식)"] \
        + (SERVC5_COLS if servc is not None else []) + ["시설간_중복", "중복_제외권장", "공고건수_최신", "메모"]
    f5 = pd.DataFrame({c: facility[c] if c in facility.columns else None for c in
                       ["프로젝트ID", "시설ID", "시설명", "세부사업명", "세부사업수", "사업유형", "표2_대분류", "표2_중분류", "수요기관", "연면적_m2", "지상층수", "지하층수", "구조",
                        "공고건수_최신", "시설간_중복", "중복_제외권장"]}, index=facility.index)
    f5["공고연도"] = facility["최종공고일"].astype(str).str[:4] if ("최종공고일" in facility.columns and len(facility)) else None
    f5["건설공사비지수_보정계수(입력)"] = 1.0
    f5["통합발주_포함공종"] = f5["프로젝트ID"].map(bundled_trades_by_project(trade))
    f5["메모"] = ""
    for c in fcols:
        if c not in f5.columns:
            f5[c] = None
    f5 = f5[fcols]
    _write_df(ws, f5, money_cols=["연면적_m2"] + [f"{tr}(수식)" for tr in TRADES] + ["건축공종계(수식)"] + TOTAL5_COLS + ADJ5_COLS
              + ["건축_순건축추정(수식)"]
              + [f"문서배분_{tr}(수식)" for tr in DETAIL_ALLOC_TRADES] + SERVC5_COLS[:len(SERVC_KINDS)],
              blue_cols=["프로젝트ID", "시설ID", "시설명", "세부사업명", "세부사업수", "사업유형", "표2_대분류", "표2_중분류", "수요기관", "연면적_m2", "지상층수", "지하층수", "구조",
                         "공고연도", "공고건수_최신", "통합발주_포함공종", "시설간_중복", "중복_제외권장"],
              yellow_cols=["건설공사비지수_보정계수(입력)", "메모"],
              widths={"시설명": 30, "세부사업명": 36, "공종누락경고(수식)": 28, "통합발주_포함공종": 34, MISS_REAL_COL: 28, MISS_CNT_COL: 12, "시설간_중복": 34, SERVC_DONE_COL: 16})
    fc = {c: get_column_letter(i + 1) for i, c in enumerate(fcols)}
    sv = None
    if servc is not None:
        sv = servc.copy()
        for c in SERVC_COLS:
            if c not in sv.columns:
                sv[c] = None
        sv = sv[SERVC_COLS]
        # 공사 공고 수집 종료일까지의 날수(시각은 버리고 날짜끼리 뺀다). 모든 줄에 넣는다 — 08 에서 합산여부·프로젝트ID 를 고쳐도 그 줄의 날수는 그대로 쓰인다.
        # 종료일을 모르거나 공고일시를 읽지 못한 줄은 빈칸(05 에서 '최근' 으로 치지 않는다)
        age = [None] * len(sv) if pd.isna(cn_end) else [None if pd.isna(t) else int((cn_end.normalize() - t.normalize()).days) for t in _parse_dt(sv["공고일시"])]
        sv.insert(SERVC8_COLS.index(SERVC_AGE_COL), SERVC_AGE_COL, pd.Series(age, index=sv.index, dtype=object))
        vcol = {c: get_column_letter(i + 1) for i, c in enumerate(SERVC8_COLS)}
        n_v = max(len(sv), 1) + 1                    # 행 범위를 지정한다(04b 와 같게)
        V = {c: f"'{S_SERVC}'!${vcol[c]}$2:${vcol[c]}${n_v}" for c in ("프로젝트ID", "용역종류", "합산여부", "추정가격", SERVC_AGE_COL)}
    tot_rng = f"'{S_TRADE}'!${col['금액합계_기초금액기준(수식)']}:${col['금액합계_기초금액기준(수식)']}"
    tot2_rng = f"'{S_TRADE}'!${col['금액합계_추정가격기준(수식)']}:${col['금액합계_추정가격기준(수식)']}"
    id_rng, tr_rng = f"'{S_TRADE}'!${col['프로젝트ID']}:${col['프로젝트ID']}", f"'{S_TRADE}'!${col['공종']}:${col['공종']}"
    # 05 연면적은 01_시설마스터(노란 셀, 사용자 수정 가능)를 참조 → 01 에서 고치면 ㎡당 공사비에 반영
    fac_cols = list(fac.columns)
    area_col = get_column_letter(fac_cols.index("연면적_m2") + 1)
    n_fac = max(len(fac), 1) + 1                     # 열 전체($V:$V) 대신 행 범위를 지정(수식 검증 도구·구형 Excel 호환)
    fac_rng = f"'01_시설마스터'!${area_col}$2:${area_col}${n_fac}"
    fac_id_rng = f"'01_시설마스터'!$A$2:$A${n_fac}"
    for r in range(2, ws.max_row + 1):
        for tr in TRADES:
            ws[f"{fc[f'{tr}(수식)']}{r}"] = f'=SUMIFS({tot_rng},{id_rng},$A{r},{tr_rng},"{tr}")'
        first, last = fc[f"{TRADES[0]}(수식)"], fc[f"{TRADES[-1]}(수식)"]
        # 건축공종계 = 요청 양식의 '건축공종'(건축 + 기계 + 토목 + 조경)
        ws[f"{fc['건축공종계(수식)']}{r}"] = "=" + "+".join(f"{fc[f'{tr}(수식)']}{r}" for tr in ("건축", "기계설비", "토목", "조경"))
        # 대표값 2종: 기초금액 기준(공종별 열의 합) / 추정가격 기준(04 금액합계_추정가격기준의 프로젝트 합)
        ws[f"{fc['총공사비합계_기초금액기준(수식)']}{r}"] = f"=SUM({first}{r}:{last}{r})"
        ws[f"{fc['총공사비합계_추정가격기준(수식)']}{r}"] = f"=SUMIFS({tot2_rng},{id_rng},$A{r})"
        A = f"{fc['연면적_m2']}{r}"
        ws[A] = f'=IFERROR(INDEX({fac_rng},MATCH($A{r},{fac_id_rng},0)),"")'
        ws[A].font = FONT
        K = f"{fc['건설공사비지수_보정계수(입력)']}{r}"
        for basis in ("기초금액기준", "추정가격기준"):
            T, U = f"{fc[f'총공사비합계_{basis}(수식)']}{r}", f"{fc[f'㎡당공사비_{basis}_원(수식)']}{r}"
            ws[U] = f'=IF(AND(ISNUMBER({A}),{A}>0),{T}/{A},"")'
            ws[f"{fc[f'보정_㎡당공사비_{basis}_원(수식)']}{r}"] = f'=IF(ISNUMBER({U}),{U}*{K},"")'
        # 공고 자체가 없으면 '누락', 공고는 있는데 금액이 0 이면 '금액없음'
        miss = "&".join(f'IF(COUNTIFS({id_rng},$A{r},{tr_rng},"{tr}")=0,"{tr} ","")' for tr in ("건축", "전기", "정보통신", "소방"))
        zero = "&".join(f'IF(AND(COUNTIFS({id_rng},$A{r},{tr_rng},"{tr}")>0,{fc[f"{tr}(수식)"]}{r}=0),"{tr} ","")'
                        for tr in ("건축", "전기", "정보통신", "소방"))
        ws[f"{fc['공종누락경고(수식)']}{r}"] = (f'=TRIM(IF(TRIM({miss})="","",TRIM({miss})&"누락 ")'
                                            f'&IF(TRIM({zero})="","",TRIM({zero})&"금액없음"))')
        # 통합발주로 다른 공고에 묶인 공종은 '정보가 없는' 것이 아니다 → 공고가 없고(누락) 이 행의 통합발주_포함공종에도 적혀 있지 않은 공종만 남긴다.
        # (묶인 공종의 금액은 묶은 공고의 공종 열에 들어 있다. 통합발주_포함공종이 빈 행은 SEARCH 가 오류 → 누락 그대로)
        bund = f"${fc['통합발주_포함공종']}{r}"
        real = "&".join(f'IF(AND(COUNTIFS({id_rng},$A{r},{tr_rng},"{tr}")=0,ISERROR(SEARCH("{tr}",{bund}))),"{tr} ","")' for tr in ("건축", "전기", "정보통신", "소방"))
        ws[f"{fc[MISS_REAL_COL]}{r}"] = f'=IF(TRIM({real})="","",TRIM({real})&" 누락")'
        # 실제 누락 공종의 수 = 위 칸의 공종 낱말 수('정보통신 소방 누락' → 빈칸 2개 → 2). 빈칸이면 0. 05-요약·05-k 시트가 이 수로 나눈다
        mr = f"${fc[MISS_REAL_COL]}{r}"
        ws[f"{fc[MISS_CNT_COL]}{r}"] = f'=IF({mr}="",0,LEN({mr})-LEN(SUBSTITUTE({mr}," ","")))'
        # 04b 참고 배분: 프로젝트·세부공종별 배분금액 합(총공사비합계에는 더하지 않음)
        # 다른 공종 공고 안에 묶인 금액만 합산(상위공종<>세부공종) — 해당 공종 공고 자신의 몫은 이미 그 공종(수식)에 전액 들어 있으므로 제외
        for tr in DETAIL_ALLOC_TRADES:
            ws[f"{fc[f'문서배분_{tr}(수식)']}{r}"] = (f'=SUMIFS({D["배분금액(수식)"]},{D["프로젝트ID"]},$A{r},{D["세부공종"]},"{tr}",'
                                                  f'{D["상위공종"]},"<>{tr}")')
        arch_all = f'SUMIFS({D["배분금액(수식)"]},{D["프로젝트ID"]},$A{r},{D["상위공종"]},"건축")'
        arch_own = f'SUMIFS({D["배분금액(수식)"]},{D["프로젝트ID"]},$A{r},{D["상위공종"]},"건축",{D["세부공종"]},"건축")'
        ws[f"{fc['건축_순건축추정(수식)']}{r}"] = f'=IF({arch_all}=0,"",{fc["건축(수식)"]}{r}-({arch_all}-{arch_own}))'
        # 중복경고: 건축 공고의 집계표에 전기·정보통신·소방 금액이 있고 그 공종의 분리발주 공고도 있을 때(06 '공종 중복 가능(문서)' 와 같은 기준)
        dup = "&".join(f'IF(AND(SUMIFS({D["배분금액(수식)"]},{D["프로젝트ID"]},$A{r},{D["상위공종"]},"건축",{D["세부공종"]},"{tr}")>0,'
                       f'N({fc[f"{tr}(수식)"]}{r})>0),"{tr} ","")' for tr in ("전기", "정보통신", "소방"))
        ws[f"{fc['문서배분_중복경고(수식)']}{r}"] = f'=TRIM(IF(TRIM({dup})="","",TRIM({dup})&" 중복가능"))'
        if servc is not None:
            # 용역 참고 열: 08 에서 이 프로젝트·그 종류·합산여부 Y 인 줄의 추정가격 합. 합산여부 Y 인 공고가 없으면 0 이 아니라 빈칸(공고 없이 집행되는 용역이 많다)
            for k in SERVC_KINDS:
                cond = f'{V["프로젝트ID"]},$A{r},{V["용역종류"]},"{k}",{V["합산여부"]},"Y"'
                ws[f"{fc[f'용역_{k}_추정가격(수식)']}{r}"] = f'=IF(COUNTIFS({cond})=0,"",SUMIFS({V["추정가격"]},{cond}))'
            ws[f"{fc['용역_공고수(수식)']}{r}"] = f'=COUNTIFS({V["프로젝트ID"]},$A{r},{V["합산여부"]},"Y")'
            # 발주완료(보조지표): 이 프로젝트에 더한(Y) 폐기물처리 용역이 하나라도 있으면 '발주완료'. 없으면 빈칸('미완료' 가 아니다 — 폐기물이 적은 공사·
            # 수의계약은 용역 공고가 없다). 08 의 노란 칸을 고치면 따라 바뀐다. 실자료(2026-11): 폐기물 용역 첫 공고는 공사 첫 공고의 3주쯤 뒤(가운데값)이고,
            # 3건 중 1건꼴로 그 뒤에 전기·정보통신·소방(드물게 건축) 공고가 처음 나왔다 → 공고가 최근인 사업에서는 이른 표시일 수 있다.
            # 그래서 그 폐기물 용역(Y) 줄이 모두 공사 공고 수집 종료일에서 SERVC_RECENT_DAYS 일 안쪽이거나 그 뒤면(= 가장 이른 공고도 최근) '발주완료(최근)' 으로 따로 적는다.
            # 08 의 일수 칸이 빈 줄(종료일 모름·끼워 넣은 줄)은 '<183' 조건에 세어지지 않으므로 '최근' 으로 치지 않는다
            done = f'{V["프로젝트ID"]},$A{r},{V["용역종류"]},"{SERVC_DONE_KIND}",{V["합산여부"]},"Y"'
            ws[f"{fc[SERVC_DONE_COL]}{r}"] = (f'=IF(COUNTIFS({done})=0,"",IF(COUNTIFS({done},{V[SERVC_AGE_COL]},"<{SERVC_RECENT_DAYS}")=COUNTIFS({done}),'
                                              f'"{SERVC_DONE_RECENT}","발주완료"))')
        for c in fcols:
            if "(수식)" in c:
                ws[f"{fc[c]}{r}"].font = FONT
    n = ws.max_row + 2
    ws[f"A{n}"] = ("주: 건설공사비지수 보정계수는 사용자 입력(기준연도 지수/공고연도 지수). 공종별 열과 '총공사비합계_기초금액기준' = 04 금액합계_기초금액기준(기초금액 + 관급 + 기타금액)의 합, "
                   "'총공사비합계_추정가격기준' = 04 금액합계_추정가격기준(추정가격 + 부가세 + 관급 + 기타금액)의 합 — 두 대표값을 함께 제시한다. 건축공종계 = 건축 + 기계설비 + 토목 + 조경. "
                   "연면적은 01_시설마스터 값을 참조(01 에서 수정). 공종누락경고: '누락'=해당 공종 공고 없음, '금액없음'=공고는 있으나 금액 0. "
                   "공종누락_통합발주제외 = 공종누락경고의 '누락' 가운데 통합발주_포함공종에 적힌 공종(다른 공고에 묶여 발주된 공종)을 뺀 것 — 공고도 없고 묶였다는 기록도 없는 공종만 남는다"
                   "(빈칸 = 그런 공종 없음. 묶인 공종의 금액은 묶은 공고의 공종 열에 들어 있고 따로 나뉘지 않는다 — 04b 에 내역이 있으면 문서배분_*. "
                   "통합발주_포함공종은 나라장터의 부공종과 공고문에서 읽은 것이라, 거기 적히지 않은 통합발주는 그대로 누락으로 남는다). "
                   f"실제누락_공종수 = 공종누락_통합발주제외에 남은 공종의 수(0~4). 이 수별 프로젝트 목록은 '{S_MISS_N.format(k=0)}' ~ '{S_MISS_N.format(k=len(MISS_TRADES))}' 시트, 유형별 건수는 '{S_MISS_SUM}' 시트. "
                   "세부사업명·세부사업수: 한 시설 후보(시설ID×사업유형)의 공고가 여러 사업(다른 단지·구역·연도·별개 공사)이면 세부사업으로 나눠 행을 따로 둔다"
                   "(프로젝트ID 끝 번호 -01, -02 …, 세부사업수 = 그 시설 후보가 나뉜 개수). 같은 세부사업·같은 공종의 공고가 여럿이면 모두 합산(06 '같은 공종 합산'). "
                   "시설간_중복: 같은 공고가 여러 시설 후보에 걸려 같은 사업이 두 행 이상으로 집계된 경우 — '같은 사업: X'(대표 공고가 모두 같음), 'X 의 일부'(이 행의 공고가 모두 X 에 있음), "
                   "'X 를 포함(n/m건)', 'X 와 일부 겹침(n/m건)', 'X + Y 에 나뉘어 있음'(이 행의 공고가 다른 시설 후보의 세부사업 X, Y 에 나뉘어 모두 있음). 중복_제외권장 = Y 인 행을 빼면 같은 공고가 두 번 합산되지 않는다"
                   "(같은 사업끼리는 세부사업으로 나뉘지 않은 시설의 행을 남김. '일부 겹침' 은 그것만으로는 Y 가 붙지 않음 — Y 행을 뺀 뒤에도 두 번 남는 공고 수는 06_검증로그 '시설 간 중복 프로젝트' 에 표시). 여러 프로젝트의 합계·평균을 낼 때는 Y 행을 필터로 제외할 것. "
                   "㎡당 공사비는 사업유형(신축/증축/리모델링)이 같은 프로젝트끼리만 비교할 것 — 리모델링의 연면적은 공사 대상 연면적임. "
                   + ("용역_*_추정가격(수식)은 08 시트에서 이 프로젝트에 이어진 용역 공고(합산여부 Y)의 추정가격 합(부가세 제외) — 공사비가 아닌 참고값이며 총공사비합계·㎡당 공사비에 들어가지 않는다. "
                      "빈칸은 '더한(합산여부 Y) 공고 없음' 이지 용역비 0 이 아니다(README '용역(08·05)'). "
                      "발주완료_폐기물용역(수식) = 08 에 이 프로젝트의 폐기물처리 용역 공고(합산여부 Y)가 하나라도 있으면 '발주완료' — 그 사업의 공사 발주가 끝났다고 보는 보조지표다. "
                      "빈칸은 '판단할 근거 없음' 이지 '미완료' 가 아니다(폐기물이 적은 공사·수의계약은 용역 공고가 없다). "
                      "폐기물 용역 공고는 대개 건축 공고 무렵에 나고, 그 뒤에 전기·정보통신·소방(드물게 건축) 공고가 처음 나오는 사업이 3건 중 1건꼴이다(넷 중 셋은 석 달 안, 열에 하나는 일곱 달 넘게 뒤). "
                      + (f"그래서 이 프로젝트의 폐기물 용역 공고(합산여부 Y)가 모두 공사 공고 수집 종료일({cn_end_s})에서 6개월({SERVC_RECENT_DAYS}일) 안쪽이거나 그 뒤면 '{SERVC_DONE_RECENT}' 으로 따로 적는다"
                         f"(08 의 '{SERVC_AGE_COL}' 가 모두 {SERVC_RECENT_DAYS} 미만. {cn_end_def}). '{SERVC_DONE_RECENT}' 은 아직 발주 중일 수 있으니 발주가 끝난 사업을 고를 때는 '발주완료' 만 고르고, "
                         "쓰기 전에 08 의 공고일시를 확인할 것(용역 공고는 공사 공고보다 늦은 달까지 수집된다). "
                         "'발주완료' 인 행도 그 사업의 공사 공고가 최근이면(01 의 '최종공고일' 이 공사 공고 수집 종료일에서 6개월 안쪽) 아직 발주 중일 수 있다. "
                         if cn_end_s else
                         "폐기물 용역 공고가 최근인 사업은 아직 발주 중일 수 있으니 08 의 공고일시를 확인할 것(용역 공고는 공사 공고보다 늦은 달까지 수집된다. "
                         f"이 워크북은 실린 공사 공고에서 공고일을 읽지 못해 '{SERVC_DONE_RECENT}' 을 가리지 않았다). ") +
                      "'발주완료' 는 이 행에 공사비가 모두 들어 있다는 뜻이 아니다 — 공종누락경고는 따로 본다(통합발주이거나 그 공종 공고가 이 DB 에 없는 경우. 철거·부지 정리 공고만 있는 행도 '발주완료' 가 될 수 있다). "
                      "같은 사업이 두 행으로 잡힌 경우(시설간_중복) 두 행이 함께 '발주완료' 가 될 수 있다. " if servc is not None else "") +
                   "문서배분_*(수식)은 다른 공종 공고 안에 묶인 그 공종의 참고 배분(04b 비율 × 04 금액합계_기초금액기준, 자기 공종 공고의 몫 제외, 총공사비합계에 미포함), 건축_순건축추정 = 건축(수식) − 건축 공고 안의 비건축 배분(04b 배분적용이 Y 인 건축 공고의 것만 뺀다 — 배분되지 않은 건축 공고가 섞인 프로젝트는 그만큼 덜 빠진 값).")
    ws[f"A{n}"].font = FONT

    # ── 05-요약 / 05-0 … 05-4 — 실제 누락 공종 수(05 '실제누락_공종수')별 프로젝트 목록과 유형별 요약 ──
    # 값은 모두 05 를 읽는 수식이다(01·04 를 고치면 따라 바뀐다). 어느 프로젝트가 어느 시트에 오는지만 만들 때 정한다 — 05 의 수식과 같은 조건
    # (04 에 그 프로젝트·그 공종 줄이 없고, 통합발주_포함공종 글에도 그 공종 이름이 없음)을 여기서 한 번 계산한다.
    n5 = max(len(f5), 1) + 1
    R5 = lambda c: f"'{S_FAC}'!${fc[c]}$2:${fc[c]}${n5}"
    txt5 = lambda c: ["" if _clean(v) is None else str(_clean(v)) for v in f5[c]]
    have: Dict[str, set] = {}
    for pid, tr in zip(t["프로젝트ID"], t["공종"]):
        have.setdefault(str(pid), set()).add(str(tr))
    pids5 = txt5("프로젝트ID")
    nmiss = [sum(1 for tr in MISS_TRADES if tr not in have.get(p, set()) and tr not in b) for p, b in zip(pids5, txt5("통합발주_포함공종"))]
    ks = list(range(len(MISS_TRADES) + 1))
    # 유형별 요약: 왼쪽 = 전체, 오른쪽 = 중복_제외권장 Y 를 뺀 것. 한 줄 = 유형, 칸 = COUNTIFS(05)
    ws = wb.create_sheet(S_MISS_SUM)
    crit = lambda v: '"=' + re.sub(r"([~*?])", r"~\1", v).replace('"', '""') + '"'     # COUNTIFS 조건 글자('=' 를 붙여 글자 그대로 견준다. 빈 값은 "=" → 빈 칸)
    typ, big, mid = txt5("사업유형"), txt5("표2_대분류"), txt5("표2_중분류")
    def _order(vals, first=()):                      # 건수가 많은 순(같으면 이름순). first 에 적힌 값은 그 순서로 앞에
        cnt: Dict[str, int] = {}
        for v in vals:
            cnt[v] = cnt.get(v, 0) + 1
        return [v for v in first if v in cnt] + sorted((v for v in cnt if v not in first), key=lambda v: (-cnt[v], v))
    o_typ, o_big = _order(typ, ("신축", "증축", "리모델링", "증축·리모델링")), _order(big)
    pairs_mid, pairs_typ = set(zip(big, mid)), set(zip(big, typ))
    blocks = [("1. 사업유형별", [((v, ""), [("사업유형", v)]) for v in o_typ]),
              ("2. 시설 유형별(표2_대분류)", [((v, ""), [("표2_대분류", v)]) for v in o_big]),
              ("3. 시설 유형별(표2_대분류 · 표2_중분류)", [((b, m), [("표2_대분류", b), ("표2_중분류", m)]) for b in o_big
                                                 for m in _order([m for bb, m in zip(big, mid) if bb == b]) if (b, m) in pairs_mid]),
              ("4. 시설 유형 · 사업유형별", [((b, v), [("표2_대분류", b), ("사업유형", v)]) for b in o_big for v in o_typ if (b, v) in pairs_typ])]
    sum_head = ["사업 수"] + [f"실제 누락 {k}개" for k in ks] + ["누락 없음 비율"] + [f"{tr} 실제 누락" for tr in MISS_TRADES]
    c_all, c_nodup = 3, 3 + len(sum_head) + 1          # 두 표의 첫 열(C, 그리고 한 칸 띄운 뒤)
    ws.append([f"실제 누락 공종 수 — 유형별 프로젝트 수 (05 '{MISS_CNT_COL}')"])
    ws["A1"].font = Font(name="Arial", size=12, bold=True)
    ws.append(["실제 누락 = 건축·전기·정보통신·소방 가운데 그 공종 공고가 없고, 통합발주로 다른 공고에 묶였다는 기록(05 '통합발주_포함공종')도 없는 공종. "
               "모든 칸은 05 시트를 세는 수식이다. 오른쪽 표는 같은 사업이 두 번 잡힌 행(05 '중복_제외권장' = Y)을 뺀 것."])
    ws["A2"].font = FONT

    def _sum_row(label, conds):
        ws.append([x or None for x in label])
        r = ws.max_row
        for c0, dup in ((c_all, ""), (c_nodup, f",{R5('중복_제외권장')},\"<>Y\"")):
            base = "".join(f"{R5(c)},{crit(v)}," for c, v in conds)
            L = [get_column_letter(c0 + i) for i in range(len(sum_head))]
            ws[f"{L[0]}{r}"] = f'=COUNTIFS({base}{R5(MISS_CNT_COL)},">=0"{dup})'
            for k in ks:
                ws[f"{L[1 + k]}{r}"] = f"=COUNTIFS({base}{R5(MISS_CNT_COL)},{k}{dup})"
            ws[f"{L[2 + len(ks) - 1]}{r}"] = f'=IF({L[0]}{r}=0,"",{L[1]}{r}/{L[0]}{r})'
            ws[f"{L[2 + len(ks) - 1]}{r}"].number_format = "0.0%"
            for i, tr in enumerate(MISS_TRADES):
                ws[f"{L[2 + len(ks) + i]}{r}"] = f'=COUNTIFS({base}{R5(MISS_REAL_COL)},"*{tr}*"{dup})'
        for c in range(1, c_nodup + len(sum_head)):
            ws.cell(row=r, column=c).font = BOLD if label[0] == "계" else FONT

    for title, rows in blocks:
        ws.append([])
        ws.append([title] + [None] * (c_all - 2) + ["전체"] + [None] * len(sum_head) + ["중복_제외권장 Y 제외"])
        ws.append(["구분", "세부 구분"] + sum_head + [None] + sum_head)
        for r in (ws.max_row - 1, ws.max_row):
            for c in range(1, c_nodup + len(sum_head)):
                cell = ws.cell(row=r, column=c)
                cell.font = BOLD
                if r == ws.max_row and cell.value:
                    cell.fill, cell.alignment = HEAD, Alignment(wrap_text=True, vertical="center")
        for label, conds in rows:
            _sum_row(tuple("(빈칸)" if (x == "" and i < len(conds)) else x for i, x in enumerate(label)), conds)
        _sum_row(("계", ""), [])
    ws.append([])
    ws.append(["주: '실제 누락 k개' = 실제 누락 공종이 k개인 프로젝트 수(프로젝트 목록은 "
               f"'{S_MISS_N.format(k=0)}' ~ '{S_MISS_N.format(k=ks[-1])}' 시트). '○○ 실제 누락' = 그 공종이 실제 누락인 프로젝트 수(한 프로젝트가 여러 공종에 세어진다). "
               "누락 없음 비율 = 실제 누락 0개 ÷ 사업 수. 통합발주 기록은 나라장터의 부공종과 공고문에서 읽은 것이라, 기록되지 않은 통합발주는 누락으로 세어진다. "
               "누락 공종이 있다고 그 공사가 없었다는 뜻은 아니다 — 수집 기간 밖에 공고됐거나, 소액이라 공고 없이 집행됐거나, 아직 발주 전일 수 있다."])
    ws.cell(row=ws.max_row, column=1).font = FONT
    ws.column_dimensions["A"].width, ws.column_dimensions["B"].width = 26, 18
    for c in range(c_all, c_nodup + len(sum_head)):
        ws.column_dimensions[get_column_letter(c)].width = 4 if c == c_nodup - 1 else 12
    ws.freeze_panes = "C3"
    # 누락 수별 프로젝트 목록: 프로젝트ID(값) + 05 의 그 프로젝트 행을 읽는 수식
    mcols = ["프로젝트ID", "시설명", "세부사업명", "사업유형", "표2_대분류", "표2_중분류", "수요기관", "공고연도", "연면적_m2", "지상층수", "지하층수", "구조"] \
        + [f"{tr}(수식)" for tr in TRADES] + ["건축공종계(수식)"] + TOTAL5_COLS + ["공종누락경고(수식)", "통합발주_포함공종", MISS_REAL_COL, MISS_CNT_COL] \
        + ([SERVC_DONE_COL] if servc is not None else []) + ["중복_제외권장", "공고건수_최신", MISS_ROW_COL]
    mc = {c: get_column_letter(i + 1) for i, c in enumerate(mcols)}
    for k in ks:
        ws = wb.create_sheet(S_MISS_N.format(k=k))
        ids = [p for p, m in zip(pids5, nmiss) if m == k]
        _write_df(ws, pd.DataFrame({c: (ids if c == "프로젝트ID" else [None] * len(ids)) for c in mcols}),
                  money_cols=["연면적_m2"] + [f"{tr}(수식)" for tr in TRADES] + ["건축공종계(수식)"] + TOTAL5_COLS, blue_cols=["프로젝트ID"],
                  widths={"시설명": 30, "세부사업명": 36, "수요기관": 24, "공종누락경고(수식)": 28, "통합발주_포함공종": 34, MISS_REAL_COL: 28, MISS_CNT_COL: 12, SERVC_DONE_COL: 16})
        for r in range(2, len(ids) + 2):
            H = f"${mc[MISS_ROW_COL]}{r}"
            ws[f"{mc[MISS_ROW_COL]}{r}"] = f"=MATCH($A{r},{R5('프로젝트ID')},0)"
            for c in mcols[1:-1]:
                ws[f"{mc[c]}{r}"] = f'=IF(INDEX({R5(c)},{H})="","",INDEX({R5(c)},{H}))'
        nk = len(ids) + 3
        ws[f"A{nk}"] = (f"주: 실제 누락 공종이 {k}개인 프로젝트(05 '{MISS_CNT_COL}' = {k}). 실제 누락 = 건축·전기·정보통신·소방 가운데 그 공종 공고가 없고 통합발주_포함공종에도 적히지 않은 공종. "
                        "프로젝트ID 를 뺀 칸은 모두 05 시트의 그 프로젝트 행을 읽는 수식이다(공종별 금액·총공사비합계는 05 와 같은 값, 연면적은 01 의 값) — 값을 고치려면 01·04 에서 고칠 것. "
                        f"어느 프로젝트가 이 시트에 오는지는 워크북을 만들 때 정해진다: 04 를 고쳐 누락 수가 달라지면 '{MISS_CNT_COL}' 열이 {k} 이 아닌 값으로 보이고, excel 단계를 다시 실행하면 맞는 시트로 옮겨진다. "
                        + ("누락 0개에는 통합발주로 묶인 공종이 있는 프로젝트도 들어 있다 — 그 공종의 금액은 묶은 공고의 공종 열(주로 건축)에 들어 있어 그 공종 열은 0 으로 보인다. " if k == 0 else
                           "이 시트의 총공사비합계·㎡당 공사비에는 누락된 공종의 금액이 들어 있지 않다. ") +
                        "합계·평균을 낼 때는 중복_제외권장 Y 행을 빼고, ㎡당 공사비는 사업유형이 같은 프로젝트끼리 비교할 것.")
        ws[f"A{nk}"].font = FONT

    # ── 06 / 07 ──
    ws = wb.create_sheet("06_검증로그")
    lcols = ["공고번호", "항목", "API값", "문서값", "차이", "차이율", "판정", "비고"]
    lg = logs.copy() if not logs.empty else pd.DataFrame(columns=lcols)
    for c in lcols:
        if c not in lg.columns:
            lg[c] = None
    _write_df(ws, lg[lcols], money_cols=["API값", "문서값", "차이"], blue_cols=lcols, widths={"비고": 70, "항목": 26})
    for r in range(2, ws.max_row + 1):
        ws[f"F{r}"].number_format = "0.00%"
    ws = wb.create_sheet("07_추출노트")
    ncols = ["공고번호", "파일명", "형식", "우선순위", "다운로드", "추출성공", "추출글자수", "파서", "LLM호출", "오류", "URL"]
    nt = notes.copy() if not notes.empty else pd.DataFrame(columns=ncols)
    for c in ncols:
        if c not in nt.columns:
            nt[c] = None
    _write_df(ws, nt[ncols], blue_cols=ncols, widths={"파일명": 45, "오류": 50, "URL": 50})

    # ── 08 용역공고(참고) — link-servc 결과가 있을 때만. 05 의 용역 참고 열이 이 시트를 SUMIFS 로 읽는다 ──
    if sv is not None:
        ws = wb.create_sheet(S_SERVC)
        _write_df(ws, sv, money_cols=["추정가격", "배정예산"], blue_cols=[c for c in SERVC8_COLS if c != "합산여부"], yellow_cols=["합산여부", "프로젝트ID"],
                  widths={"시설명": 26, "공고명": 52, "합산제외_까닭": 44, "연결근거": 50, "용역세부": 22, "수요기관": 24, "공고기관": 20, "프로젝트_후보": 24, "상세URL": 30})
        rc = get_column_letter(SERVC8_COLS.index("공사비대비_비율") + 1)
        for r in range(2, ws.max_row + 1):
            ws[f"{rc}{r}"].number_format = "0.0%"
        nv = len(sv) + 3
        ws[f"A{nv}"] = ("주: 용역 입찰공고(나라장터) 가운데 이 DB 의 시설에 이어진 공고. 공사비가 아닌 참고 자료다. 추정가격은 부가세 제외, 배정예산은 보통 부가세 포함. "
                        "합산여부 Y 인 줄만 05 의 '용역_○○_추정가격' 에 더해진다 — N 의 까닭은 '합산제외_까닭'. 프로젝트ID 가 빈 줄은 시설까지만 이어진 공고(어느 프로젝트인지 가리지 못함 — '프로젝트_후보'). "
                        "손으로 고칠 때: 05 는 프로젝트ID·용역종류·합산여부 세 칸을 함께 본다 — 프로젝트ID 가 빈 줄은 합산여부를 Y 로 바꾸는 것만으로는 더해지지 않고 프로젝트ID(노란 칸)도 적어야 하며, "
                        "용역종류가 '기타' 인 줄은 Y 로 바꿔도 금액 열에 들어가지 않는다(용역_공고수 에는 세어진다). "
                        "용역종류가 '폐기물처리' 이고 합산여부가 Y 인 줄이 있는 프로젝트는 05 의 '발주완료_폐기물용역' 이 '발주완료' 가 된다"
                        + (f"(그 줄들의 '{SERVC_AGE_COL}' 가 모두 {SERVC_RECENT_DAYS} 미만이면 '{SERVC_DONE_RECENT}'). "
                           f"{SERVC_AGE_COL} = 공사 공고 수집 종료일({cn_end_s}) − 용역 공고일(날수. {cn_end_def}. 음수 = 그날 뒤에 난 용역 공고. "
                           "이 칸이 빈 줄은 '최근' 으로 치지 않으므로 폐기물처리 줄을 끼워 넣었으면 이 칸도 적을 것). "
                           if cn_end_s else f". {SERVC_AGE_COL} 는 실린 공사 공고에서 공고일을 읽지 못해 비워 두었다('{SERVC_DONE_RECENT}' 을 가리지 않는다). ") +
                        "표 아래에 새로 적은 줄은 더해지지 않는다(표의 둘째 줄부터 마지막 줄 사이에 행을 끼워 넣을 것 — 머리글 바로 아래 첫 줄 위에 끼운 줄도 더해지지 않는다). "
                        "상태: 대표 / 재공고 이전(같은 용역이 뒤에 다시 공고됨) / 취소. 공사대비_개월 = 용역 공고일 − 그 프로젝트(없으면 가장 가까운 프로젝트) 공사 첫 공고일. "
                        "공사비대비_비율 = 추정가격 ÷ 그 프로젝트 공사 추정가격 합(지나치게 크면 06_검증로그 '용역비 비율'). "
                        "연결은 공고명과 발주 기관만으로 한 것이므로 금액을 쓰기 전에 공고명을 확인할 것(상세URL).")
        ws[f"A{nv}"].font = FONT

    wb.save(path)
