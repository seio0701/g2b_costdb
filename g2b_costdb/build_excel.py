"""S7 Excel DB 생성 (openpyxl, 수식 기반).

시트: README / 01_시설마스터 / 02_공고목록_전체 / 03_공고목록_최신 / 04_공사비DB_공종별 / 05_공사비DB_시설합산 / 06_검증로그 / 07_추출노트
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
                   trade_detail: Optional[pd.DataFrame] = None) -> None:
    wb = Workbook()
    # ── README ──
    ws = wb.active
    ws.title = "README"
    meta = meta or {}
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
        ["05_공사비DB_시설합산", "시설 단위 공종 합산(SUMIFS, 공종별 열은 기초금액 기준), 건축공종계, 총공사비합계 2종(기초금액 기준·추정가격 기준)과 각각의 ㎡당 공사비, 보정계수 입력, 공종 누락 경고. 행 단위 = 프로젝트(시설ID×사업유형, 여러 사업이 묶인 시설 후보는 세부사업별로 나눈 행 — 세부사업명·세부사업수). 문서배분_*(수식) = 다른 공종 공고 안에 묶인 그 공종의 배분금액 합(해당 공종 공고 자신의 몫은 제외, 참고), 건축_순건축추정 = 건축(수식) − 건축 공고 안의 비건축 배분, 문서배분_중복경고 = 건축 집계표에 전기·정보통신·소방 금액이 있는데 분리발주 공고도 있음. 시설간_중복 = 다른 시설 후보의 프로젝트와 대표 공고가 겹치는 경우 그 관계, 중복_제외권장 = Y 인 행을 빼면 같은 공고가 두 번 합산되지 않는다(합계·평균을 낼 때는 Y 행을 필터로 제외. '일부 겹침' 은 그것만으로는 Y 가 붙지 않는다 — Y 행을 뺀 뒤에도 두 번 남는 공고가 있는지는 06_검증로그 '시설 간 중복 프로젝트' 의 남은 건수로 확인하고, 남아 있을 때만 검수에서 판단)"],
        ["06_검증로그", "API-문서 교차검증, 재발주 금액변동, 정합성 점검 결과"],
        ["07_추출노트", "첨부파일별 다운로드·텍스트 추출 결과"],
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
        + ["공종누락경고(수식)", "통합발주_포함공종"] \
        + [f"문서배분_{tr}(수식)" for tr in DETAIL_ALLOC_TRADES] + ["건축_순건축추정(수식)", "문서배분_중복경고(수식)", "시설간_중복", "중복_제외권장", "공고건수_최신", "메모"]
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
              + [f"문서배분_{tr}(수식)" for tr in DETAIL_ALLOC_TRADES],
              blue_cols=["프로젝트ID", "시설ID", "시설명", "세부사업명", "세부사업수", "사업유형", "표2_대분류", "표2_중분류", "수요기관", "연면적_m2", "지상층수", "지하층수", "구조",
                         "공고연도", "공고건수_최신", "통합발주_포함공종", "시설간_중복", "중복_제외권장"],
              yellow_cols=["건설공사비지수_보정계수(입력)", "메모"],
              widths={"시설명": 30, "세부사업명": 36, "공종누락경고(수식)": 28, "통합발주_포함공종": 34, "시설간_중복": 34})
    fc = {c: get_column_letter(i + 1) for i, c in enumerate(fcols)}
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
        for c in fcols:
            if "(수식)" in c:
                ws[f"{fc[c]}{r}"].font = FONT
    n = ws.max_row + 2
    ws[f"A{n}"] = ("주: 건설공사비지수 보정계수는 사용자 입력(기준연도 지수/공고연도 지수). 공종별 열과 '총공사비합계_기초금액기준' = 04 금액합계_기초금액기준(기초금액 + 관급 + 기타금액)의 합, "
                   "'총공사비합계_추정가격기준' = 04 금액합계_추정가격기준(추정가격 + 부가세 + 관급 + 기타금액)의 합 — 두 대표값을 함께 제시한다. 건축공종계 = 건축 + 기계설비 + 토목 + 조경. "
                   "연면적은 01_시설마스터 값을 참조(01 에서 수정). 공종누락경고: '누락'=해당 공종 공고 없음, '금액없음'=공고는 있으나 금액 0. "
                   "세부사업명·세부사업수: 한 시설 후보(시설ID×사업유형)의 공고가 여러 사업(다른 단지·구역·연도·별개 공사)이면 세부사업으로 나눠 행을 따로 둔다"
                   "(프로젝트ID 끝 번호 -01, -02 …, 세부사업수 = 그 시설 후보가 나뉜 개수). 같은 세부사업·같은 공종의 공고가 여럿이면 모두 합산(06 '같은 공종 합산'). "
                   "시설간_중복: 같은 공고가 여러 시설 후보에 걸려 같은 사업이 두 행 이상으로 집계된 경우 — '같은 사업: X'(대표 공고가 모두 같음), 'X 의 일부'(이 행의 공고가 모두 X 에 있음), "
                   "'X 를 포함(n/m건)', 'X 와 일부 겹침(n/m건)', 'X + Y 에 나뉘어 있음'(이 행의 공고가 다른 시설 후보의 세부사업 X, Y 에 나뉘어 모두 있음). 중복_제외권장 = Y 인 행을 빼면 같은 공고가 두 번 합산되지 않는다"
                   "(같은 사업끼리는 세부사업으로 나뉘지 않은 시설의 행을 남김. '일부 겹침' 은 그것만으로는 Y 가 붙지 않음 — Y 행을 뺀 뒤에도 두 번 남는 공고 수는 06_검증로그 '시설 간 중복 프로젝트' 에 표시). 여러 프로젝트의 합계·평균을 낼 때는 Y 행을 필터로 제외할 것. "
                   "㎡당 공사비는 사업유형(신축/증축/리모델링)이 같은 프로젝트끼리만 비교할 것 — 리모델링의 연면적은 공사 대상 연면적임. "
                   "문서배분_*(수식)은 다른 공종 공고 안에 묶인 그 공종의 참고 배분(04b 비율 × 04 금액합계_기초금액기준, 자기 공종 공고의 몫 제외, 총공사비합계에 미포함), 건축_순건축추정 = 건축(수식) − 건축 공고 안의 비건축 배분(04b 배분적용이 Y 인 건축 공고의 것만 뺀다 — 배분되지 않은 건축 공고가 섞인 프로젝트는 그만큼 덜 빠진 값).")
    ws[f"A{n}"].font = FONT

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

    wb.save(path)
