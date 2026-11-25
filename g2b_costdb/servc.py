"""3단계(b): 용역 입찰공고(설계·건설사업관리/감리·폐기물처리)를 공사비 DB 의 시설·프로젝트에 연결한다 — 참고용(총공사비 합계에는 넣지 않는다).

입력  data/servc_all.parquet          collect --servc 가 남긴 용역 공고 원문(보존 규칙에 걸린 것)
      data/notices_hist.parquet       dedup 결과(시설·프로젝트·공고 이력)
      data/notices_latest.parquet     dedup 결과(프로젝트별 대표 공고)
      output/facility_candidates.xlsx 검수 파일(시설 이름·별칭·수요기관)
출력  data/servc_linked.parquet       연결된 용역 공고(한 줄 = 용역 공고 × 시설) → excel 의 '08_용역공고(참고)' 시트와 05 의 용역 참고 열

API·LLM 호출 없음. 연결 조건은 '시설 이름이 용역 공고명에 있고, 발주 기관이 그 시설의 공사 공고와 같은 기관' 이다.
어느 프로젝트의 용역인지 가릴 수 없으면 시설까지만 잇고 프로젝트 칸은 비워 둔다(05 의 합에 들어가지 않는다).
규칙의 원칙: 틀린 금액을 05 에 보이느니 빈칸으로 둔다 — 가릴 수 없는 공고는 08 에 까닭과 함께 N 으로 싣고, 사람이 공고명을 보고 합산여부·프로젝트ID 를 고칠 수 있다."""
from __future__ import annotations

import logging
import re
from functools import lru_cache
from typing import Dict, List, Optional, Tuple

import pandas as pd

from . import works
from .classify import classify_notice_kind, parse_amount
from .dedup import _no, _parse_dt

log = logging.getLogger(__name__)

LINK_COLS = ["프로젝트ID", "시설ID", "시설명", "용역종류", "용역세부", "합산여부", "합산제외_까닭", "연결강도", "연결근거", "공고번호", "공고차수", "공고명", "공고일시",
             "공고기관", "수요기관", "추정가격", "배정예산", "상태", "공고종류", "용역구분", "조달분류", "프로젝트_후보", "공사_최초공고일", "공사대비_개월",
             "공사비대비_비율", "상세URL"]
KINDS = ("설계", "건설사업관리·감리", "폐기물처리")          # 05 에 열을 두는 종류. 그 밖('기타')은 08 에만 싣는다


def _txt(v) -> str:
    if v is None or (isinstance(v, float) and v != v):
        return ""
    t = str(v).strip()
    return "" if t.lower() in ("nan", "none", "<na>", "nat") else t


def _ns(v) -> str:
    return re.sub(r"\s+", "", _txt(v))


# ── 1. 표준화 ────────────────────────────────────────────────────────────────────────────────────────────
_RAW_KEYS = ("bid_no", "bid_ord", "name", "kind", "re_notice", "reg_type", "notice_dt", "notice_inst", "demand_inst", "demand_inst_cd", "presmpt_price",
             "servc_budget", "servc_div", "servc_class", "servc_class_mid", "servc_class_lrg", "prev_bid_no", "detail_url")


def raw_columns(cfg: dict) -> List[str]:
    """servc_all.parquet 에서 읽을 원문 열 이름(config fields 매핑 기준)."""
    f = cfg["fields"]
    return list(dict.fromkeys(f[k] for k in _RAW_KEYS if f.get(k)))


def standardize(raw: pd.DataFrame, cfg: dict) -> pd.DataFrame:
    """용역 공고 원문 → 표준 열. 없는 필드는 빈 값."""
    f = cfg["fields"]
    empty = pd.Series([""] * len(raw), index=raw.index, dtype=object)
    g = lambda k: raw[f[k]].map(_txt) if (k in f and f[k] in raw.columns) else empty
    d = pd.DataFrame({
        "공고번호": g("bid_no").map(_no), "공고차수": g("bid_ord"), "공고명": g("name"), "공고종류_원문": g("kind"), "재공고여부": g("re_notice"),
        "등록유형": g("reg_type"), "공고일시": g("notice_dt"), "공고기관": g("notice_inst"), "수요기관": g("demand_inst"),
        "추정가격": g("presmpt_price").map(parse_amount), "배정예산": g("servc_budget").map(parse_amount),
        "용역구분": g("servc_div"), "조달분류": g("servc_class"), "조달분류_중": g("servc_class_mid"), "조달분류_대": g("servc_class_lrg"),
        "이전공고번호": g("prev_bid_no").map(_no), "상세URL": g("detail_url"), "수요기관코드": g("demand_inst_cd").map(_no),
    })
    d["공고차수"] = d["공고차수"].map(lambda s: s.zfill(3) if s.isdigit() else (s or "000"))
    d["공고종류"] = [classify_notice_kind(k, r, t, n) for k, r, t, n in zip(d["공고종류_원문"], d["재공고여부"], d["등록유형"], d["공고명"])]
    return d[d["공고번호"] != ""].reset_index(drop=True)


def latest_orders(d: pd.DataFrame) -> pd.DataFrame:
    """공고번호마다 가장 늦은 차수 한 줄(변경공고의 금액·취소 여부는 마지막 차수가 맞다). 구간 경계에서 두 번 받은 같은 줄도 여기서 하나가 된다."""
    if d.empty:
        return d.assign(_dt=pd.Series(dtype="datetime64[ns]"))
    d = d.assign(_ord=pd.to_numeric(d["공고차수"], errors="coerce").fillna(0), _dt=_parse_dt(d["공고일시"]))
    d = d.sort_values(["공고번호", "_ord", "_dt"], kind="stable").drop_duplicates("공고번호", keep="last")
    return d.drop(columns=["_ord"]).reset_index(drop=True)


# ── 2. 용역 종류 ─────────────────────────────────────────────────────────────────────────────────────────
# 공고명에서 시설 이름을 가린 뒤 판정한다(link 가 가려서 넘긴다) — '어울림플랫폼'·'데이터센터'·'정보통신공제조합 회관' 처럼 시설 이름 안에 든 말에 걸리지 않게.
# 판정 순서가 뜻을 가진다: ① 이 사업의 용역이 아닌 것(운영 폐기물·정보시스템) ② 금액을 더하면 안 되는 부대 용역(VE·공모 관리·타당성 등) ③ 세 종류.
_K_WASTE_OPS = re.compile(r"생활\s*폐기물|음식물|재활용|대형\s*(?:생활\s*)?폐기물|의료\s*폐기물|감염성|해양\s*폐기물|침적\s*폐기물|영농\s*폐기물|방치\s*폐기물|불법\s*투기|무단\s*투기|"
                          r"슬러지|오니|가축\s*분뇨|폐\s*수(?![가-힣])|폐\s*시약|폐\s*농약|폐\s*비닐|위탁\s*소각|"
                          r"폐기물\s*(?:종합\s*|공공\s*|광역\s*)?(?:처리|소각|매립|처분|재활용|자원화|에너지화|선별|전처리)\s*(?:시설|장(?!기\s*계속)|센터|타운|단지)|"
                          r"폐기물\s*(?:에너지|감량|관리\s*시스템|수수료|처분\s*부담금|처리\s*기본\s*계획)|소각\s*(?:시설|장)|매립\s*(?:시설|장|지)|자원\s*(?:회수|순환)|"
                          r"환경\s*에너지\s*(?:센터|타운|시설)|클린\s*센터|적환\s*시설|수집\s*[·ㆍ,및]?\s*운반\s*대행|대행\s*용역|민간\s*위탁|청소\s*용역|가로\s*청소|"
                          r"원가\s*(?:계산|산정)|연구\s*용역|실태\s*조사")
_K_WASTE_CLASS_OPS = re.compile(r"생활폐기물|음식물|재활용|의료폐기물|하수|분뇨|슬러지|청소")
_K_WASTE = re.compile(r"폐기물|폐\s*콘크리트|폐\s*아스콘|폐\s*석면|폐\s*보드|폐\s*임목|폐\s*목재|폐\s*(?:판넬|패널)|폐\s*토석|석면\s*(?:폐기물|처리)|건설\s*오니")
# 공사에서 나온 폐기물임을 보이는 말. 이 말이 있으면 '재활용'(처리 방법)·'민간위탁'(계약 방식)·'오니'(건설오니)는 운영 폐기물의 표시로 보지 않는다
_K_WASTE_CONSTR = re.compile(r"건설\s*폐기물|건설\s*오니|건설\s*폐재류|폐\s*아스콘|폐\s*콘크리트|임목\s*폐기물|폐\s*임목|폐\s*목재|폐\s*(?:판넬|패널)|폐\s*토석")
_K_WASTE_CONSTR_OK = re.compile(r"재활용(?!\s*(?:시설|센터|선별|품))|민간\s*위탁|(?:건설\s*)?오니")      # '재활용시설'·'재활용품 선별장' 은 처리 방법이 아니라 시설이다
# 공고명에 든 공사·사업 말 — 이 말이 있으면 그 공사·사업의 폐기물이다(조달분류가 운영 폐기물 쪽이어도). 유지보수 공사인지는 뒤에서 사업유형으로 가린다
# ('○○ 리모델링공사 폐기물처리', '개보수 공사 가연성폐기물', '조성사업 인조잔디 재활용처리'). 공사 말과 폐기물 말 사이에 다른 말이 끼면('신축 이전에 따른 사무집기 폐기물',
# '증축공사 준공 청소 및 폐기물') 그 공사의 폐기물이 아니다. 기관 이름 속의 '공사'('○○도시공사')와 '공사기간' 은 공사 말이 아니다
_K_WASTE_WORK = re.compile(
    r"(?:(?<!도시)(?<!개발)(?<!주택)(?<!관광)(?<!교통)(?<!시설)공사(?!\s*기간)|사업(?!장)|건립|신축|증축|개축|재건축|리모델링|개보수|대수선|철거|해체)"
    r"\s*(?:[\(（][^\(\)（）]{0,14}[\)）]\s*)?(?:에\s*따른\s*)?(?:현장\s*)?(?:발생\s*)?(?:추가\s*)?"
    r"(?:[\(（]?\s*(?:건설|혼합|가연성|불연성|매립|임목|지정|사업장|일반|인조\s*잔디)\s*[\)）]?\s*)*"
    r"(?:폐\s*기\s*물|폐\s*[가-힣]{2,4}|인조\s*잔디\s*(?:분리\s*회수|재활용|철거))")
# 폐기물 '처리' 용역임을 보이는 말(마지막에 오는 용역 이름이 이것이면 폐기물 처리 용역, 설계·감리이면 그쪽 용역)
_K_WASTE_JOB = re.compile(r"(?:건설\s*오니|폐(?:기물|석면|콘크리트|아스콘|보드|임목))[^,]{0,25}?(?:처리|운반|수집|수거|위탁|상차|반출)|(?:처리|운반|수집|수거)[^,]{0,12}?폐기물|(?:처리|운반)\s*용역")
_K_SERVICE_NOUN = re.compile(r"설계|감리|건설\s*사업\s*관리|(?<![A-Za-z])CM(?![A-Za-z])|공모")
_K_UNIT = re.compile(r"연간\s*단가|단가\s*계약|(?<![가-힣])연간(?![가-힣])|단가\s*[\)）]|(?<![가-힣])관내(?![가-힣])|(?:각종|소규모|일반)\s*(?:건설)?\s*공사|읍\s*[·ㆍ]\s*면|"
                     r"수시\s*발생|적환장|야적장|집하장|임시\s*보관|(?:상|하)반기|\d\s*분기|(?<![0-9])\d{1,2}\s*(?:권역|구역)(?![가-힣])|"
                     r"(?:20\d\d\s*년(?:도)?|년도)\s*(?:건설\s*|사업장\s*|지정\s*)?폐기물")
_K_ENV_FACILITY = re.compile(r"(?:소각|매립|자원\s*회수|자원\s*순환|자원화|재활용\s*선별|에너지화|슬러지\s*(?:건조|처리|소각)?|가축\s*분뇨\s*(?:공공\s*)?처리|하수\s*처리|"
                             r"폐수\s*처리|정수)\s*(?:시설|장|센터|타운|단지)|(?:쓰레기|위생)\s*매립|환경\s*에너지\s*(?:센터|타운|시설)|클린\s*센터")
_K_IT = re.compile(r"정보\s*시스템|정보화|전산|시스템\s*(?:구축|고도화|개발|통합|운영)|통합\s*시스템|소프트웨어|(?<![A-Za-z])SW(?![A-Za-z])|(?<![A-Za-z])ISP(?![A-Za-z])|"
                   r"네트워크|유무선|통신망|플랫폼\s*구축|정보\s*자원|통합\s*구축|홈페이지|데이터\s*(?:구축|베이스)|클라우드|차세대\s*[가-힣A-Za-z]{0,12}\s*(?:시스템|구축)|감리\s*법인|개인\s*정보")
_K_SIDE = re.compile(r"설계\s*(?:의\s*)?경제성|(?<![A-Za-z])VE(?![A-Za-z])|가치\s*공학|설계\s*의도|사후\s*설계\s*관리|설계\s*안전성|안전성\s*검토|(?<![A-Za-z])DFS(?![A-Za-z])|"
                     r"안전\s*관리\s*계획|유해\s*위험|기술\s*지도|재해\s*예방(?!\s*센터)|공모\s*(?:의\s*)?(?:관리|운영|대행|지원|진행|홍보|심사)|공모\s*지침|"
                     r"설계\s*(?:검토|심의|자문|적정성|검증)|원가\s*(?:계산|검토)|(?<![가-힣])적산(?![가-힣])|공사비\s*(?:검증|검토)|"
                     r"타당성|기본\s*계획|기본\s*구상|건축\s*기획(?!\s*(?:설계|및\s*(?:기본|실시)))|지구\s*단위|도시\s*계획|연구\s*용역|영향\s*평가|"
                     r"(?<!시험)인증(?!\s*센터|\s*지원\s*센터|생산|동(?![가-힣]))|컨설팅|측량|지반\s*조사|지질\s*조사|"
                     r"문화재\s*(?:지표|발굴|시굴)|발굴\s*조사|감정\s*평가|석면\s*조사|해체\s*계획|안전\s*진단|안전\s*점검|내진\s*성능|구조\s*검토|적정성\s*검토|"
                     r"사후\s*평가|백서|홍보(?!\s*(?:관(?![가-힣])|교육관|체험관|전시관|판매장|원(?![가-힣])))|이전\s*용역|이사\s*용역")
# 부대 용역 판정에서 보지 않는 자리: 사업 이름 머리의 '도시계획시설(…)'·'군계획시설', 괄호 속 '… 포함'(본 용역에 딸린 일)
_K_SIDE_SKIP = re.compile(r"(?:도시|군)\s*계획\s*시설|[\(（][^\(\)（）]{0,30}포함\s*[\)）]|[\(（\[]\s*공모\s*대행\s*[\)）\]]")
# 설계 용역에 딸려 함께 발주되는 일이 설계 앞에 이어 적힌 꼴 — '조사측량 및 실시설계', '기본계획 및 실시설계', '공간컨설팅, 실시설계' 는 설계 용역이다(부대 용역이 아니다)
_K_SIDE_WITH_DESIGN = re.compile(r"(?:측량|지반\s*조사|지질\s*조사|기본\s*계획|기본\s*구상|컨설팅)\s*(?:수립\s*)?(?:및|[,·ㆍ/])\s*(?:기본\s*및\s*)?(?:실시|기본)\s*설계|"
                                 r"(?:실시|기본)\s*설계\s*(?:및|[,·ㆍ/])\s*(?:조사\s*)?(?:측량|지반\s*조사|지질\s*조사)")
_K_CM = re.compile(r"건설\s*사업\s*관리|감독\s*권한\s*대행|(?<![A-Za-z])CM(?![A-Za-z])|(?<![A-Za-z])PM\s*용역")
_K_SUPERVISE = re.compile(r"감리")
_K_COMBINED = re.compile(r"설계\s*(?:[·ㆍ,/]|및)\s*(?:감리|건설\s*사업\s*관리)|설계\s*감리(?!\s*단)")
_K_COMPETE = re.compile(r"설계\s*(?:제안\s*)?공모|설계안\s*공모|현상\s*설계|설계\s*경기|설계\s*재\s*공모|국제\s*지명|지명\s*공모|간이\s*공모|[\(（]\s*공모\s*[\)）]")
_K_PROPOSAL = re.compile(r"제안\s*공모")
_K_BUILD_WORD = re.compile(r"건립|신축|증축|개축|재건축|리모델링|개보수|대수선|조성|공사")       # '설계' 낱말 없는 제안공모가 건축 설계 공모임을 보이는 말(조달분류 '건축설계…' 와 함께 있을 때)
_K_DESIGN = re.compile(r"설계")
_K_DESIGN_NOT = re.compile(r"전시\s*(?:물|설계|기획|연출|콘텐츠|컨텐츠)|설계\s*(?:[·ㆍ,/]|및)?\s*제작|제작\s*(?:[·ㆍ,/]|및)?\s*설치\s*용역|미디어\s*아트|조형물|인테리어\s*설계")
_K_SPLIT_DESIGN = re.compile(r"(?:전기|정보\s*통신|통신|소방|기계\s*설비|기계)(?:\s*(?:분야|부문|공사|시설))?(?:\s*(?:[·ㆍ,/]|및)?\s*(?:전기\s*소방|기계\s*소방|전기|정보\s*통신|통신|소방|기계\s*설비|기계)"
                             r"(?:\s*(?:분야|부문|공사|시설))?)*\s*(?:실시\s*)?설계")
_K_REDESIGN = re.compile(r"[\(（]\s*재\s*[\)）]\s*(?:기본|실시)|(?<!자)재\s*설계|보완\s*설계|설계\s*보완")
_K_STAGE_DESIGN = re.compile(r"설계\s*단계|설계\s*및\s*시공\s*단계|실시\s*설계에\s*대한")
_K_SUP_SUB = (("석면", r"석면"), ("해체", r"해체|철거"), ("전기", r"전기|전력"), ("소방", r"소방"), ("정보통신", r"통신"), ("건축", r"건축|상주|책임"))
_K_DEMOLITION = re.compile(r"석면|해체|철거")
_K_MAIN_WORK = re.compile(r"(?:신축|증축|개축|재건축|건립)(?!\s*(?:예정\s*)?부지)|리모델링|개보수|대수선|포함")
_K_PAREN = re.compile(r"[\(（][^\(\)（）]*[\)）]")


def _demolition_supervision(t: str) -> bool:
    """감리 대상이 석면·해체(철거) 공사 자체인가 — 마지막 '석면·해체·철거' 와 '감리' 사이에 본 공사 말이 없을 때만.
    '◇ 철거 및 신축공사 감리용역'·'◇ 리모델링공사(철거 포함) 감리용역' 은 본 공사의 감리, '◇ 리모델링공사 석면해체 제거공사 감리용역'·'◇ 해체공사 감리용역' 은 해체 감리."""
    head = t.split("감리")[0]
    if not _K_DEMOLITION.search(head):
        head = t
    last = None
    for last in _K_DEMOLITION.finditer(head):
        pass
    if last is None:
        return False
    # 괄호 속 말은 '… 포함' 일 때만 본다('석면해체·제거공사(리모델링 구간) 감리' 는 해체 감리, '리모델링공사(철거 포함) 감리' 는 본 공사 감리)
    tail = _K_PAREN.sub(lambda m: m.group(0) if "포함" in m.group(0) else " ", head[last.end():])
    return not _K_MAIN_WORK.search(tail)


_K_NOT_BUILDING = re.compile(r"숲\s*가꾸기|조림|사방\s*(?:사업|댐)|임도|풀베기|산불|하천|관망|관로|하수\s*처리|공공\s*하수|정수장|상수도|문화재|국가\s*유산|"
                             r"도로\s*(?:확포장|개설|건설|관리)|확포장|농촌\s*중심지|기초\s*생활\s*거점")
_K_HOUSING_SUPERVISOR = re.compile(r"감리자\s*(?:[\(（][^\)）]*[\)）]\s*)?(?:지정|모집)")
_K_PUBLIC_HOUSING = re.compile(r"공공|임대|행복\s*주택")
BUNDLE = re.compile(r"(?:외|등)\s*\d+\s*(?:개\s*)?(?:교|건|개소|곳|동|블록|시설|사업|단지|지구|현장|구간|기관|학교)(?![가-힣])|\d+\s*개\s*(?:교|학교)(?![가-힣])")


_UNIT_WHY = "연간·단가 계약으로 보임"


def _unit_words_from(title: str, names: List[str]) -> bool:
    """공고명에서 연간·단가 계약으로 읽힌 말이 모두 그 프로젝트의 공사 공고명에도 있는 말인가
    ('○○ 조성사업(2차)_1구역', '2025년 상반기 ○○ 리모델링공사' 처럼 공사 이름에 든 '구역'·'상반기')."""
    hits = [_ns(h.group(0)) for h in _K_UNIT.finditer(_txt(title))]
    return bool(hits) and all(any(h in n for n in names) for h in hits)


def classify_kind(title, cls: str = "") -> Tuple[str, str, str]:
    """용역 공고명(+공공조달분류명) → (용역종류, 세부, 합산하지 않는 까닭).
    용역종류: '설계' / '건설사업관리·감리' / '폐기물처리' / '기타'. 까닭이 빈 문자열이면 그 종류의 금액으로 더해도 되는 공고다.
    '기타' 는 연결돼도 08 에만 싣는다(운영 폐기물·정보시스템 감리·타당성조사·VE·공모 관리 등 — 이 사업의 설계비·감리비·폐기물 처리비가 아니다)."""
    kind, sub, why = _classify_kind(_txt(title), _txt(cls))
    t = _txt(title)
    if not why and kind != "기타":
        if kind != "폐기물처리" and _K_UNIT.search(t):
            why = _UNIT_WHY + "(여러 사업을 한 계약으로 처리)"
        elif _K_HOUSING_SUPERVISOR.search(t) and not _K_PUBLIC_HOUSING.search(t):
            why = "주택법에 따른 감리자 지정·모집 공고(사업 주체가 민간일 수 있음)"
    return kind, sub, why


def _classify_kind(t: str, c: str) -> Tuple[str, str, str]:
    if _K_WASTE.search(t) or "폐기물" in c:
        # 공고명이 공사·사업에서 나온 폐기물임을 밝혔으면('건설폐기물'·'폐아스콘', '○○ 리모델링공사 폐기물처리') 조달분류는 보지 않는다 —
        # 공사 폐기물에도 '폐기물재활용서비스'·'생활폐기물처리서비스' 가 흔히 붙는다(2026-11 실자료: 이 때문에 공사 폐기물 처리 용역 60여 줄이 운영 폐기물로 읽혔다)
        constr = bool(_K_WASTE_CONSTR.search(t)) or bool(_K_WASTE_WORK.search(t))
        if _K_WASTE_OPS.search(_K_WASTE_CONSTR_OK.sub(" ", t) if constr else t) or (not constr and _K_WASTE_CLASS_OPS.search(c)):
            return "기타", "운영 폐기물·폐기물 시설", "공사에서 나온 폐기물의 처리 용역이 아님(생활·음식물 폐기물, 폐기물 처리시설의 운영·설계·감리 등)"
        job = [m.end() for m in _K_WASTE_JOB.finditer(t)]
        noun = [m.end() for m in _K_SERVICE_NOUN.finditer(t)]
        if job and (not noun or max(job) >= max(noun)) or (not noun and "폐기물" in c):
            sub = "지정폐기물(석면)" if re.search(r"석면", t) else "지정폐기물" if re.search(r"지정", t) else "임목폐기물" if re.search(r"임목", t) \
                else "건설폐기물" if re.search(r"건설|혼합|가연성|불연성|폐\s*콘크리트|폐\s*아스콘|폐\s*보드", t + c) else "폐기물"
            if _K_UNIT.search(t):
                return "폐기물처리", sub, _UNIT_WHY + "(여러 공사의 폐기물을 한 계약으로 처리)"
            return "폐기물처리", sub, ""
        # 폐기물이라는 말이 있지만 마지막 용역 이름이 설계·감리 → 아래의 설계·감리 판정으로('○○청사 철거공사(석면·폐기물 포함) 감리용역')
    if _K_ENV_FACILITY.search(t):                     # 공고명에 '폐기물' 이 없어도 소각장·매립장·자원순환시설·하수처리시설 사업의 설계·건설사업관리
        return "기타", "환경기초시설", "폐기물·하수 처리시설 사업의 용역으로 보임(건축 사업의 용역이 아님)"
    if _K_IT.search(t):
        return "기타", "정보시스템·통신망", "건축 사업의 용역이 아님"
    if _K_COMBINED.search(t):
        return "기타", "설계·감리 일괄", "설계와 감리(건설사업관리)를 한 계약으로 묶은 용역 — 한 종류의 금액으로 넣을 수 없음"
    side = _K_SIDE.search(_K_SIDE_WITH_DESIGN.sub(" 실시설계", _K_SIDE_SKIP.sub(" ", t)))
    if _K_CM.search(t):
        if side and not _K_STAGE_DESIGN.search(t):
            return "기타", "건설사업관리 부대(" + side.group(0) + ")", "건설사업관리비가 아닌 부대 용역"
        return "건설사업관리·감리", "건설사업관리" + ("(설계단계 포함)" if (_K_STAGE_DESIGN.search(t) or side) else ""), ""
    if _K_SUPERVISE.search(t):
        if side:
            return "기타", "감리 부대(" + side.group(0) + ")", "감리비가 아닌 부대 용역"
        subs = [lab for lab, pat in _K_SUP_SUB if re.search(pat, t)]
        if _demolition_supervision(t):
            return "건설사업관리·감리", "감리(석면·해체)", "석면·해체 공사의 감리(본 공사의 감리비와 따로 봄)"
        subs = [s for s in subs if s not in ("석면", "해체")]
        return "건설사업관리·감리", "감리" + (f"({'·'.join(subs)})" if subs else ""), ""
    # '제안공모' 만으로는 설계가 아니다(협상계약 방식의 콘텐츠 제작·운영·행사 용역도 '제안공모' 다) — 설계·공모 낱말이 있거나 조달분류가 설계 용역이어야 한다
    if _K_COMPETE.search(t) or _K_DESIGN.search(t) or (_K_PROPOSAL.search(t) and c.startswith("건축설계") and _K_BUILD_WORD.search(t)):
        if side:
            return "기타", "설계 부대(" + side.group(0) + ")", "설계비가 아닌 부대 용역"
        if _K_DESIGN_NOT.search(t):
            return "기타", "전시·제작 설계", "건축 설계비가 아님(전시물 제작·설치 등)"
        if _K_COMPETE.search(t) or _K_PROPOSAL.search(t):
            return "설계", "설계공모", ""
        if _K_REDESIGN.search(t):
            return "설계", "설계용역(재설계·보완)", ""
        if _K_SPLIT_DESIGN.search(t) and not re.search(r"건축", t):
            return "설계", "설계용역(전기·통신·소방 분리발주)", ""
        return "설계", "설계용역", ""
    return "기타", "", "설계·건설사업관리·감리·폐기물 처리 용역이 아님"


# ── 3. 기관 맞춤 ─────────────────────────────────────────────────────────────────────────────────────────
_CORP = re.compile(r"[\(（](?:주|재|사|학|의|특)[\)）]|주식회사|재단법인|사단법인")
_PROV_TOKEN = re.compile(r"(.+?)(특별자치도|특별자치시|통합특별시|특별시|광역시|자치도|도)$")
_MUNI_TOKEN = re.compile(r".{1,8}(?:시|군|구)$")
_ORG_WORD = re.compile(r"사업소|공사|공단|본부|센터|재단|학교|대학|병원|조합|협회|위원회|연구|사업단|교육청|지원청|관리소|박물관|도서관|미술관|회관|체육회|농원|기구")
_PROV_ABBR = {"전라북": "전북", "전라남": "전남", "경상북": "경북", "경상남": "경남", "충청북": "충북", "충청남": "충남"}


@lru_cache(maxsize=200_000)
def inst_key(v) -> str:
    """기관 이름 비교용 키: 법인 표시와 공백을 뗀다."""
    return re.sub(r"\s+", "", _CORP.sub("", _txt(v)))


def _prov_stems(tok: str) -> frozenset:
    """광역 이름 → 권역 낱말. 이름이 바뀌어도 같은 곳으로 본다: '강원도'·'강원특별자치도' → {강원}, '전라북도'·'전북특별자치도' → {전북},
    합쳐진 이름은 양쪽 모두('전남광주통합특별시' → {전남, 광주})."""
    m = _PROV_TOKEN.fullmatch(tok)
    if not m:
        return frozenset()
    stem = _PROV_ABBR.get(m.group(1), m.group(1))
    if len(stem) >= 4 and len(stem) % 2 == 0:
        return frozenset(stem[i:i + 2] for i in range(0, len(stem), 2))
    return frozenset([stem]) if 2 <= len(stem) <= 3 else frozenset()


@lru_cache(maxsize=200_000)
def inst_sig(v) -> Tuple[str, frozenset, str, str]:
    """기관 이름 → (갈래, 권역, 시·군·구, 첫 어절). 갈래: 'gov'(지자체: 광역·기초와 그 사업소) / 'edu'(교육청과 그 소속) / 'org'(그 밖)."""
    toks = [t for t in re.split(r"\s+", _CORP.sub("", _txt(v))) if t]
    if not toks:
        return "", frozenset(), "", ""
    head = toks[0]
    if head.endswith("교육청"):
        return "edu", _prov_stems(head[:-3]), "", head
    prov = _prov_stems(head) if not _ORG_WORD.search(head) else frozenset()
    muni = ""
    for t in toks[(1 if prov else 0):]:
        t = re.sub(r"(?<=[시군구])청$", "", t)          # '○○광역시 동구청' 은 자치구 '동구' 다(광역 사업소로 읽으면 그 광역의 모든 사업소와 같은 계열이 된다)
        if _MUNI_TOKEN.fullmatch(t) and not _ORG_WORD.search(t) and not _prov_stems(t):
            muni = t                                   # 광역 다음의 첫 시·군·구('경기도 수원시 영통구' → 수원시)
            break
    if prov or muni:
        return "gov", prov, muni, head
    return "org", frozenset(), "", head


def inst_same(a, b) -> bool:
    """두 기관이 같은 발주 주체(같은 지자체·같은 교육청·같은 기관과 그 소속)인지.
    - 이름이 같다
    - 지자체: 시·군·구가 같고 광역이 어긋나지 않는다(광역 이름이 바뀐 '강원도 강릉시' = '강원특별자치도 강릉시', 사업소 '경상북도 경주시 시립도서관' = '경상북도 경주시').
      광역 본청·사업소끼리는 같은 광역('서울특별시 한강사업본부' = '서울특별시 도시기반시설본부'). 광역 본청과 그 안의 시·군·구는 다른 기관
    - 교육청: 같은 시·도 교육청(지원청·학교 포함)
    - 그 밖(공사·공단·대학·국가기관): 첫 어절이 같거나 한쪽 이름이 다른 쪽에 들어 있다('한국토지주택공사 경기남부지역본부' = '한국토지주택공사')"""
    ka, kb = inst_key(a), inst_key(b)
    if not ka or not kb:
        return False
    if ka == kb:
        return True
    ta, pa, ma, ha = inst_sig(a)
    tb, pb, mb, hb = inst_sig(b)
    if ta != tb:
        return False
    if ta == "gov":
        if ma or mb:
            return ma == mb and (not pa or not pb or bool(pa & pb))
        return bool(pa & pb)
    if ta == "edu":
        return bool(pa & pb) or ha == hb
    s, l = (ka, kb) if len(ka) <= len(kb) else (kb, ka)
    return ha == hb or (len(s) >= 4 and s in l)


_PROV_END = re.compile(r"특별자치도|특별자치시|통합특별시|특별시|광역시|자치도|도")


def _sans_prov(v, stems: frozenset) -> str:
    """머리 어절(광역·교육청)을 뺀 나머지 어절에서 그 광역의 이름으로 된 머리를 뗀 꼴: '강원특별자치도교육청 강원특별자치도가온교육지원청' → '가온교육지원청'."""
    out = []
    for t in _CORP.sub("", _txt(v)).split()[1:]:
        for m in _PROV_END.finditer(t):
            if m.end() < len(t) and _prov_stems(t[:m.end()]) & stems:
                t = t[m.end():]
                break
        out.append(t)
    return "".join(out)


def inst_narrow(a, b) -> bool:
    """inst_same 가운데서도 '바로 그 기관' 인 경우: 이름이 같다 / 한쪽 이름이 다른 쪽에 들어 있다(본청·본사와 그 소속) / 같은 시·군·구(부서·사업소가 달라도) /
    광역 본청·교육청 본청과 그 소속(이름이 바뀐 '강원도' ↔ '강원특별자치도 ○○사업소' 포함).
    부처·공사의 다른 소속기관('국가보훈처 국립대전현충원' ↔ '국가보훈처 국립이천호국원'), 같은 시·도 교육청의 다른 학교·지원청, 광역의 다른 사업소끼리는 아니다 —
    이런 기관의 공고는 공고명에 프로젝트의 사업 이름이 통째로 있을 때만 더한다(link). 공사 쪽 재검색이 검수_수요기관이 들어 있는 공고만 받는 것과 맞춘 것이다."""
    if not inst_same(a, b):
        return False
    ka, kb = inst_key(a), inst_key(b)
    s, l = (ka, kb) if len(ka) <= len(kb) else (kb, ka)
    if s in l:
        return True
    ta, pa, ma, _ = inst_sig(a)
    tb, pb, mb, _ = inst_sig(b)
    if ta == "gov" and (ma or mb):
        return True                                    # inst_same 이 이미 같은 시·군·구임을 확인했다
    if ta in ("gov", "edu") and (pa & pb):
        if len(_txt(a).split()) == 1 or len(_txt(b).split()) == 1:
            return True
        # 광역 이름만 바뀐 같은 소속기관: '강원도교육청 강원도가온교육지원청' = '강원특별자치도교육청 강원특별자치도가온교육지원청', '전라북도 농업기술원' = '전북특별자치도 농업기술원'
        xa, xb = _sans_prov(a, pa | pb), _sans_prov(b, pa | pb)
        return bool(xa) and xa == xb
    return False


_WIDE_WHY = "발주 기관이 이 시설의 수요기관과 같은 계열의 다른 기관(소속기관·지사·학교가 다름)"
_GENERIC_WHY = "일반적인 시설 이름만으로 이어졌고"
_DELEGATE = re.compile(r"도시\s*공사|개발\s*공사|시설\s*관리\s*공단|시설\s*공단|도시\s*관리\s*공단|문화\s*재단|체육회|종합\s*건설\s*본부|도시\s*기반\s*시설\s*본부|건설\s*본부")


def inst_related(a, b) -> bool:
    """같은 기관은 아니지만 그 지자체의 일을 대신 발주하는 기관으로 보이는지: 한쪽이 지자체이고 다른 쪽 이름에 그 지자체 이름(시·군·구 또는 광역의 권역 낱말)과
    대행 기관 말(도시공사·시설관리공단·문화재단·체육회·건설본부 …)이 함께 있다('경기도 성남시' ↔ '성남도시개발공사', '서울특별시 송파구' ↔ '서울특별시 도시기반시설본부')."""
    for x, y in ((a, b), (b, a)):
        tx, px, mx, _ = inst_sig(x)
        if tx != "gov":
            continue
        ky = inst_key(y)
        if not _DELEGATE.search(_txt(y)):
            continue
        stems = ([mx[:-1]] if len(mx) >= 3 else []) + [p for p in px if not mx]
        ty, py, my, _ = inst_sig(y)
        if ty == "gov" and not my and px and (px & py):            # 광역 본청의 건설본부 ↔ 그 광역 안의 시·군·구
            return True
        if any(st and st in ky for st in stems):
            return True
    return False


# ── 4. 시설의 검색 열쇠 ──────────────────────────────────────────────────────────────────────────────────
_NAME_MIN = 3            # 공백을 뺀 시설 이름이 이보다 짧으면 잇지 않는다
_NAME_SHORT = 3          # 이 길이 이하의 이름(지자체 접두를 뗀 '미술관'·'수장고')은 공고명에 띄어쓰기 없이 그대로 있을 때만 인정한다
_NAME_GENERIC = 4        # 검수한 이름·별칭은 공고명에 없고 지자체·기관 접두를 뗀 이름만 있는데 그것이 이 길이 이하('주차타워'·'문화회관'·'공공주택')면
                         # 같은 기관의 다른 사업일 수 있다 → 공사와 시점이 가까울 때만 더한다(검수한 이름이 원래 짧은 '○○회관' 은 해당하지 않는다)
_BR_CHARS = re.compile(r"[\[\]\(\)（）【】<>〈〉「」『』]")
_BR_SPAN = re.compile(r"[\[\(（【<〈][^\[\]\(\)（）【】<>〈〉]{0,40}[\]\)）】>〉]")


def facility_keys(reviewed: pd.DataFrame, hist: pd.DataFrame) -> Dict[str, dict]:
    """시설ID → {"name": 검수_시설명, "keyword": 검색어, "pats": [공백 뺀 이름·별칭 …], "loose": [어절 간격 허용 정규식 …],
                "insts": [그 시설 공사 공고의 수요기관 + 검수_수요기관 …], "codes": {수요기관코드 …}, "full": [pats 가운데 검수한 이름·별칭 그대로인 것]}.
    이름 만들기는 research(discover.research_by_facility)와 같다: 검수_시설명 + 별칭, 지자체 접두를 뗀 꼴도 함께."""
    from .classify import strip_institution_prefix
    from .discover import _blank, _gap_pattern
    keys: Dict[str, dict] = {}
    if hist.empty:
        return keys
    hf = hist["시설ID"].astype(str)
    hist_inst = {f: sorted({_txt(v) for v in g if _txt(v)}) for f, g in hist.groupby(hf)["수요기관"]}
    hist_code = {f: {_no(v) for v in g if _no(v)} for f, g in hist.groupby(hf)["수요기관코드"]} if "수요기관코드" in hist.columns else {}
    for fid, grp in reviewed.groupby(reviewed["시설ID"].astype(str).str.strip(), sort=False):
        if fid not in hist_inst:                       # 재검색에서 공고가 하나도 남지 않은 시설(프로젝트 없음)에는 이을 곳이 없다
            continue
        names, insts, full = [], list(hist_inst[fid]), []
        for _, r in grp.iterrows():
            nm = [r["검수_시설명"]] + [a.strip() for a in str(r.get("검수_별칭(;구분)") or "").split(";") if a.strip()]
            inst_src = str(r.get("검수_수요기관") or r.get("수요기관") or "")
            names += [n for n in nm if n and not _blank(n)]
            full += [_ns(n) for n in nm if n and not _blank(n)]
            names += [strip_institution_prefix(n, inst_src) for n in nm if n and not _blank(n)]
            if _txt(r.get("검수_수요기관")):
                insts.append(_txt(r.get("검수_수요기관")))
        pats = [p for p in dict.fromkeys(_ns(n) for n in names) if len(p) >= _NAME_MIN]
        if not pats:
            continue
        loose = [g for g in dict.fromkeys(_gap_pattern(n) for n in names) if g]
        keys[fid] = {"name": _txt(grp.iloc[0]["검수_시설명"]), "keyword": _txt(grp.iloc[0].get("검색어")), "pats": pats, "loose": loose,
                     "insts": list(dict.fromkeys(insts)), "codes": hist_code.get(fid, set()), "full": [p for p in dict.fromkeys(full) if len(p) >= _NAME_MIN]}
    return keys


_JOIN = re.compile(r"[및·ㆍ・,/&+]")


@lru_cache(maxsize=1024)
def _span_forms(title: str) -> Tuple[str, str, str]:
    """이름을 찾는 세 가지 꼴(match_facilities 와 같다): 0 = 괄호 기호만 뺀 꼴, 1 = 괄호 속까지 뺀 꼴, 2 = 괄호 그대로(공백만 뺀 꼴)."""
    return _ns(_BR_CHARS.sub("", title)), _ns(_BR_SPAN.sub(" ", title)), _ns(title)


def _name_spans(title: str, pats: List[str], loose: Tuple[str, ...] = ()) -> List[Tuple[int, int, int]]:
    """공고명에서 이 시설의 이름이 놓인 자리 (꼴, 시작, 끝) 들. match_facilities 가 이름을 찾는 꼴마다 따로 찾는다 —
    '가온(국민체육)센터' 처럼 괄호 속을 빼야 이름('가온센터')이 맞는 시설, 어절 사이에 다른 말이 낀 꼴(loose)로만 맞은 시설도 자리를 갖게."""
    out = []
    for f, t in enumerate(_span_forms(title)):
        for p in pats:
            q = p if f == 2 else _BR_CHARS.sub("", p)
            i = t.find(q) if q else -1
            while i >= 0:
                out.append((f, i, i + len(q)))
                i = t.find(q, i + 1)
        if f == 2 and not out:                         # 정확한 이름이 어느 꼴에도 없을 때만(= 느슨한 꼴로 이어진 줄)
            for g in loose:
                out += [(f, m.start(), m.end()) for m in re.finditer(g, t)]
    return out


def _named_apart(title: str, a: List[Tuple[int, int, int]], b: List[Tuple[int, int, int]]) -> bool:
    """두 시설의 이름이 공고명의 서로 다른 자리에 '및'·'·'·',' 로 이어져 있는가 — '○○ 체육관 및 △△ 수영장 건립 설계용역' 같은 묶음 발주.
    한 이름이 다른 이름 안에 든 경우('○○ 체육관' ⊂ '○○ 체육관 수영장')나 잇는 말 없이 붙은 경우('□□지구 A1블록 공공주택')는 해당하지 않는다."""
    forms = _span_forms(title)
    for f1, s1, e1 in a:
        for f2, s2, e2 in b:
            if f1 != f2:                               # 같은 꼴에서 찾은 자리끼리만 견준다
                continue
            t = forms[f1]
            gap = t[e1:s2] if e1 <= s2 else (t[e2:s1] if e2 <= s1 else None)
            if gap is not None and _JOIN.search(gap):
                return True
    return False


def _mask_name(title: str, pats: List[str]) -> str:
    """공고명에서 시설 이름(띄어쓰기는 달라도 됨)을 '◇' 로 가린다 — 종류·잡음 판정이 시설 이름 속의 말에 걸리지 않게."""
    out = title
    for p in sorted(pats, key=len, reverse=True):
        out = re.sub(r"\s*".join(re.escape(ch) for ch in p), " ◇ ", out)
    return out


# ── 5. 연결 ──────────────────────────────────────────────────────────────────────────────────────────────
def match_facilities(sv: pd.DataFrame, keys: Dict[str, dict]) -> pd.DataFrame:
    """용역 공고(최신 차수) × 시설 → 연결 후보. 조건: 공고명에 시설 이름이 있고(공백 무시) 수요기관(없으면 공고기관)이 그 시설의 기관과 같다.
    - 이름: 정확히 들어 있으면 '확실', 어절 사이에 다른 말이 낀 꼴로만 맞으면 '추정'. 괄호는 그대로 둔 꼴·괄호 기호만 뺀 꼴·괄호 속까지 뺀 꼴 모두에서 찾는다
      ('교양(국민체육)센터' ↔ 시설 이름 '교양센터')
    - 기관: 수요기관코드가 같거나 이름이 같은 발주 주체(inst_same)면 그대로, 그 지자체의 대행 기관(inst_related)이면 '추정'. 그 밖의 기관이면 잇지 않는다
      (같은 이름의 다른 지자체 시설 — 용역 공고명에는 지자체 이름이 없는 경우가 많다). 공고기관은 수요기관이 빈 공고에서만 대신 본다(조달청이 대신 공고하는 일이 흔하다).
    - 같은 발주 주체 가운데서도 '바로 그 기관'(inst_narrow)이 아닌 기관의 공고는 _wide 로 표시한다 → link 가 사업 이름이 통째로 있을 때만 더한다."""
    cols = ["_i", "시설ID", "연결강도", "연결근거", "_wide"]
    if sv.empty or not keys:
        return pd.DataFrame(columns=cols)
    spaced = sv["공고명"].map(_txt)
    t_a = spaced.map(_ns)                              # 괄호 그대로
    t_b = spaced.map(lambda t: _ns(_BR_CHARS.sub("", t)))        # 괄호 기호만 뺌
    t_c = spaced.map(lambda t: _ns(_BR_SPAN.sub(" ", t)))        # 괄호 속까지 뺌
    dem, ntc = sv["수요기관"].map(_txt), sv["공고기관"].map(_txt)
    who = dem.where(dem != "", ntc)                    # 발주 주체: 수요기관(비어 있을 때만 공고기관)
    code = sv["수요기관코드"].map(_no) if "수요기관코드" in sv.columns else pd.Series("", index=sv.index)
    by_inst = {v: ix.to_numpy() for v, ix in who.groupby(who).groups.items() if v}                 # 기관 이름 → 그 기관의 용역 공고(행 라벨)
    by_code = {v: ix.to_numpy() for v, ix in code.groupby(code).groups.items() if v}
    uniq = list(by_inst)
    delegates = [v for v in uniq if _DELEGATE.search(v)]
    cache: Dict[str, Tuple[List[str], List[str], List[str]]] = {}

    def insts_for(i: str) -> Tuple[List[str], List[str], List[str]]:
        """시설의 기관 하나 → (같은 발주 주체인 용역 기관들, 대행 기관으로 보이는 용역 기관들, 앞의 것 가운데 바로 그 기관인 것들). 같은 기관을 가진 시설이 많아 결과를 기억해 둔다."""
        if i not in cache:
            same = [v for v in uniq if inst_same(v, i)]
            ss = set(same)
            cand = delegates if not _DELEGATE.search(i) else uniq
            cache[i] = (same, [v for v in cand if v not in ss and inst_related(v, i)], [v for v in same if inst_narrow(v, i)])
        return cache[i]

    rows = []
    for fid, k in keys.items():
        same, related, narrow = set(), set(), set()
        for i in k["insts"]:
            sm, rl, nr = insts_for(i)
            same.update(sm)
            related.update(rl)
            narrow.update(nr)
        related -= same
        code_parts = [by_code[c] for c in k["codes"] if c in by_code]
        near = {x for v in narrow for x in by_inst[v]} | {x for p in code_parts for x in p}      # 바로 그 기관(또는 수요기관코드가 같은 기관)의 용역 공고
        parts = [by_inst[v] for v in same] + code_parts
        i_same = pd.Index(pd.unique(pd.Series([x for p in parts for x in p], dtype=object))) if parts else pd.Index([])
        rel_parts = [by_inst[v] for v in related]
        i_rel = pd.Index([x for p in rel_parts for x in p]).difference(i_same) if rel_parts else pd.Index([])
        idx = i_same.append(i_rel)
        if not len(idx):
            continue
        rel_set = set(i_rel)
        exact = pd.Series(False, index=idx)
        for p in k["pats"]:
            if len(p) <= _NAME_SHORT:                  # 짧은 이름은 띄어쓰기를 가로질러 맞추지 않는다('학용정수장 고도정수처리' 안의 '수장고')
                exact |= spaced[idx].str.contains(p, regex=False, na=False)
            else:
                exact |= t_a[idx].str.contains(p, regex=False, na=False) | t_b[idx].str.contains(p, regex=False, na=False) | t_c[idx].str.contains(p, regex=False, na=False)
        loose = pd.Series(False, index=idx)
        if not exact.any():                            # research 와 같게: 정확한 이름으로 하나도 없을 때만 느슨한 꼴을 본다
            for g in k["loose"]:
                loose |= t_a[idx].str.contains(g, regex=True, na=False)
        for i in idx[exact.to_numpy() | loose.to_numpy()]:
            ex, rel = bool(exact.at[i]), i in rel_set
            wide = not rel and i not in near
            basis = ("공고명에 시설 이름" if ex else "시설 이름의 어절이 순서대로 있음(사이에 다른 말)") \
                + (" + 그 지자체의 대행 기관으로 보이는 발주 기관" if rel else " + 같은 계열의 다른 기관(소속기관·지사·학교가 다름)" if wide else " + 같은 발주 기관")
            rows.append((i, fid, "확실" if (ex and not rel) else "추정", basis, wide))
    return pd.DataFrame(rows, columns=cols).sort_values(["_i", "시설ID"], kind="stable").reset_index(drop=True)


# 용역 공고명에서 사업 이름만 남기기 위해 지우는 말(용역 이름·공종·행정 문구)
_SVC_WORDS = re.compile(
    r"사업\s*수행\s*능력\s*평가서?(?:\s*제출)?(?:\s*안내)?|[\(（]\s*PQ\s*[\)）]|"
    r"감독\s*권한\s*대행\s*등?|통합\s*건설\s*사업\s*관리|건설\s*사업\s*관리|시공\s*단계|설계\s*단계|책임\s*감리|상주\s*감리|비상주\s*감리|전면\s*책임|감리자|감리|"
    r"기본\s*및\s*실시\s*설계|실시\s*설계|기본\s*설계|기획\s*설계|건축\s*설계|설계\s*(?:제안\s*)?공모|설계안\s*공모|현상\s*설계|설계\s*경기|제안\s*공모|국제\s*지명|지명\s*공모|일반\s*공모|설계|"
    r"(?:건설|지정|혼합|사업장)?\s*폐기물\s*(?:수집\s*[·ㆍ,및]?\s*운반\s*(?:및)?\s*)?(?:위탁\s*)?(?:처리|운반)?|석면|"
    r"엔지니어링|용역|(?<![A-Za-z])CM(?![A-Za-z])|입찰|모집|공고|계속비|총괄|장기\s*계속|긴급|재공모|재공고|재입찰|재수정|정정|수정|변경|가칭")
_SVC_TRADE = re.compile(r"정보\s*통신|통신|전기|전력\s*시설물|소방|기계\s*설비|기계|건축|토목|조경|공사|사업|분야")
_TYPE_CODE = {"신축": ("신축",), "증축": ("증축",), "리모델링": ("리모델링",), "증축·리모델링": ("증축", "리모델링")}
# 사업유형(신축·증축·리모델링·유지보수)을 볼 때 용역 이름 속의 '건설' 은 지운다 — '건설사업관리'·'건설폐기물' 의 '건설' 이 신축 어휘로 읽히면
# 유지보수 공사에 딸린 용역이 '신축' 이 되어 유지보수 거름을 빠져나가고, '어느 공사의 폐기물인지 모름' 규칙도 듣지 않는다
_WT_SVC = re.compile(r"건설\s*사업\s*관리|건설\s*(?:혼합\s*)?폐(?:기물|재류|토석|목재|자재|콘크리트|아스콘)|건설\s*오니|건설\s*(?:기술|엔지니어링)")


def project_core(name) -> str:
    """공고명 → 사업 이름만(공백·기호·용역 이름·공종어·'공사'·'사업' 제거). 공사 공고명과 용역 공고명을 같은 꼴로 만들어 견준다."""
    s = works._strip_admin(name, drop_ord=True)
    s = _SVC_WORDS.sub(" ", s)
    s = _SVC_TRADE.sub(" ", s)
    return works._PUNCT.sub("", s)


def rebid_key(name) -> str:
    """같은 용역이 다시 난 공고(재공고·재재공고·수정/정정/사전 공고·PQ 안내와 본 공고)를 한 묶음으로 보는 열쇠.
    괄호 속 차수 '(1차)'·'(2차)'·'(3차 공고)' 는 뗀다(재공고 표시일 때가 많다 — 폐기물 처리의 차수분인지는 _mark_status 가 금액으로 가린다).
    '2차분'·'N단계'·'(2025년)'·'(재)실시설계' 는 다른 용역이므로 남긴다."""
    s = _REBID_STAGE.sub(" ", _txt(name))
    s = re.sub(r"사업\s*수행\s*능력\s*평가(?:\s*서)?(?:\s*[\(（]\s*PQ\s*[\)）])?(?:\s*서류)?(?:\s*제출)?(?:\s*안내)?(?:\s*공고)?|[\(（]\s*PQ\s*[\)）]|(?:재\s*){2,}(?:공고|입찰)|재\s*공모|재\s*수정|"
               r"(?:수정|정정|변경|사전)\s*공고|계속비[_\s]*총괄|계속비|모집\s*공고", " ", s)
    return re.sub(r"공사|사업|용역", "", works.title_key(works._strip_admin(s, drop_ord=True)))


# 같은 용역의 입찰 단계를 알리는 말 — 사업수행능력평가(PQ) 안내 공고와 그 뒤의 가격입찰·지명경쟁 공고, 수의시담 공고는 한 용역이다
# (2026-11 실자료: 이 말들 때문에 이름 열쇠가 달라져 같은 설계·건설사업관리 용역이 두 번 더해진 쌍이 110개·1,292억)
_PQ = r"P\s*\.?\s*Q\s*\.?"
_REBID_STAGE = re.compile(
    r"(?:용역\s*)?(?:사업\s*)?집행\s*계획\s*및|"
    r"(?:사업\s*)?수행\s*능력\s*(?:[\(（]\s*" + _PQ + r"\s*[\)）]\s*)?평가\s*(?:서류|자료|서)?\s*(?:[\(（]\s*" + _PQ + r"\s*[\)）]\s*)?(?:서류|자료)?\s*(?:제출)?\s*(?:안내)?\s*(?:공고)?|"
    r"평가\s*자료\s*(?:[\(（]\s*" + _PQ + r"\s*[\)）]\s*)?제출\s*(?:안내)?\s*(?:공고)?|"
    r"(?<![A-Za-z])" + _PQ + r"(?![A-Za-z])\s*(?:평가서?|심사)?\s*(?:제출)?\s*(?:안내)?\s*(?:공고)?\s*,?\s*(?:후)?\s*(?:가격\s*입찰|가격|입찰|지명\s*경쟁)?|"
    r"(?:평가\s*)?후\s*가격\s*입찰|가격\s*입찰|지명\s*경쟁|지명\s*공고|[\(（\[]\s*지명\s*[\)）\]]|(?:전자\s*)?수의\s*시담(?:\s*공고)?|감독\s*권한\s*대행\s*등?")


# 용역 종류별로 '그 프로젝트의 용역일 수 있는' 공고 시점[개월]: (공사 첫 공고보다 앞으로, 공사 마지막 공고보다 뒤로)
WINDOWS = {"설계": (84, 6), "건설사업관리·감리": (48, 36), "폐기물처리": (12, 48), "기타": (84, 48)}
# 같은 이름의 프로젝트가 여럿일 때 시점으로 고르는 좁은 창[개월]: (공사 첫 공고보다 앞으로, 뒤로 — 폐기물 처리는 공사 마지막 공고 기준)
TIGHT = {"설계": (30, 0), "건설사업관리·감리": (6, 6), "폐기물처리": (1, 18)}


DESIGN_LEAD_MONTHS = 2    # 공고명에 사업 이름이 없는 설계 공고는 공사 첫 공고보다 적어도 이만큼 앞서야 그 공사의 설계로 본다
TYPE_MISMATCH_SIM = 0.8   # 공고명의 사업유형이 프로젝트와 달라도, 후보가 하나뿐이고 사업 이름 유사도가 이 이상이면 그 프로젝트로 본다
WASTE_SOLE_MONTHS = (1, 12)   # 공사 이름 없는 '건설폐기물' 공고를 그 시설의 하나뿐인 프로젝트에 더하는 시점[개월]: 공사 첫 공고보다 (앞으로, 뒤로)

# 'A 및 B … 통합 건설사업관리·감리·설계' 의 A·B(어절). 공종·용역 이름·일반어뿐인 어절('전기·통신 통합감리', '기본 및 실시설계')은 현장 이름이 아니다
_JOINT_PAIR = re.compile(r"([0-9A-Za-z가-힣◇]+)\s*(?:및|[·ㆍ・,/&])\s*([0-9A-Za-z가-힣◇]+)")
_JOINT_SKIP = re.compile(r"(?:기본|실시|설계|기획|계획|감독|권한|대행|건설|사업|관리|감리|건축|토목|조경|기계|설비|전기|전력|정보통신|정보|통신|소방|시공|단계|공사|용역|철거|해체|석면|폐기물|"
                         r"수집|운반|처리|조성|건립|신축|증축|개축|리모델링|개보수|부대|시설물?|긴급|입찰|공고|총괄|\d+차분?|[A-Za-z]{1,4})+")
_JOINT_WORK = re.compile(r"공사|사업|건립|신축|증축|개축|재건축|건설|조성|리모델링|개보수|대수선")


def _joint_partner(masked: str, names: List[str], k: dict) -> str:
    """'○○ 및 △△ … 통합 건설사업관리·감리·설계' 에서 이 프로젝트의 공사 이름(과 이 시설의 이름·기관)에 없는 상대 현장 이름을 돌려준다(없으면 '').
    masked 는 시설 이름을 '◇' 로 가린 공고명. 상대 현장은 사업 이름 자리에만 온다 — 시설 이름 앞('마루공원 및 ◇ 건립공사')이나, 시설 이름 바로 뒤에서 첫 공사 말 앞까지
    ('◇ 및 주차복합빌딩 건립사업'. 이때는 '◇' 와 짝을 이룬 어절만). 그 뒤의 어절('건립공사 전력시설물·정보통신공사 통합감리', '◇ 본관·별관 건립공사')은 공종·건물 부분의 이름이다.
    공사 이름 자체가 '○○ 체육관 및 급식소 증축공사' 이면 두 어절이 모두 공사 이름에 있으므로 묶음이 아니다. '통합' 이라는 말만으로도(건축·전기·통신 통합 감리) 묶음이 아니다."""
    t = _REBID_STAGE.sub(" ", _txt(masked))
    m = re.search(r"통합", t)
    at = t.find("◇")
    if not m or at < 0 or at > m.start():
        return ""
    before, after = t[:at], t[at + 1:m.start()]
    w = _JOINT_WORK.search(after)
    after = after[:w.start()] if w else after
    own = "".join(names) + "".join(_ns(x) for x in k["pats"]) + "".join(inst_key(i) for i in k["insts"])
    pairs = _JOINT_PAIR.findall(before + "◇") + [p for p in _JOINT_PAIR.findall("◇" + after) if "◇" in p[0]]
    for pair in pairs:
        for tok in pair:
            tok = tok.replace("◇", "")
            if len(tok) >= 2 and not _JOINT_SKIP.fullmatch(tok) and tok not in own:
                return tok
    return ""


def _months(a, b) -> Optional[float]:
    if a is None or b is None or pd.isna(a) or pd.isna(b):
        return None
    return (a - b).days / 30.44


def _nearest(kind: str, dt, cands: List[dict]) -> Optional[dict]:
    """같은 이름의 프로젝트가 여럿일 때(같은 이름으로 다시 나는 공사) 공고 시점으로 고른다 — 좁은 창(TIGHT)에 드는 프로젝트가 꼭 하나일 때만.
    설계는 공사보다 먼저, 폐기물 처리는 공사 중·뒤에, 건설사업관리·감리는 공사 첫 공고 무렵에 난다. 해마다 같은 이름으로 나는 공사는 창에 둘이 들어 정하지 않는다."""
    before, after = TIGHT.get(kind, (0, 0))
    hit = []
    for p in cands:
        m_first, m_last = _months(dt, p["first"]), _months(dt, p["last"])
        if m_first is None:
            return None
        if -before <= m_first and (m_last if (kind == "폐기물처리" and m_last is not None) else m_first) <= after:
            hit.append(p)
    return hit[0] if len(hit) == 1 else None


def assign_project(kind: str, title: str, dt, projs: List[dict], wtype: str = "", pats: Tuple[str, ...] = ()) -> Tuple[str, str, List[str], float]:
    """한 시설의 프로젝트 목록 가운데 이 용역이 속한 프로젝트 → (프로젝트ID 또는 '', 까닭, 후보 프로젝트ID들, 근거의 세기).
    근거의 세기: 3.x = 공고명에 그 프로젝트의 사업 이름이 통째로 있음(이름이 길수록 큼) / 2 = 시기가 맞는 프로젝트가 하나뿐 / 1 = 이름이 가장 비슷함 / 0 = 정하지 못함.
    ① 시기가 맞는 프로젝트만 후보(종류별 창 WINDOWS)
    ② 후보의 사업 이름(공사 공고명에서 공종·'공사' 를 뺀 것. 시설 이름뿐인 것은 빼고 넘겨받는다)이 용역 공고명에 통째로 들어 있으면 그 프로젝트.
       여럿이면 가장 긴 이름('○○ 건립' 과 '○○ 건립 경관조명 조성' 이 둘 다 들어 있으면 뒤의 것 — 더 구체적인 사업), 그래도 같으면 사업유형
       (wtype: 공고명으로 본 신축·증축·리모델링), 좁은 창(TIGHT)에 드는 하나
    ③ 통째로 든 것이 없으면: 공고명이 사업유형을 밝혔으면 그 유형의 프로젝트만('○○회관 리모델링 설계용역' 을 신축 프로젝트에 붙이지 않는다).
       남은 후보가 하나면 그것, 여럿이면 좁은 창에 드는 프로젝트가 꼭 하나일 때 그것
    ④ 그래도 여럿이면 가장 비슷한 이름 — 차이가 뚜렷할 때만. 그 밖에는 정하지 않는다(05 에 더하지 않는다)."""
    before, after = WINDOWS.get(kind, WINDOWS["기타"])
    cands = []
    for p in projs:
        m_first, m_last = _months(dt, p["first"]), _months(dt, p["last"])
        if m_first is None:
            cands.append(p)                            # 날짜를 모르면 시기로 거르지 않는다
        elif -before <= m_first and (m_last if m_last is not None else m_first) <= after:
            cands.append(p)
    if not cands:
        return "", "공고 시점이 맞는 프로젝트가 없음(이 시설의 다른 사업일 수 있음)", [], 0
    ids = [p["pid"] for p in cands]
    want = _TYPE_CODE.get(wtype, ())
    fits = lambda p: bool(want) and any(w in p["type"] for w in want)
    core = project_core(title)
    inside = []
    for p in cands:
        ks = [k for k in p["cores"] if k and len(k) >= 4 and k in core]
        if ks:
            inside.append((max(len(k) for k in ks), p))
    if inside:
        inside.sort(key=lambda x: -x[0])
        top = [p for n, p in inside if n == inside[0][0]]
        lvl = 3 + inside[0][0] / 1000
        if len(top) == 1:
            return top[0]["pid"], "공고명에 그 프로젝트의 사업 이름이 들어 있음", ids, lvl
        typed = [p for p in top if fits(p)]
        if len(typed) == 1:
            return typed[0]["pid"], "공고명에 사업 이름이 들어 있고 사업유형(" + wtype + ")이 맞는 프로젝트", ids, lvl
        near = _nearest(kind, dt, typed or top)
        if near is not None:
            return near["pid"], f"같은 사업 이름의 프로젝트 {len(top)}개 가운데 공고 시점이 맞는 하나", ids, lvl
        return "", f"같은 사업 이름의 프로젝트 {len(top)}개 가운데 가릴 수 없음", ids, 0
    if want:                                          # 공고명이 사업유형을 밝혔으면 그 유형의 프로젝트만 본다
        typed = [p for p in cands if fits(p)]
        if not typed:
            # 공고명의 사업유형 말('건립'·'조성')은 공사 쪽 분류(증축·리모델링)와 어긋나기 쉽다('○○ 확장 건립사업' 은 증축 프로젝트, 그 감리 공고는 '○○ 건립사업 … 감리').
            # 시기가 맞는 프로젝트가 하나뿐이고 사업 이름이 거의 같으면 그 프로젝트로 본다(2026-11 실자료: 유사도 0.80 이상인 줄은 모두 그 프로젝트의 용역, 0.79 이하는 다른 사업)
            # 다만 사업유형 말만 다른 이름('○○ 건립공사 감리' ↔ 프로젝트 '○○ 증축공사')은 다른 사업이다(수집 기간 전에 지은 건물의 용역일 수 있다) —
            # 공고명의 사업유형 말(건립·신축·조성 …)이 그 프로젝트의 사업 이름에도 들어 있을 때만('○○ 확장 건립사업')
            if len(cands) == 1 and core:
                verbs = set(works._VERB.findall(core))
                s, kk = max([(works.similarity(core, k), k) for k in cands[0]["cores"] if k] or [(0.0, "")])
                if s >= TYPE_MISMATCH_SIM and any(v in kk for v in verbs):
                    return cands[0]["pid"], f"공고명의 사업유형({wtype})은 다르지만 사업 이름이 거의 같은 프로젝트 하나(유사도 {s:.2f})", ids, 1
            return "", f"공고명의 사업유형({wtype})에 맞는 프로젝트가 없음", ids, 0
        cands = typed
    if len(cands) == 1:
        return cands[0]["pid"], ("시기가 맞는 프로젝트 하나" if len(projs) > 1 else "시설의 프로젝트 하나"), ids, 2
    near = _nearest(kind, dt, cands)                  # 여럿이면 좁은 창(TIGHT)에 드는 프로젝트가 꼭 하나일 때 그것
    if near is not None:
        return near["pid"], f"프로젝트 {len(cands)}개 가운데 공고 시점이 맞는 하나", ids, 2
    # 유사도는 양쪽에 사업을 가리는 말이 있을 때만 쓴다. 공고명에 시설 이름뿐이면 점수는 프로젝트 이름에 덧붙은 말의 길이로 정해지고('○○ 건립' 0.80 > '○○ 객석 개보수' 0.62),
    # 사업 이름이 시설 이름뿐이라 cores 가 비워진 프로젝트는 0 점이 되어 공고명이 그 공사 이름과 글자 그대로 같아도 진다
    rest = core
    for q in sorted((works._PUNCT.sub("", x) for x in pats), key=len, reverse=True):
        rest = rest.replace(q, "")
    if (pats and len(rest) < 2) or any(not p["cores"] for p in cands):
        return "", f"프로젝트 {len(cands)}개 가운데 가릴 수 없음(공고명이나 프로젝트 이름에 사업을 가리는 말이 없음)", ids, 0
    scored = sorted(((max([works.similarity(core, k) for k in p["cores"] if k] or [0.0]) if core else 0.0, p) for p in cands), key=lambda x: -x[0])
    if scored[0][0] >= 0.75 and scored[0][0] - scored[1][0] >= 0.15:
        return scored[0][1]["pid"], f"공고명이 가장 가까운 프로젝트(유사도 {scored[0][0]:.2f}, 다음 {scored[1][0]:.2f})", ids, 1
    return "", f"프로젝트 {len(cands)}개 가운데 가릴 수 없음", ids, 0


_PHASE_WORD = re.compile(r"\d+|단계|공구|차분|차년도|블록|[A-Za-z]")


def _names_this_project(p: Optional[dict], title: str, k: dict) -> bool:
    """(같은 계열의 다른 기관이 낸 공고에서) 공고명에 든 사업 이름이 바로 이 프로젝트의 것이라고 볼 수 있는가.
    ① 그 사업 이름의 앞뒤에, 이 시설의 이름·기관에 없는 다른 현장 말이 없다 — '세경지구 수출전문 스마트팜 온실 신축공사' 의 '세경지구' 는 다른 지사의 같은 이름 사업이라는 표시다
    ② 검수한 시설 이름이 공고명에 그대로 있거나, 사업 이름에 시설 이름·일반어(건립·신축 …) 말고도 현장을 가리는 말이 있다 —
       '스마트팜 온실 신축'(시설 '○○교육청 스마트팜온실')뿐이면 다른 학교의 같은 이름 공사와 가릴 수 없다"""
    if p is None:
        return False
    core = project_core(title)
    best = max((c for c in p["cores"] if len(c) >= 4 and c in core), key=len, default="")
    if not best:
        return False
    names = sorted({works._PUNCT.sub("", x) for x in k["pats"]}, key=len, reverse=True)
    own = "".join(names) + "".join(inst_key(i) for i in k["insts"])
    for frag in core.split(best):
        frag = _PHASE_WORD.sub("", works._GENERIC.sub("", frag))
        if len(frag) >= 2 and frag not in own:
            return False
    flat = _ns(_BR_CHARS.sub("", title))
    if any(_BR_CHARS.sub("", x) in flat for x in k.get("full", ())):
        return True
    rest = best
    for q in names:
        rest = rest.replace(q, "")
    return len(works._GENERIC.sub("", rest)) >= 2


# 사업 이름 앞뒤에 붙어도 '다른 현장' 을 뜻하지 않는 말: 폐기물 종류·처리 방법, 감리 방식, 입찰 문구
_SITE_NOISE = re.compile(r"파쇄|소각|매립|처리|운반|수집|수거|가연성|불연성|혼합|임목|지정|폐[가-힣]{1,3}|추가분?|총괄분?|긴급|비상주|상주|책임|입찰|안내|제출|견적|수의|소액|시행\s*구역|구역|구간|공모|당선작?|제안|통합")
_REGION_HEAD = re.compile(r"[가-힣0-9]{1,4}(?:읍|면|동|(?<!거)리)")        # 읍·면·동·리 이름 하나(위치 표시). '○○시장'·'○○지구'·'원도심'·'문화의거리' 는 현장 이름이다
# 시설 종류 앞에 흔히 붙는 꾸밈말 — 이것뿐이면 현장을 가리는 말이 아니다('공영 주차타워 건립' 은 '주차타워 건립' 과 다를 바 없다)
_SITE_QUAL = re.compile(r"공영|공공|공용|공립|시립|군립|구립|도립|국립|시민|군민|구민|도민|종합|복합|통합|다목적|제?\d+")


def _only_this_site(masked: str, k: dict) -> bool:
    """(시설 이름을 '◇' 로 가린) 폐기물 공고명에서 폐기물 말 앞에 이 시설 말고 다른 현장·다른 공사를 가리키는 말이 없는가 — '◇ 건설폐기물 처리용역'.
    '◇ 및 △△ 건설폐기물'·'◇ 외 2개소'·'◇ 등'(묶음), '◇ 별관'·'◇ 사무실 칸막이 공사'(다른 공사)는 아니다."""
    m = _K_WASTE.search(masked)
    head = masked[:m.start()] if m else masked
    if "◇" not in head or re.search(r"및|[,·ㆍ・/&]|(?<![가-힣])(?:등|외)(?![가-힣])", head):
        return False
    own = "".join(works._PUNCT.sub("", x) for x in k["pats"]) + "".join(inst_key(i) for i in k["insts"])
    rest = _PHASE_WORD.sub("", _SITE_NOISE.sub("", works._GENERIC.sub("", project_core(head.replace("◇", " ")))))
    return len(rest) < 2 or rest in own


def _site_fit(p: Optional[dict], title: str, k: dict) -> str:
    """(짧은 일반 이름으로 이어진 공고에서) 공고명에 든 사업 이름이 이 프로젝트를 얼마나 가리는가.
    'foreign'  사업 이름 앞에 이 시설·기관에 없는 다른 현장 말이 있거나('가온역 환승 주차타워 건립공사' ↔ '주차타워 건립'),
               뒤에 다른 대상 + 건설·건립 말이 이어진다('○○지구 3지구 9단지 아파트 건설공사' ↔ '○○지구 3지구 (조경공사)') — 다른 사업
    'specific' 사업 이름에 시설 이름·일반어(건립·신축 …) 말고도 현장을 가리는 말이 있다('나래지구 3단지 공공주택 건설') — 시점과 무관하게 이 사업
    'plain'    사업 이름이 시설 종류 + 건립·신축뿐('주차타워 건립')이거나, 통째로 있지는 않지만 그 사업의 현장 말이 공고명에 그대로 있다 — 시점까지 맞아야 한다
    ''         공고명에 사업 이름도, 그 사업의 현장을 가리는 말도 없음"""
    if p is None:
        return ""
    core = project_core(_REBID_STAGE.sub(" ", _txt(title)))      # 입찰 단계 말('집행계획 및 … 평가서 제출 안내')은 현장 이름이 아니다
    best = max((c for c in p["cores"] if len(c) >= 4 and c in core), key=len, default="")
    names = sorted({works._PUNCT.sub("", x) for x in k["pats"]}, key=len, reverse=True)
    own = "".join(names) + "".join(inst_key(i) for i in k["insts"])
    bare = lambda s: _PHASE_WORD.sub("", _SITE_NOISE.sub("", works._GENERIC.sub("", s)))      # 일반어·용역 수식어·차수를 뺀 나머지

    def site_words(s: str) -> str:                 # 시설 이름·일반어(건립·신축·공사 …)를 뺀 나머지 = 현장을 가리는 말
        for q in names:
            s = s.replace(q, "")
        return works._GENERIC.sub("", s)
    if not best:
        # 사업 이름이 통째로 있지는 않아도('○○현충원 봉안당 설계용역' ↔ 공사 '○○현충원 봉안당 신축공사'), 그 사업의 현장을 가리는 말('○○현충원')이 공고명에 그대로 있으면 'plain' 과 같이 본다
        # 지자체 이름('○○군')은 현장을 가리는 말이 아니다(그 군의 어느 사업에나 붙는다). 부처·공사·교육청의 소속기관 이름('국립○○현충원')은 현장 말로 친다
        region = "".join(names) + "".join(inst_key(i) if inst_sig(i)[0] == "gov" else inst_key(_txt(i).split()[0] if _txt(i).split() else "") for i in k["insts"])
        mine = site_words(core)
        if len(bare(mine)) < 2:
            return "plain"                         # 공고명이 '유리온실 신축 설계용역' 처럼 시설 종류 + 건립·신축뿐(다른 현장을 가리키는 말이 없다) — 시점이 맞으면 이 사업
        for c in p["cores"]:
            w = site_words(c)
            if len(_PHASE_WORD.sub("", w)) >= 2 and w not in region and w in mine:
                return "plain"
        return ""
    at = core.find(best)
    head = bare(core[:at])
    if _REGION_HEAD.fullmatch(head):                       # 읍·면·동 이름('임피면 승화원 …')은 다른 현장이 아니라 위치 표시다
        head = ""
    if len(head) >= 2 and head not in own:
        return "foreign"
    # 사업 이름 뒤에 이 프로젝트의 공사 이름에 없는 다른 대상이 이어지면('○○지구 (A1) 공동주택 전기 설계' ↔ 사업 이름 '○○지구') 그 지구의 다른 사업이다
    tail = bare(core[at + len(best):])
    if len(tail) >= 2 and tail not in own and tail not in "".join(p.get("names", ())):
        return "foreign"
    rest = best
    for q in names:
        rest = rest.replace(q, "")
    return "specific" if len(_SITE_QUAL.sub("", works._GENERIC.sub("", rest))) >= 2 else "plain"


_K_PREP_NOTICE = re.compile(r"철거|해체|석면|지장물")      # 본 공사에 앞서 나는 준비 공사 공고(설계 공고 시점을 재는 기준에서 뺀다)


def project_table(latest: pd.DataFrame) -> Dict[str, List[dict]]:
    """시설ID → [{"pid", "type", "first", "last", "cores", "amt", "keys"} …] (대표 공고 기준)."""
    out: Dict[str, List[dict]] = {}
    if latest.empty:
        return out
    lt = latest.assign(_dt=_parse_dt(latest["공고일시"]), _amt=pd.to_numeric(latest["추정가격"], errors="coerce").fillna(0))
    for pid, g in lt.groupby("프로젝트ID", sort=False):
        names = [n for n in g["공고명"].map(_txt) if n]
        if "세부사업명" in g.columns and _txt(g["세부사업명"].iloc[0]):
            names.append(_txt(g["세부사업명"].iloc[0]))
        cores = [c for c in dict.fromkeys(project_core(n) for n in names) if len(c) >= 4]
        out.setdefault(str(g["시설ID"].iloc[0]), []).append({
            "pid": str(pid), "type": _txt(g["사업유형"].iloc[0]), "first": g["_dt"].min(), "last": g["_dt"].max(), "cores": cores,
            "first_main": (lambda d: None if pd.isna(d) else d)(g.loc[~g["공고명"].map(_txt).str.contains(_K_PREP_NOTICE), "_dt"].min()),
            "amt": float(g["_amt"].sum()), "keys": set(g["공고번호"].map(_no)), "names": [_ns(n) for n in names]})
    return out


def link(raw: pd.DataFrame, reviewed: pd.DataFrame, hist: pd.DataFrame, latest: pd.DataFrame, cfg: dict) -> pd.DataFrame:
    """용역 공고 원문 → 연결 표(LINK_COLS). 한 줄 = 용역 공고(최신 차수) × 시설."""
    from .classify import work_type_for
    sv = latest_orders(standardize(raw, cfg))
    keys = facility_keys(reviewed, hist)
    m = match_facilities(sv, keys)
    if m.empty:
        return pd.DataFrame(columns=LINK_COLS)
    projs = project_table(latest)
    named: Dict[str, List[dict]] = {}

    def projects_of(fid: str) -> List[dict]:
        """그 시설의 프로젝트 목록. 사업 이름 가운데 시설 이름과 다를 바 없는 것('○○미술관 전기공사' → '○○미술관')은 뺀다 —
        그런 이름은 그 시설의 어느 용역 공고명에나 들어 있어 프로젝트를 가리는 근거가 못 된다."""
        if fid not in named:
            # cores 는 기호를 뺀 꼴(project_core)이므로 시설 이름도 같은 꼴로 견준다: '의령4·26추모공원' ↔ '의령426추모공원', 'The-K호텔' ↔ 'TheK호텔'
            pats = sorted({works._PUNCT.sub("", p) for p in keys[fid]["pats"]}, key=len, reverse=True)

            def specific(core: str) -> bool:
                rest = core
                for p in pats:
                    rest = rest.replace(p, "")
                return len(rest) >= 2
            named[fid] = [dict(p, cores=[c for c in p["cores"] if specific(c)]) for p in projs.get(fid, [])]
        return named[fid]

    rows = []
    for i, fid, strength, basis, wide in zip(m["_i"], m["시설ID"], m["연결강도"], m["연결근거"], m["_wide"]):
        r, k = sv.loc[i], keys[fid]
        title = _txt(r["공고명"])
        masked = _mask_name(title, k["pats"])
        # 조달분류는 가장 구체적인 이름만 본다 — 대분류 '폐기물 처리 및 재활용서비스' 의 '재활용' 때문에 건설폐기물 처리 용역이 운영 폐기물로 읽히지 않게
        kind, sub, why = classify_kind(masked, r["조달분류"] or r["조달분류_중"])
        wtype = work_type_for(_WT_SVC.sub(" ", title), k["keyword"], k["name"])
        pl = projects_of(fid)
        pid, why_p, cand, lvl = assign_project(kind, title, r["_dt"], pl, wtype, tuple(k["pats"]))
        p = next((x for x in pl if x["pid"] == pid), None)
        if why.startswith(_UNIT_WHY) and lvl >= 3 and p is not None and _unit_words_from(masked, p["names"]):
            why = ""              # '1구역'·'상반기'·'(연간단가)' 가 그 공사의 이름에 든 말일 뿐 — 용역이 여러 사업을 묶은 계약이라는 뜻이 아니다
        if not why and BUNDLE.search(title):
            why = "여러 시설·여러 건을 한 공고로 묶은 용역 — 금액을 나눌 수 없음"
        if not why and kind in KINDS and p is not None:
            other = _joint_partner(masked, p["names"], k)
            if other:             # '○○ 및 △△ … 통합 건설사업관리': 묶인 상대 현장이 DB 에 없어도(한 시설에만 이어져도) 묶음 발주다
                why = f"여러 현장을 묶은 통합 용역(묶음 발주)으로 보임(공고명의 '{other}' 이(가) 이 프로젝트의 공사 이름에 없음) — 금액을 나눌 수 없음"
        price = r["추정가격"]
        if not why and kind in KINDS and p is not None and p["amt"] > 0 and price is not None and not pd.isna(price) and abs(float(price) - p["amt"]) <= p["amt"] * 0.001:
            why = "용역 추정가격이 이 프로젝트의 공사 추정가격 합과 같음 — 공사비를 용역 금액 칸에 적은 공고로 보임(공고문 확인)"
        flat = _ns(_BR_CHARS.sub("", title))
        hit = lambda names: [x for x in names if _BR_CHARS.sub("", x) in flat]
        found = max((len(x) for x in hit(k["pats"])), default=99)                                  # 느슨한 꼴로만 맞은 줄은 어차피 '추정'
        # '추정' 연결(대행 기관·느슨한 이름)을 더해도 되는가: 공고명에 든 사업 이름이 이 현장을 가리킬 때만 —
        # 현장을 가리는 말이 든 사업 이름('○○혁신센터 공영주차장 건립')이 통째로 있거나, 검수한 시설 이름이 그대로 있고 사업 이름 앞뒤에 다른 현장 말이 없을 때.
        # 접두를 뗀 일반 이름 + '건립' 뿐인 사업 이름('국민체육센터 건립')은 광역 건설본부가 낸 다른 자치구의 같은 이름 사업에도 들어 있다
        is_named = False
        if strength != "확실" and lvl >= 3 and p is not None:
            fit0 = _site_fit(p, title, k)
            is_named = fit0 == "specific" or (fit0 == "plain" and bool(hit(k.get("full", ()))))
        if found <= _NAME_GENERIC and not hit(k.get("full", ())) and p is not None:
            # 짧은 일반 이름('공공주택'·'주차타워')은 같은 기관의 어느 사업 공고명에나 들어 있다 → 시점이 아니라 사업 이름으로 가린다:
            # 공고명에 그 프로젝트의 사업 이름이 통째로 있으면(lvl 3) 공사 공고와 멀어도 더하고, 없으면 시점이 맞아도 더하지 않는다
            # (2026-11 실자료: 사업 이름 없이 시점만으로 더한 17줄 가운데 10줄이 다른 단지·다른 사업 / 사업 이름이 있는데 시점이 멀다고 뺀 13줄 가운데 11줄은 맞는 용역)
            why_p = f"일반적인 시설 이름({found}글자)으로 이어짐 — " + why_p
            if not why and kind in KINDS and not wide:                  # 같은 계열의 다른 기관이 낸 공고는 아래에서 그 까닭으로 뺀다
                fit = _site_fit(p, title, k)
                if not fit:
                    why = _GENERIC_WHY + " 공고명에 그 프로젝트의 사업 이름이 없음 — 같은 기관의 다른 사업일 수 있음(공고명 확인)"
                elif fit == "foreign":                                  # '○○역 환승 주차타워 건립공사' ↔ 사업 이름 '주차타워 건립', '3지구 9단지 아파트 건설공사' ↔ '3지구 (조경)'
                    why = _GENERIC_WHY + " 공고명의 사업 이름 앞뒤에 다른 현장·다른 공사의 이름이 있음 — 같은 기관의 다른 사업일 수 있음(공고명 확인)"
                elif fit == "plain" and _nearest(kind, r["_dt"], [p]) is None:      # 사업 이름이 '주차타워 건립' 뿐이면 시점까지 맞아야 한다
                    why = _GENERIC_WHY + " 사업 이름도 시설 종류 + 건립·신축뿐인데 공고 시점이 그 공사와 가깝지 않음 — 같은 기관의 다른 사업일 수 있음(공고명 확인)"
        # 아래 세 가지(유지보수·폐기물 출처·건축 아님)는 공고명에 DB 프로젝트의 사업 이름이 통째로 들어 있으면(lvl 3) 보지 않는다 — 그 사업은 이미 DB 에 실린 공사다.
        # '같은 계열의 다른 기관' 은 lvl 3 이어도 그 사업 이름이 바로 이 프로젝트의 것인지 한 번 더 본다(_names_this_project)
        # (시설 사업에 딸린 '관로 연결공사'·'진입도로 개설공사' 처럼 토목 공종 공고도 프로젝트에 들어 있다)
        if not why and lvl < 3 and kind in KINDS and wtype == "유지보수":
            why = "유지보수 성격의 공사에 딸린 용역으로 보임(이 DB 는 신축·증축·리모델링 공사만 싣는다)"
        if not why and lvl < 3 and kind == "폐기물처리" and wtype not in _TYPE_CODE:
            # 공사 이름은 없어도 공고명이 '건설폐기물'·'폐아스콘' 이라고 밝혔고, 그 시설의 프로젝트가 하나뿐이며, 그 공사 공고 무렵(1개월 전~12개월 뒤)에 났으면 그 공사의 폐기물로 본다
            # (2026-11 실자료: 이 까닭으로 뺀 줄의 22% 가 그 공사의 폐기물이었고, 이 조건에 드는 줄은 대부분 맞았다)
            lead = _months(r["_dt"], p["first"]) if p is not None else None
            sole = (p is not None and len(pl) == 1 and _K_WASTE_CONSTR.search(masked) and lead is not None and -WASTE_SOLE_MONTHS[0] <= lead <= WASTE_SOLE_MONTHS[1]
                    and _only_this_site(masked, k))
            if not sole:
                why = "어느 공사에서 나온 폐기물인지 공고명으로 알 수 없음(공사 이름·신축·증축·리모델링 말이 없음 — 시설 운영·유지관리 폐기물일 수 있음)"
        if not why and lvl < 3 and kind in KINDS and _K_NOT_BUILDING.search(masked):
            why = "건축 사업의 용역이 아닌 것으로 보임(산림·하천·상하수도·도로·국가유산 등)"
        if not why and wide and kind in KINDS:
            if lvl < 3:
                why = _WIDE_WHY + "이고 공고명에 이 프로젝트의 사업 이름이 통째로 있지 않음 — 다른 기관의 같은 이름 시설일 수 있음"
            elif not _names_this_project(p, title, k):
                # 사업 이름이 통째로 있어도: 그 앞뒤에 다른 현장 이름이 있거나('○○지구 스마트팜 온실 신축공사'), 사업 이름이 '접두를 뗀 시설 이름 + 건립·신축' 뿐이면 가릴 수 없다
                why = (_WIDE_WHY + "이고 공고명의 사업 이름이 이 프로젝트의 것인지 가릴 수 없음(다른 현장 이름이 함께 있거나, 사업 이름이 시설 종류 + 건립·신축뿐) "
                       "— 다른 기관의 같은 이름 시설일 수 있음")
        if not why and kind == "설계" and lvl < 3 and p is not None:
            # 설계 용역 공고는 공사 공고보다 여러 달 앞선다. 공고명에 사업 이름도 없는 설계 공고가 공사 공고의 직전·뒤에 났으면 그 공사의 설계가 아니다
            # (2026-11 실자료: 이런 Y 15줄 가운데 9줄이 같은 시설의 다음 단계·다른 사업의 설계였다). 사업 이름이 통째로 있으면(철거 공고가 먼저 난 사업 등) 그대로 더한다
            lead = _months(r["_dt"], p.get("first_main"))      # 철거·해체·지장물 공고를 뺀 첫 공사 공고(그런 공고뿐인 프로젝트는 이 규칙을 보지 않는다 — 본 공사는 아직 나지 않았다)
            if lead is not None and lead > -DESIGN_LEAD_MONTHS:
                why = (f"설계 공고가 그 공사 공고의 직전이거나 뒤({lead:+.0f}개월)이고 공고명에 그 프로젝트의 사업 이름이 없음 "
                       "— 그 공사의 설계로 보기 어려움(다음 단계·다른 사업의 설계일 수 있음)")
        ref = p or (min(pl, key=lambda x: (lambda m: 1e9 if m is None else abs(m))(_months(r["_dt"], x["first"]))) if pl else None)
        gap = _months(r["_dt"], ref["first"]) if ref is not None else None
        rows.append({
            "프로젝트ID": pid, "시설ID": fid, "시설명": k["name"], "용역종류": kind, "용역세부": sub, "합산여부": "", "합산제외_까닭": why,
            "연결강도": strength, "연결근거": basis + " / " + why_p, "공고번호": r["공고번호"], "공고차수": r["공고차수"], "공고명": title,
            "공고일시": r["공고일시"], "공고기관": r["공고기관"], "수요기관": r["수요기관"], "추정가격": price, "배정예산": r["배정예산"], "상태": "",
            "공고종류": r["공고종류"], "용역구분": r["용역구분"], "조달분류": r["조달분류"], "프로젝트_후보": ", ".join(cand) if not pid else "",
            "공사_최초공고일": (str(ref["first"])[:10] if ref is not None and not pd.isna(ref["first"]) else ""),
            "공사대비_개월": round(gap, 1) if gap is not None else None,
            "공사비대비_비율": (round(float(price) / p["amt"], 4) if (p and p["amt"] > 0 and price is not None and not pd.isna(price)) else None),
            "상세URL": r["상세URL"], "_dt": r["_dt"], "_prev": r["이전공고번호"], "_lvl": lvl, "_named": is_named, "_masked": masked, "_alien": why.startswith((_WIDE_WHY, _GENERIC_WHY)), "_spans": _name_spans(title, k["pats"], tuple(k["loose"]))})
    out = pd.DataFrame(rows)
    out = _mark_status(out)
    out = _mark_sum(out, projs)
    out = out.sort_values(["시설ID", "프로젝트ID", "용역종류", "_dt", "공고번호"], kind="stable").drop(columns=["_dt", "_prev", "_lvl", "_spans", "_alien", "_named", "_masked"])
    return out[LINK_COLS].reset_index(drop=True)


REBID_NEAR_DAYS = 120     # 같은 이름으로 이 기간 안에 다시 난 공고는 재공고(유찰·정정 뒤 다시 냄). 폐기물 처리의 차수분(차수 표기와 금액이 모두 다름)은 예외
REBID_MAX_DAYS = 550      # 그보다 떨어져도 이 기간 안이고 금액이 거의 같거나(±15%) 뒤 공고가 '재공고' 로 표시돼 있거나 같은 프로젝트의 설계·건설사업관리이면 재공고. 그 밖에는 다음 회차·다음 단계의 다른 용역
REBID_AMT_BAND = 0.15
REBID_SAME_AMT = 0.01     # '같은 금액' 으로 보는 차이(같은 프로젝트의 설계·건설사업관리/감리가 이 안의 금액으로 REBID_NEAR_DAYS 안에 다시 났으면 이름이 달라도 같은 용역)
_DUP_TRADE = re.compile(r"전기|통신|소방|기계|건축")
_DUP_SPECIAL = re.compile(r"토목|조경|무대|전시|석면|해체|철거")
_DUP_PHASE = re.compile(r"\d+(?:\s*-\s*\d+)?\s*(?:단계|차분|차년도|차|공구|블록|BL|단지|지구|구역|권역|호기|관|동)|[A-Za-z]\s*-?\s*\d+\s*(?:BL|블록)?")


def _other_service(a: str, b: str) -> bool:
    """두 용역 공고명이 서로 다른 용역임을 보이는 표시가 있는가 — 공종('전기공사 감리' ↔ '소방공사 감리')이나 단계·차수·블록('1단계' ↔ '2단계')이 양쪽에 적혀 있고 겹치지 않을 때.
    한쪽에만 적혀 있으면('… 감리용역' ↔ '… 감리용역(건축)') 다르다고 보지 않는다. 다만 따로 발주되는 일('무대장치 감리'·'석면 해체 감리'·'조경 설계')은 한쪽에만 적혀 있어도 다른 용역이다."""
    for pat in (_DUP_TRADE, _DUP_PHASE):
        x, y = {_ns(m) for m in pat.findall(a)}, {_ns(m) for m in pat.findall(b)}
        if x and y and not (x & y):
            return True
    return set(_DUP_SPECIAL.findall(a)) != set(_DUP_SPECIAL.findall(b))


_ORD_REBID = re.compile(r"\d+\s*차\s*(?:재\s*)?(?:공고|입찰)")                    # '(2차 공고)' — 다시 낸 공고라는 표시
_ORD_MARK = re.compile(r"(\d+)\s*차(?!\s*(?:재\s*)?(?:공고|입찰))")          # '(2차)'·'2차분' — '2차 공고'·'2차 입찰' 은 다시 낸 공고라는 표시이므로 뺀다


def _mark_status(out: pd.DataFrame) -> pd.DataFrame:
    """상태: '취소' / '재공고 이전'(같은 시설·같은 종류에서 같은 이름으로 다시 난 공고가 있는 앞 공고) / '대표'.
    유찰 뒤 재공고는 공고번호가 새로 나오므로 그대로 더하면 같은 용역비가 두 번 들어간다. 사업수행능력평가(PQ) 안내 공고와 본 공고도 같은 묶음으로 본다.
    해마다·단계마다 같은 이름으로 다시 나는 용역(폐기물 처리 1차·2차 연도)은 간격이 길고 금액이 달라 재공고로 보지 않는다 — 둘 다 더한다."""
    out = out.copy()
    out["상태"] = ["취소" if k == "취소" else "대표" for k in out["공고종류"]]
    out["_key"] = [rebid_key(n) for n in out["공고명"]]
    live = out[out["상태"] == "대표"]
    if "_alien" in out.columns:
        # 다른 기관·다른 사업의 공고일 수 있어 더하지 않는 줄(같은 계열의 다른 기관, 일반 이름 + 먼 시점)은 '같은 용역이 다시 난 공고' 판정에 넣지 않는다 —
        # 다른 학교가 같은 이름으로 낸 공고가 제 학교의 공고를 '재공고 이전' 으로 만들지 않게
        live = live[~live["_alien"].astype(bool)]
    replaced = set()
    nos = set(zip(live["시설ID"], live["공고번호"]))
    for f, n, p in zip(live["시설ID"], live["공고번호"], live["_prev"]):      # 이전공고번호가 가리키는 공고(같은 시설에 연결돼 있을 때)
        if p and p != n and (f, p) in nos:
            replaced.add((f, p))
    price = pd.to_numeric(out["추정가격"], errors="coerce")
    for _, g in live.groupby(["시설ID", "용역종류", "_key"], sort=False):
        if len(g) < 2:
            continue
        g = g.sort_values(["_dt", "공고번호"], kind="stable")
        idx = list(g.index)
        waste = out.at[idx[0], "용역종류"] == "폐기물처리"
        marks = {i: tuple(_ORD_MARK.findall(_txt(out.at[i, "공고명"]))) for i in idx}
        pairs = []
        for j in range(1, len(idx)):                   # 바로 앞 공고와 견준다. 폐기물 처리는 차수분 사이에 재공고가 끼어 날 수 있으므로 같은 차수 표기의 앞 공고가 따로 있으면 그것과도 견준다
            own = [i for i in idx[:j] if marks[i] == marks[idx[j]]] if waste else []
            if own and own[-1] != idx[j - 1]:
                pairs += [(own[-1], idx[j], False), (idx[j - 1], idx[j], True)]
            else:
                pairs.append((idx[j - 1], idx[j], False))
        for a, b, has_own in pairs:                    # 뒤 공고가 같은 이름으로 다시 난 것이면 앞의 것은 재공고 이전
            da, db = out.at[a, "_dt"], out.at[b, "_dt"]
            gap = (db - da).days if not (pd.isna(da) or pd.isna(db)) else 0
            pa, pb = price.at[a], price.at[b]
            same_amt = bool(pa > 0 and pb > 0 and abs(pb - pa) <= max(pa, pb) * REBID_AMT_BAND) or not (pa > 0 and pb > 0)
            # 설계·건설사업관리는 한 사업에 한 번 하는 용역 — 같은 프로젝트에 같은 이름으로 다시 났으면 금액이 달라져도(예산을 고쳐 다시 낸 공고) 같은 용역이다.
            # 폐기물 처리는 차수마다 따로 나므로 금액을 본다
            one_off = out.at[a, "용역종류"] in ("설계", "건설사업관리·감리") and out.at[a, "프로젝트ID"] == out.at[b, "프로젝트ID"] != ""
            again = out.at[b, "공고종류"] == "재공고" and not has_own      # 제 차수의 앞 공고가 따로 있는 재공고는 그 공고를 대체한 것이다(다른 차수의 공고가 아니라)
            # 폐기물 처리의 '(1차)'·'(2차)' 는 차수분일 수 있다: 차수 표기가 서로 다르고 금액도 다르면(뒤 공고가 재공고로 표시돼 있지 않으면) 따로 난 용역으로 보아 둘 다 남긴다
            lots = (waste and not same_amt and not again and marks[a] != marks[b] and not _ORD_REBID.search(_txt(out.at[b, "공고명"])))
            if (gap <= REBID_NEAR_DAYS and not lots) or (gap <= REBID_MAX_DAYS and (same_amt or one_off or again)):
                replaced.add((out.at[a, "시설ID"], out.at[a, "공고번호"]))
    # 한 사업에 한 번 하는 용역(설계·건설사업관리/감리)이 같은 프로젝트에 같은 금액으로 가까운 날에 다시 났으면, 이름이 꽤 달라도 같은 용역이다 —
    # 평가 안내 공고와 본 공고('… 제출 안내 공고' ↔ '(PQ후 가격입찰)'), 머리말·어순을 바꿔 다시 낸 공고. 공종이나 단계·차수 표기가 서로 다르면 다른 용역으로 둔다
    # 더할 수 있는 줄끼리만 견준다(연간 단가·묶음 따위로 어차피 더하지 않는 줄이 멀쩡한 공고를 '재공고 이전' 으로 만들지 않게). 건설사업관리와 감리는 서로 다른 용역이다
    one = live[live["용역종류"].isin(("설계", "건설사업관리·감리")) & (live["프로젝트ID"].map(_txt) != "") & (live["합산제외_까닭"].map(_txt) == "")]
    is_cm = lambda i: _txt(out.at[i, "용역세부"]).startswith("건설사업관리")
    name_col = "_masked" if "_masked" in out.columns else "공고명"       # 공종·단계 낱말은 시설 이름을 가린 공고명에서 본다('가온소방서 신축공사 전기공사 감리' 의 '소방' 은 공종이 아니다)
    fset = out.groupby("공고번호")["시설ID"].agg(frozenset).to_dict()       # 이어진 시설이 서로 다른 두 공고('○○ 및 △△ …' ↔ '□□ 및 △△ …')는 같은 용역이 아니다
    for _, g in one.groupby(["시설ID", "프로젝트ID", "용역종류"], sort=False):
        if len(g) < 2:
            continue
        idx = list(g.sort_values(["_dt", "공고번호"], kind="stable").index)
        for j in range(1, len(idx)):
            for a in idx[:j]:
                b = idx[j]
                da, db = out.at[a, "_dt"], out.at[b, "_dt"]
                pa, pb = price.at[a], price.at[b]
                if pd.isna(da) or pd.isna(db) or not (pa > 0 and pb > 0) or (db - da).days > REBID_NEAR_DAYS or is_cm(a) != is_cm(b) \
                        or fset.get(out.at[a, "공고번호"]) != fset.get(out.at[b, "공고번호"]):
                    continue
                if abs(pb - pa) <= max(pa, pb) * REBID_SAME_AMT and not _other_service(_txt(out.at[a, name_col]), _txt(out.at[b, name_col])):
                    replaced.add((out.at[a, "시설ID"], out.at[a, "공고번호"]))
    mask =pd.Series([(f, n) in replaced for f, n in zip(out["시설ID"], out["공고번호"])], index=out.index)
    out.loc[mask & (out["상태"] == "대표"), "상태"] = "재공고 이전"
    return out.drop(columns=["_key"])


def _mark_sum(out: pd.DataFrame, projs: Dict[str, List[dict]]) -> pd.DataFrame:
    """합산여부(Y/N): 05 의 용역 참고 열에 더할 줄. Y = 세 종류 가운데 하나 + 확실한 연결 + 프로젝트가 정해짐 + 대표 공고 + 한 사업의 용역.
    한 용역 공고가 시설 둘 이상에 이어졌을 때:
      · 공고명에 서로 다른 시설의 이름이 '및'·'·'·',' 로 이어져 따로 들어 있으면(묶음 발주 '○○ 및 △△ 건립 …') 금액을 나눌 수 없어 모두 N
      · 어느 한 시설의 프로젝트에만 사업 이름이 통째로(더 길게) 맞으면 그쪽만 Y, 나머지는 N(이름이 비슷한 다른 시설 후보)
      · 같은 공사 공고를 함께 가진 시설끼리(같은 사업이 두 시설 후보로 잡힌 경우 — 05 '시설간_중복')는 한 사업으로 보아 양쪽 모두 Y
      · 근거가 같은 서로 다른 사업의 시설들은 모두 N
      · 사업 이름이 통째로 있지 않은데(시기·유사도로만 정함) 다른 시설에도 '확실' 로 이어진 공고는 N(프로젝트를 정하지 못한 쪽의 용역일 수 있다)"""
    out = out.copy()
    fac_keys = {f: (set().union(*[p["keys"] for p in pl]) if pl else set()) for f, pl in projs.items()}
    multi: Dict[Tuple[str, str], str] = {}
    live = out[out["상태"] == "대표"]                  # 프로젝트를 정하지 못한 줄도 '다른 시설에 이어졌다' 는 사실은 같다(근거의 세기 0 으로 견준다)
    for no, g in live.groupby("공고번호", sort=False):
        if g["시설ID"].nunique() < 2:
            continue
        # 묶음 발주: 한 공고명에 서로 다른 시설의 이름이 '및'·'·'·',' 로 이어져 따로 들어 있으면 금액을 나눌 수 없다 — 사업 이름이 통째로 맞는 쪽이 있어도 마찬가지다
        # ('○○ 체육관 및 △△ 수영장 건립 설계용역' 은 '△△수영장건립' 만 통째로 맞는다). 같은 공사 공고를 함께 가진 시설 후보끼리는 한 사업이므로 뺀다
        sp = dict(zip(g["시설ID"], g["_spans"]))
        fs = list(sp)
        apart = [x for i, a in enumerate(fs) for b in fs[i + 1:] if not (fac_keys.get(a, set()) & fac_keys.get(b, set()))
                 and _named_apart(g["공고명"].iloc[0], sp[a], sp[b]) for x in (a, b)]
        if apart:
            named = list(dict.fromkeys(apart))
            for f in fs:
                multi[(no, f)] = (f"한 공고명에 서로 다른 시설 {len(named)}개의 이름이 함께 들어 있음({', '.join(named[:4])}{' …' if len(named) > 4 else ''}) — "
                                  "묶음 발주로 보여 금액을 나눌 수 없음")
            continue
        best = float(g["_lvl"].max())
        strong = g[g["_lvl"] == best]
        for f in g.loc[g["_lvl"] < best, "시설ID"]:
            multi[(no, f)] = f"같은 공고가 다른 시설의 프로젝트({', '.join(x for x in dict.fromkeys(strong['프로젝트ID']) if x)})에 더 정확히 맞음"
        sf = list(dict.fromkeys(strong["시설ID"]))
        base = fac_keys.get(sf[0], set())
        if len(sf) >= 2:
            if not all(base & fac_keys.get(f, set()) for f in sf[1:]):
                for f in sf:
                    multi[(no, f)] = f"한 공고가 서로 다른 사업의 시설 {len(sf)}개에 이어짐({', '.join(sf[:4])}{' …' if len(sf) > 4 else ''}) — 금액을 나눌 수 없음"
        elif best < 3:
            # 공고명에 이 프로젝트의 사업 이름이 통째로 있지 않은데(시기·유사도로만 정함) 같은 공고가 다른 시설에도 '확실' 로 이어졌다 → 어느 시설의 용역인지 가릴 수 없다
            # (이름이 다른 시설 이름에 들어 있는 시설: '○○체육관' 과 '○○체육관 수영장' — 뒤의 시설에서 프로젝트를 정하지 못했을 뿐일 수 있다)
            others = [f for f in dict.fromkeys(g.loc[g["연결강도"] == "확실", "시설ID"]) if f != sf[0] and not (base & fac_keys.get(f, set()))]
            if others:
                multi[(no, sf[0])] = (f"같은 공고가 다른 시설({', '.join(others[:4])}{' …' if len(others) > 4 else ''})에도 이어지고 공고명에 이 프로젝트의 사업 이름이 통째로 있지 않음 "
                                      "— 어느 시설의 용역인지 가릴 수 없음")
    why = []
    for _, r in out.iterrows():
        w = _txt(r["합산제외_까닭"])
        if r["상태"] != "대표":
            w = "취소 공고" if r["상태"] == "취소" else "재공고로 대체됨(같은 용역이 뒤에 다시 공고됨)"
        if not w and r["용역종류"] not in KINDS:
            w = "세 종류(설계·건설사업관리/감리·폐기물처리)가 아님"
        if not w and r["연결강도"] != "확실" and not r["_named"]:
            # 어절 사이에 다른 말이 낀 이름·대행 기관의 공고는 '추정' 이지만, 공고명에 든 사업 이름이 이 현장을 가리키면(link 의 _named) 그 사업의 용역으로 본다
            w = "연결이 추정임(" + str(r["연결근거"]).split(" / ")[0] + ")"
        if not w and not _txt(r["프로젝트ID"]):
            w = "프로젝트를 정하지 못함(" + str(r["연결근거"]).split(" / ")[-1] + ")"
        if not w and (r["공고번호"], r["시설ID"]) in multi:
            w = multi[(r["공고번호"], r["시설ID"])]
        if not w and (r["추정가격"] is None or pd.isna(r["추정가격"]) or float(r["추정가격"]) <= 0):
            w = "추정가격 없음"
        why.append(w)
    out["합산제외_까닭"] = why
    out["합산여부"] = ["N" if w else "Y" for w in why]
    return out


# ── 6. 점검·요약 ─────────────────────────────────────────────────────────────────────────────────────────
RATIO_BAND = {"설계": (0.015, 0.12), "건설사업관리·감리": (0.005, 0.15), "폐기물처리": (0.0, 0.10)}     # 공사 추정가격 합 대비 용역 추정가격 합의 흔한 범위(벗어나면 확인)


def ratio_logs(linked: pd.DataFrame, latest: pd.DataFrame) -> List[Dict]:
    """프로젝트별 용역 금액(합산여부 Y)이 공사 추정가격 합에 견주어 지나치게 크면 06 검증로그에 '경고' 로 남긴다 —
    다른 사업의 용역이 섞였거나, 묶음 발주·나뉘어 다시 난 공고가 겹쳤을 수 있다. 작은 쪽은 알리지 않는다(용역이 공고 없이 집행되는 일이 흔하다)."""
    if linked.empty or latest.empty:
        return []
    amt = pd.to_numeric(latest["추정가격"], errors="coerce").fillna(0).groupby(latest["프로젝트ID"].astype(str)).sum()
    y = linked[linked["합산여부"] == "Y"]
    logs = []
    for (pid, kind), g in y.groupby(["프로젝트ID", "용역종류"], sort=True):
        base, tot = float(amt.get(pid, 0)), float(pd.to_numeric(g["추정가격"], errors="coerce").sum())
        hi = RATIO_BAND.get(kind, (0, 1))[1]
        if base > 0 and tot / base > hi:
            logs.append({"공고번호": pid, "항목": f"용역비 비율({kind})", "API값": base, "문서값": tot, "차이": None, "차이율": tot / base, "판정": "경고",
                         "비고": f"{kind} 용역 추정가격 합 {tot:,.0f}원이 이 프로젝트 공사 추정가격 합 {base:,.0f}원의 {tot / base:.1%}(흔한 범위는 {hi:.0%} 이하) — "
                                 f"다른 사업의 용역이 섞였는지, 묶음 발주·나뉘어 다시 난 공고가 겹쳤는지 08 시트에서 확인({len(g)}건: {', '.join(g['공고번호'].astype(str).head(5))})"})
    return logs


def profile(sv_raw: pd.DataFrame, cfg: dict, top: int = 12) -> str:
    """수집된 용역 공고의 분포(연결 규칙을 손볼 때 쓰는 진단). 공고명은 싣지 않는다 — 값의 종류와 건수만."""
    d = latest_orders(standardize(sv_raw, cfg))
    if d.empty:
        return "용역 공고 0건"
    lines = [f"[진단] 용역 공고 {len(sv_raw):,}줄 → 공고번호 기준 {len(d):,}건(최신 차수)"]
    vc = lambda s: " / ".join(f"{k or '(빈값)'} {v:,}" for k, v in s.value_counts(dropna=False).head(top).items())
    lines.append("  용역구분: " + vc(d["용역구분"]))
    lines.append("  공고종류(판정): " + vc(d["공고종류"]) + "  | 원문: " + vc(d["공고종류_원문"]))
    lines.append("  조달분류(많은 순): " + vc(d["조달분류"]))
    lines.append(f"  공고기관이 조달청인 공고: {int(d['공고기관'].str.contains('조달청', na=False).sum()):,}건 / 수요기관 빈값: {int((d['수요기관'] == '').sum()):,}건 / "
                 f"수요기관코드 있음: {int((d['수요기관코드'] != '').sum()):,}건 / 이전공고번호 있음: {int((d['이전공고번호'] != '').sum()):,}건")
    p, b = pd.to_numeric(d["추정가격"], errors="coerce"), pd.to_numeric(d["배정예산"], errors="coerce")
    ok = (p > 0) & (b > 0)
    r = (b[ok] / p[ok])
    lines.append(f"  추정가격 없음(0 포함): {int((~(p > 0)).sum()):,}건 / 배정예산÷추정가격: 1.09~1.11 {int(((r > 1.09) & (r < 1.11)).sum()):,}건, 1 미만 {int((r < 1).sum()):,}건, "
                 f"1.11 초과 {int((r >= 1.11).sum()):,}건(비교 가능한 {int(ok.sum()):,}건 중)")
    kinds = pd.Series([classify_kind(t, c)[0:2] for t, c in zip(d["공고명"], d["조달분류"])])
    lines.append("  공고명으로 본 종류: " + vc(kinds.map(lambda k: k[0] + (f"·{k[1]}" if k[1] else ""))))
    lines.append(f"  묶음 표기('외 N건'·'등 N교' 등)가 있는 공고: {int(d['공고명'].str.contains(BUNDLE, na=False).sum()):,}건")
    return "\n".join(lines)


def _signed(v) -> str:
    n = int(round(float(v)))
    return f"{n:+d}" if n else "0"


def summary(linked: pd.DataFrame, n_raw: int, n_latest: int, latest: pd.DataFrame, fac: Optional[pd.DataFrame] = None) -> str:
    """link-servc 화면 요약(글)."""
    lines = [f"용역 공고 {n_raw:,}줄(공고번호 기준 {n_latest:,}건) 가운데 시설에 이어진 것: {linked['공고번호'].nunique() if len(linked) else 0:,}건 "
             f"(연결 줄 {len(linked):,}, 시설 {linked['시설ID'].nunique() if len(linked) else 0:,}개)"]
    if linked.empty:
        return "\n".join(lines + ["이어진 공고가 없습니다. 검수 파일의 시설 이름·수요기관과 용역 공고의 수요기관이 맞는지 확인하세요."])
    y = linked[linked["합산여부"] == "Y"]
    n_proj = latest["프로젝트ID"].nunique() if not latest.empty else 0
    lines.append(f"05 참고 열에 더하는 줄(합산여부 Y): {len(y):,}줄, 프로젝트 {y['프로젝트ID'].nunique():,}개 / 전체 {n_proj:,}개")
    for k in KINDS:
        g = y[y["용역종류"] == k]
        gap = pd.to_numeric(g["공사대비_개월"], errors="coerce").dropna()
        when = f", 공사 첫 공고 대비 시점 중앙값 {_signed(gap.median())}개월(10%~90%: {_signed(gap.quantile(0.1))}~{_signed(gap.quantile(0.9))})" if len(gap) else ""
        u = g.drop_duplicates("공고번호")              # 같은 공고가 시설 후보 둘에 함께 실린 줄(05 '시설간_중복')은 건수·금액에서 한 번만 센다
        lines.append(f"  · {k}: {len(u):,}건({len(g):,}줄), 프로젝트 {g['프로젝트ID'].nunique():,}개, 추정가격 합 {pd.to_numeric(u['추정가격'], errors='coerce').sum() / 1e8:,.1f}억"
                     f"(공고마다 한 번만 셈){when}")
    n = linked[linked["합산여부"] == "N"]
    if len(n):
        top = n["합산제외_까닭"].map(lambda s: re.split(r"[\(（—]", s)[0].strip()[:34]).value_counts().head(10)
        lines.append(f"더하지 않는 줄(합산여부 N) {len(n):,}줄. 까닭별: " + " / ".join(f"{k} {v:,}" for k, v in top.items()))
    big = linked.groupby(["시설ID", "시설명"])["공고번호"].nunique().sort_values(ascending=False).head(8)
    lines.append("연결이 많은 시설(이름이 일반적이면 다른 사업의 용역이 섞였을 수 있음. 08 시트에서 확인): "
                 + ", ".join(f"{f} {str(nm)[:16]} {c}건" for (f, nm), c in big.items()))
    if fac is not None and not fac.empty and "연면적_m2" in fac.columns:
        blank = set(fac.loc[fac["연면적_m2"].isna(), "프로젝트ID"].astype(str))
        # 과업지시서에 연면적이 적히는 용역만: 더하는 줄(Y) 가운데 설계와 건설사업관리(전기·소방·통신 감리, 추정 연결, 묶음 공고는 빼고)
        has = set(y.loc[(y["용역종류"] == "설계") | y["용역세부"].astype(str).str.startswith("건설사업관리"), "프로젝트ID"])
        lines.append(f"연면적이 빈 프로젝트 {len(blank):,}개 가운데 설계·건설사업관리 용역(합산여부 Y)이 이어진 프로젝트: {len(blank & has):,}개"
                     "(용역 과업지시서로 연면적을 보완할 수 있는 후보)")
    return "\n".join(lines)
