"""세부사업 분리: 한 시설 후보(시설ID×사업유형)에 묶인 공고를 '세부사업'(한 건물·한 단지·한 현장을 한 번에 짓는 공사)으로 나누고,
세부사업 안에서 재공고·재발주로 대체된 공고를 가린다.

배경: 시설 후보가 여러 사업을 묶은 경우(예: '서울주택도시개발공사 공공주택' = 고덕강일 3단지·마곡 16단지·위례 A1-14BL …)
'같은 공종의 최신 공고 1건만 대표' 규칙이 서로 다른 사업을 대체시켜 금액이 빠지고 한 행에 여러 현장의 금액이 섞였다.

 세부사업 = 같은 현장·같은 건물의 공사를 공종별(건축·전기·통신·소방·기계·토목·조경)·분야별(기계소방/전기소방)·차수별(1차분/2차분)·
            부대공사(무대기계·승강기·전시·인테리어 등)로 나눠 발주한 공고의 묶음 → 모두 합산
 다른 세부사업 = 다른 단지·지구·블록·동·권역·현장, 다른 연도의 연례 사업, 같은 시설 안이라도 공간·목적과 시기가 다른 별개 공사
 대체(superseded) = 같은 세부사업·같은 공종에서 같은 공고명으로 며칠 뒤 다시 난 공고(재공고·유찰 후 재발주)가 있는 앞 공고, 또는 API 이전공고번호로 연결된 앞 공고.
            공고명에 차수·단계 표기가 다르면('조성사업' / '조성사업(2차분)', '(1차)' / '(2차)') 재공고가 아니라 차수별 발주 → 둘 다 합산
 제안 안내 공고 = '공법·자재 선정 기술제안서 제출 안내' 처럼 공사 발주가 아닌 공고(is_proposal). 그 금액은 본 공사 공고에 들어 있으므로 합산하지 않는다(dedup 에서 제외)

방법: 공고명에서 머리말·공종어·일반어·시설명을 지우고 남는 '구분어'(현장·건물·구간 이름)를 비교한다.
 ① 구분어가 같으면(둘 다 비었거나 거의 같음) 같은 세부사업  ② 한쪽 구분어가 다른 쪽에 들어 있으면(공종 공고 ⊂ 본공사) 같은 세부사업
 ③ 구분어가 빈 본공사와 구분어가 있는 공고는, 뒤쪽이 자기 목적어를 가진 별개 사업('○○ 신축', '△△ 조성')이 아닐 때만 같은 세부사업
 어느 경우든 두 공고의 간격이 공사 규모에 맞는 기간 안이어야 한다(작은 공사는 짧게, 큰 공사는 길게).
"""
from __future__ import annotations

import re
from difflib import SequenceMatcher
from typing import Dict, List

import pandas as pd

# 판정 기준(실제 공고 이력의 표본 프로젝트를 사람이 나눈 결과와 대조해 정한 값)
PARAMS = dict(
    disc_same=0.75,                   # 구분어가 오타 수준으로 같은지 따져 볼 최소 유사도(그 뒤 글자 단위 조건을 다시 본다)
    windows=((3e8, 270), (3e9, 600), (1e10, 800), (float("inf"), 1200)),   # (추정가격 상한, 같은 세부사업으로 보는 최대 간격[일])
    near_days=120,                    # 본공사와 이 기간 안에 난 공고는 이름에 자기 사업어('○○ 리모델링')가 있어도 같은 세부사업의 분할 발주로 본다
    year_slack=(60, 120),             # 연도 표기('2024년')가 한쪽에만 있을 때: 다른 쪽 공고일이 그 해의 60일 전 ~ 120일 후 안이면 같은 해 사업
    rebid_min_days=3, rebid_max_days=300, rebid_lo=0.6, rebid_hi=1.67,      # 같은 공고명으로 다시 난 공고를 재공고로 보는 간격[일]·금액비
                                                                            # (300일을 넘겨 같은 이름·같은 공종으로 다시 나면 다음 해의 연례 사업 → 다른 세부사업)
                                                                            # 공고종류·공고명에 '재공고' 가 있으면 같은 날·다음 날 다시 난 것도 재공고
    oneoff_max_days=1200, oneoff_lo=0.5, oneoff_hi=2.0,                     # 건립·신축·증축·조성처럼 한 번 하는 공사가 같은 공고명·같은 공종으로 다시 나면
                                                                            # 300일이 넘어도(사업 지연 뒤 재발주) 같은 사업의 재공고로 본다
    recur_max_days=180,               # 유지관리·연간단가처럼 해마다(반기마다) 다시 발주하는 공사는 같은 이름이 180일 넘게 떨어지면 다음 회차(재공고 표기가 있어도 재공고 아님)
    rebid_fuzzy=0.92, rebid_fuzzy_lo=0.85, rebid_fuzzy_hi=1.18,             # 공고명이 조금 바뀐 재공고: 이름 유사도·금액비(더 좁게)
    repeat_sim=0.9,                   # 같은 공종에서 300일 넘게 떨어진 두 공고의 이름이 이만큼 비슷하면 연례 사업의 다른 해 분(다른 세부사업)
    anc_max_ratio=5.0,                # 부대공사로 보고 본공사에 붙이는 공고의 금액은 그 본공사 공고의 5배 이하(큰 공사가 작은 공사의 부대공사일 수는 없다)
    anc_max_len=8,                    # 부대공사 이름으로 보는 구분어의 최대 길이('연결통로', '공중화장실'). 더 길면 자기 이름을 가진 별개 공사로 본다
)

_BR_OPEN, _BR_CLOSE = r"\[\(（【<〈", r"\]\)）】>〉"
# 행정 문구: 괄호 안이 이런 말(과 숫자·'차'·'제' 등)로만 이루어진 경우에만 꼬리표로 보고 지운다.
# '(수정동 공영주차장)'·'(안내동)' 처럼 지명에 같은 글자가 들어간 괄호는 현장 이름이므로 남긴다
_ADMIN = (r"긴급|재공고|재입찰|입찰\s*공고|입찰\s*대행|수의\s*계약|수의\s*견적|소액\s*수의|수의|견적\s*제출|견적|안내\s*공고|안내|공고|제출|장기\s*계속\s*계약|장기계속|"
          r"별도\s*재무관|전자\s*입찰|지역\s*제한|총액|공사\s*계약|계약|변경|정정|수정|취소|소액")
_BRACKET = re.compile(r"[%s]([^%s%s]{0,40})[%s]" % (_BR_OPEN, _BR_OPEN, _BR_CLOSE, _BR_CLOSE))
_ADMIN_ONLY = re.compile(r"(?:%s|\d|차|제|회|후|건|분|및|[\s,·ㆍ/\-_.:])+" % _ADMIN)
_ADMIN_RE = re.compile(_ADMIN)
_LOT_IN = re.compile(r"(?:제\s*)?\d+(?:\s*-\s*\d+)*\s*차\s*분")                        # 행정 괄호 안에 함께 적힌 차수 표기: '(2차분 재공고)' 의 '2차분'
_LOT1 = r"(?:제\s*)?\d+(?:\s*-\s*\d+)*\s*(?:차(?:\s*분)?)?"
_LOT_TAG = re.compile(r"\s*%s(?:\s*(?:[,·ㆍ/]|및)\s*%s)*\s*" % (_LOT1, _LOT1))     # 괄호 전체가 차수 표기: '(2차)', '(2차분)', '(3-2)', '(2)'
_LOT_BARE = re.compile(r"(?:제\s*)?\d+(?:\s*-\s*\d+)+|(?:제\s*)?\d+\s*차(?!\s*분)")   # '(2차, 긴급)' 의 '2차', '(긴급, 3-2)' 의 '3-2'
_ORDINAL_ADMIN = re.compile(r"재공고|재입찰|공고|변경|정정|수정|취소")               # 숫자·차수가 공고 회차를 뜻하게 만드는 행정 문구('(2차 재공고)', '(2차 변경)')
_REBID_WORD = re.compile(r"재공고|재입찰|새로운\s*입찰")
_NOISE = re.compile(r"(?:전체|총\s*차|총괄)\s*분(?![가-힣])|(?:제\s*)?\d+\s*차\s*(?:재공고|재입찰)|(?:재공고|재입찰)\s*[\(\[（]?\s*(?:제\s*)?\d+\s*차\s*[\)\]）]?|새로운\s*입찰|"
                    r"긴급|재공고|재입찰|전자\s*입찰\s*공고|전자\s*입찰|입찰\s*공고|입찰\s*대행|"
                    r"(?:소액\s*수의|수의\s*(?:계약|견적))(?:\s*견적)?(?:\s*제출)?(?:\s*안내)?(?:\s*공고)?|견적\s*(?:제출\s*)?(?:안내\s*)?(?:공고)?|안내\s*공고|"
                    r"(?:여성|장애인|사회적)[^)\]）]{0,24}대상|공법\s*제안\s*공고|공법\s*제안|공사\s*계약|장기\s*계속\s*계약|계속비|별도\s*재무관|총괄|가칭")


_BARE_ORD = re.compile(r"\s*(?:제\s*)?\d+\s*차\s*")


def _strip_admin(name, drop_ord: bool = False) -> str:
    """공고명에서 행정 꼬리표 괄호('[긴급]', '(재공고)', '(2차 변경)')와 행정 문구('입찰공고', '수의견적 제출 안내공고')를 지운다.
    괄호 전체가 차수·단계 표기인 '(2차)'·'(2차분)'·'(3-2)' 는 다른 계약을 가리키므로 남긴다. '(2차분 재공고)' 는 '2차분', '(2차, 긴급)' 은 '2차' 만 남기고,
    '(2차 재공고)'·'(2차 변경)'·'(2회)'·'(제2회)' 는 공고 회차이므로 지운다."""
    def tag(m):
        x = m.group(1)
        if not (x.strip() and _ADMIN_ONLY.fullmatch(x)):
            return m.group(0)
        if not _ADMIN_RE.search(x):
            if not _LOT_TAG.fullmatch(x):
                return " "                                             # '(2회)'·'(제2회)'·'(1건)' 은 공고 회차 → 예전처럼 지운다
            return " " if (drop_ord and _BARE_ORD.fullmatch(x)) else m.group(0)
        lots = _LOT_IN.findall(x)
        if not lots and not _ORDINAL_ADMIN.search(x):
            lots = _LOT_BARE.findall(x)                                # '(2차, 긴급)' → '2차'
            if drop_ord:
                lots = [t for t in lots if not _BARE_ORD.fullmatch(t)]
        return (" " + " ".join(lots) + " ") if lots else " "
    return _NOISE.sub(" ", _BRACKET.sub(tag, str(name or "")))


# 제안 안내 공고: 특허공법·자재 선정을 위한 기술제안서 제출 안내·공모. 공사 발주 공고가 아니며, 금액이 적혀 있어도 본 공사 공고의 일부다.
# '실시설계 기술제안 입찰공고'(기술제안입찰 = 실제 공사 발주)와 '…건립공사(특허공법)'(특허공법을 포함한 공사)은 해당하지 않는다
# 무엇을 제안받는지(공법·자재·신기술·특허)가 이름에 있어야 한다 — '전시물 제작·설치 제안서 제출 안내'(협상에 의한 계약 = 실제 발주)는 해당하지 않는다
_PROPOSAL = re.compile(r"제안서\s*(?:제출|접수)\s*(?:안내|공고)|(?:공법|자재|기술)\s*제안\s*(?:공모|공고|안내)"
                       r"|(?:공법|자재)\s*선정[^,)\]）]{0,12}(?:공모|안내\s*공고)|(?:공법|자재)\s*(?:선정\s*)?(?:공모|공고)(?![가-힣])")
_PROPOSAL_CTX = re.compile(r"공법|(?<!기)자재|신기술(?!\s*(?:창업|사업화|센터|단지))|특허")          # '기자재'·'신기술창업센터' 는 공법·자재 선정이 아니다
# 제안서를 받더라도 공사·제작 발주인 공고: 실시설계 기술제안 입찰, 협상계약 방식의 제작·설치(전시물, 무대 기자재)
_PROPOSAL_NOT = re.compile(r"(?:실시|기본)\s*설계\s*기술\s*제안|기술\s*제안\s*입찰|입찰\s*공고(?!\s*(?:가\s*)?아님)|(?:제작|구매)\s*[·ㆍ,및]?\s*설치|제작\s*[·ㆍ,및]?\s*구매")
_ORD_REBID_BR = re.compile(r"[%s]?\s*((?:제\s*)?\d+\s*차)\s*(?:재공고|재입찰)\s*[%s]?|(?:재공고|재입찰)\s*[%s]?\s*((?:제\s*)?\d+\s*차)\s*[%s]?"
                           % (_BR_OPEN, _BR_CLOSE, _BR_OPEN, _BR_CLOSE))   # '(2차 재공고)', '2차 재공고', '재공고(2차)'


def is_proposal(name) -> bool:
    """공법·자재 선정용 제안서 제출 안내 공고(공사 발주가 아닌 공고)인지."""
    s_ = str(name or "")
    return bool(_PROPOSAL.search(s_)) and bool(_PROPOSAL_CTX.search(s_)) and not _PROPOSAL_NOT.search(s_)


_TRADE = "정보통신|통신설비|통신|전기설비|전기|소방시설|소방|기계설비|기계|설비|토목|조경|건축|골조|마감"
_TRADE_RE = re.compile(_TRADE)
_YEAR = re.compile(r"(20[12]\d)\s*년(?:도)?")
# 차수·단계 표기: '3단계', '1-1공구', '(3-2단계)', '(3-2)', '2차분', 낱말로 떨어져 있는 '2차'·'제2차'.
# 아파트·단지 이름에 붙은 'N차'('이지더원2차', '푸르지오3차')는 다른 단지를 가리키는 이름이므로 지우지 않는다
_PHASE = re.compile(r"(?:제\s*)?(\d+(?:\s*-\s*\d+)*)\s*(단계|공구|차분|차년도)"
                    r"|[\(\[（]\s*(\d+(?:\s*-\s*\d+)+)\s*()[\)\]）]"
                    r"|(?<![0-9A-Za-z가-힣])(?:제\s*)?(\d+)\s*(차)(?![0-9A-Za-z가-힣])")
_LOT_LIST = re.compile(r"[%s]\s*(%s(?:\s*(?:[,·ㆍ/]|및)\s*%s)+)\s*[%s]" % (_BR_OPEN, _LOT1, _LOT1, _BR_CLOSE))
_ONEOFF = re.compile(r"건립|신축|증축|개축|재건축|건설|조성|신설")
_RECUR = re.compile(r"유지|관리|연간|단가|정비|보수|위탁|운영")
# 해마다 다시 발주하는 것이 분명한 공사. '관리동 증축'·'개보수'·'국유지', 건물 이름 '유지관리동·유지관리센터', 낱말 속 글자 '산단가로등' 에는 걸리지 않게
_ANNUAL = re.compile(r"유지\s*(?:[·ㆍ/]|및)?\s*(?:관리|보수)(?!\s*(?:동|센터|사무소|청사|실(?![가-힣])))|(?<![가-힣])연간|연간\s*(?:단가|유지|관리|보수|정비)|(?<![0-9가-힣])단가|단가\s*(?:계약|공사|입찰)")
# 연례 공사 말 뒤에 오는 한 번짜리 공사 말: '유지관리 사무실 신축공사'. '및' 으로 나란히 적은 말('유지보수 및 신설공사')은 연례 공사의 일부
_ANNUAL_NOT = re.compile(r"(?<!및)(?<!및\s)(?:건립|신축|증축|개축|재건축|신설)(?=$|[^가-힣]|공사|사업)")
_ANNUAL_STRONG = re.compile(r"연간\s*단가|단가\s*(?:계약|공사|입찰)")   # 이름에 있으면 뒤에 무슨 말이 오든 연례 공사: '시설물 유지보수공사(연간단가, 신설 포함)'


def is_annual(name) -> bool:
    """해마다 다시 발주하는 공사 이름인지. '유지관리 사무실 신축공사'·'유지관리용 창고 신축공사' 처럼 연례 공사 말 뒤에 건물 공사 말이 오면 한 번짜리 공사.
    다만 '연간단가'·'단가계약' 이 있거나 건물 공사 말이 '및' 으로 이어져 있으면('유지보수 및 신설공사') 연례 공사."""
    s_ = str(name or "")
    if _ANNUAL_STRONG.search(s_):
        return True
    m = _ANNUAL.search(s_)
    return bool(m) and not _ANNUAL_NOT.search(s_, m.end())
# '부지'·'이전' 은 낱말로 쓰였을 때만 일반어('부지조성'·'주차장부지 조성', '이전 조성'·'이전에 따른'). 낱말 속 글자('서부지구', '본부지하주차장', '어린이전용', '기업이전단지')와
# 장소 이름 '고수부지' 는 지우지 않는다
_SITE = (r"(?<![동서남북중])(?<!고수)부지(?=$|[^가-힣]|조성|정비|정리|매입|성토)"
         r"|(?<!어린)이전(?=$|[^가-힣]|공사|사업|설치|신축|건립|조성|에\s*따른)")
_SITE_RE = re.compile(_SITE)
# 일반어: 사업 종류·공사 종류·부대공사 이름처럼 현장을 가리키지 않는 말(지워도 어느 현장인지는 달라지지 않는다)
_GENERIC = re.compile(
    r"건립|신축|증축|개축|재건축|건설|조성|리모델링|개보수|보수|정비|공사|사업(?![화단소장부비])|시설물|일원|일대|"
    + _SITE + r"|"
    r"인테리어|실내|내부|외부|부대|보강|개설|설치|철거|해체|석면|지장물|지정폐기물|폐기물|포장|가설|울타리|정화조|승강기|엘리베이터|"
    r"무대\s*(?:기계|장치|조명|음향)?|전시물?|집기|가구|사인물|간판|제작|구매|관급\s*자재|자재|기타|일괄|에\s*따른|관련|시설|구조|수선|분야|부문|시행|및")
_SITE_WORD = re.compile(r"(?<![동서남북중])(?<!고수)부지(?=$|[^가-힣])|(?<!어린)이전(?=$|[^가-힣])")
_BUILD_RE = re.compile(r"(?:건립|신축|증축|개축|재건축|건설|신설)(?!\s*(?:예정|사업|공사)?\s*부지)")   # 부지 정리가 아니라 건물 공사를 함께 발주한 공고
_VERB = re.compile(r"건립|신축|증축|개축|재건축|건설|조성|리모델링|개보수|신설|확장|이전")
_REGION_TAIL = re.compile(r"(특별자치도|특별자치시|특별시|광역시|자치도|도|시|군|구)$")
_PUNCT = re.compile(r"[\s\-_·ㆍ・,.\[\]\(\)（）【】<>〈〉/~:;'\"]")


def clean_title(name) -> str:
    """머리말·꼬리말('[긴급]', '(재공고)', '입찰공고')을 뺀 공고명(표시용)."""
    t = re.sub(r"\s+", " ", _strip_admin(name))
    t = re.sub(r"\s*,\s*(?=[\]\)）])", "", re.sub(r"(?<=[\[\(（])\s*,\s*", "", t))      # 문구를 지우고 남은 '(전기, )' → '(전기)'
    t = re.sub(r"(?<=[\[\(（/])\s*및\s*(?=(?:제\s*)?\d)", "", re.sub(r"(?<=[\[\(（])\s*및\s+", "", t))   # '(전체분 및 1차분)'·'(총괄및1차분)' → '(1차분)'
    t = re.sub(r"(?<=,)\s*및\s*(?=(?:제\s*)?\d)", " ", t)                           # '(장기계속, 전체분 및 1차분)' → '(장기계속, 1차분)'
    t = re.sub(r"[\[\(（]\s*[\]\)）]", "", t)                                               # 빈 괄호
    return re.sub(r"\s+", " ", t).strip(" -")


def title_key(name) -> str:
    """재공고 판별용 공고명 비교 키: 머리말·꼬리말과 공백·괄호·기호를 지운다."""
    return _PUNCT.sub("", _strip_admin(name))


def similarity(a: str, b: str) -> float:
    """두 문자열의 유사도(0~1)."""
    if not a or not b:
        return 0.0
    return SequenceMatcher(None, a, b, autojunk=False).ratio()


def _facility_parts(facility) -> List[str]:
    """시설명에서 공고명에서 지울 조각: 어절과, 행정구역 꼬리(시·군·구·도)를 뗀 어절. 긴 것부터."""
    parts = set()
    for t in re.split(r"\s+", _PUNCT.sub(" ", str(facility or "")).strip()):
        if len(t) >= 2:
            parts.add(t)
            stem = _REGION_TAIL.sub("", t)
            if len(stem) >= 2:
                parts.add(stem)
            bare = _SITE_WORD.sub("", t)                # 시설명 어절에 붙은 '부지'·'이전'('환승주차장부지')은 공고명에서 이미 지워졌으므로 뗀 꼴도 지운다
            if len(bare) >= 2:
                parts.add(bare)
    whole = _PUNCT.sub("", str(facility or ""))
    if len(whole) >= 2:
        parts.add(whole)
    return sorted(parts, key=lambda x: (-len(x), x))


_LIST_BR = re.compile(r"[\[\(（]([^\]\)）]*,[^\]\)）]*)[\]\)）]")


def features(name, facility="") -> Dict:
    """공고명 → 구분어(disc)·연도·식별 숫자·자기 사업어 유무·본공사 접두 비교용 키(base)·묶음 괄호의 낱개 구분어(alts)."""
    s = _strip_admin(name)
    years = frozenset(_YEAR.findall(s))
    s = _YEAR.sub(" ", s)
    lots = [_PUNCT.sub("", m.group(1)) for m in _LOT_LIST.finditer(s)]      # '(2,3차분)', '(2차 및 3차)' — 여러 차수를 묶은 괄호
    s = _LOT_LIST.sub(" ", s)
    phases = frozenset("".join(x for x in m.groups() if x).replace(" ", "") for m in _PHASE.finditer(s)) | frozenset(lots)
    s = _PHASE.sub(" ", s)
    verb = bool(_VERB.search(s))
    s = _TRADE_RE.sub(" ", s)
    base = _PUNCT.sub("", re.sub(r"공사|사업(?![화단소장부비])", " ", s))      # 공종어와 '공사·사업'만 뺀 이름: '무등야구장리모델링' 이 '무등야구장리모델링야구장클라이밍장조성' 의 앞부분인지 볼 때 쓴다
    site = frozenset(m.group(0) for m in _SITE_RE.finditer(s))      # 일반어로 지워지는 '부지'·'이전' (같은 이름의 '부지조성공사' 와 '조성공사' 는 다른 계약)
    s = _GENERIC.sub(" ", s)
    parts = _facility_parts(facility)

    def squash(t: str) -> str:
        # 시설명 조각을 지운다. 띄어쓰기가 달라도 지우되('걷고 싶은 거리' ↔ '걷고싶은거리'), 낱말 가운데에서 시작해 다음 낱말로 넘어가는 우연한 일치는
        # 지우지 않는다('둔덕정수장 고도정수처리' 안의 '수장고' — 시설명 '수장고' 와 무관)
        t = _PUNCT.sub("\x00", t)
        t = re.sub("\x00+", "\x00", t).strip("\x00")
        for part in parts:
            pat = re.compile("\x00?".join(re.escape(ch) for ch in part))
            out, pos = [], 0
            for mm in pat.finditer(t):
                crosses = "\x00" in mm.group(0)
                at_start = mm.start() == 0 or t[mm.start() - 1] == "\x00"
                if crosses and not at_start:
                    continue
                out.append(t[pos:mm.start()])
                pos = mm.end()
            t = "".join(out) + t[pos:]
        t = t.replace("\x00", "")
        return re.sub(r"^(?:등|외|내|의)+|(?:등|외|내|의)+$", "", t)

    disc = squash(s)
    # '(지산, 범물)' 처럼 여러 현장을 한 괄호에 묶은 공고(공통 전기 공고 등): 낱개 현장만 남긴 구분어도 만들어 현장별 공고와 이어지게 한다
    alts = set()
    m = None
    for m in _LIST_BR.finditer(s):
        pass
    if m:
        for item in (x.strip() for x in m.group(1).split(",")):
            if _PUNCT.sub("", item):
                alt = squash(s[:m.start()] + " " + item + " " + s[m.end():])
                if alt and alt != disc:
                    alts.add(alt)
    raw = str(name or "")
    raw_nf = raw
    for part in parts:
        raw_nf = raw_nf.replace(part, " ")
    return {"disc": disc, "years": years, "phases": phases, "digits": frozenset(re.findall(r"\d+", disc)), "verb": verb, "base": base,
            "alts": frozenset(alts), "rebid": bool(_REBID_WORD.search(raw)),
            "oneoff": bool(_ONEOFF.search(raw)) and not _RECUR.search(raw) and not years, "site": site,
            "site_prep": "부지" in site and not _SITE_RE.search(str(facility or "")) and not _BUILD_RE.search(raw_nf)}


def _window(price) -> int:
    p = float(price) if (price is not None and pd.notna(price)) else 0.0
    for limit, days in PARAMS["windows"]:
        if p < limit:
            return days
    return PARAMS["windows"][-1][1]


def _contained(short: str, long: str) -> bool:
    """짧은 구분어가 긴 구분어에 그대로 들어 있는지(2자 이상)."""
    return bool(short) and 2 <= len(short) < len(long) and short in long


_AREA_SUFFIX = ("지구", "단지", "마을", "지역", "권역", "구역")
_OPPOSITES = ("동서", "남북", "본별", "상하", "좌우", "전후", "신구", "내외", "대소", "AB", "가나")      # 한 글자만 달라도 서로 다른 곳을 가리키는 짝


def _typo_same(a: str, b: str, rare: bool = True) -> bool:
    """오타·표기 차이 수준으로 같은 이름인지('교육관동…'/'교육관종…', '종합상황실'/'종합감시실', '위례'/'위례지구').
    다른 부분이 1~2자 바뀐 것(숫자·방위 짝 제외)이거나 '지구·단지·권역' 같은 꼬리말이 한쪽에만 붙은 경우이고, 그 앞뒤로 같은 글자가 2자 이상 이어질 때만 같다고 본다.
    글자가 바뀐 경우는 두 이름 중 하나가 그 시설에서 한 번만 나올 때(rare)만 오타로 본다 — '북수지구'·'공수지구', '신사'·'신길' 처럼 양쪽 다 여러 공고에 쓰인 이름은 서로 다른 현장.
    이름의 앞머리가 다른 '상반기…'/'하반기…', '눈소공원'/'마동공원', 가운데에 말이 더 들어간 '논현1권역'/'논현고잔1권역' 도 다른 현장.
    (호출부는 공종이 다른 두 공고 사이에서만 이 판정을 쓴다 — 같은 공종으로 나란히 난 '(축구장)'·'(야구장)' 은 다른 현장)"""
    if a == b:
        return True
    if min(len(a), len(b)) < 4 or similarity(a, b) < PARAMS["disc_same"]:
        return False
    ops = SequenceMatcher(None, a, b, autojunk=False).get_opcodes()
    for n, (tag, i1, i2, j1, j2) in enumerate(ops):
        if tag == "equal":
            continue
        da, db = a[i1:i2], b[j1:j2]
        if re.search(r"\d", da + db):
            return False
        swap = rare and len(da) == len(db) and 1 <= len(da) <= 2 and not any(da[k] != db[k] and (da[k] + db[k] in o or db[k] + da[k] in o)
                                                                                  for k in range(len(da)) for o in _OPPOSITES)
        one = rare and max(len(da), len(db)) == 1 and min(len(da), len(db)) == 0        # 한 글자 빠짐·더함
        suffix = (not da and db in _AREA_SUFFIX) or (not db and da in _AREA_SUFFIX)
        if not (swap or one or suffix):
            return False
        left = ops[n - 1] if n > 0 else None
        right = ops[n + 1] if n + 1 < len(ops) else None
        if suffix and left and left[0] == "equal" and left[2] - left[1] >= 2:
            continue                                   # 꼬리말은 이름 끝에 붙어도 된다('위례' / '위례지구')
        if swap and len(da) == 1 and right is None and left and left[0] == "equal" and left[2] - left[1] >= 4:
            continue                                   # 이름 끝 한 글자가 바뀐 오타('제4봉안동' / '제4봉안당') — 앞 4자 이상이 같을 때
        if not (left and left[0] == "equal" and left[2] - left[1] >= 2 and right and right[0] == "equal" and right[2] - right[1] >= 2):
            return False
    return True


def _year_ok(fa: Dict, fb: Dict, dta, dtb) -> bool:
    """연례 사업의 연도 구분. 둘 다 연도 표기가 있으면 같아야 하고, 한쪽에만 있으면 다른 쪽 공고일이 그 해(앞 60일~뒤 120일 여유) 안이어야 한다."""
    ya, yb = fa["years"], fb["years"]
    if ya and yb:
        return ya == yb
    if not ya and not yb:
        return True
    ys, dt = (ya, dtb) if ya else (yb, dta)
    before, after = PARAMS["year_slack"]
    return any(pd.Timestamp(f"{y}-01-01") - pd.Timedelta(days=before) <= dt <= pd.Timestamp(f"{y}-12-31") + pd.Timedelta(days=after) for y in ys)


def _link(fa: Dict, fb: Dict, gap: int, window: int, ta: str, tb: str, new_build: bool, pa: float = 0.0, pb: float = 0.0) -> str:
    """두 공고가 같은 세부사업일 근거: 'same' / 'contain' / 'main' / ''(없음). 연도·기간 조건은 호출부에서 본다.
    gap = 두 공고 간격(일), window = 규모에 따른 허용 간격, ta·tb = 두 공고의 공종, new_build = 신축 프로젝트인지, pa·pb = 추정가격."""
    same_trade = ta == tb
    da, db = fa["disc"], fb["disc"]
    if not da and not db:
        return "same"
    if da and db:
        if fa["digits"] != fb["digits"] and (fa["digits"] or fb["digits"]) and not (fa["digits"] < fb["digits"] or fb["digits"] < fa["digits"]):
            return ""
        # 오타 수준 차이는 서로 빠진 공종을 채워 주는 공고 사이에서만 같은 이름으로 본다: 건축 공고만 '제4봉안동', 전기·소방·통신 공고는 '제4봉안당'.
        # 같은 공종으로 나란히 난 한두 글자 다른 이름('(축구장)'·'(야구장)', '고잔역'·'중앙역')이나, 상대 이름 묶음에 이미 같은 공종 공고가 있는 경우
        # ('(신사)' 건축·전기가 있는데 '(신길)' 건축) 는 서로 다른 현장
        if da == db:
            return "same"
        if (fa["digits"] == fb["digits"] and not same_trade and ta not in fb.get("trs", ()) and tb not in fa.get("trs", ())
                and _typo_same(da, db, rare=min(fa.get("dn", 1), fb.get("dn", 1)) <= 1)):
            return "same"
        # 포함 관계는 공종이 다른 공고 사이에서만 본다(여러 현장에 공통인 전기 공고 ↔ 현장별 건축 공고, 본공사 ↔ 그 건물의 공종 공고).
        # 같은 공종이면 이름이 짧은 쪽도 독립된 계약이므로('○○ 조성사업' 과 같은 날의 '○○ 조성사업(영월군청)') 서로 다른 사업
        if same_trade:
            return ""
        if da in fb["alts"] or db in fa["alts"]:
            return "contain"
        short, long_, fs, fl = (da, db, fa, fb) if len(da) <= len(db) else (db, da, fb, fa)
        # 숫자가 더 붙은 이름('신사' / '신사2', '1단지' / '11단지')은 다른 현장이므로 숫자 구성이 같을 때만 포함 관계로 본다
        if fs["digits"] == fl["digits"] and (_contained(short, long_) or (gap <= PARAMS["near_days"] and long_.startswith(short))):
            return "contain"
        return ""
    child, main, pc, pm = (fa, fb, pa, pb) if da else (fb, fa, pb, pa)
    # 구분어가 빈 본공사 ↔ 구분어가 있는 공고: 자기 목적어를 가진 별개 사업('○○ 신축', '△△ 조성')이면 연결하지 않는다.
    # 다만 본공사와 거의 같은 때 난 공고(권역별·구역별 분할 발주)나, 본공사 이름 뒤에 세부 이름을 붙인 공고('…리모델링 사업 클라이밍장 조성공사')는 연결
    if len(main["base"]) >= 6 and child["base"].startswith(main["base"]):
        return "main"
    if main.get("site_prep"):
        return ""                                      # '부지'를 지워서 구분어가 빈 공고(지구 전체의 부지조성)는 이름이 다른 공고의 본공사 노릇을 하지 않는다
    if not child["verb"]:
        # 부대공사로 보이는 이름('연결통로공사', '자동관수시설공사'): 신축은 준공 때까지 부대공사가 이어지므로 허용 간격 전체,
        # 증축·리모델링 시설은 이름이 다른 소규모 공사가 대개 별개 사업이므로 본공사와 가까운 때에만
        if (pc > 0 and pm > 0 and pc > pm * PARAMS["anc_max_ratio"]) or len(child["disc"]) > PARAMS["anc_max_len"]:
            return ""                                  # 본공사 공고보다 훨씬 크거나 긴 자기 이름을 가진 공고는 부대공사가 아니다(시설명이 우연히 걸린 무관한 공사 방지)
        if child.get("ntr", 1) >= 3 and main.get("ntr0", 0) >= 2:
            return ""                                  # 그 이름으로 공종 공고가 3개 이상 따로 난 공사는 자기 공종 묶음을 가진 별개 사업
        return "main" if gap <= (window if new_build else PARAMS["near_days"]) else ""
    # 자기 사업어가 있는 이름: 본공사('○○ 건립 전기공사'처럼 사업어가 있는 공고)와 거의 같은 때 난 다른 공종 공고만.
    # 같은 공종이면 범위가 다른 별개 계약이고, 구분어가 빈 쪽이 부대공사('정화조 설치공사')면 본공사 노릇을 하지 못한다
    return "main" if (gap <= PARAMS["near_days"] and not same_trade and main["verb"]) else ""


def _assign_group(g: pd.DataFrame) -> Dict:
    g = g.sort_values(["_dt", "공고번호"])
    idx: List = list(g.index)
    fac = g["시설명"].iloc[0] if "시설명" in g.columns else ""
    new_build = ("사업유형" not in g.columns) or str(g["사업유형"].iloc[0]) == "신축"
    feat = {i: features(g.at[i, "공고명"], fac) for i in idx}
    n_disc: Dict = {}
    for f in feat.values():
        n_disc[f["disc"]] = n_disc.get(f["disc"], 0) + 1
    trades_of: Dict = {}
    for i, f in feat.items():
        trades_of.setdefault(f["disc"], set()).add(str(g.at[i, "공종"]))
    for f in feat.values():
        f["dn"] = n_disc[f["disc"]]                     # 같은 구분어를 쓰는 공고 수(오타 판정에 사용)
        f["trs"] = frozenset(trades_of[f["disc"]])      # 그 구분어로 난 공고들의 공종
        f["ntr"] = len(trades_of[f["disc"]])            # 그 구분어로 난 공고의 공종 수
        f["ntr0"] = len(trades_of.get("", ()))          # 구분어가 빈(본공사) 공고의 공종 수
    pos = {i: n for n, i in enumerate(idx)}
    parent = {i: i for i in idx}

    def find(i):
        while parent[i] != i:
            parent[i] = parent[parent[i]]
            i = parent[i]
        return i

    def union(i, j):
        a, b = find(i), find(j)
        if a != b:
            if pos[a] > pos[b]:
                a, b = b, a
            parent[b] = a

    by_no = {str(g.at[i, "공고번호"]).strip(): i for i in idx}
    prev = {}
    for i in idx:
        pv = str(g.at[i, "이전공고번호"] or "").strip() if "이전공고번호" in g.columns else ""
        prev[i] = "" if pv in ("nan", "None", "<NA>", "NaT") else pv
        if prev[i] and prev[i] in by_no:               # API 이전공고번호로 연결된 공고는 같은 세부사업
            union(i, by_no[prev[i]])
    price = {i: (float(g.at[i, "_price"]) if pd.notna(g.at[i, "_price"]) else 0.0) for i in idx}
    tk = {i: title_key(g.at[i, "공고명"]) for i in idx}
    # '(2차 재공고)' 는 보통 '두 번째 재공고'(공고 회차)라서 지우지만, 같은 이름의 '(2차)' 공고가 따로 있으면 그 2차 공고의 재공고다 → 그 공고와 같은 키로 맞춘다
    # (그러지 않으면 차수 표기 없는 1차 공고가 대체되고 2차 공고가 두 번 합산된다)
    for i in idx:
        nm = str(g.at[i, "공고명"] or "")
        if _ORD_REBID_BR.search(nm):
            for fmt in (lambda x: x, lambda x: re.sub(r"제\s*", "", x), lambda x: "제" + re.sub(r"제\s*", "", x)):      # '제2차 재공고' ↔ '(2차)', '2차 재공고' ↔ '(제2차)'
                lot_key = title_key(_ORD_REBID_BR.sub(lambda m: "(" + fmt(m.group(1) or m.group(2)) + ")", nm))
                if lot_key != tk[i] and any(k != i and tk[k] == lot_key and g.at[k, "공종"] == g.at[i, "공종"] and g.at[k, "_dt"] <= g.at[i, "_dt"] for k in idx):
                    tk[i] = lot_key
                    break
    tk0 = {i: _PUNCT.sub("", _strip_admin(g.at[i, "공고명"], True)) for i in idx}        # 괄호로만 적은 'N차' 를 뺀 키
    annual = {i: is_annual(g.at[i, "공고명"]) for i in idx}
    members = {i: [i] for i in idx}
    for i in idx:                                       # 이전공고번호로 이미 묶인 것 반영
        r = find(i)
        if r != i:
            members[r].append(i)
            members.pop(i, None)

    def is_rebid_notice(i) -> bool:
        """공고 자체가 재공고임을 밝힌 경우: 공고종류 '재공고', 재공고여부 'Y', 공고명의 '재공고·재입찰·새로운 입찰'."""
        if feat[i]["rebid"]:
            return True
        if "공고종류" in g.columns and "재공고" in str(g.at[i, "공고종류"]):
            return True
        return "재공고여부" in g.columns and str(g.at[i, "재공고여부"]).strip().upper() == "Y"

    def rebid_limit(i, j) -> int:
        """같은 공고명으로 다시 난 공고를 재공고로 볼 수 있는 최대 간격[일]. 유지관리·연간단가 공사는 더 짧게 —
        그보다 떨어진 앞 공고는 지난 회차의 계약이다(뒤 공고가 '재공고' 여도 그 원 공고는 이번 회차에 따로 났다가 취소·누락된 것)."""
        return PARAMS["recur_max_days"] if (annual[i] and annual[j]) else PARAMS["rebid_max_days"]

    def retender(i, j) -> bool:
        """한 번 하는 공사(건립·신축·증축·조성 …)가 같은 공고명·같은 공종으로 다시 난 경우(사업이 늦어져 1~3년 뒤 재발주한 것 포함).
        연례 사업(관리·유지·연간단가, 연도 표기)은 해당하지 않는다."""
        if g.at[i, "공종"] != g.at[j, "공종"] or tk[i] != tk[j] or not (feat[i]["oneoff"] and feat[j]["oneoff"]):
            return False
        if abs((g.at[j, "_dt"] - g.at[i, "_dt"]).days) > PARAMS["oneoff_max_days"]:
            return False
        ratio = (price[j] / price[i]) if (price[i] > 0 and price[j] > 0) else 1.0
        return PARAMS["oneoff_lo"] <= ratio <= PARAMS["oneoff_hi"]

    def apart(i, j) -> bool:
        """같은 공종에서 (거의) 같은 공고명으로 300일 넘게 떨어져 다시 난 공고(연례 사업의 다음 해 분)는 한 세부사업이 될 수 없다.
        같은 사업의 뒤늦은 분할 발주라면 공종이 다르거나 이름에 다른 내용(부대공사 이름)이 붙는다. 한 번 하는 공사의 재발주(retender)는 예외."""
        if g.at[i, "공종"] != g.at[j, "공종"] or abs((g.at[j, "_dt"] - g.at[i, "_dt"]).days) <= rebid_limit(i, j):
            return False
        if (prev[i] and prev[i] == str(g.at[j, "공고번호"]).strip()) or (prev[j] and prev[j] == str(g.at[i, "공고번호"]).strip()):
            return False
        if retender(i, j):
            return False
        if feat[i]["phases"] != feat[j]["phases"] and feat[i]["oneoff"] and feat[j]["oneoff"]:
            return False                               # 한 번 하는 공사의 차수·단계 표기가 다른 공고('조성사업' / '조성사업(2차분)')는 차수별 발주 → 같은 세부사업일 수 있다
        return tk[i] == tk[j] or similarity(tk[i], tk[j]) >= PARAMS["repeat_sim"]

    def years_clash(a, b) -> bool:
        """묶음 a 의 공고가 묶음 b 의 사업 연도와 맞지 않는지. 연도 표기가 없는 공고도 한번 어느 해 사업에 묶이면 그 해를 따른다
        (2022년 사업의 전기 공고에 연도 표기가 없어도 2021년 사업과 이어지지 않게)."""
        yb = set().union(*[feat[y]["years"] for y in members[b]])
        if not yb:
            return False
        fy = {"years": frozenset(yb)}
        return any(not _year_ok(feat[x], fy, g.at[x, "_dt"], g.at[x, "_dt"]) for x in members[a])

    def try_union(i, j) -> None:
        a, b = find(i), find(j)
        if a == b or any(apart(x, y) for x in members[a] for y in members[b]) or years_clash(a, b) or years_clash(b, a):
            return
        union(i, j)
        r = find(i)
        other = b if r == a else a
        members[r] = members[r] + members.pop(other)

    # 연결 후보를 근거가 강한 순(같은 구분어 → 포함 → 본공사), 같은 근거 안에서는 공고일이 가까운 순으로 잇는다
    # (연도 표기가 없는 공종 공고가 먼 해의 사업보다 가까운 해의 사업에 먼저 묶이도록)
    rank = {"same": 0, "contain": 1, "main": 2}
    edges = []
    for n, i in enumerate(idx):
        for j in idx[n + 1:]:
            gap = abs((g.at[j, "_dt"] - g.at[i, "_dt"]).days)
            if retender(i, j):                          # 같은 이름으로 다시 발주한 한 번짜리 공사는 규모별 허용 간격을 넘어도 같은 사업
                edges.append((-1, gap, pos[i], pos[j], i, j))
                continue
            win = _window(max(price[i], price[j]))
            if gap > win or not _year_ok(feat[i], feat[j], g.at[i, "_dt"], g.at[j, "_dt"]):
                continue
            kind = _link(feat[i], feat[j], gap, win, str(g.at[i, "공종"]), str(g.at[j, "공종"]), new_build, price[i], price[j])
            if kind:
                edges.append((rank[kind], gap, pos[i], pos[j], i, j))
    for _, _, _, _, i, j in sorted(edges, key=lambda e: e[:4]):
        try_union(i, j)
    cl: Dict = {}
    for i in idx:
        cl.setdefault(find(i), []).append(i)
    order = sorted(cl, key=lambda r: pos[r])
    work = {i: n + 1 for n, r in enumerate(order) for i in cl[r]}
    role = {i: "count" for i in idx}
    why = {i: "" for i in idx}
    repl = {i: "" for i in idx}
    def supersede(old, new, reason) -> None:
        if role[old] == "superseded":
            return
        t, seen = new, set()                            # new 가 (여러 단계를 거쳐) old 로 대체되어 있으면 서로를 가리키는 잘못된 연결 → 하나는 남긴다
        while t is not None and role[t] == "superseded" and t not in seen:
            seen.add(t)
            t = by_no.get(repl[t])
        if t == old:
            return
        role[old], repl[old], why[old] = "superseded", str(g.at[new, "공고번호"]).strip(), reason

    ann = set()                                         # 괄호 'N차' 가 차수가 아니라 공고 회차(몇 번째 공고)임이 드러난 공고

    def round_ord(i) -> bool:
        """i 의 괄호 'N차' 가 공고 회차인지: 앞에서 그렇게 판정됐거나, 재공고라고 밝힌 공고인데 같은 이름(같은 차수)의 앞 공고가 없을 때.
        같은 이름의 앞 공고가 있으면 그 차수 공고의 재공고이므로 차수 표기다('(1차)' → '(1차)' 재공고 → 금액이 같은 '(2차)' 는 다른 차수)."""
        return i in ann or (is_rebid_notice(i) and not any(pos[k] < pos[i] and tk[k] == tk[i] and work[k] == work[i] and g.at[k, "공종"] == g.at[i, "공종"] and k not in ann for k in idx))

    for n, i in enumerate(idx):
        for j in idx[n + 1:]:                           # j 는 i 보다 공고일시가 같거나 뒤
            if work[i] != work[j]:
                continue
            no_i, no_j = str(g.at[i, "공고번호"]).strip(), str(g.at[j, "공고번호"]).strip()
            if prev[j] and prev[j] == no_i:             # API: j 가 i 의 재공고
                supersede(i, j, "API 이전공고번호 연결")
                continue
            if prev[i] and prev[i] == no_j:             # API: i 가 j 의 재공고인데 j 의 변경 차수가 더 늦게 난 경우 → 대체되는 쪽은 j
                supersede(j, i, "API 이전공고번호 연결")
                continue
            if g.at[i, "공종"] != g.at[j, "공종"] or role[i] == "superseded":
                continue
            days = (g.at[j, "_dt"] - g.at[i, "_dt"]).days
            ratio = (price[j] / price[i]) if (price[i] > 0 and price[j] > 0) else 1.0
            flagged = is_rebid_notice(j) and g.at[j, "_dt"] > g.at[i, "_dt"]      # 재공고라고 밝힌 공고는 같은 날·다음 날이어도 재공고
            in_time = (0 if flagged else PARAMS["rebid_min_days"]) <= days <= rebid_limit(i, j)
            # 괄호로만 적은 '(2차)' 는 차수 발주일 수도, '두 번째 공고'(재공고)일 수도 있다. 금액이 그대로이고(차수 발주는 금액이 다르다),
            # 앞 공고가 차수 표기 없는 같은 이름(또는 그 재공고)이며, 그 차수의 원 공고가 따로 없을 때만 재공고로 본다.
            # 두 공고 사이에 금액이 다른 '(N차)' 공고가 있으면 차수 발주다('', (2차) 60억, (3차) 50억 — 1차분과 3차분의 금액이 같을 뿐)
            loose = (tk[i] != tk[j] and tk0[j] != tk[j] and tk0[i] == tk0[j] and (tk[i] == tk0[i] or j in ann or round_ord(i))
                     and price[i] > 0 and price[j] > 0 and 0.99 <= ratio <= 1.01 and (flagged or days <= PARAMS["near_days"])
                     and not any(pos[i] < pos[k] < pos[j] and tk0[k] == tk0[j] and tk[k] != tk0[k] and g.at[k, "공종"] == g.at[j, "공종"]
                                 and price[k] > 0 and not 0.99 <= price[k] / price[j] <= 1.01 for k in idx)
                     and not any(k != j and tk[k] == tk[j] and work[k] == work[j] and g.at[k, "공종"] == g.at[j, "공종"]
                                 and not (pos[k] > pos[j] and round_ord(j)) for k in idx))      # 재공고라고 밝힌 첫 '(2차)' 뒤에 같은 이름으로 또 난 공고는 그 차수의 원 공고가 아니다
            exact = (tk[i] == tk[j] or loose) and in_time and PARAMS["rebid_lo"] <= ratio <= PARAMS["rebid_hi"]
            late = (not exact and retender(i, j) and days >= (0 if flagged else PARAMS["rebid_min_days"]))
            fuzzy = (not exact and not late and in_time and PARAMS["rebid_fuzzy_lo"] <= ratio <= PARAMS["rebid_fuzzy_hi"] and price[i] > 0
                     and re.findall(r"\d+", tk[i]) == re.findall(r"\d+", tk[j]) and similarity(tk[i], tk[j]) >= PARAMS["rebid_fuzzy"]
                     and feat[i]["disc"] == feat[j]["disc"] and feat[i]["site"] == feat[j]["site"])
            if exact and loose:
                ann.add(j)
            elif exact and tk[i] != tk0[i] and round_ord(i):
                ann.update((i, j))                      # 공고 회차 '(2차)' 를 같은 이름으로 한 번 더 낸 공고도 공고 회차
            if exact or late or fuzzy:
                supersede(i, j, "같은 공고명으로 다시 공고" if exact else ("같은 공고명으로 다시 발주(사업 지연 뒤 재발주)" if late else "거의 같은 공고명·금액으로 다시 공고"))
    # 재공고가 거듭된 경우(A ← B ← C) 대체공고번호는 끝까지 남은 공고(C)를 가리키게 한다
    for i in idx:
        if role[i] == "superseded":
            seen, t = {i}, by_no.get(repl[i])
            while t is not None and role[t] == "superseded" and t not in seen:
                seen.add(t)
                t = by_no.get(repl[t])
            if t is not None and role[t] != "superseded":
                repl[i] = str(g.at[t, "공고번호"]).strip()
    names = {}
    for r in order:
        mem = sorted(cl[r], key=lambda i: (0 if g.at[i, "공종"] == "건축" else 1, -price[i], pos[i]))
        names[work[r]] = clean_title(g.at[mem[0], "공고명"])
    return {i: (work[i], role[i], names[work[i]], why[i], repl[i]) for i in idx}


def assign_works(d: pd.DataFrame, group_col: str) -> pd.DataFrame:
    """유효 공고(취소·구차수 제외) d 를 group_col(시설ID×사업유형) 별로 세부사업에 배정.
    필요한 열: group_col, 공고번호, 공고명, 공종, _dt(공고일시), _price(추정가격), 시설명·이전공고번호(선택).
    반환: d 와 같은 index 의 DataFrame[세부사업번호(1부터, 최초 공고일 순), 역할('count'|'superseded'), 세부사업명, 사유, 대체공고번호]."""
    out: Dict = {}
    for _, g in d.groupby(group_col, sort=False):
        out.update(_assign_group(g))
    return pd.DataFrame.from_dict(out, orient="index", columns=["세부사업번호", "역할", "세부사업명", "사유", "대체공고번호"]).reindex(d.index)
