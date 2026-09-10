"""설정 로딩 + 규칙 기반 분류(공종 / 사업유형 / 공고종류 / 시설명 정규화)."""
from __future__ import annotations

import os
import re
from functools import lru_cache
from typing import Dict, List, Optional, Tuple

import yaml

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))


@lru_cache(maxsize=1)
def load_config(path: Optional[str] = None) -> dict:
    """config.yaml 로드. paths.* 의 상대경로는 현재 작업 폴더가 아니라 프로젝트 루트(config.yaml 위치) 기준으로 해석한다."""
    cfg_path = path or os.path.join(ROOT, "config.yaml")
    with open(cfg_path, encoding="utf-8") as f:
        cfg = yaml.safe_load(f)
    base = os.path.dirname(os.path.abspath(cfg_path))
    for k, v in list((cfg.get("paths") or {}).items()):
        if isinstance(v, str) and k != "excel_name" and not os.path.isabs(v):
            cfg["paths"][k] = os.path.normpath(os.path.join(base, v))
    return cfg


@lru_cache(maxsize=1)
def load_keywords(path: Optional[str] = None) -> dict:
    with open(path or os.path.join(ROOT, "keywords.yaml"), encoding="utf-8") as f:
        return yaml.safe_load(f)


# ── 정규화 ──────────────────────────────────────────────────────
_STATUS_TOKENS = (r"(재공고|재입찰|긴급|변경|취소|정정|수정|연기|일부변경|재발주|\d차공고|수의계약|수의견적|소액수의|수액수의|전자견적|견적\s*제출|안내\s*공고|"
                  r"입찰\s*공고|업체\s*선정|제한경쟁|일반경쟁|장기계속|민간입찰대행|입찰대행|전자입찰|일부특허|혁신|총괄|전체분|\d+차분|국체전\s*대비|전국체전\s*대비|"
                  r"전국체육대회|\d{4}년도?|"
                  r"견적공고|견적제출|제출안내|(?<![가-힣])(입찰|공고|제출|안내|시행|견적서|견적)(?![가-힣]))")      # 홀로 선 상태어만('부산공고'·'안내판' 같은 이름의 일부는 보존)
_BRACKET_INNER = re.compile(r"[\(（]([^()（）\[\]{}【】]*)[\)）]|[\[【]([^()（）\[\]{}【】]*)[\]】]|\{([^()（）\[\]{}【】]*)\}")
# 상태어만 남은 괄호 밖 문구('시설공사 수의견적 제출 안내 공고[…]' → '시설') 판정용
_GENERIC_OUTER = re.compile(r"^(학교|시설|공사|건축|기타|긴급|소액|수의|견적|제출|안내|공고|시행|사업|시행공고|외\d*\S*)*$")
_PHASE = re.compile(r"(\d+\s*단계|\d+\s*차|\d+\s*공구|\d+\s*차분|골조|마감)")
_TRADE_WORDS = re.compile(
    r"(건축\s*공사|전기\s*공사|정보통신\s*공사|통신\s*공사|소방\s*공사|소방시설\s*공사|조경\s*공사|기계설비\s*공사|기계\s*공사|설비\s*공사|"
    r"토목\s*공사|건립\s*공사|신축\s*공사|증축\s*공사|건설\s*공사|조성\s*공사|공사|건립|신축|증축|건설|조성|사업|용역)"
)
# '2019년도', '23년~24년', '25~26년', '26,27년' 같은 연도 접두어(뒤에 글자가 붙는 '2024목동주경기장' 은 건드리지 않음)
_YEAR = re.compile(r"(?<![\d가-힣])[′'‘]?(?:20)?\d{2}(?:\s*년도?)?(?:\s*[~∼\-,]\s*[′'‘]?(?:20)?\d{2})*\s*년도?(?![\d가-힣])")
_SPACES = re.compile(r"\s+")


def strip_brackets(text: str) -> str:
    """괄호를 안쪽부터 반복 제거(중첩 '[…(리모델링)공사 (전기공사)]' 도 남김없이)."""
    s = text or ""
    for _ in range(6):
        s2 = _BRACKET_INNER.sub(" ", s)
        if s2 == s:
            break
        s = s2
    return s


def bracket_contents(text: str) -> List[str]:
    """최상위 괄호 안 문구 목록(중첩 괄호는 안쪽을 포함한 채로). '공고[A(리모델링) (전기)]' → ['A(리모델링) (전기)']"""
    opens, closes = "([{（【", ")]}）】"
    out, depth, buf = [], 0, []
    for ch in text or "":
        if ch in opens:
            if depth > 0:
                buf.append(ch)
            depth += 1
        elif ch in closes and depth > 0:
            depth -= 1
            if depth == 0:
                out.append("".join(buf)); buf = []
            else:
                buf.append(ch)
        elif depth > 0:
            buf.append(ch)
    return out


def clean_notice_name(name: str) -> str:
    """상태어·괄호·연도 제거 후 공백 정리 (시설명 추출·프로젝트키용)."""
    s = name or ""
    s = strip_brackets(s)
    s = re.sub(r"[()\[\]{}（）【】]", " ", s)          # 짝이 없는 괄호('구)문화예술회관', '긴급) …')는 공백으로
    for _ in range(3):                      # '입찰 취소 공고' 처럼 상태어가 겹쳐 있으면 반복 제거
        s2 = re.sub(_STATUS_TOKENS, " ", s)
        if s2 == s:
            break
        s = s2
    s = _YEAR.sub(" ", s)
    s = re.sub(r"[「」『』<>《》\"'“”‘’]", " ", s)
    s = _SPACES.sub(" ", s).strip(" -–—·,.")
    return s


def normalize_facility_key(facility: str) -> str:
    """프로젝트키용 시설명 정규화: 공백 제거, 단계/차수 토큰은 보존."""
    s = clean_notice_name(facility)
    phases = "".join(p.replace(" ", "") for p in _PHASE.findall(s))
    s = _TRADE_WORDS.sub(" ", s)
    s = _SPACES.sub("", s)
    return (s + ("|" + phases if phases else "")).lower()


def extract_facility_name(notice_name: str, keyword: str) -> str:
    """공고명에서 검색어를 포함한 어절과 그 앞 어절(지자체·수식어)을 묶어 시설명 후보 생성.
    예) '고흥군 문화예술회관 건립공사(건축)' + '문화예술회관' → '고흥군 문화예술회관'
        '수원시립미술관 건립 전기공사' + '미술관' → '수원시립미술관'
    """
    s = clean_notice_name(notice_name)
    tokens = s.split(" ")
    kw_parts = keyword.split()                       # '온실 건립' 처럼 띄어쓴 검색어는 첫 어절로 위치를 찾는다
    kw_head = kw_parts[0] if kw_parts else keyword
    idx = next((i for i, t in enumerate(tokens) if kw_head in t), None)
    if idx is None:
        idx = next((i for i, t in enumerate(tokens) if keyword.replace(" ", "") in t), None)
    if idx is None:
        # 검색어가 어절 두세 개에 걸쳐 있는 경우('지하 공영주차장' ← '지하공영주차장') → 어절을 합쳐 찾는다
        kw_ns = keyword.replace(" ", "")
        for span in (2, 3):
            for i in range(len(tokens) - span + 1):
                if kw_ns in "".join(tokens[i:i + span]):
                    tokens[i:i + span] = [" ".join(tokens[i:i + span])]
                    idx, kw_head = i, tokens[i]
                    break
            if idx is not None:
                break
    if idx is None:
        # 검색어가 괄호 안에만 있는 경우: '제2안식의 집(봉안당) 건립공사' → 괄호 밖 이름('제2안식의 집')을 시설명으로,
        # 괄호 밖이 상태어뿐이면('입찰 취소 공고[진해아트홀 시설 개선공사]') 괄호 안 문구로 다시 추출
        inner = [c for c in bracket_contents(notice_name or "")
                 if kw_head in c or keyword.replace(" ", "") in c.replace(" ", "")]
        outer = _SPACES.sub(" ", _TRADE_WORDS.sub(" ", s)).strip(" -–—·,.")
        outer = " ".join(_tidy_parts(outer.split(" ")))
        outer_core = re.sub(r"[^가-힣A-Za-z0-9]", "", outer)
        if inner and (len(re.sub(r"[^가-힣A-Za-z]", "", outer)) < 2 or _GENERIC_OUTER.fullmatch(outer_core)):
            return extract_facility_name(inner[0], keyword)
        return outer if outer else s
    # 앞 어절 중 지자체/수식어로 보이는 것 최대 2개 포함 (공종어·숫자만 있는 어절·'제1' 같은 순번 제외)
    start = idx
    for j in range(idx - 1, max(-1, idx - 3), -1):
        t = tokens[j]
        if _TRADE_WORDS.fullmatch(t) or re.fullmatch(r"[\d\-\.]+", t) or re.match(r"제\s*\d", t):
            break
        start = j
    # 검색어가 포함된 어절에서 뒤에 붙은 공종어 제거 ('미술관건립공사' → '미술관')
    head = tokens[idx]
    kw_pos = head.find(kw_head)
    if kw_pos >= 0:
        # 검색어 뒤에 공종·사업유형 어휘가 붙어 있으면 거기서 자르고('미술관건립공사' → '미술관', '종합운동장공원조성사업' → '종합운동장공원'),
        # 이름이 이어지면 보존('호텔인터시티' → '호텔인터시티')
        m = _TRADE_WORDS.search(head, kw_pos + len(kw_head))
        if m:
            head = head[: m.start()]
    tail = []
    if len(kw_parts) > 1 and head.endswith(kw_head):
        # '스마트팜 온실 설비…' 처럼 띄어쓴 검색어의 나머지 어절이 이어지면 시설명에 포함('스마트팜' → '스마트팜 온실')
        for k, part in enumerate(kw_parts[1:], start=1):
            # '온실 건립' 의 '건립' 처럼 공종·사업유형 어휘는 시설명이 아니다
            if idx + k < len(tokens) and tokens[idx + k].startswith(part) and not _TRADE_WORDS.fullmatch(part):
                tail.append(part)
            else:
                break
    return " ".join(_tidy_parts(tokens[start:idx] + [head] + tail)).strip()


def _tidy_parts(parts: List[str]) -> List[str]:
    """시설명 어절 정리: 기호·숫자뿐인 어절('′23∼′') 제거, 앞의 '및/외/내/등' 제거, 끝에 남은 접두어 조각('월배공원 재') 제거."""
    parts = [t for t in parts if t and re.search(r"[가-힣A-Za-z]", t)]
    while parts and parts[0] in ("및", "외", "내", "등", "-", "·"):
        parts.pop(0)
    while len(parts) > 1 and parts[-1] in ("재", "구", "신", "본", "제", "舊"):   # '집'·'관' 같은 진짜 이름 어절은 남긴다
        parts.pop()
    return parts


# ── 단지명 없는 시설명 보완 ───────────────────────────────────────
# 검색어와 일반어(아파트·주택·지구·단지…)만 남은 시설명('영구임대아파트', '공공주택지구')은 어느 시설인지 알 수 없으므로 수요기관의 지자체명을 앞에 붙인다
_GENERIC_NAME_WORDS = ("아파트", "주택", "지구", "단지", "사업", "시설", "공사", "노후", "공공", "임대", "통합", "신축", "건립", "건설", "조성",
                       "공영", "공설", "시립", "군립", "구립", "도립", "국립", "종합", "생활", "체육", "다목적", "제", "구", "신", "본", "동", "관내")
_INST_ORG = re.compile(r"사업소|공사|공단|본부|청$|센터|재단|학교|대학|병원|조합|협회|위원회|연구|사업단")


def is_generic_facility_name(facility: str, keyword: str) -> bool:
    """시설명이 검색어 + 일반어뿐인지('영구임대아파트' + '영구임대' → True, '마음에온 일도1차 통합공공임대주택' → False)."""
    s = (facility or "").replace(" ", "")
    for w in sorted({keyword.replace(" ", "")} | set(keyword.split()), key=len, reverse=True):
        if w:
            s = s.replace(w, "")
    for w in sorted(_GENERIC_NAME_WORDS, key=len, reverse=True):
        s = s.replace(w, "")
    return len(re.sub(r"[^가-힣A-Za-z]", "", s)) == 0


def short_institution(inst: str) -> str:
    """수요기관에서 시설명 앞에 붙일 짧은 이름: 가장 구체적인 시·군·구('전남광주통합특별시 구례군' → '구례군', '부산광역시 강서구' → '강서구').
    시·군·구 어절이 없으면 첫 어절('서울주택도시개발공사', '경상북도교육청')."""
    tokens = [t for t in re.split(r"\s+", re.sub(r"[\(（](주|재|사|학|의|특)[\)）]", "", inst or "").strip()) if t]   # '(주)강원랜드' → '강원랜드'
    if not tokens:
        return ""
    cands = [t for t in tokens if re.fullmatch(r".{1,12}(시|군|구)", t) and not _INST_ORG.search(t)]
    return cands[-1] if cands else tokens[0]


def prefix_institution(facility: str, keyword: str, inst: str) -> str:
    """단지명 없는 시설명이면 지자체명을 앞에 붙인다. ('영구임대아파트', '영구임대', '대구도시개발공사') → '대구도시개발공사 영구임대아파트'"""
    short = short_institution(inst)
    if not short or not is_generic_facility_name(facility, keyword):
        return facility
    if (facility or "").replace(" ", "").startswith(short.replace(" ", "")):
        return facility
    return f"{short} {facility}".strip()


def strip_institution_prefix(name: str, inst: str) -> str:
    """검수_시설명 앞에 붙은 지자체명을 떼어 공고명 검색용 패턴을 만든다(공고명에는 지자체명이 없는 경우가 많음). 없으면 그대로."""
    short = short_institution(inst).replace(" ", "")
    n = (name or "").strip()
    core = re.sub(r"[\(（](주|재|사|학|의|특)[\)）]", "", n).replace(" ", "")      # '(주)강원랜드 유리온실' 도 '강원랜드' 접두로 인식
    if short and core.startswith(short) and len(core) > len(short) + 1:
        return core[len(short):]
    return n


# ── 분류 ────────────────────────────────────────────────────────
WORK_TYPE_CODE = {"신축": "N", "증축": "E", "리모델링": "R", "증축·리모델링": "ER", "유지보수": "M", "미분류": "U"}


def classify_work_type(notice_name: str, rules: Optional[Dict[str, List[str]]] = None,
                       ancillary: Optional[List[str]] = None) -> str:
    """사업유형: 신축 / 증축 / 리모델링 / 증축·리모델링 / 유지보수 / 미분류.
    yaml 순서 = 우선순위(리모델링 → 증축 → 신축 → 유지보수). 증축과 리모델링 어휘가 함께 있으면 '증축·리모델링'.
    부속·외부 공사 어휘(진입로·주차장·조명·설비 등)가 있으면: 리모델링/증축은 유지, 신축이면 '미분류'(검토필요), 아니면 '유지보수'."""
    kw = None
    if rules is None or ancillary is None:
        kw = load_keywords()
    rules = rules or kw["work_type_rules"]
    ancillary = ancillary if ancillary is not None else (kw.get("ancillary_words") or [])
    s = (notice_name or "").replace(" ", "")
    hits = [label for label, words in rules.items() if any(w.replace(" ", "") in s for w in words)]
    # '기계설비공사'·'건축설비공사'는 공종명이지 부속 설비 공사가 아니다 → 부속어 판정에서 제외
    s_anc = re.sub(r"기계설비|건축설비|전기설비|소방설비|통신설비", "", s)
    anc = any(w.replace(" ", "") in s_anc for w in ancillary)
    if "리모델링" in hits and "증축" in hits:
        return "증축·리모델링"
    for label in rules:                      # 첫 매칭 우선
        if label in hits:
            if label == "신축" and anc:
                return "미분류"
            return label
    return "유지보수" if anc else "미분류"


def work_type_for(notice_name: str, keyword: str = "", facility_name: str = "") -> str:
    """검색어·시설명에 들어 있는 부속어는 제외하고 사업유형 판정.
    예) '지하 공영주차장 조성 소방공사' 는 시설 자체가 주차장이므로 '주차장' 을 부속어로 보지 않는다 → 신축."""
    kw = load_keywords()
    ctx = ((keyword or "") + (facility_name or "")).replace(" ", "")
    anc = [w for w in (kw.get("ancillary_words") or []) if w.replace(" ", "") not in ctx]
    return classify_work_type(notice_name, kw["work_type_rules"], anc)


def classify_trade(notice_name: str = "", main_cnstty: str = "", license_names: str = "",
                   rules: Optional[Dict[str, List[str]]] = None) -> Tuple[str, str]:
    """공종 분류. 우선순위: 면허제한 업종명 → 주공종명 → 공고명. 반환 (공종, 근거)."""
    rules = rules or load_keywords()["trade_rules"]

    def _match(text: str) -> Optional[str]:
        t = re.sub(r"[\sㆍ·・∙]", "", text or "")     # '기계설비ㆍ가스공사업' 같은 구분자 제거
        if not t:
            return None
        hits = [label for label, words in rules.items() if any(re.sub(r"[\sㆍ·・∙]", "", w) in t for w in words)]
        if not hits:
            return None
        # '토목건축공사업' 처럼 토목·건축이 함께 걸리면 건축(건물 본공사)으로 본다
        if "건축" in hits and "토목" in hits:
            hits.remove("토목")
        for label in rules:                      # 규칙 순서 = 우선순위(건축은 마지막)
            if label in hits and label != "건축":
                return label
        return "건축" if "건축" in hits else None

    for src, text in (("면허제한", license_names), ("주공종명", main_cnstty), ("공고명", notice_name)):
        lab = _match(text)
        if lab:
            return lab, src
    return "기타", "미분류"


def classify_notice_kind(kind: str = "", re_notice: str = "", reg_type: str = "", notice_name: str = "") -> str:
    """'취소' / '변경' / '재공고' / '일반' 판정.
    취소는 공고종류·등록유형(구조화 필드) 또는 공고명의 '취소공고'로만 판정한다 — '(취소 후 재공고)' 같은 유효 공고를 취소로 오인하지 않도록."""
    fields = " ".join([str(kind or ""), str(reg_type or "")])
    name = str(notice_name or "")
    if "취소" in fields:
        return "취소"
    if str(re_notice).upper() == "Y" or "재공고" in fields or "재공고" in name or "재입찰" in name:
        return "재공고"
    if "취소공고" in name.replace(" ", ""):
        return "취소"
    if "변경" in fields or "정정" in fields or re.search(r"[\(\[【]\s*(변경|정정)", name):
        return "변경"
    return "일반"


def match_categories(notice_name: str, kw: Optional[dict] = None, best_only: bool = True) -> List[Tuple[str, str, str]]:
    """공고명에 매칭되는 (대분류, 중분류, 검색어) 목록. exclude 단어 포함 시 제외.
    best_only=True 면 여러 분류에 걸릴 때 가장 구체적인(긴) 검색어 1건만 반환한다
    (예: '야외공연장 조성공사'는 '공연장'과 '야외공연장' 모두 매칭 → '야외공연장'만 채택하여 시설 중복 생성을 막음)."""
    kw = kw or load_keywords()
    s = (notice_name or "").replace(" ", "")
    if any(x.replace(" ", "") in s for x in kw.get("exclude_common", []) or []):   # 모든 분류 공통 제외어(지하철역·도로 공사 등)
        return []
    hits = []
    for cat in kw["categories"]:
        if any(x.replace(" ", "") in s for x in cat.get("exclude", [])):
            continue
        for w in sorted(cat["include"], key=lambda x: -len(x.replace(" ", ""))):
            if w.replace(" ", "") in s:
                hits.append((cat["major"], cat["minor"], w))
                break
    if best_only and len(hits) > 1:
        best = max(hits, key=lambda h: len(h[2].replace(" ", "")))
        # 서로 포함 관계가 아닌 독립 검색어(예: '미술관'+'수장고')는 모두 유지
        hits = [h for h in hits if h == best or (h[2].replace(" ", "") not in best[2].replace(" ", "")
                                                 and best[2].replace(" ", "") not in h[2].replace(" ", ""))]
    return hits


def parse_amount(v) -> Optional[int]:
    """'1,234,567' / 1234567.0 / '' → int 또는 None"""
    if v is None or (isinstance(v, float) and v != v):  # None / NaN
        return None
    s = str(v).replace(",", "").strip()
    if s.lower() in ("nan", "none", "null"):
        return None
    if not s or s in ("0", "0.0"):
        return None
    try:
        return int(float(s))
    except ValueError:
        return None
