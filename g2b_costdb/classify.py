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
_STATUS_TOKENS = r"(재공고|재입찰|긴급|변경|취소|정정|수정|연기|일부변경|재발주|2차공고|3차공고)"
_BRACKET = re.compile(r"[\(\[\{（【].*?[\)\]\}）】]")
_PHASE = re.compile(r"(\d+\s*단계|\d+\s*차|\d+\s*공구|\d+\s*차분|골조|마감)")
_TRADE_WORDS = re.compile(
    r"(건축\s*공사|전기\s*공사|정보통신\s*공사|통신\s*공사|소방\s*공사|소방시설\s*공사|조경\s*공사|기계설비\s*공사|"
    r"토목\s*공사|건립\s*공사|신축\s*공사|증축\s*공사|건설\s*공사|조성\s*공사|공사|건립|신축|증축|건설|조성|사업|용역)"
)
_YEAR = re.compile(r"(20\d{2})\s*년(도)?")
_SPACES = re.compile(r"\s+")


def clean_notice_name(name: str) -> str:
    """상태어·괄호·연도 제거 후 공백 정리 (시설명 추출·프로젝트키용)."""
    s = name or ""
    s = _BRACKET.sub(" ", s)
    s = re.sub(_STATUS_TOKENS, " ", s)
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
        return s
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
        head = head[: kw_pos + len(kw_head)]
    parts = tokens[start:idx] + [head]
    return " ".join(parts).strip()


# ── 분류 ────────────────────────────────────────────────────────
WORK_TYPE_CODE = {"신축": "N", "증축": "E", "리모델링": "R", "증축·리모델링": "ER", "유지보수": "M", "미분류": "U"}


def classify_work_type(notice_name: str, rules: Optional[Dict[str, List[str]]] = None) -> str:
    """사업유형: 신축 / 증축 / 리모델링 / 증축·리모델링 / 유지보수 / 미분류.
    yaml 순서 = 우선순위(리모델링 → 증축 → 신축 → 유지보수). 증축과 리모델링 어휘가 함께 있으면 '증축·리모델링'."""
    rules = rules or load_keywords()["work_type_rules"]
    s = (notice_name or "").replace(" ", "")
    hits = [label for label, words in rules.items() if any(w.replace(" ", "") in s for w in words)]
    if "리모델링" in hits and "증축" in hits:
        return "증축·리모델링"
    for label in rules:                      # 첫 매칭 우선
        if label in hits:
            return label
    return "미분류"


def classify_trade(notice_name: str = "", main_cnstty: str = "", license_names: str = "",
                   rules: Optional[Dict[str, List[str]]] = None) -> Tuple[str, str]:
    """공종 분류. 우선순위: 면허제한 업종명 → 주공종명 → 공고명. 반환 (공종, 근거)."""
    rules = rules or load_keywords()["trade_rules"]

    def _match(text: str) -> Optional[str]:
        t = (text or "").replace(" ", "")
        if not t:
            return None
        hits = [label for label, words in rules.items() if any(w.replace(" ", "") in t for w in words)]
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
