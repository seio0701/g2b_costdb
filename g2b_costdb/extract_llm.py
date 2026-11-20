"""S6 LLM 구조화 추출 + API 값 교차검증.

- 공고문 텍스트에서 공사개요·금액구성을 JSON으로 추출 (Claude API, 모델은 config.llm.model)
- 키워드 주변 창을 우선 배치하여 입력 길이 제한(max_input_chars) 안에서 핵심 정보 손실 방지
- 구조화 출력(output_config.format=json_schema)으로 JSON 형식을 보장하고, 미지원 시 일반 응답 파싱으로 폴백
- 추출값 vs API 값(추정가격·기초금액) 차이율, 기초금액≈추정가격×1.1 정합성, 단위 오류 점검 → 검증로그
SDK 사용법·모델명은 https://docs.claude.com/en/api/overview 기준으로 확인.
"""
from __future__ import annotations

import json
import logging
import math
import os
import re
from typing import Dict, List, Optional, Tuple

log = logging.getLogger(__name__)

FIELDS = {
    "공사명": "string", "공사위치": "string", "용도": "string", "사업유형": "신축|증축|리모델링|증축·리모델링|유지보수|기타",
    "연면적_m2": "number|null", "건축면적_m2": "number|null", "대지면적_m2": "number|null",
    "지하층수": "integer|null", "지상층수": "integer|null", "구조": "string",
    "공사기간_일": "integer|null", "공사기간_원문": "string",
    "추정가격_원": "integer|null", "기초금액_원": "integer|null", "부가가치세_원": "integer|null",
    "도급자관급액_원": "integer|null", "관급자관급액_원": "integer|null", "총공사금액_원": "integer|null",
    "설계금액_원": "integer|null", "공종": "건축|전기|정보통신|소방|조경|기계설비|토목|기타",
    "분리발주_언급공종": "string", "포함공종_문서": "string", "근거문구": "object(항목명→원문 인용 30자 이내)", "신뢰도": "high|medium|low",
}
_NUM_FIELDS = ["연면적_m2", "건축면적_m2", "대지면적_m2", "지하층수", "지상층수", "공사기간_일", "추정가격_원", "기초금액_원",
               "부가가치세_원", "도급자관급액_원", "관급자관급액_원", "총공사금액_원", "설계금액_원"]
_STR_FIELDS = ["공사명", "공사위치", "용도", "구조", "공사기간_원문", "분리발주_언급공종", "포함공종_문서"]


def output_schema() -> Dict:
    """구조화 출력용 JSON 스키마. 모든 필드 필수이며 값이 없으면 빈 문자열("") — 숫자 필드도 문자열(자릿수만)로 받아 normalize_doc 이 숫자로 바꾼다.
    API 의 구조화 출력은 union 타입(["number","null"]) 파라미터를 요청당 16개까지만 허용해(초과 시 400 'Schemas contains too many parameters
    with union types') null 대신 빈 문자열을 쓴다. 근거문구는 [{항목, 원문}] 배열로 받아 dict 로 변환."""
    props: Dict[str, Dict] = {}
    for k in _STR_FIELDS:
        props[k] = {"type": "string"}
    for k in _NUM_FIELDS:
        props[k] = {"type": "string"}
    props["사업유형"] = {"type": "string", "enum": ["신축", "증축", "리모델링", "증축·리모델링", "유지보수", "기타"]}
    props["공종"] = {"type": "string", "enum": ["건축", "전기", "정보통신", "소방", "조경", "기계설비", "토목", "기타"]}
    props["신뢰도"] = {"type": "string", "enum": ["high", "medium", "low"]}
    props["근거문구"] = {"type": "array", "items": {"type": "object", "properties": {"항목": {"type": "string"}, "원문": {"type": "string"}},
                                                "required": ["항목", "원문"], "additionalProperties": False}}
    return {"type": "object", "properties": props, "required": list(props.keys()), "additionalProperties": False}


SYSTEM = (
    "당신은 한국 공공건축 공사 입찰공고문에서 공사개요와 공사금액 구성을 정확히 추출하는 분석가입니다. "
    "반드시 JSON 객체 하나만 출력합니다(마크다운 코드블록·설명 금지). 문서에 없는 값은 비워 둡니다(null 또는 빈 문자열 \"\") — 추정하지 않습니다. "
    "숫자 항목은 쉼표·단위 없이 자릿수만 적습니다(예: \"13731.3\", \"23215938082\"). "
    "금액은 원 단위 정수로 변환합니다(예: '1,234백만원'→1234000000, '12.3억원'→1230000000, '천원' 단위 표는 ×1000). "
    "면적은 ㎡ 단위 숫자로 변환하고(평→×3.3058), 층수는 정수로 분리합니다. "
    "기초금액은 통상 추정가격+부가가치세이며, 관급자재(관급자관급액·도급자관급액)는 별도 항목으로 추출합니다. "
    "리모델링·개보수 공사는 '사업유형'을 리모델링으로 하고, 연면적_m2에는 공사 대상 연면적(부분 리모델링이면 대상 부분 면적)을 넣고 "
    "근거문구에 '대상면적'임을 표시합니다. 증축이 함께 있으면 '증축·리모델링'으로 표기합니다. "
    "각 값의 근거가 된 원문 문구를 '근거문구'에 항목명별로 짧게 인용합니다(근거문구는 [{\"항목\": …, \"원문\": …}] 배열). "
    "메타데이터(공고명·수요기관)는 문서 식별용이며, 금액·면적은 반드시 문서 본문에 표기된 값만 추출하고 없으면 비워 둡니다. "
    "'포함공종_문서'에는 공사금액 표의 구분란처럼 문서가 이 계약 금액에 포함된다고 명시한 공종을 쉼표로 적습니다(예: '건축, 토목, 기계, 조경'). "
    "분리발주·별도 발주라고 적힌 공종은 '분리발주_언급공종'에만 적습니다."
)

_KEYS = ["공사개요", "공사 개요", "공사금액", "공사 금액", "규모", "연면적", "건축면적", "층수", "지하", "지상", "구조", "공사기간", "준공",
         "추정가격", "기초금액", "예정가격", "관급", "부가가치세", "도급", "총공사", "설계금액", "공사금액", "사업비"]

# 모델별 1M 토큰당 요금(USD): 입력, 출력 — 견적용(정확한 값은 docs.claude.com 요금표 참조)
PRICE_USD_PER_M = {
    "claude-sonnet-5": (2.0, 10.0), "claude-opus-5": (5.0, 25.0), "claude-opus-4-8": (5.0, 25.0),
    "claude-opus-4-7": (5.0, 25.0), "claude-sonnet-4-6": (3.0, 15.0), "claude-haiku-4-5": (1.0, 5.0),
}


def focus_text(text: str, max_chars: int, window: int = 1200, keys: Optional[List[str]] = None) -> str:
    """앞부분 + 키워드 주변 창을 우선 결합하여 길이 제한 내로 축약. keys: 발췌 키워드(기본 1차 추출용 _KEYS)."""
    if len(text) <= max_chars:
        return text
    head = text[: int(max_chars * 0.35)]
    spans = []
    for k in (keys or _KEYS):
        for m in re.finditer(re.escape(k), text):
            spans.append((max(0, m.start() - window // 3), min(len(text), m.end() + window)))
    spans.sort()
    merged: List[List[int]] = []
    for s, e in spans:
        if merged and s <= merged[-1][1]:
            merged[-1][1] = max(merged[-1][1], e)
        else:
            merged.append([s, e])
    body, budget = [], max_chars - len(head) - 200
    for s, e in merged:
        s = max(s, len(head))                      # head 와 겹치는 부분 제외
        if e <= s:
            continue
        seg = text[s:e]
        if len(seg) > budget:
            seg = seg[:budget]                     # 예산을 넘는 구간은 잘라서라도 넣는다(뒤 구간 전체 탈락 방지)
        if not seg:
            break
        body.append(f"...\n{seg}\n...")
        budget -= len(seg)
        if budget <= 0:
            break
    return head + "\n\n[핵심 구간 발췌]\n" + "\n".join(body)


def focus_tiered(text: str, max_chars: int, tiers: List[List[str]], window: int = 1200, head_ratio: float = 0.3, first_n: int = 2) -> str:
    """앞부분 + 키워드 주변 창. tiers 의 앞 묶음(찾는 항목에만 쓰이는 낱말)부터 예산을 쓰고, 남으면 뒤 묶음(흔한 낱말)에 쓴다.
    한 묶음 안에서는 각 키워드의 처음 first_n 곳을 먼저 담아, 자주 나오는 낱말 하나가 예산을 다 쓰지 않게 한다. 출력은 문서 순서."""
    if len(text) <= max_chars:
        return text
    head_n = int(max_chars * head_ratio)
    left = max_chars - head_n - 200
    taken: List[List[int]] = []                       # 겹치지 않는 [시작, 끝) 구간, 정렬 상태 유지

    def add(s: int, e: int) -> None:
        nonlocal left, taken
        s, cur, parts = max(s, head_n), max(s, head_n), []
        for a, b in taken:                            # 이미 담은 구간을 뺀 나머지만 예산에서 차감
            if b <= cur:
                continue
            if a >= e:
                break
            if a > cur:
                parts.append([cur, a])
            cur = max(cur, b)
        if cur < e:
            parts.append([cur, e])
        for a, b in parts:
            if left <= 0:
                break
            b = min(b, a + left)
            taken.append([a, b])
            left -= b - a
        taken.sort()
        merged: List[List[int]] = []
        for a, b in taken:
            if merged and a <= merged[-1][1]:
                merged[-1][1] = max(merged[-1][1], b)
            else:
                merged.append([a, b])
        taken = merged

    for keys in tiers:
        hits = {k: [m.start() for m in re.finditer(re.escape(k), text)] for k in keys}
        first = sorted((pos, k) for k, ps in hits.items() for pos in ps[:first_n])
        rest = sorted((pos, k) for k, ps in hits.items() for pos in ps[first_n:])
        for pos, k in first + rest:
            if left <= 0:
                break
            add(max(0, pos - window // 3), min(len(text), pos + len(k) + window))
    body = "\n".join(f"...\n{text[a:b]}\n..." for a, b in taken)
    return text[:head_n] + "\n\n[핵심 구간 발췌]\n" + body


def _parse_json(s: str) -> Dict:
    s = s.strip()
    s = re.sub(r"^```(?:json)?|```$", "", s, flags=re.M).strip()
    start, end = s.find("{"), s.rfind("}")
    if start < 0 or end < 0:
        raise ValueError("JSON 미검출")
    try:
        obj, _ = json.JSONDecoder().raw_decode(s[start:])   # 첫 '{' 에서 시작하는 완결 객체만(뒤 설명문 무시)
        return obj
    except json.JSONDecodeError:
        return json.loads(s[start: end + 1])


def _num(v) -> Optional[float]:
    """숫자/숫자문자열('19,100,000,000', '1.2e9')/None/NaN → float 또는 None."""
    if v is None or isinstance(v, bool):
        return None
    if isinstance(v, (int, float)):
        return None if (isinstance(v, float) and (math.isnan(v) or math.isinf(v))) else float(v)
    s = str(v).replace(",", "").replace("원", "").strip()
    if not s or s.lower() in ("nan", "none", "null"):
        return None
    try:
        return float(s)
    except ValueError:
        return None


_UNIT_TAIL = re.compile(r"(㎡|m2|m²|제곱미터|%|퍼센트|미터|세대|가구|대|면|호|m|원|정)$", re.I)
_MONEY_MULT = {"조": 1e12, "억": 1e8, "천만": 1e7, "백만": 1e6, "만": 1e4, "천": 1e3}


def _num_loose(v, money: bool = False) -> Optional[float]:
    """단위가 붙은 숫자 문자열도 읽는다: '12.35%', '약 125대', '24.5m', '1,817.93㎡', '5,600천원'(money=True → 5,600,000).
    '60% 이하'·'550평'·'5억6천만원' 처럼 한 숫자로 볼 수 없는 표기는 None(호출부가 원문을 따로 남긴다)."""
    x = _num(v)
    if x is not None or v is None or isinstance(v, (int, float, bool)):
        return x
    t = re.sub(r"^(약|총|금|계)", "", str(v).replace(",", "").replace(" ", "").strip())
    mult = 1.0
    if money:
        m = re.fullmatch(r"([0-9.]+)(조|억|천만|백만|만|천)?원?정?", t)
        if not m:
            return None
        t, mult = m.group(1), _MONEY_MULT.get(m.group(2) or "", 1.0)
    else:
        t = _UNIT_TAIL.sub("", _UNIT_TAIL.sub("", t))
    try:
        return float(t) * mult
    except ValueError:
        return None


def jsonable(v):
    """numpy/pandas 스칼라·NaN 을 JSON 직렬화 가능한 값으로."""
    if v is None:
        return None
    if hasattr(v, "item"):
        try:
            v = v.item()
        except (ValueError, TypeError):
            return str(v)
    if isinstance(v, float) and (math.isnan(v) or math.isinf(v)):
        return None
    if isinstance(v, (str, int, float, bool)):
        return v
    return str(v)


def normalize_doc(data: Dict) -> Dict:
    """LLM 출력 정규화: 숫자 필드 → 숫자, 근거문구 배열 → dict, 정수 필드 반올림."""
    out = dict(data)
    for k in _NUM_FIELDS:
        x = _num(out.get(k))
        if x is not None and k not in ("연면적_m2", "건축면적_m2", "대지면적_m2"):
            x = int(round(x))
        out[k] = x
    ev = out.get("근거문구")
    if isinstance(ev, list):
        out["근거문구"] = {str(e.get("항목", "")): str(e.get("원문", "")) for e in ev if isinstance(e, dict)}
    elif not isinstance(ev, dict):
        out["근거문구"] = {}
    return out


def estimate_tokens(chars: int) -> int:
    """한국어 문서 기준 대략 1.5자 ≈ 1토큰 (견적용)."""
    return int(chars / 1.5) + 600


def estimate_cost(texts: Dict[str, str], keys: List[str], cfg_llm: Dict, usd_krw: float = 1400.0,
                  max_chars: Optional[int] = None, out_per_doc: int = 2000) -> Dict:
    max_chars = int(max_chars or cfg_llm.get("max_input_chars", 60000))
    n, chars = 0, 0
    for k in keys:
        t = texts.get(k)
        if t:
            n += 1
            chars += min(len(t), max_chars)
    in_tok = sum(estimate_tokens(min(len(texts[k]), max_chars)) for k in keys if texts.get(k))
    out_tok = n * int(out_per_doc)            # 기본 2,000 = 실측(3,975건 평균 2,012 토큰, 사고 토큰 포함)
    model = cfg_llm.get("model", "claude-sonnet-5")
    pin, pout = PRICE_USD_PER_M.get(model, (5.0, 25.0))
    usd = in_tok / 1e6 * pin + out_tok / 1e6 * pout
    return {"docs": n, "chars": chars, "input_tokens": in_tok, "output_tokens": out_tok, "model": model,
            "usd": round(usd, 2), "krw": int(usd * usd_krw)}


def build_user_prompt(text: str, api_hint: Dict, cfg_llm: Dict) -> str:
    hint = {k: jsonable(v) for k, v in (api_hint or {}).items()}
    schema = json.dumps(FIELDS, ensure_ascii=False, indent=1)
    return (f"[공고 메타데이터(API)]\n{json.dumps(hint, ensure_ascii=False)}\n\n"
            f"[출력 스키마]\n{schema}\n\n"
            f"[공고문 텍스트]\n{focus_text(text, int(cfg_llm.get('max_input_chars', 60000)))}")


DEFAULT_MAX_TOKENS = 16000     # 사고(thinking) 토큰이 여기에 포함된다 — Sonnet 5 는 기본으로 사고하므로 8,000 이면 긴 공고에서 JSON 전에 잘린다


def build_request_params(text: str, api_hint: Dict, cfg_llm: Dict, structured: Optional[bool] = None,
                         max_tokens: Optional[int] = None) -> Dict:
    """messages.create / Batches 요청에 공통으로 쓰는 파라미터. max_tokens 를 주면 설정값 대신 사용(잘린 건 재제출용)."""
    if structured is None:
        structured = bool(cfg_llm.get("structured_output", True))
    params = dict(model=cfg_llm.get("model", "claude-sonnet-5"),
                  max_tokens=int(max_tokens or cfg_llm.get("max_tokens", DEFAULT_MAX_TOKENS)),
                  system=SYSTEM, messages=[{"role": "user", "content": build_user_prompt(text, api_hint, cfg_llm)}])
    if structured:
        params["output_config"] = {"format": {"type": "json_schema", "schema": output_schema()}}
    return params


def parse_response_text(out: str, model: str = "") -> Dict:
    data = normalize_doc(_parse_json(out))
    data["_model"] = model
    return data


_STRUCTURED_STATE = {"ok": True}      # 한 번 스키마가 거부되면(400) 이 프로세스에서는 더 시도하지 않는다


NONSTREAM_MAX_TOKENS = 21_333      # SDK 가 비스트리밍 요청에 허용하는 상한(10분 안에 끝날 수 있는 크기). 넘으면 스트리밍으로 받아야 한다


def _messages_call(client, **kw):
    """max_tokens 가 비스트리밍 상한을 넘으면(잘린 건 2배 재시도 등) 스트리밍으로 받아 최종 메시지를 돌려준다."""
    if int(kw.get("max_tokens", 0) or 0) > NONSTREAM_MAX_TOKENS:
        with client.messages.stream(**kw) as s:
            return s.get_final_message()
    return client.messages.create(**kw)


def _finish_call(create, kwargs: Dict):
    """호출 후 잘림(max_tokens)이면 1회 2배로 재시도, 거부(refusal)면 예외."""
    resp = create(**kwargs)
    if getattr(resp, "stop_reason", "") == "max_tokens":     # 사고 토큰까지 포함해 잘린 경우 1회 2배로 재시도
        kwargs["max_tokens"] = kwargs["max_tokens"] * 2
        resp = create(**kwargs)
        if getattr(resp, "stop_reason", "") == "max_tokens":
            raise ValueError("응답 해석 실패(stop=max_tokens): 두 번 잘림 — config.llm.max_tokens 를 늘리세요")   # 'stop=max_tokens' 가 있어야 다음 실행에서 2배 재요청
    if getattr(resp, "stop_reason", "") == "refusal":
        raise ValueError("모델이 응답을 거부함(refusal)")
    return resp


def _structured_create(client, schema: Dict):
    """구조화 출력 → 400 이면 일반 응답으로 폴백하는 create 함수를 만든다(프로세스 내 1회만 시도)."""
    import anthropic  # type: ignore
    state = {"structured": _STRUCTURED_STATE["ok"]}

    def _create(**kw):
        try:
            if state["structured"]:
                return _messages_call(client, output_config={"format": {"type": "json_schema", "schema": schema}}, **kw)
            return _messages_call(client, **kw)
        except anthropic.BadRequestError as e:
            if not state["structured"]:
                raise
            log.warning("구조화 출력 미지원/거부(%s) → 일반 응답 파싱으로 폴백(이 실행에서는 다시 시도하지 않음)", str(e)[:160])
            state["structured"] = False
            _STRUCTURED_STATE["ok"] = False
            return _messages_call(client, **kw)
    return _create


# ── 2차 추출: 내역서 집계표의 공종별 금액 ──────────────────────────────────
TRADE_LABELS = ["건축", "전기", "정보통신", "소방", "조경", "기계설비", "토목", "철거", "공통가설", "간접비", "관급자재", "부가세", "합계", "배분제외", "기타"]
_AGG_NAME_RE = re.compile(r"도급\s*(?:비|액|금액)|총\s*공사비|순\s*공사비|직접\s*공사비|공사비\s*계|총계|합계|총\s*원가|총\s*사업비|추정가격|기초금액|공사금액|총액")
# 항목명만으로 '계약 전체 금액' 이 분명한 말. 순공사비·직접공사비는 공종별 원가계산서의 소계일 수 있어 넣지 않는다
_TOTAL_NAME_RE = re.compile(r"추정가격|기초금액|총\s*공사비|총\s*사업비|총\s*원가|도급\s*(?:비|액|금액)|공사금액|총액")
_VAT_WORD = r"부\s*가\s*가\s*치\s*세|부\s*가\s*세|(?<![A-Za-z])VAT(?![A-Za-z])"
_TRADE_RUN = r"(?:(?:전기|소방|정보통신|통신|기계설비|기계|토목|조경|건축|철거)(?:\s*(?:시설\s*공사|시설|공사|설비))?\s*[·,/∙ㆍ・및\s]*)+"
# '부가세 제외'·'관급자재비 포함'·'관급자재를 제외한'·'(도급비 미포함)'·'(전기·소방 별도)'·'추정가격+부가가치세' 처럼 뒤따르는 조건·부정어는 라벨 판별에서 지운다
_NEG_NAME_RE = re.compile(r"(?:" + _VAT_WORD + r"|관급\s*자재|관급|지급\s*자재|도급비|도급\s*금액|간접비|이윤|관리비)(?:비|대|액|분|금액)?\s*(?:을|를|은|는)?\s*(?:제외한|포함한|포함된|제외|별도|미포함|불포함|포함)"
                          r"|" + _TRADE_RUN + r"(?:제외한|제외|별도|미포함|불포함|분리\s*발주)"
                          r"|[+＋및]\s*(?:" + _VAT_WORD + r")", re.I)
_NONTRADE_RE = re.compile(r"건축물|건설\s*기계|중장비|폐기물\s*처리\s*시설|기계\s*경비")     # 공종어를 품었지만 공종이 아닌 말('건축물 철거', '건설기계 운반비')
_SOBANG_LICENSE_RE = re.compile(r"(?:일반|전문)?소방시설공사업(?:\((?:전기|기계)\))?")      # '일반소방시설공사업(전기)' = 소방 면허의 전기 분야 → 소방 하나
_LICENSE_NAMES = (("토목건축공사업", " "),)        # 면허명 '토목건축공사업' 은 토목+건축 두 공종이 아니므로 지운다('건축공사업 또는 토목건축공사업' → 건축, '토목공사업 또는 토목건축공사업' → 토목)
_GOV_NAME_RE = re.compile(r"관급|지급\s*자재|도급자\s*(설치|시공)")
_GOV_LABOR_RE = re.compile(r"설치\s*(비|공사|노무)|시공비|노무비")     # '관급자재 설치비' 는 관급 금액이 아니라 설치 노무(도급 공사비)
_VAT_NAME_RE = re.compile(_VAT_WORD, re.I)
_TEMP_NAME_RE = re.compile(r"가설")
_DEMOLISH_RE = re.compile(r"철거|해체")
_INDIRECT_NAME_RE = re.compile(r"간접|일반관리|이윤|경비(?!실|초소|원|동|업)|보험|공제|안전관리|품질관리|환경보전|관리비|조달\s*수수료|보증|폐기물|임금채권|석면")
_EXCL_NAME_RE = re.compile(r"이설비|이전비|분담금|부담금|용역비|감리\s*(비|용역|대가|료)|설계비|측량|미술\s*장식|인증\s*수수료|보상비|예비비")        # 도급 공사비 밖 → '배분제외'
_OUTSIDE_RE = re.compile(r"도급(?:비|액|금액)?\s*(?:미포함|불포함|별도|제외|외)")      # 항목명이 스스로 '도급비 미포함' 이라 하면 모델 라벨과 무관하게 배분제외
_PAREN_RE = re.compile(r"\([^()]*\)|\[[^\[\]]*\]")
TRADE_ALLOC_LABELS = ["건축", "전기", "정보통신", "소방", "조경", "기계설비", "토목", "철거", "기타"]   # 비율 배분에 쓰는 라벨(공통가설·간접비·관급자재·부가세·합계·배분제외는 제외)
TRADE_SYSTEM = (
    "당신은 한국 공공 건축공사 입찰 내역서(설계내역서·원가계산서·총괄집계표)와 공고문의 공사금액 표를 읽어 공종별 금액 구성을 추출하는 분석가입니다. "
    "반드시 JSON 객체 하나만 출력합니다. 입력은 집계표에서 추려 낸 행들입니다. "
    "'공종별금액'에는 대공종 행(건축·기계설비·토목·조경·전기·정보통신·소방·철거 등)과 공통가설·간접비(간접노무비·산재·고용보험·안전관리비·일반관리비·이윤 등 합쳐서 하나)·관급자재·부가세 행을 넣고, "
    "세부 공종(철근콘크리트·미장·창호 등)은 넣지 않습니다. 금액은 원 단위 정수 자릿수만 적고 '천원' 단위 표는 ×1000 합니다(단위 표기를 '단위'에 적음). "
    "'합계_원'에는 집계표의 총합계(총공사비·총계) 값을, '금액기준'에는 그 합계가 부가세 포함인지·제외(순공사비)인지·직접공사비인지 적고 모르면 '불명'. "
    "집계표에 합계만 있고 공종별 금액이 없으면 '공종별금액'은 빈 배열로 둡니다. 추정하거나 계산해 만들지 않습니다. "
    "여러 공종을 합친 금액(예: '도급비(건축, 토목, 기계, 조경)', 총공사비, 순공사비, 총계)은 한 공종의 금액이 아니므로 공종을 '합계'로 적습니다. "
    "도급 공사비 밖의 항목(이설비·이전비, 각종 분담금, 설계·감리 용역비, 미술장식품, 인증수수료 등 '도급비 미포함'·별도 비용)은 '배분제외'로 적습니다. "
    "발췌는 셀 위치를 탭으로 보존한 것입니다: 첫 열에 있는 행은 대공종(건축공사·기계설비공사 등), 들여쓴 행은 그 대공종의 세부 공종이며 '소계'는 바로 위 대공종의 합입니다. "
    "대공종 소계 안에 이미 든 가설공사·안전관리비 등은 '공통가설'로 따로 적지 말고, 집계표 최상위에 별도 행으로 있는 공통가설(가설전기·가설용수 포함)만 '공통가설'로 적습니다. "
    "각 값의 근거 행을 '근거문구'에 짧게 인용합니다."
)
_TRADE_KEYS = ["공사금액", "공사 금액", "공사비 내역", "공종별", "집계", "총괄", "원가계산", "공사개요", "구분"]
TRADE_FIELDS = {"금액기준": "부가세포함|부가세제외|직접공사비|불명", "관급자재_포함여부": "포함|제외|불명", "단위": "원|천원|불명",
                "합계_원": "integer|null", "공종별금액": "[{공종(" + "|".join(TRADE_LABELS) + " 중 하나), 항목명_원문, 금액_원}]",
                "출처": "string", "신뢰도": "high|medium|low", "근거문구": "string"}


def trade_output_schema() -> Dict:
    """2차 추출 구조화 출력 스키마. 모두 string/enum·required(union 0·optional 0) — 1차와 같은 제약."""
    item = {"type": "object", "additionalProperties": False,
            "properties": {"공종": {"type": "string", "enum": TRADE_LABELS}, "항목명_원문": {"type": "string"}, "금액_원": {"type": "string"}},
            "required": ["공종", "항목명_원문", "금액_원"]}
    props = {
        "금액기준": {"type": "string", "enum": ["부가세포함", "부가세제외", "직접공사비", "불명"]},
        "관급자재_포함여부": {"type": "string", "enum": ["포함", "제외", "불명"]},
        "단위": {"type": "string", "enum": ["원", "천원", "불명"]},
        "합계_원": {"type": "string"},
        "공종별금액": {"type": "array", "items": item},
        "출처": {"type": "string"},
        "신뢰도": {"type": "string", "enum": ["high", "medium", "low"]},
        "근거문구": {"type": "string"},
    }
    return {"type": "object", "properties": props, "required": list(props.keys()), "additionalProperties": False}


def keyword_windows(text: str, keys: List[str], budget: int, window: int = 1200) -> str:
    """키워드 주변 창만 모아 budget 자 안으로(앞부분 없이). 2차 추출에서 공고문·현장설명서의 공사금액 표를 끌어오는 데 쓴다."""
    if not text or budget <= 0:
        return ""
    spans = []
    for k in keys:
        for m in re.finditer(re.escape(k), text):
            spans.append((max(0, m.start() - window // 3), min(len(text), m.end() + window)))
    spans.sort()
    merged: List[List[int]] = []
    for s, e in spans:
        if merged and s <= merged[-1][1]:
            merged[-1][1] = max(merged[-1][1], e)
        else:
            merged.append([s, e])
    body, left = [], budget
    for s, e in merged:
        seg = text[s:e][:left]
        if not seg:
            break
        body.append(f"...\n{seg}\n...")
        left -= len(seg)
        if left <= 0:
            break
    return "\n".join(body)


def build_trade_prompt(snippet: str, api_hint: Dict) -> str:
    hint = {k: jsonable(v) for k, v in (api_hint or {}).items()}
    return (f"[공고 메타데이터(API)]\n{json.dumps(hint, ensure_ascii=False)}\n\n"
            f"[출력 스키마]\n{json.dumps(TRADE_FIELDS, ensure_ascii=False, indent=1)}\n\n"
            f"[집계표·공사금액 표 발췌]\n{snippet}")


def build_trade_request_params(snippet: str, api_hint: Dict, cfg_llm: Dict, structured: Optional[bool] = None,
                               max_tokens: Optional[int] = None) -> Dict:
    if structured is None:
        structured = bool(cfg_llm.get("structured_output", True))
    params = dict(model=cfg_llm.get("model", "claude-sonnet-5"),
                  max_tokens=int(max_tokens or cfg_llm.get("max_tokens", DEFAULT_MAX_TOKENS)),
                  system=TRADE_SYSTEM, messages=[{"role": "user", "content": build_trade_prompt(snippet, api_hint)}])
    if structured:
        params["output_config"] = {"format": {"type": "json_schema", "schema": trade_output_schema()}}
    return params


def _name_label(core: str) -> Tuple[Optional[str], Optional[str]]:
    """항목명(괄호 안을 뺀 핵심부)에서 분명한 비공종 라벨을 찾는다. hard(가설·관급·부가세)는 모델 라벨과 무관하게, soft(간접비·배분제외)는 모델이
    기타·합계로 적었거나 공종이 여럿 섞인 행만. '가설건물 철거' 는 철거, '관급자재 설치비' 는 설치 노무이므로 hard 에서 뺀다."""
    if _TEMP_NAME_RE.search(core) and not _DEMOLISH_RE.search(core):
        return "공통가설", None
    if _GOV_NAME_RE.search(core) and not _GOV_LABOR_RE.search(core):
        return "관급자재", None
    if _VAT_NAME_RE.search(core):
        return "부가세", None
    if _INDIRECT_NAME_RE.search(core):          # '…석면분담금+임금채권부담금+일반관리비+이윤' 같은 간접비 나열이 분담금 때문에 배분제외로 가지 않도록 간접비 먼저
        return None, "간접비"
    if _EXCL_NAME_RE.search(core):
        return None, "배분제외"
    return None, None


def _name_parts(name: str) -> Tuple[str, str, List[str], List[str]]:
    """항목명 → (조건어·비공종어를 지운 이름, 괄호 밖 핵심부, 핵심부의 배분 공종 라벨, 이름 전체의 배분 공종 라벨).
    '건 축 공 사 도급금액'·'▶ 전 기 공 사' 처럼 띄어 쓴 머리글이 흔해 공백을 모두 지운 뒤 본다."""
    from . import classify
    n = re.sub(r"\s+", "", str(name or ""))
    n = _SOBANG_LICENSE_RE.sub("소방공사업", n)
    n = _NONTRADE_RE.sub(" ", n)
    n = _NEG_NAME_RE.sub(" ", n)
    for a, b in _LICENSE_NAMES:
        n = n.replace(a, b)
    core = _PAREN_RE.sub(" ", n)
    if not re.search(r"[가-힣A-Za-z]", core):
        core = n
    labels = lambda x: [l for l in classify._doc_labels(x)[0] if l in TRADE_ALLOC_LABELS]      # 가스·승강기·창호 같은 비공종 라벨은 세지 않는다
    return n, core, labels(core), labels(n)


def _is_total_name(name: str) -> bool:
    """공종어 없이 추정가격·기초금액·총공사비·도급비 같은 '계약 전체 금액' 이름인가(공고문 공사금액 표의 열 머리글이 항목명으로 오는 경우 포함)."""
    n, _, _, trades_all = _name_parts(name)
    return not trades_all and bool(_TOTAL_NAME_RE.search(n))


def _trade_label(raw: str, item_name: str = "") -> str:
    """모델이 적은 공종명을 라벨로. 구조화 출력이 거부돼 일반 응답으로 폴백한 경우 '건축공사비'·'기계설비(공조)' 같은 원문이 올 수 있다.
    항목명 규칙(우선순위): ① 공백을 지우고('건 축 공 사'), '부가세 제외'·'관급자재비 포함'·'(도급비 미포함)'·'(전기·소방 별도)'·'추정가격+부가가치세' 같은 조건어와
    '건축물'·'건설기계' 같은 비공종어는 지우며, 면허명 '토목건축공사업'·'일반소방시설공사업(전기)' 은 공종으로 세지 않는다
    ② 가설·관급·부가세 가 항목명 핵심부(괄호 밖)에 있으면 그 라벨(가설건물 철거·관급자재 설치비는 예외) ③ 항목명이 스스로 '도급비 미포함·별도' 라 하면 '배분제외'
    ④ 보험·공제·안전관리비·폐기물 등은 '간접비', 이설비·분담금·용역비 등은 '배분제외'(모델이 기타·합계로 적었거나 공종이 여럿 섞인 행만; 철거 행의 폐기물은 철거)
    ⑤ 괄호 밖에 공종이 2개 이상이거나('건축, 기계공사', '전기소방공사'), 괄호 밖에 공종이 없고 괄호 안에 2개 이상이면('도급비(건축, 토목, 기계, 조경)') '합계'(비율 배분 제외).
    '건축(건축+무대기계+기계) 소계' 처럼 괄호 밖 공종이 하나면 그 공종(그 공종의 묶음 금액) ⑥ 모델이 '기타'로 적었는데 항목명에 공종이 하나면 그 공종('건축공사 합계' → 건축),
    '합계' 로 적었어도 괄호 밖에 합산어 없이 공종이 하나면 그 공종('일반소방시설공사업(전기)' → 소방) ⑦ 기타·합계인데 도급비·총공사비·추정가격 등 합산어면 '합계'.
    모델이 공종을 적은 '추정가격'·'도급비' 행은 여기서 바꾸지 않고 normalize_trade_doc 이 문서 단위로 판단한다(다른 공종 행이 함께 있을 때만 '합계')."""
    from . import classify
    lab = raw.strip()
    name = str(item_name or "")
    if name:
        n, core, tc, ta = _name_parts(name)
        trades = tc or ta
        multi = len(set(tc)) >= 2 or (not tc and len(set(ta)) >= 2)
        hard, soft = _name_label(core)
        if not hard and not soft and lab in ("기타", ""):
            soft = _name_label(n)[1]                 # '기타금액(분담금 등)' 처럼 괄호 안에만 단서가 있는 기타 행
        if hard:
            return hard
        if _OUTSIDE_RE.search(name):
            return "배분제외"
        if soft and (multi or lab in ("기타", "합계", "")):
            if soft == "간접비" and set(trades) == {"철거"}:
                return "철거"
            return soft
        if multi:
            return "합계"
        if lab in ("기타", "") and len(set(trades)) == 1:
            return trades[0]
        if lab == "합계" and len(set(tc)) == 1 and not _TOTAL_NAME_RE.search(core):     # '합계 - 건축(건축+무대기계) 소계' 는 건축의 묶음 소계(합계·소계는 막지 않음)
            return tc[0]
        if lab in ("기타", "합계", "") and _AGG_NAME_RE.search(n):
            return "합계"
    if lab in TRADE_LABELS:
        return lab
    if "가설" in lab:
        return "공통가설"
    if "부가" in lab:
        return "부가세"
    if "관급" in lab:
        return "관급자재"
    if "배분" in lab or lab == "제외":
        return "배분제외"
    if any(w in lab for w in ("간접", "일반관리", "이윤", "경비", "보험")):
        return "간접비"
    labels = [l for l in classify._doc_labels(lab)[0] if l in TRADE_LABELS]
    if len(labels) == 1:                       # '전기공사금액' 같은 자유 서술 라벨은 공종 하나면 그 공종(합산어 검사보다 먼저)
        return labels[0]
    if _AGG_NAME_RE.search(lab):
        return "합계"
    return "기타"


def normalize_trade_doc(data: Dict) -> Dict:
    """2차 추출 출력 정규화: 금액 → 정수(원), 빈 항목 제거, 공종 라벨 밖 값은 '기타', 비율 배분용 공종 합(_공종합_원) 계산."""
    out = dict(data)
    total = _num(out.get("합계_원"))
    out["합계_원"] = int(round(total)) if total is not None else None
    items = []
    for it in out.get("공종별금액") or []:
        if not isinstance(it, dict):
            continue
        amt = _num(it.get("금액_원"))
        if amt is None:
            continue
        lab = _trade_label(str(it.get("공종") or ""), str(it.get("항목명_원문") or ""))
        items.append({"공종": lab, "항목명_원문": str(it.get("항목명_원문") or "")[:80], "금액_원": int(round(amt))})
    # 모델이 공종을 적은 '추정가격'·'도급비'·'총공사비' 행: 같은 문서에 이름이 공종인 다른 배분 행(기타 제외)이 있으면 전체 금액을 공종 행처럼 중복 셈한 것이므로 '합계'.
    # 그 행만 있거나(단일 공종 공고의 추정가격) 모두 열 머리글 이름이면(공사금액 표의 공종별 추정가격) 모델 라벨을 믿는다.
    total_idx = [k for k, i in enumerate(items) if i["공종"] in TRADE_ALLOC_LABELS and i["공종"] != "기타" and _is_total_name(i["항목명_원문"])]
    if total_idx and any(i["공종"] in TRADE_ALLOC_LABELS and i["공종"] != "기타" and k not in total_idx for k, i in enumerate(items)):
        for k in total_idx:
            items[k]["공종"] = "합계"
    out["공종별금액"] = items
    out["_공종합_원"] = sum(i["금액_원"] for i in items if i["공종"] in TRADE_ALLOC_LABELS) or None
    if out["합계_원"] is None:
        aggs = [i["금액_원"] for i in items if i["공종"] == "합계" and not re.search(r"직접\s*공사비|순\s*공사비|소계", i["항목명_원문"])]
        if aggs:
            out["합계_원"] = max(aggs)                 # 합계 행만 있고 합계_원이 비었으면 가장 큰 합계 행을 합계로(직접공사비·순공사비 소계는 총액이 아니므로 제외)
    for k in ("금액기준", "관급자재_포함여부", "단위", "출처", "신뢰도", "근거문구"):
        out[k] = str(out.get(k) or "")
    return out


def parse_trade_response_text(out: str, model: str = "") -> Dict:
    data = normalize_trade_doc(_parse_json(out))
    data["_model"] = model
    return data


def extract_trades_with_claude(snippet: str, api_hint: Dict, cfg_llm: Dict, max_tokens: Optional[int] = None) -> Dict:
    """집계표 발췌 → 공종별 금액 dict(즉시 호출)."""
    import anthropic  # type: ignore
    client = anthropic.Anthropic(api_key=os.environ.get(cfg_llm.get("api_key_env", "ANTHROPIC_API_KEY")) or None)
    kwargs = dict(model=cfg_llm.get("model", "claude-sonnet-5"),
                  max_tokens=int(max_tokens or cfg_llm.get("max_tokens", DEFAULT_MAX_TOKENS)),
                  system=TRADE_SYSTEM, messages=[{"role": "user", "content": build_trade_prompt(snippet, api_hint)}])
    create = _structured_create(client, trade_output_schema()) if bool(cfg_llm.get("structured_output", True)) else (lambda **kw: _messages_call(client, **kw))
    resp = _finish_call(create, kwargs)
    out = "".join(getattr(b, "text", "") for b in resp.content if getattr(b, "type", "") == "text")
    data = parse_trade_response_text(out, kwargs["model"])
    data["_usage_in"] = getattr(getattr(resp, "usage", None), "input_tokens", None)
    data["_usage_out"] = getattr(getattr(resp, "usage", None), "output_tokens", None)
    return data


def cross_verify_trades(bid_no: str, api_row: Dict, tdoc: Dict, vat: float = 0.10) -> List[Dict]:
    """2차 추출 정합성: 공종 합 vs 집계표 합계, 집계표 합계의 금액 기준(부가세 포함/제외/직접공사비/천원 단위) 판별."""
    logs: List[Dict] = []
    total, tsum = _num(tdoc.get("합계_원")), _num(tdoc.get("_공종합_원"))
    items = tdoc.get("공종별금액") or []
    if not items:
        logs.append({"공고번호": bid_no, "항목": "공종별 금액(문서)", "판정": "참고", "비고": "집계표에 공종별 금액 없음(합계만 또는 미발견)"})
        return logs
    gross = tsum
    if gross is not None:
        gross += sum(i["금액_원"] for i in items if i["공종"] in ("공통가설", "간접비", "관급자재", "부가세"))
        excl = sum(i["금액_원"] for i in items if i["공종"] == "배분제외")
        if total and excl:                              # 이설비·분담금 같은 배분제외 행은 집계표 총합에 들어 있을 수도, 빠져 있을 수도 → 합계에 가까운 쪽
            gross = min((gross, gross + excl), key=lambda v: abs(v - total))
    if total and gross:
        r = (gross - total) / total
        logs.append({"공고번호": bid_no, "항목": "공종합=집계표합계", "API값": total, "문서값": gross, "차이": gross - total, "차이율": round(r, 5),
                     "판정": "정상" if abs(r) <= 0.01 else ("경고" if abs(r) <= 0.05 else "경고"),
                     "비고": "" if abs(r) <= 0.01 else "공종 행 합과 합계 행 불일치(간접비·관급·부가세 행 누락 또는 세부 공종 중복 의심)"})
    alloc = [i for i in items if i["공종"] in TRADE_ALLOC_LABELS]
    api_trade = str(api_row.get("공종") or "")
    if alloc and api_trade and all(_is_total_name(i["항목명_원문"]) for i in alloc) and {i["공종"] for i in alloc} - {"기타", api_trade}:
        labs = "·".join(sorted({i["공종"] for i in alloc}))
        logs.append({"공고번호": bid_no, "항목": "공종 라벨 의심(문서)", "문서값": sum(i["금액_원"] for i in alloc), "판정": "참고",
                     "비고": f"전체 금액 행('{alloc[0]['항목명_원문'][:20]}')을 {labs} 로 읽음 — 공고 공종({api_trade})과 다름(공종 오분류 또는 모델 라벨 오류) → 04b 확인"})
    etc = sum(i["금액_원"] for i in items if i["공종"] == "기타")
    if tsum and etc / tsum > 0.5:
        logs.append({"공고번호": bid_no, "항목": "공종 라벨 미매핑(문서)", "문서값": etc, "판정": "경고",
                     "비고": "'기타' 금액이 공종 합의 50% 초과 → 공종명을 라벨로 못 바꾼 것(비구조화 응답) 의심, 04b 세부공종 확인"})
    base = total or gross
    api_p, api_b = _num(api_row.get("추정가격")), _num(api_row.get("기초금액"))
    gov_api = (_num(api_row.get("도급자관급액_API")) or 0) + (_num(api_row.get("관급자관급액_API")) or 0) or (_num(api_row.get("관급자재_API")) or 0)
    gov_doc = sum(i["금액_원"] for i in items if i["공종"] == "관급자재")
    vat_doc = sum(i["금액_원"] for i in items if i["공종"] == "부가세")
    if base and (api_p or api_b):
        def near(a, b, tol=0.01):
            return bool(a and b and abs(a - b) / b <= tol)
        # 문서 합계는 관급·부가세를 포함했을 수도, 뺐을 수도 있다 → 보정한 값들과 API 값들을 짝지어 하나라도 맞으면 그 기준을 적는다
        doc_cands = [(base, "집계표 합계"), (base - gov_doc, "집계표 합계 − 관급자재"), (base - vat_doc, "집계표 합계 − 부가세"),
                     (base - gov_doc - vat_doc, "집계표 합계 − 관급 − 부가세")]
        api_cands = []
        for av, an in ((api_b, "기초금액(부가세 포함)"), (api_p, "추정가격(부가세 제외)")):
            if av:                                      # 값이 없는 API 항목은 후보에서 뺀다(없는 기초금액 + 관급 = 관급만 남는 가짜 후보 방지)
                api_cands.append((av, an))
                if gov_api:
                    api_cands.append((av + gov_api, an.split("(")[0] + " + 관급자재(API)"))
        match = next(((dn, an) for dv, dn in doc_cands for av, an in api_cands if dv and av and near(dv, av)), None)
        close = match or next(((dn, an) for dv, dn in doc_cands for av, an in api_cands if dv and av and near(dv, av, 0.03)), None)
        if match:
            verdict, note = "참고", f"{match[0]} = {match[1]}"
        elif near(base * 1000, api_b) or near(base * 1000, api_p):
            verdict, note = "경고", "천원 단위 표를 원으로 읽은 의심(×1000 하면 API 금액과 일치)"
        elif close:
            verdict, note = "참고", f"{close[0]} ≈ {close[1]} (3% 이내, 조달수수료·관급 금액 차이 가능)"
        elif (api_p or api_b) and 0.6 <= base / (api_p or api_b / (1 + vat)) <= 0.95:      # 추정가격이 없으면 기초금액/(1+부가세)로 대신
            verdict, note = "참고", "집계표 합계가 추정가격의 60~95% → 직접공사비(간접비 제외) 가능성"
        else:
            verdict, note = "경고", "집계표 합계가 API 추정가격·기초금액(관급 보정 포함)과 불일치(차수·변경 내역서 혼입 또는 다른 공사 집계표 의심)"
        logs.append({"공고번호": bid_no, "항목": "집계표합계 기준", "API값": api_b or api_p, "문서값": base,
                     "차이": base - (api_b or api_p), "차이율": round((base - (api_b or api_p)) / (api_b or api_p), 5), "판정": verdict, "비고": note})
    return logs


def extract_with_claude(text: str, api_hint: Dict, cfg_llm: Dict, max_tokens: Optional[int] = None) -> Dict:
    """공고문 텍스트 → 구조화 dict. api_hint: {'공고번호','공고명','수요기관'} (문서 식별용 힌트. 금액은 넣지 않는다).
    max_tokens: 설정값 대신 쓸 출력 한도(이전에 잘린 건은 2배로 시작)."""
    import anthropic  # type: ignore

    client = anthropic.Anthropic(api_key=os.environ.get(cfg_llm.get("api_key_env", "ANTHROPIC_API_KEY")) or None)
    user = build_user_prompt(text, api_hint, cfg_llm)
    kwargs = dict(model=cfg_llm.get("model", "claude-sonnet-5"),
                  max_tokens=int(max_tokens or cfg_llm.get("max_tokens", DEFAULT_MAX_TOKENS)),
                  system=SYSTEM, messages=[{"role": "user", "content": user}])
    structured = bool(cfg_llm.get("structured_output", True)) and _STRUCTURED_STATE["ok"]

    def _create(**kw):
        nonlocal structured
        try:
            if structured:
                return _messages_call(client, output_config={"format": {"type": "json_schema", "schema": output_schema()}}, **kw)
            return _messages_call(client, **kw)
        except anthropic.BadRequestError as e:
            if not structured:
                raise
            log.warning("구조화 출력 미지원/거부(%s) → 일반 응답 파싱으로 폴백(이 실행에서는 다시 시도하지 않음)", str(e)[:160])
            structured = False
            _STRUCTURED_STATE["ok"] = False
            return _messages_call(client, **kw)

    resp = _finish_call(_create, kwargs)
    out = "".join(getattr(b, "text", "") for b in resp.content if getattr(b, "type", "") == "text")
    data = normalize_doc(_parse_json(out))
    data["_model"] = cfg_llm.get("model")
    data["_usage_in"] = getattr(resp.usage, "input_tokens", None)
    data["_usage_out"] = getattr(resp.usage, "output_tokens", None)
    return data


def _doc_base_in_04(p, b, bd, v=None, vd=None) -> str:
    """04 시트의 부가세(수식)·금액합계 가운데 '문서' 기초금액을 따르는 부분(build_excel 의 04 수식과 같은 조건·같은 순서). 따르는 것이 없으면 ''.
    p = 04 추정가격(수식)에 쓰이는 값, b·bd = API·문서 기초금액, v·vd = API·문서 부가세.
    ① API 기초금액이 없고 문서 기초금액이 추정가격의 100~110.5% → 기초금액 기준 합계 = 문서 기초금액
    ② 부가세_출처가 '기초금액(문서)-추정가격'(API·문서 부가세를 못 쓰고 API 기초금액이 범위 밖) → 부가세·추정가격 기준 합계 = 문서 기초금액
    반환: 'both' = ①+② 부가세(수식)와 두 합계 모두 / 'base' = ①만: 기초금액 기준 합계만 / 'vat' = ②만: 부가세(수식)와 추정가격 기준 합계(기초금액 기준 합계는 API 기초금액).
    추정가격이 API·문서 어디에도 없는 행(합계 = 문서 기초금액)은 다루지 않는다 — 그런 행에는 _gap_hint 가 비고를 쓰지 않는다."""
    if not p or p <= 0 or bd is None:
        return ""
    okd = lambda x: x is not None and 0 <= x - p <= p * 0.105
    bb = b if b is not None else bd
    okv = lambda x: x is not None and 0 <= x <= p * 0.105 and not (x < p * 0.099 and bb is not None and abs(bb - p * 1.1) <= p * 0.011)
    vat_doc = not okv(v) and not okv(vd) and not okd(b) and okd(bd)       # 부가세_출처 = '기초금액(문서)-추정가격'
    if not (b is not None and b > 0) and okd(bd):
        return "both" if vat_doc else "base"
    return "vat" if vat_doc else ""


def _ratio_txt(x: float) -> str:
    """비고에 적는 비율: 1,000 이상은 쉼표, 0.001 미만은 1/n 로 적는다('.3g' 는 1,000 부터 1e+03 처럼 지수 표기가 된다)."""
    if x >= 999.5:
        return f"{x:,.0f}"
    if 0 < x < 0.001:
        return f"1/{1 / x:,.0f}"
    return f"{x:.3g}"


def _gap_hint(a: float, d: float, vat: float = 0.10, item: str = "", sib: Optional[Dict] = None) -> str:
    """API값과 문서값이 5% 넘게 다를 때, 비율과 같은 공고의 다른 금액으로 알 수 있는 흔한 원인을 비고에 적는다(판정은 바꾸지 않는다 — 검수자가 볼 순서를 정하는 단서).
    04 시트의 추정가격·기초금액 칸은 API값을 먼저 쓰므로 문서값이 그 칸을 바꾸지는 않는다. 그래서 API값 쪽이 어긋난 경우(문서값이 같은 공고의 다른 금액과 맞는 경우)는
    '문서 오독' 으로 적지 않고 04 를 확인하라고 적고, 04 의 부가세·합계가 문서 기초금액을 따르는 경우(_doc_base_in_04)도 그렇게 적는다.
    item = '추정가격'/'기초금액', sib = 같은 공고의 다른 금액 {api_p, api_b, api_v, doc_p, doc_b, doc_v}(없으면 비율만 본다)."""
    if not a or not d or a < 0 or d < 0:
        return ""
    sib = sib or {}
    api_p, api_b, api_v, doc_p, doc_b, doc_v = (sib.get(x) for x in ("api_p", "api_b", "api_v", "doc_p", "doc_b", "doc_v"))
    q, k = d / a, 1 + vat
    near = lambda x: abs(q - x) <= x * 0.002
    same = lambda x, y: bool(x) and bool(y) and x > 0 and y > 0 and abs(x / y - 1) <= 0.002
    use = "04 는 API값 사용"
    p04 = api_p if (api_p and api_p > 0) else doc_p              # 04 추정가격(수식): API값(0 보다 클 때) → 문서값
    how = _doc_base_in_04(*((a, api_b, doc_b) if item == "추정가격" else (p04, a, d) if item == "기초금액" else (None, None, None)), api_v, doc_v)
    if how:                                                      # 무엇이 문서 기초금액을 따르는지 그대로 적는다(없는 API 기초금액을 'API값' 이라 적지 않는다)
        use = {"both": "API 기초금액이 없어 04 부가세(수식)·금액합계(두 기준 모두)가 문서 기초금액을 따름 → 04 합계 확인",
               "base": "API 기초금액이 없어 04 금액합계_기초금액기준이 문서 기초금액을 따름 → 04 합계 확인",
               "vat": "04 금액합계_기초금액기준은 API값이지만 부가세(수식)·금액합계_추정가격기준은 문서 기초금액을 따름 → 04 합계 확인"}[how]
    tail = f" ({use})" if how else ""                            # '04 는 API값 사용' 을 적지 않는 문구에도, 문서 기초금액을 따르는 것이 있으면 그것은 적는다
    more = lambda x: f"{use}, {x}" if how else f"{use} → {x}"    # use 가 이미 '→ 04 합계 확인' 으로 끝나면 화살표를 겹쳐 쓰지 않는다
    # API 추정가격이 같은 공고의 기초금액과 맞는가 — '기초금액≈추정가격×1.1' 행과 같은 값을 본다(기초금액: API → 문서, 부가세: API → 문서).
    # 맞으면 'API 추정가격이 어긋남' 이라고 적지 않는다(그 행이 정상·'공고에 적힌 부가세와 일치' 인데 추정가격 행이 API 를 의심하면 서로 엇갈린다)
    v_any = api_v if api_v is not None else doc_v
    has_api_b = bool(api_b) and api_b > 0
    b_any, bsrc = (api_b, "API") if has_api_b else (doc_b, "문서")
    api_fit = ""
    if item == "추정가격" and same(b_any, a * k):
        api_fit = f"API 추정가격×{k:g} = {bsrc} 기초금액" + ("(API 쪽은 서로 맞음)" if has_api_b else "(API 추정가격이 기초금액과 맞음)")
    exempt = item == "추정가격" and not api_fit and v_any is not None and v_any >= 0 and same(b_any, a + v_any)
    if exempt and same(b_any, d * k):
        # 같은 기초금액이 API 추정가격+부가세(면세·부분면세)로도, 문서 추정가격×1.1(과세)로도 나온다(실자료 20230734989) → 어느 쪽도 '다른 금액' 이 아니다
        return (f"문서값×{k:g} = {bsrc} 기초금액이고 API 추정가격 + 부가세 = {bsrc} 기초금액 — 두 쪽 모두 같은 기초금액과 맞음(면세·부분면세면 API 추정가격이, 과세면 문서 추정가격이 맞는 금액): "
                f"과세 여부를 공고문에서 확인({use})")
    if item == "추정가격" and same(api_b, d * k) and not api_fit and not exempt:
        # 문서 추정가격×1.1 = API 기초금액인데 API 추정가격은 그렇지 않다 → 어긋난 쪽은 API(부가세액·부가세 포함 금액을 추정가격 칸에 넣은 공고).
        # API 추정가격+부가세 = API 기초금액이면(면세·부분면세로 API 쪽이 서로 맞음) 여기에 넣지 않는다(위)
        return (f"문서값×{k:g} = API 기초금액 — 문서 추정가격은 기초금액과 맞고 API 추정가격은 맞지 않음(API 쪽이 부가세액·부가세 포함 금액 등 다른 값일 수 있음). "
                f"04 추정가격(수식)은 API값을 쓰므로 04 의 추정가격·부가세·과세구분과 추정가격 기준 합계 확인{tail}")
    if item == "추정가격" and not has_api_b and same(a, doc_v) and same(doc_v, d * vat) and same(doc_b, d * k):
        # API 기초금액이 없어 위 규칙을 쓸 수 없지만, 문서는 추정가격×1.1 = 기초금액으로 서로 맞고 API 추정가격이 문서 부가세(= 문서 추정가격의 10%)와 같다 → 어긋난 쪽은 API
        return (f"API 추정가격 = 문서 부가세 — 문서는 추정가격×{k:g} = 기초금액으로 서로 맞음: API 추정가격 칸에 부가세액이 들어간 것으로 보임. "
                f"04 추정가격(수식)은 API값을 쓰므로 04 합계 확인{tail}")
    if item == "기초금액" and same(d, (api_p or 0) * k):
        if same(a, api_p):               # API 기초금액 = API 추정가격 → 04 는 부가세 0(면세)으로 계산한다
            return f"문서값 = 추정가격×{k:g}, API 기초금액 = API 추정가격 — 면세 공사이거나 API 기초금액에 부가세가 빠진 것으로 보임({more('04 과세구분이 맞는지 확인')})"
        if not 0.75 <= a / d <= 1.35:    # 조사금액·관급 차이로 볼 수 없는 크기(실자료 34행은 0.83~1.19) → API 기초금액 쪽을 확인하게 한다
            return (f"문서값 = 추정가격×{k:g}(추정가격+부가세) — API 기초금액이 이와 크게 다름(API값÷문서값 = {_ratio_txt(a / d)}): API 기초금액 쪽이 다른 값일 수 있음. "
                    f"04 기초금액 기준 합계는 API값을 쓰므로 기초금액 기준 합계 확인{tail}")
        if v_any is not None and v_any >= 0 and same(a, api_p + v_any):     # API 추정가격 + 부가세 = API 기초금액 → 조사금액·관급 차이가 아니라 면세·부분면세('×1.1' 행의 참고와 같은 판단)
            return (f"문서값 = 추정가격×{k:g}, API 기초금액 = API 추정가격 + 부가세(부가세율 {v_any / api_p:.1%}) — API 쪽은 면세·부분면세로 서로 맞음: "
                    f"문서 기초금액이 부가세 {vat:.0%} 를 더한 금액으로 보임({more('04 과세구분이 맞는지 확인')})")
        return f"문서값 = 추정가격×{k:g}(추정가격+부가세) — API 기초금액이 이와 다름: 조달청 조사금액 기준이거나 관급 포함 여부 차이로 보임({use})"
    if item == "기초금액" and api_p and api_p > 0 and doc_v is not None and doc_v > 0 and same(d, api_p + doc_v):
        # 문서 기초금액 = 추정가격 + 문서 부가세(10% 가 아닌 부가세 — 10% 면 위 규칙) → 문서는 면세·부분면세로 서로 맞는다(실자료 20190301453·20190301633·20190819618)
        return (f"문서값 = 추정가격 + 문서 부가세(부가세율 {doc_v / api_p:.1%}) — 문서 쪽은 면세·부분면세로 서로 맞음: API 기초금액이 이와 다름"
                f"(부가세 {vat:.0%} 를 더한 금액이거나 조달청 조사금액 기준일 수 있음. {use}, 과세 여부를 공고문에서 확인)")
    if near(k):
        if item == "기초금액" and same(a, (api_p or 0) * k):      # API 기초금액 = API 추정가격×1.1 → API 기초금액은 부가세 포함 금액이 맞다
            return f"문서값 = API값×{k:g} — API 기초금액 = API 추정가격×{k:g}(API 쪽은 서로 맞음): 문서에서 부가세를 한 번 더 더한 금액을 읽은 것으로 보임({use}, 공고문 확인)"
        if item == "기초금액":           # 기초금액은 원래 부가세 포함 금액이다 → '부가세 포함 금액을 읽음' 은 설명이 되지 않는다
            return f"문서값 = API값×{k:g} — API 기초금액이 부가세를 뺀 금액이거나 문서에서 부가세를 한 번 더 더한 금액을 읽은 것으로 보임({use}, 공고문 확인)"
        return f"문서값 = API값×{k:g} — 문서에서 부가세 포함 금액을 읽은 것으로 보임({use})"
    if near(1 / k):
        if item == "추정가격" and (api_fit or exempt):    # API 추정가격이 기초금액과 맞는다(×1.1, 또는 추정가격+부가세) → API 추정가격은 부가세를 뺀 금액이 맞다
            why = api_fit or f"API 추정가격 + 부가세 = {bsrc} 기초금액" + ("(API 쪽은 면세·부분면세로 서로 맞음)" if has_api_b else "(면세·부분면세로 서로 맞음)")
            return f"문서값 = API값÷{k:g} — {why}: 문서에서 부가세를 한 번 더 뺀 금액을 읽은 것으로 보임({use})"
        if item == "추정가격":           # 추정가격은 원래 부가세를 뺀 금액이다 → 어긋난 쪽은 API
            return f"문서값 = API값÷{k:g} — API 추정가격이 부가세 포함 금액으로 보임({more('추정가격 기준 합계 확인')})"
        if item == "기초금액":           # 같은 공고의 추정가격을 알고 그것이 문서값과 다르면 '(추정가격)' 이라고 적지 않는다
            known = [x for x in (api_p, doc_p) if x and x > 0]
            what = "" if known and not any(same(d, x) for x in known) else "(추정가격)"
            return f"문서값 = API값÷{k:g} — 문서에서 부가세를 뺀 금액{what}을 기초금액으로 읽은 것으로 보임({use})"
        return f"문서값 = API값÷{k:g} — 문서에서 부가세를 뺀 금액을 읽었거나 API값이 부가세 포함 금액({use})"
    for m, txt, why in ((10, "10배", "자릿수"), (100, "100배", "자릿수"), (1000, "1,000배", "단위(원/천원)"),
                        (0.1, "1/10", "자릿수"), (0.01, "1/100", "자릿수"), (0.001, "1/1,000", "단위(원/천원)")):
        if near(m):
            return f"문서값이 API값의 {txt} — {why} 오독 의심({use})"
    if same(doc_p, (api_p or 0) * q) and same(doc_b, (api_b or 0) * q):
        return (f"API값과 문서값이 5% 넘게 다름 — 추정가격·기초금액이 같은 비율(×{_ratio_txt(q)})로 다름: 문서가 이 공고와 다른 금액(일부 공구·차수분·변경 전후 금액)을 "
                f"적고 있을 수 있음({use}, 공고문 확인)")
    return f"API값과 문서값이 5% 넘게 다름 — 변경공고로 금액이 바뀌었거나 문서에서 다른 금액을 읽었을 수 있음({use}, 공고문 확인)"


def cross_verify(bid_no: str, api_row: Dict, doc: Dict, warn_ratio: float = 0.005, vat: float = 0.10) -> List[Dict]:
    """API 값과 문서 추출값 비교 → 검증로그 rows. 숫자가 문자열/NaN 이어도 안전."""
    logs: List[Dict] = []

    def _cmp(item: str, a: Optional[float], d: Optional[float], ratio_warn: float, note: str = ""):
        if a is None and d is None:
            logs.append({"공고번호": bid_no, "항목": item, "API값": None, "문서값": None, "판정": "미확인", "비고": "양쪽 모두 없음"})
            return
        if a is None or d is None:
            logs.append({"공고번호": bid_no, "항목": item, "API값": a, "문서값": d, "판정": "참고",
                         "비고": "한쪽만 존재 → 존재하는 값 사용"})
            return
        diff = d - a
        r = diff / a if a else (0.0 if diff == 0 else 1.0)
        verdict = "정상" if abs(r) <= ratio_warn else ("경고" if abs(r) <= 0.05 else "오류")
        if verdict == "오류" and not note:
            note = _gap_hint(a, d, vat, item, {"api_p": api_p, "api_b": api_b, "api_v": _num(api_row.get("부가세_API")),
                                               "doc_p": doc_p, "doc_b": doc_b, "doc_v": _num(doc.get("부가가치세_원"))})
        logs.append({"공고번호": bid_no, "항목": item, "API값": a, "문서값": d, "차이": diff,
                     "차이율": round(r, 5), "판정": verdict, "비고": note})

    api_p, api_b = _num(api_row.get("추정가격")), _num(api_row.get("기초금액"))
    doc_p, doc_b = _num(doc.get("추정가격_원")), _num(doc.get("기초금액_원"))
    _cmp("추정가격", api_p, doc_p, warn_ratio)
    _cmp("기초금액", api_b, doc_b, warn_ratio)

    p, b = api_p or doc_p, api_b or doc_b
    if p and b:
        exp = p * (1 + vat)
        r = (b - exp) / exp
        verdict, note = ("정상", "") if abs(r) <= 0.01 else ("경고", "부가세 비율 불일치(면세·관급 포함 여부 확인)")
        if verdict == "경고":
            # 10% 가 아닌 이유가 설명되면 경고 대신 참고: ① 공고의 부가세(API·문서)가 기초금액−추정가격과 맞는 면세·부분면세 ② 조달청 공고(기초금액 = 조사금액 기준) ③ 공동주택(국민주택 면세 가능)
            rate = (b - p) / p
            v_api, v_doc = _num(api_row.get("부가세_API")), _num(doc.get("부가가치세_원"))
            v = v_api if v_api is not None else v_doc
            low = -0.001 <= rate < (vat - 0.001)
            if low and v is not None and abs((p + v) - b) <= b * 0.01:
                verdict, note = "참고", f"부가세율 {max(rate, 0):.1%} — 공고에 적힌 부가세와 일치(면세·부분면세 공사, 조세특례제한법 제106조 등)"
            elif "조달청" in str(api_row.get("공고기관") or "") and 0.9 <= b / exp < 1.0:
                verdict, note = "참고", "조달청 공고: 기초금액이 조사금액 기준이라 추정가격×1.1 보다 낮음(면세 아님일 수 있음) → 04 의 추정가격 기준 합계 참조"
                if any(l["항목"] == "추정가격" and "추정가격 기준 합계 확인" in str(l.get("비고")) for l in logs):
                    note += " (단, 추정가격 행 비고대로 API 추정가격이 어긋난 것이면 04 의 추정가격·부가세나 추정가격 기준 합계가 틀린 쪽 — 공고문 확인)"
            elif low and str(api_row.get("표2_대분류") or "") == "공동주택":
                verdict, note = "참고", f"부가세율 {max(rate, 0):.1%} — 공동주택: 국민주택규모 이하 면세·부분면세(조세특례제한법 제106조) 가능"
        logs.append({"공고번호": bid_no, "항목": "기초금액≈추정가격×1.1", "API값": round(exp), "문서값": b, "차이": round(b - exp),
                     "차이율": round(r, 5), "판정": verdict, "비고": note})
    rt, dt_ = api_row.get("사업유형"), doc.get("사업유형")
    if rt and dt_ and dt_ not in ("기타",) and rt != dt_:
        logs.append({"공고번호": bid_no, "항목": "사업유형 일치", "API값": rt, "문서값": dt_, "판정": "경고",
                     "비고": "공고명 규칙 분류와 공고문 판독이 다름 → 프로젝트 분리(신축/리모델링) 확인"})
    area = _num(doc.get("연면적_m2"))
    if area is not None and (area < 50 or area > 500000):
        logs.append({"공고번호": bid_no, "항목": "연면적 범위", "문서값": area, "판정": "경고", "비고": "㎡/평 단위 혼동 또는 오독 의심"})
    fl = _num(doc.get("지상층수"))
    if fl is not None and fl > 60:
        logs.append({"공고번호": bid_no, "항목": "층수 범위", "문서값": fl, "판정": "경고", "비고": "층수 오독 의심"})
    for item, ak, dk in (("도급자관급액", "도급자관급액_API", "도급자관급액_원"), ("관급자관급액", "관급자관급액_API", "관급자관급액_원")):
        a, dv = _num(api_row.get(ak)), _num(doc.get(dk))
        if a is None and dv is None:
            continue
        if a is not None and dv is not None and (a or dv):
            base = a if a else dv
            r = (dv - a) / base
            logs.append({"공고번호": bid_no, "항목": item, "API값": a, "문서값": dv, "차이": dv - a, "차이율": round(r, 5),
                         "판정": "정상" if abs(r) <= 0.01 else "경고", "비고": "" if abs(r) <= 0.01 else "관급자재 금액 불일치(API 우선 사용)"})
        else:
            logs.append({"공고번호": bid_no, "항목": item, "API값": a, "문서값": dv, "판정": "참고", "비고": "한쪽만 존재 → 존재하는 값 사용"})
    tot = _num(doc.get("총공사금액_원"))
    if tot and b:
        comp = b + (_num(doc.get("도급자관급액_원")) or 0) + (_num(doc.get("관급자관급액_원")) or 0)
        r = (tot - comp) / comp if comp else 0
        logs.append({"공고번호": bid_no, "항목": "총공사금액=기초금액+관급", "API값": comp, "문서값": tot, "차이": tot - comp,
                     "차이율": round(r, 5), "판정": "정상" if abs(r) <= 0.01 else "경고", "비고": "금액 구성 정합성"})
    return logs


# ── 보강 추출(3차): 기타금액(한전불입금 등)·면세 문구·조달청 '공사비 정보광장' 공사개요 항목 ─────────────────────
# 1차 추출 결과(llm_docs.json)는 건드리지 않고 data/llm_extra.json 에 따로 저장한다. 1차에 없는 항목만 묻고, 대지면적·건축면적은 1차에서 비었을 때 채우는 용도.
EXTRA_FIELDS = {
    "대지면적_m2": "number|null", "건축면적_m2": "number|null", "조경면적_m2": "number|null",
    "건폐율_pct": "number|null(%)", "용적률_pct": "number|null(%)", "기준층층고_m": "number|null", "최고높이_m": "number|null",
    "주차대수": "integer|null", "세대수": "integer|null", "건축법상용도": "string", "설계특이사항": "string",
    "기타금액_원": "integer|null", "기타금액_항목": "string", "기타금액_위치": "별도|관급에포함|추정가격에포함|불명",
    "면세_문구": "string", "근거문구": "object(항목명→원문 인용 30자 이내)", "신뢰도": "high|medium|low",
}
_EXTRA_FLOAT = ["대지면적_m2", "건축면적_m2", "조경면적_m2", "건폐율_pct", "용적률_pct", "기준층층고_m", "최고높이_m"]
_EXTRA_INT = ["주차대수", "세대수", "기타금액_원"]
_EXTRA_STR = ["건축법상용도", "설계특이사항", "기타금액_항목", "면세_문구"]
EXTRA_SYSTEM = (
    "당신은 한국 공공건축 공사 입찰공고문·현장설명서·과업지시서에서 공사개요의 세부 항목과 공사금액의 부대 항목을 정확히 추출하는 분석가입니다. "
    "반드시 JSON 객체 하나만 출력합니다(마크다운 코드블록·설명 금지). 문서에 없는 값은 빈 문자열로 둡니다 — 추정하거나 계산해 만들지 않습니다. "
    "숫자 항목은 쉼표·단위 없이 자릿수만 적습니다. 면적은 ㎡(평→×3.3058), 건폐율·용적률은 % 숫자(예: 12.35), 높이·층고는 m, 주차대수는 대(면) 수, 금액은 원 단위 정수(천원 단위 표는 ×1,000, 백만원 단위는 ×1,000,000 으로 환산). "
    "건폐율·용적률은 이 공사의 설계값만 적고, '60% 이하' 같은 법정 한도·허용 기준은 적지 않습니다. 주차대수도 '법정 주차대수 이상' 같은 기준이면 비웁니다. "
    "'기준층층고_m' 은 기준층의 층고, '최고높이_m' 은 건축물 최고높이입니다. '건축법상용도' 는 문서에 적힌 건축물 용도(예: 업무시설(공공업무시설), 문화 및 집회시설)를 그대로 적습니다. "
    "'설계특이사항' 은 문서에 적힌 범위에서 동 수·주요 구조형식·특수 공법·인증·지반·토공 특이점을 80자 이내로 요약하고, 없으면 비웁니다. "
    "'기타금액_원' 은 추정가격·부가가치세·관급자재비와 별도로 문서에 적힌 부대 금액의 합입니다: 한전불입금·한전수탁비·표준시설부담금·시설분담금·상수도 원인자부담금·가스 시설분담금·통신 가입비·이설비 등. "
    "'기타금액_항목' 에는 항목명과 금액을 '한전불입금 5,600,000; 상수도분담금 1,200,000' 처럼 적습니다. "
    "'기타금액_위치' 는 그 금액이 문서에서 어디에 속하는지입니다: 도급·관급 금액 밖의 별도 금액이면 '별도', 관급자재비 안에 포함된 것으로 적혀 있으면 '관급에포함', "
    "추정가격(도급액) 안에 포함된 것으로 적혀 있으면 '추정가격에포함', 알 수 없으면 '불명'. 기타 금액이 없으면 세 항목 모두 비웁니다. 관급자재비·부가가치세 자체는 기타금액이 아닙니다. "
    "'면세_문구' 에는 부가가치세 면제·면세 사업임을 밝힌 문구(예: '조세특례제한법 제106조에 의한 부가가치세 면제 대상')를 60자 이내로 인용하고, 없으면 비웁니다. "
    "각 값의 근거가 된 원문 문구를 '근거문구'에 항목명별로 짧게 인용합니다(근거문구는 [{\"항목\": …, \"원문\": …}] 배열)."
)
# 발췌 키워드 2단: 찾는 항목에만 쓰이는 낱말을 먼저 담고, 남는 분량에 흔한 낱말(용도·구조·관급 …) 주변을 담는다.
# 한 단으로 섞으면 흔한 낱말이 문서 앞쪽에서 예산을 다 써서 뒤쪽의 건축개요·한전불입금 구간이 빠진다
_EXTRA_KEYS_MAIN = ["건축개요", "건축 개요", "설계개요", "설계 개요", "공사개요", "공사 개요", "사업개요", "대지면적", "대지 면적", "부지면적", "건축면적", "조경면적",
                    "건폐율", "용적률", "층고", "최고높이", "건물높이", "주차대수", "주차 대수", "한전", "불입금", "분담금", "부담금", "수탁비", "납입금",
                    "표준공사비", "기타금액", "기타 금액", "이설비", "조세특례", "면세", "부가가치세 면제", "부가세 면제"]
_EXTRA_KEYS_SUB = ["높이", "주차", "세대", "용도", "구조", "별도", "면제", "부가가치세", "부가세", "추정금액", "총공사", "관급"]
_EXTRA_KEYS = _EXTRA_KEYS_MAIN + _EXTRA_KEYS_SUB
EXTRA_INPUT_CHARS = 40000


def extra_output_schema() -> Dict:
    """보강 추출 구조화 출력 스키마. 1차와 같은 제약(모두 string/enum·required, union 0)."""
    props: Dict[str, Dict] = {k: {"type": "string"} for k in _EXTRA_FLOAT + _EXTRA_INT + _EXTRA_STR}
    props["기타금액_위치"] = {"type": "string", "enum": ["별도", "관급에포함", "추정가격에포함", "불명", ""]}
    props["신뢰도"] = {"type": "string", "enum": ["high", "medium", "low"]}
    props["근거문구"] = {"type": "array", "items": {"type": "object", "properties": {"항목": {"type": "string"}, "원문": {"type": "string"}},
                                                "required": ["항목", "원문"], "additionalProperties": False}}
    return {"type": "object", "properties": props, "required": list(props.keys()), "additionalProperties": False}


def build_extra_prompt(text: str, api_hint: Dict, cfg_llm: Dict) -> str:
    hint = {k: jsonable(v) for k, v in (api_hint or {}).items()}
    limit = int(cfg_llm.get("extra_max_input_chars", EXTRA_INPUT_CHARS))
    return (f"[공고 메타데이터(API)]\n{json.dumps(hint, ensure_ascii=False)}\n\n"
            f"[출력 스키마]\n{json.dumps(EXTRA_FIELDS, ensure_ascii=False, indent=1)}\n\n"
            f"[공고문 텍스트]\n{focus_tiered(text, limit, [_EXTRA_KEYS_MAIN, _EXTRA_KEYS_SUB])}")


def build_extra_request_params(text: str, api_hint: Dict, cfg_llm: Dict, structured: Optional[bool] = None,
                               max_tokens: Optional[int] = None) -> Dict:
    if structured is None:
        structured = bool(cfg_llm.get("structured_output", True))
    params = dict(model=cfg_llm.get("model", "claude-sonnet-5"),
                  max_tokens=int(max_tokens or cfg_llm.get("max_tokens", DEFAULT_MAX_TOKENS)),
                  system=EXTRA_SYSTEM, messages=[{"role": "user", "content": build_extra_prompt(text, api_hint, cfg_llm)}])
    if structured:
        params["output_config"] = {"format": {"type": "json_schema", "schema": extra_output_schema()}}
    return params


_EMPTY_WORDS = ("null", "none", "nan", "없음", "-", "")


def _etc_where(w) -> str:
    """기타금액_위치 표기 정리: 스키마 밖 표기('관급 포함', '도급액에 포함', '도급액 별도')도 뜻에 맞는 값으로. 알 수 없으면 ''.
    '포함' 이라고 적힌 경우에만 포함으로 본다('관급 미포함(별도)'·'관급자재와 별도' 는 별도)."""
    w = str(w or "").replace(" ", "")
    if w in ("별도", "관급에포함", "추정가격에포함", "불명"):
        return w
    inc = "포함" in w.replace("미포함", "").replace("불포함", "")
    if "별도" in w and not inc:
        return "별도"
    if inc and "관급" in w:
        return "관급에포함"
    if inc and ("추정가격" in w or "도급" in w):
        return "추정가격에포함"
    return ""


def normalize_extra_doc(data: Dict) -> Dict:
    """보강 추출 출력 정규화: 숫자 필드 → 숫자(단위가 붙은 표기도 해석, 0 이하·비현실 값은 비움), 근거문구 배열 → dict, 기타금액이 없으면 항목·위치도 비움.
    숫자로 읽지 못한 원문은 '_raw' 에 남긴다 → 저장된 결과에 다시 적용해도 같은 결과(멱등)이고, 해석 규칙을 고친 뒤 excel 만 다시 돌려도 값이 살아난다."""
    out = dict(data)
    raw = dict(out.get("_raw") or {}) if isinstance(out.get("_raw"), dict) else {}
    for k in _EXTRA_FLOAT + _EXTRA_INT:
        src = out.get(k)
        if (src is None or src == "") and raw.get(k):
            src = raw[k]
        x = _num_loose(src, money=(k == "기타금액_원"))
        if x is None and isinstance(src, str) and src.strip().lower() not in _EMPTY_WORDS:
            raw[k] = src.strip()[:60]
        else:
            raw.pop(k, None)
        if x is not None and x > 0:
            out[k] = int(round(x)) if k in _EXTRA_INT else x
        else:
            out[k] = None
    # 건폐율을 0.1235 처럼 비율로 적은 경우: 같은 문서의 건축면적/대지면적과 ×100 값이 맞을 때만 % 로 바꾼다.
    # (1% 미만 건폐율·용적률은 추모공원·운동장·온실처럼 대지가 큰 시설에서 실제로 나오므로 무조건 ×100 하지 않는다)
    bc, site, bld = out.get("건폐율_pct"), out.get("대지면적_m2"), out.get("건축면적_m2")
    if bc is not None and bc < 1 and site and bld:
        calc = bld / site * 100
        if abs(bc * 100 - calc) <= calc * 0.05 < abs(bc - calc):
            out["건폐율_pct"] = round(bc * 100, 4)
    if out.get("건폐율_pct") is not None and out["건폐율_pct"] > 100:
        out["건폐율_pct"] = None
    for k in _EXTRA_STR:
        v = str(out.get(k) or "").strip()
        out[k] = "" if v.lower() in _EMPTY_WORDS else v
    w = _etc_where(out.get("기타금액_위치") or raw.get("기타금액_위치"))
    if out.get("기타금액_원"):
        out["기타금액_위치"] = w or "불명"
        raw.pop("기타금액_위치", None)
    else:
        out["기타금액_위치"] = ""
        if raw.get("기타금액_원") and w:              # 금액을 읽지 못한 동안에도 위치는 보존(나중에 금액 해석이 되면 함께 살아남)
            raw["기타금액_위치"] = w
        else:
            raw.pop("기타금액_위치", None)
    ev = out.get("근거문구")
    if isinstance(ev, list):
        out["근거문구"] = {str(e.get("항목", "")): str(e.get("원문", "")) for e in ev if isinstance(e, dict)}
    elif not isinstance(ev, dict):
        out["근거문구"] = {"원문": str(ev)} if ev else {}
    out["신뢰도"] = str(out.get("신뢰도") or "")
    if raw:
        out["_raw"] = raw
    else:
        out.pop("_raw", None)
    return out


def parse_extra_response_text(out: str, model: str = "") -> Dict:
    data = normalize_extra_doc(_parse_json(out))
    data["_model"] = model
    return data


def extract_extra_with_claude(text: str, api_hint: Dict, cfg_llm: Dict, max_tokens: Optional[int] = None) -> Dict:
    """공고문 텍스트 → 보강 추출 dict(즉시 호출)."""
    import anthropic  # type: ignore
    client = anthropic.Anthropic(api_key=os.environ.get(cfg_llm.get("api_key_env", "ANTHROPIC_API_KEY")) or None)
    kwargs = dict(model=cfg_llm.get("model", "claude-sonnet-5"),
                  max_tokens=int(max_tokens or cfg_llm.get("max_tokens", DEFAULT_MAX_TOKENS)),
                  system=EXTRA_SYSTEM, messages=[{"role": "user", "content": build_extra_prompt(text, api_hint, cfg_llm)}])
    create = _structured_create(client, extra_output_schema()) if bool(cfg_llm.get("structured_output", True)) else (lambda **kw: _messages_call(client, **kw))
    resp = _finish_call(create, kwargs)
    out = "".join(getattr(b, "text", "") for b in resp.content if getattr(b, "type", "") == "text")
    data = parse_extra_response_text(out, kwargs["model"])
    data["_usage_in"] = getattr(getattr(resp, "usage", None), "input_tokens", None)
    data["_usage_out"] = getattr(getattr(resp, "usage", None), "output_tokens", None)
    return data


ETC_MAX_SHARE = 0.30      # 기타금액이 기초금액의 이 비율을 넘으면 도급액·관급자재를 잘못 읽은 것으로 보고 합계에 더하지 않는다


def cross_verify_extra(bid_no: str, api_row: Dict, doc: Dict, xdoc: Dict) -> List[Dict]:
    """보강 추출 정합성: 건폐율 = 건축면적/대지면적, 대지면적 ≥ 건축면적·조경면적, 1차 추출과 면적 일치, 층고·높이 범위, 숫자 해석 실패,
    기타금액과 예산금액 잔차 대조·기초금액 대비 크기, 면세 문구와 부가세율."""
    logs: List[Dict] = []
    doc = doc or {}
    for k, v in (xdoc.get("_raw") or {}).items():
        if k != "기타금액_위치":
            logs.append({"공고번호": bid_no, "항목": f"숫자 해석 실패({k})", "판정": "경고", "비고": f"문서 표기 '{v}' 를 숫자로 읽지 못해 비움 → 원문 확인"})
    site = _num(xdoc.get("대지면적_m2")) or _num(doc.get("대지면적_m2"))
    bld = _num(xdoc.get("건축면적_m2")) or _num(doc.get("건축면적_m2"))
    for k in ("대지면적_m2", "건축면적_m2"):
        a, b = _num(doc.get(k)), _num(xdoc.get(k))
        if a and b and abs(a - b) / a > 0.01:
            logs.append({"공고번호": bid_no, "항목": f"{k} 1차·보강 불일치", "API값": a, "문서값": b, "차이": b - a, "차이율": round((b - a) / a, 5),
                         "판정": "경고", "비고": "1차 추출값과 보강 추출값이 다름 → 01 시트는 1차 값을 사용, 원문 확인"})
    if site and bld and bld > site * 1.001:
        logs.append({"공고번호": bid_no, "항목": "대지면적≥건축면적", "API값": site, "문서값": bld, "판정": "경고", "비고": "건축면적이 대지면적보다 큼 → 단위·오독 의심"})
    land = _num(xdoc.get("조경면적_m2"))
    if site and land and land > site * 1.001:
        logs.append({"공고번호": bid_no, "항목": "대지면적≥조경면적", "API값": site, "문서값": land, "판정": "경고", "비고": "조경면적이 대지면적보다 큼 → 단위·오독 의심"})
    bc = _num(xdoc.get("건폐율_pct"))
    if bc and site and bld:
        calc = bld / site * 100
        r = (bc - calc) / calc
        ratio = abs(r) > 0.02 and abs(bc * 100 - calc) <= calc * 0.05
        logs.append({"공고번호": bid_no, "항목": "건폐율=건축면적/대지면적", "API값": round(calc, 2), "문서값": bc, "차이": round(bc - calc, 2), "차이율": round(r, 5),
                     "판정": "정상" if abs(r) <= 0.02 else "경고",
                     "비고": "" if abs(r) <= 0.02 else ("문서가 건폐율을 비율(0.xx)로 적은 것으로 보임 — ×100 하면 계산값과 일치" if ratio
                                                       else "문서 건폐율과 면적으로 계산한 값이 다름(법정 한도 표기 또는 면적 오독 의심)")})
    elif bc and bc < 1:
        logs.append({"공고번호": bid_no, "항목": "건폐율 1% 미만", "문서값": bc, "판정": "참고", "비고": "면적으로 대조할 수 없음 — 비율(0.xx) 표기일 수 있으니 원문 확인"})
    vl = _num(xdoc.get("용적률_pct"))
    if vl and vl > 1500:
        logs.append({"공고번호": bid_no, "항목": "용적률 범위", "문서값": vl, "판정": "경고", "비고": "용적률 1,500% 초과 → 오독 의심"})
    elif vl and vl < 1:
        logs.append({"공고번호": bid_no, "항목": "용적률 1% 미만", "문서값": vl, "판정": "참고", "비고": "비율(0.xx) 표기일 수 있으니 원문 확인"})
    for k, hi, what in (("기준층층고_m", 20, "층고 20m 초과"), ("최고높이_m", 300, "최고높이 300m 초과")):
        hv = _num(xdoc.get(k))
        if hv and hv > hi:
            logs.append({"공고번호": bid_no, "항목": f"{k} 범위", "문서값": hv, "판정": "경고", "비고": f"{what} → 단위(mm·cm)·오독 의심"})
    etc = _num(xdoc.get("기타금액_원"))
    if etc:
        b, p = _num(api_row.get("기초금액")), _num(api_row.get("추정가격"))
        gov = (_num(api_row.get("도급자관급액_API")) or 0) + (_num(api_row.get("관급자관급액_API")) or 0) or (_num(api_row.get("관급자재_API")) or 0)
        budget = _num(api_row.get("예산금액"))
        base = b or (p * 1.1 if p else None)
        note, verdict = f"{xdoc.get('기타금액_항목') or ''} (위치: {xdoc.get('기타금액_위치') or '불명'})".strip(), "참고"
        if budget and base and abs(budget - base - gov - etc) <= max(1.0, etc * 0.01):
            note += " — 예산금액(API) − 기초금액 − 관급 잔차와 일치"
        elif base and etc > base * ETC_MAX_SHARE:
            verdict, note = "경고", note + f" — 기타금액이 기초금액의 {ETC_MAX_SHARE:.0%} 초과 → 관급자재·도급액을 기타금액으로 읽었는지 확인(금액 합계에는 더하지 않음)"
        logs.append({"공고번호": bid_no, "항목": "기타금액(문서)", "문서값": etc, "판정": verdict, "비고": note})
    ex = str(xdoc.get("면세_문구") or "").strip()
    if ex:
        p, b = _num(api_row.get("추정가격")), _num(api_row.get("기초금액"))
        rate = (b - p) / p if (p and b) else None
        ok = rate is not None and rate < 0.099
        logs.append({"공고번호": bid_no, "항목": "면세 문구(문서)", "API값": p, "문서값": b, "차이율": round(rate, 5) if rate is not None else None,
                     "판정": "참고" if (ok or rate is None) else "경고",
                     "비고": f"'{ex[:60]}'" + ("" if (ok or rate is None) else " — 면세 문구가 있는데 기초금액이 추정가격×1.1 수준(부가세 10%) → 확인")})
    return logs


# ── 파일 인수인계 모드 (Cowork·사람이 채우는 방식) ─────────────────────
EXPORT_README = """# LLM 추출 작업지시 (Cowork / 사람용)

이 폴더(`llm_in/`)의 `*.txt` 파일 하나가 공고 1건입니다. 각 파일 맨 위의 [작업지시]·[출력 JSON 양식]을 따라
같은 이름의 JSON 파일을 `../llm_out/<같은 이름>.json` 으로 저장하세요. 예) `llm_in/R26BK0001-000.txt` → `llm_out/R26BK0001-000.json`

규칙 요약
- 문서에 없는 값은 null. 추정하지 않습니다.
- 금액은 원 단위 정수(예: '1,234백만원' → 1234000000, '12.3억원' → 1230000000, 천원 단위 표는 ×1000).
- 면적은 ㎡ 숫자(평 → ×3.3058), 층수는 정수. 리모델링이면 연면적_m2 에 공사 대상 면적을 넣고 근거문구에 '대상면적' 표시.
- 기초금액 = 통상 추정가격 + 부가가치세. 관급자재(도급자관급액·관급자관급액)는 별도 항목.
- 근거문구: {"항목명": "원문 인용 30자 이내", ...} 형태의 객체.
- JSON 외의 설명·코드블록 표시(```)는 넣지 않습니다(있어도 무시되지만 깨끗한 편이 안전).

Cowork/Claude Code 에 붙여 넣을 지시문 예시
  "llm_in 폴더의 txt 파일을 하나씩 읽고, 각 파일 상단의 작업지시대로 JSON 을 만들어 llm_out 폴더에 같은 이름의 .json 으로 저장해 줘.
   한 번에 20개씩 처리하고, 끝나면 몇 개 처리했는지 알려 줘."

모두 채운 뒤 프로젝트 폴더에서:  python -m g2b_costdb.pipeline extract --import
"""


def export_template() -> Dict:
    t: Dict = {}
    for k in _STR_FIELDS:
        t[k] = None
    for k in _NUM_FIELDS:
        t[k] = None
    t["사업유형"] = "신축|증축|리모델링|증축·리모델링|유지보수|기타 중 하나"
    t["공종"] = "건축|전기|정보통신|소방|조경|기계설비|토목|기타 중 하나"
    t["근거문구"] = {"연면적_m2": "원문 인용", "추정가격_원": "원문 인용"}
    t["신뢰도"] = "high|medium|low 중 하나"
    return t


def export_prompt_file(key: str, text: str, api_hint: Dict, cfg_llm: Dict) -> str:
    hint = {k: jsonable(v) for k, v in (api_hint or {}).items()}
    return (f"# 공고 추출 작업지시 — {key}\n\n"
            f"[메타데이터]\n{json.dumps(hint, ensure_ascii=False)}\n\n"
            f"[작업지시]\n{SYSTEM}\n"
            f"결과는 llm_out/{key}.json 파일로 저장합니다.\n\n"
            f"[출력 JSON 양식]\n{json.dumps(export_template(), ensure_ascii=False, indent=1)}\n\n"
            f"[공고문 텍스트]\n{focus_text(text, int(cfg_llm.get('max_input_chars', 60000)))}\n")


def export_prompts(keys: List[str], texts: Dict[str, str], hints: Dict[str, Dict], cfg_llm: Dict, in_dir: str, out_dir: str) -> int:
    os.makedirs(in_dir, exist_ok=True)
    os.makedirs(out_dir, exist_ok=True)
    with open(os.path.join(in_dir, "README_지시문.md"), "w", encoding="utf-8") as f:
        f.write(EXPORT_README)
    with open(os.path.join(in_dir, "_schema.json"), "w", encoding="utf-8") as f:
        json.dump(output_schema(), f, ensure_ascii=False, indent=1)
    n = 0
    for k in keys:
        with open(os.path.join(in_dir, f"{k}.txt"), "w", encoding="utf-8") as f:
            f.write(export_prompt_file(k, texts[k], hints.get(k, {}), cfg_llm))
        n += 1
    return n


def import_results(out_dir: str, valid_keys: Optional[set] = None) -> Tuple[Dict[str, Dict], List[Tuple[str, str]]]:
    """llm_out/*.json → {공고키: doc}. 반환 (docs, [(파일명, 오류)])."""
    docs: Dict[str, Dict] = {}
    errors: List[Tuple[str, str]] = []
    if not os.path.isdir(out_dir):
        return docs, errors
    for fn in sorted(os.listdir(out_dir)):
        if not fn.lower().endswith(".json") or fn.startswith("_"):
            continue
        key = fn[:-5]
        if valid_keys is not None and key not in valid_keys:
            errors.append((fn, "대표 공고 목록에 없는 공고키(파일명 확인)"))
            continue
        try:
            raw = open(os.path.join(out_dir, fn), encoding="utf-8-sig").read()
            doc = normalize_doc(_parse_json(raw))
        except Exception as e:  # noqa: BLE001
            errors.append((fn, f"JSON 해석 실패: {str(e)[:120]}"))
            continue
        filled = [k for k in _NUM_FIELDS + ["공사명", "구조"] if doc.get(k) not in (None, "", 0)]
        if not filled:
            errors.append((fn, "값이 하나도 없음(전부 null)"))
            continue
        doc["_model"] = "manual/cowork"
        doc["_source_files"] = fn
        docs[key] = doc
    return docs, errors


# ── Message Batches (50% 할인, 최대 24시간) ────────────────────────
BATCH_CHUNK = 500


def submit_batches(client, items: List[Tuple[str, str, Dict]], cfg_llm: Dict, chunk: int = BATCH_CHUNK,
                   max_tokens_by_key: Optional[Dict[str, int]] = None, params_builder=None, on_submitted=None) -> List[Dict]:
    """items: [(공고키, 텍스트, 힌트)]. 반환: [{'id', 'keys', 'status'}]. 구조화 출력이 거부되면 일반 요청으로 재시도.
    max_tokens_by_key: 공고키별 max_tokens 재정의(이전에 max_tokens 에서 잘린 건은 2배로 재제출).
    params_builder: 요청 파라미터 생성 함수(기본 build_request_params, 2차 공종 추출은 build_trade_request_params)."""
    out = []
    mt = max_tokens_by_key or {}
    build = params_builder or build_request_params
    for i in range(0, len(items), chunk):
        part = items[i:i + chunk]
        for structured in (bool(cfg_llm.get("structured_output", True)), False):
            reqs = [{"custom_id": k, "params": build(t, h, cfg_llm, structured, max_tokens=mt.get(k))} for k, t, h in part]
            try:
                b = client.messages.batches.create(requests=reqs)
                out.append({"id": b.id, "keys": [k for k, _, _ in part], "status": "submitted", "structured": structured})
                if on_submitted:
                    on_submitted(out[-1])                 # 청크마다 상태 파일에 바로 기록 → 뒤 청크가 실패해도 앞 배치 ID 를 잃지 않는다
                break
            except Exception as e:  # noqa: BLE001
                if structured and "output_config" in str(e) or (structured and "400" in str(e)):
                    log.warning("구조화 출력 배치가 거부됨(%s) → 일반 요청으로 재제출", str(e)[:120])
                    continue
                e.submitted = out                           # 이미 만들어진 배치를 호출부가 저장할 수 있게 예외에 붙여 전달
                raise
    return out


def fetch_batch(client, batch_id: str, model: str = "", parser=None) -> Tuple[str, Dict[str, Dict], Dict[str, int]]:
    """반환 (processing_status, {공고키: doc 또는 {'_error':…}}, request_counts dict). parser: 응답 텍스트 → dict(기본 parse_response_text)."""
    parse = parser or parse_response_text
    b = client.messages.batches.retrieve(batch_id)
    counts = {}
    rc = getattr(b, "request_counts", None)
    for name in ("processing", "succeeded", "errored", "canceled", "expired"):
        counts[name] = int(getattr(rc, name, 0) or 0) if rc is not None else 0
    if b.processing_status != "ended":
        return b.processing_status, {}, counts
    docs: Dict[str, Dict] = {}
    for r in client.messages.batches.results(batch_id):
        key = r.custom_id
        rtype = getattr(r.result, "type", "")
        if rtype == "succeeded":
            msg = r.result.message
            text = "".join(getattr(blk, "text", "") for blk in msg.content if getattr(blk, "type", "") == "text")
            try:
                d = parse(text, model)
                d["_usage_in"] = getattr(getattr(msg, "usage", None), "input_tokens", None)
                d["_usage_out"] = getattr(getattr(msg, "usage", None), "output_tokens", None)
                d["_batch_id"] = batch_id
                docs[key] = d
            except Exception as e:  # noqa: BLE001
                # 원인 진단용: 종료 사유(max_tokens 면 잘림, refusal 이면 거부)와 응답 앞부분을 함께 남긴다
                stop = getattr(msg, "stop_reason", "") or ""
                head = (text or "").strip().replace("\n", " ")[:80]
                docs[key] = {"_error": f"응답 해석 실패(stop={stop}): {str(e)[:80]} | 응답: {head}"}
        else:
            err = getattr(r.result, "error", None)
            inner = getattr(err, "error", None)            # ErrorResponse(type='error', error=ErrorObject(type, message)) — 실제 원인은 안쪽
            etype = getattr(inner, "type", None) or getattr(err, "type", "") or ""
            emsg = getattr(inner, "message", None) or getattr(err, "message", "") or ""
            docs[key] = {"_error": f"batch {rtype}: {etype} {str(emsg)[:120]}".strip()}
    return "ended", docs, counts
