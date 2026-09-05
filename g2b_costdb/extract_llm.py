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
    "분리발주_언급공종": "string", "근거문구": "object(항목명→원문 인용 30자 이내)", "신뢰도": "high|medium|low",
}
_NUM_FIELDS = ["연면적_m2", "건축면적_m2", "대지면적_m2", "지하층수", "지상층수", "공사기간_일", "추정가격_원", "기초금액_원",
               "부가가치세_원", "도급자관급액_원", "관급자관급액_원", "총공사금액_원", "설계금액_원"]
_STR_FIELDS = ["공사명", "공사위치", "용도", "구조", "공사기간_원문", "분리발주_언급공종"]


def output_schema() -> Dict:
    """구조화 출력용 JSON 스키마 (모든 필드 필수, 없으면 null). 근거문구는 [{항목, 원문}] 배열로 받아 dict로 변환."""
    props: Dict[str, Dict] = {}
    for k in _STR_FIELDS:
        props[k] = {"type": ["string", "null"]}
    for k in _NUM_FIELDS:
        props[k] = {"type": ["number", "null"]}
    props["사업유형"] = {"type": "string", "enum": ["신축", "증축", "리모델링", "증축·리모델링", "유지보수", "기타"]}
    props["공종"] = {"type": "string", "enum": ["건축", "전기", "정보통신", "소방", "조경", "기계설비", "토목", "기타"]}
    props["신뢰도"] = {"type": "string", "enum": ["high", "medium", "low"]}
    props["근거문구"] = {"type": "array", "items": {"type": "object", "properties": {"항목": {"type": "string"}, "원문": {"type": "string"}},
                                                "required": ["항목", "원문"], "additionalProperties": False}}
    return {"type": "object", "properties": props, "required": list(props.keys()), "additionalProperties": False}


SYSTEM = (
    "당신은 한국 공공건축 공사 입찰공고문에서 공사개요와 공사금액 구성을 정확히 추출하는 분석가입니다. "
    "반드시 JSON 객체 하나만 출력합니다(마크다운 코드블록·설명 금지). 문서에 없는 값은 null로 두고 추정하지 않습니다. "
    "금액은 원 단위 정수로 변환합니다(예: '1,234백만원'→1234000000, '12.3억원'→1230000000, '천원' 단위 표는 ×1000). "
    "면적은 ㎡ 단위 숫자로 변환하고(평→×3.3058), 층수는 정수로 분리합니다. "
    "기초금액은 통상 추정가격+부가가치세이며, 관급자재(관급자관급액·도급자관급액)는 별도 항목으로 추출합니다. "
    "리모델링·개보수 공사는 '사업유형'을 리모델링으로 하고, 연면적_m2에는 공사 대상 연면적(부분 리모델링이면 대상 부분 면적)을 넣고 "
    "근거문구에 '대상면적'임을 표시합니다. 증축이 함께 있으면 '증축·리모델링'으로 표기합니다. "
    "각 값의 근거가 된 원문 문구를 '근거문구'에 항목명별로 짧게 인용합니다(근거문구는 [{\"항목\": …, \"원문\": …}] 배열). "
    "메타데이터(공고명·수요기관)는 문서 식별용이며, 금액·면적은 반드시 문서 본문에 표기된 값만 추출하고 없으면 null 로 둡니다."
)

_KEYS = ["공사개요", "공사 개요", "규모", "연면적", "건축면적", "층수", "지하", "지상", "구조", "공사기간", "준공",
         "추정가격", "기초금액", "예정가격", "관급", "부가가치세", "도급", "총공사", "설계금액", "공사금액", "사업비"]

# 모델별 1M 토큰당 요금(USD): 입력, 출력 — 견적용(정확한 값은 docs.claude.com 요금표 참조)
PRICE_USD_PER_M = {
    "claude-sonnet-5": (2.0, 10.0), "claude-opus-5": (5.0, 25.0), "claude-opus-4-8": (5.0, 25.0),
    "claude-opus-4-7": (5.0, 25.0), "claude-sonnet-4-6": (3.0, 15.0), "claude-haiku-4-5": (1.0, 5.0),
}


def focus_text(text: str, max_chars: int, window: int = 1200) -> str:
    """앞부분 + 키워드 주변 창을 우선 결합하여 길이 제한 내로 축약."""
    if len(text) <= max_chars:
        return text
    head = text[: int(max_chars * 0.35)]
    spans = []
    for k in _KEYS:
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


def estimate_cost(texts: Dict[str, str], keys: List[str], cfg_llm: Dict, usd_krw: float = 1400.0) -> Dict:
    max_chars = int(cfg_llm.get("max_input_chars", 60000))
    n, chars = 0, 0
    for k in keys:
        t = texts.get(k)
        if t:
            n += 1
            chars += min(len(t), max_chars)
    in_tok = sum(estimate_tokens(min(len(texts[k]), max_chars)) for k in keys if texts.get(k))
    out_tok = n * 1200
    model = cfg_llm.get("model", "claude-sonnet-5")
    pin, pout = PRICE_USD_PER_M.get(model, (5.0, 25.0))
    usd = in_tok / 1e6 * pin + out_tok / 1e6 * pout
    return {"docs": n, "chars": chars, "input_tokens": in_tok, "output_tokens": out_tok, "model": model,
            "usd": round(usd, 2), "krw": int(usd * usd_krw)}


def extract_with_claude(text: str, api_hint: Dict, cfg_llm: Dict) -> Dict:
    """공고문 텍스트 → 구조화 dict. api_hint: {'공고번호','공고명','수요기관'} (문서 식별용 힌트. 금액은 넣지 않는다)."""
    import anthropic  # type: ignore

    client = anthropic.Anthropic(api_key=os.environ.get(cfg_llm.get("api_key_env", "ANTHROPIC_API_KEY")) or None)
    hint = {k: jsonable(v) for k, v in (api_hint or {}).items()}
    schema = json.dumps(FIELDS, ensure_ascii=False, indent=1)
    user = (
        f"[공고 메타데이터(API)]\n{json.dumps(hint, ensure_ascii=False)}\n\n"
        f"[출력 스키마]\n{schema}\n\n"
        f"[공고문 텍스트]\n{focus_text(text, int(cfg_llm.get('max_input_chars', 60000)))}"
    )
    kwargs = dict(model=cfg_llm.get("model", "claude-sonnet-5"), max_tokens=int(cfg_llm.get("max_tokens", 8000)),
                  system=SYSTEM, messages=[{"role": "user", "content": user}])
    structured = bool(cfg_llm.get("structured_output", True))

    def _create(**kw):
        nonlocal structured
        try:
            if structured:
                return client.messages.create(output_config={"format": {"type": "json_schema", "schema": output_schema()}}, **kw)
            return client.messages.create(**kw)
        except anthropic.BadRequestError as e:
            if not structured:
                raise
            log.warning("구조화 출력 미지원/거부(%s) → 일반 응답 파싱으로 폴백", str(e)[:120])
            structured = False
            return client.messages.create(**kw)

    resp = _create(**kwargs)
    if getattr(resp, "stop_reason", "") == "max_tokens":     # 사고 토큰까지 포함해 잘린 경우 1회 2배로 재시도
        kwargs["max_tokens"] = kwargs["max_tokens"] * 2
        resp = _create(**kwargs)
        if getattr(resp, "stop_reason", "") == "max_tokens":
            raise ValueError("응답이 max_tokens 에서 잘림 — config.llm.max_tokens 를 늘리세요")
    if getattr(resp, "stop_reason", "") == "refusal":
        raise ValueError("모델이 응답을 거부함(refusal)")
    out = "".join(getattr(b, "text", "") for b in resp.content if getattr(b, "type", "") == "text")
    data = normalize_doc(_parse_json(out))
    data["_model"] = cfg_llm.get("model")
    data["_usage_in"] = getattr(resp.usage, "input_tokens", None)
    data["_usage_out"] = getattr(resp.usage, "output_tokens", None)
    return data


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
        logs.append({"공고번호": bid_no, "항목": "기초금액≈추정가격×1.1", "API값": round(exp), "문서값": b, "차이": round(b - exp),
                     "차이율": round(r, 5), "판정": "정상" if abs(r) <= 0.01 else "경고",
                     "비고": "" if abs(r) <= 0.01 else "부가세 비율 불일치(면세·관급 포함 여부 확인)"})
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
    tot = _num(doc.get("총공사금액_원"))
    if tot and b:
        comp = b + (_num(doc.get("도급자관급액_원")) or 0) + (_num(doc.get("관급자관급액_원")) or 0)
        r = (tot - comp) / comp if comp else 0
        logs.append({"공고번호": bid_no, "항목": "총공사금액=기초금액+관급", "API값": comp, "문서값": tot, "차이": tot - comp,
                     "차이율": round(r, 5), "판정": "정상" if abs(r) <= 0.01 else "경고", "비고": "금액 구성 정합성"})
    return logs
