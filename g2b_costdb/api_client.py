"""나라장터 입찰공고정보서비스 Open API 클라이언트.

- 오퍼레이션 호출 / 페이지네이션(totalCount) / 월 단위 기간 분할
- 공공데이터포털 오류 응답(JSON 헤더 또는 XML cmmMsgHeader) 해석
  · 키 미등록(30)·기간만료(31)·IP 미등록(32)·서비스 접근거부(20) 등 → 재시도 없이 한국어 안내와 함께 중단
  · 일일 트래픽 초과(22) → DailyBudgetExceeded (다음 날 재개)
  · 일시 오류(01/02/04/05/99, HTTP 5xx, 네트워크) → 지수 백오프 재시도
  · 데이터 없음(03) → 정상(빈 결과)
  · 조달청 서비스 자체의 요청 오류(06 날짜 형식·07 조회 범위 초과·08 필수값 누락)는 HTTP 200 에 다른 모양으로 온다:
    {"nkoneps.com.response.ResponseError": {"header": {"resultCode": "07", "resultMsg": "입력범위값 초과 에러"}}} → 재시도 없이 ApiError
  · 알려진 형식이 아닌 응답(response.header/body 없음, body 에 totalCount·items 없음 등)은 '자료 없음'이 아니라
    UnexpectedResponse(재시도 없음)로 중단. 그런 본문은 캐시에 넣지 않고, 이미 캐시에 든 빈 본문은 없는 것으로 취급
- 조회 기간 제한: 2026-11-10 부터 입찰공고정보서비스는 한 번에 달력 날짜 7일까지만 조회된다(그전에는 1개월) → split_window 로 나눠 조회
- SQLite 캐시(오퍼레이션+파라미터 해시)로 재호출 방지, 일일 호출 예산 관리
"""
from __future__ import annotations

import calendar
import hashlib
import json
import logging
import os
import re
import sqlite3
import time
from dataclasses import dataclass
from datetime import date, datetime
from typing import Any, Dict, Iterator, List, Optional
from urllib.parse import quote, quote_plus, unquote

import requests

log = logging.getLogger(__name__)


class DailyBudgetExceeded(RuntimeError):
    """일일 호출 예산 소진(자체 예산 또는 포털 트래픽 제한 22) → 다음 날 재개(체크포인트는 캐시에 남아 있음)."""


class ApiError(ValueError):
    """재시도해도 해결되지 않는 API 오류(키·권한·파라미터). code: 포털 오류코드 문자열."""

    def __init__(self, code: str, msg: str, op: str = "", params: Optional[Dict[str, Any]] = None):
        self.code, self.msg, self.op, self.params = code, msg, op, params or {}
        super().__init__(f"API 오류 {code}: {msg} — {explain_code(code)} (op={op}, params={self.params})")


class UnexpectedResponse(ApiError):
    """HTTP 200 인데 알려진 응답 형식이 아니거나(예: {"nkoneps.com.response.ResponseError": …} 인데 코드도 없음, {"response": {}}),
    정상 코드인데 body 에 totalCount·items 가 모두 없는 응답. '자료 없음'으로 넘기지 않는다(재시도 없음, 캐시 안 함). code='XX'."""

    def __init__(self, why: str, raw: str = "", op: str = "", params: Optional[Dict[str, Any]] = None):
        self.why, self.raw = why, (raw or "")[:300]
        ApiError.__init__(self, "XX", f"{why} | 응답 앞부분: {self.raw}", op, params)


class IncompleteResult(RuntimeError):
    """페이지를 끝까지 돌았는데 받은 건수가 서버 totalCount 보다 적다 → 그 달은 완료로 표시하면 안 된다."""


# 공공데이터포털 공통 오류코드 (returnReasonCode / resultCode)
_CODE_EXPLAIN = {
    "00": "정상",
    "01": "포털 애플리케이션 오류(일시적) — 잠시 후 재시도",
    "02": "포털 DB 오류(일시적) — 잠시 후 재시도",
    "03": "데이터 없음(정상)",
    "04": "HTTP 오류(일시적) — 잠시 후 재시도",
    "05": "서비스 연결 실패/타임아웃(일시적) — 잠시 후 재시도",
    # 06~08 은 조달청 서비스 자체 오류코드(HTTP 200, 최상위 키가 'response' 가 아닌 오류 응답으로 온다)
    "06": "날짜 형식 오류 — inqryBgnDt/inqryEndDt 형식(YYYYMMDDHHMM) 확인",
    "07": "입력범위값 초과 — 한 번에 조회할 수 있는 기간을 넘음(2026-11-10 부터 달력 날짜 7일까지. config.yaml api.max_window_days 확인)",
    "08": "필수값 입력 오류 — 이 오퍼레이션의 필수 파라미터(조회구분별 필수값) 확인",
    "10": "잘못된 요청 파라미터 — 오퍼레이션명·파라미터명(inqryDiv 등)을 활용가이드와 대조",
    "11": "필수 요청 파라미터 누락 — 활용가이드의 필수 파라미터 확인",
    "12": "해당 오픈API 서비스가 없거나 폐기됨 — base_url/오퍼레이션명 확인",
    "20": "서비스 접근 거부 — 공공데이터포털에서 이 서비스의 활용신청이 승인되었는지 확인",
    "21": "일시적으로 사용할 수 없는 서비스키",
    "22": "일일 트래픽 초과 — 내일 같은 명령으로 재개(완료된 월은 자동 건너뜀)",
    "23": "초당 호출 허용량 초과(일시적) — 잠시 후 재시도",
    "29": "차단된 IP — 공공데이터포털 활용지원센터에 문의",
    "30": "등록되지 않은 서비스키 — 환경변수 G2B_SERVICE_KEY 값이 포털의 '일반 인증키(Decoding)'인지 확인(공백·따옴표 없이)",
    "31": "활용기간 만료 — 포털 마이페이지에서 활용기간 연장",
    "32": "등록되지 않은 IP — 포털 활용신청 시 등록한 IP와 다른 곳에서 호출 중",
    "33": "서명되지 않은 호출",
    "99": "알 수 없는 오류(일시적) — 잠시 후 재시도",
    "XX": "응답 형식을 해석하지 못함 — '자료 없음'이 아니라 오류입니다(아래 '응답 앞부분' 확인)",
}
RETRYABLE_CODES = {"01", "02", "04", "05", "23", "99"}
QUOTA_CODES = {"22"}
NODATA_CODES = {"03"}


def explain_code(code: str) -> str:
    return _CODE_EXPLAIN.get(str(code), "포털 활용가이드의 오류코드표 참조")


def parse_portal_error(text: str) -> Optional[Dict[str, str]]:
    """공공데이터포털 게이트웨이 오류 응답 → {'code','msg'} (오류 응답이 아니면 None).
    형식 두 가지를 모두 해석한다:
      XML : <OpenAPI_ServiceResponse><cmmMsgHeader><returnReasonCode>30</returnReasonCode>…
      JSON: {"cmmMsgHeader": {"returnReasonCode": "30", "returnAuthMsg": "…"}}  (type=json 요청 시, HTTP 403 과 함께 오기도 함)
            {"response": {"header": {"resultCode": "30", "resultMsg": "…"}}}"""
    t = (text or "").strip()
    if not t:
        return None
    if t.startswith("{"):
        try:
            data = json.loads(t)
        except ValueError:
            data = None
        if isinstance(data, dict):
            h = data.get("cmmMsgHeader")
            if not isinstance(h, dict) and isinstance(data.get("OpenAPI_ServiceResponse"), dict):
                h = data["OpenAPI_ServiceResponse"].get("cmmMsgHeader")
            if isinstance(h, dict):
                return {"code": str(h.get("returnReasonCode") or "99").strip().zfill(2),
                        "msg": str(h.get("returnAuthMsg") or h.get("errMsg") or "")[:200]}
            hdr = (data.get("response") or {}).get("header") if isinstance(data.get("response"), dict) else None
            if isinstance(hdr, dict) and hdr.get("resultCode") not in (None, "") and str(hdr["resultCode"]).strip().zfill(2) != "00":
                return {"code": str(hdr["resultCode"]).strip().zfill(2), "msg": str(hdr.get("resultMsg", ""))[:200]}
            return None
    if "cmmMsgHeader" not in t and "OpenAPI_ServiceResponse" not in t:
        # 제공기관 형식의 XML 오류: <response><header><resultCode>07</resultCode><resultMsg>…</resultMsg></header></response>
        rc = re.search(r"<resultCode>\s*(\d+)\s*</resultCode>", t)
        if t.startswith("<") and rc and rc.group(1).zfill(2) != "00":
            rm = re.search(r"<resultMsg>(.*?)</resultMsg>", t, re.S)
            return {"code": rc.group(1).zfill(2), "msg": (rm.group(1).strip() if rm else "")[:200]}
        return None
    code = re.search(r"<returnReasonCode>\s*(\d+)\s*</returnReasonCode>", t)
    msg = re.search(r"<returnAuthMsg>(.*?)</returnAuthMsg>", t, re.S) or re.search(r"<errMsg>(.*?)</errMsg>", t, re.S)
    return {"code": code.group(1).zfill(2) if code else "99", "msg": (msg.group(1).strip() if msg else t[:120])}


parse_portal_error_xml = parse_portal_error  # 하위 호환


def _norm_code(v) -> Optional[str]:
    """resultCode/returnReasonCode → 두 자리 문자열(없으면 None). 0, "0", "00", 7, " 07 " 모두 처리."""
    if v is None or isinstance(v, bool):
        return None
    s = str(v).strip()
    if not s:
        return None
    if re.fullmatch(r"\d+\.0+", s):
        s = s.split(".")[0]
    return s.zfill(2)


def body_items(body) -> List[Dict[str, Any]]:
    items = body.get("items") if isinstance(body, dict) else None
    if not items or isinstance(items, str):
        return []
    if isinstance(items, dict):
        items = items.get("item", [])
    if isinstance(items, dict):
        items = [items]
    return [it for it in items if isinstance(it, dict)] if isinstance(items, list) else []


def body_total(body) -> Optional[int]:
    try:
        return int(float(body.get("totalCount")))
    except (TypeError, ValueError, AttributeError, OverflowError):     # OverflowError: "Infinity"/1e999
        return None


def validate_body(body, page_no: int = 1, rows: int = 0) -> Optional[str]:
    """response.body 가 '믿을 수 있는 결과'인지 확인 → 문제 설명(없으면 None). 새 응답과 캐시에서 꺼낸 본문에 똑같이 적용한다.
    정상: 항목이 1건 이상 / totalCount=0(자료 없음) / 마지막 쪽을 지난 빈 쪽.  비정상: dict 아님, totalCount·items 모두 없음, totalCount>0 인데 0건."""
    if not isinstance(body, dict):
        return "body 가 객체(dict)가 아님"
    if body_items(body):
        return None
    total = body_total(body)
    if total is None:
        return "body 에 totalCount 도 items 도 없음"
    if total == 0:
        return None
    if total < 0:
        return f"totalCount 가 음수({total})"
    if rows and page_no and (page_no - 1) * rows >= total:
        return None
    return f"TRUNCATED: totalCount={total} 인데 이 쪽(pageNo={page_no})의 items 가 0건"


@dataclass
class ApiConfig:
    base_url: str
    service_key: str
    num_of_rows: int = 999
    timeout_sec: int = 60
    max_retries: int = 5
    sleep_between_calls_sec: float = 0.3
    daily_call_budget: int = 950
    max_window_days: int = 7        # 한 번에 조회할 수 있는 기간(달력 날짜 수). 2026-11-10 부터 입찰공고정보서비스는 7일


class G2BClient:
    def __init__(self, cfg: ApiConfig, cache_db: str):
        self.cfg = cfg
        # 포털의 'Encoding' 키(%2B, %3D 포함)를 넣은 경우 requests가 다시 인코딩하여 30(미등록 키)이 나므로 Decoding 키로 변환
        if "%" in (cfg.service_key or ""):
            self.cfg.service_key = unquote(cfg.service_key)
            log.info("서비스키가 URL 인코딩 형태(%s 포함)로 보여 Decoding 키로 변환하여 사용합니다.", "%")
        os.makedirs(os.path.dirname(cache_db) or ".", exist_ok=True)
        self.conn = sqlite3.connect(cache_db)
        self.conn.execute(
            "CREATE TABLE IF NOT EXISTS cache (k TEXT PRIMARY KEY, op TEXT, params TEXT, body TEXT, fetched_at TEXT)"
        )
        self.conn.execute("CREATE TABLE IF NOT EXISTS calls (day TEXT PRIMARY KEY, n INTEGER)")
        self.conn.commit()
        self.session = requests.Session()

    # ── 호출 예산 ────────────────────────────────────────────────
    def calls_today(self) -> int:
        row = self.conn.execute("SELECT n FROM calls WHERE day=?", (date.today().isoformat(),)).fetchone()
        return row[0] if row else 0

    _calls_today = calls_today  # 하위 호환

    def _bump_calls(self) -> None:
        d = date.today().isoformat()
        self.conn.execute(
            "INSERT INTO calls(day,n) VALUES(?,1) ON CONFLICT(day) DO UPDATE SET n=n+1", (d,)
        )
        self.conn.commit()

    def _mark_budget_exhausted(self) -> None:
        """포털이 22(트래픽 초과)를 돌려주면 오늘은 더 호출하지 않도록 예산을 소진 처리."""
        d = date.today().isoformat()
        self.conn.execute("INSERT INTO calls(day,n) VALUES(?,?) ON CONFLICT(day) DO UPDATE SET n=MAX(n,excluded.n)",
                          (d, self.cfg.daily_call_budget))
        self.conn.commit()

    # ── 응답 해석 ────────────────────────────────────────────────
    @staticmethod
    def _by_code(code: str, msg: str, op: str, p: Dict[str, str]) -> Dict[str, Any]:
        """오류코드별 처리: 03 → 빈 결과, 22 → DailyBudgetExceeded, 일시 오류 → RuntimeError(재시도), 그 밖 → ApiError(재시도 없음)."""
        if code in NODATA_CODES:
            return {"items": [], "totalCount": 0}
        G2BClient._raise_for_code(code, msg, op, p)
        raise RuntimeError(f"API {code}: {msg}")

    @staticmethod
    def _interpret(r: requests.Response, op: str, p: Dict[str, str], page_no: int = 1, rows: int = 0) -> Dict[str, Any]:
        """HTTP 응답 → body(dict). 재시도 대상은 RuntimeError, 비재시도는 ApiError(해석 불가는 UnexpectedResponse), 트래픽 초과는 DailyBudgetExceeded.
        정상(00)으로 인정하는 조건: response(dict).header.resultCode 가 00(또는 header 없음) 이고 body 가 validate_body 를 통과."""
        text = r.text or ""
        if r.status_code != 200:
            err = parse_portal_error(text)
            if err and err["code"] not in RETRYABLE_CODES:
                G2BClient._raise_for_code(err["code"], err["msg"], op, p)   # 키·권한 오류는 즉시 중단, 22는 예산 소진
            raise RuntimeError(f"HTTP {r.status_code}: {text[:200]}")
        try:
            data = r.json()
        except ValueError:
            err = parse_portal_error(text)
            if err:
                return G2BClient._by_code(err["code"], err["msg"], op, p)
            raise RuntimeError(f"JSON 아님: {text[:200]}")
        if not isinstance(data, dict):
            raise RuntimeError(f"예상 밖 응답: {str(data)[:200]}")
        # 1) 포털 게이트웨이 오류(JSON 형태): {"cmmMsgHeader": …} 또는 {"OpenAPI_ServiceResponse": {"cmmMsgHeader": …}}
        if "response" not in data:
            h = data.get("cmmMsgHeader")
            if h is None and isinstance(data.get("OpenAPI_ServiceResponse"), dict):
                h = data["OpenAPI_ServiceResponse"].get("cmmMsgHeader")
            if h is not None:
                if not isinstance(h, dict):
                    raise UnexpectedResponse("cmmMsgHeader 형식이 예상과 다름", text, op, p)
                return G2BClient._by_code(_norm_code(h.get("returnReasonCode")) or "99", str(h.get("returnAuthMsg") or h.get("errMsg") or ""), op, p)
        # 2) 표준 형식 {"response": {"header": …, "body": …}} 또는 다른 이름으로 감싼 오류({"nkoneps.com.response.ResponseError": {"header": …}})
        resp, where = data.get("response"), "response"
        if not isinstance(resp, dict):
            resp = None
            for k, v in data.items():
                if isinstance(v, dict) and isinstance(v.get("header"), dict) and "resultCode" in v["header"]:
                    resp, where = v, str(k)
                    break
            if resp is None:
                raise UnexpectedResponse('응답에 "response" 항목이 없고 오류코드도 찾지 못함', text, op, p)
        header, body = resp.get("header"), resp.get("body")
        if header is not None and not isinstance(header, dict):
            raise UnexpectedResponse("response.header 형식이 예상과 다름", text, op, p)
        code = _norm_code((header or {}).get("resultCode"))
        if code is not None and code != "00":
            return G2BClient._by_code(code, str((header or {}).get("resultMsg", "")), op, p)
        if where != "response":
            raise UnexpectedResponse(f'"{where}" 형식의 응답인데 오류코드가 없음', text, op, p)
        problem = validate_body(body, page_no, rows)
        if problem and problem.startswith("TRUNCATED"):
            raise RuntimeError(f"잘린 응답({problem}) — 재시도")                    # 일시적일 수 있으므로 재시도, 끝내 안 되면 그 달은 실패
        if problem:
            raise UnexpectedResponse(f"결과코드 {code or '없음'} 인데 {problem}", text, op, p)
        return body

    @staticmethod
    def _raise_for_code(code: str, msg: str, op: str, p: Dict[str, str]) -> None:
        if code in QUOTA_CODES:
            raise DailyBudgetExceeded(f"포털 일일 트래픽 초과(22): {msg} → 내일 같은 명령으로 재개")
        if code in RETRYABLE_CODES or code in NODATA_CODES:
            return
        raise ApiError(code, msg, op, p)

    @staticmethod
    def _cache_key(op: str, p: Dict[str, str], cache_salt: Optional[str] = None) -> str:
        return hashlib.sha1((op + json.dumps(p, sort_keys=True, ensure_ascii=False) + (cache_salt or "")).encode()).hexdigest()

    def forget(self, op: str, params: Dict[str, Any], cache_salt: Optional[str] = None) -> None:
        """이 조회 1건의 캐시 행을 지운다. 쪽 하나하나는 정상인데 합쳐 보니 전량이 아닌 조회(IncompleteResult)의 쪽들이 캐시에 남으면
        다음 실행 때 서버에 묻지도 않고 같은 실패를 되풀이하므로, 그런 조회는 다음에 서버에서 다시 받게 한다."""
        p = {k: str(v) for k, v in params.items() if v is not None}
        self.conn.execute("DELETE FROM cache WHERE k=?", (self._cache_key(op, p, cache_salt),))
        self.conn.commit()

    # ── 단건 호출 ────────────────────────────────────────────────
    def call(self, op: str, params: Dict[str, Any], use_cache: bool = True, cache_salt: Optional[str] = None) -> Dict[str, Any]:
        """오퍼레이션 1회 호출 → response.body(dict). 캐시 우선.
        cache_salt: 캐시 키에 섞는 문자열(예: 오늘 날짜) — 진행 중인 달처럼 내용이 바뀌는 조회에 사용."""
        p = {k: str(v) for k, v in params.items() if v is not None}
        key = self._cache_key(op, p, cache_salt)
        try:
            page_no, rows = int(p.get("pageNo") or 1), int(p.get("numOfRows") or self.cfg.num_of_rows)
        except ValueError:
            page_no, rows = 1, 0
        if use_cache:
            row = self.conn.execute("SELECT body FROM cache WHERE k=?", (key,)).fetchone()
            if row:
                try:
                    cached = json.loads(row[0])
                except ValueError:
                    cached = None
                problem = validate_body(cached, page_no, rows)
                if problem is None:
                    return cached
                # 예전 버전이 '자료 없음'으로 저장해 둔 빈 본문({} 등)은 캐시에 없는 것으로 보고 다시 호출한다
                log.warning("캐시에 든 응답이 비정상(%s) → 버리고 다시 호출합니다: op=%s params=%s", problem, op, p)
                self.conn.execute("DELETE FROM cache WHERE k=?", (key,))
                self.conn.commit()

        if self.calls_today() >= self.cfg.daily_call_budget:
            raise DailyBudgetExceeded(f"오늘 호출 {self.calls_today()}회: 예산 {self.cfg.daily_call_budget} 도달")

        url = f"{self.cfg.base_url}/{op}"
        q = dict(p)
        q.update({"serviceKey": self.cfg.service_key, "type": "json"})
        q.setdefault("numOfRows", str(self.cfg.num_of_rows))
        q.setdefault("pageNo", "1")

        delay = 2.0
        last_err: Optional[Exception] = None
        for attempt in range(1, self.cfg.max_retries + 1):
            try:
                self._bump_calls()
                r = self.session.get(url, params=q, timeout=self.cfg.timeout_sec)
                time.sleep(self.cfg.sleep_between_calls_sec)
                body = self._interpret(r, op, p, page_no, rows)
                self.conn.execute(
                    "INSERT OR REPLACE INTO cache VALUES(?,?,?,?,?)",
                    (key, op, json.dumps(p, ensure_ascii=False), json.dumps(body, ensure_ascii=False),
                     datetime.now().isoformat(timespec="seconds")),
                )
                self.conn.commit()
                return body
            except DailyBudgetExceeded:
                self._mark_budget_exhausted()
                raise
            except ApiError as e:
                e.args = (self._scrub(e.args[0]),) if e.args else e.args      # 응답 앞부분에 인증키가 섞여 있어도 화면·로그에 남지 않게
                e.msg = self._scrub(e.msg)
                raise
            except (requests.RequestException, RuntimeError, ValueError) as e:
                last_err = e
                log.warning("호출 실패(%d/%d) op=%s params=%s err=%s", attempt, self.cfg.max_retries, op, p, self._scrub(e))
                if attempt < self.cfg.max_retries:
                    time.sleep(delay)
                    delay = min(delay * 2, 60)
        raise RuntimeError(f"재시도 소진 op={op} params={p}: {self._scrub(last_err)}")

    def diagnose(self, op: str, params: Dict[str, Any]) -> Dict[str, Any]:
        """진단용 1회 호출(캐시·재시도 없음, 호출 예산 1회 사용): 응답을 _interpret 로 읽어 본 결과와 겉모양을 요약한다.
        반환 {'http', 'status', 'code', 'msg', 'keys', 'total', 'n', 'snippet'(, 'body')} — status: 'ok'(항목 있음) / 'nodata'(서버가 0건이라고 답함) /
        'error'(서버가 요청을 거절: code·msg) / 'unreadable'(읽을 수 없는 응답) / 'temporary'(일시 오류) / 'quota'(일일 한도) / 'network'(연결 실패).
        snippet 은 본문 앞부분(인증키 가림)."""
        p = {k: str(v) for k, v in params.items() if v is not None}
        out = {"http": None, "status": "network", "code": None, "msg": "", "keys": [], "total": None, "n": 0, "snippet": ""}
        if self.calls_today() >= self.cfg.daily_call_budget:
            out.update(status="quota", msg=f"오늘 호출 예산({self.cfg.daily_call_budget}회) 도달")
            return out
        q = dict(p)
        q.update({"serviceKey": self.cfg.service_key, "type": "json"})
        q.setdefault("numOfRows", str(self.cfg.num_of_rows))
        q.setdefault("pageNo", "1")
        try:
            self._bump_calls()
            r = self.session.get(f"{self.cfg.base_url}/{op}", params=q, timeout=self.cfg.timeout_sec)
            time.sleep(self.cfg.sleep_between_calls_sec)
        except requests.RequestException as e:
            out["msg"] = self._scrub(e)[:300]
            return out
        text = r.text or ""
        out.update(http=r.status_code, snippet=self._scrub(re.sub(r"\s+", " ", text))[:300])
        try:
            data = r.json()
            out["keys"] = list(data.keys())[:8] if isinstance(data, dict) else []
        except ValueError:
            pass
        try:
            page_no, rows = int(q["pageNo"]), int(q["numOfRows"])
        except ValueError:
            page_no, rows = 1, 0
        try:
            body = self._interpret(r, op, p, page_no, rows)
            n = len(self._items(body))
            out.update(status="ok" if n else "nodata", code="00", total=body.get("totalCount"), n=n, body=body)
        except DailyBudgetExceeded as e:
            self._mark_budget_exhausted()               # 포털이 일일 한도(22)라고 답했으면 오늘은 더 부르지 않는다
            out.update(status="quota", code="22", msg=self._scrub(e)[:200])
        except UnexpectedResponse as e:
            out.update(status="unreadable", msg=self._scrub(e.why)[:200])
        except ApiError as e:
            out.update(status="error", code=e.code, msg=self._scrub(e.msg)[:200])
        except RuntimeError as e:
            out.update(status="temporary", msg=self._scrub(e)[:200])
        return out

    def _scrub(self, text) -> str:
        """오류 문구에서 인증키를 가린다. 연결 오류 메시지에는 요청 URL(serviceKey=… 포함)이 그대로 들어 있어, 화면·로그에 키가 찍히지 않게 한다."""
        s = str(text)
        k = self.cfg.service_key or ""
        for v in {k, quote(k, safe=""), quote_plus(k)}:
            if v:
                s = s.replace(v, "***")
        return re.sub(r"(serviceKey=)[^&\s'\")]+", r"\1***", s, flags=re.I)

    def purge_cache(self, op: str, contains: str, before: Optional[str] = None) -> int:
        """op 의 응답 캐시 중 요청 파라미터에 contains 가 든 행을 지운다. 용역 전수 수집처럼 응답은 크고 보존하는 행은 적을 때,
        한 달을 다 받아 저장한 뒤 그 달의 캐시를 비워 cache.sqlite 가 불어나지 않게 한다(달 중간에 끊기면 캐시가 남아 이어받기 가능).
        before('YYYY-MM-DD'): 그 날짜 이전에 받은 행만 지운다 — 진행 중인 달을 날마다 다시 받을 때 전날까지의 사본을 지우고 오늘 받은 페이지는 남긴다."""
        if before:
            cur = self.conn.execute("DELETE FROM cache WHERE op=? AND params LIKE ? AND fetched_at < ?", (op, f"%{contains}%", before))
        else:
            cur = self.conn.execute("DELETE FROM cache WHERE op=? AND params LIKE ?", (op, f"%{contains}%"))
        self.conn.commit()
        return int(cur.rowcount or 0)

    # ── 페이지네이션 ─────────────────────────────────────────────
    @staticmethod
    def _items(body: Dict[str, Any]) -> List[Dict[str, Any]]:
        return body_items(body)

    def iter_all(self, op: str, params: Dict[str, Any], use_cache: bool = True,
                 cache_salt: Optional[str] = None, require_total: bool = False,
                 stats: Optional[Dict[str, Any]] = None) -> Iterator[Dict[str, Any]]:
        """전 페이지 순회. 받은 건수가 서버 totalCount 보다 적은 채 끝나면 IncompleteResult(조용히 끝내지 않는다).
        require_total: totalCount 가 없는 응답을 오류로 본다(전량 수집용). stats: {'total','seen','pages'} 를 채워 준다."""
        page = 1
        seen = 0
        total = None

        def drop_pages(upto: int) -> None:          # 전량을 받지 못한 조회: 이 조회의 쪽들을 캐시에서 지워 다음 실행 때 서버에서 다시 받는다
            if use_cache:
                for pg in range(1, upto + 1):
                    self.forget(op, {**params, "pageNo": pg, "numOfRows": self.cfg.num_of_rows}, cache_salt)
        while True:
            body = self.call(op, {**params, "pageNo": page, "numOfRows": self.cfg.num_of_rows},
                             use_cache=use_cache, cache_salt=cache_salt)
            items = self._items(body)
            page_total = body_total(body)
            if page_total is not None:
                total = page_total if total is None else max(total, page_total)     # 쪽마다 건수가 달라지면 큰 쪽을 믿는다(줄어든 값으로 일찍 끝내지 않는다)
            elif items:
                if require_total:
                    drop_pages(page)
                    raise UnexpectedResponse("응답에 totalCount 가 없어 전량 수신 여부를 확인할 수 없음", json.dumps(body, ensure_ascii=False)[:200], op, params)
                if page == 1:
                    log.warning("%s: 응답에 totalCount 가 없어 페이지가 짧아질 때까지 계속 조회합니다(전량 여부 확인 불가)", op)
            if page == 1 and items and len(items) < self.cfg.num_of_rows and (total or 0) > len(items):
                log.info("%s: 서버가 페이지당 %d건만 반환(요청 %d) → 페이지 수가 늘어남", op, len(items), self.cfg.num_of_rows)
            for it in items:
                yield it
            seen += len(items)
            if stats is not None:
                stats.update(total=total, seen=seen, pages=page)
            if not items:
                break
            if total is not None:
                if seen >= total:
                    break
            elif len(items) < self.cfg.num_of_rows:
                break
            if page > 5000:  # 안전장치
                log.warning("%s: 페이지 5000 초과 — 중단", op)
                break
            page += 1
        if total is not None and seen < total:
            drop_pages(page)
            raise IncompleteResult(f"{op}: 서버 totalCount={total:,} 인데 {seen:,}건만 받고 끝남(pageNo={page}) → 이 조회는 완료로 처리하지 않습니다")

    # ── 월 단위 기간 분할 ───────────────────────────────────────
    @staticmethod
    def month_ranges(start_ym: str, end_ym: str) -> List[tuple]:
        """'2019-01','2026-09' → [(inqryBgnDt, inqryEndDt), ...] (YYYYMMDDHHMM)"""
        y, m = map(int, str(start_ym).split("-"))
        ey, em = map(int, str(end_ym).split("-"))
        out = []
        while (y, m) <= (ey, em):
            last = calendar.monthrange(y, m)[1]
            out.append((f"{y:04d}{m:02d}010000", f"{y:04d}{m:02d}{last:02d}2359"))
            m += 1
            if m > 12:
                y, m = y + 1, 1
        return out

    @staticmethod
    def split_window(bgn: str, end: str, max_days: int, overlap: bool = True) -> List[tuple]:
        """한 달 안의 조회 기간 (YYYYMMDD0000, YYYYMMDD2359) 을 달력 날짜 max_days 일 이하의 구간으로 나눈다.
        구간의 끝은 다음 구간의 시작과 같은 분(00:00)이다: 7일이면 1일 00:00~7일 00:00, 7일~13일, 13일~19일, 19일~25일, 25일 00:00~말일 23:59.
        서버가 끝 시각의 그 1분을 포함하는지(23:59:30 같은 초 단위 시각)를 알 수 없으므로, 구간을 23:59 로 끊지 않고 경계의 1분을 두 구간에 걸치게 해
        어느 쪽이든 빠지는 공고가 없게 한다(경계 1분의 공고는 두 번 올 수 있다 — 받는 쪽에서 같은 행을 한 번만 남긴다. 달의 마지막 1분은 예전과 같다).
        max_days 가 1 이면 겹칠 수 없어 하루씩(00:00~23:59), 0 이하이거나 기간이 그보다 짧으면 그대로 한 구간.
        overlap=False: 겹치지 않는 구간(1~7일 23:59, 8~14일 …) — 표본만 보는 probe 용(수집에는 쓰지 않는다)."""
        d0, d1 = int(bgn[6:8]), int(end[6:8])
        if not max_days or max_days <= 0 or bgn[:6] != end[:6] or d1 - d0 + 1 <= max_days:
            return [(bgn, end)]
        ym, out, d = bgn[:6], [], d0
        if max_days == 1 or not overlap:
            return [(f"{ym}{k:02d}{bgn[8:] if k == d0 else '0000'}", f"{ym}{min(k + max_days - 1, d1):02d}{end[8:] if k + max_days - 1 >= d1 else '2359'}")
                    for k in range(d0, d1 + 1, max_days)]
        while True:
            start = f"{ym}{d:02d}{bgn[8:] if d == d0 else '0000'}"
            if d + max_days - 1 >= d1:                   # 남은 날이 한 구간에 들어간다 → 달의 끝까지
                out.append((start, end))
                return out
            d += max_days - 1
            out.append((start, f"{ym}{d:02d}0000"))

    def iter_period(self, op: str, start_ym: str, end_ym: str, extra: Optional[Dict[str, Any]] = None,
                    inqry_div: str = "1") -> Iterator[Dict[str, Any]]:
        """기간 조회 오퍼레이션을 월 단위(그 안에서 max_window_days 구간)로 분할 호출하여 전체 항목을 순회."""
        for bgn, end in self.month_ranges(start_ym, end_ym):
            n = 0
            for wb, we in self.split_window(bgn, end, self.cfg.max_window_days):
                params = {"inqryDiv": inqry_div, "inqryBgnDt": wb, "inqryEndDt": we}
                if extra:
                    params.update(extra)
                for it in self.iter_all(op, params):
                    it["_inqry_month"] = bgn[:6]
                    n += 1
                    yield it
            log.info("%s %s: %d건", op, bgn[:6], n)
