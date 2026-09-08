"""나라장터 입찰공고정보서비스 Open API 클라이언트.

- 오퍼레이션 호출 / 페이지네이션(totalCount) / 월 단위 기간 분할
- 공공데이터포털 오류 응답(JSON 헤더 또는 XML cmmMsgHeader) 해석
  · 키 미등록(30)·기간만료(31)·IP 미등록(32)·서비스 접근거부(20) 등 → 재시도 없이 한국어 안내와 함께 중단
  · 일일 트래픽 초과(22) → DailyBudgetExceeded (다음 날 재개)
  · 일시 오류(01/02/04/05/99, HTTP 5xx, 네트워크) → 지수 백오프 재시도
  · 데이터 없음(03) → 정상(빈 결과)
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
from urllib.parse import unquote

import requests

log = logging.getLogger(__name__)


class DailyBudgetExceeded(RuntimeError):
    """일일 호출 예산 소진(자체 예산 또는 포털 트래픽 제한 22) → 다음 날 재개(체크포인트는 캐시에 남아 있음)."""


class ApiError(ValueError):
    """재시도해도 해결되지 않는 API 오류(키·권한·파라미터). code: 포털 오류코드 문자열."""

    def __init__(self, code: str, msg: str, op: str = "", params: Optional[Dict[str, Any]] = None):
        self.code, self.msg, self.op, self.params = code, msg, op, params or {}
        super().__init__(f"API 오류 {code}: {msg} — {explain_code(code)} (op={op}, params={self.params})")


# 공공데이터포털 공통 오류코드 (returnReasonCode / resultCode)
_CODE_EXPLAIN = {
    "00": "정상",
    "01": "포털 애플리케이션 오류(일시적) — 잠시 후 재시도",
    "02": "포털 DB 오류(일시적) — 잠시 후 재시도",
    "03": "데이터 없음(정상)",
    "04": "HTTP 오류(일시적) — 잠시 후 재시도",
    "05": "서비스 연결 실패/타임아웃(일시적) — 잠시 후 재시도",
    "10": "잘못된 요청 파라미터 — 오퍼레이션명·파라미터명(inqryDiv 등)을 활용가이드와 대조",
    "11": "필수 요청 파라미터 누락 — 활용가이드의 필수 파라미터 확인",
    "12": "해당 오픈API 서비스가 없거나 폐기됨 — base_url/오퍼레이션명 확인",
    "20": "서비스 접근 거부 — 공공데이터포털에서 이 서비스의 활용신청이 승인되었는지 확인",
    "21": "일시적으로 사용할 수 없는 서비스키",
    "22": "일일 트래픽 초과 — 내일 같은 명령으로 재개(완료된 월은 자동 건너뜀)",
    "30": "등록되지 않은 서비스키 — 환경변수 G2B_SERVICE_KEY 값이 포털의 '일반 인증키(Decoding)'인지 확인(공백·따옴표 없이)",
    "31": "활용기간 만료 — 포털 마이페이지에서 활용기간 연장",
    "32": "등록되지 않은 IP — 포털 활용신청 시 등록한 IP와 다른 곳에서 호출 중",
    "33": "서명되지 않은 호출",
    "99": "알 수 없는 오류(일시적) — 잠시 후 재시도",
}
RETRYABLE_CODES = {"01", "02", "04", "05", "99"}
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
        return None
    code = re.search(r"<returnReasonCode>\s*(\d+)\s*</returnReasonCode>", t)
    msg = re.search(r"<returnAuthMsg>(.*?)</returnAuthMsg>", t, re.S) or re.search(r"<errMsg>(.*?)</errMsg>", t, re.S)
    return {"code": code.group(1).zfill(2) if code else "99", "msg": (msg.group(1).strip() if msg else t[:120])}


parse_portal_error_xml = parse_portal_error  # 하위 호환


@dataclass
class ApiConfig:
    base_url: str
    service_key: str
    num_of_rows: int = 999
    timeout_sec: int = 60
    max_retries: int = 5
    sleep_between_calls_sec: float = 0.3
    daily_call_budget: int = 950


class G2BClient:
    def __init__(self, cfg: ApiConfig, cache_db: str):
        self.cfg = cfg
        # 포털의 'Encoding' 키(%2B, %3D 포함)를 넣은 경우 requests가 다시 인코딩하여 30(미등록 키)이 나므로 Decoding 키로 변환
        if "%" in (cfg.service_key or ""):
            self.cfg.service_key = unquote(cfg.service_key)
            log.info("서비스키가 URL 인코딩 형태(%%)로 보여 Decoding 키로 변환하여 사용합니다.")
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
    def _interpret(r: requests.Response, op: str, p: Dict[str, str]) -> Dict[str, Any]:
        """HTTP 응답 → body(dict). 재시도 대상은 RuntimeError, 비재시도는 ApiError, 트래픽 초과는 DailyBudgetExceeded."""
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
                G2BClient._raise_for_code(err["code"], err["msg"], op, p)
                raise RuntimeError(f"API {err['code']}: {err['msg']}")
            raise RuntimeError(f"JSON 아님: {text[:200]}")
        if not isinstance(data, dict):
            raise RuntimeError(f"예상 밖 응답: {str(data)[:200]}")
        if "response" not in data and "cmmMsgHeader" in data:  # JSON 형태의 포털 오류
            h = data["cmmMsgHeader"] or {}
            G2BClient._raise_for_code(str(h.get("returnReasonCode", "99")).zfill(2), str(h.get("returnAuthMsg") or h.get("errMsg")), op, p)
        resp = data.get("response") or {}
        header = resp.get("header") or {}
        code = str(header.get("resultCode", "")).strip().zfill(2) if header.get("resultCode") not in (None, "") else "00"
        if code != "00":
            if code in NODATA_CODES:
                return {"items": [], "totalCount": 0}
            G2BClient._raise_for_code(code, str(header.get("resultMsg", "")), op, p)
            raise RuntimeError(f"API {code}: {header.get('resultMsg', '')}")
        return resp.get("body") or {}

    @staticmethod
    def _raise_for_code(code: str, msg: str, op: str, p: Dict[str, str]) -> None:
        if code in QUOTA_CODES:
            raise DailyBudgetExceeded(f"포털 일일 트래픽 초과(22): {msg} → 내일 같은 명령으로 재개")
        if code in RETRYABLE_CODES or code in NODATA_CODES:
            return
        raise ApiError(code, msg, op, p)

    # ── 단건 호출 ────────────────────────────────────────────────
    def call(self, op: str, params: Dict[str, Any], use_cache: bool = True, cache_salt: Optional[str] = None) -> Dict[str, Any]:
        """오퍼레이션 1회 호출 → response.body(dict). 캐시 우선.
        cache_salt: 캐시 키에 섞는 문자열(예: 오늘 날짜) — 진행 중인 달처럼 내용이 바뀌는 조회에 사용."""
        p = {k: str(v) for k, v in params.items() if v is not None}
        key = hashlib.sha1((op + json.dumps(p, sort_keys=True, ensure_ascii=False) + (cache_salt or "")).encode()).hexdigest()
        if use_cache:
            row = self.conn.execute("SELECT body FROM cache WHERE k=?", (key,)).fetchone()
            if row:
                return json.loads(row[0])

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
                body = self._interpret(r, op, p)
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
            except ApiError:
                raise
            except (requests.RequestException, RuntimeError, ValueError) as e:
                last_err = e
                log.warning("호출 실패(%d/%d) op=%s params=%s err=%s", attempt, self.cfg.max_retries, op, p, e)
                if attempt < self.cfg.max_retries:
                    time.sleep(delay)
                    delay = min(delay * 2, 60)
        raise RuntimeError(f"재시도 소진 op={op} params={p}: {last_err}")

    # ── 페이지네이션 ─────────────────────────────────────────────
    @staticmethod
    def _items(body: Dict[str, Any]) -> List[Dict[str, Any]]:
        items = (body or {}).get("items")
        if not items or isinstance(items, str):
            return []
        if isinstance(items, dict):
            items = items.get("item", [])
        if isinstance(items, dict):
            items = [items]
        return [it for it in items if isinstance(it, dict)]

    def iter_all(self, op: str, params: Dict[str, Any], use_cache: bool = True,
                 cache_salt: Optional[str] = None) -> Iterator[Dict[str, Any]]:
        page = 1
        seen = 0
        while True:
            body = self.call(op, {**params, "pageNo": page, "numOfRows": self.cfg.num_of_rows},
                             use_cache=use_cache, cache_salt=cache_salt)
            items = self._items(body)
            try:
                total = int(float(body.get("totalCount")))
            except (TypeError, ValueError):
                total = None  # totalCount 없음 → 짧은 페이지가 나올 때까지 계속
                if page == 1 and items:
                    log.warning("%s: 응답에 totalCount 가 없어 페이지가 짧아질 때까지 계속 조회합니다", op)
            if page == 1 and items and len(items) < self.cfg.num_of_rows and (total or 0) > len(items):
                log.info("%s: 서버가 페이지당 %d건만 반환(요청 %d) → 페이지 수가 늘어남", op, len(items), self.cfg.num_of_rows)
            for it in items:
                yield it
            seen += len(items)
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

    def iter_period(self, op: str, start_ym: str, end_ym: str, extra: Optional[Dict[str, Any]] = None,
                    inqry_div: str = "1") -> Iterator[Dict[str, Any]]:
        """기간 조회 오퍼레이션을 월 단위로 분할 호출하여 전체 항목을 순회."""
        for bgn, end in self.month_ranges(start_ym, end_ym):
            params = {"inqryDiv": inqry_div, "inqryBgnDt": bgn, "inqryEndDt": end}
            if extra:
                params.update(extra)
            n = 0
            for it in self.iter_all(op, params):
                it["_inqry_month"] = bgn[:6]
                n += 1
                yield it
            log.info("%s %s: %d건", op, bgn[:6], n)
