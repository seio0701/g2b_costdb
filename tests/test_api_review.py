"""수집기·응답 해석·받은 원문 점검(check-raw)·probe 에 대한 독립 검토(2026-11-17)에서 나온 확인 항목.
네트워크·인증키를 쓰지 않는다(요청은 모두 가짜 응답이거나 127.0.0.1 의 tests/mock_g2b.MockG2B).

    python -m tests.test_api_review        # tests.test_regressions 도 끝에서 이 묶음을 함께 실행한다

- check_*  : 테스트가 지키지 못하던 규칙(한 줄 변이로 확인) — 응답 해석의 경계, 캐시 지우기, 가장 큰 totalCount, 진단 호출, 인증키 가림,
             진행 중인 달의 수집, 실패 처리, .prev, 캐시 정리 범위, probe 의 판단, 가짜 서버의 기간 규칙, check-raw 의 판정
- fix_*    : 2차 검토에서 고친 결함 — probe 판단(일시 오류·일일 한도·빈 달), 인증키 가림의 빈틈, RATIO_OUTLIER, 면허제한 조회 중단, 낙찰 probe
- review2  : 구간 안의 같은 행 보존, 서버가 0건이라고 답한 빈 달, 구간으로 받은 달의 서버 건수 표시, check-raw 가 캐시를 읽는 동안의 수집
"""
from __future__ import annotations

import contextlib
import datetime
import io
import json
import logging
import os
import re
import sqlite3
import subprocess
import sys
import tempfile
import time
from unittest import mock

REPO = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, REPO)

import requests  # noqa: E402
import yaml  # noqa: E402
from g2b_costdb import collect, selfcheck  # noqa: E402
from g2b_costdb.api_client import (ApiConfig, ApiError, DailyBudgetExceeded, G2BClient, IncompleteResult, UnexpectedResponse,  # noqa: E402
                                   body_total, validate_body)

J = lambda o: json.dumps(o, ensure_ascii=False)  # noqa: E731
OPN = "getBidPblancListInfoCnstwk"
OPS = "getBidPblancListInfoServc"


class _Resp:
    def __init__(self, text, status=200):
        self.text, self.status_code, self.headers = text, status, {}

    def json(self):
        return json.loads(self.text)


def OKB(items, total, **extra):
    return J({"response": {"header": {"resultCode": "00", "resultMsg": "정상"}, "body": dict({"items": items, "totalCount": total}, **extra)}})


NODATA = J({"response": {"header": {"resultCode": "03", "resultMsg": "NODATA_ERROR"}, "body": {}}})
E08 = J({"nkoneps.com.response.ResponseError": {"header": {"resultCode": "08", "resultMsg": "필수값 입력 에러"}}})
XML = lambda code, msg="X": (f"<OpenAPI_ServiceResponse><cmmMsgHeader><errMsg>SERVICE ERROR</errMsg><returnAuthMsg>{msg}</returnAuthMsg>"  # noqa: E731
                             f"<returnReasonCode>{code}</returnReasonCode></cmmMsgHeader></OpenAPI_ServiceResponse>")
_n = [0]


def _client(tmp, retries=3, key="k", **kw):
    _n[0] += 1
    return G2BClient(ApiConfig(base_url="http://x", service_key=key, max_retries=retries, sleep_between_calls_sec=0, **kw), os.path.join(tmp, f"na{_n[0]}.sqlite"))


def make_cfg(tmp, name, start="2024-01", end="2024-01", **api):
    cfg = yaml.safe_load(open(os.path.join(REPO, "config.yaml"), encoding="utf-8"))
    d = os.path.join(tmp, name)
    cfg["api"].update(base_url="http://x/BidPublicInfoService", sleep_between_calls_sec=0, max_retries=2, num_of_rows=5, use_awards=False)
    cfg["api"].update(api)
    cfg["period"] = {"start": start, "end": end}
    cfg["servc"]["period_end"] = end
    cfg["paths"] = {"data_dir": d, "raw_dir": os.path.join(d, "raw"), "cache_db": os.path.join(d, "cache.sqlite"), "text_dir": os.path.join(d, "text"),
                    "files_dir": os.path.join(d, "files"), "out_dir": os.path.join(d, "out"), "excel_name": "x.xlsx"}
    os.makedirs(cfg["paths"]["raw_dir"], exist_ok=True)
    return cfg


class Srv:
    """Fake server for requests.Session.get: rows filtered by the query range (minute resolution, end minute included)."""

    def __init__(self, rows, hook=None):
        self.rows, self.hook, self.log = rows, hook, []

    def __call__(self, url, params=None, timeout=None, **kw):
        op = url.rstrip("/").split("/")[-1]
        self.log.append((op, dict(params)))
        if self.hook:
            r = self.hook(op, params)
            if r is not None:
                if isinstance(r, BaseException):
                    raise r
                return r
        b, e, pg, n = params.get("inqryBgnDt", ""), params.get("inqryEndDt", ""), int(params.get("pageNo", 1)), int(params.get("numOfRows", 999))
        hit = [r for r in self.rows if b <= "".join(ch for ch in r["bidNtceDt"] if ch.isdigit())[:12] <= e]
        if not hit:
            return _Resp(NODATA)
        return _Resp(OKB(hit[(pg - 1) * n: pg * n] or "", len(hit)))

    def starts(self):
        return [p.get("inqryBgnDt") for _, p in self.log]


def row(i, dt, **kw):
    return dict({"bidNtceNo": f"R{i:04d}", "bidNtceOrd": "000", "bidNtceNm": "가상 공사", "bidNtceDt": dt, "presmptPrce": "1"}, **kw)


def _lines(path):
    return sum(1 for _ in open(path, encoding="utf-8"))


@contextlib.contextmanager
def patched(srv, today=None):
    with contextlib.ExitStack() as st:
        st.enter_context(mock.patch.object(requests.Session, "get", side_effect=srv))
        st.enter_context(mock.patch("time.sleep"))
        if today is not None:
            d = st.enter_context(mock.patch.object(collect, "date"))
            d.today.return_value = today
        st.enter_context(contextlib.redirect_stderr(io.StringIO()))
        yield


# ───────────────────────────────────────── api_client ─────────────────────────────────────────
def check_body_total_and_validate_body_edges(tmp):
    """kills A01, A04, A05"""
    assert body_total({"totalCount": "12.0"}) == 12 and body_total({"totalCount": 3.0}) == 3 and body_total({"totalCount": " 7 "}) == 7      # A01
    # 쪽이 정확히 totalCount 에서 시작하면(3쪽 = 11번째 행부터, 전체 10건) 마지막 쪽을 지난 빈 쪽 = 정상                                          # A04
    assert validate_body({"items": [], "totalCount": 10}, page_no=3, rows=5) is None
    # 2쪽(6~10번째 행)이 비었는데 전체가 9건이면 잘린 쪽이다(마지막 쪽을 지난 것이 아님)                                                           # A05
    assert str(validate_body({"items": [], "totalCount": 9}, page_no=2, rows=5)).startswith("TRUNCATED")
    assert str(validate_body({"items": [], "totalCount": 5}, page_no=1, rows=5)).startswith("TRUNCATED")


def check_http_4xx_with_ok_body_is_not_data(tmp):
    """kills A11 — HTTP 403/404 must never be read as data even if the body looks like a normal response"""
    for status in (403, 404, 429):
        c = _client(tmp, retries=2)
        with mock.patch.object(c.session, "get", return_value=_Resp(OKB([{"a": 1}], 1), status)) as g, mock.patch("time.sleep"):
            try:
                c.call("op", {"a": "1"})
                raise AssertionError(f"HTTP {status} 본문을 정상 자료로 받아들임")
            except RuntimeError as e:
                assert "재시도 소진" in str(e) and g.call_count == 2, (status, str(e), g.call_count)
        assert c.conn.execute("SELECT COUNT(*) FROM cache").fetchone()[0] == 0


def check_gateway_header_without_code_and_retry_codes(tmp):
    """kills A16, A22 (A23 is killed by the existing test)"""
    c = _client(tmp, retries=3)
    with mock.patch.object(c.session, "get", return_value=_Resp(J({"cmmMsgHeader": {"errMsg": "SERVICE ERROR"}}))) as g, mock.patch("time.sleep"):
        try:
            body = c.call("op", {"a": "1"})
            raise AssertionError(f"코드 없는 cmmMsgHeader 를 자료 없음으로 읽음: {body}")                                                  # A16
        except RuntimeError as e:
            assert "재시도 소진" in str(e) and g.call_count == 3
    assert c.conn.execute("SELECT COUNT(*) FROM cache").fetchone()[0] == 0
    for code in ("01", "02", "04", "05", "23", "99"):                                                                                        # A22: 모든 일시 오류 코드가 재시도된다
        c = _client(tmp, retries=3)
        with mock.patch.object(c.session, "get", return_value=_Resp(XML(code))) as g, mock.patch("time.sleep"):
            try:
                c.call("op", {"a": code})
                raise AssertionError("오류 기대")
            except ApiError as e:
                raise AssertionError(f"일시 오류 {code} 가 재시도 없이 ApiError 로 끝남: {e}")
            except RuntimeError as e:
                assert "재시도 소진" in str(e) and g.call_count == 3, (code, g.call_count)


def check_forget_honours_cache_salt(tmp):
    """kills A31 — the pages of an incomplete query of a re-collected month (salted cache key) must be dropped too"""
    c = _client(tmp); c.cfg.num_of_rows = 5
    fake = lambda url, params=None, timeout=None: _Resp(OKB([{"i": 1}] if params["pageNo"] == "1" else "", 3))  # noqa: E731
    with mock.patch.object(c.session, "get", side_effect=fake) as g, mock.patch("time.sleep"):
        for _ in range(2):
            try:
                list(c.iter_all("op", {"a": "1"}, cache_salt="2026-11-17"))
                raise AssertionError("IncompleteResult 기대")
            except IncompleteResult:
                pass
        assert g.call_count == 4, g.call_count                       # 두 번 다 서버에 다시 묻는다(캐시에서 되풀이하지 않는다)
    assert c.conn.execute("SELECT COUNT(*) FROM cache").fetchone()[0] == 0
    # forget 는 salt 가 같은 행만 지운다
    with mock.patch.object(c.session, "get", return_value=_Resp(OKB([{"i": 1}], 1))), mock.patch("time.sleep"):
        c.call("op", {"q": "1"}); c.call("op", {"q": "1"}, cache_salt="s")
    c.forget("op", {"q": 1}, "s")
    assert c.conn.execute("SELECT COUNT(*) FROM cache").fetchone()[0] == 1
    c.forget("op", {"q": 1})
    assert c.conn.execute("SELECT COUNT(*) FROM cache").fetchone()[0] == 0


def check_iter_all_trusts_largest_total_and_reports_stats(tmp):
    """kills A34, A35, A40, A41"""
    c = _client(tmp); c.cfg.num_of_rows = 5
    pages = {"1": ([{"i": i} for i in range(5)], 12), "2": ([{"i": i} for i in range(5, 10)], 10), "3": ([{"i": 10}, {"i": 11}], 12)}   # 2쪽에서만 건수가 줄어 보인다
    st = {}
    with mock.patch.object(c.session, "get", side_effect=lambda url, params=None, timeout=None: _Resp(OKB(*pages[params["pageNo"]]))) as g, mock.patch("time.sleep"):
        got = list(c.iter_all("op", {"a": "1"}, require_total=True, stats=st))
    assert [x["i"] for x in got] == list(range(12)) and g.call_count == 3, (len(got), g.call_count)                                       # A34, A35
    assert st == {"total": 12, "seen": 12, "pages": 3}, st                                                                               # A40
    # 서버가 totalCount 보다 많이 내려 줘도 total 은 서버가 말한 값, seen 은 받은 값(둘을 섞지 않는다)                                       # A41
    c = _client(tmp); c.cfg.num_of_rows = 10
    st = {}
    with mock.patch.object(c.session, "get", return_value=_Resp(OKB([{"i": i} for i in range(6)], 5))), mock.patch("time.sleep"):
        assert len(list(c.iter_all("op", {"a": "2"}, stats=st))) == 6
    assert (st["total"], st["seen"]) == (5, 6), st


def check_diagnose_budget_and_quota(tmp):
    """kills A50, A51, A52, A56"""
    c = _client(tmp, daily_call_budget=3)
    with mock.patch.object(c.session, "get", return_value=_Resp(OKB([{"a": 1}], 1))) as g, mock.patch("time.sleep"):
        assert c.diagnose("op", {"a": "1"})["status"] == "ok" and c.calls_today() == 1                                                   # A51: 진단 호출도 예산에 센다
        c.diagnose("op", {"a": "2"}); c.diagnose("op", {"a": "3"})
        d = c.diagnose("op", {"a": "4"})                                                                                                 # A50: 예산을 다 쓰면 부르지 않는다
        assert d["status"] == "quota" and d["http"] is None and g.call_count == 3 and c.calls_today() == 3, (d, g.call_count)
    c = _client(tmp, daily_call_budget=50)
    with mock.patch.object(c.session, "get", return_value=_Resp(XML("22", "LIMITED_NUMBER_OF_SERVICE_REQUESTS_EXCEEDS_ERROR"))) as g, mock.patch("time.sleep"):
        d = c.diagnose("op", {"a": "1"})
        assert (d["status"], d["code"]) == ("quota", "22"), d                                                                           # A56
        assert c.calls_today() >= 50, c.calls_today()                                                                                   # A52: 22 를 받으면 오늘 예산 소진
        assert c.diagnose("op", {"a": "2"})["status"] == "quota" and g.call_count == 1, g.call_count
        try:
            c.call("op", {"a": "3"})
            raise AssertionError("DailyBudgetExceeded 기대")
        except DailyBudgetExceeded:
            assert g.call_count == 1


def check_scrub_each_defence_on_its_own(tmp):
    """kills A55, A57, A58"""
    key = "SECRET+KEY/abc=="               # made-up value
    c = _client(tmp, key=key)
    assert "SECRET" not in c._scrub("invalid key SECRET%2BKEY%2Fabc%3D%3D given")                                                      # A58: 'serviceKey=' 없이 인코딩된 키
    assert c._scrub("GET /op?inqryDiv=1&serviceKey=SOMETHING-ELSE%2b1&type=json") == "GET /op?inqryDiv=1&serviceKey=***&type=json"      # A57: 값이 달라도 serviceKey= 뒤는 가린다
    err = requests.exceptions.ConnectionError("HTTPSConnectionPool(host='apis.data.go.kr', port=443): Max retries exceeded with url: "
                                              "/1230000/ad/BidPublicInfoService/op?inqryDiv=1&serviceKey=SECRET%2BKEY%2Fabc%3D%3D&type=json (Caused by X) key='SECRET+KEY/abc=='")
    with mock.patch.object(c.session, "get", side_effect=err), mock.patch("time.sleep"):
        d = c.diagnose("op", {"a": "1"})                                                                                                 # A55
    assert d["status"] == "network" and "SECRET" not in J(d) and "***" in d["msg"], d


# ───────────────────────────────────────── collect: _collect_monthly ─────────────────────────────────────────
def check_current_month_collection(tmp):
    """kills C01, C22, C24, C28, C29 — the month that contains today: every window that has started is fetched (also the one that starts today),
    no done mark, months.json says current, and a stale done mark does not make it skip."""
    rows = [row(i, f"2024-01-{d:02d} 10:00:00") for i, d in enumerate((2, 8, 13, 13, 20, 26))]
    srv = Srv(rows)
    cfg = make_cfg(tmp, "current")
    raw = cfg["paths"]["raw_dir"]
    today = datetime.date(2024, 1, 13)                              # 구간: 1~7, 7~13, 13~19(오늘 시작), 19~25·25~31(아직 오지 않음)
    with patched(srv, today):
        collect.collect_notices(cfg)
    assert "202401130000" in srv.starts(), srv.starts()                                                                                 # C01, C29
    assert not [s for s in srv.starts() if s > "202401139999"], srv.starts()
    assert _lines(os.path.join(raw, "notices_202401.jsonl")) == 4
    assert not os.path.exists(os.path.join(raw, "notices_done.txt")) or "202401" not in open(os.path.join(raw, "notices_done.txt")).read()   # C24
    meta = json.load(open(os.path.join(raw, "notices_months.json"), encoding="utf-8"))["202401"]
    assert meta["current"] is True and meta["windows"] == 3 and meta["saved"] == 4, meta                                                # C22, C01
    open(os.path.join(raw, "notices_done.txt"), "w", encoding="utf-8").write("202401")      # 손으로 넣은 완료 표시가 있어도 진행 중인 달은 다시 받는다
    n0 = len(srv.log)
    with patched(srv, datetime.date(2024, 1, 14)):
        collect.collect_notices(cfg)
    assert len(srv.log) > n0, "진행 중인 달을 건너뜀"                                                                                    # C28


def check_collect_requires_total_count(tmp):
    """kills C04 — a window whose pages carry no totalCount cannot be verified → the month must not be marked done"""
    rows = [row(i, "2024-01-03 10:00:00") for i in range(3)]

    def hook(op, p):
        if p["inqryBgnDt"].startswith("20240101"):
            return _Resp(J({"response": {"header": {"resultCode": "00"}, "body": {"items": rows}}}))        # totalCount 없음
    srv = Srv(rows, hook)
    cfg = make_cfg(tmp, "nototal")
    with patched(srv):
        collect.collect_notices(cfg)
    raw = cfg["paths"]["raw_dir"]
    assert not os.path.exists(os.path.join(raw, "notices_done.txt")) and not os.path.exists(os.path.join(raw, "notices_202401.jsonl")), os.listdir(raw)


def check_key_error_stops_collection_at_once(tmp):
    """kills C07 — code 30 (unregistered key) ends the run with [중단] after one call, it is not '3 failed months'"""
    srv = Srv([], lambda op, p: _Resp(XML("30", "SERVICE_KEY_IS_NOT_REGISTERED_ERROR")))
    cfg = make_cfg(tmp, "key30", end="2024-06")
    with patched(srv):
        try:
            collect.collect_notices(cfg)
            raise AssertionError("SystemExit 기대")
        except SystemExit as e:
            assert "[중단]" in str(e) and "30" in str(e), str(e)
    assert len(srv.log) == 1, len(srv.log)


def check_failure_counter_resets_after_a_good_month(tmp):
    """kills C10 — only CONSECUTIVE failed months stop the run"""
    rows = [row(i, f"2024-{m:02d}-03 10:00:00") for i, m in enumerate(range(1, 8))]
    bad = {"202401", "202403", "202405"}
    srv = Srv(rows, lambda op, p: _Resp("{}") if p["inqryBgnDt"][:6] in bad else None)
    cfg = make_cfg(tmp, "nonconsec", end="2024-07")
    with patched(srv):
        collect.collect_notices(cfg)
    done = open(os.path.join(cfg["paths"]["raw_dir"], "notices_done.txt"), encoding="utf-8").read().split()
    assert done == ["202402", "202404", "202406", "202407"], done


def check_prev_only_when_fewer(tmp):
    """kills C16 — re-collecting a month with the same number of rows must not leave a .prev copy"""
    rows = [row(i, "2024-01-03 10:00:00") for i in range(7)]
    srv = Srv(rows)
    cfg = make_cfg(tmp, "sameprev")
    raw = cfg["paths"]["raw_dir"]
    with patched(srv):
        collect.collect_notices(cfg)
    os.remove(os.path.join(raw, "notices_done.txt"))
    with patched(srv, datetime.date(2030, 1, 2)):
        collect.collect_notices(cfg)
    assert sorted(fn for fn in os.listdir(raw) if fn.startswith("notices_2024")) == ["notices_202401.jsonl"], os.listdir(raw)
    assert _lines(os.path.join(raw, "notices_202401.jsonl")) == 7


def check_purge_pattern_and_same_day_resume(tmp):
    """kills C25, C26 — a filtered collection (용역) empties the cache of a finished month; a re-collected month that fails midway keeps
    the pages fetched today so the same-day rerun only asks for the missing windows"""
    rows = [row(i, f"2024-01-{d:02d} 10:00:00", bidNtceNm="건설폐기물 처리용역") for i, d in enumerate((2, 8, 14, 20, 26))]
    srv = Srv(rows)
    cfg = make_cfg(tmp, "purge")
    cfg["api"]["ops"]["servc_list"] = OPS
    raw = cfg["paths"]["raw_dir"]
    count = lambda: sqlite3.connect(cfg["paths"]["cache_db"]).execute("SELECT COUNT(*) FROM cache WHERE op=?", (OPS,)).fetchone()[0]  # noqa: E731
    with patched(srv):
        collect.collect_servc(cfg, "2024-01", "2024-01")
    assert open(os.path.join(raw, "servc_done.txt")).read().split() == ["202401"] and count() == 0, count()                             # C25
    # 다시 받는 달(완료 표시를 지움): 네 번째 구간에서 실패 → 같은 날 다시 실행하면 앞의 세 구간은 캐시에서
    os.remove(os.path.join(raw, "servc_done.txt"))
    srv.hook = lambda op, p: _Resp("{}") if p["inqryBgnDt"].startswith("20240119") else None
    day = datetime.date(2030, 1, 5)
    with patched(srv, day), mock.patch("g2b_costdb.api_client.datetime") as dt:
        dt.now.return_value = datetime.datetime(2030, 1, 5, 9, 0, 0)
        collect.collect_servc(cfg, "2024-01", "2024-01")
        assert count() == 3, count()
        srv.hook, n0 = None, len(srv.log)
        collect.collect_servc(cfg, "2024-01", "2024-01")
    assert len(srv.log) - n0 == 2, (len(srv.log) - n0, srv.starts()[n0:])                                                               # C26: 실패한 구간과 그 뒤 구간만
    assert open(os.path.join(raw, "servc_done.txt")).read().split() == ["202401"] and _lines(os.path.join(raw, "servc_202401.jsonl")) == 5


def check_license_lookup_continues_after_a_success(tmp):
    """kills C33 — the early stop applies only while not a single lookup has succeeded"""
    cfg = make_cfg(tmp, "lic2")
    ok = OKB([{"lcnsLmtNm": "건축공사업"}], 1)
    n = [0]

    def fake(url, params=None, timeout=None, **kw):
        n[0] += 1
        return _Resp(ok if n[0] == 1 else E08)
    with mock.patch.object(requests.Session, "get", side_effect=fake), mock.patch("time.sleep"), contextlib.redirect_stderr(io.StringIO()):
        df = collect.collect_license_limits(cfg, [f"R{i}" for i in range(12)])
    assert n[0] == 12 and len(df) == 1, (n[0], len(df))


# ───────────────────────────────────────── collect: probes ─────────────────────────────────────────
def _probe_out(fn, cfg, ym, srv, today=None):
    buf = io.StringIO()
    with patched(srv, today), contextlib.redirect_stdout(buf):
        fn(cfg, ym)
    return buf.getvalue()


def check_probe_moves_on_from_an_empty_first_window(tmp):
    """kills C37, C38, C44"""
    rows = [row(i, "2024-03-10 10:00:00", bidNtceNm="건설폐기물 처리용역", ntceInsttNm="가상군", dminsttNm="가상군") for i in range(3)]
    srv = Srv(rows, lambda op, p: _Resp(OKB([{"lcnsLmtNm": "건축공사업"}], 1)) if "LicenseLimit" in op else None)
    cfg = make_cfg(tmp, "probe1", start="2024-03", end="2024-03", num_of_rows=999)
    out = _probe_out(collect.probe, cfg, "2024-03", srv)
    assert f"=== {OPN} (20240308~20240314) totalCount=3 표본 3건 ===" in out, out                                                       # C37
    srv = Srv(rows)
    cfg = make_cfg(tmp, "probe2", start="2024-03", end="2024-03", num_of_rows=999)
    out = _probe_out(collect.probe_servc, cfg, "2024-03", srv)
    assert f"=== {OPS} (20240308~20240314, 7일) totalCount=3 표본 3건 ===" in out and "표본을 받지 못했습니다" not in out, out             # C38
    # 평년 2월: 표본용 구간은 4개(1~7 … 22~28)지만 수집은 경계를 겹쳐 5구간으로 조회한다 → 안내·호출 수 추정은 5구간
    rows = [row(i, "2023-02-03 10:00:00", bidNtceNm="건설폐기물 처리용역") for i in range(3)]
    cfg = make_cfg(tmp, "probe3", start="2023-02", end="2023-02", num_of_rows=999)
    out = _probe_out(collect.probe_servc, cfg, "2023-02", Srv(rows))
    assert "7일 이하 5구간으로 나눠 조회합니다" in out and "한 달(5구간) 약 5회" in out, out                                               # C44


def check_probe_servc_quota_is_not_diagnosed(tmp):
    """kills C39 — with the daily budget used up the probe says so and spends nothing (no [판단] built from 'quota' lines)"""
    cfg = make_cfg(tmp, "probeq", start="2024-03", end="2024-03", daily_call_budget=4)
    c = collect.make_client(cfg)
    c._mark_budget_exhausted(); c.conn.close()
    srv = Srv([])
    out = _probe_out(collect.probe_servc, cfg, "2024-03", srv)
    assert "오늘의 호출 한도에 도달했습니다" in out and "[판단]" not in out and "①" not in out and len(srv.log) == 0, (out, len(srv.log))


def check_probe_all_variants_rejected_is_not_no_data(tmp):
    """kills C41, C42"""
    rows = [row(1, "2024-03-03 10:00:00")]
    srv = Srv(rows, lambda op, p: _Resp(E08) if "LicenseLimit" in op or "Scsbid" in op else None)
    cfg = make_cfg(tmp, "probev", start="2024-03", end="2024-03", num_of_rows=999, use_awards=True, award_base_url="http://x/ScsbidInfoService")
    out = _probe_out(collect.probe, cfg, "2024-03", srv)
    assert "[면허제한] 모든 조회구분이 오류로 거절됨" in out and "면허제한이 없을 수 있음" not in out, out                                 # C41
    assert "[낙찰정보] 모든 조회구분이 오류로 거절됨" in out and "[낙찰정보] 어떤 조회구분으로도 항목이 없음" not in out, out               # C42


# ───────────────────────────────────────── mock_g2b (the test double) ─────────────────────────────────────────
def check_mock_span_rule_and_bsis_range(tmp):
    """kills M01, M07 — the fake server must reject exactly what the real one rejects (8 calendar dates) and filter every list by range"""
    from tests import mock_g2b
    srv = mock_g2b.MockG2B()
    base = srv.start()
    try:
        s = requests.Session(); s.trust_env = False
        q = lambda op, b, e: s.get(f"{base}/BidPublicInfoService/{op}", params={"serviceKey": mock_g2b.KEY, "type": "json", "inqryDiv": "1",  # noqa: E731
                                                                                   "inqryBgnDt": b, "inqryEndDt": e, "pageNo": 1, "numOfRows": 999}, timeout=10).json()
        assert "response" in q(OPN, "202401010000", "202401072359") and "response" in q(OPN, "202401010000", "202401070000")
        assert q(OPN, "202401010000", "202401080000")["nkoneps.com.response.ResponseError"]["header"]["resultCode"] == "07"                # M01: 달력 날짜 8일
        opb = "getBidPblancListInfoCnstwkBsisAmount"
        jan = [b for b in srv.data["bsis"] if b["bidNtceDt"].startswith("2024-01")]
        got = q(opb, "202401010000", "202401072359")["response"]
        assert got["header"]["resultCode"] == "03", got["header"]                                                                        # M07: 1~7일에는 기초금액이 없다
        got = sum(len(q(opb, f"202401{a:02d}0000", f"202401{b:02d}2359")["response"]["body"].get("items", [])) for a, b in ((8, 14), (15, 21), (22, 28), (29, 31)))
        assert got == len(jan) > 0, (got, len(jan))
    finally:
        srv.stop()


# ───────────────────────────────────────── selfcheck ─────────────────────────────────────────
class Raw:
    """a data dir with raw/<prefix>YYYYMM.jsonl, done.txt, months.json and a cache.sqlite with hand-made rows (page size 5)"""

    def __init__(self, tmp, name):
        self.d = os.path.join(tmp, name)
        self.raw = os.path.join(self.d, "raw")
        os.makedirs(self.raw)
        self.con = sqlite3.connect(os.path.join(self.d, "cache.sqlite"))
        self.con.execute("CREATE TABLE cache (k TEXT PRIMARY KEY, op TEXT, params TEXT, body TEXT, fetched_at TEXT)")
        self.k = 0

    def month(self, prefix, ym, lines, done=True):
        open(os.path.join(self.raw, f"{prefix}{ym}.jsonl"), "w", encoding="utf-8").write('{"x": 1}\n' * lines)
        if done:
            with open(os.path.join(self.raw, f"{prefix}done.txt"), "a", encoding="utf-8") as f:
                f.write(ym + "\n")

    def meta(self, prefix, recs):
        json.dump(recs, open(os.path.join(self.raw, f"{prefix}months.json"), "w", encoding="utf-8"))

    def put(self, op, ym, page, body, fetched="2026-09-10T10:00:00", bgn=None, end=None, rows=5):
        import calendar
        self.k += 1
        last = calendar.monthrange(int(ym[:4]), int(ym[4:]))[1]
        p = {"inqryDiv": "1", "inqryBgnDt": bgn or ym + "010000", "inqryEndDt": end or f"{ym}{last:02d}2359", "pageNo": str(page), "numOfRows": str(rows)}
        self.con.execute("INSERT INTO cache VALUES(?,?,?,?,?)", (f"k{self.k}", op, json.dumps(p), body if isinstance(body, str) else json.dumps(body), fetched))

    def whole(self, op, ym, n, total=None, fetched="2026-09-10T10:00:00"):
        for page in range(1, (n + 4) // 5 + 1):
            self.put(op, ym, page, {"items": [{"x": 1}] * min(5, n - (page - 1) * 5), "totalCount": n if total is None else total}, fetched)

    def analyse(self, today="202611", **kw):
        self.con.commit()
        rep, bad, meta = selfcheck.analyse(self.d, page_size=5, use_cache=True, ops=dict(selfcheck.DEFAULT_OPS), today_ym=today, **kw)
        return {p: {r["ym"]: r for r in v["rows"]} for p, v in rep.items() if not p.startswith("_")}, rep, bad


def check_selfcheck_records_and_confirmed(tmp):
    """kills S01, S04, S06, S12, S28, S32"""
    x = Raw(tmp, "sc1")
    for ym, n in (("202301", 6), ("202302", 6), ("202303", 4)):
        x.month("notices_", ym, n)
    x.month("notices_", "202611", 3, done=False)                    # 오늘이 속한 달(받는 중)
    x.meta("notices_", {"202301": {"total": 6, "received": 6, "dup": 0, "saved": 6, "current": True},       # 진행 중일 때 쓴 기록 → 완료 판정에 쓰지 않는다
                        "202302": {"total": 6, "received": 6, "dup": 0, "saved": 6, "current": False},
                        "202303": {"total": 6, "received": 6, "dup": 0, "saved": 6, "current": False}})    # 6행 저장했다는데 4행
    x.month("servc_", "202301", 10)                                  # 걸러 저장하는 수집: 기록 없음, 줄 수(10)가 쪽 크기의 배수, 캐시에는 한 달 전체 조회(20건)가 온전히 남아 있음
    x.whole(selfcheck.DEFAULT_OPS["servc_"], "202301", 20)
    r, rep, _ = x.analyse()
    n, s = r["notices_"], r["servc_"]
    assert n["202301"]["confirmed"] is False and n["202301"]["server"] is None and n["202301"]["flags"] == [], n["202301"]              # S01
    assert n["202302"]["confirmed"] is True and n["202302"]["server"] == 6
    assert n["202303"]["flags"] == ["SHORT_VS_SAVED"] and n["202303"]["confirmed"] is False, n["202303"]                                # S04
    assert n["202611"]["flags"] == [] and n["202611"]["current"] is True, n["202611"]                                                   # S28: 받는 중인 달은 표시하지 않는다
    assert s["202301"]["flags"] == [] and s["202301"]["server"] is None, s["202301"]                                                    # S06, S12
    assert selfcheck._kind(True, -1, 0, 1, 5)[0] == "EMPTY_NO_TOTAL" and selfcheck._kind(True, 12, 0, 2, 5)[0] == "TRUNCATED"            # S32


def check_selfcheck_confirmed_months_get_no_heuristic_flags(tmp):
    """kills S09, S13, S14, S15, S16, S26 — the January case on the user's PC: a month far below its neighbours whose line count equals
    the server count is NOT flagged; the same shape without a server count IS flagged (and is not reopened: it is a SOFT flag)"""
    OPB = selfcheck.DEFAULT_OPS["bsis_"]
    x = Raw(tmp, "sc2")
    counts = {"202301": 4, "202302": 12, "202303": 13, "202304": 11, "202305": 12, "202306": 4, "202307": 13, "202308": 12, "202309": 11}
    for ym, n in counts.items():
        x.month("notices_", ym, n)
        if ym != "202306":
            x.whole(OPN, ym, n)                                      # 202306 만 서버 건수를 모른다
        b = {"202304": 3, "202308": 3}.get(ym, n - 1)                # 기초금액÷공고 비율: 202304·202308 만 크게 낮다
        x.month("bsis_", ym, b)
        if ym != "202308":
            x.whole(OPB, ym, b)                                      # 202304 는 서버 건수와 맞음(실제로 적은 달), 202308 은 모름
    x.put(OPN, "202302", 9, "{}", fetched="2026-09-10T11:00:00")     # 서버 건수와 맞는 달의 캐시에 남은 낡은 '읽지 못한 응답' 행
    r, rep, bad = x.analyse()
    n, b = r["notices_"], r["bsis_"]
    assert n["202301"]["confirmed"] and n["202301"]["flags"] == [], n["202301"]                                                         # S13
    assert n["202306"]["flags"] == ["LOW_VS_NEIGHBOURS"], n["202306"]                                                                   # S14
    assert n["202302"]["confirmed"] and n["202302"]["flags"] == [] and len(bad) == 1, (n["202302"], bad)                                 # S09
    assert b["202304"]["flags"] == [] and n["202304"]["flags"] == [], (b["202304"], n["202304"])                                        # S15
    assert "RATIO_OUTLIER" in b["202308"]["flags"], b["202308"]                                                                         # S16
    assert selfcheck.hard_months(rep) == {"notices_": [], "bsis_": []}
    assert selfcheck.reopen(x.raw, rep) == {}                                                                                           # S26: 눈으로 확인할 달은 완료 표시를 지우지 않는다
    assert sorted(open(os.path.join(x.raw, "notices_done.txt"), encoding="utf-8").read().split()) == sorted(counts)


def check_scan_cache_grouping(tmp):
    """kills S17, S18, S19, S20, S21"""
    x = Raw(tmp, "sc3")
    # 202301: 한 달 통째로 받은 예전 수집분(13행 = totalCount 13). 같은 조회 조건으로 probe 가 나중에 남긴 50건짜리 '읽지 못한 응답' 행은 쓰지 않는다
    x.month("notices_", "202301", 13); x.whole(OPN, "202301", 13)
    x.put(OPN, "202301", 1, "{}", fetched="2026-11-17T07:00:00", rows=50)
    # … 새 probe 가 남긴 첫 7일 구간의 행(50건 요청, 그 구간 전체는 3,148건)도 쓰지 않는다 — 쓰면 '쪽이 빠짐' 으로 잘못 표시된다
    x.put(OPN, "202301", 1, {"items": [{"x": 1}] * 5, "totalCount": 3148}, fetched="2026-11-17T07:01:00", bgn="202301010000", end="202301072359", rows=50)
    # 202302: 같은 쪽을 두 번 받음 — 예전 것은 읽지 못한 응답, 나중 것은 정상 → 나중 것을 본다
    x.month("notices_", "202302", 4)
    x.put(OPN, "202302", 1, "{}", fetched="2026-09-10T10:00:00")
    x.put(OPN, "202302", 1, {"items": [{"x": 1}] * 4, "totalCount": 4}, fetched="2026-09-11T10:00:00")
    # 202303: 7일 구간으로 받았고(경계의 1건이 두 구간에 걸려 구간별 합 13) 수집 기록은 없음 → 캐시로는 그 달의 서버 건수를 알 수 없다(표시도 확인도 없음)
    x.month("notices_", "202303", 12)
    for (b, e), k in zip(G2BClient.split_window("202303010000", "202303312359", 7), (3, 3, 3, 2, 2)):
        x.put(OPN, "202303", 1, {"items": [{"x": 1}] * k, "totalCount": k}, fetched="2026-11-20T10:00:00", bgn=b, end=e)
    # 202304: 예전에 한 달 통째로 받은 행(totalCount 20)이 남아 있고, 그 뒤 7일 구간으로 다시 받음(11행) → 나중에 받은 쪽을 본다
    x.month("notices_", "202304", 11)
    x.whole(OPN, "202304", 20, fetched="2026-09-10T10:00:00")
    for (b, e), k in zip(G2BClient.split_window("202304010000", "202304302359", 7), (3, 2, 2, 2, 2)):
        x.put(OPN, "202304", 1, {"items": [{"x": 1}] * k, "totalCount": k}, fetched="2026-11-20T10:00:00", bgn=b, end=e)
    r, rep, bad = x.analyse()
    n = r["notices_"]
    assert n["202301"]["confirmed"] and n["202301"]["flags"] == [] and n["202301"]["server"] == 13, n["202301"]                         # S17
    assert n["202302"]["confirmed"] and n["202302"]["flags"] == [], n["202302"]                                                         # S18
    assert n["202303"]["flags"] == [] and n["202303"]["confirmed"] is False and n["202303"]["server"] is None, n["202303"]              # S19, S20
    assert n["202304"]["flags"] == [] and n["202304"]["server"] is None, n["202304"]                                                    # S21
    assert selfcheck.hard_months(rep) == {"notices_": []}


# ── 2차 검토에서 고친 결함 ─────────────────────────────────────────────────────
E07 = J({"nkoneps.com.response.ResponseError": {"header": {"resultCode": "07", "resultMsg": "입력범위값 초과 에러"}}})
SV = [row(i, f"2024-03-{d:02d} 10:00:00", bidNtceNm="건설폐기물 처리용역", ntceInsttNm="가상군", dminsttNm="가상군") for i, d in enumerate((5, 7, 11, 20))]


def fix_D1_quota_during_diagnosis(tmp):
    n = [0]
    def hook(op, p):
        n[0] += 1
        return _Resp(E07) if n[0] == 1 else _Resp(XML("22", "LIMITED_NUMBER_OF_SERVICE_REQUESTS_EXCEEDS_ERROR"))
    srv = Srv(SV, hook)
    out = _probe_out(collect.probe_servc, make_cfg(tmp, "d1", start="2024-03", end="2024-03", num_of_rows=999), "2024-03", srv)
    verdict = out.split("[판단]")[1]
    assert "호출 한도" in verdict and "인증키" not in verdict and len(srv.log) == 2, (verdict, len(srv.log))


def fix_D2_one_transient_failure_is_not_a_span_limit(tmp):
    for first in (_Resp("<html>500</html>", 500), _Resp('{"weird": 1}'), requests.exceptions.ConnectionError("reset")):
        n = [0]
        def hook(op, p, first=first):
            n[0] += 1
            return first if n[0] == 1 else None
        out = _probe_out(collect.probe_servc, make_cfg(tmp, f"d2{id(first)}", start="2024-03", end="2024-03", num_of_rows=999), "2024-03", Srv(SV, hook))
        verdict = out.split("[판단]")[1]
        assert "max_window_days 를 줄여야" not in verdict and "일시적" in verdict, verdict
    # the real span limit is still recognised: 7 days rejected with 07, one day accepted
    def hook(op, p):
        return _Resp(E07) if int(p["inqryEndDt"][6:8]) - int(p["inqryBgnDt"][6:8]) + 1 > 3 else None
    out = _probe_out(collect.probe_servc, make_cfg(tmp, "d2s", start="2024-03", end="2024-03", num_of_rows=999), "2024-03", Srv(SV, hook))
    assert "하루는 조회되는데 7일은 안 됩니다" in out and "max_window_days 를 줄여야" in out, out


def fix_D3_empty_month_costs_one_call_per_window(tmp):
    srv = Srv(SV)
    out = _probe_out(collect.probe_servc, make_cfg(tmp, "d3", start="2024-02", end="2024-02", num_of_rows=999), "2024-02", srv)
    assert "서버가 이 기간은 0건이라고 답했습니다" in out and "[판단]" in out and len(srv.log) == 5, (len(srv.log), out)
    assert "20240201~20240229" in out and "5개 구간" in out, out


def fix_D4_scrub_gaps(tmp):
    key = "SeCrEtaB3+dE5/gH7jK9lM1nO3pQ5+rS7tU9/vW1xY3zA5bC7dE9fG1hI3+jK5lM7nO9pQ1/rS3tU5vW7xY9zA=="      # made-up value
    from urllib.parse import quote
    enc = quote(key, safe="")
    def texts(resp):
        c = _client(tmp, retries=2, key=key)
        out = []
        with mock.patch.object(c.session, "get", return_value=resp), mock.patch("time.sleep"):
            try:
                c.call("op", {"a": "1"})
            except BaseException as e:
                out += [str(e), str(getattr(e, "msg", "")), str(getattr(e, "raw", "")), str(getattr(e, "why", ""))]
            out.append(J({k: v for k, v in c.diagnose("op", {"a": "2"}).items() if k != "body"}))
        return " || ".join(out)
    for name, resp in (("22 with raw key", _Resp(XML("22", f"LIMITED key={key}"))),
                       ("22 with serviceKey=<encoded>", _Resp(J({"cmmMsgHeader": {"returnReasonCode": "22", "returnAuthMsg": f"LIMITED serviceKey={enc}"}}))),
                       ("lower-case %xx", _Resp(J({"error": "x", "detail": "invalid key " + enc.replace("%2B", "%2b").replace("%2F", "%2f").replace("%3D", "%3d")}))),
                       ("escaped slashes", _Resp('{"error": "bad key", "key": "' + key.replace("/", "\\/") + '"}')),
                       ("cut at 200 (HTTP 500)", _Resp("<html>" + "x" * 164 + key + "</html>", 500)),
                       ("cut at 300 (unreadable)", _Resp(J({"pad": "y" * 236, "key": key}))),
                       ("unreadable, raw attr", _Resp(J({"error": "bad key", "key": key})))):
        t = texts(resp)
        assert "SeCrEt" not in t and "gH7jK9lM1nO3pQ5" not in t, (name, t[:400])


def fix_D5_ratio_outlier_goes_to_the_unconfirmed_side(tmp):
    x = Raw(tmp, "d5")
    for ym in ("202301", "202302", "202303", "202304", "202305", "202306", "202307"):
        x.month("notices_", ym, 12); x.whole(OPN, ym, 12)
        x.month("bsis_", ym, 31 if ym == "202304" else 11)
    r, rep, bad = x.analyse()
    assert r["notices_"]["202304"]["confirmed"] and r["notices_"]["202304"]["flags"] == [], r["notices_"]["202304"]
    assert "RATIO_OUTLIER" in r["bsis_"]["202304"]["flags"], r["bsis_"]["202304"]


def fix_D6_license_lookup_stops_on_repeated_temporary_failures(tmp):
    cfg = make_cfg(tmp, "d6", max_retries=2)
    with mock.patch.object(requests.Session, "get", return_value=_Resp("<html>500</html>", 500)) as g, mock.patch("time.sleep"), contextlib.redirect_stderr(io.StringIO()):
        assert collect.collect_license_limits(cfg, [f"R{i}" for i in range(50)]).empty and g.call_count == 6, g.call_count


def fix_O3_O4_award_probe(tmp):
    aw = [dict(row(1, "2024-03-03 10:00:00"), opengDt="2024-03-04 11:00:00", sucsfbidAmt="1", sucsfbidRate="88", bidwinnrNm="x", prtcptCnum="3")]
    def hook(op, p):
        if "Scsbid" in op:
            return _Resp(OKB(aw, 1)) if p.get("inqryDiv") == "1" else _Resp(XML("10", "INVALID_REQUEST_PARAMETER_ERROR"))
        if "LicenseLimit" in op:
            return _Resp(NODATA)
    cfg = make_cfg(tmp, "o3", start="2024-03", end="2024-03", num_of_rows=999, use_awards=True, award_base_url="http://x/ScsbidInfoService", award_query={"inqryDiv": "2"})
    out = _probe_out(collect.probe, cfg, "2024-03", Srv([row(1, "2024-03-03 10:00:00")], hook))
    assert "[낙찰정보] 기간 조회 파라미터 {'inqryDiv': '1'} 로 결과 확인" in out, out[-600:]
    def hook12(op, p):
        if "Scsbid" in op:
            return _Resp(XML("12", "NO_OPENAPI_SERVICE_ERROR"))
        if "LicenseLimit" in op:
            return _Resp(NODATA)
    srv = Srv([row(1, "2024-03-03 10:00:00")], hook12)
    cfg = make_cfg(tmp, "o4", start="2024-03", end="2024-03", num_of_rows=999, use_awards=True, award_base_url="http://x/ScsbidInfoService")
    out = _probe_out(collect.probe, cfg, "2024-03", srv)
    assert "End Point" in out and sum(1 for op, _ in srv.log if "Scsbid" in op) == 1, out[-600:]


# ── 수정분 검증 검토의 확인 항목 ───────────────────────────────────────────────
def review2_checks(tmp):
    import requests
    import yaml
    from g2b_costdb import collect, selfcheck
    out = []
    J = lambda o: json.dumps(o, ensure_ascii=False)
    os.environ["G2B_SERVICE_KEY"] = "DUMMY-PLACEHOLDER"          # not a key: every HTTP call is mocked

    def make_cfg(name, start="2024-05", end="2024-05", **api):
        cfg = yaml.safe_load(open(os.path.join(REPO, "config.yaml"), encoding="utf-8"))
        d = os.path.join(tmp, name)
        cfg["api"].update(base_url="http://x/BidPublicInfoService", sleep_between_calls_sec=0, max_retries=2, num_of_rows=50, use_awards=False, **api)
        cfg["period"] = {"start": start, "end": end}
        cfg["paths"] = {"data_dir": d, "raw_dir": os.path.join(d, "raw"), "cache_db": os.path.join(d, "cache.sqlite"), "text_dir": os.path.join(d, "text"),
                        "files_dir": os.path.join(d, "files"), "out_dir": os.path.join(d, "out"), "excel_name": "x.xlsx"}
        os.makedirs(cfg["paths"]["raw_dir"], exist_ok=True)
        return cfg

    def server(rows, end_inclusive=True):
        def get(url, params=None, timeout=None):
            b, e, pg, per = params["inqryBgnDt"], params["inqryEndDt"], int(params["pageNo"]), int(params["numOfRows"])
            hit = [r for r in rows if b + "00" <= re.sub(r"\D", "", r["bidNtceDt"]) <= e + ("59" if end_inclusive else "00")]
            return _Resp(J({"response": {"header": {"resultCode": "00", "resultMsg": "정상"},
                                          "body": {"items": hit[(pg - 1) * per: pg * per], "totalCount": len(hit), "numOfRows": per, "pageNo": pg}}}))
        return get

    def row(no, dt, **kw):
        return dict({"bidNtceNo": no, "bidNtceOrd": "000", "bidNtceNm": "가상 공사", "bidNtceDt": dt, "presmptPrce": "1"}, **kw)

    # ① 한 구간 안에서 서버가 똑같은 행을 두 번 주면 그대로 두 줄(월 단위 조회 때와 같게). 구간 경계에서 두 번 온 행만 한 번 저장 — 끝 시각 처리 두 가지 모두
    base = [row(f"R{d:02d}", f"2024-05-{d:02d} 12:00:00") for d in range(1, 32)]
    twin = row("TWIN", "2024-05-10 09:00:00")
    edge = [row("E07", "2024-05-07 00:00:00"), row("E13", "2024-05-13 00:00:30"), row("E13b", "2024-05-13 00:00:30"), row("E13b", "2024-05-13 00:00:30")]
    rows = base + [twin, dict(twin)] + edge
    for inclusive in (True, False):
        cfg = make_cfg(f"twin{int(inclusive)}")
        with mock.patch.object(requests.Session, "get", side_effect=server(rows, inclusive)), mock.patch("time.sleep"):
            collect.collect_notices(cfg)
        got = [json.loads(l)["bidNtceNo"] for l in open(os.path.join(cfg["paths"]["raw_dir"], "notices_202405.jsonl"), encoding="utf-8")]
        meta = json.load(open(os.path.join(cfg["paths"]["raw_dir"], "notices_months.json"), encoding="utf-8"))["202405"]
        want = sorted(r["bidNtceNo"] for r in rows)
        out.append((f"① rows stored exactly as the server has them (end_inclusive={inclusive})", sorted(got) == want and meta["saved"] == len(want)
                    and meta["received"] - meta["dup"] == len(want), f"stored {len(got)} want {len(want)} TWIN x{got.count('TWIN')} E13b x{got.count('E13b')} meta {meta}"))

    # ② check-raw: 서버가 0건이라고 답한 0행짜리 완료 달은 '확인된 달' 이라도 눈으로 확인할 달에 남는다. 서버 건수와 맞는 적은 달(1월)은 표시하지 않는다
    d = os.path.join(tmp, "zero"); raw = os.path.join(d, "raw"); os.makedirs(raw)
    months = [f"2024{m:02d}" for m in range(1, 13)]
    counts = {m: (0 if m == "202407" else 40 if m == "202401" else 120) for m in months}
    con = sqlite3.connect(os.path.join(d, "cache.sqlite"))
    con.execute("CREATE TABLE cache (k TEXT PRIMARY KEY, op TEXT, params TEXT, body TEXT, fetched_at TEXT)")
    import calendar
    for m in months:
        open(os.path.join(raw, f"notices_{m}.jsonl"), "w", encoding="utf-8").write("".join("{}\n" for _ in range(counts[m])))
        p = {"inqryDiv": "1", "inqryBgnDt": m + "010000", "inqryEndDt": f"{m}{calendar.monthrange(2024, int(m[4:]))[1]:02d}2359", "pageNo": "1", "numOfRows": "999"}
        con.execute("INSERT INTO cache VALUES(?,?,?,?,?)", (m, "getBidPblancListInfoCnstwk", J(p), J({"items": [{"a": i} for i in range(counts[m])], "totalCount": counts[m]}), "2026-09-08T10:00:00"))
    con.commit(); con.close()
    open(os.path.join(raw, "notices_done.txt"), "w").write("\n".join(months))
    rep, _, _ = selfcheck.analyse(d, 999, True, dict(selfcheck.DEFAULT_OPS), ("202401", "202412"), today_ym="202611")
    r = {x["ym"]: x for x in rep["notices_"]["rows"]}
    out.append(("② zero-row done month whose cached answer was totalCount 0 is still listed (soft)", bool(r["202407"]["flags"]) and not any(f in selfcheck.HARD for f in r["202407"]["flags"]),
                f"202407 flags {r['202407']['flags']} confirmed {r['202407']['confirmed']}"))
    out.append(("②b low month that matches the server count gets no flag", r["202401"]["flags"] == [] and r["202401"]["confirmed"], f"202401 flags {r['202401']['flags']}"))

    # ③ check-raw 표의 서버건수: 구간으로 받은 달은 경계에서 두 번 센 행을 뺀 값(저장행과 같아야 읽는 사람이 '모자란다'고 오해하지 않는다)
    d = os.path.join(tmp, "col"); raw = os.path.join(d, "raw"); os.makedirs(raw)
    open(os.path.join(raw, "notices_202609.jsonl"), "w").write("{}\n" * 2632)
    open(os.path.join(raw, "notices_done.txt"), "w").write("202609")
    json.dump({"202609": {"total": 2637, "received": 2637, "dup": 5, "saved": 2632, "pages": 5, "windows": 5, "current": False}}, open(os.path.join(raw, "notices_months.json"), "w"))
    rep, _, _ = selfcheck.analyse(d, 999, False, dict(selfcheck.DEFAULT_OPS), ("202609", "202609"), today_ym="202611")
    r = rep["notices_"]["rows"][0]
    out.append(("③ server column of a windowed month equals the stored lines when nothing is missing", r["confirmed"] and r["server"] == r["lines"] == 2632, f"lines {r['lines']} server {r['server']} confirmed {r['confirmed']}"))

    # ④ check-raw 가 캐시를 읽는 동안(큰 캐시는 1~2분) 수집기가 'database is locked' 로 죽지 않고 기다린다
    cfg = make_cfg("lock")
    with mock.patch.object(requests.Session, "get", side_effect=server(base)), mock.patch("time.sleep"):
        collect.collect_notices(cfg)                                                     # 캐시를 만든다
    open(os.path.join(cfg["paths"]["raw_dir"], "notices_done.txt"), "w").write("")       # 다시 받게
    reader = subprocess.Popen([sys.executable, "-c", f"import sqlite3,time\ncon=sqlite3.connect({cfg['paths']['cache_db']!r})\ncur=con.execute('SELECT body FROM cache')\ncur.fetchone()\n"
                               "print('held',flush=True)\ntime.sleep(7)\ncur.fetchall()\ncon.close()"], stdout=subprocess.PIPE, text=True)
    reader.stdout.readline()
    t0, err = time.time(), None
    try:
        with mock.patch.object(requests.Session, "get", side_effect=server(base)):
            with mock.patch.object(collect, "date") as dd:                                # 다른 날(캐시를 다시 쓰지 않게)
                dd.today.return_value = datetime.date(2030, 1, 2)
                collect.collect_notices(cfg)
    except sqlite3.OperationalError as e:
        err = e
    reader.wait()
    out.append(("④ collector survives a reader holding the cache for 7 s", err is None, f"after {time.time() - t0:.1f} s: {err!r}"))
    return out


CHECKS = [v for k, v in sorted(globals().items()) if k.startswith("check_")]
FIXES = [v for k, v in sorted(globals().items()) if k.startswith("fix_")]


def run() -> None:
    keys = ("G2B_SERVICE_KEY", "NO_PROXY", "no_proxy", "HTTP_PROXY", "HTTPS_PROXY", "http_proxy", "https_proxy", "ALL_PROXY", "all_proxy")
    saved = {k: os.environ.get(k) for k in keys}
    for k in keys[3:]:
        os.environ.pop(k, None)
    os.environ["NO_PROXY"] = os.environ["no_proxy"] = "127.0.0.1,localhost"
    os.environ["G2B_SERVICE_KEY"] = "k"            # 인증키가 아니다: 아래의 모든 요청은 가짜 응답
    logging.disable(logging.CRITICAL)
    try:
        for fn in CHECKS + FIXES:
            with tempfile.TemporaryDirectory() as tmp:
                try:
                    fn(tmp)
                except BaseException as e:
                    raise AssertionError(f"{fn.__name__}: {type(e).__name__}: {str(e)[:600]}") from e
        with tempfile.TemporaryDirectory() as tmp:
            bad = [(n, info) for n, ok, info in review2_checks(tmp) if not ok]
            assert not bad, bad
    finally:
        logging.disable(logging.NOTSET)
        for k, v in saved.items():
            if v is None:
                os.environ.pop(k, None)
            else:
                os.environ[k] = v
    print(f"OK: 수집기 검토 테스트 통과 ({len(CHECKS) + len(FIXES)}개 + 2차 검토 항목)")


if __name__ == "__main__":
    run()
