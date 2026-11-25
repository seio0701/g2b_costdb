"""실데이터에서 드러날 수 있는 결함에 대한 회귀 테스트(네트워크·API 키 불필요). 실행: python -m tests.test_regressions
※ 모든 공고·URL은 가상의 예시."""
from __future__ import annotations

import json
import os
import struct
import sys
import tempfile
import zlib
from unittest import mock

import pandas as pd

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
from g2b_costdb import attachments, dedup, discover, extract_llm  # noqa: E402
from g2b_costdb.api_client import ApiConfig, ApiError, DailyBudgetExceeded, G2BClient  # noqa: E402
from g2b_costdb.classify import (classify_notice_kind, classify_trade, classify_work_type, extract_facility_name,  # noqa: E402
                                 load_config, match_categories, parse_amount, work_type_for)


class _Resp:
    def __init__(self, text, status=200, headers=None):
        self.text, self.status_code, self.headers = text, status, headers or {}

    def json(self):
        return json.loads(self.text)


_n = [0]


def _client(tmp, retries=3):
    _n[0] += 1   # 클라이언트마다 별도 캐시 DB(예산 카운터는 DB 단위)
    return G2BClient(ApiConfig(base_url="http://x", service_key="k", max_retries=retries, sleep_between_calls_sec=0),
                     os.path.join(tmp, f"c{_n[0]}.sqlite"))


def test_api_client(tmp):
    xml = ('<OpenAPI_ServiceResponse><cmmMsgHeader><errMsg>SERVICE ERROR</errMsg>'
           '<returnAuthMsg>SERVICE_KEY_IS_NOT_REGISTERED_ERROR</returnAuthMsg><returnReasonCode>30</returnReasonCode>'
           '</cmmMsgHeader></OpenAPI_ServiceResponse>')
    c = _client(tmp)
    with mock.patch.object(c.session, "get", return_value=_Resp(xml)), mock.patch("time.sleep"):
        try:
            c.call("op", {"a": "1"})
            raise AssertionError("ApiError 기대")
        except ApiError as e:
            assert e.code == "30" and "G2B_SERVICE_KEY" in str(e)
    assert c.calls_today() == 1, "키 오류는 재시도 없이 1회만 호출해야 함"
    # 실제 게이트웨이 형식: HTTP 403 + JSON {"cmmMsgHeader": {...}} → 역시 재시도 없이 30
    cj = _client(tmp)
    jerr = json.dumps({"cmmMsgHeader": {"errMsg": "SERVICE ERROR", "returnAuthMsg": "SERVICE_KEY_IS_NOT_REGISTERED_ERROR", "returnReasonCode": "30"}})
    with mock.patch.object(cj.session, "get", return_value=_Resp(jerr, status=403)), mock.patch("time.sleep"):
        try:
            cj.call("op", {"a": "1"})
            raise AssertionError("ApiError 기대")
        except ApiError as e:
            assert e.code == "30"
    assert cj.calls_today() == 1
    from g2b_costdb.api_client import parse_portal_error
    assert parse_portal_error('{"response":{"header":{"resultCode":"20","resultMsg":"SERVICE_ACCESS_DENIED_ERROR"}}}')["code"] == "20"
    assert parse_portal_error('{"response":{"header":{"resultCode":"00"},"body":{}}}') is None
    # 22 트래픽 초과 → DailyBudgetExceeded + 오늘 예산 소진 처리
    xml22 = xml.replace("30", "22").replace("SERVICE_KEY_IS_NOT_REGISTERED_ERROR", "LIMITED_NUMBER_OF_SERVICE_REQUESTS_EXCEEDS_ERROR")
    with mock.patch.object(c.session, "get", return_value=_Resp(xml22)), mock.patch("time.sleep"):
        try:
            c.call("op", {"a": "2"})
            raise AssertionError("DailyBudgetExceeded 기대")
        except DailyBudgetExceeded:
            pass
    assert c.calls_today() >= c.cfg.daily_call_budget
    # 03 NODATA → 빈 결과(정상)
    c2 = _client(tmp)
    nodata = json.dumps({"response": {"header": {"resultCode": "03", "resultMsg": "NODATA_ERROR"}, "body": {}}})
    with mock.patch.object(c2.session, "get", return_value=_Resp(nodata)), mock.patch("time.sleep"):
        assert list(c2.iter_all("op", {"a": "3"})) == []
    # totalCount 누락 → 짧은 페이지까지 계속
    c3 = _client(tmp)
    c3.cfg.num_of_rows = 2
    pages = {"1": [{"i": 1}, {"i": 2}], "2": [{"i": 3}, {"i": 4}], "3": [{"i": 5}]}

    def fake_get(url, params=None, timeout=None):
        return _Resp(json.dumps({"response": {"header": {"resultCode": "00"}, "body": {"items": pages[params["pageNo"]]}}}))
    with mock.patch.object(c3.session, "get", side_effect=fake_get), mock.patch("time.sleep"):
        assert [it["i"] for it in c3.iter_all("op", {"a": "4"})] == [1, 2, 3, 4, 5]
    # 서버가 요청보다 작은 페이지(예: 100건)만 돌려줘도 totalCount 까지 계속 조회
    c6 = _client(tmp); c6.cfg.num_of_rows = 999
    pages6 = {"1": [{"i": i} for i in range(100)], "2": [{"i": i} for i in range(100, 150)]}

    def fake_get6(url, params=None, timeout=None):
        return _Resp(json.dumps({"response": {"header": {"resultCode": "00"}, "body": {"items": pages6.get(params["pageNo"], []), "totalCount": 150}}}))
    with mock.patch.object(c6.session, "get", side_effect=fake_get6), mock.patch("time.sleep"):
        assert len(list(c6.iter_all("op", {"a": "6"}))) == 150
    # Encoding 키 자동 변환
    c4 = G2BClient(ApiConfig(base_url="http://x", service_key="ab%2Bcd%3D%3D"), os.path.join(tmp, "d.sqlite"))
    assert c4.cfg.service_key == "ab+cd=="
    # 캐시 salt: 같은 파라미터라도 salt 가 다르면 재호출
    c5 = _client(tmp)
    ok = json.dumps({"response": {"header": {"resultCode": "00"}, "body": {"items": [{"x": 1}], "totalCount": 1}}})
    with mock.patch.object(c5.session, "get", return_value=_Resp(ok)) as g, mock.patch("time.sleep"):
        c5.call("op", {"q": "1"}); c5.call("op", {"q": "1"}); c5.call("op", {"q": "1"}, cache_salt="2026-09-05")
        assert g.call_count == 2


def test_api_response_shapes_and_windows(tmp):
    """2026-11-17: 용역 조회가 'totalCount=None 표본 0건 / 항목 없음' 으로 나온 원인 — 서버가 조회 기간 초과(07)를 HTTP 200 의 다른 모양
    ({"nkoneps.com.response.ResponseError": …})으로 돌려주는데 그것을 '자료 없음'으로 읽었다. 읽을 수 없는 응답·잘린 쪽은 오류로 드러내고
    (캐시에 넣지 않음, 달을 완료로 표시하지 않음, 기존 원문을 지우지 않음), 한 달을 7일 이하 구간으로 나눠 조회한다. check-raw 로 이미 받은 자료를 점검한다."""
    import contextlib
    import copy
    import io
    import sqlite3
    from g2b_costdb import collect, selfcheck
    from g2b_costdb.api_client import IncompleteResult, UnexpectedResponse, validate_body
    J = lambda o: json.dumps(o, ensure_ascii=False)
    E07 = J({"nkoneps.com.response.ResponseError": {"header": {"resultCode": "07", "resultMsg": "입력범위값 초과 에러"}}})
    OKB = lambda items, total: J({"response": {"header": {"resultCode": "00", "resultMsg": "정상"}, "body": {"items": items, "numOfRows": 5, "pageNo": 1, "totalCount": total}}})

    def outcome(text, status=200, retries=3):
        """(결과 종류, HTTP 호출 수, 캐시 행 수). 종류: 'data'/'nodata'/ApiError 코드/'unreadable'/'retried'"""
        c = _client(tmp, retries)
        with mock.patch.object(c.session, "get", return_value=_Resp(text, status)) as g, mock.patch("time.sleep"):
            try:
                body = c.call("op", {"inqryDiv": "1", "pageNo": 1, "numOfRows": 5})
                kind = "data" if c._items(body) else "nodata"
                assert validate_body(body, 1, 5) is None, body
            except UnexpectedResponse as e:
                kind = "unreadable"
                assert e.code == "XX" and "응답 앞부분" in str(e)
            except ApiError as e:
                kind = e.code
            except RuntimeError as e:
                kind = "retried" if "재시도 소진" in str(e) else f"?{e}"
            return kind, g.call_count, c.conn.execute("SELECT COUNT(*) FROM cache").fetchone()[0]

    # ① 응답 모양별 판정: 오류·읽을 수 없는 응답은 한 번도 빈 결과가 되지 않고 캐시에도 들어가지 않는다
    for text, want, calls in (
            (E07, "07", 1),                                                                                  # 조회 기간 초과: 재시도 없이 오류
            (J({"nkoneps.com.response.ResponseError": {"header": {"resultCode": "08", "resultMsg": "필수값 입력 에러"}}}), "08", 1),
            (J({"OpenAPI_ServiceResponse": {"cmmMsgHeader": {"returnReasonCode": "30", "returnAuthMsg": "SERVICE_KEY_IS_NOT_REGISTERED_ERROR"}}}), "30", 1),
            (J({"cmmMsgHeader": {"returnReasonCode": "04", "returnAuthMsg": "HTTP_ERROR"}}), "retried", 3),   # 일시 오류: 재시도
            (J({"cmmMsgHeader": {"returnReasonCode": "23", "returnAuthMsg": "PER_SECOND"}}), "retried", 3),   # 초당 한도: 재시도
            ("{}", "unreadable", 1), (J({"response": {}}), "unreadable", 1), (J({"response": None}), "unreadable", 1),
            (J({"response": "ERROR"}), "unreadable", 1), (J({"error": {"code": 500}}), "unreadable", 1),
            (J({"response": {"header": {"resultCode": "00"}}}), "unreadable", 1),                              # 정상 코드인데 body 없음
            (J({"response": {"header": {"resultCode": "00"}, "body": ""}}), "unreadable", 1),
            (J({"response": {"header": {"resultCode": ""}, "body": {}}}), "unreadable", 1),
            (J({"response": {"header": {"resultCode": "00"}, "body": ["x"]}}), "unreadable", 1),
            (OKB([], 37), "retried", 3),                                                                      # 잘린 쪽(totalCount>0, 0건): 재시도 뒤 실패
            (J({"response": {"header": {"resultCode": 7, "resultMsg": "x"}}}), "07", 1),                       # 숫자 코드
            ("<response><header><resultCode>07</resultCode><resultMsg>입력범위값 초과 에러</resultMsg></header></response>", "07", 1),
            ("<html>502</html>", "retried", 3), ("[1, 2]", "retried", 3)):
        got = outcome(text)
        assert got == (want, calls, 0), (text[:70], got)
    # 정상·자료 없음은 그대로: 항목 있음 / totalCount 0 / 03(표준·오류 모양·XML) / items 가 빈 문자열
    for text, want in ((OKB([{"a": 1}], 1), "data"), (OKB([], 0), "nodata"), (OKB("", 0), "nodata"),
                       (J({"response": {"header": {"resultCode": "03", "resultMsg": "NODATA_ERROR"}, "body": {}}}), "nodata"),
                       (J({"nkoneps.com.response.ResponseError": {"header": {"resultCode": "03", "resultMsg": "데이터 없음"}}}), "nodata"),
                       ("<OpenAPI_ServiceResponse><cmmMsgHeader><returnReasonCode>03</returnReasonCode></cmmMsgHeader></OpenAPI_ServiceResponse>", "nodata")):
        got = outcome(text)
        assert got == (want, 1, 1), (text[:70], got)
    # 오류 문구에 서버가 준 말이 그대로 실리고, 응답에 인증키가 섞여 있어도 가려진다
    c = G2BClient(ApiConfig(base_url="http://x", service_key="SECRET+KEY/abc==", sleep_between_calls_sec=0), os.path.join(tmp, "scrub.sqlite"))
    with mock.patch.object(c.session, "get", return_value=_Resp(E07)), mock.patch("time.sleep"):
        try:
            c.call("op", {"a": "1"})
            raise AssertionError("ApiError 기대")
        except ApiError as e:
            assert e.code == "07" and "입력범위값 초과 에러" in str(e) and "7일" in str(e), str(e)
    echo = J({"error": "bad", "url": "/x?serviceKey=SECRET%2BKEY%2Fabc%3D%3D&type=json", "key": "SECRET+KEY/abc=="})
    with mock.patch.object(c.session, "get", return_value=_Resp(echo)), mock.patch("time.sleep"):
        try:
            c.call("op", {"a": "2"})
            raise AssertionError("UnexpectedResponse 기대")
        except UnexpectedResponse as e:
            assert "SECRET" not in str(e) and "SECRET" not in e.msg and "***" in str(e), str(e)
        d = c.diagnose("op", {"a": "3"})
        assert d["status"] == "unreadable" and "SECRET" not in J(d) and d["keys"] == ["error", "url", "key"], d
    with mock.patch.object(c.session, "get", return_value=_Resp(E07)), mock.patch("time.sleep"):
        d = c.diagnose("op", {"a": "4"})
        assert (d["status"], d["code"], d["msg"], d["http"]) == ("error", "07", "입력범위값 초과 에러", 200), d
    with mock.patch.object(c.session, "get", return_value=_Resp(OKB([{"a": 1}], 9))), mock.patch("time.sleep"):
        d = c.diagnose("op", {"a": "5"})
        assert (d["status"], d["total"], d["n"]) == ("ok", 9, 1) and d["body"]["totalCount"] == 9, d
    assert c.conn.execute("SELECT COUNT(*) FROM cache").fetchone()[0] == 0, "진단 호출은 캐시에 쓰지 않는다"
    # 예전 버전이 캐시에 넣어 둔 빈 본문('{}')·잘린 쪽은 없는 것으로 보고 다시 호출한다(정상 본문은 그대로 쓴다)
    c = _client(tmp)
    import hashlib
    for i, body in enumerate(("{}", J({"items": [], "totalCount": 37}), '"x"', J({"items": [{"a": 1}], "totalCount": 1}))):
        p = {"q": str(i), "pageNo": "1", "numOfRows": "5"}
        k = hashlib.sha1(("op" + json.dumps(p, sort_keys=True, ensure_ascii=False)).encode()).hexdigest()
        c.conn.execute("INSERT OR REPLACE INTO cache VALUES(?,?,?,?,?)", (k, "op", json.dumps(p), body, "2026-09-10T00:00:00"))
    c.conn.commit()
    with mock.patch.object(c.session, "get", return_value=_Resp(OKB([{"new": 1}], 1))) as g, mock.patch("time.sleep"):
        got = [c.call("op", {"q": str(i), "pageNo": 1, "numOfRows": 5}) for i in range(4)]
        assert g.call_count == 3 and [b["items"][0] for b in got] == [{"new": 1}, {"new": 1}, {"new": 1}, {"a": 1}], (g.call_count, got)

    # ② 쪽을 끝까지 돌았는데 받은 건수가 totalCount 보다 적으면 조용히 끝내지 않는다 / 전량 수집은 totalCount 없는 응답을 오류로 본다
    c = _client(tmp); c.cfg.num_of_rows = 5
    pages = {"1": [{"i": i} for i in range(5)], "2": [{"i": i} for i in range(5, 8)]}
    with mock.patch.object(c.session, "get", side_effect=lambda url, params=None, timeout=None: _Resp(J(
            {"response": {"header": {"resultCode": "00"}, "body": {"items": pages.get(params["pageNo"], []), "totalCount": 9}}}))), mock.patch("time.sleep"):
        try:
            list(c.iter_all("op", {"a": "1"}))
            raise AssertionError("IncompleteResult 기대")
        except IncompleteResult as e:
            assert "totalCount=9" in str(e) and "8건" in str(e), str(e)
    c = _client(tmp); c.cfg.num_of_rows = 5
    with mock.patch.object(c.session, "get", return_value=_Resp(J({"response": {"header": {"resultCode": "00"}, "body": {"items": [{"i": 1}]}}}))), mock.patch("time.sleep"):
        assert len(list(c.iter_all("op", {"a": "2"}))) == 1                      # 기본은 관대(짧은 쪽에서 끝)
        try:
            list(c.iter_all("op", {"a": "3"}, require_total=True))
            raise AssertionError("UnexpectedResponse 기대")
        except UnexpectedResponse:
            pass

    # ③ 조회 구간 나누기: 달력 날짜 7일 이하, 구간의 끝 = 다음 구간의 시작(경계의 1분을 두 구간에 걸쳐 조회 — 끝 1분의 처리가 어느 쪽이어도 빠지지 않게)
    sw = G2BClient.split_window
    assert sw("202608010000", "202608312359", 7) == [("202608010000", "202608070000"), ("202608070000", "202608130000"), ("202608130000", "202608190000"),
                                                     ("202608190000", "202608250000"), ("202608250000", "202608312359")]
    assert len(sw("202602010000", "202602282359", 7)) == 5 and sw("202402010000", "202402292359", 7)[-1] == ("202402250000", "202402292359")
    assert sw("202608010000", "202608312359", 0) == [("202608010000", "202608312359")] == sw("202608010000", "202608312359", 31)
    assert sw("202608010000", "202608072359", 7) == [("202608010000", "202608072359")]
    assert sw("202608010000", "202608312359", 1)[:2] == [("202608010000", "202608012359"), ("202608020000", "202608022359")] and len(sw("202608010000", "202608312359", 1)) == 31
    for last in (28, 29, 30, 31):
        for n in (2, 3, 6, 7, 8, 10, 27):
            w = sw("202601010000", f"202601{last:02d}2359", n)
            assert w[0][0] == "202601010000" and w[-1][1] == f"202601{last:02d}2359" and all(e1 == b2 for (_, e1), (b2, _) in zip(w, w[1:])), (last, n, w)
            assert all(int(e[6:8]) - int(b[6:8]) + 1 <= n for b, e in w), (last, n, w)

    # ④ 수집(가짜 서버: 7일 넘는 기간은 07 로 거절, 조회 기간으로 걸러 돌려줌)
    import yaml
    from tests.mock_g2b import KEY, MockG2B
    server = MockG2B()
    base = server.start()
    old_env = {k: os.environ.get(k) for k in ("G2B_SERVICE_KEY", "NO_PROXY")}
    os.environ.update(G2B_SERVICE_KEY=KEY, NO_PROXY="127.0.0.1,localhost")
    root = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))

    def make_cfg(name, **api):
        cfg = yaml.safe_load(open(os.path.join(root, "config.yaml"), encoding="utf-8"))
        d = os.path.join(tmp, name)
        cfg["api"].update(base_url=base + "/BidPublicInfoService", sleep_between_calls_sec=0, max_retries=2, use_awards=False, **api)
        cfg["period"] = {"start": "2024-01", "end": "2024-04"}
        cfg["paths"] = {"data_dir": d, "raw_dir": os.path.join(d, "raw"), "cache_db": os.path.join(d, "cache.sqlite"), "text_dir": os.path.join(d, "text"),
                        "files_dir": os.path.join(d, "files"), "out_dir": os.path.join(d, "out"), "excel_name": "x.xlsx"}
        os.makedirs(cfg["paths"]["raw_dir"], exist_ok=True)
        return cfg

    def state(cfg, prefix="notices_"):
        raw = cfg["paths"]["raw_dir"]
        dp = os.path.join(raw, prefix + "done.txt")
        files = {fn[len(prefix):-6]: sum(1 for _ in open(os.path.join(raw, fn), encoding="utf-8")) for fn in sorted(os.listdir(raw)) if fn.startswith(prefix) and fn.endswith(".jsonl")}
        return (sorted(open(dp, encoding="utf-8").read().split()) if os.path.exists(dp) else []), files

    OPN = "getBidPblancListInfoCnstwk"
    by_month = {}
    for n in server.data["notices"]:
        by_month[n["bidNtceDt"][:7].replace("-", "")] = by_month.get(n["bidNtceDt"][:7].replace("-", ""), 0) + 1
    try:
        with mock.patch("time.sleep"):
            # 7일 구간으로 나눠 받으면 모든 공고를 한 번씩 받는다(구간 사이 중복·누락 없음)
            cfg = make_cfg("win7")
            c0 = server.calls[OPN]
            df = collect.collect_notices(cfg)
            done, files = state(cfg)
            assert done == ["202401", "202402", "202403", "202404"] and {k: v for k, v in files.items() if v} == by_month, (done, files, by_month)
            assert len(df) == len(server.data["notices"]) and df[["bidNtceNo", "bidNtceOrd"]].drop_duplicates().shape[0] == len(df)
            assert server.calls[OPN] - c0 == 20, "4개월 × 5구간(각 1쪽)"
            meta = json.load(open(os.path.join(cfg["paths"]["raw_dir"], "notices_months.json"), encoding="utf-8"))
            assert meta["202401"]["total"] == by_month["202401"] == meta["202401"]["received"] and meta["202401"]["windows"] == 5 and meta["202404"]["total"] == 0, meta
            c1 = server.calls[OPN]
            collect.collect_notices(cfg)
            assert server.calls[OPN] == c1, "완료한 달은 다시 호출하지 않는다"
            # 한 달 통째로 조회(예전 방식)하면 서버가 07 로 거절한다 → 빈 달로 저장하지 않고, 완료 표시도 하지 않고, 세 달 연속 실패에서 멈춘다
            cfg = make_cfg("month", max_window_days=31)
            c0 = server.calls[OPN]
            with contextlib.redirect_stderr(io.StringIO()):
                df = collect.collect_notices(cfg)
            assert state(cfg) == ([], {}) and len(df) == 0, state(cfg)
            assert server.calls[OPN] - c0 == 3, "오류는 재시도하지 않고(달마다 1회) 연속 3개월 실패에서 중단"
            assert sqlite3.connect(cfg["paths"]["cache_db"]).execute("SELECT COUNT(*) FROM cache").fetchone()[0] == 0, "오류 응답은 캐시에 넣지 않는다"
            # 한 구간의 응답이 읽을 수 없는 모양이면 그 달만 실패: 완료 표시 없음·파일 없음. 다음 실행 때 받은 구간은 캐시에서, 실패한 구간만 다시 받는다
            cfg = make_cfg("fault")
            server.faults[(OPN, "20240213", 1)] = (200, b"{}")
            c0 = server.calls[OPN]
            collect.collect_notices(cfg)
            done, files = state(cfg)
            assert done == ["202401", "202403", "202404"] and "202402" not in files, (done, files)
            assert server.calls[OPN] - c0 == 18, "2월은 세 번째 구간에서 멈춤(5+3+5+5)"
            server.faults.clear()
            c0 = server.calls[OPN]
            collect.collect_notices(cfg)
            done, files = state(cfg)
            assert "202402" in done and files["202402"] == by_month["202402"] and server.calls[OPN] - c0 == 3, (done, files, server.calls[OPN] - c0)
            # 잘린 쪽(totalCount>0 인데 0건)도 같은 처리
            cfg = make_cfg("trunc")
            server.faults[(OPN, "20240107", 1)] = (200, OKB([], 4).encode("utf-8"))
            collect.collect_notices(cfg)
            done, files = state(cfg)
            assert "202401" not in done and "202401" not in files, (done, files)
            server.faults.clear()
            # 다시 받은 결과가 기존 파일보다 적으면 기존 원문을 .prev 로 남긴다(원문을 조용히 잃지 않는다)
            cfg = make_cfg("shrink")
            collect.collect_notices(cfg)
            raw = cfg["paths"]["raw_dir"]
            open(os.path.join(raw, "notices_done.txt"), "w", encoding="utf-8").write("202402\n202403\n202404")      # 1월을 다시 받게 함
            keep = server.data["notices"]
            server.data["notices"] = [n for n in keep if not n["bidNtceDt"].startswith("2024-01-1")]
            collect.collect_notices(cfg)
            server.data["notices"] = keep
            prev = os.path.join(raw, "notices_202401.jsonl.prev")
            assert os.path.exists(prev) and sum(1 for _ in open(prev, encoding="utf-8")) == by_month["202401"] > state(cfg)[1]["202401"], state(cfg)

            # ⑤ probe --servc: 표본을 받지 못하면 조회 조건을 바꿔 가며 서버의 답(코드·문구)을 그대로 보여 준다 — '항목 없음' 으로 뭉뚱그리지 않는다
            cfg = make_cfg("probe")
            cfg["servc"]["period_end"] = "2024-04"
            server.max_span_days = 3                                 # 서버 제한이 설정값(7일)보다 짧아진 경우
            buf = io.StringIO()
            with contextlib.redirect_stdout(buf):
                collect.probe_servc(cfg, "2024-03")
            out = buf.getvalue()
            assert "표본을 받지 못했습니다" in out and "① 20240301~20240307(7일), 50건: HTTP 200 | 서버가 요청을 거절 | 코드 07" in out, out
            assert "서버 문구: 입력범위값 초과 에러" in out and "③ 한 달 전체, 50건: HTTP 200 | 서버가 요청을 거절 | 코드 07" in out, out
            assert "하루는 조회되는데 7일은 안 됩니다" in out and "max_window_days" in out and "항목 없음" not in out and "totalCount=None" not in out, out
            server.max_span_days = 7
            server.faults[("getBidPblancListInfoServc", "20240301", 1)] = (200, b'{"weird": 1}')
            buf = io.StringIO()
            with contextlib.redirect_stdout(buf):
                collect.probe_servc(cfg, "2024-03")
            out = buf.getvalue()
            assert "읽을 수 없는 응답" in out and "최상위 키 ['weird']" in out and '본문 앞부분: {"weird": 1}' in out, out
            assert "⑥ 대조: 공사 공고 목록, 7일: HTTP 200 | 정상" in out and "공사 공고는 조회되는데 용역 공고만 안 됩니다" in out, out
            server.faults.clear()
            buf = io.StringIO()
            with contextlib.redirect_stdout(buf):
                collect.probe_servc(cfg, "2024-02")                  # 2024-02 에는 용역 공고가 없다 → 오류가 아니라 서버가 0건이라고 답한 것
            assert "서버가 이 기간은 0건이라고 답했습니다" in buf.getvalue() and "[판단]" in buf.getvalue(), buf.getvalue()
    finally:
        server.stop()
        for k, v in old_env.items():
            if v is None:
                os.environ.pop(k, None)
            else:
                os.environ[k] = v

    # ⑥ check-raw: 예전 수집기가 남겼을 흔적을 API 호출 없이 찾는다(쪽 크기 5)
    d = os.path.join(tmp, "selfcheck")
    raw = os.path.join(d, "raw")
    os.makedirs(raw)
    con = sqlite3.connect(os.path.join(d, "cache.sqlite"))
    con.execute("CREATE TABLE cache (k TEXT PRIMARY KEY, op TEXT, params TEXT, body TEXT, fetched_at TEXT)")
    months = {"202301": 13, "202302": 12, "202303": 0, "202304": 10, "202305": 5, "202306": 14, "202307": 15, "202308": 11, "202309": 12}
    truth = {"202301": 13, "202302": 12, "202303": 14, "202304": 23, "202305": 12, "202306": 14, "202307": 15, "202308": 11, "202309": 12}
    last = {"202301": 31, "202302": 28, "202303": 31, "202304": 30, "202305": 31, "202306": 30, "202307": 31, "202308": 31, "202309": 30}
    k = [0]

    def put(op, ym, page, body, fetched="2026-09-10T10:00:00", bgn=None, end=None, rows=5):
        k[0] += 1
        p = {"inqryDiv": "1", "inqryBgnDt": bgn or ym + "010000", "inqryEndDt": end or f"{ym}{last[ym]:02d}2359", "pageNo": str(page), "numOfRows": str(rows)}
        con.execute("INSERT INTO cache VALUES(?,?,?,?,?)", (f"k{k[0]}", op, json.dumps(p), body if isinstance(body, str) else json.dumps(body), fetched))
    for ym, n in months.items():
        open(os.path.join(raw, f"notices_{ym}.jsonl"), "w", encoding="utf-8").write("".join('{"x": 1}\n' for _ in range(n)))
        total = truth[ym]
        for page in range(1, (n + 4) // 5 + 1):                       # 저장된 만큼의 쪽은 정상 본문으로 캐시에 있다
            put(OPN, ym, page, {"items": [{"x": 1}] * min(5, n - (page - 1) * 5), "totalCount": total})
    put(OPN, "202303", 1, "{}")                                       # 1쪽이 읽지 못한 응답 → 0행인데 완료
    put(OPN, "202304", 3, "{}")                                       # 3쪽이 읽지 못한 응답 → 10행(쪽 배수)에서 끊김
    put(OPN, "202305", 2, {"items": [], "totalCount": 12})            # 2쪽이 잘린 쪽 → 5행에서 끊김
    put("getBidPblancListInfoServc", "202308", 1, "{}", fetched="2026-11-17T07:13:24", rows=50)     # 오늘의 probe 가 남긴 행
    con.commit(); con.close()
    open(os.path.join(raw, "notices_done.txt"), "w", encoding="utf-8").write("\n".join(m for m in months if m != "202309"))
    rep, bad, meta = selfcheck.analyse(d, page_size=5, use_cache=True, ops=dict(selfcheck.DEFAULT_OPS), today_ym="202611")
    fl = {r["ym"]: set(r["flags"]) for r in rep["notices_"]["rows"] if r["flags"]}
    assert set(fl) == {"202303", "202304", "202305", "202309"}, fl
    assert {"EMPTY_DONE", "BAD_CACHE_BODY"} <= fl["202303"] and {"SHORT_VS_TOTAL", "BAD_CACHE_BODY", "PAGE_MULTIPLE"} <= fl["202304"], fl
    assert {"SHORT_VS_TOTAL", "BAD_CACHE_BODY", "CACHE_SUM_LT_TOTAL"} <= fl["202305"] and fl["202309"] == {"NOT_DONE"}, fl
    assert selfcheck.hard_months(rep) == {"notices_": ["202303", "202304", "202305"]} and len(bad) == 4, (selfcheck.hard_months(rep), bad)
    assert not [r for r in rep["notices_"]["rows"] if r["ym"] == "202307" and r["flags"]], "줄 수가 쪽 크기의 배수(15)여도 서버 건수와 맞으면 정상"
    buf = io.StringIO()
    with contextlib.redirect_stdout(buf):
        n_hard = selfcheck.print_report(rep, bad, meta)
    out = buf.getvalue()
    assert n_hard == 3 and "다시 받아야 할 달 3개: 202303, 202304, 202305" in out and "눈으로 확인할 달 1개: 202309" in out and "check-raw --reopen" in out, out
    assert "getBidPblancListInfoServc 조회 시작일=20230801" in out, out
    before = {fn: os.path.getmtime(os.path.join(raw, fn)) for fn in os.listdir(raw) if fn.endswith(".jsonl")}
    assert selfcheck.reopen(raw, rep) == {"notices_": ["202303", "202304", "202305"]}
    assert open(os.path.join(raw, "notices_done.txt"), encoding="utf-8").read().split() == ["202301", "202302", "202306", "202307", "202308"]
    assert before == {fn: os.path.getmtime(os.path.join(raw, fn)) for fn in os.listdir(raw) if fn.endswith(".jsonl")}, "원문 파일은 건드리지 않는다"
    # 캐시가 없으면 줄 수만으로: 0행 완료·쪽 배수·이웃 달보다 적음
    rep2, _, meta2 = selfcheck.analyse(d, page_size=5, use_cache=False, ops=dict(selfcheck.DEFAULT_OPS), today_ym="202611")
    assert not meta2["have_cache"] and "PAGE_MULTIPLE" in {f for r in rep2["notices_"]["rows"] for f in r["flags"]}
    # 7일 구간으로 받은 달: 수집 기록(months.json)과 대조하고, 서버가 0건이라고 답한 달은 빈 달이어도 정상. 기록이 없으면 캐시의 구간별 쪽만 본다
    d2 = os.path.join(tmp, "selfcheck_win")
    raw2 = os.path.join(d2, "raw")
    os.makedirs(raw2)
    con = sqlite3.connect(os.path.join(d2, "cache.sqlite"))
    con.execute("CREATE TABLE cache (k TEXT PRIMARY KEY, op TEXT, params TEXT, body TEXT, fetched_at TEXT)")
    for ym, per in (("202301", [3, 2, 4, 1, 2]), ("202302", [0, 0, 0, 0, 0]), ("202303", [2, 2, 2, 2, 2])):
        for i, (b, e) in enumerate(G2BClient.split_window(ym + "010000", f"{ym}{last[ym]:02d}2359", 7)):
            put(OPN, ym, 1, {"items": [{"x": 1}] * per[i], "totalCount": per[i]}, bgn=b, end=e)
    put(OPN, "202304", 1, {"items": [{"x": 1}] * 5, "totalCount": 9}, bgn="202304010000", end="202304070000")       # 받다 만 달(완료 표시 없음): 쪽이 모자란 것이 정상
    put("getBidPblancListInfoServc", "202303", 1, "{}", bgn="202303010000", end="202303312359", rows=50)             # 예전 probe 가 남긴 행(쪽 크기가 다름)
    con.commit(); con.close()
    open(os.path.join(raw2, "notices_202301.jsonl"), "w", encoding="utf-8").write('{"x": 1}\n' * 11)       # 수집 때 12행을 저장했다고 기록돼 있는데 11행
    open(os.path.join(raw2, "notices_202302.jsonl"), "w", encoding="utf-8").write("")
    open(os.path.join(raw2, "notices_202303.jsonl"), "w", encoding="utf-8").write('{"x": 1}\n' * 10)
    open(os.path.join(raw2, "notices_done.txt"), "wb").write("\ufeff202301\r\n202302\r\n202303\r\n메모\r\n".encode("utf-8"))      # 메모장이 붙인 BOM·CRLF·메모
    json.dump({"202301": {"total": 12, "received": 12, "dup": 0, "saved": 12}, "202302": {"total": 0, "received": 0, "dup": 0, "saved": 0},
               "202303": {"total": 10, "received": 10, "dup": 0, "saved": 10}}, open(os.path.join(raw2, "notices_months.json"), "w", encoding="utf-8"))
    for fast in (True, False):                                         # SQLite 의 JSON 함수로 읽는 빠른 길과 파이썬으로 읽는 길의 결과가 같다
        keep_sql = selfcheck._FAST_SQL
        if not fast:
            selfcheck._FAST_SQL = "SELECT nosuchfunc(body) FROM cache"
        try:
            rep3, bad3, meta3 = selfcheck.analyse(d2, page_size=5, use_cache=True, ops=dict(selfcheck.DEFAULT_OPS), today_ym="202611", period=("202301", "202304"))
        finally:
            selfcheck._FAST_SQL = keep_sql
        r3 = {r["ym"]: r for r in rep3["notices_"]["rows"]}
        assert r3["202301"]["flags"] == ["SHORT_VS_SAVED"] and r3["202302"]["flags"] == [] and r3["202303"]["flags"] == [] and r3["202303"]["confirmed"], (fast, r3)
        assert r3["202304"]["flags"] == ["NEVER_COLLECTED"] and selfcheck.hard_months(rep3) == {"notices_": ["202301"]}, r3["202304"]     # 받는 중인 달의 모자란 쪽은 문제가 아니다
        assert rep3["notices_"]["junk"] == ["메모"] and rep3["notices_"]["done"] == 3 and len(bad3) == 1, (rep3["notices_"]["junk"], bad3)
    buf = io.StringIO()
    with contextlib.redirect_stdout(buf):
        selfcheck.print_report(rep3, bad3, meta3)
    out = buf.getvalue()
    assert "달(YYYYMM)이 아닌 글자가 있어 무시했습니다" in out and "참고: 캐시에 남아 있는 '읽지 못한 응답' 행 1개" in out and "위의 달 판정과는 별개" in out, out
    assert selfcheck.reopen(raw2, rep3) == {"notices_": ["202301"]} and collect._read_done(os.path.join(raw2, "notices_done.txt")) == {"202302", "202303"}
    # 걸러서 저장하는 수집(용역): 줄 수는 수집 기록의 '저장 건수' 와 대조 — 한 건도 남기지 않은 달은 정상, 파일이 줄어든 달은 다시 받는다
    open(os.path.join(raw2, "servc_202301.jsonl"), "w", encoding="utf-8").write("")
    open(os.path.join(raw2, "servc_202302.jsonl"), "w", encoding="utf-8").write('{"x": 1}\n' * 3)
    open(os.path.join(raw2, "servc_done.txt"), "w", encoding="utf-8").write("202301\n202302")
    json.dump({"202301": {"total": 20000, "received": 20000, "dup": 0, "saved": 0}, "202302": {"total": 18000, "received": 18000, "dup": 2, "saved": 900}},
              open(os.path.join(raw2, "servc_months.json"), "w", encoding="utf-8"))
    rep4, bad4, meta4 = selfcheck.analyse(d2, page_size=5, use_cache=True, ops=dict(selfcheck.DEFAULT_OPS), today_ym="202611")
    r4 = {r["ym"]: r["flags"] for r in rep4["servc_"]["rows"]}
    assert r4["202301"] == [] and r4["202302"] == ["SHORT_VS_SAVED"] and "202303" not in r4, r4          # 202303 의 예전 probe 행은 달 판정에 쓰지 않는다
    buf = io.StringIO()
    with contextlib.redirect_stdout(buf):
        selfcheck.print_report(rep4, bad4, meta4)
    assert "pipeline collect --servc" in buf.getvalue() and "걸러서 저장하는 수집" in buf.getvalue(), buf.getvalue()
    # 캐시 파일을 읽을 수 없으면(빈 파일·깨진 파일) 죽지 않고 캐시 없이 점검한다 / 달이 아닌 글자만 든 done.txt 에도 멈추지 않는다
    d3 = os.path.join(tmp, "selfcheck_bad")
    os.makedirs(os.path.join(d3, "raw"))
    open(os.path.join(d3, "cache.sqlite"), "wb").write(b"not a database, just text" * 50)
    open(os.path.join(d3, "raw", "notices_202301.jsonl"), "w", encoding="utf-8").write('{"x": 1}\n' * 7)
    open(os.path.join(d3, "raw", "notices_done.txt"), "wb").write("\ufeffdone 메모 2026-09".encode("utf-8"))
    rep5, bad5, meta5 = selfcheck.analyse(d3, page_size=5, use_cache=True, ops=dict(selfcheck.DEFAULT_OPS), today_ym="202611")
    assert not meta5["have_cache"] and meta5["cache_error"] and [r["flags"] for r in rep5["notices_"]["rows"] if r["ym"] == "202301"] == [["NOT_DONE"]], (meta5, rep5)
    buf = io.StringIO()
    with contextlib.redirect_stdout(buf):
        selfcheck.print_report(rep5, bad5, meta5)
    assert "읽지 못했습니다" in buf.getvalue() and selfcheck.month_seq("\ufeff202301", "202303") == [], buf.getvalue()


def test_collector_review(tmp):
    """수집기 수정분(7일 구간·전량 확인) 독립 검토에서 나온 결함: ① 전량을 받지 못한 조회의 쪽이 캐시에 남아 서버가 정상이 된 뒤에도 같은 실패를 되풀이
    ② 다시 받은 결과가 두 번 줄면 .prev 가 덮여 가장 큰 원본이 사라짐 ③ 면허제한 조회가 모두 08(필수값)로 거절돼도 공고 수만큼 호출
    ④ max_window_days 가 빈 값이면 TypeError ⑤ totalCount 음수·무한대."""
    import requests
    import yaml
    from g2b_costdb import collect
    from g2b_costdb.api_client import IncompleteResult, UnexpectedResponse, validate_body, body_total
    J = lambda o: json.dumps(o, ensure_ascii=False)
    root = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))

    def make_cfg(name, **api):
        cfg = yaml.safe_load(open(os.path.join(root, "config.yaml"), encoding="utf-8"))
        d = os.path.join(tmp, name)
        cfg["api"].update(base_url="http://x/BidPublicInfoService", sleep_between_calls_sec=0, max_retries=2, num_of_rows=5, use_awards=False, **api)
        cfg["period"] = {"start": "2024-01", "end": "2024-01"}
        cfg["paths"] = {"data_dir": d, "raw_dir": os.path.join(d, "raw"), "cache_db": os.path.join(d, "cache.sqlite"), "text_dir": os.path.join(d, "text"),
                        "files_dir": os.path.join(d, "files"), "out_dir": os.path.join(d, "out"), "excel_name": "x.xlsx"}
        os.makedirs(cfg["paths"]["raw_dir"], exist_ok=True)
        return cfg

    rows = [{"bidNtceNo": f"R{i:03d}", "bidNtceOrd": "000", "bidNtceNm": "가상 공사", "bidNtceDt": "2024-01-03 10:00:00", "presmptPrce": "1"} for i in range(12)]
    state = {"rows": rows, "skip": 7, "calls": 0}

    def server(url, params=None, timeout=None):                       # 조회 기간으로 걸러 돌려주는 가짜 서버(skip: 건수에는 들어가지만 내려 주지 않는 행)
        state["calls"] += 1
        b, e, pg, n = params["inqryBgnDt"], params["inqryEndDt"], int(params["pageNo"]), int(params["numOfRows"])
        hit = [r for r in state["rows"] if b <= r["bidNtceDt"].replace("-", "").replace(" ", "").replace(":", "")[:12] <= e]
        give = [r for i, r in enumerate(hit) if i != state["skip"]]
        return _Resp(J({"response": {"header": {"resultCode": "00"}, "body": {"items": give[(pg - 1) * n: pg * n] or "", "totalCount": len(hit)}}}))

    old_env = os.environ.get("G2B_SERVICE_KEY")
    os.environ["G2B_SERVICE_KEY"] = "k"
    try:
        with mock.patch.object(requests.Session, "get", side_effect=server), mock.patch("time.sleep"):
            # ① 서버가 건수(12)보다 적게(11) 내려 준 달은 실패 — 그 조회의 쪽은 캐시에 남기지 않아, 서버가 정상이 되면 다음 실행에서 끝난다
            cfg = make_cfg("sticky")
            raw = cfg["paths"]["raw_dir"]
            collect.collect_notices(cfg)
            assert not os.path.exists(os.path.join(raw, "notices_done.txt")) and not os.path.exists(os.path.join(raw, "notices_202401.jsonl"))
            state["skip"], c0 = None, state["calls"]
            collect.collect_notices(cfg)
            assert open(os.path.join(raw, "notices_done.txt")).read().split() == ["202401"] and state["calls"] - c0 == 7, state["calls"] - c0   # 3쪽 + 빈 구간 4개
            assert sum(1 for _ in open(os.path.join(raw, "notices_202401.jsonl"), encoding="utf-8")) == 12
            # ② 두 번 줄어도 처음 남긴 .prev(가장 큰 원본)는 그대로, 두 번째는 .prev2
            for keep in (10, 6):
                os.remove(os.path.join(raw, "notices_done.txt"))
                state["rows"] = rows[:keep]
                with mock.patch.object(collect, "date") as d:        # 받은 날이 달라야 캐시를 다시 쓰지 않는다
                    d.today.return_value = __import__("datetime").date(2030, 1, keep)
                    collect.collect_notices(cfg)
            n = lambda fn: sum(1 for _ in open(os.path.join(raw, fn), encoding="utf-8"))
            assert (n("notices_202401.jsonl"), n("notices_202401.jsonl.prev"), n("notices_202401.jsonl.prev2")) == (6, 12, 10)
            state["rows"] = rows
            # ④ max_window_days 가 빈 값(YAML 'max_window_days:')이면 기본 7
            cfg = make_cfg("nullwin")
            cfg["api"]["max_window_days"] = None
            assert collect.make_client(cfg).cfg.max_window_days == 7
        # ③ 면허제한 조회가 모두 요청 조건 오류(08)면 첫 건에서 멈춘다(오류 응답은 캐시에 넣지 않으므로 실행 때마다 공고 수만큼 호출하게 된다)
        e08 = J({"nkoneps.com.response.ResponseError": {"header": {"resultCode": "08", "resultMsg": "필수값 입력 에러"}}})
        cfg = make_cfg("lic")
        with mock.patch.object(requests.Session, "get", return_value=_Resp(e08)) as g, mock.patch("time.sleep"):
            assert collect.collect_license_limits(cfg, [f"R{i}" for i in range(30)]).empty and g.call_count == 1, g.call_count
        with mock.patch.object(requests.Session, "get", return_value=_Resp("{}")) as g, mock.patch("time.sleep"):
            assert collect.collect_license_limits(cfg, [f"R{i}" for i in range(30)]).empty and g.call_count == 10, g.call_count
    finally:
        if old_env is None:
            os.environ.pop("G2B_SERVICE_KEY", None)
        else:
            os.environ["G2B_SERVICE_KEY"] = old_env
    # ⑤ totalCount 음수는 '마지막 쪽을 지난 빈 쪽'이 아니고, 무한대는 예외 없이 '없음'
    assert validate_body({"items": "", "totalCount": -1}, 1, 5) and body_total({"totalCount": "Infinity"}) is None and body_total({"totalCount": 1e999}) is None
    # iter_all: 전량이 아닌 채 끝난 조회의 쪽은 캐시에서 지운다(쪽 자체는 정상이어도)
    c = _client(tmp); c.cfg.num_of_rows = 5
    with mock.patch.object(c.session, "get", side_effect=lambda url, params=None, timeout=None: _Resp(J({"response": {"header": {"resultCode": "00"}, "body": {
            "items": [{"i": 1}] if params["pageNo"] == "1" else "", "totalCount": 3}}}))) as g, mock.patch("time.sleep"):
        for _ in range(2):                                              # 1쪽(1건) + 2쪽(마지막 쪽을 지난 빈 쪽으로 보이는 응답) — 두 번 다 서버에 다시 묻는다
            try:
                list(c.iter_all("op", {"a": "1"}))
                raise AssertionError("IncompleteResult 기대")
            except IncompleteResult:
                pass
        assert g.call_count == 4, g.call_count
    assert c.conn.execute("SELECT COUNT(*) FROM cache").fetchone()[0] == 0
    # ⑥ 구간 경계: 서버가 끝 시각의 1분을 포함하든(23:59:59 까지) 아니든(00초까지), 경계에 걸린 공고가 빠지지도 두 번 저장되지도 않는다
    from tests.mock_g2b import KEY, MockG2B
    srv = MockG2B()
    base = srv.start()
    old_env = {k: os.environ.get(k) for k in ("G2B_SERVICE_KEY", "NO_PROXY")}
    os.environ.update(G2B_SERVICE_KEY=KEY, NO_PROXY="127.0.0.1,localhost")
    try:
        tmpl = srv.data["notices"][0]
        edge = [dict(tmpl, bidNtceNo=f"EDGE{i}", bidNtceOrd="000", bidNtceDt=dt) for i, dt in enumerate((
            "2024-01-06 23:59:30", "2024-01-07 00:00:00", "2024-01-07 00:00:30", "2024-01-12 23:59:59", "2024-01-13 00:00:59", "2024-01-25 00:00:10", "2024-01-31 23:58:59"))]
        srv.data["notices"] = srv.data["notices"] + edge
        want = sorted(f'{n["bidNtceNo"]}-{n["bidNtceOrd"]}' for n in srv.data["notices"] if n["bidNtceDt"].startswith("2024-01"))
        for inclusive in (True, False):
            srv.end_inclusive = inclusive
            cfg = make_cfg(f"edge{int(inclusive)}")
            cfg["api"].update(base_url=base + "/BidPublicInfoService", num_of_rows=999)
            with mock.patch("time.sleep"):
                collect.collect_notices(cfg)
            raw = cfg["paths"]["raw_dir"]
            got = sorted(f'{r["bidNtceNo"]}-{r["bidNtceOrd"]}' for r in map(json.loads, open(os.path.join(raw, "notices_202401.jsonl"), encoding="utf-8")))
            assert got == want, (inclusive, sorted(set(want) - set(got)), [x for x in got if got.count(x) > 1])
            meta = json.load(open(os.path.join(raw, "notices_months.json"), encoding="utf-8"))["202401"]
            # 경계 1분에 걸린 행: 끝 1분을 포함하는 서버는 4건(7일 00:00:00·00:00:30, 13일 00:00:59, 25일 00:00:10)을, 아닌 서버는 정각의 1건만 두 번 준다
            assert meta["saved"] == len(want) and meta["received"] - meta["dup"] == len(want) and meta["dup"] == (4 if inclusive else 1), (inclusive, meta)
    finally:
        srv.stop()
        for k, v in old_env.items():
            if v is None:
                os.environ.pop(k, None)
            else:
                os.environ[k] = v
    # ⑦ 완료 표시 파일을 메모장·PowerShell 로 고쳐 BOM·CRLF·메모가 섞여도 달만 읽는다(첫 달을 다시 받지 않는다)
    dp = os.path.join(tmp, "done_bom.txt")
    open(dp, "wb").write("\ufeff202401\r\n202402\r\n# 메모\r\n2024-03\r\n".encode("utf-8"))
    assert collect._read_done(dp) == {"202401", "202402"}
    open(dp, "wb").write("202401\n202402".encode("utf-16"))
    assert collect._read_done(dp) == {"202401", "202402"}


def test_classify():
    assert extract_facility_name("제주시 문화예술회관 건립공사", "문화예술회관") == "제주시 문화예술회관"
    assert extract_facility_name("○○시 유리온실 건립공사 전기공사", "온실 건립") == "○○시 유리온실"
    assert classify_trade("OO시 문화예술회관 건립공사", "", "건축공사업 / 토목건축공사업")[0] == "건축"
    assert classify_trade("OO소방서 신축공사", "", "")[0] == "건축"
    assert classify_trade("OO미술관 건립 소방시설공사", "", "")[0] == "소방"
    assert classify_notice_kind("일반공고", "Y", "", "가상시 미술관 건립공사(입찰취소 후 재공고)") == "재공고"
    assert classify_notice_kind("취소공고", "N", "", "x") == "취소"
    assert classify_notice_kind("일반공고", "N", "", "OO 건립공사 취소공고") == "취소"
    assert [h[2] for h in match_categories("○○시 야외공연장 조성공사")] == ["야외공연장"]
    # 2026-09 실데이터에서 확인된 오탐·오분류
    assert match_categories("2정수장 고압 간선케이블 교체 전기공사") == [] and match_categories("경일고등학교 학과재구조화(호텔앤리조트과) 실습실 환경개선사업 전기공사") == []
    assert match_categories("경북고등학교 학교체육시설(야구장) 환경개선공사") == [] and match_categories("서천군 유소년 축구장 관리동 건립공사(건축)")
    assert classify_work_type("영구임대주택 승강기 안전장치(부품) 설치공사") == "유지보수"
    assert classify_work_type("정남면 야외공연장 설치공사") == "미분류" and classify_work_type("죽전야외음악당 주차장 및 진입도로 조성공사") == "미분류"
    assert classify_work_type("화성예술의전당 소공연장 조성 건축(기계) 공사 (전체분 및 1차분)") == "신축"
    assert classify_work_type("상주박물관 수장고 증축사업 건축공사(총괄, 1차분)") == "증축"
    assert extract_facility_name("입찰 취소 공고[진해아트홀 시설 개선공사] ", "아트홀") == "진해아트홀"
    assert extract_facility_name("제2안식의 집(봉안당) 건립공사(조경)", "봉안당") == "제2안식의 집"
    assert work_type_for("이현삼거리 서편 지하 공영주차장 조성 소방공사", "지하공영주차장") == "신축"
    assert extract_facility_name("이현삼거리 서편 지하 공영주차장 조성 기계공사", "지하공영주차장") == "이현삼거리 서편 지하 공영주차장"
    assert work_type_for("종합운동장 부설주차장 유료화 대비 조성 공사", "종합운동장") == "미분류"
    assert classify_work_type("홍천군 추모공원 봉안묘 석축 설치 공사") == "미분류"
    assert extract_facility_name("화성예술의전당 소공연장 조성 건축(기계) 공사 (전체분 및 1차분)", "예술의전당") == "화성예술의전당"
    assert classify_trade("x", "기계설비ㆍ가스공사업", "")[0] == "기계설비" and classify_trade("x", "실내건축공사업", "")[0] == "건축"
    assert classify_trade("x", "조경식재ㆍ시설물공사업", "")[0] == "조경" and classify_trade("x", "지반조성ㆍ포장공사업", "")[0] == "토목"
    assert match_categories("○○군 매입임대주택 리모델링") == []
    assert parse_amount(float("nan")) is None and parse_amount("nan") is None and parse_amount("1,000") == 1000
    cfg = load_config()
    assert os.path.isabs(cfg["paths"]["data_dir"]), "paths 는 프로젝트 루트 기준 절대경로여야 함"


def test_discover_and_dedup(tmp):
    cfg = load_config()
    raw = pd.DataFrame([
        dict(bidNtceNo="A1", bidNtceOrd="00", bidNtceNm="가상군 문화예술회관 건립공사", ntceKindNm="일반공고", bidNtceDt="2024-01-05 10:00:00",
             dminsttNm="가상군", presmptPrce="30000000000", mainCnsttyNm=None, ntceSpecDocUrl1=None, ntceSpecFileNm1=None),
        dict(bidNtceNo="A2", bidNtceOrd=0, bidNtceNm="가상군 문화예술회관 건립 전기공사", ntceKindNm="일반공고", bidNtceDt="2024-01-06 10:00:00",
             dminsttNm="가상군", presmptPrce=None, mainCnsttyNm="전기공사"),
        dict(bidNtceNo="A3", bidNtceOrd="00", bidNtceNm="가상군 문화예술회관 관리사무소 건립공사", ntceKindNm="일반공고", bidNtceDt="2025-03-01 10:00:00",
             dminsttNm="가상군", presmptPrce="400000000", mainCnsttyNm="건축공사"),
        dict(bidNtceNo="B1", bidNtceOrd="00", bidNtceNm="다른군 문화예술회관 건립공사", ntceKindNm="일반공고", bidNtceDt="2024-02-05 10:00:00",
             dminsttNm="다른군", presmptPrce="25000000000", mainCnsttyNm="건축공사"),
        dict(bidNtceNo="C1", bidNtceOrd="00", bidNtceNm="문화예술회관 건립공사(1단계)", ntceKindNm="일반공고", bidNtceDt="2023-01-05 10:00:00",
             dminsttNm="셋째군", presmptPrce="5000000000", mainCnsttyNm="건축공사"),
        dict(bidNtceNo="C2", bidNtceOrd="00", bidNtceNm="문화예술회관 건립공사(2단계)", ntceKindNm="일반공고", bidNtceDt="2023-06-05 10:00:00",
             dminsttNm="셋째군", presmptPrce="7000000000", mainCnsttyNm="건축공사"),
        dict(bidNtceNo="C3", bidNtceOrd="00", bidNtceNm="문화예술회관 건립공사 2차 재공고", ntceKindNm="일반공고", reNtceYn="Y", bidNtceDt=None,
             dminsttNm="셋째군", presmptPrce="7100000000", mainCnsttyNm="건축공사"),
    ])
    std = discover.standardize(raw, cfg, None)
    assert std["주공종명"].iloc[0] == "" and std["첨부URL1"].iloc[0] == "" and std["공고차수"].iloc[1] == "000", "None/NaN 은 공란으로, 차수는 3자리"
    # 전량 parquet 는 필요한 컬럼만 읽는다(메모리) — 그 컬럼만으로 standardize 한 결과가 전체 컬럼으로 한 것과 같아야 함
    need = discover.raw_columns_needed(cfg)
    assert {"bidNtceNo", "bidNtceOrd", "bidNtceNm", "presmptPrce", "ntceSpecDocUrl10", "sptDscrptDocUrl5", "subsiCnsttyNm9", "befBidBbancNo"} <= set(need)
    raw_extra = raw.assign(extraField="x")
    a = discover.standardize(raw_extra[[c for c in need if c in raw_extra.columns]], cfg, None)
    b = discover.standardize(raw_extra, cfg, None)
    assert a.equals(b) and "extraField" not in b.columns, "필요 컬럼만 읽어도 동일 결과"
    assert "nan" not in set(std["주공종명"]) and "None" not in set(std["첨부파일명1"])
    agg = discover.discover_candidates(std)
    # 같은 이름이라도 수요기관이 다르면 별도 시설
    assert len(agg) == 3 and set(agg["수요기관"]) == {"가상군", "다른군", "셋째군"}
    # 검수 내용 이어받기(시설ID·검수 컬럼)
    prev = agg.copy()
    prev.loc[prev["수요기관"] == "가상군", "검수_별칭(;구분)"] = "가상군문예회관"
    prev.loc[prev["수요기관"] == "가상군", "시설ID"] = "F0007"
    # 가상군 행의 검수_포함여부는 옛 규칙 기본값 그대로('제외')였다고 가정 → 새 규칙 기본값으로 다시 계산되어야 함
    prev.loc[prev["수요기관"] == "가상군", "사업유형"] = "유지보수"
    prev.loc[prev["수요기관"] == "가상군", "검수_포함여부"] = "제외"
    prev = prev[prev["수요기관"] != "셋째군"]            # 셋째군은 이전 파일에 없던 새 시설 → 이어지는 번호(F0008)
    agg2 = discover.discover_candidates(std, previous_review=prev, recalc=True)
    r = agg2[agg2["수요기관"] == "가상군"].iloc[0]
    assert r["검수_포함여부"] == "포함", "--recalc: 기본값 그대로였던 '제외'는 이어받지 않고 새 규칙(신축 → 포함)으로"
    keep = discover.discover_candidates(std, previous_review=prev)
    assert keep[keep["수요기관"] == "가상군"].iloc[0]["검수_포함여부"] == "제외", "기본 동작(검수 유지): 기본값과 같아 보이는 칸도 적혀 있던 그대로"
    prev3 = agg.copy()
    prev3.loc[prev3["수요기관"] == "다른군", "검수_포함여부"] = "제외"        # 사용자가 기본값(포함)을 고침 → 이어받아야 함
    prev3.loc[prev3["수요기관"] == "다른군", "검수_시설명"] = "다른군 문예회관(별칭)"
    r_b = discover.discover_candidates(std, previous_review=prev3).query("수요기관 == '다른군'").iloc[0]
    assert r_b["검수_포함여부"] == "제외" and r_b["검수_시설명"] == "다른군 문예회관(별칭)", "사용자가 고친 값은 이어받음"
    r_f = discover.discover_candidates(std, previous_review=prev3, ids_only=True).query("수요기관 == '다른군'").iloc[0]
    assert r_f["검수_포함여부"] == "포함" and r_f["검수_시설명"] == r_f["시설명_후보"] and r_f["시설ID"] == r_b["시설ID"], "--fresh: 시설ID만 유지, 검수 값은 기본값"
    assert r["시설ID"] == "F0007" and r["검수_별칭(;구분)"] == "가상군문예회관"
    assert agg2["시설ID"].is_unique and agg2[agg2["수요기관"] == "셋째군"]["시설ID"].iloc[0] == "F0008"
    # 검수 파일 왕복: 빈 검수_시설명 → 시설명_후보 대체, 수요기관 필터
    path = os.path.join(tmp, "facility_candidates.xlsx")
    agg2.loc[agg2["수요기관"] == "가상군", "검수_시설명"] = None
    discover.write_review_workbook(agg2, path)
    reviewed = discover.read_reviewed(path)
    assert set(reviewed["검수_시설명"]) == {"가상군 문화예술회관", "다른군 문화예술회관", "셋째군 문화예술회관"}, \
        ("지자체명 없는 '문화예술회관' 은 수요기관을 접두", set(reviewed["검수_시설명"]))
    hits = discover.research_by_facility(std, reviewed)
    # '문화예술회관'(셋째군)은 이름만으로는 전국 매칭이지만 수요기관 필터로 셋째군 공고만
    third = hits[hits["시설명"] == "셋째군 문화예술회관"]
    assert set(third["공고번호"]) == {"C1", "C2", "C3"}
    assert set(hits[hits["시설명"] == "가상군 문화예술회관"]["공고번호"]) == {"A1", "A2", "A3"}
    hist, latest, logs = dedup.dedup_latest(hits, 0.30)
    pid = latest.set_index("공고번호")["프로젝트ID"].to_dict()
    key = latest.set_index("공고번호")["프로젝트키"].to_dict()
    fid = reviewed.set_index("검수_시설명")["시설ID"]
    # 세부사업 분리: 본공사(A1, 300억)와 전기(A2)는 한 세부사업, 14개월 뒤의 '관리사무소 건립공사'(A3, 4억)는 자기 사업어를 가진 별개 세부사업
    # → 예전처럼 A3 가 A1 을 대체하지도, '소액' 으로 빠지지도 않고 각자 대표(프로젝트ID 끝에 -01, -02)
    f1 = fid["가상군 문화예술회관"]
    assert pid["A1"] == pid["A2"] == f"{f1}-N-01" and pid["A3"] == f"{f1}-N-02", pid
    assert any(l["항목"] == "세부사업 분리" and f1 in l["비고"] for l in logs)
    assert set(hist[hist["공고번호"].isin(["A1", "A2", "A3"])]["세부사업수"]) == {2} and hist[hist["공고번호"] == "A3"]["세부사업명"].iloc[0].startswith("가상군 문화예술회관 관리사무소")
    # (1단계)/(2단계)는 같은 세부사업의 차수별 발주 → 프로젝트키만 다르고 둘 다 합산. 날짜 없는 C3('2차 재공고')는 시기를 알 수 없어 따로 남는다
    f3 = fid["셋째군 문화예술회관"]
    assert pid["C1"] == pid["C2"] != pid["C3"] and key["C1"].endswith("|건축|1단계") and key["C2"].endswith("|건축|2단계")
    assert hist[hist["공고번호"] == "C3"]["최신여부"].iloc[0] and any(l["항목"] == "같은 공종 합산" and pid["C1"] in l["비고"] for l in logs)
    assert set(hist[hist["공고번호"] == "B1"]["프로젝트ID"]) == {f"{fid['다른군 문화예술회관']}-N"} and hist[hist["공고번호"] == "B1"]["세부사업명"].iloc[0] == "", "나뉘지 않은 시설은 번호·이름 없음"
    # 동일 공고번호가 두 시설에 걸려도 한쪽이 '구차수'로 탈락하지 않음
    dup = pd.concat([hits[hits["공고번호"] == "A1"].assign(시설ID="F0099", 시설명="복제")] + [hits], ignore_index=True)
    h2, l2, _ = dedup.dedup_latest(dup)
    assert (h2[h2["공고번호"] == "A1"]["최신차수여부"]).all()
    # 취소 차수: 최신 차수가 취소이면 원 차수는 '취소됨'
    canc = hits[hits["공고번호"] == "B1"].copy()
    canc2 = canc.copy(); canc2["공고차수"], canc2["공고종류"], canc2["공고키"] = "01", "취소", "B1-01"
    h3, l3, lg3 = dedup.dedup_latest(pd.concat([canc, canc2], ignore_index=True))
    assert l3.empty and h3["대표선정사유"].str.contains("취소").all()
    # 빈 입력에서도 facility_summary 스키마 유지
    assert "프로젝트ID" in dedup.facility_summary(l3, h3).columns
    # API 이전공고번호 연결: 프로젝트키가 달라도(단계토큰 유무) 재공고가 이전 공고를 대체
    a = hits[hits["공고번호"] == "A1"].copy(); a["공고명"] = "가상군 문화예술회관 건립공사(1단계)"; a["이전공고번호"] = ""
    b = hits[hits["공고번호"] == "A1"].copy(); b["공고번호"], b["공고키"], b["공고명"], b["이전공고번호"], b["공고일시"] = "A9", "A9-00", "가상군 문화예술회관 건립공사 재공고", "A1", "2024-06-01 10:00:00"
    h4, l4, lg4 = dedup.dedup_latest(pd.concat([a, b], ignore_index=True))
    assert set(l4["공고번호"]) == {"A9"} and h4[h4["공고번호"] == "A1"]["대체공고번호"].iloc[0] == "A9"


def test_discover_keeps_review(tmp):
    """검수를 마친 뒤 discover 를 다시 돌려도(공고가 늘었거나 이름 정리 규칙이 바뀌었어도) 이전 행의 시설ID·검수 칸이 그대로 남는다."""
    from g2b_costdb import pipeline
    cfg = load_config()
    n = lambda no, nm, dt, org, mc="건축공사": dict(bidNtceNo=no, bidNtceOrd="00", bidNtceNm=nm, ntceKindNm="일반공고", bidNtceDt=dt, dminsttNm=org,
                                                presmptPrce="1000000000", mainCnsttyNm=mc)
    rows1 = [n("A1", "가상군 문화예술회관 건립공사", "2024-01-05 10:00:00", "가상군"), n("B1", "다른군 문화예술회관 건립공사", "2024-02-05 10:00:00", "다른군"),
             n("M1", "셋째군 문화예술회관 방수공사", "2024-03-01 10:00:00", "셋째군")]
    rows2 = rows1 + [n("M2", "셋째군 문화예술회관 증축공사", "2026-09-20 10:00:00", "셋째군"), n("N1", "넷째군 문화예술회관 건립공사", "2026-09-21 10:00:00", "넷째군")]
    std1, std2 = discover.standardize(pd.DataFrame(rows1), cfg, None), discover.standardize(pd.DataFrame(rows2), cfg, None)
    agg0 = discover.discover_candidates(std1)
    assert list(agg0.columns).index(discover.STATUS_COL) == list(agg0.columns).index("검수_포함여부") - 1 and set(agg0[discover.STATUS_COL]) == {""}, "표시 열은 노란 칸 바로 앞, 처음 만들 때는 빈칸"
    by = lambda df, org: df[df["수요기관"] == org].iloc[0]
    assert by(agg0, "셋째군")["사업유형"] == "유지보수" and by(agg0, "셋째군")["검수_포함여부"] == "제외"
    prev = agg0.copy()
    ix = lambda org: prev.index[prev["수요기관"] == org][0]
    prev.loc[ix("가상군"), ["시설ID", "검수_수요기관", "검수_별칭(;구분)"]] = ["F0100", None, "가상문예"]            # 수요기관을 비움(이름만으로 찾기) — 기본값으로 되돌리면 안 된다
    old_key = prev.at[ix("다른군"), "시설키"]
    prev.loc[ix("다른군"), ["시설ID", "시설키", "시설명_후보", "검수_시설명", "검수_메모"]] = ["F0002", old_key + "옛", "다른군 문화예술회관 옛", "다른군 문화예술회관 옛", "확인함"]   # 옛 규칙의 시설키
    prev.loc[ix("셋째군"), "시설ID"] = "F0003"
    gone = dict(prev.iloc[0]); gone.update({"시설ID": "F0200", "시설키": "없어진시설", "시설명_후보": "없어진 시설", "수요기관": "없는군", "공고명_예시": None,
                                           "검수_포함여부": "포함", "검수_시설명": "없어진 시설", "검수_별칭(;구분)": None, "검수_수요기관": "없는군", "검수_메모": "직접 넣음"})
    prev = pd.concat([prev, pd.DataFrame([gone, {c: None for c in prev.columns}])], ignore_index=True)        # 이번 후보에 없는 행 + 통째로 빈 줄
    path = os.path.join(tmp, "keep", "facility_candidates.xlsx")
    discover.write_review_workbook(prev, path)
    assert not os.path.exists(path + ".tmp"), "임시 파일을 남기지 않는다"
    prev_x = discover.read_review_file(path)                                                                # 엑셀을 거친 값으로 시험(빈칸은 None)
    out = discover.discover_candidates(std2, previous_review=prev_x)
    st = lambda org: by(out, org)[discover.STATUS_COL]
    a, b, m, nw, g = by(out, "가상군"), by(out, "다른군"), by(out, "셋째군"), by(out, "넷째군"), by(out, "없는군")
    assert (a["시설ID"], a["검수_수요기관"], a["검수_별칭(;구분)"], a["검수_포함여부"], st("가상군")) == ("F0100", "", "가상문예", "포함", ""), dict(a)
    assert b["시설ID"] == "F0002" and b["시설키"] == old_key and b["검수_시설명"] == "다른군 문화예술회관 옛" and b["검수_메모"] == "확인함" \
        and st("다른군").startswith("후보 이름 바뀜: '다른군 문화예술회관 옛' → '다른군 문화예술회관'"), ("시설키가 바뀐 행은 예시 공고명으로 맞춘다", dict(b))
    assert m["시설ID"] == "F0003" and m["사업유형"] == "증축" and m["검수_포함여부"] == "제외" and "확인: 사업유형 유지보수→증축, 새 기본값 '포함'(지금 '제외' 그대로 둠)" in st("셋째군"), dict(m)
    assert nw["시설ID"] == "F0201" and st("넷째군") == "신규" and nw["검수_포함여부"] == "포함", ("새 시설의 번호는 이전 파일의 모든 번호 뒤(남겨 둔 F0200 다음)", dict(nw))
    assert g["시설ID"] == "F0200" and g["검수_포함여부"] == "포함" and g["검수_메모"] == "직접 넣음" and st("없는군").startswith("이번 후보에 없음") and out.index[out["수요기관"] == "없는군"][0] == len(out) - 1
    assert len(out) == 5 and out["시설ID"].is_unique, out[["시설ID", "수요기관", discover.STATUS_COL]]
    discover.write_review_workbook(out, path)
    rv = discover.read_reviewed(path)
    assert sorted(rv["시설ID"]) == ["F0002", "F0100", "F0200", "F0201"] and rv.set_index("시설ID").loc["F0100", "검수_수요기관"] in ("", None), rv[["시설ID", "검수_수요기관"]]
    again = discover.discover_candidates(std2, previous_review=discover.read_review_file(path))            # 한 번 더 돌려도 그대로. 아직 확인하지 않았을 수 있는 표시는 '지난번: ' 을 붙여 남긴다
    key = ["시설ID"] + discover.REVIEW_COLS
    st2 = lambda org: by(again, org)[discover.STATUS_COL]
    assert again[key].equals(out[key]) and st2("넷째군") == "지난번: 신규" and st2("셋째군").startswith("지난번: 확인: 사업유형 유지보수→증축") \
        and st2("다른군") == "" and st2("가상군") == "" and st2("없는군").startswith("이번 후보에 없음"), again[key + [discover.STATUS_COL]]
    discover.write_review_workbook(again, path)
    third = discover.read_review_file(path)
    third.loc[third["수요기관"] == "넷째군", "검수_메모"] = "확인함"                                       # 새 시설을 보고 메모를 남김 → 표시가 사라진다
    third.loc[third["수요기관"] == "셋째군", "검수_포함여부"] = "포함"                                     # 확인 대상의 값을 고침 → 표시가 사라진다
    t3 = discover.discover_candidates(std2, previous_review=third)
    assert by(t3, "넷째군")[discover.STATUS_COL] == "" and by(t3, "셋째군")[discover.STATUS_COL] == "" and by(t3, "넷째군")["검수_메모"] == "확인함"
    t4 = discover.discover_candidates(std2, previous_review=discover.read_review_file(path))               # 손대지 않고 또 돌리면 표시는 그대로(겹쳐 붙지 않는다)
    assert by(t4, "넷째군")[discover.STATUS_COL] == "지난번: 신규" and by(t4, "셋째군")[discover.STATUS_COL] == st2("셋째군")
    discover.write_review_workbook(out, path)
    # --recalc(검수 전 단계용 옛 동작)과의 차이: 손대지 않은 칸을 다시 계산하고, 이번 후보에 없는 행은 싣지 않는다
    rc = discover.discover_candidates(std2, previous_review=prev_x, recalc=True)
    assert by(rc, "셋째군")["검수_포함여부"] == "포함" and by(rc, "가상군")["검수_수요기관"] == "가상군" and len(rc) == 4 and set(rc[discover.STATUS_COL]) == {""}
    # 시설키·시설ID 열이 없는 옛 형식 파일은 맞출 기준이 없다 → 새로 만든다(이전 행을 덧붙여 두 배로 만들지 않는다)
    assert len(discover.discover_candidates(std2, previous_review=prev_x.drop(columns=["시설키"]))) == 4
    # 머리글이 망가진 파일(열 삭제·이름 변경, 첫 줄 위에 줄 삽입)은 이어받을 수 없다 — 까닭을 알려 주고, discover 단계는 검수 파일을 덮어쓰지 않고 멈춘다
    assert discover.review_file_problem(prev_x) == "" and discover.review_file_problem(None) == "" and discover.review_file_problem(prev_x.iloc[0:0]) == ""
    for bad in (prev_x.drop(columns=["시설키"]), prev_x.rename(columns={"시설ID": "ID"}), prev_x.drop(columns=["수요기관"]),
                pd.DataFrame([list(prev_x.columns)] + prev_x.values.tolist(), columns=["2026년 검수본"] + ["None"] * (len(prev_x.columns) - 1))):
        assert "열이 없습니다" in discover.review_file_problem(bad), list(bad.columns)
    own = prev_x.copy()
    own["None"] = [None] * len(own)                                                                       # 서식만 있는 빈 열은 괜찮다
    assert discover.review_file_problem(own) == ""
    own.loc[own.index[0], "None"] = "따로 적은 글"                                                          # 이름 없는 열에 적은 글은 옮길 수 없다 → 말없이 빼지 않고 알린다
    assert "이름이 없는 열에 글이 1칸" in discover.review_file_problem(own), discover.review_file_problem(own)
    cfg3 = load_config()
    cfg3["paths"] = dict(cfg3["paths"], data_dir=os.path.join(tmp, "keep", "data3"), raw_dir=os.path.join(tmp, "keep", "raw3"), out_dir=os.path.join(tmp, "keep", "out3"))
    cfg3["period"] = dict(cfg3["period"], start="2024-01", end="2024-01")
    for k in ("data_dir", "raw_dir", "out_dir"):
        os.makedirs(cfg3["paths"][k], exist_ok=True)
    pd.DataFrame(rows2).to_parquet(os.path.join(cfg3["paths"]["data_dir"], "notices_all.parquet"))
    path3 = os.path.join(cfg3["paths"]["out_dir"], "facility_candidates.xlsx")
    discover.write_review_workbook(prev_x.rename(columns={"시설ID": "ID"}), path3)
    before3 = open(path3, "rb").read()
    try:
        pipeline.stage_discover(cfg3)
        raise AssertionError("머리글이 망가진 검수 파일은 덮어쓰지 않고 멈춰야 한다")
    except SystemExit as e:
        assert "[중단]" in str(e) and "'시설ID'" in str(e), str(e)
    assert open(path3, "rb").read() == before3 and os.listdir(cfg3["paths"]["out_dir"]) == ["facility_candidates.xlsx"], "검수 파일은 그대로, 사본도 만들지 않는다"
    # ── _keep_review 단위: 같은 (시설키, 수요기관)이 검색어 둘에 걸린 시설, 예전 공고로 새로 잡힌 후보
    mk = lambda w, k, org, nm, first, last, typ="신축", inc="포함": {"표2_대분류": "X", "표2_중분류": "x", "검색어": w, "시설키": k, "시설명_후보": nm, "수요기관": org, "공고건수": 1,
                                                              "최초공고일": first, "최종공고일": last, "사업유형": typ, "공고명_예시": nm + " 건립공사",
                                                              "검수_포함여부": inc, "검수_시설명": nm, "검수_별칭(;구분)": "", "검수_수요기관": org, "검수_메모": ""}
    pv = pd.DataFrame([dict(mk("k1", "K", "O", "케이", "2024-01-01", "2024-02-01", inc="제외"), 시설ID="F0001"), dict(mk("k2", "K", "O", "케이", "2024-01-01", "2024-02-01"), 시설ID="F0001"),
                       dict(mk("추모공원", "의령4·26추모공원", "의령군", "의령4·26추모공원", "2023-12-13", "2024-07-22"), 시설ID="F0002")])
    ag = pd.DataFrame([mk("k1", "K", "O", "케이", "2024-01-01", "2024-02-01"), mk("k2", "K", "O", "케이", "2024-01-01", "2024-02-01"), mk("k3", "K", "O", "케이", "2026-09-20", "2026-09-20"),
                       mk("추모공원", "의령4·26추모공원", "의령군", "의령4·26추모공원", "2023-12-13", "2024-07-22"), mk("추모공원", "의령추모공원", "의령군", "의령 추모공원", "2024-02-06", "2024-02-06", typ="미분류", inc="검토필요"),
                       mk("추모공원", "전혀다른곳", "의령군", "전혀 다른 곳", "2026-09-25", "2026-09-25")])
    kr = discover._keep_review(ag, pv)
    assert list(kr["검수_포함여부"][:3]) == ["제외", "포함", "제외"] and list(kr["시설ID"][:3]) == ["F0001"] * 3, ("검색어가 다른 두 행은 각자 자기 값을, 새 검색어 행은 그 시설의 값을", kr[["시설ID", "검색어", "검수_포함여부"]])
    assert kr[discover.STATUS_COL][2].startswith("신규 행(같은 시설 F0001") and kr[discover.STATUS_COL][0] == kr[discover.STATUS_COL][1] == ""
    assert kr["시설ID"][4] == "F0003" and "예전 공고가 있음" in kr[discover.STATUS_COL][4] and "비슷한 기존 시설: F0002 '의령4·26추모공원'" in kr[discover.STATUS_COL][4], kr[discover.STATUS_COL][4]
    assert kr["시설ID"][5] == "F0004" and kr[discover.STATUS_COL][5] == "신규", kr[discover.STATUS_COL][5]
    # 같은 시설ID 로 합친 두 행은 이전 파일의 순서를 지킨다(재검색은 묶음의 첫 행에서 시설명·분류를 가져온다) — 검수자가 정렬해 저장한 경우
    pv2 = pd.DataFrame([dict(mk("z9", "ZZ", "O", "둘째 이름", "2024-01-01", "2024-02-01"), 시설ID="F0007"), dict(mk("a1", "AA", "O", "첫째 이름", "2024-01-01", "2024-02-01"), 시설ID="F0007")])
    ag2 = pd.DataFrame([mk("a1", "AA", "O", "첫째 이름", "2024-01-01", "2024-02-01"), mk("a0", "AA", "O", "첫째 이름", "2026-09-20", "2026-09-20"), mk("z9", "ZZ", "O", "둘째 이름", "2024-01-01", "2024-02-01")])
    kr2 = discover._keep_review(ag2, pv2)
    pv2["내확인"] = ["봤음", None]                                                                              # 검수자가 직접 더한 열
    assert list(discover._keep_review(ag2, pv2)["내확인"].map(lambda v: "" if discover._blank(v) else v)) == ["봤음", "", ""], "검수자가 더한 열은 그 행에 그대로"
    assert list(kr2["검색어"]) == ["z9", "a1", "a0"] and list(kr2["시설ID"]) == ["F0007"] * 3 and kr2[discover.STATUS_COL][2].startswith("신규 행("), kr2[["시설ID", "검색어", discover.STATUS_COL]]
    # ── _keep_review 단위(이어서): 맞추는 순서 ①②③ 과 표시 조건을 하나씩 고정한다
    from g2b_costdb.classify import extract_facility_name, normalize_facility_key, prefix_institution
    ST, d1, d2 = discover.STATUS_COL, "2024-01-01", "2024-02-01"
    key_of = lambda title, w, org: normalize_facility_key(prefix_institution(extract_facility_name(title, w), w, org))
    row = lambda sid, w, k, org, nm, **kw: dict(mk(w, k, org, nm, d1, d2), 시설ID=sid, **kw)
    # ① 같은 (시설키, 수요기관)의 두 검색어 행이 이전 파일에 다른 순서로 있어도 각자 자기 검색어 행에 붙고, ② 검색어만 바뀐 행은 시설키로 붙는다
    k2 = discover._keep_review(pd.DataFrame([mk("k1", "K", "O", "케이", d1, d2), mk("k2", "K", "O", "케이", d1, d2), mk("새검색어", "L", "O", "엘", d1, d2)]),
                               pd.DataFrame([row("F0012", "k2", "K", "O", "케이"), row("F0011", "k1", "K", "O", "케이", 검수_포함여부="제외"), row("F0013", "옛검색어", "L", "O", "엘", 검수_포함여부="제외")]))
    assert list(k2["시설ID"]) == ["F0011", "F0012", "F0013"] and list(k2["검수_포함여부"]) == ["제외", "포함", "제외"] and set(k2[ST]) == {""}, k2[["시설ID", "검색어", "검수_포함여부", ST]]
    # 사람이 한 줄을 복사해 두 줄로 만든 행(묶음 열쇠가 같음), 노란 칸만 적어 넣은 행, 번호가 빈 행 — 어느 것도 잃지 않고, 앞뒤 공백은 떼며, 번호는 겹치지 않는다
    pv3 = pd.DataFrame([row(" F0021 ", "k", "K", "O", "케이", 검수_포함여부=" 제외 "), row("F0022", "k", "K", "O", "케이", 검수_시설명="케이 B단지", 검수_메모="복사한 줄"),
                        {"검수_포함여부": "포함", "검수_시설명": "손으로 넣은 시설"}, row(None, "k", "M", "오 ", "엠", 검수_메모="번호 없음"), row(" F0050 ", "k", "없어진키", "O", "없어진 곳", 공고명_예시=None)])
    k3 = discover._keep_review(pd.DataFrame([mk("k", "K", "O", "케이", d1, d2), mk("k", "M", "오 ", "엠", d1, d2), mk("k", "N", "O", "엔", "2026-09-01", "2026-09-01")]), pv3)
    got = sorted(zip(k3["시설ID"], k3["검수_포함여부"], k3["검수_시설명"], k3["검수_메모"]))
    assert got == sorted([("F0021", "제외", "케이", ""), ("F0022", "포함", "케이 B단지", "복사한 줄"), ("", "포함", "손으로 넣은 시설", ""), ("F0051", "포함", "엠", "번호 없음"),
                          ("F0052", "포함", "엔", ""), ("F0050", "포함", "없어진 곳", "")]), got
    assert k3[ST][1] == "" and k3[ST][2] == "신규" and list(k3[ST][3:]) == ["이번 후보에 없음(이전 내용 그대로 둠)"] * 3, ("수요기관은 공백까지 같은 글자로 맞춘다(전량 표와 같은 방식)", list(k3[ST]))
    # ③ 예시 공고명으로 다시 낸 시설키: 지자체명 접두까지 지금 규칙대로 / 이미 쓴 이전 행은 다시 쓰지 않는다 / 두 이전 행이 한 새 행에 겹치면 하나는 끝에 남긴다 / 같은 검색어 행 먼저
    t1, t2 = "문화예술회관 건립공사", "가나다 문화회관 건립공사"
    g1, g2 = key_of(t1, "문화예술회관", "넷째군"), key_of(t2, "문화회관", "O")
    assert g1 == "넷째군문화예술회관" and g2 == "가나다문화회관", (g1, g2)
    k4 = discover._keep_review(pd.DataFrame([mk("문화예술회관", g1, "넷째군", "넷째군 문화예술회관", d1, d2), mk("문화회관", "K", "O", "케이", d1, d2), mk("문화회관", g2, "O", "가나다 문화회관", d1, d2)]),
                               pd.DataFrame([row("F0031", "문화예술회관", "옛키", "넷째군", "문화예술회관", 공고명_예시=t1, 검수_포함여부="제외"), row("F0032", "문화회관", "K", "O", "케이", 공고명_예시=t2),
                                             row("F0033", "문화회관", "옛1", "O", "가나다 문화회관", 공고명_예시=t2, 검수_포함여부="제외"), row("F0034", "문화회관", "옛2", "O", "가나다 문화회관", 공고명_예시=t2)]))
    assert list(k4["시설ID"]) == ["F0031", "F0032", "F0033", "F0034"] and list(k4["검수_포함여부"]) == ["제외", "포함", "제외", "포함"], k4[["시설ID", "검색어", "시설키", "검수_포함여부", ST]]
    assert k4[ST][0].startswith("후보 이름 바뀜: '문화예술회관' → '넷째군 문화예술회관'") and k4[ST][1] == "" and k4[ST][2] == "시설키 바뀜(이름 정리 규칙)" and k4[ST][3].startswith("이번 후보에 없음"), list(k4[ST])
    k4b = discover._keep_review(pd.DataFrame([mk("가나다", g2, "O", "가나다 문화회관", d1, d2), mk("문화회관", g2, "O", "가나다 문화회관", d1, d2)]),
                                pd.DataFrame([row("F0035", "문화회관", "옛3", "O", "가나다 문화회관", 공고명_예시=t2, 검수_포함여부="제외")]))
    st4 = dict(zip(k4b["검색어"], k4b[ST]))
    assert list(k4b["시설ID"]) == ["F0035"] * 2 and st4["가나다"].startswith("신규 행(같은 시설 F0035") and st4["문화회관"] == "시설키 바뀜(이름 정리 규칙)" and list(k4b["검수_포함여부"]) == ["제외"] * 2, st4
    # 검수_시설명이 빈칸(= 시설명_후보를 씀)이던 행: 후보 이름이 달라지면 쓰이던 이름을 적어 두고, 같으면 빈칸 그대로
    k5 = discover._keep_review(pd.DataFrame([mk("k", "K", "O", "케이 새이름", d1, d2), mk("k", "L", "O", "엘", d1, d2)]),
                               pd.DataFrame([row("F0041", "k", "K", "O", "케이 옛이름", 검수_시설명=None), row("F0042", "k", "L", "O", "엘", 검수_시설명=None)]))
    assert list(k5["검수_시설명"]) == ["케이 옛이름", ""] and set(k5[ST]) == {""}, k5[["시설ID", "검수_시설명", ST]]
    # '확인' 은 기본값이 달라지고(유형만 바뀐 것은 아님) 지금 값이 새 기본값과 다를 때만. 이전 파일에 사업유형이 없으면 표시하지 않는다
    chk = lambda old_t, new_t, kept: discover._keep_review(pd.DataFrame([mk("k", "K", "O", "케이", d1, d2, typ=new_t)]), pd.DataFrame([row("F0001", "k", "K", "O", "케이", 사업유형=old_t, 검수_포함여부=kept)]))[ST][0]
    assert chk("유지보수", "증축", "제외").startswith("확인: 사업유형 유지보수→증축") and chk("미분류", "신축", "검토필요").startswith("확인:")
    assert chk("신축", "증축", "제외") == "" and chk("유지보수", "증축", "포함") == "" and chk(None, "신축", "제외") == "", "기본값이 같거나, 지금 값이 새 기본값과 같거나, 이전 유형을 모르면 표시 없음"
    # 비슷한 기존 시설: 한쪽 이름이 다른 쪽에 통째로 들어 있으면(3자 이상) 길이 차이가 커도 알린다 / 예전 공고 판정은 이전 파일의 가장 늦은 공고일까지(같은 시각 포함)
    k6 = discover._keep_review(pd.DataFrame([mk("k", "가나다라마바사아자차문화회관", "O", "긴 이름", d2, d2), mk("k", "문화회관", "O", "문화회관", d1, d2)]), pd.DataFrame([row("F0061", "k", "문화회관", "O", "문화회관")]))
    assert k6[ST][0].startswith("신규(예전 공고가 있음") and "비슷한 기존 시설: F0061 '문화회관'" in k6[ST][0], k6[ST][0]
    # 시설ID 열이 없는 파일도 옛 형식으로 본다(새로 만든다)
    no_id = discover.discover_candidates(std2, previous_review=prev_x.drop(columns=["시설ID"]))
    assert len(no_id) == 4 and set(no_id[discover.STATUS_COL]) == {""}
    # ── 파이프라인: 덮어쓰기 전 사본, 전량 표가 수집분보다 적을 때의 알림
    b1 = pipeline._backup_review_file(path)
    b2 = pipeline._backup_review_file(path)
    assert b1 and b2 and b1 != b2 and os.path.exists(b1) and os.path.exists(b2) and os.path.getsize(b1) == os.path.getsize(path), (b1, b2)
    assert pipeline._backup_review_file(os.path.join(tmp, "keep", "없는파일.xlsx")) is None
    d = os.path.join(tmp, "keep", "data"); os.makedirs(d, exist_ok=True)
    cfg2 = {"paths": {"data_dir": d}}
    assert pipeline._std_behind_collect(cfg2) == 0, "파일이 없으면 조용히 0"
    pd.DataFrame({"a": range(7)}).to_parquet(os.path.join(d, "notices_all.parquet")); pd.DataFrame({"a": range(5)}).to_parquet(os.path.join(d, "notices_std.parquet"))
    assert pipeline._std_behind_collect(cfg2) == 2
    # 건수가 같아도 collect 가 discover 보다 나중이면(뒤늦게 올라온 기초금액·다시 받은 달) 알린다
    t_std = os.path.getmtime(os.path.join(d, "notices_std.parquet"))
    assert not pipeline._collect_after_discover({"paths": {"data_dir": os.path.join(tmp, "keep", "없는폴더")}})
    pd.DataFrame({"a": [1]}).to_parquet(os.path.join(d, "bsis_all.parquet"))
    os.utime(os.path.join(d, "notices_all.parquet"), (t_std - 60, t_std - 60)); os.utime(os.path.join(d, "bsis_all.parquet"), (t_std - 60, t_std - 60))
    assert not pipeline._collect_after_discover(cfg2), "discover 가 나중이면 조용히"
    os.utime(os.path.join(d, "bsis_all.parquet"), (t_std + 60, t_std + 60))
    assert pipeline._collect_after_discover(cfg2)
    # ── discover 단계의 멈춤: 이어받을 수 없는 파일·열려 있는 파일·쓰기 실패 — 어느 경우에도 검수 파일은 그대로다
    def stops(msg_part, **kw):
        before = open(path3, "rb").read()
        try:
            pipeline.stage_discover(cfg3, **kw)
            raise AssertionError("멈춰야 한다: " + msg_part)
        except SystemExit as e:
            assert "[중단]" in str(e) and msg_part in str(e), str(e)
        assert open(path3, "rb").read() == before, "검수 파일은 그대로"
    dup = prev_x.copy()
    dup.insert(len(dup.columns), "검수_메모_", dup["검수_메모"])
    dup.columns = [c.rstrip("_") for c in dup.columns]                                                    # 열을 복사해 같은 이름이 둘
    discover.write_review_workbook(dup, path3)
    stops("'검수_메모' 열이 두 번 있습니다")
    discover.write_review_workbook(prev_x.rename(columns={"검수_메모": "메모"}), path3)                 # 검수 열 이름을 고침 → 그 열을 말없이 비우지 않는다
    stops("'검수_메모' 열이 없습니다")
    stops("'검수_메모' 열이 없습니다", fresh=True)
    assert os.listdir(cfg3["paths"]["out_dir"]) == ["facility_candidates.xlsx"], "멈춘 실행은 사본을 만들지 않는다"
    discover.write_review_workbook(prev_x, path3)
    real_open = open
    def locked(p, mode="r", *a, **k):                                                                      # Excel 이 열고 있는 파일: 쓰기용으로 열 수 없다
        if str(p) == path3 and "+" in mode:
            raise PermissionError(13, "Permission denied")
        return real_open(p, mode, *a, **k)
    pipeline.open = locked
    try:
        stops("Excel 에서 이 파일이 열려 있으면 닫은 뒤")
    finally:
        del pipeline.open
    assert os.listdir(cfg3["paths"]["out_dir"]) == ["facility_candidates.xlsx"] and not os.path.exists(os.path.join(cfg3["paths"]["data_dir"], "notices_std.parquet")), \
        "열려 있는 파일은 계산을 시작하기 전에 알린다(사본도, 전량 표도 만들지 않는다)"
    real_replace = discover.os.replace
    def full(src, dst):
        raise OSError(28, "No space left on device")
    discover.os.replace = full
    try:
        stops("검수 파일을 새로 쓰지 못했습니다")
    finally:
        discover.os.replace = real_replace
    left = sorted(os.listdir(cfg3["paths"]["out_dir"]))
    assert len(left) == 2 and left[0] == "facility_candidates.xlsx" and ".이전_" in left[1] and not any(x.endswith(".tmp") for x in left), left
    pipeline.stage_discover(cfg3)                                                                         # 멈춘 뒤 그냥 다시 돌리면 이어받는다
    done = discover.read_review_file(path3)
    tx = lambda v: "" if discover._blank(v) else str(v)
    rv3 = lambda df: sorted((tx(a), tx(b), tx(c)) for a, b, c in zip(df["시설ID"], df["검수_포함여부"], df["검수_메모"]) if tx(a))
    assert set(rv3(done)) >= set(rv3(prev_x)) and len(rv3(prev_x)) == 4 and len(done) == 5, (rv3(done), rv3(prev_x))
    # 사본을 쓰다 실패하면 쓰다 만 파일을 '사본' 이름으로 남기지 않는다
    real_copy = pipeline.shutil.copy2
    def half(src, dst):
        real_open(dst, "wb").write(b"half")
        raise OSError(28, "No space left on device")
    pipeline.shutil.copy2 = half
    n_before = len(os.listdir(cfg3["paths"]["out_dir"]))
    try:
        pipeline._backup_review_file(path3)
        raise AssertionError("사본 실패는 올려 보내야 한다")
    except OSError:
        pass
    finally:
        pipeline.shutil.copy2 = real_copy
    assert len(os.listdir(cfg3["paths"]["out_dir"])) == n_before
    # 번호 칸에 날짜 같은 긴 숫자가 들어 있어도 새 번호는 평소 번호 뒤에 이어진다
    odd = prev_x.copy()
    odd.loc[odd["시설ID"] == "F0001", "시설ID"] = "20240105000000"
    new_ids = set(discover.discover_candidates(std2, previous_review=odd)["시설ID"]) - set(odd["시설ID"])
    assert new_ids and all(len(x) == 5 for x in new_ids), new_ids


def test_discover_keeps_merged_row_order():
    """같은 시설ID 로 합친 행들의 순서는 discover 를 다시 돌려도 이전 파일 그대로 — 재검색이 묶음의 첫 행에서 읽는 대표 이름·표2 분류·검색어가 바뀌지 않는다."""
    mk = lambda t2, w, k, nm, inc="포함": {"표2_대분류": t2, "표2_중분류": t2, "검색어": w, "시설키": k, "시설명_후보": nm, "수요기관": "O", "공고건수": 1, "최초공고일": "2024-01-01",
                                         "최종공고일": "2024-02-01", "사업유형": "신축", "공고명_예시": nm + " 건립공사", "검수_포함여부": inc, "검수_시설명": nm, "검수_별칭(;구분)": "",
                                         "검수_수요기관": "O", "검수_메모": ""}
    ag = pd.DataFrame([mk("가", "공연장", "A", "에이 공연장"), mk("가", "공연장", "D", "디 공연장"), mk("나", "주차장", "B", "비 주차장"), mk("나", "주차장", "C", "씨 주차장")])
    # 검수자가 B 를 대표로 두고 A 와 없어진 Z 를 같은 번호로 합친 뒤 줄 순서를 바꿔 둔 파일(Z, B, C, A 순). D 는 이번에 새로 잡힌, B 와 같은 시설의 다른 검색어 행
    pv = pd.DataFrame([dict(mk("다", "옛", "Z", "지 옛시설"), 시설ID="F0002", 공고명_예시=None), dict(mk("나", "주차장", "B", "비 주차장"), 시설ID="F0002"),
                       dict(mk("나", "주차장", "C", "씨 주차장"), 시설ID="F0003"), dict(mk("가", "공연장", "A", "에이 공연장"), 시설ID="F0002")])
    ag.loc[1, "시설키"] = "B"                                               # ('공연장', B) = F0002 의 다른 검색어 행
    out = discover._keep_review(ag, pv)
    assert list(out["시설ID"]) == ["F0002", "F0002", "F0002", "F0002", "F0003"], list(out["시설ID"])
    assert list(out["검수_시설명"]) == ["지 옛시설", "비 주차장", "에이 공연장", "비 주차장", "씨 주차장"], ("합친 묶음은 첫 행 자리에 모아 이전 파일 순서대로(Z, B, A), 새 검색어 행은 묶음의 뒤", list(out["검수_시설명"]))
    assert out[discover.STATUS_COL][0].startswith("이번 후보에 없음") and out[discover.STATUS_COL][3].startswith("신규 행(같은 시설 F0002") and list(out["검색어"][:4]) == ["옛", "주차장", "공연장", "공연장"]
    again = discover._keep_review(ag, out)                                # 한 번 더 돌려도 줄 순서까지 그대로
    assert list(again["검수_시설명"]) == list(out["검수_시설명"]) and list(again["시설ID"]) == list(out["시설ID"]) and list(again["검색어"]) == list(out["검색어"])
    solo = discover._keep_review(ag, pv.assign(시설ID=["F0009", "F0002", "F0003", "F0008"]))   # 번호가 겹치지 않으면 줄 순서는 후보 순서 그대로, 없어진 행은 끝
    assert list(solo["시설키"]) == ["A", "B", "B", "C", "Z"] and list(solo["시설ID"]) == ["F0008", "F0002", "F0002", "F0003", "F0009"], solo[["시설ID", "시설키"]]


def test_institution_prefix():
    """단지명 없는 시설명은 수요기관 지자체명을 접두하고, 재검색은 접두어를 뗀 패턴으로도 공고를 찾는다."""
    from g2b_costdb.classify import is_generic_facility_name, prefix_institution, short_institution, strip_institution_prefix
    assert short_institution("전남광주통합특별시 구례군") == "구례군" and short_institution("부산광역시 상수도사업본부 동래통합사업소") == "부산광역시"
    assert short_institution("서울주택도시개발공사") == "서울주택도시개발공사" and short_institution("경상북도교육청 한국생명과학고등학교") == "경상북도교육청"
    assert short_institution("(주)강원랜드") == "강원랜드" and short_institution("(재)경기문화재단") == "경기문화재단"
    assert is_generic_facility_name("영구임대아파트", "영구임대") and is_generic_facility_name("공공주택지구", "공공주택") and is_generic_facility_name("종합운동장", "종합운동장")
    assert not is_generic_facility_name("공덕동 행복주택", "행복주택") and not is_generic_facility_name("마음에온 일도1차 통합공공임대주택", "공공임대주택")
    assert prefix_institution("영구임대아파트", "영구임대", "대구도시개발공사") == "대구도시개발공사 영구임대아파트"
    assert prefix_institution("잠실야구장", "야구장", "서울특별시 체육시설관리사업소") == "잠실야구장"
    assert strip_institution_prefix("구례군 국민임대아파트", "전남광주통합특별시 구례군") == "국민임대아파트"
    cfg = load_config()
    raw = pd.DataFrame([
        dict(bidNtceNo="G1", bidNtceOrd="000", bidNtceNm="국민임대아파트 건립사업", ntceKindNm="일반공고", bidNtceDt="2024-01-05 10:00:00",
             dminsttNm="전남광주통합특별시 구례군", presmptPrce="20000000000", mainCnsttyNm="건축공사"),
        dict(bidNtceNo="G2", bidNtceOrd="000", bidNtceNm="국민임대아파트 건립 전기공사", ntceKindNm="일반공고", bidNtceDt="2024-01-06 10:00:00",
             dminsttNm="전남광주통합특별시 구례군", presmptPrce="1500000000", mainCnsttyNm="전기공사"),
        dict(bidNtceNo="G3", bidNtceOrd="000", bidNtceNm="국민임대아파트 도배장판 교체공사", ntceKindNm="일반공고", bidNtceDt="2024-02-06 10:00:00",
             dminsttNm="광주광역시 도시공사", presmptPrce="300000000", mainCnsttyNm="건축공사"),
        dict(bidNtceNo="G4", bidNtceOrd="000", bidNtceNm="공덕동 행복주택 건설공사", ntceKindNm="일반공고", bidNtceDt="2024-02-07 10:00:00",
             dminsttNm="서울주택도시개발공사", presmptPrce="50000000000", mainCnsttyNm="건축공사"),
    ])
    std = discover.standardize(raw, cfg, None)
    agg = discover.discover_candidates(std)
    names = dict(zip(agg["수요기관"], agg["시설명_후보"]))
    assert names["전남광주통합특별시 구례군"] == "구례군 국민임대아파트" and names["광주광역시 도시공사"] == "광주광역시 국민임대아파트", names
    assert names["서울주택도시개발공사"] == "공덕동 행복주택", "단지명이 있으면 그대로"
    rev = agg[agg["수요기관"] == "전남광주통합특별시 구례군"].copy()
    rev["검수_시설명"], rev["검수_수요기관"] = rev["시설명_후보"], rev["수요기관"]
    hits = discover.research_by_facility(std, rev)
    assert set(hits["공고번호"]) == {"G1", "G2"}, ("접두어를 뗀 '국민임대아파트'로 구례군 공고만 재검색", set(hits["공고번호"]))
    assert (hits["시설명"] == "구례군 국민임대아파트").all()


def test_discover_new_row_hints():
    """새 후보의 단서: 같은 이름의 시설이 다른 수요기관 이름으로 이미 있으면 알린다(기관 이름이 바뀐 경우). 새 검색어 행은 그 시설이 쓰던 이름을 그대로 쓴다."""
    mk = lambda w, k, org, nm, first="2024-01-01": {"표2_대분류": "X", "표2_중분류": "x", "검색어": w, "시설키": k, "시설명_후보": nm, "수요기관": org, "공고건수": 1, "최초공고일": first,
                                                   "최종공고일": first, "사업유형": "신축", "공고명_예시": nm + " 건립공사", "검수_포함여부": "포함", "검수_시설명": nm, "검수_별칭(;구분)": "",
                                                   "검수_수요기관": org, "검수_메모": ""}
    pv = pd.DataFrame([dict(mk("안식의 집", "제2안식의집", "강원도 춘천시", "제2안식의 집"), 시설ID="F0001", 검수_포함여부="제외"),
                       dict(mk("k1", "K", "O", "케이 옛이름"), 시설ID="F0002", 검수_시설명=None)])
    ag = pd.DataFrame([mk("k1", "K", "O", "케이 새이름"), mk("k2", "K", "O", "케이 새이름"), mk("안식의 집", "제2안식의집", "강원특별자치도 춘천시", "제2안식의 집", "2026-09-01"),
                       mk("k1", "전혀다른곳", "강원특별자치도 춘천시", "전혀 다른 곳", "2026-09-01")])
    out = discover._keep_review(ag, pv)
    st = list(out[discover.STATUS_COL])
    assert st[2] == "신규 / 이름이 같은 기존 시설: F0001 (수요기관 '강원도 춘천시') → 같은 시설이면 '제외' 하거나 시설ID 를 그 번호로" and st[3] == "신규", st
    assert list(out["검수_시설명"][:2]) == ["케이 옛이름"] * 2 and list(out["시설ID"][:2]) == ["F0002"] * 2 and st[1].startswith("신규 행(같은 시설 F0002"), out[["시설ID", "검수_시설명"]]
    old = discover._keep_review(pd.DataFrame([mk("k1", "N", "O", "엔", "2023-06-01")]), pv)
    assert old[discover.STATUS_COL][0].startswith("신규(예전 공고가 있음: 뒤늦게 수집된 공고이거나"), old[discover.STATUS_COL][0]
    # 확인하지 않은 채 한 번 더 돌려도 새 시설 표시와 단서는 '지난번: ' 을 붙여 남는다(같은 이름 단서 포함)
    again = discover._keep_review(ag, out)
    assert list(again[discover.STATUS_COL])[2] == "지난번: 신규 / 지난번: 이름이 같은 기존 시설: F0001 (수요기관 '강원도 춘천시') → 같은 시설이면 '제외' 하거나 시설ID 를 그 번호로", list(again[discover.STATUS_COL])


def test_audit_1121_fixes(tmp):
    """2026-11-21 워크북 점검에서 나온 결함: 도급액을 관급으로 읽어 합계가 두 배, 부분 공사 공고의 면적이 연면적으로, 포함공종의 비공종 어휘,
    띄어쓰기를 가로지른 사업유형 어휘, 기초금액 없는 통합발주 공고의 집계표 기준, 합계 밖에 적힌 관급 행, 견적서 안내공고 이름."""
    from g2b_costdb import classify, works, pipeline
    from g2b_costdb.extract_llm import gov_misreads, cross_verify, cross_verify_trades, normalize_trade_doc
    # ── ① 문서의 '도급액' 을 도급자관급액으로 읽은 값: 그 공고의 기초금액·추정가격과 같은 금액이면 관급이 아니다
    api = {"추정가격": 1_399_460_000, "기초금액": 1_539_406_000, "공고기관": "x"}
    doc = {"추정가격_원": 1_399_460_000, "기초금액_원": 1_539_406_000, "도급자관급액_원": 1_539_406_000, "관급자관급액_원": 206_451_000, "총공사금액_원": 1_745_857_000}
    assert gov_misreads(api, doc) == {"도급자관급액_원": "기초금액과 같은 금액"}, gov_misreads(api, doc)
    lg = {l["항목"]: l for l in cross_verify("G1", api, doc)}
    assert lg["도급자관급액"]["판정"] == "경고" and "04 에 싣지 않음" in lg["도급자관급액"]["비고"] and lg["도급자관급액"]["문서값"] == 1_539_406_000
    assert lg["관급자관급액"]["판정"] == "참고" and lg["총공사금액=기초금액+관급"]["판정"] == "정상", "오독한 값을 빼면 문서의 총공사금액과 맞는다"
    # 같은 금액이지만 문서의 총공사금액 때문에 그대로 실은 값은 06 에 따로 경고로 적는다('총공사금액=기초금액+관급' 경고 줄만으로는 찾을 수 없다)
    lk = {l["항목"]: l for l in cross_verify("G2", {"추정가격": 20_970_000, "기초금액": 23_067_000, "부가세_API": 2_097_000, "공고기관": "x"},
                                           {"도급자관급액_원": 20_970_000, "관급자관급액_원": 2_097_000, "총공사금액_원": 42_315_000})}
    assert lk["도급자관급액"]["판정"] == "경고" and "이 공고의 추정가격과 같은 금액이지만" in lk["도급자관급액"]["비고"] and "그대로 실음" in lk["도급자관급액"]["비고"], lk["도급자관급액"]
    assert lk["관급자관급액"]["판정"] == "경고" and "부가세와 같은 금액이지만" in lk["관급자관급액"]["비고"], lk["관급자관급액"]
    lk2 = {l["항목"]: l for l in cross_verify("G3", {"추정가격": 20_970_000, "기초금액": 23_067_000, "부가세_API": 2_097_000, "도급자관급액_API": 19_000_000, "공고기관": "x"},
                                            {"도급자관급액_원": 20_970_000, "관급자관급액_원": 2_097_000, "총공사금액_원": 42_315_000})}
    assert "그대로 실음" not in lk2["도급자관급액"]["비고"] and "API 우선 사용" in lk2["도급자관급액"]["비고"], "API 관급이 있으면 04 는 API 값 — 예전의 비교 행"
    assert "그대로 실음" in lk2["관급자관급액"]["비고"], "API 값이 없는 칸은 문서 값이 04 에 실린다"
    # 관급자재가 우연히 기초금액과 비슷한 대형 공사: 문서의 총공사금액이 (기초금액 + 관급)이면 그대로 둔다 / 원 단위까지 같지 않으면 걸지 않는다
    big = {"기초금액_원": 4_230_408_000, "도급자관급액_원": 4_230_408_000, "총공사금액_원": 8_460_816_000}
    assert gov_misreads({"추정가격": 3_845_825_455, "기초금액": 4_230_408_000}, big) == {}
    assert gov_misreads({"추정가격": 3_845_825_455, "기초금액": 4_230_408_000}, {"도급자관급액_원": 4_231_688_042}) == {}, "0.03% 다른 값은 관급일 수 있다"
    assert gov_misreads({"추정가격": 100_000_000, "기초금액": 110_000_000}, {"도급자관급액_원": 30_000_000, "관급자관급액_원": 5_000_000}) == {}
    assert gov_misreads({}, {}) == {} and gov_misreads({"추정가격": "", "기초금액": None}, {"도급자관급액_원": "nan"}) == {}
    # 추정가격·부가세를 두 관급 칸에 나눠 적은 경우, 총공사금액을 관급자관급액으로 읽은 경우
    both = gov_misreads({"추정가격": 20_970_000, "기초금액": 23_067_000, "부가세_API": 2_097_000}, {"도급자관급액_원": 20_970_000, "관급자관급액_원": 2_097_000, "총공사금액_원": 42_315_000})
    assert both == {}, "빼면 문서의 총공사금액(4,231만)에서 더 멀어진다(더하면 +9%, 빼면 −45%) → 손대지 않고 06 총공사금액 경고로 남긴다"
    both = gov_misreads({"추정가격": 20_970_000, "기초금액": 23_067_000, "부가세_API": 2_097_000}, {"도급자관급액_원": 20_970_000, "관급자관급액_원": 2_097_000})
    assert both == {"도급자관급액_원": "추정가격과 같은 금액", "관급자관급액_원": "부가세와 같은 금액"}, both       # 문서 총액이 없으면 같은 금액이라는 것만으로 뺀다
    assert gov_misreads({"추정가격": 47_890_000, "기초금액": 52_679_000}, {"관급자관급액_원": 54_340_000, "총공사금액_원": 54_340_000}) == {"관급자관급액_원": "문서의 총공사금액과 같은 금액"}
    # 문서의 '총공사금액' 이 이 공고의 총액일 수 없는 값이면(여러 공종의 총사업비 / 관급 소계) 그 값에 기대지 않는다
    assert gov_misreads({"추정가격": 100_000_000, "기초금액": 110_000_000}, {"도급자관급액_원": 110_000_000, "총공사금액_원": 300_000_000}) == {"도급자관급액_원": "기초금액과 같은 금액"}
    assert gov_misreads({"추정가격": 454_545_455, "기초금액": 500_000_000}, {"관급자관급액_원": 54_000_000, "총공사금액_원": 54_000_000}) == {}
    # API 에 기초금액이 없으면 추정가격 + 부가세와 견준다
    assert gov_misreads({"추정가격": 100_000_000}, {"도급자관급액_원": 110_000_000}) == {"도급자관급액_원": "추정가격+부가세와 같은 금액"}
    # 문서의 총공사금액 = 도급 + 관급: API 기초금액과 값이 달라도 '도급' 이 도급자관급액 칸에 든 것(관급 소계를 총액으로 읽은 문서는 건드리지 않는다)
    sp = gov_misreads({"추정가격": 1_725_295_454, "기초금액": 1_897_825_000}, {"도급자관급액_원": 2_023_680_000, "관급자관급액_원": 578_720_000, "총공사금액_원": 2_602_400_000})
    assert list(sp) == ["도급자관급액_원"] and "총공사금액에서 관급자관급액을 뺀" in sp["도급자관급액_원"], sp
    assert gov_misreads({"추정가격": 1_725_295_454, "기초금액": 1_897_825_000}, {"도급자관급액_원": 300_000_000, "관급자관급액_원": 200_000_000, "총공사금액_원": 500_000_000}) == {}
    # (검토 보강) 같은 금액으로 보는 폭: ±1,000원 또는 0.01% — 그보다 다르면 걸지 않는다
    small = {"추정가격": 5_000_000, "기초금액": 5_500_000}
    assert gov_misreads(small, {"도급자관급액_원": 5_500_900}) == {"도급자관급액_원": "기초금액과 같은 금액"}, "1,000원 안쪽"
    assert gov_misreads(small, {"도급자관급액_원": 5_502_000}) == {}, "2,000원 다른 값은 관급일 수 있다"
    big_api = {"추정가격": 3_845_825_455, "기초금액": 4_230_408_000}
    assert gov_misreads(big_api, {"도급자관급액_원": 4_230_808_000}) == {"도급자관급액_원": "기초금액과 같은 금액"}, "0.0095% 차이"
    assert gov_misreads(big_api, {"도급자관급액_원": 4_231_008_000}) == {}, "0.014% 차이"
    # (검토 보강) API 금액이 없을 때 문서의 기초금액·추정가격, 문서 부가세, 기초금액 − 추정가격과 견준다. 도급자관급액은 부가세와 견주지 않는다
    assert gov_misreads({"추정가격": None, "기초금액": None}, {"기초금액_원": 66_174_000, "도급자관급액_원": 66_174_000}) == {"도급자관급액_원": "문서의 기초금액과 같은 금액"}
    assert gov_misreads({}, {"추정가격_원": 60_158_181, "도급자관급액_원": 60_158_181}) == {"도급자관급액_원": "문서의 추정가격과 같은 금액"}
    assert gov_misreads({"추정가격": 100_000_000}, {"부가가치세_원": 9_500_000, "관급자관급액_원": 9_500_000}) == {"관급자관급액_원": "문서의 부가세와 같은 금액"}
    assert gov_misreads({"추정가격": 100_000_000, "기초금액": 108_000_000}, {"관급자관급액_원": 8_000_000}) == {"관급자관급액_원": "부가세(기초금액−추정가격)와 같은 금액"}
    assert gov_misreads({"추정가격": 100_000_000, "기초금액": 110_000_000, "부가세_API": 10_000_000}, {"도급자관급액_원": 10_000_000}) == {}
    # (검토 보강) 총액 = 두 관급의 합 규칙의 폭(기초금액의 125% 까지, 합이 0.1% 안에서 맞을 때만)
    a110 = {"추정가격": 100_000_000, "기초금액": 110_000_000}
    assert gov_misreads(a110, {"도급자관급액_원": 150_000_000, "관급자관급액_원": 30_000_000, "총공사금액_원": 180_000_000}) == {}, "기초금액의 136% 는 도급액으로 보지 않는다"
    assert gov_misreads(a110, {"도급자관급액_원": 120_000_000, "관급자관급액_원": 30_000_000, "총공사금액_원": 153_000_000}) == {}, "총액과 2% 어긋나면 걸지 않는다"
    # (검토 보강) '빼면 더 멀어진다' 판단은 남는 관급(걸리지 않은 칸)을 더한 값으로 한다 / 기준 금액은 API 가 먼저, 기초금액이 없으면 추정가격 + 부가세
    assert gov_misreads({"추정가격": 396_736_364, "기초금액": 436_410_000}, {"도급자관급액_원": 436_410_000, "관급자관급액_원": 906_400_000, "총공사금액_원": 1_342_810_000}) \
        == {"도급자관급액_원": "기초금액과 같은 금액"}, "실제 관급(9.06억)이 따로 있는 문서: 도급액만 뺀다"
    assert gov_misreads(a110, {"기초금액_원": 50_000_000, "도급자관급액_원": 110_000_000, "총공사금액_원": 110_000_000}) == {"도급자관급액_원": "기초금액과 같은 금액"}, "API 기초금액이 먼저"
    assert gov_misreads({"추정가격": 20_970_000, "부가세_API": 2_097_000}, {"도급자관급액_원": 20_970_000, "관급자관급액_원": 2_097_000, "총공사금액_원": 42_315_000}) == {}, "기초금액이 없으면 추정가격 + 부가세가 기준"
    # (검토 보강) 설정의 부가세율이 그대로 전달된다
    lg5 = {l["항목"]: l for l in cross_verify("G5", {"추정가격": 100_000_000, "공고기관": "x"}, {"도급자관급액_원": 105_000_000}, vat=0.05)}
    assert lg5["도급자관급액"]["판정"] == "경고" and "추정가격+부가세와 같은 금액" in lg5["도급자관급액"]["비고"], lg5["도급자관급액"]
    # 04 표: 오독한 값은 문서 열에 싣지 않는다(04 수식은 API 관급이 없으면 문서 관급을 그대로 더한다) — 다른 공고의 관급은 그대로
    base = {"프로젝트ID": "P1", "시설ID": "F1", "시설명": "s", "사업유형": "신축", "공종": "건축", "공고차수": "000", "공고명": "n", "공고일시": "2025-01-01", "수요기관": "d",
            "주공종명": "건축공사업", "부공종명": ""}
    lt = pd.DataFrame([dict(base, 공고키="A-000", 공고번호="A", 추정가격=1_399_460_000, 기초금액=1_539_406_000),
                       dict(base, 공고키="B-000", 공고번호="B", 추정가격=100_000_000, 기초금액=110_000_000)])
    tt = pipeline.assemble_trade_table(lt, {"A-000": doc, "B-000": {"도급자관급액_원": 30_000_000, "관급자관급액_원": 5_000_000}}).set_index("공고번호")
    assert pd.isna(tt.at["A", "도급자관급액_문서"]) and tt.at["A", "관급자관급액_문서"] == 206_451_000, tt.loc["A"].to_dict()
    assert tt.at["B", "도급자관급액_문서"] == 30_000_000 and tt.at["B", "관급자관급액_문서"] == 5_000_000
    ltc = pd.DataFrame([dict(base, 공고키="C-000", 공고번호="C", 추정가격=47_890_000, 기초금액=52_679_000), dict(base, 공고키="V-000", 공고번호="V", 추정가격=100_000_000, 기초금액=None)])
    dc = {"C-000": {"관급자관급액_원": 54_340_000, "총공사금액_원": 54_340_000}, "V-000": {"도급자관급액_원": 105_000_000, "관급자관급액_원": 7_000_000}}
    tc = pipeline.assemble_trade_table(ltc, dc, vat=0.05).set_index("공고번호")
    assert pd.isna(tc.at["C", "관급자관급액_문서"]), "관급자관급액 칸에 든 총공사금액도 싣지 않는다"
    assert pd.isna(tc.at["V", "도급자관급액_문서"]) and tc.at["V", "관급자관급액_문서"] == 7_000_000, "부가세율은 설정값(5%)으로 견준다"
    assert pipeline.assemble_trade_table(ltc, dc).set_index("공고번호").at["V", "도급자관급액_문서"] == 105_000_000
    # ── ② 연면적: 부분 공사 공고(건축·기타, 프로젝트 최대 공고의 20% 미만)의 값은 쓰지 않고, 건축 공고에 값이 없으면 여러 공고가 같게 적은 값
    c = lambda v, trade, amt, key="k": {"v": v, "raw": v, "trade": trade, "amt": amt, "key": key}
    pick, why = pipeline._pick_area([c(113, "건축", 22_383_636, "철거"), c(5181.08, "소방", 300e6), c(5181.08, "전기", 500e6), c(5200, "정보통신", 200e6)], 4_000e6)
    assert pick["v"] == 5181.08 and why == "", (pick, why)
    pick, why = pipeline._pick_area([c(40, "건축", 175_818_182, "목업-000")], 60_000e6)
    assert pick is None and "연면적 미채택" in why and "목업-000" in why and "40㎡" in why, why
    assert pipeline._pick_area([c(644.88, "건축", 3_000e6), c(16959.27, "전기", 400e6), c(16959.27, "소방", 200e6)], 3_000e6)[0]["v"] == 644.88, "본 건축 공고의 값이 먼저"
    assert pipeline._pick_area([c(900, "전기", 60e6), c(1200, "소방", 40e6)], 2_000e6)[0]["v"] == 900, "일치하는 값이 없으면 남은 후보의 첫째"
    pick, why = pipeline._pick_area([c(976, "기타", 44_575_455, "석면-000")], 912_928_182)
    assert pick is None and "기타 공고 석면-000" in why and "4.9%" in why, "기타 공종의 석면 철거 공고(최대 공고의 4.9%)도 부분 공사"
    assert pipeline.AREA_MINOR_SHARE == 0.20
    assert pipeline._pick_area([c(300, "건축", 400e6, "부속동")], 5_000e6)[0] is None, "최대 공고의 8% 인 건축 공고는 부분 공사"
    assert pipeline._pick_area([c(198.36, "건축", 519_640_909, "부대시설")], 3_770e6)[0] is None, "최대 공고의 13.8% 인 건축 공고도 부분 공사(실자료 F2882-N: 토목 37.7억 조성사업의 부대시설)"
    assert pipeline._pick_area([c(300, "건축", 950e6, "부속동")], 5_000e6)[0] is None, "19% 도 부분 공사"
    assert pipeline._pick_area([c(300, "건축", 1_000e6, "부속동")], 5_000e6)[0]["v"] == 300, "꼭 20% 면 부분 공사가 아니다"
    assert pipeline._pick_area([c(300, "전기", 10e6)], 5_000e6)[0]["v"] == 300, "건축·기타가 아닌 공종은 금액이 작아도 쓴다"
    agree = [c(1000, "전기", 600e6), c(1300, "소방", 50e6), c(1310, "정보통신", 40e6)]
    assert pipeline._pick_area(agree, 600e6)[0]["v"] == 1300, "건축 공고에 값이 없으면 두 공고가 같게(1% 이내) 적은 값이 금액 큰 한 공고의 값보다 먼저"
    assert pipeline._pick_area(agree[:2], 600e6)[0]["v"] == 1000
    assert pipeline._pick_area([c(900, "전기", 600e6), c(1200, "소방", 400e6), c(1250, "정보통신", 300e6)], 9_000e6)[0]["v"] == 900, "4% 다른 값은 같은 값이 아니다"
    assert pipeline._pick_area([c(900, "전기", 600e6), c(1200, "소방", 500e6), c(1200, "정보통신", 300e6), c(900.5, "기계설비", 100e6)], 9_000e6)[0]["v"] == 1200, "일치한 공고 수가 같으면 금액 합이 큰 쪽"
    lt7 = pd.DataFrame([dict(base, 프로젝트ID="P7", 공고키="M7-000", 공고번호="M7", 공종="건축", 추정가격=1_000_000_000),
                        dict(base, 프로젝트ID="P7", 공고키="A7-000", 공고번호="A7", 공종="건축", 추정가격=220_000_000),
                        dict(base, 프로젝트ID="P7", 공고키="E7-000", 공고번호="E7", 공종="전기", 추정가격=100_000_000),
                        dict(base, 프로젝트ID="P7", 공고키="S7-000", 공고번호="S7", 공종="소방", 추정가격=100_000_000),
                        dict(base, 프로젝트ID="P6", 공고키="M6-000", 공고번호="M6", 공종="건축", 추정가격=1_000_000_000),
                        dict(base, 프로젝트ID="P6", 공고키="E6-000", 공고번호="E6", 공종="전기", 추정가격=100_000_000),
                        dict(base, 프로젝트ID="P5", 공고키="M5-000", 공고번호="M5", 공종="건축", 추정가격=1_000_000_000),
                        dict(base, 프로젝트ID="P5", 공고키="K5-000", 공고번호="K5", 공종="기타", 추정가격=49_000_000)])
    e7 = pipeline.enrich_facility(pd.DataFrame([{"프로젝트ID": "P7"}, {"프로젝트ID": "P6"}, {"프로젝트ID": "P5"}]), lt7,
                                  {"A7-000": {"연면적_m2": 300.0}, "M6-000": {"연면적_m2": 0}, "E6-000": {"연면적_m2": 500.0},
                                   "M5-000": {"구조": "철골조"}, "K5-000": {"연면적_m2": 976}}).set_index("프로젝트ID")
    assert e7.at["P7", "연면적_m2"] == 300.0, "최대 공고의 22%(프로젝트 합계의 15%)인 건축 공고는 부분 공사가 아니다 — 기준은 최대 공고"
    assert e7.at["P6", "연면적_m2"] == 500.0, "0 으로 적힌 연면적은 값이 아니다"
    assert pd.isna(e7.at["P5", "연면적_m2"]) and e7.at["P5", "개요_출처"].startswith("건축 / 연면적 미채택: 기타 공고 K5-000 의 976㎡"), e7.at["P5", "개요_출처"]
    assert pipeline._pick_area([c(36.96, "건축", 80e6)], 80e6)[0]["v"] == 36.96, "공고 하나짜리 프로젝트는 그 공고가 본 공사"
    assert pipeline._pick_area([c(300, "건축", 1_200e6, "부속동")], 5_000e6)[0]["v"] == 300, "최대 공고의 20% 이상인 건축 공고는 그대로"
    assert pipeline._pick_area([], 1e9) == (None, "")
    # 금액을 모르는 공고(추정가격 없음 → 0)는 '작은 공고' 로 보지 않는다 / 까닭에 적는 면적은 자릿수를 줄이지 않는다
    assert pipeline._pick_area([c(3000, "건축", 0.0, "본공사")], 36_508_000)[0]["v"] == 3000
    assert "12,345.67㎡" in pipeline._pick_area([c(12345.67, "기타", 1e6, "철거")], 1e9)[1]
    lt2 = pd.DataFrame([dict(base, 프로젝트ID="P9", 공고키="M-000", 공고번호="M", 공종="건축", 추정가격=4_000_000_000),
                        dict(base, 프로젝트ID="P9", 공고키="D-000", 공고번호="D", 공종="건축", 추정가격=22_383_636),
                        dict(base, 프로젝트ID="P9", 공고키="E-000", 공고번호="E", 공종="전기", 추정가격=500_000_000),
                        dict(base, 프로젝트ID="P9", 공고키="S-000", 공고번호="S", 공종="소방", 추정가격=300_000_000),
                        dict(base, 프로젝트ID="P8", 공고키="H-000", 공고번호="H", 공종="건축", 추정가격=90_000_000_000),
                        dict(base, 프로젝트ID="P8", 공고키="U-000", 공고번호="U", 공종="건축", 추정가격=175_818_182)])
    dd = {"M-000": {"구조": "철근콘크리트"}, "D-000": {"연면적_m2": 113, "건축면적_m2": 113, "지상층수": 1}, "E-000": {"연면적_m2": 5181.08}, "S-000": {"연면적_m2": 5181.08},
          "H-000": {"용도": "숙박시설"}, "U-000": {"연면적_m2": 40}}
    ef = pipeline.enrich_facility(pd.DataFrame([{"프로젝트ID": "P9"}, {"프로젝트ID": "P8"}]), lt2, dd).set_index("프로젝트ID")
    assert ef.at["P9", "연면적_m2"] == 5181.08 and ef.at["P9", "지상층수"] == 1 and "전기" in ef.at["P9", "개요_출처"], ef.loc["P9"].to_dict()
    assert pd.isna(ef.at["P9", "건축면적_m2"]), "철거 공고의 건축면적(철거한 건물)도 싣지 않는다"
    assert pd.isna(ef.at["P8", "연면적_m2"]) and "연면적 미채택" in ef.at["P8", "개요_출처"] and ef.at["P8", "용도"] == "숙박시설", ef.loc["P8"].to_dict()
    lt2z = pd.DataFrame([dict(base, 프로젝트ID="P7", 공고키="M2-000", 공고번호="M2", 공종="건축", 추정가격=4_000_000_000), dict(base, 프로젝트ID="P7", 공고키="E2-000", 공고번호="E2", 공종="전기", 추정가격=500_000_000)])
    efz = pipeline.enrich_facility(pd.DataFrame([{"프로젝트ID": "P7"}]), lt2z, {"M2-000": {"연면적_m2": 0}, "E2-000": {"연면적_m2": 777}}).set_index("프로젝트ID")
    assert efz.at["P7", "연면적_m2"] == 777, "0㎡ 는 값이 아니다 → 다른 공고의 값"
    # ── ③ 포함공종_문서: 공종 라벨만 남긴다. 2차 대상은 주요 공종이 둘 이상일 때만
    n = classify.normalize_trade_list
    assert n("전기, 1식") == "전기" and n("정보통신·CCTV설비·주차관제설비") == "정보통신" and n("조경, 식재, 포장공, 부대공") == "조경"
    assert n("건축(+100%·상대업종·철강구조물·철근·콘크리트)") == "건축" and n("가설·수장·토공") == "" and n("건축, 토목, 기계, 조경") == "건축·토목·기계설비·조경"
    assert n("전기공사 및 소방시설공사(철거 포함)") == "전기·소방·철거"
    assert classify.main_trade_count("전기·철거·가스") == 1 and classify.main_trade_count("건축·토목·1식") == 2 and classify.main_trade_count("") == 0
    lt3 = pd.DataFrame([dict(base, 공고키="E1-000", 공고번호="E1", 공종="전기", 주공종명="전기공사업"), dict(base, 공고키="E2-000", 공고번호="E2", 공종="전기", 주공종명="전기공사업"),
                        dict(base, 공고키="E3-000", 공고번호="E3", 공종="전기", 주공종명="전기공사업"), dict(base, 공고키="K1-000", 공고번호="K1", 공종="기타", 주공종명="")])
    tg = pipeline._trade_targets(lt3, {"E1-000": {"포함공종_문서": "전기, 1식"}, "E2-000": {"포함공종_문서": "전기, 소방"}, "E3-000": {"포함공종_문서": "전기, 철거"},
                                      "K1-000": {"포함공종_문서": "가설, 철거, 포장"}}, {})
    assert tg == ["E2-000"], tg
    t3 = pipeline.assemble_trade_table(lt3, {"E1-000": {"포함공종_문서": "전기, 1식"}, "K1-000": {"포함공종_문서": "가설, 철거, 포장"}}).set_index("공고번호")
    assert t3.at["E1", "포함공종(종합)"] == "전기" and t3.at["K1", "포함공종_문서"] == "철거"
    # 소방 공고의 '소방(기계, 전기)'·'기계소방, 전기소방' 은 소방의 두 분야이지 기계설비·전기 공종이 아니다(API 에 그 공종이 따로 있으면 그대로)
    assert classify.drop_fire_subfields(n("소방(기계, 전기)"), "소방") == "소방" and classify.drop_fire_subfields(n("기계소방, 전기소방"), "소방") == "소방"
    assert classify.drop_fire_subfields("소방·기계설비", "소방", "소방·기계설비") == "소방·기계설비" and classify.drop_fire_subfields("전기·소방", "전기") == "전기·소방"
    assert classify.drop_fire_subfields("소방·정보통신", "소방") == "소방·정보통신" and classify.drop_fire_subfields("", "소방") == ""
    lt3f = pd.DataFrame([dict(base, 공고키="S1-000", 공고번호="S1", 공종="소방", 주공종명="소방시설공사업"), dict(base, 공고키="S2-000", 공고번호="S2", 공종="소방", 주공종명="소방시설공사업"),
                         dict(base, 공고키="S3-000", 공고번호="S3", 공종="소방", 주공종명="소방시설공사업", 부공종명="기계설비공사업")])
    dfire = {"S1-000": {"포함공종_문서": "소방(기계, 전기)"}, "S2-000": {"포함공종_문서": "소방, 통신"}, "S3-000": {"포함공종_문서": "기계소방"}}
    assert pipeline._trade_targets(lt3f, dfire, {}) == ["S2-000", "S3-000"], pipeline._trade_targets(lt3f, dfire, {})
    t3f = pipeline.assemble_trade_table(lt3f, dfire).set_index("공고번호")
    assert t3f.at["S1", "포함공종(종합)"] == "소방" and t3f.at["S1", "포함공종_문서"] == "소방" and t3f.at["S3", "포함공종(종합)"] == "소방·기계설비", t3f[["포함공종(종합)", "포함공종_문서"]].to_dict()
    # ── ④ 사업유형 어휘: 앞 어절의 끝 글자와 뒤 어절의 첫 글자가 이어져 생긴 어휘, 다른 낱말 속의 어휘는 치지 않는다
    w = classify.classify_work_type
    assert w("공원녹지 권역별 관리공사(만수권역)") == "미분류", "'권역별 관리' 는 별관(증축)이 아니다"
    assert w("2026년 노후영구임대 수선유지공사") == "유지보수", "'임대 수선' 은 대수선(리모델링)이 아니다"
    assert w("쌍신축구장 천연잔디 유지관리공사") == "유지보수" and w("쌍신 축구장 조명시설 설치공사(전기)") == "유지보수", "'쌍신축구장' 의 신축"
    assert w("종합운동장 통신설비 및 선로교체 도급공사") == "유지보수", "'통신설비' 의 신설"
    assert w("○○센터 건축 조경공사") == "미분류" and w("혁신 축산단지 진입로 정비공사") == "유지보수"
    assert w("대전 신축구장 관람석 설치공사") == "미분류" and w("신축구장 주변 정비공사") == "미분류", "어절 첫머리의 '신축구장' 은 신축 + 구장일 수 있다 → 예전처럼 검토 대상(미분류)"
    # 유지보수 어휘는 예전처럼 공백을 뺀 문자열에서 찾는다(유지보수 → 미분류로 바뀌어 재검색에 새로 들어오는 공고가 생기지 않게)
    assert w("외벽방 수공사") == "유지보수" and w("노후시설 개선공사") == "유지보수"
    # 띄어 쓴 어휘·낱자 띄어쓰기·온전한 어휘는 그대로 잡는다
    assert w("○○초등학교 별관 증축공사") == "증축" and w("문화회관 별관 건립공사") == "증축" and w("시민회관 대수선 공사") == "리모델링"
    assert w("주민센터 환경 개선 공사") == "유지보수" and w("청사 대 수선 공사") == "리모델링" and w("○ ○ 도 서 관 증 축 공 사") == "증축"
    assert w("풋살장 및 축구장 신축공사") == "신축" and w("축구센터 건립공사") == "신축" and w("정보통신설비 신설공사") == "신축"
    assert w("기능보강 리모델링 공사") == "리모델링" and w("기능보강리모델링공사") == "리모델링" and w("재건축 조경공사") == "리모델링"
    assert w("도시철도 1호선 구암역 환승주차장 건설 기계설비공사") == "미분류", "신축 어휘 + 부속어(주차장·설비)는 예전대로"
    # 붙여 쓴 '건축조경' 의 축조 / 공고명 맨 앞에서 시작하는 낱자 띄어쓰기 / 숫자 뒤에서 걸린 어휘('2개 보수', '2건 설치') / 우연히 걸린 자리 뒤에 진짜 어휘가 또 있는 공고명
    assert w("○○센터 건축조경공사") == "미분류" and w("증 축 공 사") == "증축" and w("쌍신축구센터 잔디 보수") == "유지보수"
    assert w("화장실 2개 보수공사") == "유지보수" and w("소방시설 등 2건 설치공사") == "미분류"
    assert w("혁신 축산단지 관리동 신 축 공 사") == "신축", "앞의 '혁신 축산' 은 버리되 뒤의 '신 축' 은 잡는다"
    # ── ⑤ 2차 검증: 기초금액이 없는 공고는 추정가격 + 부가세도 기준 후보 / 합계 밖에 따로 적힌 관급·부가세 행
    td = normalize_trade_doc({"합계_원": "109100000000", "공종별금액": [{"공종": "건축", "항목명_원문": "건축공사", "금액_원": "80000000000"}, {"공종": "전기", "항목명_원문": "전기공사", "금액_원": "18140000000"},
                                                                {"공종": "정보통신", "항목명_원문": "통신공사", "금액_원": "8350000000"}, {"공종": "소방", "항목명_원문": "소방공사", "금액_원": "2610000000"}]})
    row = {"공고번호": "V", "공종": "건축", "추정가격": 99_181_818_182, "기초금액": None}
    basis = [l for l in cross_verify_trades("V", row, td) if l["항목"] == "집계표합계 기준"][0]
    assert basis["판정"] == "참고" and "추정가격 + 부가세" in basis["비고"], basis
    assert pipeline._alloc_decision(row, td) == "Y"
    row2 = dict(row, 도급자관급액_API=3_000_000_000, 관급자관급액_API=30_000_000_000)      # 관급자 설치분은 집계표에 없고 도급자 설치분만 든 문서
    td2 = normalize_trade_doc({"합계_원": "112100000000", "공종별금액": td["공종별금액"] + [{"공종": "관급자재", "항목명_원문": "도급자설치 관급자재", "금액_원": "3000000000"}]})
    b2 = [l for l in cross_verify_trades("V2", row2, td2) if l["항목"] == "집계표합계 기준"][0]
    assert b2["판정"] == "참고" and "도급자관급" in b2["비고"], b2
    far = [l for l in cross_verify_trades("V3", row, normalize_trade_doc({"합계_원": "150000000000", "공종별금액": td["공종별금액"]})) if l["항목"] == "집계표합계 기준"][0]
    assert far["판정"] == "경고", "맞는 기준이 없는 집계표는 여전히 경고"
    out = normalize_trade_doc({"합계_원": "6624341000", "관급자재_포함여부": "제외", "공종별금액": [
        {"공종": "건축", "항목명_원문": "건축공사", "금액_원": "4954661000"}, {"공종": "토목", "항목명_원문": "토목공사", "금액_원": "1370487000"},
        {"공종": "기계설비", "항목명_원문": "기계설비공사", "금액_원": "299193000"}, {"공종": "관급자재", "항목명_원문": "관급자재비", "금액_원": "2090000000"}]})
    r4 = {"공고번호": "X4", "공종": "건축", "추정가격": 6_022_128_182, "기초금액": 6_624_341_000}
    assert pipeline._alloc_decision(r4, out) == "Y", pipeline._alloc_decision(r4, out)
    assert [l for l in cross_verify_trades("X4", r4, out) if l["항목"] == "공종합=집계표합계"][0]["판정"] == "정상"
    dup = normalize_trade_doc({"합계_원": "1000", "공종별금액": [{"공종": "건축", "항목명_원문": "건축공사", "금액_원": "900"}, {"공종": "전기", "항목명_원문": "전기공사", "금액_원": "600"}]})
    assert "중복·다른 집계표 의심" in pipeline._alloc_decision({"공고번호": "X5", "추정가격": 909, "기초금액": 1000}, dup), "공종 행 합이 합계의 150% 인 문서는 여전히 배분하지 않는다"
    # 맞아떨어지는 조합이 없는 문서는 예전대로 비공종 행을 모두 합계에서 뺀다(문서의 '관급자재_포함여부' 표기는 보지 않는다 — 금액이 1% 안에서 맞는지만 본다):
    # 관급이 합계의 대부분인 문서는 '일부 행만 읽힘', 공종 행 합이 합계를 넘는 문서는 '중복 의심' 그대로
    part_only = normalize_trade_doc({"합계_원": "72380000", "관급자재_포함여부": "포함", "공종별금액": [{"공종": "기계설비", "항목명_원문": "기계설비공사", "금액_원": "4222570"},
                                                                                       {"공종": "관급자재", "항목명_원문": "관급자재", "금액_원": "70020000"}, {"공종": "부가세", "항목명_원문": "부가세", "금액_원": "6580000"}]})
    dec = pipeline._alloc_decision({"공고번호": "X7", "추정가격": 65_800_000, "기초금액": 72_380_000}, part_only)
    assert dec.startswith("N(") and "일부 행만" in dec, dec
    unk = normalize_trade_doc({"합계_원": "7417173000", "관급자재_포함여부": "불명", "공종별금액": [{"공종": "건축", "항목명_원문": "건축공사", "금액_원": "5000000000"}, {"공종": "전기", "항목명_원문": "전기공사", "금액_원": "3221529304"},
                                                                                       {"공종": "관급자재", "항목명_원문": "관급자재", "금액_원": "2800000000"}]})
    dec = pipeline._alloc_decision({"공고번호": "X8", "추정가격": 6_742_884_545, "기초금액": 7_417_173_000}, unk)
    assert "중복·다른 집계표 의심" in dec, dec
    inside = normalize_trade_doc({"합계_원": "1300", "공종별금액": [{"공종": "건축", "항목명_원문": "건축공사", "금액_원": "1000"}, {"공종": "관급자재", "항목명_원문": "관급자재비", "금액_원": "300"}]})
    assert [l for l in cross_verify_trades("X6", {"추정가격": 909, "기초금액": 1000, "관급자재_API": 300}, inside) if l["항목"] == "공종합=집계표합계"][0]["판정"] == "정상", "관급이 합계 안에 든 문서도 정상"
    # 합계 안팎은 라벨 단위가 아니라 행 단위로 고른다(실자료 20230541184: 부가세·도급자 설치 관급은 합계 안, 관급자 설치 관급·분담금은 합계 밖)
    from g2b_costdb.extract_llm import inside_total_rows
    mix = normalize_trade_doc({"합계_원": "2880650000", "공종별금액": [
        {"공종": "전기", "항목명_원문": "추정가격", "금액_원": "1989304773"}, {"공종": "부가세", "항목명_원문": "부가가치세", "금액_원": "198930477"},
        {"공종": "관급자재", "항목명_원문": "도급자설치관급액", "금액_원": "692415000"}, {"공종": "관급자재", "항목명_원문": "관급자설치관급액", "금액_원": "522563000"},
        {"공종": "배분제외", "항목명_원문": "기타금액(분담금 등)", "금액_원": "46936000"}]})
    assert inside_total_rows(mix["공종별금액"], 2_880_650_000, 1_989_304_773) == (198_930_477 + 692_415_000, True)
    r9 = {"공고번호": "X9", "공종": "전기", "추정가격": 1_989_304_773, "기초금액": 2_188_235_000, "도급자관급액_API": 692_415_000, "관급자관급액_API": 522_563_000}
    b9 = [l for l in cross_verify_trades("X9", r9, mix) if l["항목"] == "집계표합계 기준"][0]
    assert pipeline._alloc_decision(r9, mix) == "Y" and b9["판정"] == "참고" and "기초금액 + 도급자관급(API)" in b9["비고"], b9
    b10 = [l for l in cross_verify_trades("V4", dict(row, 도급자관급액_API=3_000_000_000), td2) if l["항목"] == "집계표합계 기준"][0]
    assert "관급자재(API)" in b10["비고"] and "도급자관급" not in b10["비고"], "관급자 설치분이 없으면 '+ 도급자관급' 후보를 따로 두지 않는다"
    it = lambda lab, amt: {"공종": lab, "항목명_원문": lab, "금액_원": amt}
    assert inside_total_rows([it("건축", 1000), it("간접비", 300), it("관급자재", 200)], 1300, 1000) == (300, True), "공통가설·간접비는 늘 합계 안"
    assert inside_total_rows([it("건축", 1000), it("배분제외", 50)], 1050, 1000) == (50, True) and inside_total_rows([it("건축", 1000), it("배분제외", 50)], 1000, 1000) == (0, True)
    assert inside_total_rows([it("건축", 1000)], 1005, 1000)[1] is True and inside_total_rows([it("건축", 1000)], 1015, 1000)[1] is False, "맞음 = 1% 이내"
    many = [it("건축", 1000)] + [it("관급자재", 10)] * 13            # 비공종 행이 12개를 넘으면 라벨 단위로만 넣고 뺀다(행 5개만 넣는 조합은 찾지 않는다)
    assert inside_total_rows(many, 1130, 1000) == (130, True) and inside_total_rows(many, 1050, 1000)[1] is False
    # 분리발주 중복 가능 로그: 배분적용 N 인 공고는 05 가 아니라 04b 를 보라고 적고, 여러 시설에 배정된 공고도 한 번만
    lt5 = pd.DataFrame([dict(base, 프로젝트ID=pid, 공고키="A-000", 공고번호="A", 공종="건축", 추정가격=909, 기초금액=1000) for pid in ("P1", "P2")] +
                       [dict(base, 프로젝트ID=pid, 공고키="E-000", 공고번호="E", 공종="전기", 추정가격=10, 기초금액=11) for pid in ("P1", "P2")])
    d5 = pipeline._trade_dup_logs(lt5, {"A-000": dup})
    assert len(d5) == 1 and "배분적용 N" in d5[0]["비고"] and "04b" in d5[0]["비고"], d5
    okdoc = normalize_trade_doc({"합계_원": "1000", "공종별금액": [{"공종": "건축", "항목명_원문": "건축공사", "금액_원": "800"}, {"공종": "전기", "항목명_원문": "전기공사", "금액_원": "200"}]})
    d6 = pipeline._trade_dup_logs(lt5, {"A-000": okdoc})
    assert len(d6) == 1 and "05 '문서배분_중복경고' 참고" in d6[0]["비고"], d6
    # ── ⑥ '견적서 제출 안내공고' 가 세부사업명에 '서 제출' 로 남지 않는다
    assert works.clean_title("△△ 체육관 방수공사 견적서제출안내공고") == "△△ 체육관 방수공사" and works.clean_title("2인 견적서 제출 서산문화회관 보수공사") == "2인 서산문화회관 보수공사"
    assert works.clean_title("수의견적서울시립미술관 방수공사") == "서울시립미술관 방수공사" and works.clean_title("○○도서관 견적서고 리모델링 공사") == "○○도서관 서고 리모델링 공사", "'견적' 에 붙은 다음 낱말의 첫 글자 '서' 는 남긴다"
    assert works.clean_title("용인시문예회관 별관 개량공사 수의계약 견적서 제출 안내공고") == "용인시문예회관 별관 개량공사"
    assert works.clean_title("△△ 체육관 방수공사 (수의계약 견적서 제출 안내공고)") == "△△ 체육관 방수공사"
    assert works.clean_title("예산군 문예회관 외부 리모델링 공사 소액수의견적제출 안내  공고") == "예산군 문예회관 외부 리모델링 공사"
    assert works.clean_title("△△ 체육관 방수공사 견적서 제출 안내공고") == "△△ 체육관 방수공사", "'수의계약' 없이 '견적서 제출' 만 있는 공고명"
    # ── ⑦ discover 요약: '포함' 으로 들어갈 새 시설을 먼저·빠짐없이 보이고, 검수하지 않으면 어떻게 되는지 알린다. research 도 같은 행을 알린다
    import io, contextlib
    from g2b_costdb import discover
    ST = discover.STATUS_COL
    rows = [{"시설ID": f"F{i:04d}", "시설명_후보": f"시설{i}", "검수_시설명": f"시설{i}", "검수_포함여부": "제외", ST: "신규"} for i in range(1, 19)]
    rows[16]["검수_포함여부"] = rows[17]["검수_포함여부"] = "포함"                       # 17·18번째 행(예전에는 15행에서 잘려 화면에 나오지 않았다)
    rows += [{"시설ID": "F0100", "시설명_후보": "기존", "검수_시설명": "기존", "검수_포함여부": "포함", ST: "신규 행(같은 시설 F0001 의 다른 검색어: 검수 값은 그 행과 같게 둠)"},
             {"시설ID": "F0101", "시설명_후보": "지난", "검수_시설명": "지난", "검수_포함여부": "포함", ST: "지난번: 신규 / 지난번: 비슷한 기존 시설: F0001 '시설1' → 같은 시설이면 …"},
             {"시설ID": "F0102", "시설명_후보": "그대로", "검수_시설명": "그대로", "검수_포함여부": "포함", ST: ""}]
    agg = pd.DataFrame(rows)
    assert list(pipeline._unreviewed_new_included(agg)["시설ID"]) == ["F0017", "F0018", "F0101"]
    assert pipeline._unreviewed_new_included(agg.drop(columns=[ST])).empty and pipeline._unreviewed_new_included(pd.DataFrame()).empty
    buf = io.StringIO()
    with contextlib.redirect_stdout(buf):
        pipeline._review_change_report(agg)
    out = buf.getvalue()
    new_part = out[out.index("새 시설(기본값 확인): 18행"):]
    assert new_part.index("F0017") < new_part.index("F0001") and "F0018 시설18 [포함]" in new_part and "… 외 3행" in new_part, out
    assert "[주의] 새로 잡힌 시설 가운데 검수_포함여부가 '포함'(기본값)인 3행: F0017, F0018, F0101" in out and "유료 추출" in out, out
    many = pd.DataFrame([{"시설ID": f"F{i:04d}", "시설명_후보": f"시설{i}", "검수_시설명": f"시설{i}", "검수_포함여부": "포함" if i > 3 else "제외", ST: "신규"} for i in range(1, 24)])
    buf = io.StringIO()
    with contextlib.redirect_stdout(buf):
        pipeline._review_change_report(many)
    assert all(f"F{i:04d} 시설{i} [포함]" in buf.getvalue() for i in range(4, 24)) and "… 외 3행" in buf.getvalue(), "'포함' 인 새 시설 20행은 15행에서 잘리지 않는다"
    assert len(pipeline._unreviewed_new_included(pd.DataFrame([{"시설ID": "F1", "검수_포함여부": " 포함 ", ST: " 신규"}]))) == 1, "재검색(read_reviewed)과 같은 기준: 앞뒤 공백은 뗀다"
    # 지난 실행 때 새로 잡혀 아직 그대로인 행(검수 없이 discover 를 한 번 더 돌린 경우): 사업유형·시설키가 달라져 다른 표시가 앞에 붙어도 새 시설로 세고,
    # '지난 실행 때 표시됐고 …' 묶음에서도 '포함' 인 행을 먼저·빠짐없이 보인다(2026-11-21 PC 의 다음 실행이 바로 이 경우: 새 18행이 모두 '지난번: 신규')
    late = [{"시설ID": f"F{i:04d}", "시설명_후보": f"시설{i}", "검수_시설명": f"시설{i}", "검수_포함여부": "제외", ST: "지난번: 신규"} for i in range(1, 19)]
    late[15]["검수_포함여부"] = late[16]["검수_포함여부"] = late[17]["검수_포함여부"] = "포함"
    late[16][ST] = "확인: 사업유형 증축→미분류, 새 기본값 '검토필요'(지금 '포함' 그대로 둠) / 지난번: 신규"
    late[17][ST] = "후보 이름 바뀜: 'a' → 'b'(이전 줄의 공고명_예시로 다시 맞춤. 다른 시설로 보이면 검수 칸 확인) / 지난번: 신규"
    agg2 = pd.DataFrame(late)
    assert list(pipeline._unreviewed_new_included(agg2)["시설ID"]) == ["F0016", "F0017", "F0018"]
    buf = io.StringIO()
    with contextlib.redirect_stdout(buf):
        pipeline._review_change_report(agg2)
    part = buf.getvalue()[buf.getvalue().index("지난 실행 때 표시됐고"):]
    assert all(f"F00{i} 시설{i} [포함]" in part for i in (16, 17, 18)) and part.index("F0016") < part.index("F0001") and "… 외 3행" in part, buf.getvalue()
    # 갱신_표시 열을 복사해 같은 이름의 열이 둘인 파일에서도 research 의 알림 때문에 멈추지 않는다
    dup_cols = pd.concat([agg, agg[[ST]]], axis=1)
    assert list(pipeline._unreviewed_new_included(dup_cols)["시설ID"]) == ["F0017", "F0018", "F0101"]
    # research: 재검색을 시작하기 전에 같은 행을 알린다
    cfg_r = dict(load_config())
    cfg_r["paths"] = dict(cfg_r["paths"], data_dir=os.path.join(tmp, "rs", "data"), out_dir=os.path.join(tmp, "rs", "out"))
    os.makedirs(cfg_r["paths"]["data_dir"], exist_ok=True)
    pd.DataFrame([{"공고명": "x", "수요기관": "y"}]).to_parquet(os.path.join(cfg_r["paths"]["data_dir"], "notices_std.parquet"), index=False)
    discover.write_review_workbook(agg, os.path.join(cfg_r["paths"]["out_dir"], "facility_candidates.xlsx"))
    real_rs, buf = discover.research_by_facility, io.StringIO()
    discover.research_by_facility = lambda std, reviewed: pd.DataFrame()
    try:
        with contextlib.redirect_stdout(buf):
            pipeline.stage_research(cfg_r)
        raise AssertionError("재검색 결과가 없으면 멈춰야 한다")
    except SystemExit as e:
        assert "재검색 결과 없음" in str(e)
    finally:
        discover.research_by_facility = real_rs
    assert "[주의] 검수 파일에 새로 잡힌 시설 3행이 '포함' 으로 들어 있습니다" in buf.getvalue() and all(x in buf.getvalue() for x in ("F0017 시설17", "F0018 시설18", "F0101 지난")), buf.getvalue()
    # discover 마지막 표: 이번 후보에 없어 남겨 둔 행은 세지 않는다(머리줄의 후보 시설 수와 합이 같다)
    from g2b_costdb import collect
    cfg_d = dict(load_config())
    cfg_d["paths"] = dict(cfg_d["paths"], data_dir=os.path.join(tmp, "dc", "data"), raw_dir=os.path.join(tmp, "dc", "raw"), out_dir=os.path.join(tmp, "dc", "out"))
    cfg_d["api"] = dict(cfg_d["api"], use_awards=False)
    os.makedirs(cfg_d["paths"]["data_dir"], exist_ok=True)
    agg_d = pd.DataFrame([{"시설ID": "F0001", "표2_대분류": "문화 및 집회시설", "사업유형": "신축", ST: "", "검수_포함여부": "포함"},
                          {"시설ID": "F0002", "표2_대분류": "숙박시설", "사업유형": "증축", ST: "신규", "검수_포함여부": "포함"},
                          {"시설ID": "F0003", "표2_대분류": "문화 및 집회시설", "사업유형": "신축", ST: "이번 후보에 없음(이전 내용 그대로 둠)", "검수_포함여부": "제외"}])
    saved = (collect.incomplete_months, pipeline._read_raw_notices, discover.standardize, discover.discover_candidates)
    collect.incomplete_months, pipeline._read_raw_notices = (lambda cfg, prefix: []), (lambda cfg: pd.DataFrame())
    discover.standardize = lambda raw, cfg, bsis, awards=None: pd.DataFrame([{"공고명": "x"}] * 5)
    discover.discover_candidates = lambda std, previous_review=None, ids_only=False, recalc=False: agg_d
    buf = io.StringIO()
    try:
        with contextlib.redirect_stdout(buf):
            pipeline.stage_discover(cfg_d)
    finally:
        collect.incomplete_months, pipeline._read_raw_notices, discover.standardize, discover.discover_candidates = saved
    import re as _re
    out = buf.getvalue()
    assert "후보 시설 2개(이번 후보에 없어 남겨 둔 1행은 따로)" in out and _re.search(r"문화 및 집회시설\s+1\n", out) and _re.search(r"신축\s+1\n", out), out
    # extract 견적 줄: 같은 공고가 여러 시설에 배정돼 있어도 고유 공고로 센다(대상 + 완료 + 텍스트 없음 = 대표 공고)
    buf = io.StringIO()
    with contextlib.redirect_stdout(buf):
        pipeline._print_estimate(load_config(), {"A-000": "t", "B-000": "t"}, ["B-000"], {"A-000": {"신뢰도": "high"}, "OLD-000": {"신뢰도": "high"}},
                                 pd.DataFrame([{"공고키": "A-000"}, {"공고키": "A-000"}, {"공고키": "B-000"}, {"공고키": "C-000"}, {"공고키": "C-000"}]))
    assert "LLM 추출 대상: 1건 (대표 공고 3건(고유) 중 이미 완료 1건, 텍스트 없음 1건)" in buf.getvalue(), buf.getvalue()
    # ── ⑧ attach: 같은 공고가 여러 시설에 배정돼 있어도 한 번만 처리하고, 이전에 겹쳐 쌓인 똑같은 기록은 한 줄로
    from g2b_costdb import attachments
    cfg = load_config()
    cfg["paths"] = dict(cfg["paths"], data_dir=os.path.join(tmp, "att", "data"), files_dir=os.path.join(tmp, "att", "files"), text_dir=os.path.join(tmp, "att", "text"))
    for k in ("data_dir", "files_dir", "text_dir"):
        os.makedirs(cfg["paths"][k], exist_ok=True)
    lt8 = pd.DataFrame([dict(base, 프로젝트ID=pid, 시설ID=fid, 공고키="Z-000", 공고번호="Z", 추정가격=1, 기초금액=1, 첨부URL1="http://x/a.hwp", 첨부파일명1="a.hwp")
                        for pid, fid in (("P1", "F1"), ("P2", "F2"), ("P3", "F3"))])
    lt8.to_parquet(os.path.join(cfg["paths"]["data_dir"], "notices_latest.parquet"), index=False)
    old_rec = {"공고번호": "OLD", "파일명": "b.hwp", "우선순위": 1, "URL": "http://x/b.hwp", "다운로드": "Y", "추출성공": "Y", "추출글자수": 10, "파서": "p", "오류": ""}
    pd.DataFrame([old_rec, old_rec, old_rec]).to_parquet(os.path.join(cfg["paths"]["data_dir"], "notes_attach.parquet"), index=False)
    calls = []
    real = attachments.process_notice_attachments
    def fake(rd, files_dir, text_dir, attach_max=10):
        calls.append(rd["공고키"])
        return "", [{"공고번호": "Z", "파일명": "a.hwp", "우선순위": 1, "URL": "http://x/a.hwp", "다운로드": "N", "추출성공": "N", "추출글자수": 0, "파서": "", "오류": "다운로드 실패"}]
    attachments.process_notice_attachments = fake
    try:
        with contextlib.redirect_stdout(io.StringIO()):
            pipeline.stage_attach(cfg)
    finally:
        attachments.process_notice_attachments = real
    nt = pd.read_parquet(os.path.join(cfg["paths"]["data_dir"], "notes_attach.parquet"))
    assert calls == ["Z-000"] and sorted(nt["공고번호"]) == ["OLD", "Z"], (calls, nt["공고번호"].tolist())


def test_research_fallback_and_merge():
    """정확한 이름으로 못 찾으면 어절 간격을 허용해 찾고, 같은 시설ID의 검수 행은 한 시설로 병합한다."""
    from g2b_costdb.classify import extract_facility_name, clean_notice_name
    cfg = load_config()
    raw = pd.DataFrame([
        dict(bidNtceNo="S1", bidNtceOrd="000", bidNtceNm="송파 창의혁신 공공주택 건설사업 [1,2단지]", ntceKindNm="일반공고", bidNtceDt="2024-01-05 10:00:00",
             dminsttNm="서울주택도시개발공사", presmptPrce="900000000000", mainCnsttyNm="건축공사"),
        dict(bidNtceNo="S2", bidNtceOrd="000", bidNtceNm="송파 창의혁신 공공주택 건설 전기공사", ntceKindNm="일반공고", bidNtceDt="2024-01-06 10:00:00",
             dminsttNm="서울주택도시개발공사", presmptPrce="20000000000", mainCnsttyNm="전기공사"),
        dict(bidNtceNo="S3", bidNtceOrd="000", bidNtceNm="마곡 창의 공공주택 건설공사", ntceKindNm="일반공고", bidNtceDt="2024-01-06 10:00:00",
             dminsttNm="서울주택도시개발공사", presmptPrce="1000000000", mainCnsttyNm="건축공사"),
        dict(bidNtceNo="H1", bidNtceOrd="000", bidNtceNm="화성 함백산추모공원 건립공사", ntceKindNm="일반공고", bidNtceDt="2024-02-05 10:00:00",
             dminsttNm="화성도시공사", presmptPrce="80000000000", mainCnsttyNm="건축공사"),
        dict(bidNtceNo="H2", bidNtceOrd="000", bidNtceNm="화성 함백산추모공원 장례식장 증축공사", ntceKindNm="일반공고", bidNtceDt="2025-02-05 10:00:00",
             dminsttNm="경기도 화성시", presmptPrce="5000000000", mainCnsttyNm="건축공사"),
        dict(bidNtceNo="H3", bidNtceOrd="000", bidNtceNm="함백산추모공원 봉안당 건립공사", ntceKindNm="일반공고", bidNtceDt="2025-03-05 10:00:00",
             dminsttNm="강원도 정선군", presmptPrce="3000000000", mainCnsttyNm="건축공사"),
    ])
    std = discover.standardize(raw, cfg, None)
    fac = extract_facility_name("송파 창의혁신 공공주택 건설사업 [1,2단지]", "공공주택")
    assert fac == "송파 창의 공공주택", fac                       # '혁신'은 상태어로 지워져 원 공고명의 부분 문자열이 아님
    assert clean_notice_name("′23∼′24년 영덕군 스마트팜 온실 신축") == "영덕군 스마트팜 온실 신축"
    assert extract_facility_name("월배공원 재조성사업(야구장 등) 공사", "야구장") == "월배공원"
    rev = pd.DataFrame([
        dict(시설ID="F0001", 검수_시설명=fac, **{"검수_별칭(;구분)": ""}, 검수_수요기관="서울주택도시개발공사", 표2_대분류="공동주택", 표2_중분류="공동주택", 검색어="공공주택"),
        dict(시설ID="F0002", 검수_시설명="화성 함백산추모공원", **{"검수_별칭(;구분)": ""}, 검수_수요기관="화성도시공사", 표2_대분류="묘지관련시설", 표2_중분류="-", 검색어="추모공원"),
        dict(시설ID="F0002", 검수_시설명="화성 함백산추모공원", **{"검수_별칭(;구분)": "함백산추모공원"}, 검수_수요기관="경기도 화성시", 표2_대분류="묘지관련시설", 표2_중분류="-", 검색어="추모공원"),
    ])
    hits = discover.research_by_facility(std, rev)
    by = hits.groupby("시설ID")["공고번호"].apply(set).to_dict()
    assert by["F0001"] == {"S1", "S2"}, ("예비 검색(어절 간격 허용)으로 찾고 '마곡 창의'는 제외", by)
    assert by["F0002"] == {"H1", "H2"}, ("같은 시설ID 두 행 → 수요기관 합집합으로 병합, 정선군 공고는 제외", by)
    assert (hits[hits["시설ID"] == "F0002"]["시설명"] == "화성 함백산추모공원").all()


def test_research_other_orders():
    """같은 공고번호의 다른 차수가 공고명이 바뀌어 시설 이름으로 안 잡혀도 보충된다 — 뒤 차수(금액 변경·취소)를 놓치면 옛 차수가 최신으로 남는다."""
    from g2b_costdb import dedup
    cfg = load_config()
    inst = "가상광역시"
    mk = lambda no, o, nm, price, kind="일반공고", dt="2024-03-05 10:00:00": dict(
        bidNtceNo=no, bidNtceOrd=o, bidNtceNm=nm, ntceKindNm=kind, bidNtceDt=dt, dminsttNm=inst, presmptPrce=str(price), mainCnsttyNm="건축공사")
    raw = pd.DataFrame([
        mk("N1", "000", "가나문화회괸 건립공사", 1_000_000_000, dt="2024-03-01 10:00:00"),               # 오타 → 이름으로 못 찾는 앞 차수
        mk("N1", "001", "가나문화회관 건립공사", 1_000_000_000, kind="변경공고", dt="2024-03-03 10:00:00"),
        mk("N2", "000", "가나문화회관 증축공사", 2_000_000_000, dt="2024-04-01 10:00:00"),
        mk("N2", "001", "가나 문화예술회관 증축공사", 2_600_000_000, kind="변경공고", dt="2024-04-09 10:00:00"),     # 이름이 바뀐 뒤 차수(금액도 바뀜)
        mk("N3", "000", "가나문화회관 별관 리모델링공사", 700_000_000, dt="2024-05-01 10:00:00"),
        mk("N3", "001", "가나 문예회관 별관 리모델링공사", 700_000_000, kind="취소공고", dt="2024-05-08 10:00:00"),    # 이름이 바뀐 취소 차수
        mk("N4", "000", "가나문화회관 외벽 보수공사", 90_000_000, dt="2024-06-01 10:00:00"),               # 유지보수 → 빠짐
        mk("N4", "001", "가나 문예회관 외벽 보수공사", 95_000_000, kind="변경공고", dt="2024-06-03 10:00:00"),      # 그 공고의 다른 차수도 넣지 않는다
        mk("N5", "000", "가나문화회관 주차장 신축공사", 400_000_000, dt="2024-07-01 10:00:00"),
        mk("N5", "001", "가나문화회관 주차장 유지보수공사", 400_000_000, kind="변경공고", dt="2024-07-03 10:00:00"),  # 이름으로 찾았지만 유지보수라 뺀 차수 → 다시 넣지 않는다
        mk("N6", "000", "다라체육관 건립공사", 5_000_000_000, dt="2024-08-01 10:00:00"),                 # 다른 시설(차수 하나) — 그대로
        mk("Z9", "000", "가나 문화예술회관 증축공사", 1, dt="2024-04-09 10:00:00"),                    # 이름이 같아도 공고번호가 다르면 보충하지 않는다
    ])
    std = discover.standardize(raw, cfg, None)
    rev = pd.DataFrame([
        dict(시설ID="F0001", 검수_시설명="가나문화회관", **{"검수_별칭(;구분)": ""}, 검수_수요기관=inst, 표2_대분류="문화 및 집회시설", 표2_중분류="공연장", 검색어="문화회관"),
        dict(시설ID="F0002", 검수_시설명="다라체육관", **{"검수_별칭(;구분)": ""}, 검수_수요기관=inst, 표2_대분류="운동시설", 표2_중분류="체육관", 검색어="체육관"),
        dict(시설ID="F0003", 검수_시설명="가나 문화예술회관", **{"검수_별칭(;구분)": ""}, 검수_수요기관=inst, 표2_대분류="문화 및 집회시설", 표2_중분류="공연장", 검색어="문화예술회관"),
    ])
    hits = discover.research_by_facility(std, rev)
    keys = lambda fid: sorted(hits.loc[hits["시설ID"] == fid, "공고키"])
    assert keys("F0001") == ["N1-000", "N1-001", "N2-000", "N2-001", "N3-000", "N3-001", "N5-000"], keys("F0001")
    assert keys("F0002") == ["N6-000"] and keys("F0003") == ["N2-000", "N2-001", "Z9-000"], (keys("F0002"), keys("F0003"))
    assert not hits.duplicated(["시설ID", "공고키"]).any() and list(hits.columns) == list(discover.research_by_facility(std, rev.iloc[[1]]).columns)
    f1 = hits[hits["시설ID"] == "F0001"].set_index("공고키")
    assert f1.loc["N2-001", "시설명"] == "가나문화회관" and f1.loc["N2-001", "사업유형"] == f1.loc["N2-000", "사업유형"] == "증축", f1.loc[["N2-000", "N2-001"], ["시설명", "사업유형"]]
    assert f1.loc["N1-000", "표2_중분류"] == "공연장" and f1.loc["N3-001", "검색어"] == "문화회관" and f1.loc["N3-001", "수요기관_대표"] == inst
    hits["공종"], hits["공종근거"] = "건축", "주공종명"
    hist, latest, lg0 = dedup.dedup_latest(hits, 0.30)
    sup_log = sorted((l["공고번호"], l["비고"][:5], l["판정"], "취소공고라" in l["비고"]) for l in lg0 if l["항목"] == "최신 차수 보충")
    assert sup_log == [("N2", "F0001", "참고", False), ("N3", "F0001", "참고", True)], sup_log      # 최신 차수가 보충으로 바뀐 공고는 06 에 남는다(앞 차수만 보충된 N1·F0003 의 N2 는 아님)
    assert set(hits.loc[hits["차수보충"] == "Y", "공고키"]) == {"N1-000", "N2-001", "N3-001", "N2-000"} and set(hits["차수보충"]) == {"", "Y"}
    assert "차수보충" in hist.columns and "_shown" not in hist.columns and "_shown" not in latest.columns
    assert hist.set_index(["시설ID", "공고키"]).loc[("F0001", "N2-001"), "공고명"] == "가나 문화예술회관 증축공사", "표시 공고명은 그 차수의 것"
    l1 = latest[latest["시설ID"] == "F0001"]
    assert sorted(l1["공고키"]) == ["N1-001", "N2-001", "N5-000"], ("최신 차수는 N2-001, 취소된 N3 은 합산하지 않음", sorted(l1["공고키"]))
    assert float(l1.set_index("공고키").loc["N2-001", "추정가격"]) == 2_600_000_000
    assert sorted(latest.loc[latest["시설ID"] == "F0003", "공고키"]) == ["N2-001", "Z9-000"]
    n3 = hist[(hist["시설ID"] == "F0001") & (hist["공고번호"] == "N3")].set_index("공고키")["대표선정사유"].to_dict()
    assert n3["N3-000"].startswith("취소됨") and n3["N3-001"] == "취소공고", n3
    # 보충할 차수가 없으면 결과가 그대로다(행·순서·열)
    one = discover.research_by_facility(std, rev.iloc[[1]])
    same, n, newer = discover.add_other_orders(one, std, set(zip(one["시설ID"], one["공고키"])))
    assert n == 0 and newer == [] and same.equals(one)
    # 반환값: 보충 행 수와 '최신 차수가 바뀐 공고' 목록
    base = hits[(hits["시설ID"] == "F0001") & hits["공고키"].isin(["N1-001", "N2-000", "N3-000", "N5-000"])].drop(columns=["공종", "공종근거"])
    out, n, newer = discover.add_other_orders(base, std, {("F0001", k) for k in ("N1-001", "N2-000", "N3-000", "N4-000", "N5-000", "N5-001")})
    assert n == 3 and newer == ["N2(F0001: 000→001)", "N3(F0001: 000→001)"], (n, newer)
    assert discover.add_other_orders(base.iloc[0:0], std)[1] == 0 and discover.add_other_orders(base, std.iloc[0:0])[1] == 0
    # 공고번호가 빈 행끼리는 같은 공고가 아니다(빈 값·결측이 서로 묶여 엉뚱한 공고가 딸려 오지 않게)
    blank = base.head(2).copy()
    blank["공고번호"] = ["", None]
    std_b = pd.concat([std, std.head(3).assign(공고번호=["", None, "nan"], 공고키=["-000", "x-000", "nan-000"])], ignore_index=True)
    assert discover.add_other_orders(blank, std_b)[1] == 0


def test_research_other_orders_more():
    """차수 보충의 나머지 동작: 물려받는 차수(이름으로 찾아 남은 가장 늦은 차수), 사업유형은 자기 공고명이 아니라 물려받은 값, 여러 차수·여러 시설, 순서·index·dtype, 로그."""
    import io
    import logging
    from g2b_costdb import dedup
    cfg = load_config()
    inst = "가상광역시"
    mk = lambda no, o, nm, price, kind="일반공고", dt="2024-03-05 10:00:00": dict(
        bidNtceNo=no, bidNtceOrd=o, bidNtceNm=nm, ntceKindNm=kind, bidNtceDt=dt, dminsttNm=inst, presmptPrce=str(price), mainCnsttyNm="건축공사")
    raw = pd.DataFrame([
        mk("Z1", "000", "가나문화회괸 리모델링공사", 900_000_000, dt="2024-07-01 10:00:00"),                    # 못 찾는 앞 차수(오타)
        mk("Z1", "001", "가나문화회관 리모델링공사", 900_000_000, kind="변경공고", dt="2024-07-02 10:00:00"),     # 찾음: 리모델링
        mk("Z1", "002", "가나문화회관 증축공사", 950_000_000, kind="변경공고", dt="2024-07-03 10:00:00"),         # 찾음: 증축 ← 물려줄 차수
        mk("Z1", "003", "가나 문예회관 시설공사", 980_000_000, kind="변경공고", dt="2024-07-05 10:00:00"),        # 못 찾음(자기 이름으로는 미분류)
        mk("Z1", "004", "가나 문예회관 지붕 방수공사", 990_000_000, kind="변경공고", dt="2024-07-09 10:00:00"),   # 못 찾음(자기 이름으로는 유지보수) → 그래도 증축으로 들어간다
        mk("A2", "000", "가나문화회관 별관 건립공사", 500_000_000, dt="2024-08-01 10:00:00"),
        mk("A2", "001", "가나 문예회관 별관 건립공사", 520_000_000, kind="변경공고", dt="2024-08-03 10:00:00"),
        mk("B3", "000", "가나문화회관 주차장 신축공사", 400_000_000, dt="2024-09-01 10:00:00"),                  # 찾음(남음)
        mk("B3", "001", "가나문화회관 주차장 유지보수공사", 410_000_000, kind="변경공고", dt="2024-09-03 10:00:00"),   # 찾음 → 유지보수라 빠짐
        mk("B3", "002", "가나 문예회관 주차장 유지보수공사", 420_000_000, kind="변경공고", dt="2024-09-05 10:00:00"),  # 못 찾음
        mk("C4", "000", "가나문화회관 증축 전기공사", 100_000_000, dt="2024-10-01 10:00:00"),
        mk("C4", "001", "가나 문화예술회관 증축 전기공사", 110_000_000, kind="변경공고", dt="2024-10-05 10:00:00"),
        mk("C4", "002", "가나 예술의전당 건립 전기공사", 120_000_000, kind="변경공고", dt="2024-10-09 10:00:00"),
    ])
    std = discover.standardize(raw, cfg, None)
    row = lambda fid, nm, kw: dict(시설ID=fid, 검수_시설명=nm, **{"검수_별칭(;구분)": ""}, 검수_수요기관=inst, 표2_대분류="문화 및 집회시설", 표2_중분류="공연장", 검색어=kw)
    rev = pd.DataFrame([row("F0001", "가나문화회관", "문화회관"), row("F0003", "가나 문화예술회관", "문화예술회관"), row("F0005", "가나 예술의전당", "예술의전당")])
    buf = io.StringIO()
    h = logging.StreamHandler(buf)
    lg = logging.getLogger(discover.__name__)
    old_level = lg.level
    lg.addHandler(h); lg.setLevel(logging.INFO)
    try:
        hits = discover.research_by_facility(std, rev)
    finally:
        lg.removeHandler(h); lg.setLevel(old_level)
    f1 = hits[hits["시설ID"] == "F0001"].set_index("공고키")
    # ① 물려받는 값은 이름으로 찾아 남은 차수 가운데 가장 늦은 것(Z1-002 증축)의 것 — 앞 차수(000)·뒤 차수(003·004) 모두. 자기 공고명의 분류(미분류·유지보수)는 쓰지 않는다
    assert sorted(f1.index[f1.index.str.startswith("Z1")]) == ["Z1-000", "Z1-001", "Z1-002", "Z1-003", "Z1-004"], sorted(f1.index)
    assert f1.loc["Z1-001", "사업유형"] == "리모델링" and f1.loc["Z1-002", "사업유형"] == "증축"
    assert list(f1.loc[["Z1-000", "Z1-003", "Z1-004"], "사업유형"]) == ["증축"] * 3, f1.loc[["Z1-000", "Z1-003", "Z1-004"], "사업유형"].tolist()
    assert std.set_index("공고키").loc["Z1-004", "사업유형"] == "유지보수", "이 단언의 전제: 보충 행 자기 공고명의 분류는 물려받은 값과 다르다"
    assert (f1["표2_대분류"] == "문화 및 집회시설").all() and (f1["표2_중분류"] == "공연장").all() and (f1["시설명"] == "가나문화회관").all()
    # ② 보충 행 자기 값(공고명·금액·공고종류·공고일시)은 전량 표의 그 차수 것
    assert f1.loc["Z1-004", "공고명"] == "가나 문예회관 지붕 방수공사" and float(f1.loc["Z1-004", "추정가격"]) == 990_000_000 and f1.loc["Z1-003", "공고일시"].startswith("2024-07-05")
    # ③ 이름으로 찾은 가장 늦은 차수(B3-001)가 유지보수라 빠진 공고: 그 차수는 다시 넣지 않고, 못 찾은 뒤 차수(B3-002)는 남은 차수(B3-000)의 값을 물려받아 들어간다
    #    (지금 동작을 고정한 것 — 'B3-002 도 넣지 않는다'로 바꾸려면 이 단언과 add_other_orders 설명을 함께 고친다)
    assert sorted(f1.index[f1.index.str.startswith("B3")]) == ["B3-000", "B3-002"] and f1.loc["B3-002", "사업유형"] == f1.loc["B3-000", "사업유형"]
    # ④ 한 공고를 세 시설이 서로 다른 차수로 찾은 경우: 시설마다 모든 차수, 자기 시설의 값
    for fid, kw in (("F0001", "문화회관"), ("F0003", "문화예술회관"), ("F0005", "예술의전당")):
        g = hits[(hits["시설ID"] == fid) & (hits["공고번호"] == "C4")]
        assert sorted(g["공고키"]) == ["C4-000", "C4-001", "C4-002"] and (g["검색어"] == kw).all(), (fid, sorted(g["공고키"]), g["검색어"].tolist())
    assert set(hits.loc[(hits["시설ID"] == "F0005") & (hits["공고번호"] == "C4"), "사업유형"]) == {"신축"}      # 예술의전당 '건립' — 다른 시설의 '증축' 과 섞이지 않는다
    # ⑤ 결과 꼴: 이름으로 찾은 행이 앞에 그대로(값·순서), 보충 행은 뒤에. index 는 겹치지 않고, 열 dtype 은 보충 전과 같다. 입력은 바뀌지 않는다
    keep = hits[[(f, k) in {("F0001", "Z1-001"), ("F0001", "Z1-002"), ("F0001", "A2-000"), ("F0001", "B3-000"), ("F0001", "C4-000"), ("F0003", "C4-001")}
                 for f, k in zip(hits["시설ID"], hits["공고키"])]].sort_values(["시설ID", "공고키"], ascending=False)
    keep.index = [5, 5, 3, 2, 1, 0]                                                                     # 겹치는 index
    before = keep.copy()
    out, n, newer = discover.add_other_orders(keep, std, {("F0001", "B3-001")})
    assert keep.equals(before) and "_no" not in std.columns, "입력 표를 바꾸지 않는다"
    assert out.iloc[:len(keep)].reset_index(drop=True).equals(keep.reset_index(drop=True)), "기존 행은 앞에 그대로"
    assert out.index.is_unique and list(out.index) == list(range(len(out))) and hits.index.is_unique
    assert list(out.columns) == list(keep.columns) and {c: str(t) for c, t in out.dtypes.items()} == {c: str(t) for c, t in keep.dtypes.items()}
    assert n == 9 and len(out) == len(keep) + 9, n                                                      # Z1 3 + A2 1 + B3 1 + C4 (F0001 2 + F0003 2)
    # ⑥ newer: (시설, 공고)마다 한 줄, '→' 뒤는 보충한 가장 늦은 차수, 정렬된 목록
    assert newer == ["A2(F0001: 000→001)", "B3(F0001: 000→002)", "C4(F0001: 000→002)", "C4(F0003: 001→002)", "Z1(F0001: 002→004)"], newer
    # ⑦ 전량 표에 같은 행이 두 번 있어도(같은 공고가 두 달 파일에 실린 경우) 결과가 같다
    out2, n2, newer2 = discover.add_other_orders(keep, pd.concat([std, std, std.iloc[::-1]], ignore_index=True), {("F0001", "B3-001")})
    assert n2 == n and newer2 == newer and not out2.duplicated(["시설ID", "공고키"]).any() and out2.equals(out)
    # ⑧ 차수는 숫자로 견준다('9' < '10') — 표준화를 거치지 않은 표를 넘겨도
    s9 = pd.DataFrame({"공고번호": ["Q"] * 3, "공고차수": ["9", "10", "11"], "공고키": ["Q-9", "Q-10", "Q-11"], "공고명": ["a", "b", "c"], "사업유형": ["신축", "증축", "미분류"]})
    d9 = s9.iloc[[0, 1]].assign(시설ID="F9", 시설명="x")
    o9, n9, newer9 = discover.add_other_orders(d9, s9)
    assert n9 == 1 and newer9 == ["Q(F9: 010→011)"] and o9.set_index("공고키").loc["Q-11", "사업유형"] == "증축", (n9, newer9)
    # ⑨ research 가 보충 건수와 최신 차수가 바뀐 공고를 로그로 알린다
    msg = buf.getvalue()
    assert "다른 차수 11건 보충" in msg and "최신 차수가 바뀌는 공고 5건" in msg and "Z1(F0001: 002→004)" in msg, msg
    # ⑩ dedup 까지: 보충된 최신 차수가 대표가 되고 사업유형은 물려받은 값
    hits["공종"], hits["공종근거"] = "건축", "주공종명"
    hist, latest, lgz = dedup.dedup_latest(hits, 0.30)
    zl = [l["비고"] for l in lgz if l["항목"] == "최신 차수 보충" and l["공고번호"] == "Z1"]
    assert len(zl) == 1 and "최신 차수 004 '가나 문예회관 지붕 방수공사'" in zl[0] and "가장 늦은 차수는 002 '가나문화회관 증축공사'" in zl[0], zl      # 묶음에는 검색으로 찾은 '가장 늦은' 차수의 공고명
    z = latest[(latest["시설ID"] == "F0001") & (latest["공고번호"] == "Z1")]
    assert list(z["공고키"]) == ["Z1-004"] and z["사업유형"].iloc[0] == "증축" and float(z["추정가격"].iloc[0]) == 990_000_000
    assert set(hist.loc[(hist["시설ID"] == "F0001") & (hist["공고번호"] == "Z1"), "사업유형"]) == {"증축"}
    # 보충된 최신 차수의 공고명에 시설 이름이 없어도(사업명 변경) 세부사업은 이름으로 찾은 차수의 공고명으로 나눈다 → 같은 시설의 공종 공고와 한 프로젝트로 남는다
    raw3 = pd.DataFrame([mk("K1", "000", "가나문화회관 건립공사", 9_000_000_000, dt="2024-03-01 10:00:00"),
                         mk("K1", "001", "가나문화예술회관 건립공사", 9_500_000_000, kind="변경공고", dt="2024-03-20 10:00:00"),
                         mk("K2", "000", "가나문화회관 건립 전기공사", 900_000_000, dt="2024-09-10 10:00:00"),
                         mk("K3", "000", "가나문화회관 건립 소방공사", 500_000_000, dt="2024-09-15 10:00:00")])
    h3 = discover.research_by_facility(discover.standardize(raw3, cfg, None), rev.iloc[[0]])
    h3["공종"], h3["공종근거"] = h3["공고번호"].map({"K1": "건축", "K2": "전기", "K3": "소방"}), "주공종명"
    hist3, latest3, logs3 = dedup.dedup_latest(h3, 0.30)
    assert sorted(latest3["공고키"]) == ["K1-001", "K2-000", "K3-000"] and latest3["프로젝트ID"].nunique() == 1 and (hist3["세부사업수"] == 1).all(), latest3[["프로젝트ID", "공고키"]]
    ren = [l for l in logs3 if l["항목"] == "최신 차수 보충"]                                              # 그렇게 묶은 공고는 06 에 남는다(금액이 바뀌었는지 볼 단서)
    assert len(ren) == 1 and ren[0]["공고번호"] == "K1" and ren[0]["판정"] == "참고" and "000 '가나문화회관 건립공사'" in ren[0]["비고"] and "최신 차수 001 '가나문화예술회관 건립공사'" in ren[0]["비고"], ren
    # 보충은 사업 묶음을 바꾸지 않는다: 보충 차수를 뺀 입력(예전 재검색 결과)과 프로젝트ID 가 같다
    h3_old = h3[h3["차수보충"] != "Y"].drop(columns=["차수보충"])
    hist3_old, latest3_old, _ = dedup.dedup_latest(h3_old, 0.30)
    pid = lambda h: h.drop_duplicates(["시설ID", "공고번호"]).set_index("공고번호")["프로젝트ID"].to_dict()
    assert pid(hist3) == pid(hist3_old) and sorted(latest3_old["공고키"]) == ["K1-000", "K2-000", "K3-000"]
    # 검색으로 찾은 행의 공고명은 묶음에서도 그대로 쓴다 — 별칭으로 잡힌 새 이름의 공고는 그 이름으로 묶인다(공고명에 검수_시설명이 없다는 것만으로 옛 공고명으로 바꾸지 않는다).
    # ① 사업명이 바뀐 뒤 공종 공고가 새 이름으로 발주된 경우: 한 프로젝트 ② 새 이름으로 같은 공고명의 재공고: 앞 공고는 대체됨(두 번 합산하지 않는다)
    rev_al = pd.DataFrame([dict(시설ID="F0001", 검수_시설명="가나문화회관", **{"검수_별칭(;구분)": "가나아트센터"}, 검수_수요기관=inst, 표2_대분류="문화 및 집회시설", 표2_중분류="공연장", 검색어="문화회관")])
    raw_a = pd.DataFrame([mk("A1", "000", "가나문화회관 건립공사", 9_000_000_000, dt="2024-03-01 10:00:00"),
                          mk("A1", "001", "가나아트센터 건립공사", 9_500_000_000, kind="변경공고", dt="2024-03-20 10:00:00"),
                          mk("A2", "000", "가나아트센터 건립 전기공사", 900_000_000, dt="2024-09-26 10:00:00"),
                          mk("A3", "000", "가나아트센터 건립 소방공사", 500_000_000, dt="2024-09-27 10:00:00")])
    ha = discover.research_by_facility(discover.standardize(raw_a, cfg, None), rev_al)
    ha["공종"], ha["공종근거"] = ha["공고번호"].map({"A1": "건축", "A2": "전기", "A3": "소방"}), "주공종명"
    hist_a, latest_a, logs_a = dedup.dedup_latest(ha, 0.30)
    assert (ha["차수보충"] == "").all() and sorted(latest_a["공고키"]) == ["A1-001", "A2-000", "A3-000"] and latest_a["프로젝트ID"].nunique() == 1, latest_a[["프로젝트ID", "공고키"]]
    assert not [l for l in logs_a if l["항목"] in ("최신 차수 보충", "세부사업 분리")]
    raw_b = pd.DataFrame([mk("B1", "000", "가나문화회관 건립 전기공사", 900_000_000, dt="2024-03-01 10:00:00"),
                          mk("B1", "001", "가나아트센터 건립 전기공사", 900_000_000, kind="변경공고", dt="2024-03-10 10:00:00"),
                          mk("B2", "000", "가나아트센터 건립 전기공사", 900_000_000, dt="2024-04-09 10:00:00")])
    hb = discover.research_by_facility(discover.standardize(raw_b, cfg, None), rev_al)
    hb["공종"], hb["공종근거"] = "전기", "주공종명"
    hist_b, latest_b, _ = dedup.dedup_latest(hb, 0.30)
    assert list(latest_b["공고키"]) == ["B2-000"] and float(latest_b["추정가격"].astype(float).sum()) == 900_000_000, latest_b[["프로젝트ID", "공고키", "대표선정사유"]]
    assert hist_b.set_index("공고키").loc["B1-001", "대표선정사유"].startswith("대체됨"), hist_b[["공고키", "대표선정사유"]]
    # 보충 차수의 공종이 그 차수 자기 공고명에서 나온 것이면(주공종명·면허제한 근거가 아님) 검색으로 찾은 차수의 공종을 따른다 — 바뀐 공고명에서 '(전기)' 가 빠졌다고
    # 재공고 판정·같은 공종 합산이 달라지지 않게. E1 은 같은 공고명으로 다시 난 E2 로 대체된 그대로(두 번 합산하지 않음)이고, 주공종명에서 나온 공종은 그 차수 것을 쓴다
    mt = lambda no, o, nm, dt, trade, basis, sup="", price=900_000_000: dict(시설ID="F1", 시설명="가나문화회관", 공고번호=no, 공고차수=o, 공고키=f"{no}-{o}", 공고명=nm, 공고일시=dt,
                                                                       사업유형="신축", 공종=trade, 공종근거=basis, 추정가격=price, 공고종류="일반", 이전공고번호="", 차수보충=sup)
    tr = [mt("M1", "000", "가나문화회관 건립공사", "2024-03-01 10:00:00", "건축", "주공종명", price=9_000_000_000),
          mt("E1", "000", "가나문화회관 건립공사(전기)", "2024-03-05 10:00:00", "전기", "공고명(꼬리표)"), mt("E2", "000", "가나문화회관 건립공사(전기)", "2024-04-05 10:00:00", "전기", "공고명(꼬리표)")]
    ht0, lt0, _ = dedup.dedup_latest(pd.DataFrame(tr), 0.30)
    ht1, lt1, _ = dedup.dedup_latest(pd.DataFrame(tr + [mt("E1", "001", "가나아트홀 건립공사", "2024-03-05 10:00:00", "건축", "공고명", "Y")]), 0.30)
    assert sorted(lt0["공고키"]) == sorted(lt1["공고키"]) == ["E2-000", "M1-000"] and float(lt1["추정가격"].astype(float).sum()) == 9_900_000_000, lt1[["공고키", "공종", "대표선정사유"]]
    e1 = ht1.set_index("공고키").loc["E1-001"]
    assert e1["공종"] == "전기" and e1["공종근거"] == "공고명(꼬리표)" and e1["대표선정사유"].startswith("대체됨") and e1["공고명"] == "가나아트홀 건립공사", e1.to_dict()
    ht2, lt2, _ = dedup.dedup_latest(pd.DataFrame(tr + [mt("E1", "001", "가나아트홀 건립공사", "2024-03-05 10:00:00", "정보통신", "주공종명", "Y")]), 0.30)
    assert ht2.set_index("공고키").loc["E1-001", "공종"] == "정보통신" and sorted(lt2["공고키"]) == ["E1-001", "E2-000", "M1-000"], "주공종명(구조화 항목)에서 나온 공종은 최신 차수 것"
    # 보충 차수가 늘었다고 다른 공고가 물려받는 사업유형(미분류 → 가까운 건축 공고, 없으면 시설의 최빈 유형)이 달라지지 않는다(실자료 F4643 꼴):
    # 미분류 P1 은 가장 가까운 건축 공고 R1 과 730일 넘게 떨어져 최빈 유형(신축 3 : 리모델링 2)을 따른다. R1·R2 에 보충 차수가 붙어도 신축 그대로
    mr = lambda no, o, nm, dt, typ, trade, sup="": dict(시설ID="F1", 시설명="가나승화원", 공고번호=no, 공고차수=o, 공고키=f"{no}-{o}", 공고명=nm, 공고일시=dt, 사업유형=typ,
                                                    공종=trade, 공종근거="주공종명", 추정가격=100_000_000, 공고종류="일반", 이전공고번호="", 차수보충=sup)
    rows = [mr("P1", "000", "가나승화원 녹지관리공사", "2019-03-05 10:00:00", "미분류", "기타"),
            mr("R1", "000", "가나승화원 통제실 리모델링 건축공사", "2021-09-29 10:00:00", "리모델링", "건축"), mr("R2", "000", "가나승화원 통제실 리모델링 전기공사", "2021-09-29 11:00:00", "리모델링", "전기"),
            mr("Q1", "000", "가나승화원 제2관 건립공사", "2023-03-01 10:00:00", "신축", "건축"), mr("Q2", "000", "가나승화원 제2관 건립 전기공사", "2023-03-02 10:00:00", "신축", "전기"),
            mr("Q3", "000", "가나승화원 제2관 건립 소방공사", "2023-03-03 10:00:00", "신축", "소방")]
    extra = [mr("R1", "001", "가나추모원 통제실 개선 건축공사", "2021-10-05 10:00:00", "리모델링", "건축", "Y"), mr("R2", "001", "가나추모원 통제실 개선 전기공사", "2021-10-05 11:00:00", "리모델링", "전기", "Y"),
             mr("R2", "002", "가나추모원 통제실 개선 전기공사(정정)", "2021-10-07 11:00:00", "리모델링", "전기", "Y")]
    typ = lambda rr: dedup.dedup_latest(pd.DataFrame(rr))[0].drop_duplicates("공고번호").set_index("공고번호")
    t0, t1 = typ(rows), typ(rows + extra)
    assert t0.loc["P1", "사업유형"] == "신축" and t1.loc["P1", "사업유형"] == "신축" and dict(t0["프로젝트ID"]) == dict(t1["프로젝트ID"]), (dict(t0["프로젝트ID"]), dict(t1["프로젝트ID"]))
    assert t1.loc["R2", "사업유형"] == "리모델링" and sorted(dedup.dedup_latest(pd.DataFrame(rows + extra))[1]["공고키"]) == ["P1-000", "Q1-000", "Q2-000", "Q3-000", "R1-001", "R2-002"]
    assert latest3.set_index("공고키").loc["K1-001", "공고명"] == "가나문화예술회관 건립공사" and float(latest3["추정가격"].astype(float).sum()) == 10_900_000_000
    # 차수가 'inf' 로 읽히는 값이어도 로그 문구 때문에 멈추지 않는다
    s_inf = pd.DataFrame({"공고번호": ["Q"] * 2, "공고차수": ["000", "inf"], "공고키": ["Q-000", "Q-inf"], "공고명": ["a", "b"]})
    assert discover.add_other_orders(s_inf.iloc[[0]].assign(시설ID="F9", 시설명="x"), s_inf)[2] == ["Q(F9: 000→inf)"]


def test_supplement_marker_decisions():
    """차수보충 = Y 행에는 묶음용 공고명(같은 시설·같은 공고에서 검색으로 찾은 가장 늦은 차수의 공고명)을 쓴다 — 그 공고명이 닿는 판정마다 하나씩(변이 검사에서 살아남던 줄)."""
    from g2b_costdb import dedup
    m = lambda fid, no, o, nm, dt, sup="", trade="건축", kind="일반", typ="신축", price=1_000_000_000, fac="가나문화회관": dict(
        시설ID=fid, 시설명=fac, 공고번호=no, 공고차수=o, 공고키=f"{no}-{o}", 공고명=nm, 공고일시=dt, 사업유형=typ, 공종=trade, 공종근거="주공종명",
        추정가격=price, 공고종류=kind, 이전공고번호="", 차수보충=sup)
    run = lambda rows: dedup.dedup_latest(pd.DataFrame(rows), 0.30)
    note = lambda logs, fid, no: [l["비고"] for l in logs if l["항목"] == "최신 차수 보충" and l["공고번호"] == no and l["비고"].startswith(fid + ":")]
    # ① '가장 늦은 차수'는 차수 번호로 고른다(공고일시가 앞서 적힌 차수여도). 단계토큰도 그 공고명에서 — 보충 행 자기 공고명의 '(3단계)' 가 아니다
    h, _, g = run([m("F1", "N1", "000", "가나문화회관 건립공사(1단계)", "2024-03-10 10:00:00"), m("F1", "N1", "001", "가나문화회관 건립공사(2단계)", "2024-03-05 10:00:00", kind="변경"),
                   m("F1", "N1", "002", "가나아트홀 건립공사(3단계)", "2024-03-20 10:00:00", "Y", kind="변경")])
    r = h.set_index("공고키").loc["N1-002"]
    assert "가장 늦은 차수는 001 '가나문화회관 건립공사(2단계)'" in note(g, "F1", "N1")[0], note(g, "F1", "N1")
    assert r["단계토큰"] == "2단계" and r["프로젝트키"].endswith("|건축|2단계") and r["공고명"] == "가나아트홀 건립공사(3단계)", r.to_dict()
    # ② 같은 공고가 두 시설에 있으면 시설마다 자기 시설에서 검색으로 찾은 차수의 공고명을 쓴다(다른 시설에서 찾은 차수가 더 늦어도)
    two = [("F1", "가나문화회관", ["", "Y", "Y"]), ("F2", "다라책마루", ["Y", "", "Y"])]
    names = ["가나문화회관 및 다라도서관 건립 통신공사", "다라책마루 외 1개소 건립 통신공사", "문화시설 2개소 건립 통신공사"]
    _, _, g = run([m(fid, "M1", f"00{k}", names[k], f"2024-08-0{k + 1} 10:00:00", sup[k], trade="정보통신", fac=fac) for fid, fac, sup in two for k in range(3)])
    assert "가장 늦은 차수는 000 '가나문화회관 및 다라도서관 건립 통신공사'" in note(g, "F1", "M1")[0] and "가장 늦은 차수는 001 '다라책마루 외 1개소 건립 통신공사'" in note(g, "F2", "M1")[0], \
        (note(g, "F1", "M1"), note(g, "F2", "M1"))
    # ③ 제안 안내 공고 판정도 묶음용 공고명으로: 검색으로 찾은 차수가 제안 안내 공고면 보충된 최신 차수도 합산하지 않고, 보충 행 공고명만 제안 안내 꼴이면 합산한다
    h, lat, _ = run([m("F1", "P1", "000", "가나문화회관 건립공사 방수공법 선정 기술제안서 제출 안내 공고", "2024-03-01 10:00:00"),
                     m("F1", "P1", "001", "가나아트홀 건립 방수공사", "2024-03-05 10:00:00", "Y", kind="변경"),
                     m("F1", "Q1", "000", "가나문화회관 별관 증축공사", "2024-04-01 10:00:00", typ="증축"),
                     m("F1", "Q1", "001", "가나아트홀 방수공법 선정 기술제안서 제출 안내 공고", "2024-04-05 10:00:00", "Y", kind="변경", typ="증축"),
                     m("F1", "K1", "000", "가나문화회관 건립공사", "2024-03-02 10:00:00", price=9_000_000_000)])
    assert sorted(lat["공고키"]) == ["K1-000", "Q1-001"] and h.set_index("공고키").loc["P1-001", "대표선정사유"] == dedup.PROPOSAL_REASON, h[["공고키", "최신여부", "대표선정사유"]]
    # ④ 취소 차수가 보충 행이면(공고명이 바뀐 취소공고) 그 취소공고를 놓을 세부사업·대체공고번호는 묶음용 공고명으로 고른다 → 같은 이름으로 다시 낸 N2
    h, _, _ = run([m("F1", "C1", "000", "가나문화회관 별관 건립공사", "2024-01-10 10:00:00"), m("F1", "C1", "001", "가나아트홀 신관 건립공사", "2024-01-20 10:00:00", "Y", kind="취소"),
                   m("F1", "N2", "000", "가나문화회관 별관 건립공사", "2024-03-01 10:00:00"), m("F1", "N9", "000", "가나문화회관 신관 건립공사", "2024-03-02 10:00:00", price=3_000_000_000)])
    x = h.set_index("공고키")
    assert x.loc["C1-001", "프로젝트ID"] == x.loc["N2-000", "프로젝트ID"] != x.loc["N9-000", "프로젝트ID"] and x.loc["C1-001", "대체공고번호"] == "N2", x[["프로젝트ID", "대체공고번호"]]
    # ⑤ 표시 없는 취소공고를 놓을 세부사업은 유효 공고의 묶음용 공고명에서 낸 구분어('별관')로 고른다 — 표시 공고명('신관')이 아니다
    h, _, _ = run([m("F1", "C1", "000", "가나문화회관 별관 소방공사", "2024-01-10 10:00:00", trade="소방", kind="취소", typ="리모델링", price=300_000_000),
                   m("F1", "N1", "000", "가나문화회관 별관 리모델링", "2024-02-01 10:00:00", typ="리모델링"), m("F1", "N1", "001", "가나아트홀 신관 리모델링", "2024-02-05 10:00:00", "Y", kind="변경", typ="리모델링"),
                   m("F1", "N9", "000", "가나문화회관 본관 소방공사", "2024-02-10 10:00:00", trade="소방", typ="리모델링", price=300_000_000)])
    x = h.set_index("공고키")
    assert x.loc["C1-000", "프로젝트ID"] == x.loc["N1-001", "프로젝트ID"] != x.loc["N9-000", "프로젝트ID"], x["프로젝트ID"].to_dict()
    # ⑥ 취소공고의 대체공고번호: 같은 세부사업·같은 공종에 남은 공고 가운데 묶음용 공고명이 가장 비슷한 공고(N1) — 표시 공고명 '가나아트홀 신축' 으로 견주면 N9 가 된다
    h, _, _ = run([m("F1", "C1", "000", "가나문화회관 건립공사", "2024-01-10 10:00:00", kind="취소"),
                   m("F1", "N1", "000", "가나문화회관 건립공사", "2024-02-01 10:00:00"), m("F1", "N1", "001", "가나아트홀 신축", "2024-02-05 10:00:00", "Y", kind="변경"),
                   m("F1", "N9", "000", "가나문화회관 건립공사 중 창호공사", "2024-02-20 10:00:00", price=150_000_000)])
    assert h.set_index("공고키").loc["C1-000", "대체공고번호"] == "N1", h[["공고키", "대표선정사유", "대체공고번호"]]
    # ⑦ 보충 행의 공종이 면허제한 업종(구조화 항목)에서 나온 것이면 그 차수 것을 쓴다(주공종명과 같게). 공고명에서 나온 것만 검색으로 찾은 차수의 공종을 따른다
    lic = [dict(m("F1", "E1", "000", "가나문화회관 건립공사(전기)", "2024-03-05 10:00:00", trade="전기", price=900_000_000), 공종근거="공고명(꼬리표)"),
           dict(m("F1", "E1", "001", "가나아트홀 건립공사", "2024-03-06 10:00:00", "Y", trade="소방", kind="변경", price=900_000_000), 공종근거="면허제한")]
    e = run(lic)[0].set_index("공고키").loc["E1-001"]
    assert e["공종"] == "소방" and e["공종근거"] == "면허제한", e.to_dict()
    # ⑧ add_other_orders 의 '최신 차수가 바뀌는 공고'는 차수 번호가 더 클 때만('001' 과 '1' 처럼 같은 번호의 다른 표기는 아니다)
    s1 = pd.DataFrame({"공고번호": ["Q", "Q"], "공고차수": ["001", "1"], "공고키": ["Q-001", "Q-1"], "공고명": ["a", "b"]})
    assert discover.add_other_orders(s1.iloc[[0]].assign(시설ID="F9", 시설명="x"), s1)[1:] == (1, [])


def test_attach_merge_and_redo():
    from g2b_costdb.pipeline import _merge_attachment_slots, _redo_keys
    rep = {"공고번호": "X1", "공고키": "X1-002", "첨부URL1": "http://x/정정.hwp", "첨부파일명1": "정정공고서.hwp", "첨부URL2": "", "첨부파일명2": "", "현장설명서URL1": ""}
    o1 = {"공고번호": "X1", "공고키": "X1-001", "첨부URL1": "http://x/정정.hwp", "첨부파일명1": "정정공고서.hwp", "첨부URL2": "http://x/공고문.hwp", "첨부파일명2": "입찰공고문.hwp", "현장설명서URL1": "http://x/spt1"}
    o0 = {"공고번호": "X1", "공고키": "X1-000", "첨부URL1": "http://x/공고문.hwp", "첨부파일명1": "입찰공고문.hwp", "첨부URL2": "http://x/도면.zip", "첨부파일명2": "도면.zip", "현장설명서URL1": "http://x/spt1"}
    m, n = _merge_attachment_slots(rep, [o1, o0], attach_max=2)
    urls = [m[f"첨부URL{i}"] for i in range(1, n + 1) if m[f"첨부URL{i}"]]
    assert urls == ["http://x/정정.hwp", "http://x/공고문.hwp", "http://x/도면.zip"] and m["현장설명서URL1"] == "http://x/spt1" and n >= 3, (urls, n)
    latest = pd.DataFrame({"공고키": ["A-000", "B-000", "C-000", "D-000"], "공종": ["건축", "건축", "전기", "건축"]})
    docs = {"A-000": {"신뢰도": "high", "연면적_m2": 1000}, "B-000": {"신뢰도": "high"}, "C-000": {"신뢰도": "high"}, "D-000": {"신뢰도": "low", "연면적_m2": 500}, "E-000": {"_error": "x"}}
    assert _redo_keys(latest, docs) == ["B-000", "D-000"], "건축인데 연면적 없음(B), 신뢰도 low(D). 전기(C)는 연면적 없어도 제외"


def test_awards():
    """낙찰정보 병합: config 매핑이 없으면 후보 필드명 사용, 공고번호+차수 → 공고번호 폴백, 같은 공고의 여러 행은 최신 개찰일시 행."""
    from g2b_costdb.collect import award_column
    cfg = load_config()
    aw = pd.DataFrame([
        dict(bidNtceNo="A1", bidNtceOrd="001", scsbidAmt="900", scsbidRate="90.0", bidwinnrNm="갑", prtcptCnum="3", opengDt="2024-02-01 11:00:00"),
        dict(bidNtceNo="A2", bidNtceOrd=0, scsbidAmt="500", scsbidRate="80%", bidwinnrNm="을", prtcptCnum="2", opengDt="2024-03-01 11:00:00"),
        dict(bidNtceNo="A2", bidNtceOrd="000", scsbidAmt="520", scsbidRate="83", bidwinnrNm="병", prtcptCnum="4", opengDt="2024-02-10 11:00:00"),
        dict(bidNtceNo="Z9", bidNtceOrd="000", scsbidAmt="1", scsbidRate="1", bidwinnrNm="무", prtcptCnum="1", opengDt="2024-03-01 11:00:00"),
        dict(bidNtceNo="A3", bidNtceOrd="000", scsbidAmt=None, scsbidRate=None, bidwinnrNm="", prtcptCnum="", opengDt=""),
    ])
    assert award_column(cfg, aw, "award_amt") == "scsbidAmt" and award_column(cfg, aw, "award_rate") == "scsbidRate", "config 매핑(sucsfbid*) 없음 → 후보"
    assert award_column(cfg, aw, "award_bidder_bizno") is None and award_column(cfg, aw.iloc[0:0], "award_amt") is None
    raw = pd.DataFrame([
        dict(bidNtceNo="A1", bidNtceOrd="000", bidNtceNm="가상군 문화예술회관 건립공사", ntceKindNm="일반공고", bidNtceDt="2024-01-05 10:00:00",
             dminsttNm="가상군", presmptPrce="1000", mainCnsttyNm="건축공사"),
        dict(bidNtceNo="A1", bidNtceOrd="001", bidNtceNm="가상군 문화예술회관 건립공사 [변경]", ntceKindNm="변경공고", bidNtceDt="2024-01-08 10:00:00",
             dminsttNm="가상군", presmptPrce="1000", mainCnsttyNm="건축공사"),
        dict(bidNtceNo="A2", bidNtceOrd="000", bidNtceNm="가상군 문화예술회관 건립 전기공사", ntceKindNm="일반공고", bidNtceDt="2024-01-06 10:00:00",
             dminsttNm="가상군", presmptPrce="600", mainCnsttyNm="전기공사"),
        dict(bidNtceNo="A3", bidNtceOrd="000", bidNtceNm="가상군 문화예술회관 건립 소방공사", ntceKindNm="일반공고", bidNtceDt="2024-01-06 10:00:00",
             dminsttNm="가상군", presmptPrce="100", mainCnsttyNm="소방공사"),
        dict(bidNtceNo="B1", bidNtceOrd="000", bidNtceNm="다른군 문화예술회관 건립공사", ntceKindNm="일반공고", bidNtceDt="2024-02-05 10:00:00",
             dminsttNm="다른군", presmptPrce="2000", mainCnsttyNm="건축공사"),
    ])
    std = discover.standardize(raw, cfg, None, awards=aw).set_index("공고키")
    assert std.loc["A1-001", "낙찰금액_API"] == 900 and std.loc["A1-001", "낙찰자"] == "갑" and std.loc["A1-001", "참가업체수"] == 3
    assert std.loc["A1-000", "낙찰금액_API"] == 900, "차수가 다른 행은 공고번호로 폴백"
    assert std.loc["A2-000", "낙찰자"] == "을" and std.loc["A2-000", "낙찰금액_API"] == 500 and std.loc["A2-000", "낙찰률_API"] == 80.0, \
        "같은 공고 2행 → 개찰일시가 최신인 행(정수 차수 0 도 000 으로), '80%' → 80.0"
    assert pd.isna(std.loc["B1-000", "낙찰금액_API"]) and pd.isna(std.loc["A3-000", "낙찰금액_API"]), "낙찰 없음·금액 없는 낙찰 행은 공란"
    assert std["낙찰개찰일시"].loc["A2-000"] == "2024-03-01 11:00:00"
    std0 = discover.standardize(raw, cfg, None)
    assert "낙찰금액_API" in std0.columns and std0["낙찰금액_API"].isna().all(), "낙찰 미사용이어도 컬럼은 존재(빈 값)"
    # collect_awards 가 만든 parquet(_award_amt 선계산) 형태로 들어와도 같은 결과
    aw2 = aw.assign(_award_amt=aw["scsbidAmt"].map(parse_amount), _award_rate=pd.to_numeric(aw["scsbidRate"].astype(str).str.replace("%", ""), errors="coerce"))
    std2 = discover.standardize(raw, cfg, None, awards=aw2).set_index("공고키")
    assert std2.loc["A2-000", "낙찰금액_API"] == 500 and std2.loc["A1-000", "낙찰률_API"] == 90.0


def test_error_summary():
    from g2b_costdb.pipeline import _error_summary
    docs = {"A": {"_error": "응답 해석 실패(stop=max_tokens): JSON 아님 | 응답: {\"공사명\": ..."},
            "B": {"_error": "응답 해석 실패(stop=max_tokens): JSON 아님 | 응답: 다른 내용"},
            "C": {"_error": "batch errored: invalid_request_error x"}, "D": {"공사명": "정상"}}
    out = _error_summary(docs)
    assert "추출 실패 3건" in out and "    2  응답 해석 실패(stop=max_tokens): JSON 아님" in out and "batch errored" in out, out
    assert _error_summary({"D": {"공사명": "정상"}}) == ""


def test_output_schema_limits():
    """구조화 출력 스키마: union 타입 파라미터 0개(API 한도 16), 선택 파라미터 0개(한도 24), 빈 문자열 → None."""
    from g2b_costdb.extract_llm import output_schema, normalize_doc
    sc = output_schema()
    unions = [k for k, v in sc["properties"].items() if isinstance(v.get("type"), list) or "anyOf" in v]
    assert not unions, unions
    assert set(sc["required"]) == set(sc["properties"]) and sc["additionalProperties"] is False
    d = normalize_doc({"연면적_m2": "13731.3", "추정가격_원": "23,215,938,082", "지상층수": "", "공사명": "", "근거문구": [{"항목": "연면적_m2", "원문": "연면적 13,731.3㎡"}]})
    assert d["연면적_m2"] == 13731.3 and d["추정가격_원"] == 23215938082 and d["지상층수"] is None and d["근거문구"] == {"연면적_m2": "연면적 13,731.3㎡"}


def test_extract_and_verify():
    doc = extract_llm.normalize_doc({"추정가격_원": "19,100,000,000", "기초금액_원": 21010000000, "연면적_m2": "9,850.5",
                                     "근거문구": [{"항목": "연면적_m2", "원문": "연면적 9,850㎡"}], "지상층수": "4"})
    assert doc["추정가격_원"] == 19100000000 and doc["연면적_m2"] == 9850.5 and doc["지상층수"] == 4
    assert doc["근거문구"] == {"연면적_m2": "연면적 9,850㎡"}
    logs = extract_llm.cross_verify("N1", {"추정가격": float("nan"), "기초금액": "21,010,000,000"}, doc)
    by = {l["항목"]: l for l in logs}
    assert by["추정가격"]["판정"] == "참고" and by["기초금액"]["판정"] == "정상"
    assert extract_llm.jsonable(float("nan")) is None
    assert extract_llm._parse_json('```json\n{"a": 1}\n```\n참고 {b} }')["a"] == 1
    ft = extract_llm.focus_text("x" * 30000 + "추정가격 123원" + "y" * 30000 + "공사기간 720일" + "z" * 30000, 20000)
    assert "추정가격 123원" in ft and "공사기간 720일" in ft and len(ft) <= 20500
    import numpy as np
    assert json.dumps({"a": extract_llm.jsonable(np.int64(5))}) == '{"a": 5}'
    from g2b_costdb.pipeline import _logs_frame
    lf = _logs_frame([{"공고번호": "N1", "항목": "사업유형 일치", "API값": "신축", "문서값": "리모델링", "판정": "경고", "비고": "x"},
                      {"공고번호": 7, "항목": "추정가격", "API값": 1.0, "문서값": "2", "차이": 1, "차이율": 1.0, "판정": "오류"}])
    assert lf["API값"].dtype == float and lf["비고"].iloc[0] == "API값=신축; 문서값=리모델링; x" and lf["문서값"].iloc[1] == 2.0
    est = extract_llm.estimate_cost({"k1": "가" * 30000, "k2": "나" * 100000}, ["k1", "k2", "k3"], {"model": "claude-sonnet-5", "max_input_chars": 60000})
    assert est["docs"] == 2 and est["usd"] > 0
    schema = extract_llm.output_schema()
    assert schema["additionalProperties"] is False and set(schema["required"]) == set(schema["properties"])
    # API값과 문서값이 5% 넘게 다르면('오류') 비율로 알 수 있는 원인을 비고에 적는다. 판정과 숫자는 그대로
    note = lambda a, d, item="추정가격": next(l for l in extract_llm.cross_verify("H", {item: a}, {item + "_원": d}) if l["항목"] == item)
    assert note(61_100_000, 67_210_000)["판정"] == "오류" and "부가세 포함" in note(61_100_000, 67_210_000)["비고"]
    assert "API 추정가격이 부가세 포함" in note(127_644_000, 116_040_000)["비고"] and "10배" in note(53_200_000, 532_000_000)["비고"]
    assert "부가세를 뺀 금액(추정가격)" in note(1_100_000, 1_000_000, "기초금액")["비고"] and "부가세 포함 금액을 읽은" not in note(1_000_000, 1_100_000, "기초금액")["비고"]
    # 같은 공고의 다른 금액을 함께 본다: 문서 추정가격×1.1 = API 기초금액이면 어긋난 쪽은 API 추정가격(실자료 20200702645·20200817832) — '문서 오독' 으로 적지 않는다
    both = lambda ap, ab, dp, db: {l["항목"]: l for l in extract_llm.cross_verify("H", {"추정가격": ap, "기초금액": ab}, {"추정가격_원": dp, "기초금액_원": db})}
    for ap, ab, dp, db in ((53_200_000, 585_200_000, 532_000_000, 585_200_000), (127_644_000, 127_644_000, 116_040_000, 127_644_000)):
        r = both(ap, ab, dp, db)["추정가격"]
        assert r["판정"] == "오류" and "API 추정가격은 맞지 않음" in r["비고"] and "오독" not in r["비고"] and "추정가격 기준 합계 확인" in r["비고"], r
    assert "부가세 포함 금액을 읽은" in both(179_272_727, 197_200_000, 197_200_000, None)["추정가격"]["비고"], "API 쪽이 서로 맞으면 문서가 부가세 포함 금액을 읽은 것"
    part = {l["항목"]: l for l in extract_llm.cross_verify("H", {"추정가격": 290_210_000, "기초금액": 293_547_000, "부가세_API": 3_337_000}, {"추정가격_원": 266_860_909, "기초금액_원": 293_547_000})}
    assert "API 추정가격은 맞지 않음" not in part["추정가격"]["비고"] and "두 쪽 모두 같은 기초금액과 맞음" in part["추정가격"]["비고"], "API 추정가격+부가세 = 기초금액(부분면세)이면 API 쪽을 의심하지 않는다(실자료 20230734989)"
    # 문서 기초금액 = 추정가격×1.1 인데 API 기초금액이 다르면(조달청 조사금액 기준 등) 변경공고·오독이 아니다(실자료 '오류' 기초금액 42행 가운데 34행)
    r = both(1_327_400_000, 1_365_143_000, 1_327_400_000, 1_460_140_000)
    assert r["기초금액"]["판정"] == "오류" and "추정가격+부가세" in r["기초금액"]["비고"] and r["추정가격"]["판정"] == "정상"
    assert "API 기초금액 = API 추정가격" in both(1_000_000, 1_000_000, 1_000_000, 1_100_000)["기초금액"]["비고"], "기초금액 행의 ×1.1 은 '문서가 부가세 포함 금액을 읽음' 이 아니다"
    # 04 의 부가세·합계가 문서 기초금액을 따르는 경우(API 기초금액 없음 + 문서 기초금액이 추정가격의 100~110.5%)에는 '04 는 API값 사용' 이라고 적지 않는다(실자료 20200629585)
    r = both(55_947_170_000, None, 52_100_340_000, 57_372_500_000)["추정가격"]
    assert r["판정"] == "오류" and "04 는 API값 사용" not in r["비고"] and "문서 기초금액을 따름" in r["비고"], r
    assert extract_llm._doc_base_in_04(1000, 1200, 1050) and not extract_llm._doc_base_in_04(1000, 1100, 1300) and not extract_llm._doc_base_in_04(1000, 1200, 1050, 100)
    r = both(2_118_130_000, 2_329_943_000, 1_090_909_091, 1_200_000_000)
    assert all("같은 비율" in r[k]["비고"] and "5% 넘게" in r[k]["비고"] for k in ("추정가격", "기초금액")), r
    assert "1/1,000" in note(1_500_000_000, 1_500_000, "기초금액")["비고"] and "5% 넘게" in note(1000, 1300)["비고"]
    assert note(1000, 1040)["판정"] == "경고" and note(1000, 1040)["비고"] == "" and note(1000, 1000)["비고"] == "", "5% 이하에는 붙이지 않는다"
    assert all("04 는 API값 사용" in extract_llm._gap_hint(100, d) for d in (110, 1000, 50)) and extract_llm._gap_hint(0, 5) == "" and extract_llm._gap_hint(5, 0) == ""
    # ── 한 줄 변이 점검으로 찾은 빠진 단언: _doc_base_in_04 의 경계·가지(build_excel 04 수식의 okd·okv 와 같은 값이어야 한다)
    D = extract_llm._doc_base_in_04
    assert D(1000, None, 1000) and D(1000, None, 1105) and not D(1000, None, 999) and not D(1000, None, 1106) and not D(1000, None, 1108), "문서 기초금액 100~110.5%(양 끝 포함)"
    assert D(1000, 0, 1050, 50) and D(1000, None, 1050, 50), "① API 기초금액이 없거나 0 이면 부가세가 있어도 기초금액 기준 합계는 문서 기초금액"
    assert not D(1000, 1100, 1050) and not D(1000, 1200, 1300) and not D(0, None, 1050) and not D(None, None, 1050) and not D(1000, None, None)
    assert not D(1000, 1200, 1050, None, 100) and D(1000, 1200, 1050, 106) and D(1000, 1200, 1050, None, 106) and D(1000, 1200, 1050, -1), "② 는 쓸 수 있는 부가세(0~10.5%)가 없을 때만"
    assert D(1000, 1110, 1050, 60) and not D(1000, 1110, 1050, 99) and not D(1000, 1130, 1050, 60), "9.9% 미만 부가세는 기초금액이 추정가격×1.1±1% 일 때만 버린다(그 기초금액은 API 값이 먼저)"
    # _gap_hint: 음수·허용오차·가지별 문구
    assert extract_llm._gap_hint(-100, -110) == "" and extract_llm._gap_hint(100, -110) == ""
    assert "5% 넘게" in note(1000, 1115)["비고"] and "부가세 포함" in note(1_000_000, 1_101_000)["비고"], "×1.1 로 보는 폭은 ±0.2%"
    assert "추정가격+부가세" in both(1_000_000, 1_200_000, 1_000_000, 1_101_000)["기초금액"]["비고"] and "추정가격+부가세" not in both(1_000_000, 1_200_000, 1_000_000, 1_111_000)["기초금액"]["비고"]
    rb = both(1_327_400_000, 1_365_143_000, 1_327_400_000, 1_460_140_000)["기초금액"]["비고"]
    assert "04 는 API값 사용" in rb, rb
    ex0 = both(1_000_000, 1_000_000, 900_000, 1_100_000)["기초금액"]["비고"]
    assert "API 기초금액 = API 추정가격" in ex0 and "과세구분" in ex0, "면세 판단은 API 추정가격과 견준다(문서 추정가격이 아니라)"
    assert "04 는 API값 사용" in note(1_000_000, 1_100_000, "기초금액")["비고"] and "04 는 API값 사용" in note(1_100_000, 1_000_000, "기초금액")["비고"]
    assert "추정가격 기준 합계 확인" in note(127_644_000, 116_040_000)["비고"]
    n100, n1000, n10, n001 = note(1_000_000, 100_000_000)["비고"], note(1_500_000, 1_500_000_000)["비고"], note(1_000_000, 100_000)["비고"], note(1_000_000, 10_000)["비고"]
    assert "API값의 100배 — 자릿수" in n100 and "API값의 1,000배 — 단위(원/천원)" in n1000 and "API값의 1/10 — 자릿수" in n10 and "API값의 1/100 — 자릿수" in n001
    assert "API값의 1/1,000 — 단위(원/천원)" in note(1_500_000_000, 1_500_000, "기초금액")["비고"] and "API값의 10배 — 자릿수" in note(53_200_000, 532_000_000)["비고"]
    # 같은 비율 규칙은 두 금액이 모두 그 비율일 때만
    assert "같은 비율" not in both(2_000_000, 2_200_000, 2_000_000, 1_300_000)["기초금액"]["비고"] and "같은 비율" not in both(2_000_000, 2_200_000, 1_300_000, 2_200_000)["추정가격"]["비고"]
    r = both(2_118_130_000, 2_329_943_000, 1_090_909_091, 1_200_000_000)
    assert all("04 는 API값 사용" in r[k]["비고"] for k in ("추정가격", "기초금액")), r
    # 04 수식과 같은 값을 넘긴다: 추정가격(수식) = API값(0 보다 클 때) → 문서값, 부가세는 API·문서 둘 다
    cv = lambda api, doc, **kw: {l["항목"]: l["비고"] for l in extract_llm.cross_verify("H", api, doc, **kw)}
    assert "문서 기초금액을 따름" in cv({"기초금액": 1200}, {"추정가격_원": 1000, "기초금액_원": 1050})["기초금액"], "API 추정가격이 없으면 문서 추정가격이 04 추정가격(수식)"
    assert "문서 기초금액을 따름" in cv({"추정가격": 1000, "기초금액": 1200}, {"기초금액_원": 1050})["기초금액"]
    assert "04 는 API값 사용" in cv({"추정가격": 1000, "기초금액": 1200, "부가세_API": 100}, {"추정가격_원": 1000, "기초금액_원": 1050})["기초금액"], "API 부가세를 쓰는 행"
    assert "04 는 API값 사용" in cv({"추정가격": 1000, "기초금액": 1200}, {"추정가격_원": 1000, "기초금액_원": 1050, "부가가치세_원": 100})["기초금액"], "문서 부가세를 쓰는 행"
    assert "API값×1.2 " in cv({"추정가격": 1000}, {"추정가격_원": 1200}, vat=0.2)["추정가격"], "설정의 부가세율(verify.vat_ratio)을 쓴다"
    # ── 검토 반영: 문서 기초금액을 따르는 부분을 그대로 적는다(없는 API 기초금액을 'API값' 이라 적지 않는다). LibreOffice 로 다시 계산한 04 와 대조한 값
    assert (D(1000, None, 1050), D(1000, None, 1050, 50), D(1000, 0, 1050), D(1000, 1200, 1050), D(1000, 1100, 1300)) == ("both", "base", "both", "vat", "")
    assert "두 기준 모두" in both(55_947_170_000, None, 52_100_340_000, 57_372_500_000)["추정가격"]["비고"], "실자료 20200629585: API 기초금액 없음"
    assert "두 기준 모두" in both(61_100_000, None, 67_210_000, 67_210_000)["추정가격"]["비고"], "실자료 20240127469"
    n = cv({"추정가격": 1_000_000_000}, {"추정가격_원": 1_200_000_000, "기초금액_원": 1_100_000_000, "부가가치세_원": 100_000_000})["추정가격"]
    assert "금액합계_기초금액기준이 문서 기초금액을 따름" in n and "부가세(수식)" not in n, n
    n = cv({"기초금액": 1200}, {"추정가격_원": 1000, "기초금액_원": 1050})["기초금액"]
    assert "금액합계_기초금액기준은 API값이지만 부가세(수식)·금액합계_추정가격기준은 문서 기초금액을 따름" in n and "추정가격·기초금액 칸은 API값" not in n, n
    # ÷1.1 인데 API 추정가격×1.1 = API 기초금액이면 API 추정가격을 의심하지 않는다(같은 공고의 '기초금액≈추정가격×1.1' 행이 정상인데 'API 가 부가세 포함' 이라 적던 것)
    r = both(1_000_000_000, 1_100_000_000, 909_090_909, None)["추정가격"]["비고"]
    assert "API 쪽은 서로 맞음" in r and "API 추정가격이 부가세 포함" not in r, r
    # 문서 기초금액 = 추정가격×1.1 인데 API 기초금액이 조사금액·관급 차이로 볼 수 없게 다르면(0.75~1.35 밖) 기초금액 기준 합계를 확인하게 한다
    r = cv({"추정가격": 909_090_909, "기초금액": 100_000_000, "부가세_API": 90_909_091}, {"추정가격_원": 909_090_909, "기초금액_원": 1_000_000_000})["기초금액"]
    assert "크게 다름" in r and "기초금액 기준 합계 확인" in r and "조달청" not in r, r
    assert "조달청 조사금액" in both(1_000_000, 830_000, 1_000_000, 1_100_000)["기초금액"]["비고"] and "조달청 조사금액" in both(1_000_000, 1_470_000, 1_000_000, 1_100_000)["기초금액"]["비고"]
    assert "크게 다름" in both(1_000_000, 800_000, 1_000_000, 1_100_000)["기초금액"]["비고"] and "크게 다름" in both(1_000_000, 1_500_000, 1_000_000, 1_100_000)["기초금액"]["비고"]
    # API 기초금액이 없어도, 문서가 서로 맞고(추정가격×1.1 = 기초금액) API 추정가격 = 문서 부가세이면 어긋난 쪽은 API(실자료 20200702645 에서 API 기초금액만 뺀 경우)
    r = cv({"추정가격": 53_200_000}, {"추정가격_원": 532_000_000, "기초금액_원": 585_200_000, "부가가치세_원": 53_200_000})["추정가격"]
    assert "API 추정가격 = 문서 부가세" in r and "오독" not in r, r
    assert "10배" in cv({"추정가격": 53_200_000}, {"추정가격_원": 532_000_000, "기초금액_원": 585_200_000})["추정가격"], "문서 부가세가 없으면 판단하지 않는다"
    assert "5% 넘게" in cv({"추정가격": 169_409_090_911, "부가세_API": 16_940_909_089}, {"추정가격_원": 140_609_000_000, "기초금액_원": 154_670_500_000})["추정가격"], "문서가 서로 맞는 것만으로는 API 를 의심하지 않는다(실자료 R26BK01335412)"
    # 조달청 공고에서 추정가격 행이 'API 추정가격은 맞지 않음' 이면 '추정가격 기준 합계 참조' 안내에 단서를 붙인다(판정은 그대로 참고)
    lg = {l["항목"]: l for l in extract_llm.cross_verify("H", {"추정가격": 127_644_000, "기초금액": 127_644_000, "공고기관": "조달청 대전지방조달청"}, {"추정가격_원": 116_040_000, "기초금액_원": 127_644_000})}
    assert lg["기초금액≈추정가격×1.1"]["판정"] == "참고" and "API 추정가격이 어긋난 것이면" in lg["기초금액≈추정가격×1.1"]["비고"], lg
    lg = {l["항목"]: l for l in extract_llm.cross_verify("H", {"추정가격": 1000, "기초금액": 1075, "공고기관": "조달청"}, {"추정가격_원": 1000, "기초금액_원": 1075})}
    assert lg["기초금액≈추정가격×1.1"]["비고"].endswith("추정가격 기준 합계 참조"), "다른 조달청 공고의 안내는 그대로"
    # ── 독립 재검토(한 줄 변이)에서 살아남던 변이에 대한 단언
    assert D(1000, 0, 1100, 50) == "base", "API 기초금액 칸이 0 이면 04 의 bb 는 0 이다(문서 기초금액으로 바꿔 보지 않는다) → API 부가세 50 을 그대로 쓴다"
    assert D(1000, 1200, 1050, 105) == "" and D(1000, 1200, 1050, 0) == "", "부가세 10.5%(끝 포함)와 0 은 쓸 수 있는 부가세"
    assert D(1000, 1111, 1050, 50) == "vat" and D(1000, 1112, 1050, 50) == "", "9.9% 미만 부가세를 버리는 폭은 추정가격×1.1 의 ±1%(끝 포함)"
    assert "10배" in cv({"추정가격": 53_200_000}, {"추정가격_원": 532_000_000, "기초금액_원": 700_000_000, "부가가치세_원": 53_200_000})["추정가격"], "문서가 서로 맞지 않으면(추정가격×1.1 ≠ 기초금액) 'API 에 부가세액' 이라 적지 않는다"
    r = cv({"추정가격": 909_090_909, "기초금액": 100_000_000, "부가세_API": 90_909_091}, {"추정가격_원": 909_090_909, "기초금액_원": 1_000_000_000})["기초금액"]
    assert "API값÷문서값 = 0.1)" in r, r
    assert "1,000배" in note(1_000_000, 1_001_500_000)["비고"] and "1,000배" in note(1_000_000, 998_500_000)["비고"] and "1,000배" not in note(1_000_000, 1_003_000_000)["비고"], "1,000배로 보는 폭은 ±0.2%"
    assert "같은 비율(×0.515)" in both(2_118_130_000, 2_329_943_000, 1_090_909_091, 1_200_000_000)["추정가격"]["비고"], "비율은 문서값÷API값"
    assert "04 는 API값 사용" in ex0, ex0
    assert "04 는 API값 사용" in both(1_000_000_000, 1_100_000_000, 909_090_909, None)["추정가격"]["비고"] and "04 는 API값 사용" in note(127_644_000, 116_040_000)["비고"]
    lg = {l["항목"]: l for l in extract_llm.cross_verify("H", {"추정가격": 127_644_000, "기초금액": 127_644_000, "공고기관": "조달청 대전지방조달청"}, {"추정가격_원": 116_040_000, "기초금액_원": 127_644_000})}
    assert "추정가격 기준 합계가 틀린 쪽" in lg["기초금액≈추정가격×1.1"]["비고"], lg
    assert "04 의 추정가격·부가세·과세구분과 추정가격 기준 합계 확인" in lg["추정가격"]["비고"] and "04 의 추정가격·부가세나 추정가격 기준 합계가 틀린 쪽" in lg["기초금액≈추정가격×1.1"]["비고"], \
        "API 추정가격 = API 기초금액이고 부가세가 없으면 04 의 두 합계는 같고 틀리는 칸은 추정가격·부가세(0)·과세구분(면세)이다 → 그 칸도 함께 적는다: " + str(lg)
    for dp in (None, 1300):      # 추정가격 행에 다른 비고('한쪽만 존재', 5% 넘게 다름)가 있어도 조달청 안내에 단서를 붙이지 않는다
        lg = {l["항목"]: l for l in extract_llm.cross_verify("H", {"추정가격": 1000, "기초금액": 1075, "공고기관": "조달청"}, {"추정가격_원": dp, "기초금액_원": 1075})}
        assert lg["추정가격"]["비고"] and lg["기초금액≈추정가격×1.1"]["비고"].endswith("추정가격 기준 합계 참조"), lg
    # ── 독립 재검토 반영(H1~H10): 같은 공고의 다른 행과 엇갈리는 문구
    # H1: API 추정가격 + 부가세(API → 문서) = API 기초금액이면(면세·부분면세로 API 쪽이 서로 맞음) API 추정가격을 의심하는 문구를 쓰지 않는다 — '×1.1' 행은 '공고에 적힌 부가세와 일치'
    for api, doc in (({"추정가격": 1_000_000_000, "기초금액": 1_000_000_000, "부가세_API": 0}, {"추정가격_원": 909_090_909, "기초금액_원": 1_000_000_000}),
                     ({"추정가격": 1_000_000_000, "기초금액": 1_000_000_000}, {"추정가격_원": 909_090_909, "기초금액_원": 1_000_000_000, "부가가치세_원": 0}),
                     ({"추정가격": 290_210_000, "기초금액": 293_547_000, "부가세_API": 3_337_000}, {"추정가격_원": 266_860_909, "기초금액_원": 293_547_000})):
        r = cv(api, doc)
        assert "두 쪽 모두 같은 기초금액과 맞음" in r["추정가격"] and "API 추정가격이 부가세 포함" not in r["추정가격"] and "공고에 적힌 부가세와 일치" in r["기초금액≈추정가격×1.1"], r
    r = cv({"추정가격": 1_000_000_000, "기초금액": 1_040_000_000, "부가세_API": 40_000_000}, {"추정가격_원": 909_090_909})["추정가격"]
    assert "면세·부분면세로 서로 맞음" in r and "API 추정가격이 부가세 포함" not in r, r
    r = cv({"추정가격": 1_000_000_000}, {"추정가격_원": 909_090_909, "기초금액_원": 1_000_000_000, "부가가치세_원": 0})
    assert "문서 기초금액 — 두 쪽 모두 같은 기초금액과 맞음" in r["추정가격"] and "문서 기초금액을 따름" in r["추정가격"] and "공고에 적힌 부가세와 일치" in r["기초금액≈추정가격×1.1"], "API 기초금액이 없으면 '×1.1' 행처럼 문서 기초금액으로 본다: " + str(r)
    # H3: API 기초금액이 없고 문서 기초금액 = API 추정가격×1.1 이면('×1.1' 행 정상) API 추정가격을 의심하지 않는다. 화살표를 겹쳐 쓰지 않는다
    r = cv({"추정가격": 1_000_000_000}, {"추정가격_원": 909_090_909, "기초금액_원": 1_100_000_000})
    assert "문서 기초금액(API 추정가격이 기초금액과 맞음)" in r["추정가격"] and "API 추정가격이 부가세 포함" not in r["추정가격"] and r["기초금액≈추정가격×1.1"] == "", r
    r = cv({"추정가격": 1_000_000_000}, {"추정가격_원": 909_090_909, "기초금액_원": 1_000_000_000})["추정가격"]
    assert "→ 04 합계 확인, 추정가격 기준 합계 확인" in r and "확인 →" not in r, r
    # H2: 조달청 안내의 단서는 추정가격 행이 '추정가격 기준 합계 확인' 을 적은 모든 경우에 붙는다(÷1.1 문구 포함)
    lg = cv({"추정가격": 1_000_000_000, "기초금액": 1_020_000_000, "공고기관": "조달청 서울지방조달청"}, {"추정가격_원": 909_090_909, "기초금액_원": 1_020_000_000})
    assert "API 추정가격이 부가세 포함" in lg["추정가격"] and "추정가격 기준 합계가 틀린 쪽" in lg["기초금액≈추정가격×1.1"], lg
    # H4: 문서 기초금액 = 추정가격×1.1 이라도 API 추정가격 + 부가세 = API 기초금액이면 조사금액·관급 차이가 아니라 면세·부분면세
    r = cv({"추정가격": 1_000_000_000, "기초금액": 1_040_000_000, "부가세_API": 40_000_000}, {"추정가격_원": 1_000_000_000, "기초금액_원": 1_100_000_000})["기초금액"]
    assert "부가세율 4.0%" in r and "조달청 조사금액" not in r and "04 는 API값 사용" in r, r
    # H5: API 기초금액 = API 추정가격×1.1 이면 'API 기초금액이 부가세를 뺀 금액' 이라고 적지 않는다
    r = cv({"추정가격": 1_000_000_000, "기초금액": 1_100_000_000}, {"추정가격_원": 1_000_000_000, "기초금액_원": 1_210_000_000})["기초금액"]
    assert "API 쪽은 서로 맞음" in r and "부가세를 뺀 금액이거나" not in r, r
    assert "부가세를 뺀 금액이거나" in note(1_000_000, 1_100_000, "기초금액")["비고"], "API 추정가격을 모르면 두 원인을 함께 적는다"
    # H6: 같은 공고의 추정가격이 문서 기초금액과 다르면 '(추정가격)' 이라고 적지 않는다
    assert "(추정가격)" not in cv({"추정가격": 800_000_000, "기초금액": 1_100_000_000}, {"추정가격_원": 800_000_000, "기초금액_원": 1_000_000_000})["기초금액"]
    assert "(추정가격)" in cv({"추정가격": 1_000_000_000, "기초금액": 1_100_000_000}, {"추정가격_원": 1_000_000_000, "기초금액_원": 1_000_000_000})["기초금액"]
    # H7: 1,000 이상·0.001 미만의 비율을 지수 표기로 적지 않는다
    assert [extract_llm._ratio_txt(x) for x in (1004.4, 999.6, 999.4, 0.0909, 0.00005, 0.515)] == ["1,004", "1,000", "999", "0.0909", "1/20,000", "0.515"]
    r = cv({"추정가격": 1_234_567, "기초금액": 1_358_024}, {"추정가격_원": 1_240_000_000, "기초금액_원": 1_364_000_000})
    assert "같은 비율(×1,004)" in r["추정가격"] and "e+" not in r["추정가격"] + r["기초금액"], r
    r = cv({"추정가격": 1_000_000_000, "기초금액": 1_100_000_000_000}, {"추정가격_원": 1_000_000_000, "기초금액_원": 1_100_000_000})["기초금액"]
    assert "API값÷문서값 = 1,000)" in r and "문서 기초금액을 따름" in r, "H10: '크게 다름' 문구에도 04 의 부가세·추정가격 기준 합계가 문서 기초금액을 따르면 그것을 적는다: " + r
    # H8: 문서 부가세가 문서 추정가격의 10% 가 아니면 'API 추정가격 칸에 부가세액' 이라고 적지 않는다
    assert "부가세액이 들어간" not in cv({"추정가격": 900_000_000}, {"추정가격_원": 1_000_000_000, "기초금액_원": 1_100_000_000, "부가가치세_원": 900_000_000})["추정가격"]
    # H9: 문서 기초금액 = 추정가격 + 문서 부가세(부분면세)면 '변경공고·다른 금액' 이 아니다(실자료 20190819618). 문서 부가세 0 이나 API 추정가격이 없을 때는 판단하지 않는다
    r = cv({"추정가격": 3_908_982_000, "기초금액": 4_560_197_000}, {"추정가격_원": 3_908_982_000, "기초금액_원": 3_999_170_000, "부가가치세_원": 90_188_000})["기초금액"]
    assert "추정가격 + 문서 부가세(부가세율 2.3%)" in r and "변경공고" not in r and "04 는 API값 사용" in r, r
    assert "변경공고" in cv({"추정가격": 1000, "기초금액": 1200}, {"추정가격_원": 1000, "기초금액_원": 1000, "부가가치세_원": 0})["기초금액"]
    assert cv({"기초금액": 1200}, {"기초금액_원": 100, "부가가치세_원": 100})["기초금액"], "API 추정가격이 없어도 예외 없이"


def test_handoff_and_batch(tmp):
    """파일 인수인계(export/import)와 Message Batches 흐름(가짜 클라이언트)."""
    from types import SimpleNamespace as NS
    cfg_llm = {"model": "claude-sonnet-5", "max_input_chars": 60000, "structured_output": True}
    texts = {"K1-000": "연면적 1,000㎡ 추정가격 100,000,000원", "K2-000": "본문"}
    hints = {"K1-000": {"공고번호": "K1", "공고명": "가상 미술관 건립공사", "수요기관": "가상시"}}
    in_dir, out_dir = os.path.join(tmp, "llm_in"), os.path.join(tmp, "llm_out")
    assert extract_llm.export_prompts(["K1-000", "K2-000"], texts, hints, cfg_llm, in_dir, out_dir) == 2
    body = open(os.path.join(in_dir, "K1-000.txt"), encoding="utf-8").read()
    assert "llm_out/K1-000.json" in body and "연면적 1,000㎡" in body and os.path.exists(os.path.join(in_dir, "README_지시문.md"))
    open(os.path.join(out_dir, "K1-000.json"), "w", encoding="utf-8").write('```json\n{"연면적_m2": "1,000", "추정가격_원": 100000000, "근거문구": {"연면적_m2": "연면적 1,000㎡"}}\n```')
    open(os.path.join(out_dir, "K2-000.json"), "w", encoding="utf-8").write('{"공사명": null}')
    open(os.path.join(out_dir, "K9-000.json"), "w", encoding="utf-8").write('{"연면적_m2": 5}')
    open(os.path.join(out_dir, "bad.json"), "w", encoding="utf-8").write('{not json')
    docs, errors = extract_llm.import_results(out_dir, {"K1-000", "K2-000", "bad"})
    assert set(docs) == {"K1-000"} and docs["K1-000"]["연면적_m2"] == 1000.0 and docs["K1-000"]["_model"] == "manual/cowork"
    assert {fn for fn, _ in errors} == {"K2-000.json", "K9-000.json", "bad.json"}, errors
    # Batches: 제출 → 진행중 → 완료(성공 1, 오류 1)
    calls = {"create": 0, "retrieve": 0}

    class FakeBatches:
        def create(self, requests):
            calls["create"] += 1
            assert all("output_config" in r["params"] for r in requests) and requests[0]["custom_id"] == "K1-000"
            calls["max_tokens"] = {r["custom_id"]: r["params"]["max_tokens"] for r in requests}
            return NS(id="msgbatch_1", processing_status="in_progress")

        def retrieve(self, bid):
            calls["retrieve"] += 1
            st = "in_progress" if calls["retrieve"] == 1 else "ended"
            return NS(processing_status=st, request_counts=NS(processing=0, succeeded=1, errored=1, canceled=0, expired=0))

        def results(self, bid):
            ok = NS(custom_id="K1-000", result=NS(type="succeeded", message=NS(content=[NS(type="text", text='{"연면적_m2": 1000, "추정가격_원": 100000000, "근거문구": []}')],
                                                                                   usage=NS(input_tokens=500, output_tokens=80))))
            # 실제 SDK 형태: result.error = ErrorResponse(type='error', error=ErrorObject(type, message)) — 안쪽 원인이 메시지에 남아야 한다
            bad = NS(custom_id="K2-000", result=NS(type="errored", error=NS(type="error", error=NS(type="overloaded_error", message="Overloaded"))))
            return [ok, bad]
    client = NS(messages=NS(batches=FakeBatches()))
    sub = extract_llm.submit_batches(client, [(k, texts[k], hints.get(k, {})) for k in ("K1-000", "K2-000")], cfg_llm,
                                     max_tokens_by_key={"K2-000": 32000})
    assert sub[0]["id"] == "msgbatch_1" and sub[0]["keys"] == ["K1-000", "K2-000"] and calls["create"] == 1
    assert calls["max_tokens"] == {"K1-000": extract_llm.DEFAULT_MAX_TOKENS, "K2-000": 32000}, calls["max_tokens"]
    st, got, counts = extract_llm.fetch_batch(client, "msgbatch_1", "claude-sonnet-5")
    assert st == "in_progress" and got == {}
    st, got, counts = extract_llm.fetch_batch(client, "msgbatch_1", "claude-sonnet-5")
    assert st == "ended" and got["K1-000"]["연면적_m2"] == 1000.0 and got["K1-000"]["_usage_in"] == 500 and "_error" in got["K2-000"]
    assert got["K2-000"]["_error"] == "batch errored: overloaded_error Overloaded", got["K2-000"]
    assert counts["succeeded"] == 1 and counts["errored"] == 1
    # 이전에 잘린 건은 2배 max_tokens 로 재요청 (pipeline._truncated_retry_tokens)
    from g2b_costdb import pipeline as _pl
    docs_prev = {"K1-000": {"_error": "응답 해석 실패(stop=max_tokens): JSON 미검출 | 응답: "}, "K2-000": {"_error": "batch errored: api_error x"}}
    assert _pl._truncated_retry_tokens({"llm": {"max_tokens": 8000}}, docs_prev, ["K1-000", "K2-000", "K3-000"]) == {"K1-000": 16000}
    assert _pl._truncated_retry_tokens({"llm": {}}, {}, ["K1-000"]) == {}
    assert extract_llm.build_request_params("본문", {}, {"max_tokens": 8000}, max_tokens=24000)["max_tokens"] == 24000
    params = extract_llm.build_request_params("본문", {"공고명": "x"}, cfg_llm)
    assert params["model"] == "claude-sonnet-5" and params["output_config"]["format"]["type"] == "json_schema" and "본문" in params["messages"][0]["content"]


def test_included_trades():
    """한 계약에 묶인 공종 표기(금액은 나누지 않음): 배정 공종 + API 주공종명·부공종명 → 공종 라벨."""
    from g2b_costdb.classify import included_trades
    from g2b_costdb.build_excel import bundled_trades_by_project
    assert included_trades("건축공사", "토목공사업 / 조경공사업", "건축") == "건축·토목·조경"
    assert included_trades("전기공사", "", "전기") == "전기"
    assert included_trades("건축공사", "기계설비ㆍ가스공사업 / 토공사업", "건축") == "건축·기계설비·토공사", "라벨 없는 업종명은 '업' 만 떼고 유지"
    assert included_trades("토목건축공사업", "", "건축") == "건축", "토목건축공사업은 건축 하나로"
    assert included_trades(None, None, "기타") == "" and included_trades("nan", "nan", "소방") == "소방"
    tr = pd.DataFrame({"프로젝트ID": ["P1", "P1", "P1", "P2"], "공종": ["건축", "전기", "기계설비", "건축"],
                       "포함공종_API": ["건축·토목·조경", "전기", "기계설비·가스시설시공", "건축"]})
    assert bundled_trades_by_project(tr) == {"P1": "건축(+토목·조경); 기계설비(+가스시설시공)"}, bundled_trades_by_project(tr)
    assert bundled_trades_by_project(pd.DataFrame()) == {}
    from g2b_costdb.build_excel import _clean
    assert _clean("도장\x01·습식·방수·석공사업") == "도장·습식·방수·석공사업", "API 업종명의 제어문자는 Excel 셀에서 제거"
    assert _clean("a\ud800b") == "ab" and _clean("정상") == "정상"
    # 문서 쪽: 현장설명서 공사금액 표의 '건축, 토목, 기계, 조경' 셀(실제 HWP 추출 형태 — 셀마다 한 줄, 빈 줄 섞임)
    from g2b_costdb.classify import included_trades_from_text, normalize_trade_list, merge_trade_lists
    doc = ("2. 위 치 :\n\n - 서귀포시 성산읍\n\n 3. 공사금액\n\n   (단위: 원/ 부가가치세 포함)\n\n구 분\n\n총공사비\n\n도급비\n\n관급자재비\n\n"
           "건축, 토목, 기계, 조경\n\n9,394,053,000\n\n7,893,683,000\n\n1,500,370,000\n")
    assert included_trades_from_text(doc, "건축") == "건축·토목·기계설비·조경", included_trades_from_text(doc, "건축")
    assert included_trades_from_text("1. 공사개요\n분리발주 공종\n전기, 통신, 소방\n", "건축") == "", "분리발주 목록은 포함공종이 아님"
    assert included_trades_from_text("공사금액\n구분\n전기, 통신, 소방공사\n", "건축") == "", "배정 공종(건축)이 없는 목록은 채택하지 않음"
    assert included_trades_from_text("공사개요\n본 공사는 건축, 토목, 조경 등 모든 공종을 포함한 턴키 공사이며 상세는 별첨 참조\n", "건축") == "", "긴 문장은 채택하지 않음"
    assert included_trades_from_text("2. 공사금액(부가가치세 포함)\n구 분\n건축공사·기계설비공사·조경공사\n1,000원\n", "건축") == "건축·기계설비·조경"
    assert included_trades_from_text("", "건축") == ""
    assert normalize_trade_list("건축, 토목, 기계, 조경") == "건축·토목·기계설비·조경" and normalize_trade_list("건축공사 및 토목공사(철거 포함)") == "건축·토목·철거"
    assert merge_trade_lists("건축·토목", "건축·토목·기계설비·조경", assigned="건축") == "건축·토목·기계설비·조경"
    assert merge_trade_lists("", "", assigned="소방") == "소방" and merge_trade_lists("nan", None, assigned="기타") == ""
    # LLM 값이 있으면 규칙보다 우선, 없으면 텍스트 규칙으로 채움 (assemble_trade_table)
    from g2b_costdb.pipeline import assemble_trade_table
    latest = pd.DataFrame([{"공고키": "A-000", "프로젝트ID": "P1", "시설ID": "F1", "시설명": "s", "사업유형": "신축", "공종": "건축",
                            "공고번호": "A", "공고차수": "000", "공고명": "n", "공고일시": "2025-01-01", "수요기관": "d", "주공종명": "건축공사업", "부공종명": ""},
                           {"공고키": "B-000", "프로젝트ID": "P1", "시설ID": "F1", "시설명": "s", "사업유형": "신축", "공종": "건축",
                            "공고번호": "B", "공고차수": "000", "공고명": "n", "공고일시": "2025-01-01", "수요기관": "d", "주공종명": "건축공사업", "부공종명": "토목공사업"}])
    docs = {"A-000": {"포함공종_문서": "건축, 토목, 기계"}, "B-000": {}}
    tt = assemble_trade_table(latest, docs, {"B-000": doc})
    a, b = tt[tt["공고번호"] == "A"].iloc[0], tt[tt["공고번호"] == "B"].iloc[0]
    assert (a["포함공종_문서"], a["포함공종_출처"], a["포함공종(종합)"]) == ("건축·토목·기계설비", "LLM", "건축·토목·기계설비"), a.to_dict()
    assert (b["포함공종_API"], b["포함공종_문서"], b["포함공종_출처"], b["포함공종(종합)"]) == ("건축·토목", "건축·토목·기계설비·조경", "규칙", "건축·토목·기계설비·조경"), b.to_dict()


def test_trade_second_pass(tmp):
    """2차 추출(내역서 집계표의 공종별 금액): 집계표 행 추림, 스키마 한도, 정규화, 검증, 대상 선정, 04b 표, 본문 창."""
    from openpyxl import Workbook
    from g2b_costdb import attachments, pipeline
    from g2b_costdb.extract_llm import (trade_output_schema, normalize_trade_doc, cross_verify_trades, keyword_windows, _TRADE_KEYS,
                                        build_trade_request_params)
    # 집계표 시트 선택: 총괄집계표는 3,000행까지, 세부내역은 400행에서 끊김. 추림은 공종·합계·단위 행만
    d = os.path.join(tmp, "files", "B1")
    os.makedirs(d)
    wb = Workbook()
    ws = wb.active
    ws.title = "세부내역"
    for i in range(700):
        ws.append([f"세부{i}", 1, 100, 100])
    ws2 = wb.create_sheet("총괄집계표")
    ws2.append(["(단위: 천원)"])
    ws2.append(["공종", "합계"])
    for name, amt in (("건축공사", 20_000_000), ("기계설비공사", 3_000_000), ("조경공사", 1_000_000), ("소계", 24_000_000),
                      ("간접노무비", 1_000_000), ("부가가치세", 2_500_000), ("총공사비", 27_500_000)):
        ws2.append([name, amt])
    for i in range(450):
        ws2.append([f"세부{i}", 5])
    p = os.path.join(d, "공내역서_B1.xlsx")
    wb.save(p)
    txt = attachments.extract_xlsx(p)
    assert "세부699" not in txt and "세부399" in txt, "일반 시트는 400행 제한"
    assert txt.count("세부449") == 1, "집계표류 시트는 행 제한이 커서 끝까지"
    snip = attachments.cost_sheet_snippets(p)
    assert snip.startswith("## [공내역서_B1.xlsx] 시트: 총괄집계표") and "(단위: 천원)" in snip and "기계설비공사\t3000000" in snip and "총공사비\t27500000" in snip
    assert "세부1\t" not in snip and "세부내역" not in snip, snip
    assert attachments.find_cost_sheet_files(os.path.join(tmp, "files"), "B1") == [p] and attachments.find_cost_sheet_files(os.path.join(tmp, "files"), "B9") == []
    open(os.path.join(d, "관급자재목록.xlsx"), "wb").write(b"x")
    assert attachments.find_cost_sheet_files(os.path.join(tmp, "files"), "B1") == [p], "내역서류 이름만"
    # 스키마: union 0·optional 0, 객체 배열
    sch = trade_output_schema()
    u = o = 0

    def walk(x):
        nonlocal u, o
        if isinstance(x, dict):
            if "anyOf" in x or isinstance(x.get("type"), list):
                u += 1
            if x.get("type") == "object":
                o += len([k for k in x.get("properties", {}) if k not in set(x.get("required", []))])
            for v in x.values():
                walk(v)
        elif isinstance(x, list):
            for v in x:
                walk(v)
    walk(sch)
    assert (u, o) == (0, 0) and sch["properties"]["공종별금액"]["items"]["properties"]["공종"]["enum"][0] == "건축"
    prm = build_trade_request_params("발췌", {"공고번호": "B1"}, {"max_tokens": 8000}, max_tokens=16000)
    assert prm["max_tokens"] == 16000 and prm["output_config"]["format"]["schema"] == sch and "발췌" in prm["messages"][0]["content"]
    # 정규화·검증: 천원 단위 표를 원으로 읽은 경우 → ×1000 하면 API 와 일치 → 경고
    t = normalize_trade_doc({"합계_원": "27,500,000", "금액기준": "부가세포함", "공종별금액": [
        {"공종": "건축", "항목명_원문": "건축공사", "금액_원": "20000000"}, {"공종": "기계설비", "항목명_원문": "기계설비공사", "금액_원": "3000000"},
        {"공종": "조경", "항목명_원문": "조경공사", "금액_원": "1000000"}, {"공종": "간접비", "항목명_원문": "간접노무비", "금액_원": "1000000"},
        {"공종": "부가세", "항목명_원문": "부가가치세", "금액_원": "2500000"}, {"공종": "?", "항목명_원문": "빈값", "금액_원": ""}]})
    assert t["_공종합_원"] == 24_000_000 and t["합계_원"] == 27_500_000 and len(t["공종별금액"]) == 5
    logs = cross_verify_trades("B1", {"추정가격": 25_000_000_000, "기초금액": 27_500_000_000}, t)
    assert [l["항목"] for l in logs] == ["공종합=집계표합계", "집계표합계 기준"] and logs[0]["판정"] == "정상" and "천원 단위" in logs[1]["비고"], logs
    ok = cross_verify_trades("B2", {"추정가격": 25_000_000_000, "기초금액": 27_500_000_000}, normalize_trade_doc({"합계_원": "27500000000", "공종별금액": [
        {"공종": "건축", "항목명_원문": "a", "금액_원": "27500000000"}]}))
    assert ok[-1]["판정"] == "참고" and "부가세 포함" in ok[-1]["비고"]
    assert cross_verify_trades("B3", {}, normalize_trade_doc({"공종별금액": []}))[0]["비고"].startswith("집계표에 공종별 금액 없음")
    # 대상 선정: 건축 또는 포함공종 2개 이상. 04b 표. 분리발주 중복 경고. 본문 창은 Excel 블록 제외
    latest = pd.DataFrame([
        {"공고키": "A-000", "프로젝트ID": "P1", "시설ID": "F1", "시설명": "s", "사업유형": "신축", "공종": "건축", "공고번호": "A", "공고차수": "000", "공고명": "n", "수요기관": "d", "추정가격": 100, "주공종명": "건축공사업", "부공종명": ""},
        {"공고키": "E-000", "프로젝트ID": "P1", "시설ID": "F1", "시설명": "s", "사업유형": "신축", "공종": "전기", "공고번호": "E", "공고차수": "000", "공고명": "n", "수요기관": "d", "추정가격": 10, "주공종명": "전기공사업", "부공종명": ""},
        {"공고키": "M-000", "프로젝트ID": "P1", "시설ID": "F1", "시설명": "s", "사업유형": "신축", "공종": "기계설비", "공고번호": "M", "공고차수": "000", "공고명": "n", "수요기관": "d", "추정가격": 20, "주공종명": "기계설비공사업", "부공종명": "토목공사업"},
        {"공고키": "A-000", "프로젝트ID": "P2", "시설ID": "F2", "시설명": "s2", "사업유형": "신축", "공종": "건축", "공고번호": "A", "공고차수": "000", "공고명": "n", "수요기관": "d", "추정가격": 100, "주공종명": "건축공사업", "부공종명": ""}])
    assert pipeline._trade_targets(latest, {}, {}) == ["A-000", "M-000"], "건축 + 통합발주(기계설비·토목), 복수 시설 배정은 한 번"
    tdocs = {"A-000": dict(t, _files=["공내역서_B1.xlsx"], 공종별금액=t["공종별금액"] + [{"공종": "전기", "항목명_원문": "전기공사", "금액_원": 500}])}
    det = pipeline.assemble_trade_detail_table(latest, tdocs)
    assert list(det.columns) == pipeline.DETAIL_COLS and len(det) == 12 and set(det["프로젝트ID"]) == {"P1", "P2"}, "복수 시설 배정은 프로젝트마다"
    assert det.iloc[0]["출처파일"] == "공내역서_B1.xlsx" and det.iloc[0]["상위공종"] == "건축"
    dup = pipeline._trade_dup_logs(latest, tdocs)
    assert len(dup) == 1 and "전기" in dup[0]["비고"], dup       # P1 에만 전기 분리발주 공고가 있음(P2 는 건축뿐)
    assert pipeline.assemble_trade_detail_table(latest, {}).empty
    text = "공고문 본문 … 2. 공사금액 표 … 건축 토목 조경\n===== [공내역서.xlsx] =====\n## 세부내역\n세부항목0\t1\n===== [현장설명서.docx] =====\n집계 참고"
    body = pipeline._non_excel_text(text)
    assert "세부항목0" not in body and "집계 참고" in body and "공사금액 표" in body
    assert "공사금액 표" in keyword_windows(body, _TRADE_KEYS, 500) and keyword_windows("", _TRADE_KEYS, 500) == ""
    cfg = {"paths": {"files_dir": os.path.join(tmp, "files")}}
    snip2, used = pipeline._trade_input(cfg, "B1", text, 20000)
    assert used == ["공내역서_B1.xlsx"] and "총괄집계표" in snip2 and "세부항목0" not in snip2 and "[공고문·현장설명서 발췌]" in snip2
    assert pipeline._trade_input(cfg, "B9", "공고문에 공사금액 합계만 있고 다른 단서 없음", 20000) == ("", []), "내역서도 단서도 없으면 제외"
    assert pipeline._trade_input(cfg, "B9", "공종별 금액: 건축 1,000,000,000 토목 200,000,000", 20000)[0].startswith("## [공고문·현장설명서 발췌]")


def test_trade_review_fixes(tmp):
    """반박 검증에서 나온 결함의 회귀: 계층형 집계표 발췌, 세부내역만 있는 파일, .xls 다중 시트, 라벨 매핑, 본문 단서, 배치 부분 실패 보존, 스트리밍 전환."""
    from types import SimpleNamespace as NS
    from unittest import mock
    from openpyxl import Workbook
    from g2b_costdb import attachments, pipeline
    from g2b_costdb.extract_llm import (_trade_label, normalize_trade_doc, cross_verify_trades, submit_batches, fetch_batch, _messages_call,
                                        build_trade_request_params, parse_trade_response_text, TRADE_SYSTEM, trade_output_schema)
    d = os.path.join(tmp, "files2", "C1")
    os.makedirs(d)
    # ① 계층형 집계표: 머리글·대공종 헤더(금액 없음)·들여쓴 세부·소계가 탭 위치를 유지한 채 남는다
    wb = Workbook()
    ws = wb.active
    ws.title = "공종별집계표"
    for row in ([["(단위: 원)"], ["공종", "합계", "재료비", "노무비", "경비"], ["건축공사", None, None, None, None],
                 [None, "가설공사", 300000000, 1, 1], [None, "철근콘크리트공사", 5000000000, 1, 1], [None, "소계", 5500000000, 1, 1],
                 ["기계설비공사", None, None, None, None], [None, "위생기구설비", 400000000, 1, 1], [None, "소계", 400000000, 1, 1], ["합계", 5900000000, 1, 1, 1]]):
        ws.append(row)
    p = os.path.join(d, "설계내역서.xlsx")
    wb.save(p)
    s = attachments.cost_sheet_snippets(p)
    lines = s.splitlines()
    assert lines[1] == "(단위: 원)" and lines[2] == "공종\t합계\t재료비\t노무비\t경비" and "건축공사" in lines and "기계설비공사" in lines, lines
    assert "\t소계\t5500000000\t1\t1" in lines and "\t가설공사\t300000000\t1\t1" in lines, "들여쓰기(빈 첫 셀) 보존"
    # ② 집계표류 시트가 없고 세부내역만 있는 파일 → 빈 결과(건너뜀), 산출서·관급 파일은 후보에서 제외, 집계표 시트 있는 파일이 먼저
    wb2 = Workbook()
    ws2 = wb2.active
    ws2.title = "공종별내역서"
    for i in range(300):
        ws2.append(["건축공사", "철근콘크리트공사", f"콘크리트 {i}", "m3", 10, 120000, 1200000])
    wb2.save(os.path.join(d, "내역서.xlsx"))
    wb2.save(os.path.join(d, "수량산출서.xlsx"))
    wb2.save(os.path.join(d, "관급자재내역서.xlsx"))
    assert attachments.cost_sheet_snippets(os.path.join(d, "내역서.xlsx")) == ""
    assert [os.path.basename(f) for f in attachments.find_cost_sheet_files(os.path.join(tmp, "files2"), "C1")] == ["설계내역서.xlsx", "내역서.xlsx"]
    # ③ .xls 다중 시트(가짜 xlrd book): 집계표 시트의 행이 읽혀야 한다(지연 바인딩 결함)
    import xlrd

    class _Sh:
        def __init__(self, name, rows):
            self.name, self._r, self.nrows = name, rows, len(rows)

        def row_values(self, i):
            return list(self._r[i])
    book = NS(sheets=lambda: [_Sh("총괄집계표", [("공종", "금액"), ("건축공사", 20000000.0), ("기계설비공사", 3000000.0), ("총공사비", 27500000.0)]),
                              _Sh("세부내역", [("품명", "금액")] + [(f"품{i}", 100.0) for i in range(5)])])
    with mock.patch.object(xlrd, "open_workbook", return_value=book):
        sx = attachments.cost_sheet_snippets(os.path.join(d, "내역서.xls"))
        tx = attachments.extract_xls(os.path.join(d, "내역서.xls"))
    assert sx.startswith("## [내역서.xls] 시트: 총괄집계표") and "건축공사\t20000000" in sx and "총공사비\t27500000" in sx and "품0" not in sx, sx
    assert "## 총괄집계표" in tx and "품4\t100" in tx
    # ④ 라벨 매핑(비구조화 폴백 응답) + '기타' 과다 경고
    assert [_trade_label(x) for x in ("건축공사비", "기계설비(공조)", "전기(가설)", "산재보험료", "토목부문", "승강기설치", "조경")] == \
        ["건축", "기계설비", "공통가설", "간접비", "토목", "기타", "조경"]
    t = normalize_trade_doc({"합계_원": "100", "공종별금액": [{"공종": "승강기", "항목명_원문": "a", "금액_원": "90"}, {"공종": "건축", "항목명_원문": "b", "금액_원": "10"}]})
    assert any(l["항목"] == "공종 라벨 미매핑(문서)" for l in cross_verify_trades("C1", {}, t)), "'기타' 90% → 라벨 미매핑 경고"
    # ⑤ 본문 단서: 상투어는 대상 아님, 공종별 금액 표 표현만
    assert not pipeline._TRADE_TEXT_HINT.search("예정가격은 원가계산에 의한 가격으로 결정하며 공종별 분리발주, 총괄 책임자, 집계")
    assert pipeline._TRADE_TEXT_HINT.search("공종별 금액: 건축 1,234,567,000원") and pipeline._TRADE_TEXT_HINT.search("총괄집계표") \
        and pipeline._TRADE_TEXT_HINT.search("기계설비공사 2,100,000,000원")
    # ⑥ 2차 배치: params_builder 로 2차 system·schema, parser 로 2차 해석, 뒤 청크 실패 시 앞 배치 보존(on_submitted + e.submitted)
    calls = {"n": 0, "kept": []}

    class FakeBatches:
        def create(self, requests):
            calls["n"] += 1
            calls["reqs"] = requests
            if calls["n"] == 3:                       # 두 번째 제출의 2번째 청크에서 실패
                raise RuntimeError("overloaded_error 529")
            return NS(id=f"msgbatch_{calls['n']}", processing_status="in_progress")

        def retrieve(self, bid):
            return NS(processing_status="ended", request_counts=NS(processing=0, succeeded=1, errored=0, canceled=0, expired=0))

        def results(self, bid):
            txt = '{"금액기준": "부가세포함", "관급자재_포함여부": "제외", "단위": "원", "합계_원": "27,500,000,000", "공종별금액": [{"공종": "건축", "항목명_원문": "건축공사", "금액_원": "20000000000"}, {"공종": "기계설비공사", "항목명_원문": "기계", "금액_원": "3,000,000,000"}], "출처": "s", "신뢰도": "high", "근거문구": "r"}'
            return [NS(custom_id="K1", result=NS(type="succeeded", message=NS(content=[NS(type="text", text=txt)], usage=NS(input_tokens=5, output_tokens=6))))]
    client = NS(messages=NS(batches=FakeBatches()))
    cfg_llm = {"model": "claude-sonnet-5", "max_tokens": 8000, "structured_output": True}
    sub = submit_batches(client, [("K1", "발췌", {"공고번호": "K1"})], cfg_llm, params_builder=build_trade_request_params, on_submitted=calls["kept"].append)
    prm = calls["reqs"][0]["params"]
    assert prm["system"] == TRADE_SYSTEM and prm["output_config"]["format"]["schema"] == trade_output_schema() and sub[0]["id"] == "msgbatch_1" and calls["kept"] == sub
    st, got, counts = fetch_batch(client, "msgbatch_1", "claude-sonnet-5", parser=parse_trade_response_text)
    g = got["K1"]
    assert g["합계_원"] == 27_500_000_000 and g["공종별금액"][1] == {"공종": "기계설비", "항목명_원문": "기계", "금액_원": 3_000_000_000} and g["_공종합_원"] == 23_000_000_000 and g["_usage_in"] == 5
    calls["kept"].clear()
    try:
        submit_batches(client, [(f"K{i}", "x", {}) for i in range(1100)], cfg_llm, chunk=500, on_submitted=calls["kept"].append)
        raise AssertionError("예외가 나야 함")
    except RuntimeError as e:
        assert len(getattr(e, "submitted", [])) == 1 and calls["kept"] == e.submitted and e.submitted[0]["keys"][0] == "K0", "앞 청크 배치 보존"
    # ⑦ max_tokens 가 비스트리밍 상한을 넘으면 stream 으로 받는다
    class _Stream:
        def __init__(self, kw):
            self.kw = kw

        def __enter__(self):
            return self

        def __exit__(self, *a):
            return False

        def get_final_message(self):
            return NS(stop_reason="end_turn", content=[], kw=self.kw)
    used = {}
    cl2 = NS(messages=NS(create=lambda **kw: used.setdefault("create", kw) or NS(stop_reason="end_turn", content=[]),
                         stream=lambda **kw: _Stream(kw)))
    r1 = _messages_call(cl2, model="m", max_tokens=16000, messages=[])
    r2 = _messages_call(cl2, model="m", max_tokens=32000, messages=[])
    assert used["create"]["max_tokens"] == 16000 and r2.kw["max_tokens"] == 32000, "32,000 은 스트리밍"


def test_trade_allocation_formulas(tmp):
    """05 문서배분_* 는 다른 공종 공고 안에 묶인 몫만(자기 공종 공고 제외), 중복경고는 건축 집계표 기준 — formulas 로 실제 계산."""
    try:
        import formulas  # type: ignore
    except ImportError:
        print("(formulas 없음 → 05 배분 수식 값 검증 생략)")
        return
    import logging
    import warnings
    from g2b_costdb.build_excel import build_workbook
    fac = pd.DataFrame([{"프로젝트ID": "P1", "시설ID": "F1", "시설명": "s", "사업유형": "신축", "표2_대분류": "a", "표2_중분류": "b", "수요기관": "d",
                         "공고건수_최신": 3, "최종공고일": "2025-01-01", "연면적_m2": 1000}])
    base = {"프로젝트ID": "P1", "시설ID": "F1", "시설명": "s", "사업유형": "신축", "공고명": "n", "공고일시": "2025-01-01", "수요기관": "d", "공고차수": "000"}
    latest = pd.DataFrame([dict(base, 공고번호="A", 공고키="A-000", 공종="건축", 기초금액=100), dict(base, 공고번호="M", 공고키="M-000", 공종="기계설비", 기초금액=50),
                           dict(base, 공고번호="E", 공고키="E-000", 공종="전기", 기초금액=110)])
    trade = pd.DataFrame([dict(base, 공고번호="A", 공고키="A-000", 공종="건축", 기초금액_API=100), dict(base, 공고번호="M", 공고키="M-000", 공종="기계설비", 기초금액_API=50),
                          dict(base, 공고번호="E", 공고키="E-000", 공종="전기", 기초금액_API=110)])
    det_rows = [("A", "A-000", "건축", lab, amt) for lab, amt in (("건축", 60), ("기계설비", 30), ("토목", 10), ("기타", 5), ("철거", 5), ("간접비", 20), ("부가세", 12), ("공통가설", 3))]
    det_rows += [("M", "M-000", "기계설비", "기계설비", 40), ("M", "M-000", "기계설비", "토목", 10), ("E", "E-000", "전기", "전기", 80), ("E", "E-000", "전기", "정보통신", 20)]
    detail = pd.DataFrame([{"프로젝트ID": "P1", "시설ID": "F1", "시설명": "s", "사업유형": "신축", "상위공종": top, "공고번호": b, "공고차수": "000", "공고키": k,
                            "세부공종": lab, "항목명_원문": lab, "금액_원_문서": amt, "금액기준": "부가세포함", "관급자재_포함": "제외", "단위": "원",
                            "집계표합계_원_문서": None, "출처파일": "x.xlsx", "신뢰도": "high", "근거문구": ""} for b, k, top, lab, amt in det_rows])
    out = os.path.join(tmp, "alloc.xlsx")
    build_workbook(out, fac, latest, latest, trade, pd.DataFrame(), pd.DataFrame(), meta={"period": "t"}, trade_detail=detail)
    import openpyxl
    wb = openpyxl.load_workbook(out)
    ws5 = wb["05_공사비DB_시설합산"]
    h5 = [c.value for c in ws5[1]]
    col = lambda name: openpyxl.utils.get_column_letter(h5.index(name) + 1)
    warnings.filterwarnings("ignore")
    logging.disable(logging.CRITICAL)
    sol = formulas.ExcelModel().loads(out).finish().calculate()
    logging.disable(logging.NOTSET)

    def val(sheet, coord):
        for k, v in sol.items():
            if k.upper().endswith(f"{sheet.upper()}'!{coord}"):
                x = getattr(v, "value", v)
                x = x.tolist() if hasattr(x, "tolist") else x
                while isinstance(x, list) and x:
                    x = x[0]
                return x
    g = lambda name: val("05_공사비DB_시설합산", f"{col(name)}2")
    assert abs(float(g("기계설비(수식)")) - 50) < 1e-6 and abs(float(g("문서배분_기계설비(수식)")) - 100 * 30 / 110) < 1e-6, "기계설비 공고 자신의 몫(40) 은 제외"
    assert abs(float(g("문서배분_토목(수식)")) - (100 * 10 / 110 + 50 * 10 / 50)) < 1e-6, "건축 공고의 토목 + 기계설비 공고의 토목"
    assert float(g("문서배분_전기(수식)")) == 0 and g("문서배분_중복경고(수식)") in ("", None), "전기 공고 자신의 집계표는 중복이 아님"
    assert abs(float(g("건축_순건축추정(수식)")) - 100 * 60 / 110) < 1e-6
    assert abs(float(g("총공사비합계_기초금액기준(수식)")) - 260) < 1e-6, "총공사비합계는 배분과 무관"
    assert abs(float(g("총공사비합계_추정가격기준(수식)")) - 260) < 1e-6, "추정가격이 없으면 추정가격 기준도 기초금액으로 대신"
    assert abs(float(g("건축공종계(수식)")) - 150) < 1e-6, "건축공종계 = 건축 100 + 기계설비 50 (+토목·조경 0)"


def test_amount_structure(tmp):
    """요청 양식(2026-11-14): 04 금액 구성(추정가격|부가세|도급자관급비|계ⓐ|관급자관급비ⓑ|기타금액ⓒ|합계 2종), 부가세 출처·과세구분, 기타금액 가산 판정,
    공고명 꼬리표 공종 분류, 기타 공종 비고, 01 공사비 정보광장 틀."""
    from g2b_costdb import classify, pipeline
    from g2b_costdb.build_excel import build_workbook, FAC_OVERVIEW_COLS
    # 공고명 꼬리표: 주공종명이 비면 '(전기)'·'(통신)'·'(소방)'·'(기계)' 가 공종. 주공종명이 있으면 그쪽이 우선, 서로 다른 꼬리표 둘이면 판단 안 함
    ct = lambda n, m="": classify.classify_trade(n, m, "")
    assert ct("KAIST미술관 신축공사(3단계)(전기)") == ("전기", "공고명(꼬리표)") and ct("○○호텔 리모델링 공사(통신)")[0] == "정보통신"
    assert ct("○○ 소공연장 조성공사(소방)")[0] == "소방" and ct("○○ 신축공사(기계)")[0] == "기계설비" and ct("○○ 건립공사[토목]")[0] == "토목"
    assert ct("○○ 주차장 건립공사(건축)", "건축공사업") == ("건축", "주공종명") and ct("○○ 공사(전기)", "일반소방시설공사업(전기)")[0] == "소방"
    assert ct("○○ 신축공사(건축)(전기)")[0] == "건축" and ct("소방서 신축공사")[0] == "건축" and ct("○○ 리모델링 기계공사")[0] == "기계설비"
    assert ct("감포 축구장 개보수 공사") == ("기타", "미분류")
    note = classify.trade_note("기타", "미분류", "감포 축구장 개보수 공사")
    assert "미분류 사유" in note and "리모델링·개보수" in note and classify.trade_note("건축", "공고명", "○○ 신축공사") == ""
    assert "석면해체.제거업" in classify.trade_note("기타", "미분류", "○○ 철거공사", "석면해체.제거업") and "철거·해체" in classify.trade_note("기타", "미분류", "○○ 철거공사")
    # 기타금액 가산 판정: 문서 '별도' → 가산, 관급 칸과 같은 값·문서가 관급/추정가격에 포함이라 적음 → 제외
    ed = pipeline._etc_decision
    assert ed(5600000, "별도", {}, {}).startswith("가산") and ed(5600000, "", {}, {}).startswith("가산(위치 불명")
    assert ed(24379000, "별도", {"관급자관급액_API": 24379000}, {}).startswith("제외(관급 금액과 같은 값")
    assert ed(100, "관급에포함", {}, {}).startswith("제외") and ed(100, "추정가격에 포함", {}, {}).startswith("제외") and ed(None, "별도", {}, {}) == "" and ed(0, "별도", {}, {}) == ""
    # 검증로그 '기초금액≈추정가격×1.1': 면세·부분면세(공고 부가세와 일치)·조달청 공고·공동주택은 경고가 아니라 참고
    from g2b_costdb.extract_llm import cross_verify
    vat_log = lambda api, doc: next(l for l in cross_verify("X", api, doc) if l["항목"] == "기초금액≈추정가격×1.1")
    assert vat_log({"추정가격": 1000, "기초금액": 1100}, {})["판정"] == "정상"
    lg = vat_log({"추정가격": 1000, "기초금액": 1027, "부가세_API": 27}, {})
    assert lg["판정"] == "참고" and "면세" in lg["비고"], lg
    assert vat_log({"추정가격": 1000, "기초금액": 1000}, {"부가가치세_원": 0})["판정"] == "참고", "문서 부가세 0 = 면세"
    lg = vat_log({"추정가격": 1000, "기초금액": 1075, "부가세_API": 100, "공고기관": "조달청 충북지방조달청"}, {})
    assert lg["판정"] == "참고" and "조달청" in lg["비고"], lg
    assert vat_log({"추정가격": 1000, "기초금액": 1030, "표2_대분류": "공동주택"}, {})["판정"] == "참고"
    assert vat_log({"추정가격": 1000, "기초금액": 1030, "표2_대분류": "운동시설"}, {})["판정"] == "경고", "근거 없는 불일치는 경고 유지"
    assert vat_log({"추정가격": 1000, "기초금액": 1300, "표2_대분류": "공동주택"}, {})["판정"] == "경고"
    try:
        import formulas  # type: ignore
    except ImportError:
        print("(formulas 없음 → 04 금액 구성 수식 값 검증 생략)")
        return
    import logging
    import warnings
    import openpyxl
    base = {"프로젝트ID": "P1", "시설ID": "F1", "시설명": "s", "사업유형": "신축", "공고명": "n", "공고일시": "2025-03-02 10:00", "수요기관": "d", "공고차수": "000"}
    rows = [  # A 과세(API 부가세), 관급 도급자 20·관급자 30
        dict(공고번호="A", 공종="건축", 추정가격_API=1000, 기초금액_API=1100, 부가세_API=100, 도급자관급액_API=20, 관급자관급액_API=30),
        # B 면세(기초금액=추정가격, 부가세 원천 없음)
        dict(공고번호="B", 공종="전기", 추정가격_API=500, 기초금액_API=500),
        # C 부분면세(문서 부가세 13), 기타금액 7 가산
        dict(공고번호="C", 공종="정보통신", 추정가격_API=400, 기초금액_API=413, 부가세_문서=13, 기타금액_문서=7, 기타금액_판정="가산(문서: 별도 금액)"),
        # D 조달청형: 기초금액(조사금액)이 추정가격×1.1 보다 낮음, API 부가세는 10% → 두 기준이 다름. 기타금액은 관급과 같은 값이라 제외
        dict(공고번호="D", 공종="소방", 추정가격_API=2000, 기초금액_API=2150, 부가세_API=200, 관급자재_API=50, 기타금액_문서=50, 기타금액_판정="제외(관급 금액과 같은 값 — 이중계산 방지)"),
        # E 기초금액·부가세 없음 → 10% 가정 / F 추정가격 없음(문서 기초금액만) / G 부가세 원천 없고 기초금액이 7.5% 높음 → 확인필요
        dict(공고번호="E", 공종="조경", 추정가격_API=300),
        dict(공고번호="F", 공종="토목", 기초금액_문서=660),
        dict(공고번호="G", 공종="기계설비", 추정가격_API=1000, 기초금액_API=1075),
        dict(공고번호="H", 공종="기타", 공종_비고="미분류 사유: x", 추정가격_문서=200, 기초금액_문서=220, 도급자관급액_문서=5, 관급자관급액_문서=6)]
    trade = pd.DataFrame([dict(base, 공고키=r["공고번호"] + "-000", **r) for r in rows])
    fac = pd.DataFrame([{"프로젝트ID": "P1", "시설ID": "F1", "시설명": "s", "사업유형": "신축", "표2_대분류": "문화 및 집회시설", "표2_중분류": "공연장", "수요기관": "d",
                         "공고건수_최신": 8, "최종공고일": "2025-03-02", "연면적_m2": 100, "구조": "철근콘크리트조", "지하층수": 1, "지상층수": 4,
                         "대지면적_m2": 400, "건축면적_m2": 50}])
    latest = trade.rename(columns={"추정가격_API": "추정가격", "기초금액_API": "기초금액"})
    out = os.path.join(tmp, "amount.xlsx")
    build_workbook(out, fac, latest, latest, trade, pd.DataFrame(), pd.DataFrame(), meta={"period": "t"})
    wb = openpyxl.load_workbook(out)
    ws4, ws5, ws1 = wb["04_공사비DB_공종별"], wb["05_공사비DB_시설합산"], wb["01_시설마스터"]
    h4, h5, h1 = [c.value for c in ws4[1]], [c.value for c in ws5[1]], [c.value for c in ws1[1]]
    order = ["추정가격(수식)", "부가세(수식)", "도급자관급비(수식)", "추정금액_계ⓐ(수식)", "관급자관급비ⓑ(수식)", "기타금액ⓒ(수식)", "금액합계_추정가격기준(수식)", "금액합계_기초금액기준(수식)"]
    assert [h for h in h4 if h in order] == order and h4.index("금액합계_기초금액기준(수식)") - h4.index("추정가격(수식)") == 7, "요청 양식 순서로 연속 배치"
    assert h4.index("공종_비고") == h4.index("공종") + 1
    assert [h for h in h1 if h in FAC_OVERVIEW_COLS] == FAC_OVERVIEW_COLS, "01 공사개요는 공사비 정보광장 항목 순서"
    assert h1.index("대지면적_m2") == h1.index("연면적_m2") + 1 and "설계특이사항" in h1 and "주차대수" in h1
    warnings.filterwarnings("ignore")
    logging.disable(logging.CRITICAL)
    sol = formulas.ExcelModel().loads(out).finish().calculate()
    logging.disable(logging.NOTSET)

    def val(sheet, coord):
        for k, v in sol.items():
            if k.upper().endswith(f"{sheet.upper()}'!{coord}"):
                x = getattr(v, "value", v)
                x = x.tolist() if hasattr(x, "tolist") else x
                while isinstance(x, list) and x:
                    x = x[0]
                return x
    L4 = lambda name: openpyxl.utils.get_column_letter(h4.index(name) + 1)
    row = {ws4.cell(r, h4.index("공고번호") + 1).value: r for r in range(2, ws4.max_row + 1)}
    v4 = lambda bid, name: val("04_공사비DB_공종별", f"{L4(name)}{row[bid]}")
    num = lambda x: float(x) if x not in ("", None) else None
    # A 과세
    assert v4("A", "부가세_출처(수식)") == "API" and num(v4("A", "부가세(수식)")) == 100 and v4("A", "과세구분(수식)") == "과세"
    assert num(v4("A", "추정금액_계ⓐ(수식)")) == 1120 and num(v4("A", "관급자관급비ⓑ(수식)")) == 30
    assert num(v4("A", "금액합계_추정가격기준(수식)")) == 1150 and num(v4("A", "금액합계_기초금액기준(수식)")) == 1150
    # B 면세: 기초금액 = 추정가격 → 부가세 0
    assert v4("B", "부가세_출처(수식)") == "기초금액-추정가격" and num(v4("B", "부가세(수식)")) == 0 and v4("B", "과세구분(수식)") == "면세"
    assert num(v4("B", "금액합계_추정가격기준(수식)")) == 500 == num(v4("B", "금액합계_기초금액기준(수식)"))
    # C 부분면세 + 기타금액 가산
    assert v4("C", "부가세_출처(수식)") == "문서" and v4("C", "과세구분(수식)") == "부분면세" and num(v4("C", "기타금액ⓒ(수식)")) == 7
    assert num(v4("C", "금액합계_추정가격기준(수식)")) == 420 and num(v4("C", "금액합계_기초금액기준(수식)")) == 420
    # D 조달청형: 추정가격 기준 2000+200+50, 기초금액 기준 2150+50. API 합계만 있는 관급은 관급자 쪽으로, 기타금액은 제외
    assert num(v4("D", "도급자관급비(수식)")) == 0 and num(v4("D", "관급자관급비ⓑ(수식)")) == 50 and num(v4("D", "기타금액ⓒ(수식)")) == 0
    assert num(v4("D", "금액합계_추정가격기준(수식)")) == 2250 and num(v4("D", "금액합계_기초금액기준(수식)")) == 2200 and v4("D", "과세구분(수식)") == "과세"
    # E 10% 가정 / F 추정가격 없음 → 두 기준 모두 문서 기초금액 / G 확인필요
    assert v4("E", "부가세_출처(수식)") == "10% 가정" and v4("E", "과세구분(수식)") == "미확인(10% 가정)" and num(v4("E", "금액합계_기초금액기준(수식)")) == 330
    assert v4("F", "추정가격(수식)") in ("", None) and num(v4("F", "금액합계_추정가격기준(수식)")) == 660 == num(v4("F", "금액합계_기초금액기준(수식)"))
    assert v4("G", "부가세_출처(수식)") == "기초금액-추정가격" and str(v4("G", "과세구분(수식)")).startswith("확인필요") and abs(num(v4("G", "부가세율(수식)")) - 0.075) < 1e-9
    # H 문서값만: 추정가격·기초금액·관급 모두 문서에서
    assert num(v4("H", "추정가격(수식)")) == 200 and v4("H", "부가세_출처(수식)") == "기초금액(문서)-추정가격" and num(v4("H", "추정금액_계ⓐ(수식)")) == 225
    assert num(v4("H", "금액합계_추정가격기준(수식)")) == 231 == num(v4("H", "금액합계_기초금액기준(수식)"))
    # 05 대표값 2종과 ㎡당, 건축공종계
    L5 = lambda name: openpyxl.utils.get_column_letter(h5.index(name) + 1)
    v5 = lambda name: num(val("05_공사비DB_시설합산", f"{L5(name)}2"))
    tot_b = 1150 + 500 + 420 + 2200 + 330 + 660 + 1075 + 231
    tot_p = 1150 + 500 + 420 + 2250 + 330 + 660 + 1075 + 231       # G 는 부가세 원천이 '기초금액-추정가격'(75)이라 두 기준이 같다
    assert v5("총공사비합계_기초금액기준(수식)") == tot_b and v5("총공사비합계_추정가격기준(수식)") == tot_p, (v5("총공사비합계_기초금액기준(수식)"), v5("총공사비합계_추정가격기준(수식)"))
    assert abs(v5("㎡당공사비_기초금액기준_원(수식)") - tot_b / 100) < 1e-6 and abs(v5("㎡당공사비_추정가격기준_원(수식)") - tot_p / 100) < 1e-6
    assert v5("건축공종계(수식)") == 1150 + 330 + 660 + 1075 and v5("기타(수식)") == 231
    # 01 수식: 공사규모·건폐율 계산·건축법상용도
    L1 = lambda name: openpyxl.utils.get_column_letter(h1.index(name) + 1)
    assert val("01_시설마스터", f"{L1('공사규모(수식)')}2") == "철근콘크리트조 지하 1층~지상 4층"
    assert abs(float(val("01_시설마스터", f"{L1('건폐율_계산(수식)')}2")) - 0.125) < 1e-9
    assert val("01_시설마스터", f"{L1('건축법상용도(수식)')}2") == "문화 및 집회시설(공연장)"
    # enrich_facility: 건축 문서 우선, 없으면 다른 공종 + 보강 추출, 출처·발주년월·공사명
    lt = pd.DataFrame([dict(프로젝트ID="P1", 공고키="A-000", 공고번호="A", 공종="건축", 공고명="○○ 신축공사", 추정가격=1000, 공고일시="2025-03-02 10:00"),
                       dict(프로젝트ID="P1", 공고키="B-000", 공고번호="B", 공종="전기", 공고명="○○ 신축공사(전기)", 추정가격=500, 공고일시="2025-01-15 10:00"),
                       dict(프로젝트ID="P2", 공고키="C-000", 공고번호="C", 공종="전기", 공고명="△△ 전기공사", 추정가격=300, 공고일시="2024-11-20 10:00")])
    fc = pd.DataFrame([{"프로젝트ID": "P1", "공사현장지역": "충청북도 증평군"}, {"프로젝트ID": "P2", "공사현장지역": "서울특별시"}])
    docs = {"A-000": {"연면적_m2": 7281.94, "구조": "철근콘크리트조", "공사위치": "충북 증평군 증평읍 1"}, "B-000": {"연면적_m2": 9999, "대지면적_m2": 14717, "지상층수": 4},
            "C-000": {"_error": "x"}}
    extra = {"A-000": {"건폐율_pct": 12.35, "주차대수": 125, "건축법상용도": "업무시설"}, "B-000": {"주차대수": 1, "용적률_pct": 37.13}, "C-000": {"대지면적_m2": 55, "설계특이사항": "3개동"}}
    ef = pipeline.enrich_facility(fc, lt, docs, extra).set_index("프로젝트ID")
    p1 = ef.loc["P1"]
    assert p1["연면적_m2"] == 7281.94 and p1["대지면적_m2"] == 14717 and p1["지상층수"] == 4, "건축 우선, 빈 항목만 다른 공종에서"
    assert p1["건폐율_pct"] == 12.35 and p1["주차대수"] == 125 and p1["용적률_pct"] == 37.13 and p1["건축법상용도_문서"] == "업무시설"
    assert p1["공사명"] == "○○ 신축공사" and p1["현장위치"] == "충북 증평군 증평읍 1" and p1["발주년월"] == "2025-03" and p1["발주년월_기준"] == "건축 공고"
    assert p1["개요_출처"] == "건축·전기"
    p2 = ef.loc["P2"]
    assert p2["대지면적_m2"] == 55 and p2["설계특이사항"] == "3개동" and p2["현장위치"] == "서울특별시" and p2["발주년월"] == "2024-11"
    assert p2["발주년월_기준"] == "전기 공고(건축 공고 없음)" and p2["개요_출처"] == "전기 (건축 공고 없음)"


def test_request_review_fixes(tmp):
    """요청 양식 반영분 검토(2026-11-14) 지적 수정: 공종 분류 보강, 대표 교체 경고, 04 금액 수식의 0·오독 방어, 보강 추출 정규화·발췌·판정,
    대기 배치 이중 호출 방지, 인증키 가림, 용역 캐시·보존 규칙."""
    from g2b_costdb import classify, pipeline, dedup, collect
    from g2b_costdb.build_excel import build_workbook
    ct = lambda n, m="": classify.classify_trade(n, m, "")
    # '건축 기계 공사'·'토목,건축,기계공사' 는 건축 통합 계약, '소방기계'·'무대기계' 는 '기계공사' 규칙에 걸리지 않음
    assert ct("광주 이스포츠 상설경기장 리모델링 건축 기계 공사")[0] == "건축" and ct("스마트팜 온실신축사업 토목,건축,기계공사(긴급)")[0] == "건축"
    assert ct("동춘인라인롤러경기장 개보수 소방기계공사")[0] != "기계설비" and ct("○○ 무대기계공사")[0] != "기계설비" and ct("○○ 리모델링 기계공사")[0] == "기계설비"
    # 괄호 앞이 '소방'·'건축' 이면 괄호는 분야 표기 → 앞 공종
    assert ct("노후공공임대주택 리모델링 소방(전기)공사") == ("소방", "공고명(꼬리표)") and ct("공연장 외 조성 건축(기계)공사")[0] == "건축"
    # 꼬리표 변형: (통신설비)·(전기/2차분)·(증축공사, 전기)·【소방】. '(소방전기)' 는 다루지 않음
    assert ct("주차타워 건립사업(통신설비)")[0] == "정보통신" and ct("첨단온실 신축공사(전기/2차분)")[0] == "전기"
    assert ct("미술관 조성사업(증축공사, 전기)")[0] == "전기" and ct("○○ 신축공사【소방】")[0] == "소방" and ct("주차타워 조성공사(소방전기)")[0] == "건축"
    # 기타 비고: '안전시설'·'대전시립' 의 '전시' 는 전시 공사가 아님
    assert "전시·인테리어" not in classify.trade_note("기타", "미분류", "설악야구장 안전시설 리모델링 공사")
    assert "전시·인테리어" in classify.trade_note("기타", "미분류", "○○미술관 전시실 인테리어 공사")
    # 같은 세부사업·같은 공종에서 이름이 다른 공고는 서로 대체하지 않고 둘 다 합산(06 '같은 공종 합산' 참고). 같은 이름으로 다시 난 공고만 재공고로 대체
    mk = lambda no, name, dt, price: dict(시설ID="F1", 시설명="가상시 회관", 공고번호=no, 공고차수="000", 공고키=no + "-000", 공고명=name, 공고일시=dt, 사업유형="신축",
                                          공종="전기", 공종근거="공고명(꼬리표)", 추정가격=price, 공고종류="일반", 이전공고번호="")
    hh, lt, lg = dedup.dedup_latest(pd.DataFrame([mk("A", "가상시 회관 건립 전기공사", "2024-01-10 10:00:00", 178_500_000),
                                                   mk("B", "가상시 회관 건립공사(전기)", "2024-06-10 10:00:00", 86_500_000)]))
    assert sorted(lt["공고번호"]) == ["A", "B"] and set(lt["프로젝트ID"]) == {"F1-N"} and set(lt["대표선정사유"]) == {"합산(같은 공종 2건)"}, lt[["공고번호", "프로젝트ID", "대표선정사유"]]
    assert any(l["항목"] == "같은 공종 합산" and l["판정"] == "참고" for l in lg), lg
    hh, lt, lg = dedup.dedup_latest(pd.DataFrame([mk("A", "가상시 회관 건립 전기공사", "2024-01-10 10:00:00", 178_500_000),
                                                   mk("B", "[긴급] 가상시 회관 건립 전기공사(재공고)", "2024-06-10 10:00:00", 178_500_000)]))
    assert list(lt["공고번호"]) == ["B"] and not any(l["항목"] == "같은 공종 합산" for l in lg), lg
    ha = hh.set_index("공고번호")
    assert ha.at["A", "대표선정사유"] == "대체됨(같은 공고명으로 다시 공고)" and ha.at["A", "대체공고번호"] == "B" and "재공고로 1건 대체" in ha.at["B", "대표선정사유"]

    # ── 보강 추출 정규화: 단위가 붙은 숫자, 원문 보존(_raw), 멱등, ×100 없음, 위치 표기 변형 ──
    nx = extract_llm.normalize_extra_doc
    x = nx({"건폐율_pct": "12.35%", "주차대수": "약 125대", "최고높이_m": "24.5m", "대지면적_m2": "1,817.93㎡", "기타금액_원": "5,600천원",
            "기타금액_항목": "한전불입금 5,600천원", "기타금액_위치": "별도", "용적률_pct": "60% 이하", "근거문구": "원문 한 줄"})
    assert (x["건폐율_pct"], x["주차대수"], x["최고높이_m"], x["대지면적_m2"], x["기타금액_원"], x["기타금액_위치"]) == (12.35, 125, 24.5, 1817.93, 5600000, "별도"), x
    assert x["용적률_pct"] is None and x["_raw"] == {"용적률_pct": "60% 이하"} and x["근거문구"] == {"원문": "원문 한 줄"} and nx(x) == x
    y = nx({"건폐율_pct": "0.9", "용적률_pct": "0.95", "대지면적_m2": "89000", "건축면적_m2": "801"})
    assert (y["건폐율_pct"], y["용적률_pct"]) == (0.9, 0.95), "1% 미만 건폐율·용적률은 실제 값일 수 있어 ×100 하지 않는다"
    assert nx({"건폐율_pct": "0.1235", "대지면적_m2": "10000", "건축면적_m2": "1235"})["건폐율_pct"] == 12.35, "면적과 맞을 때만 비율 → %"
    assert [nx({"기타금액_원": "5600000", "기타금액_위치": w})["기타금액_위치"]
            for w in ("관급 포함", "도급액에 포함", "도급·관급 외 별도", "모름", "도급액 별도", "관급자재와 별도", "관급 미포함(별도)", "추정가격 포함")] == \
        ["관급에포함", "추정가격에포함", "별도", "불명", "별도", "별도", "별도", "추정가격에포함"]
    z = nx({"기타금액_원": "5억6천만원", "기타금액_위치": "관급 포함"})
    assert z["기타금액_원"] is None and z["기타금액_위치"] == "" and z["_raw"] == {"기타금액_원": "5억6천만원", "기타금액_위치": "관급에포함"} and nx(z) == z
    cv = extract_llm.cross_verify_extra
    assert any(l["항목"] == "숫자 해석 실패(용적률_pct)" and l["판정"] == "경고" for l in cv("X", {}, {}, x))
    assert not any(l["판정"] == "경고" for l in cv("X", {}, {}, y)), cv("X", {}, {}, y)
    # 기타금액이 기초금액의 30% 초과: 예산금액이 없어도 경고, 합계에는 더하지 않음. 추정가격·기초금액·부가세·예산과 같은 값도 제외
    big = nx({"기타금액_원": "5000000000", "기타금액_항목": "총공사비", "기타금액_위치": "불명"})
    api = {"추정가격": 909090909, "기초금액": 1e9, "예산금액": None}
    assert [l["판정"] for l in cv("X", api, {}, big) if l["항목"] == "기타금액(문서)"] == ["경고"]
    ed = pipeline._etc_decision
    assert ed(5e9, "불명", api, {}).startswith("제외(확인 필요: 기초금액의 30% 초과") and ed(1e9, "별도", api, {}).startswith("제외(확인 필요: 기초금액과 같은 값")
    assert ed(90909091, "별도", dict(api, 부가세_API=90909091), {}).startswith("제외(확인 필요: 부가세") and ed(38e6, "별도", api, {}).startswith("가산")
    assert ed(38e6, "관급자재비에 포함", api, {}).startswith("제외(문서: 관급") and ed(38e6, "도급·관급 외 별도", api, {}).startswith("가산")
    # 발췌 2단: 흔한 낱말(용도·구조·관급…)이 앞에서 예산을 다 쓰지 않고, 뒤쪽의 건폐율·한전불입금 구간이 들어간다
    g = ["입찰보증금은 면제한다.", "부가가치세 포함 금액으로 투찰한다.", "기타 사항은 별도 통보한다.", "구조물 시공 시 유의한다.", "건물의 용도에 맞게 시공한다.", "관급자재는 발주처가 지급한다."]
    t = "".join("가나다라마바사아자차카타파하 " * 80 + g[i % 6] + "\n" for i in range(82))
    t = t[:70000] + "\n대지면적 30,000㎡, 건폐율 15.2%\n" + t[70000:85000] + "\n한전불입금 38,000,000원(별도)\n" + t[85000:]
    prompt = extract_llm.build_extra_prompt(t, {}, {"extra_max_input_chars": 40000})
    assert "건폐율 15.2" in prompt and "한전불입금 38,000,000" in prompt and len(prompt) < 43000
    assert "건폐율 15.2" not in extract_llm.focus_text(t, 40000, keys=extract_llm._EXTRA_KEYS), "한 단으로 섞으면 빠지던 구간(회귀 기준)"

    # ── 같은 프로젝트의 두 공고가 같은 기타금액을 읽으면 하나만 가산, 옛 형식 근거문구(문자열)도 허용, 0 면적은 보강 값으로 채움 ──
    b = dict(공고차수="000", 프로젝트ID="P1", 시설ID="F1", 시설명="s", 사업유형="신축", 수요기관="가상시", 공고일시="2024-03-10 10:00")
    lt = pd.DataFrame([dict(b, 공고키="A-000", 공고번호="A", 공고명="가상 문화회관 건립공사", 공종="건축", 추정가격=3e10, 기초금액=3.3e10),
                       dict(b, 공고키="E-000", 공고번호="E", 공고명="가상 문화회관 건립공사(전기)", 공종="전기", 추정가격=2e9, 기초금액=2.2e9)])
    xd = {k: nx({"기타금액_원": "38000000", "기타금액_항목": "한전불입금 38,000,000", "기타금액_위치": "별도", "대지면적_m2": "14717", "건축면적_m2": "2100"}) for k in ("A-000", "E-000")}
    docs = {"A-000": {"근거문구": "추정가격 30,000,000,000원", "대지면적_m2": 0, "건축면적_m2": 0, "지하층수": 0, "연면적_m2": 7000}, "E-000": {"근거문구": ["a", "b"]}}
    assert pipeline._etc_duplicates(lt, docs, xd) == {("P1", "A-000"): "E-000"}, "항목에 '한전' 이 있으면 전기 공고 쪽을 남긴다"
    tt = pipeline.assemble_trade_table(lt, docs, {}, xd).set_index("공고번호")
    assert tt.loc["E", "기타금액_판정"].startswith("가산") and tt.loc["A", "기타금액_판정"].startswith("제외(같은 프로젝트의 공고 E-000"), tt["기타금액_판정"].tolist()
    assert '"원문": "추정가격 30,000,000,000원"' in tt.loc["A", "근거문구"]
    ef = pipeline.enrich_facility(pd.DataFrame([{"프로젝트ID": "P1"}]), lt, docs, xd).iloc[0]
    assert ef["대지면적_m2"] == 14717 and ef["건축면적_m2"] == 2100 and ef["지하층수"] == 0 and ef["연면적_m2"] == 7000, ef.to_dict()
    cfg = {"paths": {"data_dir": tmp}, "verify": {"price_diff_warn_ratio": 0.005, "vat_ratio": 0.10}}
    assert any(l["항목"] == "기타금액 중복 가능" and l["공고번호"] == "A" for l in pipeline._extra_logs(cfg, lt, docs, xd))
    # 1차 검증로그를 excel 때 다시 만든다(_verify_logs): 저장된 옛 로그가 아니라 현재 대표 공고·현재 규칙 기준
    pd.DataFrame([{"공고번호": "OLD", "항목": "옛 로그", "판정": "경고"}]).to_parquet(os.path.join(tmp, "logs_verify.parquet"), index=False)
    lv = pipeline._verify_logs(cfg, lt, {"A-000": {"추정가격_원": 3e10}, "E-000": {"_error": "x"}})
    saved = pd.read_parquet(os.path.join(tmp, "logs_verify.parquet"))
    assert lv and set(saved["공고번호"]) == {"A"}, saved
    # 대기 중 배치에 든 공고는 즉시 호출 대상에서 제외(이중 과금 방지)
    json.dump({"batches": [{"id": "b1", "keys": ["A-000"], "status": "submitted"}, {"id": "b0", "keys": ["E-000"], "status": "ended"}]},
              open(os.path.join(tmp, "llm_extra_batch.json"), "w", encoding="utf-8"))
    assert pipeline._skip_pending(cfg, ["A-000", "E-000", "Z-000"], "llm_extra_batch.json", "extract-extra") == ["E-000", "Z-000"]
    assert pipeline._skip_pending(cfg, ["A-000"], "없는파일.json", "extract") == ["A-000"]

    # ── 인증키 가림·캐시 정리·용역 보존 규칙 ──
    c = _client(tmp)
    key = c.cfg.service_key
    msg = c._scrub(f"HTTPSConnectionPool(host='apis.data.go.kr'): Max retries exceeded with url: /x?serviceKey={key}&type=json (Caused by ...)")
    assert key not in msg and "serviceKey=***" in msg and "type=json" in msg
    for fetched, k in (("2024-03-10T09:00:00", "k1"), ("2024-03-11T09:00:00", "k2")):
        c.conn.execute("INSERT OR REPLACE INTO cache VALUES(?,?,?,?,?)", (k, "opS", '{"inqryBgnDt": "202403010000"}', "{}", fetched))
    c.conn.commit()
    assert c.purge_cache("opS", '"inqryBgnDt": "202403010000"', before="2024-03-11") == 1
    assert [r[0] for r in c.conn.execute("SELECT k FROM cache WHERE op='opS'")] == ["k2"], "오늘 받은 페이지는 남긴다(이어받기)"
    import re as _re
    rx = _re.compile(collect.SERVC_KEEP_NAME)
    assert all(rx.search(n) for n in ("OO 건축설계 제안공모", "OO 건립사업 설계 경기", "OO 미술관 건립 건축 설계", "OO 건립사업 기획설계", "OO 건립 설계안 공모", "OO 건립사업 건축설계(제안공모)", "OO 주민센터 신축 설계제안공모", "OO 건설폐기물 처리용역"))
    assert not rx.search("청사 청소 용역") and classify.load_config()["servc"]["keep_name_regex"] == collect.SERVC_KEEP_NAME, "config.yaml 과 코드 기본값이 같아야 한다"

    # ── 04 금액 수식: 문서의 0·단위 오독이 합계를 지우거나 '면세' 를 만들지 않는다 ──
    try:
        import formulas  # type: ignore
    except ImportError:
        print("(formulas 없음 → 04 금액 수식 방어 검증 생략)")
        return
    import logging
    import warnings
    import openpyxl
    base = {"프로젝트ID": "P1", "시설ID": "F1", "시설명": "s", "사업유형": "신축", "공고명": "n", "공고일시": "2025-03-02 10:00", "수요기관": "d", "공고차수": "000", "공종": "건축"}
    rows = [dict(공고번호="R2", 추정가격_API=5_000_000_000, 기초금액_문서=0),                              # 문서 기초금액 0
            dict(공고번호="R3", 추정가격_API=5_000_000_000, 기초금액_문서=3_300_000),                      # 문서 기초금액 단위 오독
            dict(공고번호="R4", 추정가격_API=5_000_000_000, 기초금액_API=5_500_000_000, 부가세_문서=0),     # 문서 부가세 0 인데 기초금액은 10%
            dict(공고번호="R5", 추정가격_API=5_000_000_000, 기초금액_API=5_500_000_000, 부가세_문서=500_000),  # 문서 부가세 천원 오독
            dict(공고번호="R6", 추정가격_문서=0, 기초금액_API=1_100_000_000),                              # 문서 추정가격 0
            dict(공고번호="R8", 추정가격_API=1_000_000_000, 기초금액_API=1_100_000_000, 부가세_API=0),      # API 부가세 0 인데 기초금액은 10%
            dict(공고번호="R9", 추정가격_API=1_000_000_000, 기초금액_API=1_000_000_000, 부가세_문서=0),     # 진짜 면세(기초금액 = 추정가격)
            dict(공고번호="RA", 추정가격_API=1000, 부가세_문서=30),                                        # 기초금액 없음 + 문서 부가세 3% → 문서 채택
            dict(공고번호="RB", 추정가격_API=1000, 부가세_문서=0),                                         # 기초금액 없음 + 문서 부가세 0 → 문서대로 면세
            dict(공고번호="RC", 추정가격_API=2000, 기초금액_API=2100, 부가세_문서=150),                    # 조달청형(기초금액이 ×1.1 이 아님) + 문서 부가세 → 문서 채택
            dict(공고번호="RD", 추정가격_API=10_000, 기초금액_API=11_000, 부가세_문서=300)]                # 기초금액은 10% 인데 문서 부가세 3%(다른 범위) → 기초금액 쪽
    trade = pd.DataFrame([dict(base, 공고키=r["공고번호"] + "-000", **r) for r in rows])
    fac = pd.DataFrame([{"프로젝트ID": "P1", "시설ID": "F1", "시설명": "s", "사업유형": "신축", "표2_대분류": "문화 및 집회시설", "표2_중분류": "공연장", "수요기관": "d",
                         "공고건수_최신": 11, "최종공고일": "2025-03-02", "연면적_m2": 100}])
    latest = trade.rename(columns={"추정가격_API": "추정가격", "기초금액_API": "기초금액"})
    out = os.path.join(tmp, "amount_guard.xlsx")
    build_workbook(out, fac, latest, latest, trade, pd.DataFrame(), pd.DataFrame(), meta={"period": "t"})
    wbk = openpyxl.load_workbook(out)
    ws4, ws5 = wbk["04_공사비DB_공종별"], wbk["05_공사비DB_시설합산"]
    h4, h5 = [c.value for c in ws4[1]], [c.value for c in ws5[1]]
    warnings.filterwarnings("ignore")
    logging.disable(logging.CRITICAL)
    sol = formulas.ExcelModel().loads(out).finish().calculate()
    logging.disable(logging.NOTSET)

    def val(coord, sheet="04_공사비DB_공종별"):
        for k, v in sol.items():
            if k.upper().endswith(f"{sheet}'!{coord}".upper()):
                x = getattr(v, "value", v)
                x = x.tolist() if hasattr(x, "tolist") else x
                while isinstance(x, list) and x:
                    x = x[0]
                return x
    row = {ws4.cell(r, h4.index("공고번호") + 1).value: r for r in range(2, ws4.max_row + 1)}
    v4 = lambda bid, name: val(f"{openpyxl.utils.get_column_letter(h4.index(name) + 1)}{row[bid]}")
    num = lambda x: float(x) if x not in ("", None) else None
    tp, tb, src, tax = "금액합계_추정가격기준(수식)", "금액합계_기초금액기준(수식)", "부가세_출처(수식)", "과세구분(수식)"
    for bid in ("R2", "R3"):
        assert num(v4(bid, tb)) == 5_500_000_000 == num(v4(bid, tp)) and v4(bid, src) == "10% 가정", (bid, v4(bid, tb), v4(bid, src))
    for bid in ("R4", "R5"):
        assert v4(bid, src) == "기초금액-추정가격" and v4(bid, tax) == "과세" and num(v4(bid, tp)) == 5_500_000_000 == num(v4(bid, tb)), (bid, v4(bid, src), v4(bid, tax))
    assert v4("R6", "추정가격(수식)") in ("", None) and num(v4("R6", tp)) == 1_100_000_000 == num(v4("R6", tb))
    assert v4("R8", src) == "기초금액-추정가격" and v4("R8", tax) == "과세" and num(v4("R8", tp)) == 1_100_000_000
    assert v4("R9", src) == "문서" and v4("R9", tax) == "면세" and num(v4("R9", tp)) == 1_000_000_000 == num(v4("R9", tb))
    assert v4("RA", src) == "문서" and v4("RA", tax) == "부분면세" and num(v4("RA", tp)) == 1030 == num(v4("RA", tb))
    assert v4("RB", src) == "문서" and v4("RB", tax) == "면세" and num(v4("RB", tp)) == 1000 == num(v4("RB", tb))
    assert v4("RC", src) == "문서" and v4("RC", tax) == "부분면세" and num(v4("RC", tp)) == 2150 and num(v4("RC", tb)) == 2100
    assert v4("RD", src) == "기초금액-추정가격" and v4("RD", tax) == "과세" and num(v4("RD", tp)) == 11_000 == num(v4("RD", tb))
    assert "세부사업명" in h5 and "세부사업수" in h5 and "세부사업명" in h4 and not any("대체된_다른이름" in str(c) for c in h5)


def test_work_split():
    """세부사업 분리(works.assign_works): 한 시설 후보에 묶인 여러 사업을 나누고, 같은 사업 안에서는 재공고만 대체하고 나머지는 합산."""
    from g2b_costdb import works, dedup

    def run(rows, facility="가상공사 공공주택", wtype="신축"):
        d = pd.DataFrame([dict(_base="F1-N", 시설명=facility, 사업유형=wtype, 공고번호=f"N{i:02d}", 공고명=nm, 공종=tr,
                               _dt=pd.Timestamp(dt), _price=amt * 1e8, 이전공고번호=prev)
                          for i, (nm, tr, dt, amt, *rest) in enumerate(rows) for prev in [rest[0] if rest else ""]])
        w = works.assign_works(d, "_base")
        return list(w["세부사업번호"]), list(w["역할"]), list(w["세부사업명"])

    # ① 단지가 다른 공고는 다른 세부사업, 같은 단지의 공종 공고는 한 세부사업(이름 표기가 조금 달라도: '위례' / '위례지구')
    wk, role, names = run([("고덕강일 공공주택지구 3단지 아파트 건설공사", "건축", "2023-03-14", 2924), ("고덕강일 공공주택지구 3단지 아파트 전기공사", "전기", "2023-08-18", 294),
                           ("마곡 16단지 공공주택 건설공사", "건축", "2024-11-08", 1514), ("마곡 16단지 공공주택 소방시설공사", "소방", "2025-02-05", 126),
                           ("위례 A1-14BL 공공주택 건설공사", "건축", "2025-04-29", 1349), ("위례지구 A1-14BL 공공주택 전기공사", "전기", "2025-06-25", 128),
                           ("고덕강일공공주택지구 2BL 제로에너지아파트 건설공사", "건축", "2019-06-14", 900), ("고덕강일공공주택지구 2BL 제로에너지아파트 전기공사", "전기", "2022-05-18", 118)])
    assert wk == [2, 2, 3, 3, 4, 4, 1, 1] and set(role) == {"count"}, (wk, role)          # 번호는 최초 공고일 순, 큰 공사는 3년 뒤의 공종 공고도 같은 사업
    assert names[0] == "고덕강일 공공주택지구 3단지 아파트 건설공사" and names[6].startswith("고덕강일공공주택지구 2BL")
    # ② 같은 세부사업·같은 공종의 분야별 발주(기계소방/전기소방)는 서로 대체하지 않고 둘 다 합산
    wk, role, _ = run([("독산2동 마을공원 지하공영주차장 건립공사", "건축", "2022-03-02", 150), ("독산2동 마을공원 지하공영주차장 기계소방공사", "소방", "2022-05-10", 9),
                       ("독산2동 마을공원 지하공영주차장 전기소방공사", "소방", "2022-05-10", 6)], facility="독산2동 마을공원 지하공영주차장")
    assert wk == [1, 1, 1] and role == ["count"] * 3
    # ③ 재공고: 같은 이름으로 며칠 뒤 다시 난 공고만 대체. 같은 날 난 같은 이름은 병행 발주, 이름이 조금 바뀌고 금액이 비슷하면 재공고, 이전공고번호 연결은 무조건
    wk, role, _ = run([("가상 회관 건립공사(건축)", "건축", "2024-01-10", 100), ("[긴급] 가상 회관 건립공사(건축) 재공고", "건축", "2024-02-20", 102),
                       ("가상 회관 건립 전기공사", "전기", "2024-01-11", 9), ("가상 회관 건립 전기공사", "전기", "2024-01-11", 9),
                       ("가상 회관 시범단지 온실 신축공사 입찰대행", "기타", "2024-05-14", 2.0), ("가상 회관 시범단지 조성 온실 신축공사 입찰대행", "기타", "2024-06-29", 1.94),
                       ("가상 회관 건립 소방공사", "소방", "2024-01-12", 5), ("가상 회관 건립 소방시설공사(2차)", "소방", "2024-03-12", 5.3, "N06")], facility="가상 회관")
    assert role == ["superseded", "count", "count", "count", "superseded", "count", "superseded", "count"], role
    # ④ 연례 사업: 해가 다르면 다른 세부사업(연도 표기가 없어도 같은 이름·같은 공종으로 1년 뒤 다시 나면 다음 해 분 — 재공고가 아니다)
    wk, role, _ = run([("공원녹지 권역별 관리공사(논현1권역)", "기타", "2019-01-18", 3.9), ("공원녹지 권역별 관리공사(논현1권역)", "기타", "2020-01-03", 4.3),
                       ("공원녹지 권역별 관리공사(논현2권역)", "기타", "2020-01-03", 4.0), ("2021년 공원녹지 권역별 관리공사(논현1권역)", "기타", "2021-01-14", 4.3)],
                      facility="공원녹지 권역별 관리", wtype="리모델링")
    assert wk == [1, 2, 3, 4] and role == ["count"] * 4, (wk, role)
    # ⑤ 한 해 사업의 건축은 권역별, 전기·통신은 전체 한 건 → 한 세부사업. 연도 표기 없는 전기 공고는 가까운 해의 사업에 붙고 다른 해와 섞이지 않는다
    wk, role, _ = run([("2025년 노후공공임대주택 리모델링공사(건축)-동부권", "건축", "2025-07-17", 12), ("2025년 노후공공임대주택 리모델링공사(건축)-서부권", "건축", "2025-07-17", 14),
                       ("2025년 노후공공임대주택 리모델링공사(전기)", "전기", "2025-07-17", 2.3),
                       ("노후공공임대주택 리모델링(매입임대 시설개선) 건축공사", "건축", "2021-08-26", 3.5),
                       ("2022년 노후공공임대주택 리모델링(매입임대 시설개선) 공사", "건축", "2022-07-08", 17.7), ("노후공공임대주택 리모델링(매입임대 시설개선) 전기공사", "전기", "2022-07-27", 1.5)],
                      facility="가상공사 노후공공임대주택", wtype="리모델링")
    assert wk[0] == wk[1] == wk[2] and wk[4] == wk[5] != wk[3] and len(set(wk)) == 3, wk
    # ⑥ 현장을 묶은 전기 공고 '(지산, 범물)' 는 현장별 건축 공고 '(지산)'·'(범물)' 과 한 세부사업, '(상인)' 은 별개
    wk, _, _ = run([("2026년 영구임대 리모델링공사(지산)", "건축", "2026-03-03", 7.3), ("2026년 영구임대 리모델링공사(범물)", "건축", "2026-03-03", 27.6),
                    ("2026년 영구임대 리모델링 전기공사(지산, 범물)", "전기", "2026-03-20", 1.8), ("2026년 영구임대 리모델링공사(상인)", "건축", "2026-03-03", 14.8)],
                   facility="가상공사 영구임대", wtype="리모델링")
    assert wk[0] == wk[1] == wk[2] != wk[3], wk
    # ⑦ 이름의 한 글자 오타는 같은 현장, 이름 끝·숫자가 다르면 다른 현장('신사'/'신길', '신사'/'신사2', '학동1'/'학동3')
    assert works._typo_same("교육관동급식전처리실", "교육관종급식전처리실") and works._typo_same("위례A114BL", "위례지구A114BL") and works._typo_same("제4봉안동", "제4봉안당")
    assert not works._typo_same("초행지붕신사", "초행지붕신길", rare=False) and not works._typo_same("상반기지산범물", "하반기지산범물") and not works._typo_same("논현1권역", "논현고잔1권역")
    wk, _, _ = run([("빈집활용사업(초행지붕-신사) 행복주택 건설공사", "건축", "2021-06-17", 8.9), ("빈집활용사업(초행지붕-신사) 행복주택 전기공사", "전기", "2021-06-29", 0.8),
                    ("빈집활용사업(초행지붕-신사2) 행복주택 건설공사", "건축", "2021-07-22", 8.0), ("빈집활용사업(초행지붕-신길) 행복주택 건설공사", "건축", "2021-08-06", 8.4),
                    ("학동1 공영주차장 주차전용건축물 신축공사(건축,기계)", "건축", "2020-05-18", 39), ("학동3 공영주차장 주차전용건축물 신축공사(건축ㆍ기계)", "건축", "2020-05-26", 20)],
                   facility="가상시 행복주택")
    assert wk[0] == wk[1] and len(set(wk)) == 5, wk
    # ⑧ 부대공사: 신축 본공사의 '정화조 설치공사'·본공사 이름 뒤에 세부 이름을 붙인 공고는 같은 세부사업. 자기 목적어를 가진 뒤늦은 공사('제례실 신축')와
    #    증축·리모델링 시설의 이름이 다른 소규모 공사는 별개 세부사업. 작은 공사가 2년 뒤 같은 이름으로 다시 나면 새 사업
    wk, role, _ = run([("독립문 문화공원 지하공영주차장 건립 공사(장기계속계약)", "건축", "2021-12-03", 157), ("독립문 문화공원 지하공영주차장 건립 공사(전기, 장기계속계약)", "전기", "2021-12-02", 7),
                       ("독립문 문화공원 지하공영주차장 정화조 설치공사", "기타", "2024-07-12", 0.8), ("독립문 문화공원 지하공영주차장 제례실 신축공사", "건축", "2024-08-01", 4)],
                      facility="독립문 문화공원 지하공영주차장")
    assert wk == [1, 1, 1, 2] and role == ["count"] * 4, wk
    wk, _, _ = run([("무등야구장 리모델링 사업 건축공사", "건축", "2020-03-17", 224), ("무등야구장 리모델링 사업 야구장·클라이밍장 조성공사", "건축", "2022-12-13", 31),
                    ("무등야구장 화장실 리모델링사업(실내건축)", "건축", "2021-06-02", 1.4)], facility="무등야구장", wtype="리모델링")
    assert wk[0] == wk[1] != wk[2], wk
    wk, _, _ = run([("작은공연장 단 리모델링 건축공사", "건축", "2019-05-01", 0.4), ("작은공연장 단 리모델링 공사", "기타", "2021-12-20", 0.4)], facility="작은공연장", wtype="리모델링")
    assert wk == [1, 2]
    # ⑧-2 (독립 검토에서 실제 이력으로 확인된 사례) 재공고 표기가 있으면 같은 날·다음 날 다시 난 것도 재공고, 'N차 재공고'·'재공고(3차)'·'새로운 입찰' 도 같은 공고로 본다
    wk, role, _ = run([("독도연구소 독도자연생태온실 신축공사 (건축)", "건축", "2021-12-20 10:10", 20), ("독도연구소 독도자연생태온실 신축공사 (건축) 재공고", "건축", "2021-12-20 17:28", 20),
                       ("독도연구소 독도자연생태온실 신축공사 (건축) 2차 재공고", "건축", "2021-12-27 17:17", 20), ("독도연구소 독도자연생태온실 신축공사 (건축) 새로운 입찰", "건축", "2022-01-20 11:10", 20),
                       ("독도연구소 독도자연생태온실 신축공사 (건축) 새로운 입찰 재공고", "건축", "2022-01-27 18:00", 20),
                       ("독도연구소 독도자연생태온실 신축공사 (전기)", "전기", "2021-12-20 10:09", 2), ("독도연구소 독도자연생태온실 신축공사 (전기) 재공고 (2차)", "전기", "2022-01-13 13:17", 2)],
                      facility="독도연구소 독도자연생태온실")
    assert set(wk) == {1} and role == ["superseded"] * 4 + ["count", "superseded", "count"], role
    # 한 번 하는 공사(건립·신축·증축·조성)가 같은 이름으로 1~3년 뒤 다시 발주되면 같은 사업의 재발주(한 건물을 두 번 세지 않는다). 연례 사업·같은 날 병행 발주는 해당 없음
    wk, role, _ = run([("도마실 주차전용건축물 조성공사(건축 등)", "건축", "2024-12-13", 39.5), ("도마실 주차전용건축물 조성공사(전기)", "전기", "2024-12-13", 3.2),
                       ("도마실 주차전용건축물 조성공사(건축 등)", "건축", "2025-11-06", 42.0)], facility="도마실 주차전용건축물")
    assert wk == [1, 1, 1] and role == ["superseded", "count", "count"], (wk, role)
    wk, role, _ = run([("가상박물관 수장고 증축공사(건축)", "건축", "2021-10-07", 3.7), ("가상박물관 수장고 증축공사(건축)", "건축", "2023-07-27", 4.7)], facility="가상박물관 수장고", wtype="증축")
    assert wk == [1, 1] and role == ["superseded", "count"], (wk, role)
    wk, role, _ = run([("가상박물관 수장고 증축공사(건축)", "건축", "2021-10-07", 3.7), ("가상박물관 수장고 증축공사(건축)", "건축", "2021-10-07", 3.7)], facility="가상박물관 수장고", wtype="증축")
    assert role == ["count", "count"], "재공고 표기 없이 같은 날 난 같은 이름은 병행 발주"
    # '(3-2단계)' 와 '(3-2)' 는 같은 단계 표기, 아파트 이름의 'N차'(이지더원2차·3차)는 다른 단지, 이름 끝 한 글자 오타는 빠진 공종을 채우는 공고끼리만 같은 이름
    wk, role, _ = run([("인천가족공원 조성사업(3-2단계) 봉안당 증축공사", "건축", "2025-06-19", 79.4), ("인천가족공원 조성(3-2) 봉안당 증축 전기공사", "전기", "2025-10-10", 8.5),
                       ("인천가족공원 조성(3-2) 봉안당 증축 소방전기공사", "소방", "2025-10-13", 2.7), ("인천가족공원 조성사업(3-2단계) 봉안당 증축 소방기계공사", "소방", "2025-11-20", 2.9)],
                      facility="인천광역시 봉안당", wtype="증축")
    assert wk == [1, 1, 1, 1] and role == ["count"] * 4, (wk, role)
    wk, _, _ = run([("국공립어린이집(일광이지더원2차)리모델링 공사", "기타", "2022-08-19", 0.67), ("국공립어린이집(일광이지더원3차)리모델링 공사", "기타", "2023-04-13", 0.75),
                    ("국공립어린이집(일광한신더휴1) 리모델링공사", "기타", "2022-06-22", 0.59), ("국공립어린이집(일광한신더휴2) 리모델링공사", "기타", "2022-06-22", 0.58)],
                   facility="국공립어린이집 리모델링", wtype="리모델링")
    assert len(set(wk)) == 4, wk
    wk, _, _ = run([("대전추모공원 제4봉안동 건립공사", "건축", "2023-11-27", 58.7), ("대전추모공원 제4봉안당 건립 기계설비공사", "기계설비", "2023-12-01", 8.9),
                    ("대전추모공원 제4봉안당 건립 전기공사", "전기", "2024-06-26", 7.5)], facility="대전추모공원")
    assert wk == [1, 1, 1], wk
    wk, _, _ = run([("화명생태공원 생활체육시설(축구장) 리모델링 공사", "기타", "2021-04-19", 1.59), ("화명생태공원 생활체육시설(야구장) 리모델링공사", "기타", "2021-06-01", 0.66)],
                   facility="화명생태공원 생활체육시설 리모델링", wtype="리모델링")
    assert wk == [1, 2], "같은 공종으로 나란히 난 한 글자 다른 이름은 다른 현장"
    # 시설명이 우연히 걸린 무관한 큰 공사('둔덕정수장 고도정수처리' 안의 '수장고')는 작은 본공사의 부대공사가 되지 않는다
    wk, _, _ = run([("둔덕정수장 고도정수처리시설 설치사업", "건축", "2019-04-24", 161.8), ("둔덕정수장 고도정수처리시설 설치사업(전기)", "전기", "2020-07-14", 24.4),
                    ("수장고 신축공사(건축)", "건축", "2020-09-18", 4.8), ("수장고 신축공사(전기)", "전기", "2020-09-18", 0.5), ("수장고 실내건축 공사", "건축", "2021-02-24", 4.6)],
                   facility="여수시 수장고")
    assert wk == [1, 1, 2, 2, 2], wk
    # ⑨ dedup 전체 흐름: 취소공고는 이름이 가장 비슷한 유효 공고의 세부사업을 따르고 대표가 되지 않는다. 유효 공고가 없는 시설은 '전부 취소공고'
    mk = lambda fid, no, name, dt, price, kind="일반", trade="건축": dict(시설ID=fid, 시설명="가상시 주차전용건축물", 공고번호=no, 공고차수="000", 공고키=no + "-000", 공고명=name,
                                                                        공고일시=dt, 사업유형="신축", 공종=trade, 추정가격=price, 공고종류=kind, 이전공고번호="",
                                                                        표2_대분류="자동차 관련 시설", 표2_중분류="주차장", 검색어="주차전용건축물", 수요기관="가상시")
    hist, latest, logs = dedup.dedup_latest(pd.DataFrame([
        mk("F1", "A", "서교동 공영주차장 주차전용 건축물 신축공사(건축, 기계)", "2019-03-29 10:00:00", 26.6e8), mk("F1", "B", "서교동 공영주차장 주차전용 건축물 신축공사(전기)", "2019-04-05 10:00:00", 2.4e8, trade="전기"),
        mk("F1", "C", "학동1 공영주차장 주차전용건축물 신축공사(건축,기계)", "2020-05-18 10:00:00", 39.4e8), mk("F1", "D", "학동1 공영주차장 주차전용건축물 신축공사(소방)", "2020-05-01 10:00:00", 5.5e8, kind="취소", trade="소방"),
        mk("F2", "E", "수산시장 주차전용건축물 신축공사", "2021-06-03 10:00:00", 25e8, kind="취소")]))
    h = hist.set_index("공고번호")
    assert h.at["A", "프로젝트ID"] == h.at["B", "프로젝트ID"] == "F1-N-01" and h.at["C", "프로젝트ID"] == h.at["D", "프로젝트ID"] == "F1-N-02", h["프로젝트ID"].to_dict()
    assert sorted(latest["공고번호"]) == ["A", "B", "C"] and h.at["D", "대표선정사유"] == "취소공고" and h.at["E", "대표선정사유"] == "전부 취소공고" and h.at["E", "프로젝트ID"] == "F2-N"
    assert h.at["C", "세부사업명"].startswith("학동1") and int(h.at["A", "세부사업수"]) == 2 and int(h.at["E", "세부사업수"]) == 1
    fs = dedup.facility_summary(latest, hist).set_index("프로젝트ID")
    assert list(fs.index) == ["F1-N-01", "F1-N-02"] and fs.at["F1-N-02", "세부사업명"].startswith("학동1") and fs.at["F1-N-01", "공종_전기"] == "O" and fs.at["F1-N-02", "공종_전기"] == "X"
    # ⑩ API 이전공고번호: 취소된 재공고는 원 공고를 대체하지 못한다 / 재공고의 변경 차수에서 이전공고번호가 빠져도 연결된다 /
    #    원 공고의 변경 차수가 재공고보다 늦게 나도 대체되는 쪽은 원 공고 / 공종이 빈 공고도 집계에서 빠지지 않는다 / 입력 행 순서가 달라도 결과가 같다
    def mk2(no, name, dt, price, kind="일반", trade="건축", ord_="000", prev=""):
        return dict(시설ID="F1", 시설명="가상시 회관", 공고번호=no, 공고차수=ord_, 공고키=f"{no}-{ord_}", 공고명=name, 공고일시=dt, 사업유형="신축", 공종=trade,
                    추정가격=price, 공고종류=kind, 이전공고번호=prev, 표2_대분류="문화 및 집회시설", 표2_중분류="-", 검색어="회관", 수요기관="가상시")
    hh, ll, _ = dedup.dedup_latest(pd.DataFrame([mk2("A", "가상시 회관 건립공사", "2024-01-10 10:00:00", 1e10), mk2("B", "가상시 회관 건립공사", "2024-02-10 10:00:00", 1e10, kind="재공고", prev="A"),
                                               mk2("B", "가상시 회관 건립공사", "2024-02-20 10:00:00", 1e10, kind="취소", ord_="001"), mk2("E", "가상시 회관 건립 전기공사", "2024-02-10 10:00:00", 1e9, trade="전기")]))
    assert sorted(ll["공고번호"]) == ["A", "E"], ll[["공고번호", "대표선정사유"]]
    hh, ll, _ = dedup.dedup_latest(pd.DataFrame([mk2("E1", "가상시 회관 신축공사 (전기)", "2021-12-20 10:09:02", 5e8, trade="전기"),
                                               mk2("E3", "가상시 회관 신축공사 (전기) 재공고 (2차)", "2022-01-13 11:45:11", 5e8, trade="전기", kind="재공고", prev="E1"),
                                               mk2("E3", "가상시 회관 신축공사 (전기) 재공고 (2차)", "2022-01-13 13:17:48", 5e8, trade="전기", kind="재공고", ord_="001")]))
    assert sorted(ll["공고키"]) == ["E3-001"] and hh.set_index("공고키").at["E1-000", "대체공고번호"] == "E3"
    hh, ll, _ = dedup.dedup_latest(pd.DataFrame([mk2("A", "가상시 회관 건립공사", "2024-01-10 10:00:00", 1e10), mk2("A", "가상시 회관 건립공사", "2024-03-01 10:00:00", 1e10, kind="변경", ord_="001"),
                                               mk2("B", "가상시 회관 건립공사", "2024-02-10 10:00:00", 1e10, kind="재공고", prev="A")]))
    assert sorted(ll["공고번호"]) == ["B"] and (hh[hh["공고번호"] == "B"]["대체공고번호"] == "").all(), hh[["공고키", "최신여부", "대체공고번호", "대표선정사유"]]
    rows = [mk2("A", "가상시 회관 건립공사", "2024-01-10 10:00:00", 1e10), mk2("N", "가상시 회관 건립 부대공사", "2024-02-10 10:00:00", 1e9, trade=None),
            mk2("C", "나래마을 회관 건립공사", "2024-02-11 10:00:00", 5e9), mk2("D", "나래마을 회관 건립 전기공사", "2024-03-11 10:00:00", 5e8, trade="전기")]
    h1, l1, _ = dedup.dedup_latest(pd.DataFrame(rows))
    assert sorted(l1["공고번호"]) == ["A", "C", "D", "N"] and l1.set_index("공고번호").at["N", "공종"] == "기타"
    h2, _, _ = dedup.dedup_latest(pd.concat([pd.DataFrame(rows[2:]), pd.DataFrame(rows[:2])]))      # 순서를 바꾸고 행 라벨도 겹치게
    cols = ["프로젝트ID", "세부사업번호", "세부사업명", "최신여부", "대표선정사유"]
    assert h1.set_index("공고키")[cols].sort_index().equals(h2.set_index("공고키")[cols].sort_index())


def test_work_split_review(tmp):
    """세부사업 분리 최종 점검(2026-11-15)에서 나온 결함·빈틈: 차수 괄호를 재공고로 오인, 제안 안내 공고 합산, 부지·이전·전체분, 연례 공사 간격,
    대체 사슬, 번호·날짜 형식, 시설 간 중복 프로젝트 표시 + 지금까지 테스트가 지키지 못하던 규칙들."""
    import openpyxl
    from g2b_costdb import works, dedup, pipeline
    from g2b_costdb.build_excel import build_workbook

    def mk(no, name, dt, price, kind="일반", trade="건축", ord_="000", prev="", wt="신축", fac="가상시 회관", rebid="N", fid="F1"):
        return dict(시설ID=fid, 시설명=fac, 공고번호=no, 공고차수=ord_, 공고키=f"{no}-{ord_}", 공고명=name, 공고일시=dt, 사업유형=wt, 공종=trade, 추정가격=price,
                    공고종류=kind, 재공고여부=rebid, 이전공고번호=prev, 표2_대분류="문화 및 집회시설", 표2_중분류="-", 검색어="회관", 수요기관="가상시")

    def run(rows):
        h, l, lg = dedup.dedup_latest(pd.DataFrame(rows), 0.30)
        return h.set_index("공고키"), sorted(l["공고번호"]), lg

    def split(rows, facility="가상시 회관", wtype="신축"):
        d = pd.DataFrame([dict(_base="F1-N", 시설명=facility, 사업유형=wtype, 공고번호=f"N{i:02d}", 공고명=nm, 공종=tr, _dt=pd.Timestamp(dt), _price=amt * 1e8, 이전공고번호="")
                          for i, (nm, tr, dt, amt) in enumerate(rows)])
        w = works.assign_works(d, "_base")
        return list(w["세부사업번호"]), list(w["역할"])

    # ① 차수·단계 표기가 다른 공고는 재공고가 아니라 차수별 발주 → 둘 다 합산(괄호로 적은 '(2차분)'·'(2차)' 포함, 실제: 홍천군 야구장 조성사업 / …(2차분))
    assert works.title_key("가상시 회관 조성사업(2차분)") != works.title_key("가상시 회관 조성사업")
    assert works.title_key("[긴급] 가상시 회관 건립공사(2차 변경)") == works.title_key("가상시 회관 건립공사")
    assert works.title_key("가상시 회관 건립공사(2차분 재공고)") == works.title_key("가상시 회관 건립공사(2차분)")
    for a, b, days in (("(1차분)", "(2차분)", 100), ("(1차)", "(2차)", 200), ("", "(2차분)", 480), ("", "(2차)", 188)):
        d2 = (pd.Timestamp("2024-01-10 10:00:00") + pd.Timedelta(days=days)).strftime("%Y-%m-%d %H:%M:%S")
        h, cnt, _ = run([mk("A", "가상시 회관 조성사업" + a, "2024-01-10 10:00:00", 50e8, trade="토목"), mk("B", "가상시 회관 조성사업" + b, d2, 60e8, trade="토목")])
        assert cnt == ["A", "B"] and h.at["A-000", "프로젝트ID"] == h.at["B-000", "프로젝트ID"] == "F1-N", (a, b, cnt, h["프로젝트ID"].to_dict())
    # 이름이 길어 두 공고명이 거의 같아 보여도(유사도 0.9 이상) 차수 표기가 다른 한 번짜리 공사는 한 세부사업의 차수별 발주
    h, cnt, _ = run([mk("A", "가상시 회관 주변 생태문화공원 기반시설 조성사업", "2020-02-20 10:00:00", 8.3e8, trade="토목", fac="가상시 생태문화공원"),
                     mk("B", "가상시 회관 주변 생태문화공원 기반시설 조성사업(2차분)", "2021-06-15 10:00:00", 10.8e8, trade="토목", fac="가상시 생태문화공원")])
    assert cnt == ["A", "B"] and set(h["프로젝트ID"]) == {"F1-N"}, h["프로젝트ID"].to_dict()
    # 같은 차수의 재공고는 그 차수만 대체한다(1차분은 남는다)
    h, cnt, _ = run([mk("A", "가상시 회관 조성사업", "2024-01-10 10:00:00", 50e8, trade="토목"), mk("B", "가상시 회관 조성사업(2차분)", "2024-06-10 10:00:00", 60e8, trade="토목"),
                     mk("C", "가상시 회관 조성사업(2차분 재공고)", "2024-07-10 10:00:00", 60e8, trade="토목", kind="재공고", rebid="Y")])
    assert cnt == ["A", "C"] and h.at["B-000", "대체공고번호"] == "C", (cnt, h["대표선정사유"].to_dict())

    # ② 공법·자재 선정용 제안서 제출 안내 공고는 공사 발주가 아니다 → 합산하지 않는다(실제: 명장1동 공영주차장 4.38억이 본 공사 8.93억에 더해짐)
    for name, want in (("「가상동 공영주차장 건립 사업」장경간 구조 특허 공법, 기술제안서 제출 안내 공고(변경)", True), ("가상 주차타워 조성사업 특허공법선정 기술제안 공모 안내공고", True),
                       ("「가상읍 공영주차타워 건립사업」건축구조(주차타워) 공법, 자재 제안서 제출안내", True), ("가상군 주차타워 조성사업 건축구조(주차장)공법제안 공고", True),
                       ("가상미술관 건립공사 실시설계 기술제안 입찰 재공고", False), ("가상택지 주차타워 건립공사(건축/특허공법)", False), ("가상 생산단지(공모) 조성사업 온실 신축공사", False),
                       ("가상 회관 건립공사 수의견적 제출 안내공고", False)):
        assert works.is_proposal(name) is want, name
    h, cnt, lg = run([mk("P", "「가상시 회관 건립 사업」장경간 구조 특허 공법, 기술제안서 제출 안내 공고", "2024-04-23 10:00:00", 4.38e8),
                      mk("A", "가상시 회관 건립공사(건축,기계)", "2024-08-02 10:00:00", 8.93e8), mk("E", "가상시 회관 건립 공사(전기)", "2024-08-19 10:00:00", 0.36e8, trade="전기")])
    assert cnt == ["A", "E"] and h.at["P-000", "대표선정사유"] == dedup.PROPOSAL_REASON and h.at["A-000", "대표선정사유"] == "단독 공고", h["대표선정사유"].to_dict()
    assert set(h["프로젝트ID"]) == {"F1-N"} and any(l["항목"] == "제안 안내 공고 제외" and l["공고번호"] == "P" for l in lg)
    h, cnt, lg = run([mk("P", "가상시 회관 조성사업 건축구조 공법제안 공고", "2024-04-23 10:00:00", 0)])                    # 제안 안내 공고뿐인 시설: 대표 없음 + 경고
    assert cnt == [] and h.at["P-000", "대표선정사유"] == dedup.PROPOSAL_REASON and any(l["항목"] == "대표선정" and l["판정"] == "경고" for l in lg)

    # ③ '부지'·'이전'·'전체분/총차분' 은 현장을 가리키는 말이 아니다 → 부지조성 토목 공고와 2년 뒤 건축·전기 공고가 한 세부사업(실제: 정남진 종합스포츠타운, 팽성야구장, 서군산 축구장)
    wk, _ = split([("가상 종합스포츠타운 부지조성공사(총괄 및 1차분)", "토목", "2021-12-01", 27.5), ("가상 종합스포츠타운 조성사업 건축공사(총괄 및 1차분)", "건축", "2023-12-15", 80.8),
                   ("가상 종합스포츠타운 조성사업 전기공사(총괄 및 1차분)", "전기", "2023-12-14", 8.2)], facility="가상 종합스포츠타운")
    assert wk == [1, 1, 1], wk
    wk, _ = split([("가상야구장 이전 조성사업(1구장)", "토목", "2025-07-23", 11.9), ("가상야구장 이전 조성 정보통신공사", "정보통신", "2025-09-15", 0.6),
                   ("가상야구장 이전 조성 전기공사", "전기", "2025-11-12", 3.7)], facility="가상야구장")
    assert wk == [1, 1, 1], wk
    wk, _ = split([("가상 축구장 조성공사(전체분 및 1차분)", "건축", "2020-10-23", 15.2), ("가상 축구장 전기공사(총차분 및 1차분)", "전기", "2020-10-29", 1.2),
                   ("가상 축구장 통신공사(총괄분 및 1차분)", "정보통신", "2020-10-29", 0.6)], facility="가상 축구장")
    assert wk == [1, 1, 1], wk
    assert works.features("서부지구 어린이전용 놀이터 조성공사")["disc"] == "서부지구어린이전용놀이터"                      # 낱말 속의 '부지'·'이전' 은 지우지 않는다
    assert works.features("가상 주차장부지 조성공사")["disc"] == works.features("가상 주차장 조성공사")["disc"] == "가상주차장"

    # ④ 유지관리·연간단가 공사는 같은 이름이 180일 넘게 떨어지면 다음 회차(둘 다 합산, 서로 다른 세부사업). 가까우면 재공고.
    #    뒤 공고가 '재공고' 여도 180일 넘게 떨어진 앞 공고는 지난 회차의 계약(그 재공고의 원 공고는 이번 회차에 났다가 취소·누락된 것)
    base = [mk("A", "가상시 회관 천연잔디 유지관리공사", "2021-06-11 10:00:00", 0.45e8, wt="리모델링")]
    h, cnt, _ = run(base + [mk("B", "가상시 회관 천연잔디 유지관리공사", "2022-03-08 10:00:00", 0.62e8, wt="리모델링")])       # 270일
    assert cnt == ["A", "B"] and h.at["A-000", "프로젝트ID"] != h.at["B-000", "프로젝트ID"], (cnt, h["프로젝트ID"].to_dict())
    h, cnt, _ = run(base + [mk("B", "가상시 회관 천연잔디 유지관리공사", "2021-09-19 10:00:00", 0.45e8, wt="리모델링")])       # 100일
    assert cnt == ["B"], cnt
    h, cnt, _ = run(base + [mk("B", "가상시 회관 천연잔디 유지관리공사 재공고", "2022-02-16 10:00:00", 0.45e8, wt="리모델링", kind="재공고", rebid="Y")])   # 250일, 재공고 표기
    assert cnt == ["A", "B"], cnt
    h, cnt, _ = run(base + [mk("B", "가상시 회관 천연잔디 유지관리공사", "2021-11-28 10:00:00", 0.45e8, wt="리모델링")])       # 170일: 아직 재공고로 본다
    assert cnt == ["B"], cnt
    h, cnt, _ = run(base + [mk("B", "가상시 회관 천연잔디 유지관리공사", "2021-12-28 10:00:00", 0.45e8, wt="리모델링")])       # 200일: 다음 회차
    assert cnt == ["A", "B"], cnt
    yr = [mk("A", "가상시 회관 시설물 보수 연간단가공사", "2021-06-11 10:00:00", 1e8, wt="리모델링"), mk("B", "가상시 회관 시설물 보수 연간단가공사", "2021-12-13 10:00:00", 1e8, wt="리모델링")]
    assert run(yr)[1] == ["A", "B"], "'연간단가' 도 185일 뒤 같은 이름이면 다음 회차"
    # 취소된 원 공고 + 그 재공고(올해 계약) + 270일 뒤 다음 회차: 두 회차 모두 합산
    h, cnt, _ = run([mk("A", "가상시 회관 천연잔디 유지관리공사", "2021-06-11 10:00:00", 0.45e8, wt="리모델링", kind="취소"),
                     mk("A2", "가상시 회관 천연잔디 유지관리공사", "2021-06-25 10:00:00", 0.45e8, wt="리모델링", kind="재공고", rebid="Y"),
                     mk("B", "가상시 회관 천연잔디 유지관리공사", "2022-03-08 10:00:00", 0.62e8, wt="리모델링")])
    assert cnt == ["A2", "B"], cnt
    # 건물 이름 '유지관리동', 낱말 속 글자 '산단가로등' 은 연례 공사가 아니다 → 200일 뒤 같은 이름이면 재공고
    assert run([mk("A", "가상시 회관 유지관리동 신축공사", "2024-01-10 10:00:00", 20e8), mk("B", "가상시 회관 유지관리동 신축공사", "2024-07-28 10:00:00", 20e8)])[1] == ["B"]
    assert run([mk("A", "가상산단가로등 설치공사", "2024-01-10 10:00:00", 3e8, trade="전기", fac="가상산단"), mk("B", "가상산단가로등 설치공사", "2024-07-28 10:00:00", 3e8, trade="전기", fac="가상산단")])[1] == ["B"]
    h, cnt, _ = run([mk("A", "가상시 회관 관리동 증축공사", "2021-06-11 10:00:00", 5e8, wt="증축"), mk("B", "가상시 회관 관리동 증축공사", "2022-02-16 10:00:00", 5e8, wt="증축")])
    assert cnt == ["B"], "'관리동 증축' 같은 한 번짜리 공사는 250일 뒤 같은 이름이면 재공고"

    # ⑤ 재공고가 거듭되면(A ← B ← C) 앞 공고들의 대체공고번호는 끝까지 남은 공고를 가리킨다. 서로를 이전 공고로 가리키는 잘못된 연결에도 하나는 남는다
    h, cnt, _ = run([mk("A", "가상시 회관 건립공사", "2024-01-10 10:00:00", 1e10), mk("B", "가상시 회관 건립공사", "2024-02-10 10:00:00", 1e10, kind="재공고", prev="A"),
                     mk("C", "가상시 회관 건립공사", "2024-03-10 10:00:00", 1e10, kind="재공고", prev="B")])
    assert cnt == ["C"] and h.at["A-000", "대체공고번호"] == h.at["B-000", "대체공고번호"] == "C" and "재공고로 2건 대체" in h.at["C-000", "대표선정사유"], h[["대체공고번호", "대표선정사유"]]
    h, cnt, _ = run([mk("A", "가상시 회관 건립공사", "2024-01-10 10:00:00", 1e10, prev="B"), mk("B", "가상시 문화회관 건립공사(건축)", "2024-02-10 10:00:00", 1e10, prev="A")])
    assert cnt == ["B"], cnt
    h, cnt, _ = run([mk("A", "가상시 회관 건립공사", "2024-01-10 10:00:00", 1e10, prev="B"), mk("B", "가상시 회관 리모델링공사", "2024-02-10 10:00:00", 1e10, prev="A", wt="리모델링")])
    assert len(cnt) == 1, cnt                                                                                               # 사업유형이 달라 서로 다른 프로젝트에 놓인 순환 연결

    # ⑥ 번호·날짜 형식: 숫자로 읽혀 '.0' 이 붙은 이전공고번호, 공백이 붙은 공고번호, 한 열에 섞인 날짜 형식
    assert dedup._no(20240101.0) == "20240101" and dedup._no(" A1 ") == "A1" and dedup._no(float("nan")) == "" and dedup._no(None) == ""
    dts = dedup._parse_dt(pd.Series(["2024-01-10 10:00:00", "2024-02-10T10:00:00", "202403101000", "", None]))
    assert list(dts.dt.strftime("%Y-%m-%d").fillna("")) == ["2024-01-10", "2024-02-10", "2024-03-10", "", ""], dts
    rows = [mk(20240101, "가상시 회관 건립공사", "2024-01-10 10:00:00", 1e10), mk(20240202, "가상시 문화회관 건립공사(건축)", "2024-02-10T10:00:00", 1.3e10)]
    rows[0]["이전공고번호"], rows[1]["이전공고번호"] = float("nan"), 20240101.0
    h, cnt, _ = run(rows)
    assert cnt == [20240202] and h.at["20240101-000", "대체공고번호"] == "20240202", (cnt, h["대체공고번호"].to_dict())

    # ⑦ 유효 행이 없는 공고(모든 차수가 취소·구차수)는 그 공고의 최신 차수 이름으로 세부사업을 고른다 — 현장 이름이 같은 사업, 한 공고의 모든 차수는 같은 프로젝트
    f = "가상공사 공공주택"
    h, cnt, _ = run([mk("P", "가상공사 1단지 아파트 건설공사", "2024-01-10 10:00:00", 1e11, fac=f), mk("P2", "가상공사 1단지 아파트 전기공사", "2024-02-10 10:00:00", 1e10, fac=f, trade="전기"),
                     mk("Q", "가상공사 2단지 아파트 건설공사", "2024-01-10 10:00:00", 1e11, fac=f), mk("Q2", "가상공사 2단지 아파트 소방공사", "2024-02-10 10:00:00", 1e10, fac=f, trade="소방"),
                     mk("X", "가상공사 2단지 아파트 전기공사", "2024-03-10 10:00:00", 1e10, fac=f, trade="전기", kind="취소"),
                     mk("Z", "가상공사 1단지 아파트 전기공사", "2024-03-01 10:00:00", 1e10, fac=f, trade="전기"),
                     mk("Z", "가상공사 2단지 아파트 전기공사", "2024-03-20 10:00:00", 1e10, fac=f, trade="전기", ord_="001", kind="취소")])
    assert h.at["X-000", "프로젝트ID"] == h.at["Q-000", "프로젝트ID"] != h.at["P-000", "프로젝트ID"], h["프로젝트ID"].to_dict()        # 이름이 더 비슷한 1단지 전기 공고가 아니라 2단지 사업
    assert h.at["Z-000", "프로젝트ID"] == h.at["Z-001", "프로젝트ID"] == h.at["Q-000", "프로젝트ID"], h["프로젝트ID"].to_dict()

    # ⑧ 지금까지 테스트가 지키지 못하던 규칙(변이 검사에서 살아남은 것)
    h, cnt, lg = run([mk("A", "가상시 회관 건립공사", "2024-01-10 10:00:00", 1e10), mk("B", "가상시 회관 리모델링공사", "2024-02-10 10:00:00", 1e10, prev="A", wt="리모델링")])
    assert cnt == ["B"] and h.at["A-000", "대표선정사유"] == "대체됨(API 이전공고번호 연결)" and h.at["A-000", "대체공고번호"] == "B" and any(x["항목"] == "이전공고번호 연결" for x in lg)
    _, cnt, _ = run([mk("A", "가상시 회관 보수공사", "2024-01-10 10:00:00", 1e9, wt="리모델링"), mk("B", "가상시 회관 보수공사 재공고", "2024-01-10 17:00:00", 1e9, wt="리모델링", kind="재공고", rebid="Y")])
    assert cnt == ["B"], "재공고라고 밝힌 공고는 같은 날 다시 나도 앞 공고를 대체"
    _, cnt, _ = run([mk("A", "가상시 회관 보수공사", "2024-01-10 10:00:00", 1e9, wt="리모델링"), mk("B", "가상시 회관 보수공사", "2024-03-10 10:00:00", 3e9, wt="리모델링")])
    assert cnt == ["A", "B"], "같은 이름이라도 금액이 크게 다르면(0.6~1.67배 밖) 별개 계약"
    h, _, _ = run([mk("P", "가상공사 1단지 아파트 건설공사", "2024-01-10 10:00:00", 1e11, fac=f), mk("Q", "가상공사 2단지 아파트 건설공사", "2024-01-10 10:00:00", 1e11, fac=f),
                   mk("Z", "가상공사 2단지 아파트 건설공사", "2024-03-10 10:00:00", 1e10, fac=f, trade="전기"),
                   mk("Z", "가상공사 1단지 아파트 전기공사", "2024-03-20 10:00:00", 1e10, fac=f, trade="전기", ord_="001", kind="변경")])
    assert h.at["Z-000", "프로젝트ID"] == h.at["Z-001", "프로젝트ID"] == h.at["P-000", "프로젝트ID"], "구차수는 같은 공고번호의 최신 차수와 같은 프로젝트"
    h, cnt, _ = run([mk("A", "가상시 회관 건립공사", "2024-01-10 10:00:00", 1e10), mk("B", "가상시 문화회관 건립공사(건축)", "2024-02-10 10:00:00", 1.3e10, prev="A"),
                     mk("B", "가상시 문화회관 건립공사(건축)", "2024-02-12 10:00:00", 1.3e10, ord_="001", kind="변경")])
    assert cnt == ["B"] and h.at["A-000", "대체공고번호"] == "B", "변경 차수에서 이전공고번호가 빠져도 연결"
    h, _, _ = run([mk("A", "가상시 회관 리모델링공사", "2024-01-10 10:00:00", 1e10, wt="리모델링"), mk("A", "가상시 회관 증축공사", "2024-01-20 10:00:00", 1e10, ord_="001", kind="변경", wt="증축")])
    assert h.at["A-000", "프로젝트ID"] == h.at["A-001", "프로젝트ID"] == "F1-E", "한 공고의 모든 차수는 최신 차수의 사업유형을 따른다"
    h, cnt, lg = run([mk("A", "가상시 회관 건립공사", "2024-01-10 10:00:00", 1.0e10), mk("B", "가상시 회관 건립공사", "2024-03-10 10:00:00", 1.5e10)])
    assert cnt == ["B"] and any(x["항목"] == "재발주 금액변동" and x["판정"] == "경고" and x["공고번호"] == "B" for x in lg), lg
    wk, _ = split([("가상군 생태공원 북면지구 조성사업", "조경", "2024-03-01", 9), ("가상군 생태공원 북면지구 영월군청 조성사업", "조경", "2024-03-01", 7)], facility="가상군 생태공원")
    assert wk == [1, 2], "같은 공종이면 이름이 짧은 쪽도 독립된 계약(포함 관계로 묶지 않는다)"
    h, cnt, _ = run([mk("A", "가상공사 1단지 아파트 건설공사", "2024-01-10 10:00:00", 1e11, fac=f), mk("B", "가상공사 가람마을 아파트 건설공사", "2024-03-10 10:00:00", 1e11, fac=f, prev="A", kind="재공고"),
                     mk("C", "가상공사 2단지 아파트 건설공사", "2024-01-10 10:00:00", 1e11, fac=f)])
    assert h.at["A-000", "프로젝트ID"] == h.at["B-000", "프로젝트ID"] and cnt == ["B", "C"], "API 이전공고번호로 연결된 공고는 이름이 달라도 같은 세부사업"

    # ⑨ 시설 간 중복 프로젝트: 같은 공고가 여러 시설 후보에 걸려 같은 사업이 두 번 집계되는 경우를 표시하고, 뺄 행을 권장한다
    umb, own, big, part = "가상공사 공공주택", "가상 1단지 아파트", "가상 2단지 아파트", "가상 3단지 상가"
    rows = [mk("A1", "가상 1단지 아파트 건설공사", "2024-01-10 10:00:00", 9e10, fac=umb, fid="F9"), mk("A2", "가상 1단지 아파트 전기공사", "2024-02-10 10:00:00", 9e9, fac=umb, fid="F9", trade="전기"),
            mk("B1", "가상 2단지 아파트 건설공사", "2024-01-10 10:00:00", 8e10, fac=umb, fid="F9"),
            mk("C1", "가상 3단지 아파트 건설공사", "2024-01-10 10:00:00", 7e10, fac=umb, fid="F9"), mk("C2", "가상 3단지 아파트 소방공사", "2024-02-10 10:00:00", 7e9, fac=umb, fid="F9", trade="소방"),
            mk("A1", "가상 1단지 아파트 건설공사", "2024-01-10 10:00:00", 9e10, fac=own, fid="F2"), mk("A2", "가상 1단지 아파트 전기공사", "2024-02-10 10:00:00", 9e9, fac=own, fid="F2", trade="전기"),
            mk("B1", "가상 2단지 아파트 건설공사", "2024-01-10 10:00:00", 8e10, fac=big, fid="F3"), mk("B9", "가상 2단지 아파트 전기공사", "2024-02-10 10:00:00", 8e9, fac=big, fid="F3", trade="전기"),
            mk("C2", "가상 3단지 아파트 소방공사", "2024-02-10 10:00:00", 7e9, fac=part, fid="F4", trade="소방"), mk("C9", "가상 3단지 아파트 통신공사", "2024-02-10 10:00:00", 1e9, fac=part, fid="F4", trade="정보통신")]
    hist, latest, _ = dedup.dedup_latest(pd.DataFrame(rows), 0.30)
    pid = hist.drop_duplicates(["시설ID", "공고번호"]).set_index(["시설ID", "공고번호"])["프로젝트ID"]
    dup = dedup.cross_facility_duplicates(latest)
    a9, b9, c9 = pid[("F9", "A1")], pid[("F9", "B1")], pid[("F9", "C1")]
    assert len({a9, b9, c9}) == 3 and pid[("F2", "A1")] == "F2-N" and pid[("F3", "B1")] == "F3-N", pid.to_dict()
    assert dup.at[a9, "시설간_중복"] == "같은 사업: F2-N" and dup.at[a9, "중복_제외권장"] == "Y" and dup.at["F2-N", "중복_제외권장"] == "", dup.to_dict()      # 묶음 시설의 세부사업 쪽을 뺀다
    assert dup.at[b9, "시설간_중복"] == "F3-N 의 일부" and dup.at[b9, "중복_제외권장"] == "Y" and dup.at["F3-N", "시설간_중복"] == f"{b9} 를 포함(1/2건)" and dup.at["F3-N", "중복_제외권장"] == ""
    assert dup.at[c9, "시설간_중복"] == "F4-N 와 일부 겹침(1/2건)" and dup.at[c9, "중복_제외권장"] == "" and dup.at["F4-N", "중복_제외권장"] == "", dup.to_dict()     # 일부만 겹치면 권장 없음
    keep = latest[~latest["프로젝트ID"].map(lambda x: x in dup.index and dup.at[x, "중복_제외권장"] == "Y")]
    assert sorted(keep["공고번호"]) == ["A1", "A2", "B1", "B9", "C1", "C2", "C2", "C9"]                                   # Y 행을 빼면 일부 겹침(C2)만 두 번 남는다
    assert dedup.cross_facility_duplicates(latest.iloc[0:0]).empty and list(dedup.cross_facility_duplicates(pd.DataFrame()).columns) == ["시설간_중복", "중복_제외권장"]
    fs = dedup.facility_summary(latest, hist).set_index("프로젝트ID")
    assert fs.at[a9, "중복_제외권장"] == "Y" and fs.at["F2-N", "시설간_중복"] == f"같은 사업: {a9}" and fs.at["F4-N", "시설간_중복"] == f"{c9} 와 일부 겹침(1/2건)"
    fac = pipeline.enrich_facility(fs.reset_index(), latest, {})
    out = os.path.join(tmp, "dup.xlsx")
    build_workbook(out, fac, hist, latest, pipeline.assemble_trade_table(latest, {}, {}), pd.DataFrame(), pd.DataFrame(), meta={"period": "t"})
    ws5 = openpyxl.load_workbook(out)["05_공사비DB_시설합산"]
    h5 = [c.value for c in ws5[1]]
    col = lambda name: {ws5.cell(r, 1).value: ws5.cell(r, h5.index(name) + 1).value for r in range(2, ws5.max_row + 1)}
    assert col("중복_제외권장")[a9] == "Y" and col("시설간_중복")[a9] == "같은 사업: F2-N" and col("중복_제외권장")["F2-N"] in ("", None), (h5, col("중복_제외권장"))

    # ⑩ 2차 검토(2026-11-15)에서 나온 결함
    # 괄호 속 숫자·차수의 두 가지 뜻: '(2회)'·'(제2회)'·'(2차 재공고)' 는 공고 회차(재공고), '(2차)'·'(2차, 긴급)' 는 차수 발주
    for tag_, kw in (("(2회)", dict(kind="재공고", rebid="Y")), ("(제2회)", {})):
        assert run([mk("A", "가상시 회관 건립공사", "2024-01-10 10:00:00", 1e10), mk("B", "가상시 회관 건립공사" + tag_, "2024-01-30 10:00:00", 1e10, **kw)])[1] == ["B"], tag_
    h, cnt, _ = run([mk("A", "가상시 회관 조성사업", "2024-01-10 10:00:00", 50e8, trade="토목"), mk("B", "가상시 회관 조성사업(2차, 긴급)", "2024-07-16 10:00:00", 60e8, trade="토목")])
    assert cnt == ["A", "B"], "긴급 표기가 함께 든 괄호의 '2차' 도 차수 발주"
    # '(2차)' 가 금액이 그대로인 채 다시 난 공고면 '두 번째 공고'(재공고): 재공고 표기가 있든 없든. 금액이 다르면 차수 발주(위 ①)
    assert run([mk("A", "가상시 회관 건립공사", "2024-01-10 10:00:00", 1e10), mk("B", "가상시 회관 건립공사(2차)", "2024-01-30 10:00:00", 1e10, kind="재공고", rebid="Y")])[1] == ["B"]
    assert run([mk("A", "가상시 회관 건립공사", "2024-01-10 10:00:00", 1e10), mk("B", "가상시 회관 건립공사(2차)", "2024-01-30 10:00:00", 1e10)])[1] == ["B"]
    assert run([mk("A", "가상시 회관 건립공사", "2024-01-10 10:00:00", 1e10), mk("B", "가상시 회관 건립공사(2차)", "2024-01-25 10:00:00", 1e10, kind="재공고", rebid="Y"),
                mk("C", "가상시 회관 건립공사(3차)", "2024-02-09 10:00:00", 1e10, kind="재공고", rebid="Y")])[1] == ["C"]
    # 1차(표기 없음) + '(2차)' + '(2차 재공고)': 재공고는 2차 공고를 대체하고 1차는 남는다
    h, cnt, _ = run([mk("A", "가상시 회관 조성사업", "2024-01-10 10:00:00", 50e8, trade="토목"), mk("B", "가상시 회관 조성사업(2차)", "2024-06-08 10:00:00", 60e8, trade="토목"),
                     mk("C", "가상시 회관 조성사업(2차 재공고)", "2024-06-28 10:00:00", 60e8, trade="토목", kind="재공고", rebid="Y")])
    assert cnt == ["A", "C"] and h.at["B-000", "대체공고번호"] == "C", (cnt, h["대표선정사유"].to_dict())
    # 제안 안내 공고는 무엇을 제안받는지(공법·자재·신기술·특허)가 이름에 있어야 한다 — 협상계약의 '제안서 제출 안내' 는 실제 발주
    for name, want in (("가상 전시관 전시물 제작·설치 제안서 제출 안내 공고", False), ("가상 회관 건립공사 설계공모 안내공고", False),
                       ("가상 주차타워 조성사업 특허공법 선정 공모", True), ("가상 주차타워 건립사업 구조 자재 제안 공고", True),
                       ("가상 주차타워 건립사업 장경간 특허 기술제안 안내", True), ("가상 주차타워 건립사업 특허공법 제안서 제출 공고", True)):
        assert works.is_proposal(name) is want, name
    # '부지'·'이전' 은 낱말로 쓰였을 때만 일반어: '본부지하주차장', 장소 이름 '고수부지' 는 그대로. '전체 분전반' 의 '분' 도 그대로
    assert "본부지하주차장" in works.features("가상공단 본부지하주차장 방수공사")["disc"] and "고수부지" in works.features("고수부지(한빛공원) 주차장 조성공사")["disc"]
    assert "분전반" in works.title_key("가상시 회관 전체 분전반 교체공사") and works.clean_title("가상 축구장 조성공사(전체분 및 1차분)") == "가상 축구장 조성공사(1차분)"
    # '부지조성공사' 와 '조성공사' 는 이름·금액이 비슷해도 다른 계약(재공고 아님) / 시설명에 '부지' 가 들어 있어도 그 시설의 공종 공고와 한 사업
    assert run([mk("A", "가상시 회관 주차장 부지조성공사", "2024-01-10 10:00:00", 10e8, trade="토목"), mk("B", "가상시 회관 주차장 조성공사", "2024-03-10 10:00:00", 11e8, trade="토목")])[1] == ["A", "B"]
    wk, _ = split([("가상 반환부지 문화공원 조성공사", "건축", "2024-01-10", 50), ("문화공원 조성 전기공사", "전기", "2024-07-28", 5)], facility="가상 반환부지 문화공원")
    assert wk == [1, 1], wk
    # 지구 전체의 '부지조성공사' 는 이름이 다른 공고(지장물 해체 등)의 본공사 노릇을 하지 않는다(2년 넘게 떨어진 별개 공사가 한 프로젝트가 되지 않게)
    wk, _ = split([("가상 공공주택지구(문화재2구역)지장물해체공사", "건축", "2023-11-30", 60.4), ("가상 공공주택지구 부지조성공사(2공구)", "토목", "2026-03-16", 544.5)], facility="가상 공공주택지구")
    assert wk == [1, 2], wk
    # 이전공고번호 사슬: 가운데 재공고가 취소돼도(A ← B[취소] ← C, C 는 이름이 바뀜) C 가 A 를 대체 / 세 공고가 서로를 가리켜도 하나는 남는다
    h, cnt, _ = run([mk("A", "가상시 회관 건립공사", "2024-01-10 10:00:00", 1e10), mk("B", "가상시 회관 건립공사", "2024-01-30 10:00:00", 1e10, kind="취소", prev="A"),
                     mk("C", "가상시 문화회관 건립공사(건축) 새로운 입찰", "2024-02-19 10:00:00", 1.8e10, prev="B", kind="재공고", rebid="Y")])
    assert cnt == ["C"] and h.at["A-000", "대체공고번호"] == "C", (cnt, h["대표선정사유"].to_dict())
    _, cnt, _ = run([mk("A", "가상시 회관 건립공사", "2024-01-10 10:00:00", 1e10, prev="C"), mk("B", "가상시 회관 건립공사(건축)", "2024-01-30 10:00:00", 1e10, prev="A"),
                     mk("C", "가상시 회관 건립공사(건축,기계)", "2024-02-19 10:00:00", 1e10, prev="B")])
    assert len(cnt) >= 1, cnt
    # 사업유형이 달라 다른 프로젝트에 놓인 이전공고번호 연결: 번호가 숫자·소수로 읽혀도 연결, 자기 자신을 가리키면 무시, 취소된 재공고는 대체하지 못함
    rows = [mk(20240101, "가상시 회관 건립공사", "2024-01-10 10:00:00", 1e10), mk(20240202, "가상시 회관 리모델링공사", "2024-02-10 10:00:00", 1e10, wt="리모델링")]
    rows[0]["이전공고번호"], rows[1]["이전공고번호"] = float("nan"), 20240101.0
    h, cnt, _ = run(rows)
    assert cnt == [20240202] and h.at["20240101-000", "대체공고번호"] == "20240202" and h.at["20240101-000", "대표선정사유"] == "대체됨(API 이전공고번호 연결)", h[["대체공고번호", "대표선정사유"]]
    assert run([mk("A", "가상시 회관 건립공사", "2024-01-10 10:00:00", 1e10, prev="A")])[1] == ["A"]
    assert run([mk("A", "가상시 회관 건립공사", "2024-01-10 10:00:00", 1e10), mk("B", "가상시 회관 리모델링공사", "2024-02-10 10:00:00", 1e10, prev="A", wt="리모델링", kind="취소")])[1] == ["A"]
    # 유효 행이 없는 공고: 이름이 똑같은 공고가 2년 전에 있어도, 가까운 때에 난 같은 현장의 공고가 있는 세부사업을 따른다
    h, cnt, _ = run([mk("O", "가상시 회관 보수공사(건축)", "2022-03-10 10:00:00", 2e8, wt="리모델링"), mk("N", "가상시 회관 보수공사", "2024-03-10 10:00:00", 2.2e8, wt="리모델링"),
                     mk("X", "가상시 회관 보수공사(건축)", "2024-02-10 10:00:00", 2.2e8, wt="리모델링", kind="취소")])
    assert h.at["O-000", "프로젝트ID"] != h.at["N-000", "프로젝트ID"] == h.at["X-000", "프로젝트ID"], h["프로젝트ID"].to_dict()
    # 미분류 공종 공고의 사업유형 상속도 섞인 날짜 형식을 읽는다(가장 가까운 건축 공고 = 2024년 리모델링)
    h, _, _ = run([mk("1", "가상시 회관 건립공사", "2015-03-10 10:00:00", 1e10), mk("2", "가상시 회관 리모델링공사", "2024-02-10 10:00:00", 2e9, wt="리모델링"),
                   mk("3", "가상시 회관 전기공사", "2024-03-10T10:00:00", 3e8, wt="미분류", trade="전기")])
    assert h.at["3-000", "사업유형"] == "리모델링", h["사업유형"].to_dict()
    # 시설 간 중복: 남길 행은 세부사업으로 나뉘지 않은 시설 쪽(프로젝트ID 순서와 무관) / 공고번호의 공백 차이는 같은 공고 /
    # 한 프로젝트의 공고가 다른 시설의 여러 세부사업에 나뉘어 모두 들어 있으면 그 프로젝트를 빼고 세부사업들은 남긴다
    rows = [mk("A1", "가상 1단지 아파트 건설공사", "2024-01-10 10:00:00", 9e10, fac=umb, fid="F1"), mk("B1", "가상 2단지 아파트 건설공사", "2024-01-10 10:00:00", 8e10, fac=umb, fid="F1"),
            mk("A1 ", "가상 1단지 아파트 건설공사", "2024-01-10 10:00:00", 9e10, fac=own, fid="F7")]
    _, latest, _ = dedup.dedup_latest(pd.DataFrame(rows), 0.30)
    dup = dedup.cross_facility_duplicates(latest)
    assert dup.at["F1-N-01", "중복_제외권장"] == "Y" and dup.at["F7-N", "중복_제외권장"] == "" and dup.at["F7-N", "시설간_중복"] == "같은 사업: F1-N-01", dup.to_dict()
    own2, umb2 = "가상군 청년공공임대주택", "청년공공임대주택"                                                             # 실제 모양: 한 시설은 한 프로젝트, 묶음 시설은 두 세부사업
    rows = [mk("D", "가상군 청년공공임대주택건립사업 폐교 철거공사", "2024-08-20 10:00:00", 2.8e8, fac=own2, fid="F5"),
            mk("M", "가상군 청년공공임대주택 건립사업 건축·토목·조경·기계공사", "2024-12-02 10:00:00", 60.3e8, fac=own2, fid="F5"),
            mk("D", "가상군 청년공공임대주택건립사업 폐교 철거공사", "2024-08-20 10:00:00", 2.8e8, fac=umb2, fid="F6"),
            mk("M", "가상군 청년공공임대주택 건립사업 건축·토목·조경·기계공사", "2024-12-02 10:00:00", 60.3e8, fac=umb2, fid="F6"),
            mk("E", "청년공공임대주택 건립사업 전기공사", "2025-02-26 10:00:00", 7e8, fac=umb2, fid="F6", trade="전기")]
    hist, latest, _ = dedup.dedup_latest(pd.DataFrame(rows), 0.30)
    pid = hist.set_index(["시설ID", "공고번호"])["프로젝트ID"]
    assert pid[("F5", "D")] == pid[("F5", "M")] == "F5-N" and pid[("F6", "D")] != pid[("F6", "M")] == pid[("F6", "E")], pid.to_dict()     # 전제(바뀌면 이 사례를 다시 짤 것)
    dup = dedup.cross_facility_duplicates(latest)
    b1, b2 = pid[("F6", "D")], pid[("F6", "M")]
    assert dup.at["F5-N", "중복_제외권장"] == "Y" and dup.at["F5-N", "시설간_중복"].startswith(f"{b1} + {b2} 에 나뉘어 있음"), dup.to_dict()
    assert dup.at[b1, "중복_제외권장"] == "" and dup.at[b2, "중복_제외권장"] == "", dup.to_dict()
    kept = latest[~latest["프로젝트ID"].isin(dup.index[dup["중복_제외권장"] == "Y"])]
    assert sorted(kept["공고번호"]) == ["D", "E", "M"], sorted(kept["공고번호"])
    sets_ = pd.DataFrame([dict(프로젝트ID="P1", 시설ID="F5", 공고번호="X", 세부사업수=1), dict(프로젝트ID="P1", 시설ID="F5", 공고번호="Y", 세부사업수=1),
                          dict(프로젝트ID="Q1", 시설ID="F6", 공고번호="X", 세부사업수=2), dict(프로젝트ID="Q2", 시설ID="F6", 공고번호="Y", 세부사업수=2),
                          dict(프로젝트ID="Q2", 시설ID="F6", 공고번호="Z", 세부사업수=2)])
    dup = dedup.cross_facility_duplicates(sets_)
    assert dup.at["P1", "중복_제외권장"] == "Y" and dup.at["P1", "시설간_중복"].startswith("Q1 + Q2 에 나뉘어 있음") and dup.at["Q1", "중복_제외권장"] == "" and dup.at["Q2", "중복_제외권장"] == "", dup.to_dict()
    kept = sets_[~sets_["프로젝트ID"].isin(dup.index[dup["중복_제외권장"] == "Y"])]
    assert sorted(kept["공고번호"]) == ["X", "Y", "Z"]                                                                       # 빠지는 공고도, 두 번 남는 공고도 없다


def test_work_split_review2(tmp):
    """세부사업 분리 2차 독립 검토(2026-11-16)에서 나온 결함·빈틈: 괄호 'N차' 를 공고 회차로 보는 판정의 가드, 'N차 재공고' 의 여러 표기, 여러 차수를 묶은 괄호,
    '부지' 본공사 규칙의 범위, 연례 공사 오인, 제안 안내 공고의 범위, 이전공고번호 사슬과 대체공고번호, 취소공고의 대체공고번호, 공고연도, 시설 간 중복 요약 문구."""
    import io
    import contextlib
    from g2b_costdb import works, dedup, pipeline

    t0 = pd.Timestamp("2024-01-10 10:00:00")

    def mk(no, name, day, eok, kind="일반", trade="건축", ord_="000", prev="", wt="신축", fac="가상시 회관", rebid="N", fid="F1"):
        dt = day if isinstance(day, str) else (t0 + pd.Timedelta(days=day)).strftime("%Y-%m-%d %H:%M:%S")
        return dict(시설ID=fid, 시설명=fac, 공고번호=no, 공고차수=ord_, 공고키=f"{no}-{ord_}", 공고명=name, 공고일시=dt, 사업유형=wt, 공종=trade,
                    추정가격=(None if eok is None else eok * 1e8), 공고종류=kind, 재공고여부=rebid, 이전공고번호=prev,
                    표2_대분류="문화 및 집회시설", 표2_중분류="-", 검색어="회관", 수요기관="가상시")

    def run(rows):
        h, l, lg = dedup.dedup_latest(pd.DataFrame(rows), 0.30)
        return h.set_index("공고키"), sorted(str(x).strip() for x in l["공고번호"]), lg

    def check(tag, rows, want, projects=None):
        h, cnt, _ = run(rows)
        assert cnt == sorted(want), (tag, cnt, h["대표선정사유"].to_dict())
        if projects is not None:
            assert h["프로젝트ID"].nunique() == projects, (tag, h["프로젝트ID"].to_dict())
        return h

    RB, T = dict(kind="재공고", rebid="Y"), dict(trade="토목")
    N, B = "가상시 회관 조성사업", "가상시 회관 건립공사"

    # ① 괄호로만 적은 '(2차)': 차수 발주일 수도, 두 번째 공고(재공고)일 수도 있다 → 금액이 그대로(±1%)이고 앞 공고에 차수 표기가 없으며,
    #    재공고라고 밝혔거나 120일 안에 난 경우에만 재공고로 본다
    check("(1차)+(2차) 같은 금액 → 차수별 발주", [mk("A", N + "(1차)", 0, 50, **T), mk("B", N + "(2차)", 150, 50, **T)], ["A", "B"])
    check("(1차)+(2차) 같은 금액 20일", [mk("A", B + "(1차)", 0, 100), mk("B", B + "(2차)", 20, 100)], ["A", "B"])
    check("금액이 없는 '(2차)' 는 금액이 그대로인 재공고가 아니다", [mk("A", N, 0, 50, **T), mk("B", N + "(2차)", 150, None, **T)], ["A", "B"])
    check("앞 공고 금액이 없을 때도", [mk("A", N, 0, None, **T), mk("B", N + "(2차)", 20, 60, **T)], ["A", "B"])
    check("금액 없는 재공고 '(2차)'", [mk("A", B, 0, 100), mk("B", B + "(2차)", 20, None, **RB)], ["A", "B"])
    assert run([mk("A", B, 0, 100), mk("B", B + "(2차)", 20, 100.5)])[1] == ["B"]                    # +0.5% = 재공고
    assert run([mk("A", B, 0, 100), mk("B", B + "(2차)", 20, 105)])[1] == ["A", "B"]                 # +5% = 차수 발주
    check("금액이 3% 다른 '(2차)'", [mk("A", N, 0, 50, **T), mk("B", N + "(2차)", 20, 51.5, **T)], ["A", "B"])
    check("같은 금액이라도 200일 뒤(재공고 표시 없음)면 차수 발주", [mk("A", N, 0, 50, **T), mk("B", N + "(2차)", 200, 50, **T)], ["A", "B"])
    check("재공고라고 밝혔으면 200일 뒤라도 재공고", [mk("A", N, 0, 50, **T), mk("B", N + "(2차)", 200, 50, **T, **RB)], ["B"])
    check("'', '(2차)', '(3차)' 같은 금액 20일 간격 → 마지막만", [mk("A", B, 0, 100), mk("B", B + "(2차)", 20, 100), mk("C", B + "(3차)", 40, 100)], ["C"])
    check("'', (2차) 재공고, (2차) 재공고, (3차) 재공고 → 마지막만",
          [mk("A", B, 0, 100), mk("B", B + "(2차)", 15, 100, **RB), mk("B2", B + "(2차)", 30, 100, **RB), mk("C", B + "(3차)", 45, 100, **RB)], ["C"])
    check("(1차), (1차) 재공고, 같은 금액 (2차) → 1차 재공고 + 2차",
          [mk("A", N + "(1차)", 0, 50, **T), mk("A2", N + "(1차)", 10, 50, **T, **RB), mk("B", N + "(2차)", 200, 50, **T)], ["A2", "B"])
    check("(1차), (2차), (2차) 재공고, 금액이 같은 (3차) → 차수 셋",
          [mk("A", N + "(1차)", 0, 50, **T), mk("B", N + "(2차)", 100, 60, **T), mk("B2", N + "(2차)", 120, 60, **T, **RB), mk("C", N + "(3차)", 250, 60, **T)], ["A", "B2", "C"])
    h = check("표기 없는 1차와 금액이 같은 (2차) + 그 재공고 → 1차 + 2차 재공고", [mk("A", B, 0, 100), mk("B", B + "(2차)", 60, 100), mk("C", B + "(2차)", 91, 100, **RB)], ["A", "C"])
    assert h.at["B-000", "대체공고번호"] == "C"
    check("'(2차, 긴급)' 재공고 같은 금액", [mk("A", B, 0, 100), mk("B", B + "(2차, 긴급)", 20, 100, **RB)], ["B"])
    check("'(2차, 긴급)' 금액이 다르면 차수 발주", [mk("A", B, 0, 100), mk("B", B + "(2차, 긴급)", 20, 60, **RB)], ["A", "B"])
    check("'(2차분, 긴급)' 은 금액이 같아도 차수 발주", [mk("A", B, 0, 100), mk("B", B + "(2차분, 긴급)", 20, 100)], ["A", "B"])

    # ② 'N차 재공고' 의 여러 표기: 그 차수의 원 공고('(2차)')가 먼저 나 있으면 그 공고의 재공고(1차는 남는다)
    for form in ("(2차 재공고)", " 2차 재공고", " 재공고(2차)"):
        check("N차 재공고 표기 " + form, [mk("A", N, 0, 50, **T), mk("B", N + "(2차)", 150, 60, **T), mk("C", N + form, 170, 60, **T, **RB)], ["A", "C"])
    check("1차와 금액이 같은 (2차) + (2차 재공고)", [mk("A", N, 0, 60, **T), mk("B", N + "(2차)", 150, 60, **T), mk("C", N + "(2차 재공고)", 170, 60, **T, **RB)], ["A", "C"])
    check("다른 공종의 '(2차)' 는 원 공고가 아니다", [mk("A", N, 0, 50, **T), mk("E", N + "(2차)", 5, 5, trade="전기"), mk("B", N + "(2차 재공고)", 20, 55, **T, **RB)], ["B", "E"])
    check("'(2차 재공고)' 뒤에 난 '(2차)' 는 그 원 공고가 아니다", [mk("A", N, 0, 50, **T), mk("B", N + "(2차 재공고)", 20, 50, **T, **RB), mk("C", N + "(2차)", 400, 120, **T)], ["B", "C"])
    check("'재공고(2차분)'", [mk("A", N, 0, 50, **T), mk("B", N + "(2차분)", 150, 60, **T), mk("C", N + " 재공고(2차분)", 170, 60, **T, **RB)], ["A", "C"])

    # ③ 여러 차수를 묶은 괄호도 차수 표기 — 지워지면 앞 공고의 재공고로 읽힌다. 그 본공사의 다른 공종 공고는 같은 세부사업
    check("'(2,3차분)'", [mk("A", N, 0, 50, **T), mk("B", N + "(2,3차분)", 200, 80, **T)], ["A", "B"], 1)
    check("'(2차 및 3차)'", [mk("A", N, 0, 50, **T), mk("B", N + "(2차 및 3차)", 200, 80, **T)], ["A", "B"], 1)
    check("'(2,3차분)' 본공사와 그 전기 공고", [mk("A", "가상시 회관 별관 증축공사(2,3차분)", 0, 50), mk("B", "가상시 회관 별관 증축 전기공사", 30, 5, trade="전기")], ["A", "B"], 1)
    check("'(2,3차분)' 재공고", [mk("A", N + "(2,3차분)", 0, 80, **T), mk("B", N + "(2,3차분)", 20, 80, **T, **RB)], ["B"])

    # ④ '부지' 규칙은 지구 전체의 부지조성 공고에만: 시설 이름에 '부지' 가 있거나, 건물 공사를 함께 발주했거나, 그 공고 이름을 딴 부대공사는 한 세부사업
    fb = "가상 반환부지 문화공원"
    check("시설명에 '부지'", [mk("A", fb + " 조성사업", 0, 50, fac=fb), mk("B", fb + " 연결통로 공사", 60, 3, trade="토목", fac=fb)], ["A", "B"], 1)
    check("'부지조성 및 건립공사' 본공사", [mk("A", "가상시 회관 부지조성 및 건립공사", 0, 100), mk("B", "가상시 회관 연결통로 공사", 60, 3, trade="토목")], ["A", "B"], 1)
    check("부지조성공사와 그 이름을 딴 부대공사", [mk("A", "가상시 회관 부지조성공사", 0, 30, **T), mk("B", "가상시 회관 부지조성공사 방음벽 설치공사", 60, 2, trade="기타")], ["A", "B"], 1)
    check("지구 부지조성과 무관한 소규모 공사는 다른 세부사업",
          [mk("A", "가상지구 부지조성공사", 0, 300, fac="가상지구", **T), mk("B", "가상지구 재선충병 방제사업", 60, 2, trade="기타", fac="가상지구")], ["A", "B"], 2)
    assert "놀이전수관" in works.features("가상 백중놀이전수관 건립공사", "가상 문화마을")["disc"]              # 낱말 속 '이전' 은 지우지 않는다
    assert "기업이전단지" in works.features("가상 기업이전단지 조성공사", "가상시")["disc"]

    # ⑤ 연례 공사(재공고 간격 180일)는 해마다 다시 발주하는 공사 이름일 때만: 연례 공사 말 뒤에 건물 공사 말이 오거나 낱말 속 글자면 한 번짜리 공사
    for nm in ("유지관리 사무실 신축공사", "유지관리용 창고 신축공사", "공연간판 설치공사", "2단가변무대 설치공사"):
        assert not works.is_annual("가상시 회관 " + nm), nm
        check("한 번짜리 공사 200일 뒤 재발주 → 뒤 공고만: " + nm, [mk("A", "가상시 회관 " + nm, 0, 3, trade="기타"), mk("B", "가상시 회관 " + nm, 200, 3, trade="기타")], ["B"])
    for nm in ("시설물 유지보수공사", "시설물 유지·보수공사", "연간 보수공사", "(단가) 보수공사", "시설보수단가공사", "신축 청사 유지관리공사", "시설물연간유지공사"):
        assert works.is_annual("가상시 회관 " + nm), nm
        h = check("연례 공사 200일 뒤 = 다음 회차: " + nm, [mk("A", "가상시 회관 " + nm, 0, 1, wt="리모델링"), mk("B", "가상시 회관 " + nm, 200, 1, wt="리모델링")], ["A", "B"])
        assert h.at["A-000", "프로젝트ID"] != h.at["B-000", "프로젝트ID"], nm
    check("이름 속 '신축'(쌍신축구장)", [mk("A", "쌍신축구장 천연잔디 유지관리공사", 0, 0.45, wt="리모델링", fac="쌍신축구장"),
                                      mk("B", "쌍신축구장 천연잔디 유지관리공사", 270, 0.45, wt="리모델링", fac="쌍신축구장")], ["A", "B"])

    # ⑥ 제안 안내 공고: 공법·자재 선정용 제안서 안내·공모만. 제작·설치 협상계약, 기술제안 입찰, 이름 속 글자('기자재', '신기술창업센터')는 공사 발주
    for nm, want in (("가상시 회관 전시물 및 관급자재 제작·설치 제안서 제출 안내 공고", False), ("가상시 회관 무대 기자재 구매 설치 제안서 제출 안내공고", False),
                     ("가상 신기술창업센터 전시물 제작·설치 제안서 제출 안내 공고", False), ("가상시 회관 건립공사(신기술 적용) 실시설계 기술제안 공고", False),
                     ("가상시 예술의전당 건립사업 실시설계 기술제안 입찰공고", False), ("가상시 회관 특허공법 적용 방수공사", False),
                     ("가상시 회관 친환경 자재 공모전 전시관 건립공사", False),
                     ("가상시 회관 특허공법 선정 공고", True), ("가상시 회관 공법 선정 제안서 접수 공고", True), ("가상시 회관 신기술·특허공법 공모", True),
                     ("가상시 회관 관급자재(방수) 선정 기술제안서 제출 안내 공고", True), ("가상 주차타워 건립사업 장경간 구조 신기술 제안서 제출 안내 공고", True),
                     ("「가상읍 공영주차타워 건립사업」건축구조(주차타워) 공법, 자재 제안서 제출안내", True), ("가상군 자주식 주차타워 조성사업 건축구조(주차장)공법제안 공고", True)):
        assert works.is_proposal(nm) is want, nm

    # ⑦ 표시용 이름: 지운 '전체분'·'총괄' 뒤에 남는 '및'
    assert works.clean_title("가상시 회관 건립공사(소방)(장기계속, 전체분 및 1차분)") == "가상시 회관 건립공사(소방)(장기계속, 1차분)", works.clean_title("가상시 회관 건립공사(소방)(장기계속, 전체분 및 1차분)")
    assert works.clean_title("가상군 청년임대주택 건립공사(총괄및1차분)") == "가상군 청년임대주택 건립공사(1차분)", works.clean_title("가상군 청년임대주택 건립공사(총괄및1차분)")
    assert works.clean_title("가상시 회관 건립공사(전체분 및 1차분)") == "가상시 회관 건립공사(1차분)"

    # ⑧ 이전공고번호 사슬: 유효한 공고에서 멈추고(가운데 공고도 대체), 취소된 공고는 건너뛰며, 대체공고번호는 끝까지 남은 공고
    h = check("A ← B ← C 모두 유효, 이름이 서로 다름", [mk("A", B, 0, 100), mk("B", "가상시 문화회관 건립공사(건축)", 31, 180, prev="A"),
                                                       mk("C", "가상시 복합문화회관 신축공사", 60, 290, prev="B")], ["C"])
    assert h.at["A-000", "대체공고번호"] == h.at["B-000", "대체공고번호"] == "C", h["대체공고번호"].to_dict()
    h = check("A ← B[취소] ← B2[취소] ← C", [mk("A", B, 0, 100), mk("B", B, 20, 100, kind="취소", prev="A"), mk("B2", B, 31, 100, kind="취소", prev="B"),
                                              mk("C", "가상시 문화회관 건립공사(건축) 새로운 입찰", 40, 180, prev="B2", **RB)], ["C"])
    assert h.at["A-000", "대체공고번호"] == "C"
    h = check("A ← B[취소] ← C: 이름이 달라도 한 세부사업", [mk("A", B, 0, 50), mk("B", B + "(건축)", 31, 52, prev="A", kind="취소"),
                                                           mk("C", "나래마을 문화센터 신축 건축공사", 60, 55, prev="B")], ["C"])
    assert h.at["A-000", "프로젝트ID"] == h.at["C-000", "프로젝트ID"] == "F1-N" and "재공고로 1건 대체" in h.at["C-000", "대표선정사유"], h[["프로젝트ID", "대표선정사유"]].to_dict()
    h = check("사업유형이 다른 사슬 A(신축) ← B(리모델링) ← C(리모델링 재공고)", [mk("A", B, 0, 100), mk("B", "가상시 회관 리모델링공사", 31, 100, prev="A", wt="리모델링"),
                                                                             mk("C", "가상시 회관 리모델링공사", 60, 100, prev="B", wt="리모델링", **RB)], ["C"])
    assert h.at["A-000", "대체공고번호"] == "C" and h.at["B-000", "대체공고번호"] == "C", h["대체공고번호"].to_dict()      # B 를 거쳐 끝까지 남은 C
    h = check("사업유형이 다른 두 공고가 서로를 가리키면 늦게 난 공고가 남는다",
              [mk("A", B, 0, 100, prev="B"), mk("B", "가상시 회관 리모델링공사", 31, 100, prev="A", wt="리모델링")], ["B"])
    assert h.at["A-000", "대체공고번호"] == "B" and h.at["B-000", "대체공고번호"] == ""
    # 다른 프로젝트의 재공고에 대체돼 대표에서 빠지면, 같은 프로젝트·공종에 남은 공고의 사유도 남은 건수에 맞춘다
    h = check("연결로 대표가 빠진 뒤 남은 공고의 사유", [mk("Y", "가상시 회관 건립 전기공사", 0, 5, trade="전기"), mk("Z", "가상시 회관 건립 무대조명 전기공사", 1, 2, trade="전기"),
                                                       mk("R", "가상시 회관 리모델링 전기공사", 40, 5, trade="전기", wt="리모델링", prev="Y")], ["R", "Z"])
    assert h.at["Z-000", "대표선정사유"].startswith("단독 공고") and h.at["Y-000", "대표선정사유"] == "대체됨(API 이전공고번호 연결)", h["대표선정사유"].to_dict()
    # 한 공고의 차수마다 이전공고번호가 다르면 가장 늦은 차수의 값을 쓴다(입력 행 순서와 무관). 보이는 열은 API 가 준 값 그대로
    rows = [mk("A", B, 0, 50), mk("X", "가상시 회관 별관 증축공사", 5, 30, wt="증축"), mk("B", "가상시 회관 리모델링공사", 31, 52, prev="A", wt="리모델링"),
            mk("B", "가상시 회관 리모델링공사", 33, 52, prev="X", wt="리모델링", ord_="001", kind="변경")]
    for order in ([0, 1, 2, 3], [0, 1, 3, 2], [3, 2, 1, 0]):
        h, cnt, _ = run([dict(rows[k]) for k in order])
        assert cnt == ["A", "B"] and h.at["X-000", "대체공고번호"] == "B" and h.at["B-000", "이전공고번호"] == "A" and h.at["B-001", "이전공고번호"] == "X", (order, cnt, h["이전공고번호"].to_dict())
    # 공고번호가 숫자로 읽혀 '.0' 이 붙어도 이전공고번호 연결을 세부사업 배정에서 쓴다(이름이 다른 재공고가 원 공고와 같은 세부사업)
    for a, b, pv in ((20240100001.0, 20240100002.0, 20240100001.0), ("20240100001.0", "20240100002.0", "20240100001.0"), (20240100001, 20240100002, 20240100001.0)):
        h, l, _ = dedup.dedup_latest(pd.DataFrame([mk(a, B, 0, 50, prev=("" if isinstance(pv, str) else float("nan"))), mk(b, "나래마을 문화센터 신축 건축공사", 31, 50, prev=pv)]), 0.30)
        assert len(l) == 1 and dedup._no(l["공고번호"].iloc[0]) == "20240100002" and set(h["프로젝트ID"]) == {"F1-N"}, (a, list(l["공고번호"]), h["프로젝트ID"].to_dict())

    # ⑨ 취소공고의 대체공고번호: 같은 세부사업·공종에서 이름이 가장 비슷한 대표 공고 — 단, 취소공고보다 먼저 난 공고는 가리키지 않는다(이전공고번호로 이어진 원 공고는 예외)
    h = check("취소 → 뒤에 다시 난 공고", [mk("X", "가상시 회관 건립 전기공사", 0, 5, trade="전기", kind="취소"), mk("Y", "가상시 회관 건립 전기공사", 31, 5, trade="전기"),
                                         mk("Z", "가상시 회관 건립 무대조명 전기공사", 32, 2, trade="전기")], ["Y", "Z"])
    assert h.at["X-000", "대체공고번호"] == "Y" and h.at["X-000", "대표선정사유"] == "취소공고", h["대체공고번호"].to_dict()
    h = check("취소된 재공고 → 원 공고가 그대로 유효", [mk("A", B, 0, 100), mk("X", B, 31, 100, kind="취소", prev="A")], ["A"])
    assert h.at["X-000", "대체공고번호"] == "A"
    fac = "가상공사 온실"
    rows = [mk("A", "가상군 스마트농업 실습농장 조성사업 온실공사", "2025-08-06 10:00:00", 60, fac=fac), mk("P", "가상 2호 스마트팜 온실신축사업 건축공사", "2023-07-04 10:00:00", 11, fac=fac),
            mk("X", "다른군 스마트농업 육성지구 조성사업(서부단지) 온실공사", "2026-04-28 10:00:00", 180, fac=fac, kind="취소")]
    hist, latest, _ = dedup.dedup_latest(pd.DataFrame(rows), 0.30)
    hk = hist.set_index("공고키")
    assert hk.at["X-000", "대체공고번호"] == "" and hk.at["X-000", "대표선정사유"] == "취소공고", hk[["프로젝트ID", "대체공고번호"]].to_dict()      # 먼저 난 다른 사업의 공고를 가리키지 않는다
    # 05 '공고연도'의 근거인 최종공고일은 합산 대상 공고 기준(유효 차수가 없는 취소공고가 붙어 있어도 연도가 밀리지 않는다)
    fs = dedup.facility_summary(latest, hist).set_index("프로젝트ID")
    assert str(fs.at[hk.at["A-000", "프로젝트ID"], "최종공고일"]).startswith("2025-08-06") and str(fs.at[hk.at["X-000", "프로젝트ID"], "최종공고일"])[:4] in ("2025", "2023"), fs["최종공고일"].to_dict()

    # ⑩ 시설 간 중복: 같은 사업이고 둘 다 나뉘지 않았으면 프로젝트ID 가 앞선 행을 남긴다(입력 순서와 무관) / 서로의 세부사업에 나뉜 경우 / 합친 것과 같은 경우
    rows = [mk("A1", "가상 1단지 아파트 건설공사", 0, 900, fac="가상 1단지 아파트", fid="F2"), mk("A1", "가상 1단지 아파트 건설공사", 0, 900, fac="가상 아파트", fid="F7")]
    _, latest, _ = dedup.dedup_latest(pd.DataFrame(rows), 0.30)
    dup = dedup.cross_facility_duplicates(latest)
    assert dup.at["F2-N", "중복_제외권장"] == "" and dup.at["F7-N", "중복_제외권장"] == "Y" and dup.at["F7-N", "시설간_중복"] == "같은 사업: F2-N", dup.to_dict()
    assert dedup.cross_facility_duplicates(latest.iloc[::-1]).sort_index().equals(dup.sort_index())
    sets_ = pd.DataFrame([dict(프로젝트ID=p_, 시설ID=f_, 공고번호=n_, 세부사업수=2) for p_, f_, nos in (("A-N-01", "A", "12"), ("A-N-02", "A", "34"), ("B-N-01", "B", "13"), ("B-N-02", "B", "24")) for n_ in nos])
    dup = dedup.cross_facility_duplicates(sets_)
    assert list(dup["중복_제외권장"]) == ["", "", "Y", "Y"] and dup.at["B-N-01", "시설간_중복"].startswith("A-N-01 + A-N-02 에 나뉘어 있음"), dup.to_dict()
    sets_ = pd.DataFrame([dict(프로젝트ID=p_, 시설ID=f_, 공고번호=n_, 세부사업수=k_) for p_, f_, k_, nos in (("A-N", "A", 1, "123"), ("B-N-01", "B", 3, "1"), ("B-N-02", "B", 3, "2"), ("B-N-03", "B", 3, "3")) for n_ in nos])
    dup = dedup.cross_facility_duplicates(sets_)
    assert dup.at["A-N", "중복_제외권장"] == "" and list(dup.loc[["B-N-01", "B-N-02", "B-N-03"], "중복_제외권장"]) == ["Y", "Y", "Y"], dup.to_dict()

    # ⑪ dedup 단계 요약: '복수 시설 배정' 건수(공고번호 기준 — 시설마다 다른 차수가 대표여도 한 건)와 '시설 간 중복 프로젝트' 1줄.
    #    Y 행을 빼도 두 번 남는 공고(일부만 겹치는 프로젝트)가 있으면 '두 번 합산되지 않음' 대신 남은 건수를 적는다
    def stage(rows, name):
        d_ = os.path.join(tmp, name)
        os.makedirs(d_, exist_ok=True)
        pd.DataFrame(rows).to_parquet(os.path.join(d_, "notices_research.parquet"), index=False)
        buf = io.StringIO()
        with contextlib.redirect_stdout(buf):
            pipeline.stage_dedup({"paths": {"data_dir": d_}, "dedup": {"price_change_warn_ratio": 0.30}})
        lg = pd.read_parquet(os.path.join(d_, "logs_dedup.parquet"))
        return buf.getvalue(), lg

    own, umb = "가상 1단지 아파트", "가상공사 공공주택"
    out, lg = stage([mk("A1", "가상 1단지 아파트 건설공사", 0, 900, fac=umb, fid="F9"), mk("A2", "가상 1단지 아파트 전기공사", 31, 90, fac=umb, fid="F9", trade="전기"),
                     mk("B1", "가상 2단지 아파트 건설공사", 0, 800, fac=umb, fid="F9"),
                     mk("A1", "가상 1단지 아파트 건설공사", 0, 900, fac=own, fid="F2"), mk("A2", "가상 1단지 아파트 전기공사", 31, 90, fac=own, fid="F2", trade="전기"),
                     mk("B1", "가상 2단지 아파트 건설공사", 0, 800, fac="가상 2단지 아파트", fid="F3"), mk("B9", "가상 2단지 아파트 전기공사", 31, 80, fac="가상 2단지 아파트", fid="F3", trade="전기")], "dup_clean")
    row = lg[lg["항목"] == "시설 간 중복 프로젝트"]
    assert "복수 시설 배정 3건" in out and int((lg["항목"] == "복수 시설 배정").sum()) == 3, out
    assert (lg.loc[lg["항목"] == "복수 시설 배정", "판정"] == "참고").all(), "06 시트에서 판정 칸이 비지 않게"
    assert len(row) == 1 and row["판정"].iloc[0] == "참고" and "프로젝트 4개" in row["비고"].iloc[0] and "Y 인 2개 행을 빼면 같은 공고가 두 번 합산되지 않음" in row["비고"].iloc[0], row["비고"].tolist()
    assert "시설 간 중복 프로젝트 4개" in out and "Y 인 2개 행을 빼면 같은 공고가 두 번 합산되지 않음" in out, out
    umb = "가상도시공사 공공주택"                       # 묶음 시설 F9 는 단지별로 나뉘고, 단지 시설 F2·F3 이 진입도로 공고 R 을 함께 가진 경우 → Y 행을 빼도 R 은 두 번 남는다
    part = [mk("A1", "가상 1단지 아파트 건설공사", 0, 900, fac=umb, fid="F9"), mk("A2", "가상 1단지 아파트 전기공사", 31, 90, fac=umb, fid="F9", trade="전기"),
            mk("R", "가상지구 진입도로 개설공사", 60, 50, fac=umb, fid="F9", trade="토목"),
            mk("B1", "가상 2단지 아파트 건설공사", 0, 800, fac=umb, fid="F9"), mk("B2", "가상 2단지 아파트 전기공사", 31, 80, fac=umb, fid="F9", trade="전기"),
            mk("A1", "가상 1단지 아파트 건설공사", 0, 900, fac="가상 1단지 아파트", fid="F2"), mk("A2", "가상 1단지 아파트 전기공사", 31, 90, fac="가상 1단지 아파트", fid="F2", trade="전기"),
            mk("R", "가상지구 진입도로 개설공사", 60, 50, fac="가상 1단지 아파트", fid="F2", trade="토목"),
            mk("B1", "가상 2단지 아파트 건설공사", 0, 800, fac="가상 2단지 아파트", fid="F3"), mk("B2", "가상 2단지 아파트 전기공사", 31, 80, fac="가상 2단지 아파트", fid="F3", trade="전기"),
            mk("R", "가상지구 진입도로 개설공사", 60, 50, fac="가상 2단지 아파트", fid="F3", trade="토목")]
    out, lg = stage(part, "dup_partial")
    note = lg.loc[lg["항목"] == "시설 간 중복 프로젝트", "비고"].iloc[0]
    assert "두 번 합산되는 공고가 1건 남음" in note and "두 번 합산되지 않음" not in note and "두 번 합산되는 공고가 1건 남음" in out, (note, out)
    # 같은 공고가 한 시설에서는 000 차수, 다른 시설에서는 001 차수로 대표가 된 경우도 '복수 시설 배정' 한 건
    out, lg = stage([mk("A1", "가상 1단지 아파트 건설공사", 0, 900, fac=own, fid="F2"), mk("A1", "가상 1단지 아파트 건설공사", 0, 900, fac="가상 아파트", fid="F7"),
                     mk("A1", "가상 1단지 아파트 건설공사(정정)", 5, 900, fac="가상 아파트", fid="F7", ord_="001", kind="변경")], "dup_rev")
    assert "복수 시설 배정 1건" in out and int((lg["항목"] == "복수 시설 배정").sum()) == 1, out


def test_work_split_review3(tmp):
    """세부사업 분리 3차 독립 검토(2026-11-16)에서 나온 결함·빈틈: 공고 회차 판정(round_ord)의 범위, 금액이 다른 '(2차)' 를 사이에 둔 '(3차)',
    '제N차 재공고' 표기, 지구 이름 속 '신설', '및' 으로 이어진 연례 공사, 제안 안내 공고의 '입찰' 낱말, 취소공고의 대체공고번호(변경 차수),
    '같은 공종 합산' 로그 갱신, 자기 번호를 가리키는 이전공고번호, 공고일시 없는 대표 공고의 공고연도, 요약 줄의 고유 공고 수 — 그리고 테스트가 지키지 못하던 가드들."""
    import io
    import contextlib
    from g2b_costdb import works, dedup, pipeline

    t0 = pd.Timestamp("2024-01-10 10:00:00")

    def mk(no, name, day, eok, kind="일반", trade="건축", ord_="000", prev="", wt="신축", fac="가상시 회관", rebid="N", fid="F1"):
        dt = day if (day is None or isinstance(day, str)) else (t0 + pd.Timedelta(days=day)).strftime("%Y-%m-%d %H:%M:%S")
        return dict(시설ID=fid, 시설명=fac, 공고번호=no, 공고차수=ord_, 공고키=f"{no}-{ord_}", 공고명=name, 공고일시=dt, 사업유형=wt, 공종=trade,
                    추정가격=(None if eok is None else eok * 1e8), 공고종류=kind, 재공고여부=rebid, 이전공고번호=prev,
                    표2_대분류="문화 및 집회시설", 표2_중분류="-", 검색어="회관", 수요기관="가상시")

    def run(rows):
        h, l, lg = dedup.dedup_latest(pd.DataFrame(rows), 0.30)
        return h.set_index("공고키"), sorted(str(x).strip() for x in l["공고번호"]), lg

    def check(tag, rows, want, projects=None):
        h, cnt, _ = run(rows)
        assert cnt == sorted(want), (tag, cnt, h["대표선정사유"].to_dict())
        if projects is not None:
            assert h["프로젝트ID"].nunique() == projects, (tag, h["프로젝트ID"].to_dict())
        return h

    def stage(rows, name):
        d_ = os.path.join(tmp, name)
        os.makedirs(d_, exist_ok=True)
        pd.DataFrame(rows).to_parquet(os.path.join(d_, "notices_research.parquet"), index=False)
        buf = io.StringIO()
        with contextlib.redirect_stdout(buf):
            pipeline.stage_dedup({"paths": {"data_dir": d_}, "dedup": {"price_change_warn_ratio": 0.30}})
        return buf.getvalue(), pd.read_parquet(os.path.join(d_, "logs_dedup.parquet"))

    RB, T, EL, RM = dict(kind="재공고", rebid="Y"), dict(trade="토목"), dict(trade="전기"), dict(wt="리모델링")
    N, B, BS, R = "가상시 회관 조성사업", "가상시 회관 건립공사", "가상시 회관 건립사업", "가상시 회관 옥상 방수공사"

    # ① 괄호 'N차' 가 공고 회차(몇 번째 공고)인지 — 재공고라고 밝힌 '(2차)' 는 같은 이름의 앞 공고가 없으면 공고 회차
    h = check("'', (2차) 재공고, (2차) 재공고 → 마지막만(첫 '(2차)' 뒤에 같은 이름으로 또 난 공고는 그 차수의 원 공고가 아니다)",
              [mk("A", B, 0, 100), mk("B", B + "(2차)", 15, 100, **RB), mk("B2", B + "(2차)", 30, 100, **RB)], ["B2"])
    assert h.at["A-000", "대체공고번호"] == "B2" and h.at["B-000", "대체공고번호"] == "B2", h["대체공고번호"].to_dict()
    check("'', (2차) 표시 없음, (2차) 재공고, 같은 금액 (3차) → 마지막만", [mk("A", B, 0, 100), mk("B", B + "(2차)", 20, 100), mk("B2", B + "(2차)", 40, 100, **RB), mk("C", B + "(3차)", 60, 100)], ["C"])
    check("(2차) 재공고, (3차) 재공고 — 첫 공고 없음", [mk("B", B + "(2차)", 15, 100, **RB), mk("C", B + "(3차)", 30, 100, **RB)], ["C"])
    check("(2차) 재공고 ×2, (3차) 재공고 — 첫 공고 없음", [mk("B", B + "(2차)", 15, 100, **RB), mk("B2", B + "(2차)", 30, 100, **RB), mk("C", B + "(3차)", 45, 100, **RB)], ["C"])
    check("(2차) 재공고, (2차) 재공고, 같은 금액 (3차)", [mk("B", B + "(2차)", 0, 100, **RB), mk("B2", B + "(2차)", 20, 100, **RB), mk("C", B + "(3차)", 60, 100)], ["C"])
    check("(2차) 재공고, 같은 이름 (2차), 같은 금액 (3차)", [mk("B", B + "(2차)", 0, 100, **RB), mk("B2", B + "(2차)", 20, 100), mk("C", B + "(3차)", 40, 100)], ["C"])
    # 같은 이름의 앞 공고가 있으면 그 차수 공고의 재공고(차수 표기) — 뒤 공고를 120일 안에 두어 120일 제한이 아니라 이 가드가 판정하게 한다
    check("(1차), (1차) 재공고, 90일 뒤 같은 금액 (2차) → 1차 재공고 + 2차",
          [mk("A", N + "(1차)", 0, 50, **T), mk("A2", N + "(1차)", 10, 50, **T, **RB), mk("B", N + "(2차)", 100, 50, **T)], ["A2", "B"])
    check("(1차), (2차), (2차) 재공고, 80일 뒤 금액이 같은 (3차) → 차수 셋",
          [mk("A", N + "(1차)", 0, 50, **T), mk("B", N + "(2차)", 100, 60, **T), mk("B2", N + "(2차)", 120, 60, **T, **RB), mk("C", N + "(3차)", 200, 60, **T)], ["A", "B2", "C"])
    check("다른 공종의 같은 이름 '(2차)' 는 앞 공고가 아니다", [mk("E", N + "(2차)", 0, 5, **EL), mk("B", N + "(2차)", 5, 50, **T, **RB), mk("C", N + "(3차)", 30, 50, **T)], ["C", "E"])
    check("앞의 같은 이름 공고가 모두 공고 회차면 금액이 바뀐 재공고도 공고 회차(꾸민 사례)",
          [mk("B", B + "(2차)", 0, 100, **RB), mk("B2", B + "(2차)", 20, 100, **RB), mk("B3", B + "(2차)", 40, 250, **RB), mk("C", B + "(3차)", 60, 250)], ["B2", "C"])
    # 다른 세부사업(지난 회차)의 같은 이름 '(2차)' 는 앞 공고로 치지 않는다
    check("400일 간격 두 회차, 회차마다 (2차) 재공고 + (3차) 재공고",
          [mk("B", R + "(2차)", 0, 1, **RM, **RB), mk("C", R + "(3차)", 15, 1, **RM, **RB), mk("B2", R + "(2차)", 400, 1, **RM, **RB), mk("C2", R + "(3차)", 415, 1, **RM, **RB)], ["C", "C2"])
    h = check("지난 회차는 차수 발주, 이번 회차는 (2차) 재공고 + (3차) 재공고",
              [mk("A", R, 0, 2, **RM), mk("B", R + "(2차)", 100, 1, **RM), mk("B2", R + "(2차)", 500, 1, **RM, **RB), mk("C2", R + "(3차)", 515, 1, **RM, **RB)], ["A", "B", "C2"])
    assert h.at["B2-000", "대체공고번호"] == "C2", h["대체공고번호"].to_dict()
    # 사이에 금액이 다른 '(2차)' 가 합산 대상으로 남아 있으면 '(3차)' 는 차수 발주(1차분과 금액이 같을 뿐)
    check("'', (2차) 60억, (3차) 50억 → 차수 셋", [mk("A", N, 0, 50, **T), mk("B", N + "(2차)", 50, 60, **T), mk("C", N + "(3차)", 100, 50, **T)], ["A", "B", "C"])
    check("같은 경우, (3차) 가 재공고", [mk("A", N, 0, 50, **T), mk("B", N + "(2차)", 50, 60, **T), mk("C0", N + "(3차)", 190, 50, kind="취소", **T),
                                    mk("C", N + "(3차)", 200, 50, **T, **RB)], ["A", "B", "C"])
    check("사이의 다른 금액 공고가 다른 공종이면 차수 발주의 근거가 아니다", [mk("A", N, 0, 50, **T), mk("E", N + "(3차)", 10, 5, **EL), mk("B", N + "(2차)", 20, 50, **T)], ["B", "E"])
    check("사이의 다른 금액 공고에 차수 표기가 없으면 차수 발주의 근거가 아니다", [mk("A", N, 0, 50, **T), mk("A2", N, 10, 150, **T), mk("B", N + "(2차)", 20, 50, **T)], ["A2", "B"])

    # ② 'N차 재공고' 의 '제' 유무가 차수 공고와 달라도 그 차수의 재공고
    for lot, form in (("(2차)", " 제2차 재공고"), ("(제2차)", "(2차 재공고)"), ("(제2차)", "(제2차 재공고)")):
        h = check(f"'{lot}' 의 재공고를 '{form.strip()}' 로 적은 경우", [mk("A", N, 0, 50, **T), mk("B", N + lot, 150, 60, **T), mk("C", N + form, 170, 60, **T, **RB)], ["A", "C"])
        assert h.at["B-000", "대체공고번호"] == "C", (lot, form, h["대체공고번호"].to_dict())

    # '(2차 재공고)' 는 그보다 늦게 난 '(2차)' 의 재공고가 아니다 — 첫 공고의 재공고이고, 400일 뒤의 '(2차)' 80억은 2차분(금액비가 재발주 범위 안이어도)
    h = check("'', '(2차 재공고)' 20일, 400일 뒤 '(2차)' 80억", [mk("A", N, 0, 50, **T), mk("B", N + "(2차 재공고)", 20, 50, **T, **RB), mk("C", N + "(2차)", 400, 80, **T)], ["B", "C"])
    assert h.at["A-000", "대체공고번호"] == "B", h["대체공고번호"].to_dict()

    # ③ 여러 차수를 묶은 괄호는 차수 표기로 읽는다(400일 뒤에 나도 같은 사업의 차수 발주)
    long_ = "가상시 회관 복합문화체육센터 주차장 조성사업"
    assert works.features(long_ + "(2,3차분)", "가상시 회관")["phases"] == frozenset({"23차분"}), works.features(long_ + "(2,3차분)", "가상시 회관")["phases"]
    check("'(2,3차분)' 400일 뒤", [mk("A", long_, 0, 50, **T), mk("B", long_ + "(2,3차분)", 400, 80, **T)], ["A", "B"], 1)

    # ④ '부지' 규칙: 시설(지구) 이름 속 '신설'·'건설' 은 건물 공사 말이 아니다 / '건립부지'·'건립 예정부지' 는 부지
    assert works.features("가상 신설동지구 부지조성공사", "가상 신설동지구")["site_prep"] is True
    assert works.features("가상지구 건립부지 조성공사", "가상지구")["site_prep"] is True
    assert works.features("가상지구 청사 건립 예정부지 정비공사", "가상지구")["site_prep"] is True
    assert works.features("가상시 회관 신축 및 부지조성공사", "가상시 회관")["site_prep"] is False        # 건물 공사를 함께 발주한 공고는 본공사
    for fac, a in (("가상 신설동지구", "가상 신설동지구 부지조성공사"), ("가상지구", "가상지구 건립부지 조성공사")):
        check(f"'{a}' 와 무관한 소규모 공사는 다른 세부사업", [mk("A", a, 0, 300, fac=fac, **T), mk("B", fac + " 재선충병 방제사업", 60, 2, trade="기타", fac=fac)], ["A", "B"], 2)

    # ⑤ 연례 공사: '연간단가'·'단가계약' 이 있거나 건물 공사 말이 '및' 으로 이어져 있으면 연례 공사 / 낱말 속 글자 '쌍신축구장' 은 건물 공사 말이 아니다
    for nm in ("가상시 회관 시설물 유지보수 및 신설공사(연간단가)", "가상시 회관 보안등 유지보수 및 신설공사", "가상시 회관 시설물연간정비공사",
               "가상시 회관 연간단가 시설물 보수 및 증축공사", "가상시 회관 시설물 유지보수공사(연간단가, 신설 포함)", "가상시 회관 시설물 보수 단가계약(신설 포함)",
               "가상시 회관 시설물 유지/보수공사", "가상시 회관 시설물 유지 및 보수공사",
               "가상시 회관 유지관리공사(쌍신축구장)", "가상시 회관 유지관리 쌍신축구장 잔디공사"):
        assert works.is_annual(nm), nm
    for nm in ("가상시 회관 유지관리 사무실 신축공사", "가상시 회관 유지관리용 창고 신축공사", "가상시 회관 연간회원 라운지 신축공사", "가상시 공연간판 설치공사",
               "가상시 국유지 및 관리동 증축공사", "가상시 회관 유지 및 관리동 신축공사"):
        assert not works.is_annual(nm), nm
    check("'유지보수 및 신설공사(연간단가)' 270일 뒤 = 다음 회차", [mk("A", "가상시 회관 시설물 유지보수 및 신설공사(연간단가)", 0, 1, **RM),
                                                             mk("B", "가상시 회관 시설물 유지보수 및 신설공사(연간단가)", 270, 1, **RM)], ["A", "B"], 2)

    # ⑥ 제안 안내 공고: 이름에 '입찰' 이 있다는 것만으로는 공사 발주가 아니다('입찰 아님', '전자입찰'). 실제 발주로 보는 것은 '기술제안 입찰'·'입찰공고'·제작(구매)·설치
    for nm in (BS + " 특허공법 선정 기술제안서 제출 안내 공고(입찰 아님)", BS + " 공법선정 기술제안서 제출 안내 공고(전자입찰)",
               BS + " 특허공법 선정 제안서 제출 안내 공고(입찰참가자격 등록 불필요)", BS + " 특허공법 선정 기술제안서 제출 안내 공고(입찰공고 아님)"):
        assert works.is_proposal(nm) is True, nm
    for nm, why in ((B + "(특허공법) 입찰공고", "공사 입찰공고"), (B + " 기술제안입찰 공고(특허공법 포함)", "기술제안입찰"),
                    (B + " 기술제안입찰 제안서 제출 안내 공고(특허공법 포함)", "'기술제안 입찰' 만으로 제외"),
                    ("가상시 회관 방수공법 선정 공고(입찰공고)", "'입찰공고' 만으로 제외"),
                    ("가상시 회관 무대 기자재 납품 제안서 제출 안내 공고", "'기자재' 는 자재 선정이 아님"),
                    ("가상 신기술창업센터 리모델링 제안서 제출 안내 공고", "'신기술창업센터' 는 신기술 선정이 아님"),
                    ("가상시 회관 전시 콘텐츠 제안서 제출 안내 공고", "무엇을 제안받는지(공법·자재·신기술·특허)가 없음"),
                    ("가상시 회관 관급자재 제작구매 제안서 제출 안내 공고", "'제작구매'"), ("가상시 회관 관급자재 구매 설치 제안서 제출 안내 공고", "'구매 설치'"),
                    (BS + " 실험 기자재 선정 공고", "'기자재 선정 공고' 는 구매 발주")):
        assert works.is_proposal(nm) is False, (nm, why)
    check("'입찰 아님' 이라고 적은 제안 안내 공고는 합산하지 않는다", [mk("A", BS + " 특허공법 선정 기술제안서 제출 안내 공고(입찰 아님)", 0, 4), mk("B", B, 120, 100)], ["B"])

    # ⑦ 대체 사유가 대표선정사유에 그대로 실린다 / 유효 차수가 없는 취소공고는 같은 이름으로 다시 난 공고의 세부사업을 따른다
    h, _, _ = run([mk("A", B, 0, 100), mk("B", B, 20, 100)])
    assert h.at["A-000", "대표선정사유"] == "대체됨(같은 공고명으로 다시 공고)", h["대표선정사유"].to_dict()
    h, _, _ = run([mk("A", B, "2022-01-10 10:00:00", 100), mk("B", B, "2023-06-10 10:00:00", 120)])
    assert h.at["A-000", "대표선정사유"] == "대체됨(같은 공고명으로 다시 발주(사업 지연 뒤 재발주))", h["대표선정사유"].to_dict()
    h, _, _ = run([mk("A", "가상시 회관 건립공사(건축,기계)", 0, 100), mk("B", "가상시 회관 건립공사(건축, 기계설비)", 20, 100)])
    assert h.at["A-000", "대표선정사유"] == "대체됨(거의 같은 공고명·금액으로 다시 공고)", h["대표선정사유"].to_dict()
    h, cnt, _ = run([mk("X", "가상시 회관 스프링클러 추가설치 소방공사", "2024-11-04 10:00:00", 2.47, trade="소방", kind="취소", **RM),
                     mk("N", "가상시 회관 스프링클러 추가 설치 소방공사", "2024-11-08 10:00:00", 2.47, trade="소방", **RM),
                     mk("M", "가상시 회관 스프링클러 추가설치 소방공사 입찰공고", "2025-07-23 10:00:00", 0.9, trade="소방", **RM)])
    assert cnt == ["M", "N"] and h.at["N-000", "프로젝트ID"] != h.at["M-000", "프로젝트ID"], (cnt, h["프로젝트ID"].to_dict())     # 전제: 두 계약은 다른 세부사업
    assert h.at["X-000", "프로젝트ID"] == h.at["N-000", "프로젝트ID"], h["프로젝트ID"].to_dict()

    # ⑧ 취소공고의 대체공고번호: '뒤에 난 공고' 는 두 공고 모두 처음 난 날로 견준다
    ELN = "가상시 회관 건립 전기공사"
    h = check("취소 차수(001)는 늦게 났지만 공고 자체는 Y 보다 먼저 났다 → Y", [mk("X", ELN, 0, 5, **EL), mk("X", ELN, 50, 5, ord_="001", kind="취소", **EL), mk("Y", ELN, 20, 5, **EL)], ["Y"])
    assert h.at["X-001", "대체공고번호"] == "Y", h["대체공고번호"].to_dict()
    h = check("취소공고와 같은 시각에 난 공고도 가리킨다", [mk("X", ELN, 0, 5, kind="취소", **EL), mk("Y", ELN, 0, 5, **EL)], ["Y"])
    assert h.at["X-000", "대체공고번호"] == "Y", h["대체공고번호"].to_dict()
    h = check("먼저 난 공고의 변경 차수가 취소공고보다 늦어도 '뒤에 난 공고' 가 아니다",
              [mk("A", B, "2024-01-10 10:00:00", 50), mk("A", B, "2024-07-10 10:00:00", 50, ord_="001", kind="변경"), mk("C", B, "2024-05-10 10:00:00", 50, kind="취소")], ["A"])
    assert h.at["C-000", "대체공고번호"] == "", h["대체공고번호"].to_dict()
    h = check("취소공고가 가리키던 공고가 다른 사업유형의 재공고에 대체되면 끝까지 남은 공고를 가리킨다",
              [mk("X", ELN, 0, 5, kind="취소", **EL), mk("Y", ELN, 31, 5, **EL), mk("R", "가상시 회관 리모델링 전기공사", 60, 5, prev="Y", **EL, **RM)], ["R"])
    assert h.at["X-000", "대체공고번호"] == "R" and h.at["Y-000", "대체공고번호"] == "R", h["대체공고번호"].to_dict()
    h = check("사업유형이 모두 다른 사슬 A ← B ← C ← D", [mk("A", B, 0, 100), mk("B", "가상시 회관 리모델링공사", 31, 100, prev="A", **RM),
                                                    mk("C", "가상시 회관 별관 증축공사", 60, 100, prev="B", wt="증축"), mk("D", "가상시 문화센터 대수선공사", 90, 100, prev="C", **RM)], ["D"])
    assert h.at["A-000", "대체공고번호"] == h.at["B-000", "대체공고번호"] == h.at["C-000", "대체공고번호"] == "D", h["대체공고번호"].to_dict()
    # 이전공고번호: 빈 차수는 다른 차수의 값으로 채우고, 자기 번호를 가리키는 값은 연결로 쓰지 않는다(입력 행 순서와 무관)
    h, cnt, _ = run([mk("A", B, 0, 50), mk("B", "가상시 회관 리모델링공사", 31, 52, **RM), mk("B", "가상시 회관 리모델링공사", 33, 52, prev="A", ord_="001", kind="변경", **RM)])
    assert cnt == ["B"] and h.at["B-000", "이전공고번호"] == "A" and h.at["B-001", "이전공고번호"] == "A", (cnt, h["이전공고번호"].to_dict())
    rows = [mk("A", B, 0, 50), mk("B", "가상시 문화회관 신축 건축공사 2024", 60, 55, prev="A"), mk("B", "가상시 문화회관 신축 건축공사 2024", 62, 55, prev="B", ord_="001", kind="변경")]
    for order in ([0, 1, 2], [2, 1, 0], [1, 2, 0]):
        assert run([dict(rows[k]) for k in order])[1] == ["B"], order
    # '.0' 이 붙은 공고번호: '재공고로 N건 대체' 와 '재발주 금액변동' 로그
    h, l, lg = dedup.dedup_latest(pd.DataFrame([mk(20240100001.0, B, 0, 50), mk(20240100002.0, B, 20, 80, **RB)]), 0.30)
    r = h.set_index(h["공고번호"].map(dedup._no))["대표선정사유"]
    assert len(l) == 1 and "재공고로 1건 대체" in r["20240100002"], r.to_dict()
    assert any(x["항목"] == "재발주 금액변동" and dedup._no(x["공고번호"]) == "20240100002" for x in lg), [x["항목"] for x in lg]

    # ⑨ 다른 사업유형의 재공고에 대체돼 대표가 빠지면, 남은 공고의 사유와 '같은 공종 합산' 로그를 남은 건수에 맞춘다
    S1, S2, S3, T1 = "가상시 회관 건립 기계소방공사", "가상시 회관 건립 전기소방공사", "가상시 회관 건립 소방시설 내진보강공사", "가상시 회관 리모델링 기계소방공사"
    FI = dict(trade="소방")
    h, cnt, lg = run([mk("M", B, "2024-01-05 10:00:00", 50), mk("A1", S1, "2024-01-10 10:00:00", 5, **FI), mk("A2", S2, "2024-01-20 10:00:00", 4, **FI),
                      mk("A3", S3, "2024-02-20 10:00:00", 3, **FI), mk("B", T1, "2024-04-10 10:00:00", 5, prev="A1", **FI, **RM)])
    assert cnt == ["A2", "A3", "B", "M"], cnt
    assert h.at["A2-000", "대표선정사유"].startswith("합산(같은 공종 2건)") and h.at["A3-000", "대표선정사유"].startswith("합산(같은 공종 2건)"), h["대표선정사유"].to_dict()
    sums = [x["비고"] for x in lg if x["항목"] == "같은 공종 합산"]
    assert len(sums) == 1 and "공고 2건을 모두 합산" in sums[0] and "A1 " not in sums[0] and "A2 " in sums[0] and "A3 " in sums[0], sums
    h, cnt, lg = run([mk("M", B, "2024-01-05 10:00:00", 50), mk("A1", S1, "2024-01-10 10:00:00", 5, **FI), mk("A2", S2, "2024-01-20 10:00:00", 4, **FI),
                      mk("B", T1, "2024-04-10 10:00:00", 5, prev="A1", **FI, **RM)])
    assert cnt == ["A2", "B", "M"] and h.at["A2-000", "대표선정사유"].startswith("단독 공고"), (cnt, h["대표선정사유"].to_dict())
    assert not [x for x in lg if x["항목"] == "같은 공종 합산"], [x["비고"] for x in lg if x["항목"] == "같은 공종 합산"]      # 한 건만 남으면 합산 로그도 없다

    # ⑩ 공고연도의 근거(최종공고일): 대표 공고에 공고일시가 없으면 그 프로젝트의 다른 행에서 가져온다(빈칸으로 두지 않는다)
    hist, latest, _ = dedup.dedup_latest(pd.DataFrame([mk("A", B, "2024-01-05 10:00:00", 50), mk("A", B, None, 50, ord_="001", kind="변경"),
                                                       mk("K", B, "2024-03-05 10:00:00", 50, kind="취소")]), 0.30)
    fs = dedup.facility_summary(latest, hist)
    assert len(fs) == 1 and str(fs["최종공고일"].iloc[0])[:4] == "2024", fs["최종공고일"].tolist()

    # ⑪ dedup 단계 요약: '고유 공고' 도 '복수 시설 배정' 과 같은 공고번호 기준 / 남은 이중 합산 건수는 공백이 붙은 공고번호도 같은 공고로 센다
    out, lg = stage([mk("A", B, 0, 50), mk("A", B, 0, 50, fid="F2", fac="가상시 문화회관"), mk("A", B, 2, 50, fid="F2", fac="가상시 문화회관", ord_="001", kind="변경")], "uniq_rev")
    assert "대표 2건 (고유 공고 1건" in out and "복수 시설 배정 1건" in out, out
    umb = "가상도시공사 공공주택"
    part = [mk("A1", "가상 1단지 아파트 건설공사", 0, 900, fac=umb, fid="F9"), mk("A2", "가상 1단지 아파트 전기공사", 31, 90, fac=umb, fid="F9", trade="전기"),
            mk("R", "가상지구 진입도로 개설공사", 60, 50, fac=umb, fid="F9", trade="토목"),
            mk("B1", "가상 2단지 아파트 건설공사", 0, 800, fac=umb, fid="F9"), mk("B2", "가상 2단지 아파트 전기공사", 31, 80, fac=umb, fid="F9", trade="전기"),
            mk("A1", "가상 1단지 아파트 건설공사", 0, 900, fac="가상 1단지 아파트", fid="F2"), mk("A2", "가상 1단지 아파트 전기공사", 31, 90, fac="가상 1단지 아파트", fid="F2", trade="전기"),
            mk("R", "가상지구 진입도로 개설공사", 60, 50, fac="가상 1단지 아파트", fid="F2", trade="토목"),
            mk("B1", "가상 2단지 아파트 건설공사", 0, 800, fac="가상 2단지 아파트", fid="F3"), mk("B2", "가상 2단지 아파트 전기공사", 31, 80, fac="가상 2단지 아파트", fid="F3", trade="전기"),
            mk("R ", "가상지구 진입도로 개설공사", 60, 50, fac="가상 2단지 아파트", fid="F3", trade="토목")]
    out, lg = stage(part, "dup_partial_ws")
    note = lg.loc[lg["항목"] == "시설 간 중복 프로젝트", "비고"].iloc[0]
    assert "두 번 합산되는 공고가 1건 남음" in note and "두 번 합산되는 공고가 1건 남음" in out, (note, out)


def test_trade_run_feedback(tmp):
    """첫 실제 실행(2026-11-06) 피드백: 합산 행('도급비(건축, 토목, 기계, 조경)')은 '합계'로 배분 제외, 관급·부가세 보정 기준 판별, 손상 Excel 은 calamine 폴백, 저장 결과 재정규화."""
    from unittest import mock
    from openpyxl import Workbook
    from g2b_costdb import attachments, pipeline
    from g2b_costdb.extract_llm import normalize_trade_doc, cross_verify_trades, _trade_label, TRADE_LABELS
    assert "합계" in TRADE_LABELS
    assert _trade_label("기타", "도급비(건축, 토목, 기계, 조경)") == "합계" and _trade_label("건축", "건축공사 소계") == "건축"
    assert _trade_label("건축", "건축·기계설비 합계") == "합계" and _trade_label("기타", "총공사비") == "합계" and _trade_label("승강기", "승강기설치") == "기타"
    # 2차 실제 결과(2026-11-07) 피드백: 항목명 규칙 보정
    assert _trade_label("합계", "기계소방 관급자재비") == "관급자재" and _trade_label("기타", "지급자재비(도급비 미포함, 별도발주)") == "관급자재"
    assert _trade_label("기타", "부가가치세(계)") == "부가세" and _trade_label("부가세", "부 가 가 치 세 (총원가)*10%") == "부가세"
    assert _trade_label("합계", "□ 안전관리비 (산업안전보건관리비 등, 01.건축 9,200,000+02.부대토목 2,044,312)") == "간접비"
    assert _trade_label("간접비", "간접노무비+산재보험료+석면분담금+임금채권부담금+일반관리비+이윤") == "간접비" and _trade_label("간접비", "순공사비(안전관리비)") == "간접비"
    assert _trade_label("기타", "폐기물처리비(토목)") == "간접비" and _trade_label("기타", "소방시설공사 배상책임공제") == "간접비"
    assert [_trade_label("기타", x) for x in ("이설비(도급비 미포함)", "이설비 149,239,000원", "시설분담금", "설계 및 감리 용역비", "미술장식품", "건축물 인증수수료")] == ["배분제외"] * 6
    assert _trade_label("기타", "직접공사비") == "합계" and _trade_label("기타", "추정가격") == "합계"
    # 모델이 공종을 적은 '추정가격'·'도급비' 행은 문서 단위로: 이름이 공종인 다른 배분 행이 있으면 합계(중복 셈 방지), 그 행만이거나 모두 열 머리글 이름이면 모델 라벨 유지
    assert _trade_label("건축", "도급비") == "건축" and _trade_label("소방", "추정가격(공급가액)") == "소방"
    dd = normalize_trade_doc({"합계_원": "", "공종별금액": [{"공종": "건축", "항목명_원문": "추정가격(공급가액)", "금액_원": "5000"}, {"공종": "소방", "항목명_원문": "추정가격(공급가액)", "금액_원": "300"}]})
    assert [i["공종"] for i in dd["공종별금액"]] == ["건축", "소방"] and dd["_공종합_원"] == 5300 and dd["합계_원"] is None, dd
    dd = normalize_trade_doc({"합계_원": "", "공종별금액": [{"공종": "건축", "항목명_원문": "건축공사", "금액_원": "2000"}, {"공종": "토목", "항목명_원문": "토목공사", "금액_원": "300"},
                                                     {"공종": "건축", "항목명_원문": "도급비", "금액_원": "2300"}]})
    assert [i["공종"] for i in dd["공종별금액"]] == ["건축", "토목", "합계"] and dd["_공종합_원"] == 2300 and dd["합계_원"] == 2300, dd
    dd = normalize_trade_doc({"합계_원": "", "공종별금액": [{"공종": "건축", "항목명_원문": "건축공사", "금액_원": "500"}, {"공종": "기타", "항목명_원문": "직접공사비", "금액_원": "500"}]})
    assert dd["합계_원"] is None and dd["_공종합_원"] == 500, "직접공사비 소계는 합계_원 폴백에 쓰지 않음"
    dd = normalize_trade_doc({"합계_원": "1000", "공종별금액": [{"공종": "건축", "항목명_원문": "추정가격", "금액_원": "900"}, {"공종": "기타", "항목명_원문": "미확정설계공종(PS)", "금액_원": "100"}]})
    assert [i["공종"] for i in dd["공종별금액"]] == ["건축", "기타"] and dd["_공종합_원"] == 1000, "기타 행만 더 있으면 추정가격 행은 그 공종 그대로"
    # 반박 검증(2026-11-07)에서 나온 사례들
    assert [_trade_label("건축", x) for x in ("건축공사(전기·소방 별도)", "건축공사(전기, 통신, 소방 분리발주)", "건축공사(소방시설공사 별도)")] == ["건축"] * 3
    assert _trade_label("합계", "관급자재를 제외한 총공사비") == "합계" and _trade_label("합계", "관급자재비 포함 총공사비") == "합계" and _trade_label("기계설비", "관급자재 설치비") == "기계설비"
    assert _trade_label("기타", "철거 폐기물 처리") == "철거" and _trade_label("철거", "건축물 철거공사") == "철거" and _trade_label("철거", "가설건물 철거") == "철거"
    assert _trade_label("토목", "토목공사업 또는 토목건축공사업(주계약자)") == "토목" and _trade_label("전기공사금액", "") == "전기" and _trade_label("소방시설공사 총액", "") == "소방"
    assert _trade_label("기타", "경비실 신축공사") == "기타" and _trade_label("기타", "감리사무실 설치") == "기타" and _trade_label("기타", "ELEVATOR 설치공사") == "기타"
    assert _trade_label("전기", "한전 이설비(도급비 미포함)") == "배분제외" and _trade_label("기타", "임금채권부담금") == "간접비" and _trade_label("기타", "건설기계 운반비") == "기타"
    assert _trade_label("배분 제외", "") == "배분제외"
    # 2026-11-07 Excel 검토: 띄어 쓴 머리글, 괄호 밖 공종 우선, 소방 면허 분야, 측량
    assert _trade_label("건축", "건 축 공 사 도급금액") == "건축" and _trade_label("기타", "▶ 통 신 공 사") == "정보통신" and _trade_label("기타", "도 급 액") == "합계"
    assert _trade_label("합계", "합계 - 건축(건축+무대기계+기계) 소계") == "건축" and _trade_label("합계", "▶ 소 방 공 사(전기+기계)") == "소방"
    assert _trade_label("합계", "일반소방시설공사업(전기)") == "소방" and _trade_label("합계", "전문소방시설공사업(기초금액, 부가가치세포함)") == "소방"
    assert _trade_label("합계", "총공사비(건축)") == "합계" and _trade_label("합계", "건축, 기계공사") == "합계" and _trade_label("기타", "경계측량비") == "배분제외"
    assert _trade_label("기타", "기타금액(분담금 등)") == "배분제외"
    dd = normalize_trade_doc({"합계_원": "1982557093", "공종별금액": [{"공종": "건축", "항목명_원문": "건 축 공 사 도급금액", "금액_원": "1408297000"},
                                                               {"공종": "기계설비", "항목명_원문": "기계설비공사 도급금액", "금액_원": "228777000"}]})
    assert [i["공종"] for i in dd["공종별금액"]] == ["건축", "기계설비"], "띄어 쓴 '건 축 공 사 도급금액' 은 건축 행(전체 금액 행이 아님)"
    # 04b 배분적용: 공종 행 합이 집계표(간접비·관급·부가세 제외)의 30~125% 일 때만 비율 배분. 일부 행만 읽힌 문서·API 와 안 맞는 집계표는 N(사유)
    ok = normalize_trade_doc({"합계_원": "1000", "공종별금액": [{"공종": "건축", "항목명_원문": "건축공사", "금액_원": "600"}, {"공종": "기계설비", "항목명_원문": "기계설비공사", "금액_원": "100"},
                                                          {"공종": "간접비", "항목명_원문": "간접비", "금액_원": "200"}, {"공종": "부가세", "항목명_원문": "부가세", "금액_원": "100"}]})
    assert pipeline._alloc_decision({"공고번호": "A", "추정가격": 909, "기초금액": 1000}, ok) == "Y"
    partial = normalize_trade_doc({"합계_원": "11875390000", "공종별금액": [{"공종": "기타", "항목명_원문": "04 제요율적용제외공종", "금액_원": "43539000"},
                                                                       {"공종": "관급자재", "항목명_원문": "05 도급자 설치 관급자재", "금액_원": "1259942881"}, {"공종": "부가세", "항목명_원문": "부가가치세", "금액_원": "331811000"}]})
    dec = pipeline._alloc_decision({"공고번호": "B", "추정가격": 12793466000, "기초금액": 14072813000, "도급자관급액_API": 1259942881}, partial)
    assert dec.startswith("N(") and "일부 행만" in dec, dec
    mism = normalize_trade_doc({"합계_원": "448140000", "공종별금액": [{"공종": "전기", "항목명_원문": "▶ 전 기 공 사", "금액_원": "290380000"}, {"공종": "정보통신", "항목명_원문": "▶ 통 신 공 사", "금액_원": "279140000"}]})
    dec = pipeline._alloc_decision({"공고번호": "C", "추정가격": 3723926000, "기초금액": 4096319000}, mism)
    assert dec == "N(집계표 합계가 API 금액과 불일치)", dec
    assert pipeline._alloc_decision({"공고번호": "D"}, normalize_trade_doc({"합계_원": "", "공종별금액": [{"공종": "관급자재", "항목명_원문": "관급", "금액_원": "5"}]})) == "N(공종 행 없음)"
    # 합계_원이 없어도 읽은 행의 합이 API 와 맞지 않으면(전기·통신·소방 행만 읽힌 건축 공고) 배분하지 않음
    nototal = normalize_trade_doc({"합계_원": "", "공종별금액": [{"공종": "전기", "항목명_원문": "▶ 전 기 공 사", "금액_원": "290380000"}, {"공종": "정보통신", "항목명_원문": "▶ 통 신 공 사", "금액_원": "279140000"},
                                                             {"공종": "소방", "항목명_원문": "▶ 소 방 공 사(전기+기계)", "금액_원": "448140000"}]})
    dec = pipeline._alloc_decision({"공고번호": "F", "추정가격": 2639444545, "기초금액": 2903389000, "도급자관급액_API": 1052680000, "관급자관급액_API": 140250000}, nototal)
    assert dec == "N(읽은 행의 합이 API 금액과 불일치)", dec
    # 전체 금액 행('추정가격')을 공고 공종과 다른 공종으로 읽은 경우 06 참고
    lg = cross_verify_trades("G", {"공종": "건축", "추정가격": 118992000, "기초금액": 146661000, "도급자관급액_API": 27669000},
                             normalize_trade_doc({"합계_원": "146661000", "공종별금액": [{"공종": "전기", "항목명_원문": "추정가격", "금액_원": "108174545"}, {"공종": "부가세", "항목명_원문": "부가가치세", "금액_원": "10817455"},
                                                                                {"공종": "관급자재", "항목명_원문": "총 관급자재대", "금액_원": "27669000"}]}))
    assert any(l["항목"] == "공종 라벨 의심(문서)" and l["판정"] == "참고" and "전기" in l["비고"] for l in lg), lg
    assert not any(l["항목"] == "공종 라벨 의심(문서)" for l in cross_verify_trades("H", {"공종": "건축", "추정가격": 909, "기초금액": 1000}, ok))
    noapi = normalize_trade_doc({"합계_원": "", "공종별금액": [{"공종": "기계설비", "항목명_원문": "기계설비공사", "금액_원": "210875114"}, {"공종": "건축", "항목명_원문": "건축공사", "금액_원": "0"}]})
    dec = pipeline._alloc_decision({"공고번호": "E", "추정가격": 1162618182, "기초금액": 1278880000}, noapi)
    assert dec.startswith("N(") and "불일치" in dec, dec       # 합계_원 없음 → 읽은 행의 합(2.1억)이 API 기초금액(12.8억)과 맞지 않아 N
    # 04b 시트: 배분적용 열이 있고 N 인 공고의 비율은 비며, 열이 없는(옛) 표는 Y 로 본다
    from g2b_costdb.build_excel import DETAIL_COLS as XCOLS
    assert "배분적용" in XCOLS and XCOLS.index("배분적용") == XCOLS.index("배분금액(수식)") + 1
    dd = normalize_trade_doc({"합계_원": "1100", "공종별금액": [{"공종": "건축", "항목명_원문": "건축공사", "금액_원": "1000"}, {"공종": "기타", "항목명_원문": "이설비", "금액_원": "100"}]})
    lg = cross_verify_trades("W", {"추정가격": 1100, "기초금액": 1210}, dd)
    assert lg[0]["항목"] == "공종합=집계표합계" and lg[0]["판정"] == "정상", "배분제외 행이 집계표 총합에 든 경우도 정상"
    lg = cross_verify_trades("T", {"추정가격": 1000000, "기초금액": None, "도급자관급액_API": 1020},
                             normalize_trade_doc({"합계_원": "1000", "공종별금액": [{"공종": "건축", "항목명_원문": "건축", "금액_원": "1000"}]}))
    assert lg[-1]["판정"] == "경고" and "천원" in lg[-1]["비고"], "기초금액이 없을 때 관급만 남는 가짜 후보로 ×1000 의심을 덮지 않음"
    assert _trade_label("건축", "건축공사(도급비)") == "건축" and _trade_label("기타", "건축공사 합계") == "건축" and _trade_label("기타", "추정가격(건축공사업, 업종별 공종세부 미구분)") == "건축"
    assert _trade_label("건축", "직접공사비계") == "건축" and _trade_label("토목", "◆ 직접공사비 계") == "토목", "공종별 원가계산서의 직접공사비 소계는 그 공종"
    assert _trade_label("기계설비", "부계약자 기계설비·가스공사업 기초금액(원)") == "기계설비" and _trade_label("기타", "기계설비·가스공사업(주력분야 기계설비공사)(부계약자, 12%)") == "기계설비"
    assert _trade_label("기타", "건축공사업 또는 토목건축공사업(주계약자, 88%)") == "건축", "면허명 '토목건축공사업' 은 두 공종이 아님"
    assert _trade_label("건축", "건축공사업 - 추정가격+부가가치세(비율) 51,916,293,000원(86.4%)") == "건축" and _trade_label("소방", "전문소방시설공사업 - 추정가격+부가가치세(비율) 4,171,849,000원(6.9%)") == "소방"
    assert _trade_label("기타", "추정가격(건축,기계,토목,조경 도급공사비, 부가세제외)") == "합계" and _trade_label("기타", "1.전기소방공사") == "합계"
    assert _trade_label("기계설비", "총공사비 ---->>> (기계공사(소방별도)) 144020000") == "기계설비" and _trade_label("전기", "도급자설치") == "관급자재"
    assert _trade_label("기타", "미확정설계공종(PS)") == "기타" and _trade_label("기타", "04 제요율적용제외공종") == "기타"
    assert "배분제외" in TRADE_LABELS
    from g2b_costdb.build_excel import DETAIL_EXCLUDED
    assert "배분제외" in DETAIL_EXCLUDED and "합계" in DETAIL_EXCLUDED
    d3 = normalize_trade_doc({"합계_원": "1000", "공종별금액": [{"공종": "건축", "항목명_원문": "건축공사", "금액_원": "600"}, {"공종": "기타", "항목명_원문": "이설비(도급비 미포함)", "금액_원": "100"},
                                                     {"공종": "기타", "항목명_원문": "폐기물처리비", "금액_원": "50"}]})
    assert [i["공종"] for i in d3["공종별금액"]] == ["건축", "배분제외", "간접비"] and d3["_공종합_원"] == 600, d3
    # 집계표합계 기준: 1% 안이면 '=', 3% 안이면 '≈'(참고), 그 밖은 경고
    close = cross_verify_trades("Y", {"추정가격": 1000, "기초금액": 1100}, normalize_trade_doc({"합계_원": "1125", "공종별금액": [{"공종": "건축", "항목명_원문": "건축", "금액_원": "1125"}]}))
    assert close[-1]["판정"] == "참고" and "≈" in close[-1]["비고"] and "3% 이내" in close[-1]["비고"], close
    far = cross_verify_trades("Z", {"추정가격": 1000, "기초금액": 1100}, normalize_trade_doc({"합계_원": "1300", "공종별금액": [{"공종": "건축", "항목명_원문": "건축", "금액_원": "1300"}]}))
    assert far[-1]["판정"] == "경고", far
    d = normalize_trade_doc({"합계_원": "9394053000", "금액기준": "부가세포함", "공종별금액": [
        {"공종": "기타", "항목명_원문": "도급비(건축, 토목, 기계, 조경)", "금액_원": "7893683000"}, {"공종": "관급자재", "항목명_원문": "관급자재비", "금액_원": "1500370000"}]})
    assert [i["공종"] for i in d["공종별금액"]] == ["합계", "관급자재"] and d["_공종합_원"] is None, d
    logs = cross_verify_trades("X", {"추정가격": 7645251000, "기초금액": 7893683000, "도급자관급액_API": 751632000, "관급자관급액_API": 0}, d)
    assert logs[-1]["판정"] == "참고" and logs[-1]["비고"] == "집계표 합계 − 관급자재 = 기초금액(부가세 포함)", logs
    d2 = normalize_trade_doc({"합계_원": "", "공종별금액": [{"공종": "합계", "항목명_원문": "총공사비", "금액_원": "500"}, {"공종": "건축", "항목명_원문": "건축공사", "금액_원": "400"}]})
    assert d2["합계_원"] == 500 and d2["_공종합_원"] == 400, "합계_원이 비면 합계 행으로 채움"
    # 저장된 결과 재정규화: 옛 라벨('기타' 합산 행)이 Excel 단계에서 '합계' 로 바뀌고 _files 등은 유지
    old = {"K": {"공종별금액": [{"공종": "기타", "항목명_원문": "도급비(건축, 토목)", "금액_원": 100}], "합계_원": 100, "_files": ["a.xlsx"], "_usage_in": 3, "금액기준": "부가세포함"},
           "E": {"_error": "x"}}
    rn = pipeline._renormalize_trades(old)
    assert rn["K"]["공종별금액"][0]["공종"] == "합계" and rn["K"]["_files"] == ["a.xlsx"] and rn["K"]["_usage_in"] == 3 and rn["E"] == {"_error": "x"}
    # calamine 폴백: openpyxl 이 못 여는 xlsx 도 집계표를 읽는다(python-calamine 설치 시)
    d = os.path.join(tmp, "files3", "D1")
    os.makedirs(d)
    wb = Workbook()
    ws = wb.active
    ws.title = "총괄집계표"
    for row in (["(단위: 원)"], ["공종", "합계"], ["건축공사", 1000000], ["총공사비", 1000000]):
        ws.append(row)
    p = os.path.join(d, "내역서.xlsx")
    wb.save(p)
    try:
        import python_calamine  # noqa: F401
        import openpyxl
        with mock.patch.object(openpyxl, "load_workbook", side_effect=TypeError("_NamedCellStyle.name should be str")):
            s = attachments.cost_sheet_snippets(p)
        assert "건축공사\t1000000" in s and "총공사비\t1000000" in s, s
    except ImportError:
        print("(python-calamine 없음 → 폴백 테스트 생략)")
    assert attachments.has_cost_sheet(p)
    # calamine 의 러스트 panic(BaseException 계열)은 잡아서 '읽기 실패' 로 처리, 확장자와 내용이 다른 파일(.xls 인데 xlsx)은 사본으로 읽음
    PanicException = type("PanicException", (BaseException,), {})      # pyo3 가 던지는 것과 같은 이름·계열(모듈 등록 없이도 이름으로 판별)
    bad = os.path.join(d, "bad.xls")
    open(bad, "wb").write(b"\xd0\xcf\x11\xe0" + b"\x00" * 100)
    with mock.patch.object(attachments, "_read_sheets_calamine", side_effect=PanicException("range end index 3 out of range")):
        assert attachments.cost_sheet_snippets(bad) == "" and attachments.has_cost_sheet(bad) is False
        try:
            attachments.extract_any(bad)
            raise AssertionError("손상 xls 는 예외")
        except ValueError as e:
            assert "calamine" in str(e) and "xlrd" in str(e), e
    wb2 = Workbook()
    wb2.active.title = "총괄집계표"
    wb2.active.append(["건축공사", 1000])
    px = os.path.join(d, "공내역서.xls")        # xlsx 내용인데 .xls 확장자
    wb2.save(px)
    assert attachments._read_sheets(px)[0][0] == "총괄집계표" and "건축공사\t1000" in attachments.cost_sheet_snippets(px) and os.path.exists(px + ".as.xlsx")
    try:
        with mock.patch.object(attachments, "_read_sheets_calamine", side_effect=KeyboardInterrupt()):
            attachments.cost_sheet_snippets(bad)
        raise AssertionError("KeyboardInterrupt 는 삼키면 안 됨")
    except KeyboardInterrupt:
        pass
    cfg2 = {"paths": {"files_dir": os.path.join(tmp, "files3")}}
    with mock.patch.object(pipeline, "_trade_input", side_effect=PanicException("boom")):
        assert pipeline._trade_targets.__name__ == "_trade_targets"      # (대상 선정은 영향 없음)


def _hwp_record(tag: int, payload: bytes, level: int = 0) -> bytes:
    size = len(payload)
    if size < 0xFFF:
        return struct.pack("<I", tag | (level << 10) | (size << 20)) + payload
    return struct.pack("<I", tag | (level << 10) | (0xFFF << 20)) + struct.pack("<I", size) + payload


def test_attachments(tmp):
    assert attachments.safe_name("a" * 200 + ".hwp").endswith(".hwp") and attachments.safe_name("None") == ""
    assert attachments.sniff_ext(b"%PDF-1.7") == ".pdf" and attachments.sniff_ext(b"\xd0\xcf\x11\xe0\xa1") == ".hwp"
    assert attachments.looks_like_html(b"<!DOCTYPE html><html>") and not attachments.looks_like_html(b"%PDF")
    # HWP 레코드 파서: 문단 텍스트 + 인라인 컨트롤(16바이트) + 확장 크기 레코드
    text = "추정가격 1,000원"
    payload = text.encode("utf-16-le")
    ctrl = (9).to_bytes(2, "little") + b"\x00" * 14      # 탭 컨트롤(8 WCHAR = 16 bytes)
    payload2 = ctrl + "다음".encode("utf-16-le")
    big = ("가" * 3000).encode("utf-16-le")
    stream = _hwp_record(0x10 + 51, payload) + _hwp_record(0x10 + 50, b"\x00" * 8) + _hwp_record(0x10 + 51, payload2) + _hwp_record(0x10 + 51, big)
    recs = list(attachments._hwp_records(stream))
    assert [t for t, _ in recs] == [67, 66, 67, 67] and len(recs[3][1]) == 6000
    texts = [attachments._hwp_para_text(p) for t, p in recs if t == 67]
    assert texts[0] == text and texts[1] == "\t다음" and texts[2] == "가" * 3000
    # HWPX: <hp:tab/> 뒤 텍스트 보존
    import zipfile
    hx = os.path.join(tmp, "t.hwpx")
    xml = ('<?xml version="1.0"?><hs:sec xmlns:hs="http://www.hancom.co.kr/hwpml/2011/section" xmlns:hp="http://www.hancom.co.kr/hwpml/2011/paragraph">'
           '<hp:p><hp:run><hp:t>추정가격<hp:tab/>1,234원</hp:t></hp:run></hp:p></hs:sec>')
    with zipfile.ZipFile(hx, "w") as z:
        z.writestr("Contents/section0.xml", xml)
    assert "1,234원" in attachments.extract_hwpx(hx)
    # zlib 압축 섹션이 있는 가짜 HWP(OLE)는 olefile 로 만들기 어려우므로 레코드 파서까지만 검증
    # 다운로드: HTML 응답은 실패, 정상 응답은 .part → 교체, 확장자 없으면 매직바이트 보완, 기존 파일은 요청 없이 재사용
    class R:
        def __init__(self, chunks, ct=""):
            self._c, self.headers, self.status_code = chunks, {"Content-Type": ct}, 200
        def raise_for_status(self): pass
        def iter_content(self, n): return iter(self._c)
        def __enter__(self): return self
        def __exit__(self, *a): return False
    d = os.path.join(tmp, "dl")
    with mock.patch("requests.get", return_value=R([b"<html>login</html>"], "text/html")):
        assert attachments.download("http://x/f?fileSeq=1", d, "공고문.hwp") is None
    with mock.patch("requests.get", return_value=R([b"%PDF-1.4 ", b"body"])) as g:
        p = attachments.download("http://x/f?fileSeq=2", d, "")
        assert p and p.endswith(".pdf") and open(p, "rb").read() == b"%PDF-1.4 body" and not os.path.exists(p + ".part")
    with mock.patch("requests.get", return_value=R([b"%PDF-1.4 ", b"body"])) as g:
        p2 = attachments.download("http://x/f?fileSeq=3", d, "공고문.pdf")
        assert g.call_count == 1
        attachments.download("http://x/f?fileSeq=3", d, "공고문.pdf")
        assert g.call_count == 1, "이미 받은 파일은 요청하지 않음"
    # 같은 문서의 hwp/pdf 중복: 앞 형식이 성공하면 뒤 형식은 건너뛰고, 실패하면 다른 형식으로 폴백
    row2 = {"공고번호": "N8", "첨부URL1": "http://x/f?fileSeq=11", "첨부파일명1": "공고문.hwp", "첨부URL2": "http://x/f?fileSeq=12", "첨부파일명2": "공고문.pdf"}
    calls = []
    def fake_extract(path):
        calls.append(os.path.basename(path))
        if path.endswith(".pdf"):
            raise ValueError("PDF 텍스트 없음")
        return "본문 " * 40, "olefile"
    with mock.patch("requests.get", side_effect=lambda *a, **k: R([b"\xd0\xcf\x11\xe0" + b"x" * 100])), \
            mock.patch.object(attachments, "extract_any", side_effect=fake_extract):
        text, notes2 = attachments.process_notice_attachments(row2, os.path.join(tmp, "files2"), os.path.join(tmp, "text2"))
    assert [n["파일명"] for n in notes2] == ["공고문.pdf", "공고문.hwp"] and notes2[0]["추출성공"] == "N" and notes2[1]["추출성공"] == "Y" and text
    # HWP 문단 텍스트의 UTF-16 서로게이트: 쌍은 결합(이모지), 고아는 버림(그대로 두면 UTF-8 저장이 실패해 텍스트가 통째로 사라졌던 결함)
    pay = "가".encode("utf-16le") + b"\x3d\xd8\x00\xde" + b"\x3d\xd8" + "나".encode("utf-16le") + b"\x00\xdc"
    assert attachments._hwp_para_text(pay) == "가😀나", repr(attachments._hwp_para_text(pay))
    assert attachments.sanitize_text("a\ud83db\x00c\td\n") == "abc\td\n"
    row_s = {"공고번호": "N6", "첨부URL1": "http://x/f?fileSeq=6", "첨부파일명1": "공고문.hwp"}
    with mock.patch("requests.get", side_effect=lambda *a, **k: R([b"\xd0\xcf\x11\xe0" + b"x" * 100])), \
            mock.patch.object(attachments, "extract_any", return_value=("연면적 1,000㎡ \ud83d 본문", "olefile")):
        text_s, notes_s = attachments.process_notice_attachments(row_s, os.path.join(tmp, "files6"), os.path.join(tmp, "text6"))
    assert "연면적 1,000㎡" in text_s and "\ud83d" not in text_s and notes_s[0]["추출성공"] == "Y" and not notes_s[0]["오류"], notes_s
    assert os.listdir(os.path.join(tmp, "text6")), "텍스트 파일 저장됨"
    # 텍스트 파일 저장이 실패해도(예: 경로 길이 초과) 텍스트는 유지되고 오류만 기록
    real_open = open
    def failing_open(path, *a, **k):
        if str(path).endswith(".txt") and "text5" in str(path):
            raise OSError(22, "Invalid argument (경로 너무 김)")
        return real_open(path, *a, **k)
    with mock.patch("requests.get", side_effect=lambda *a, **k: R([b"\xd0\xcf\x11\xe0" + b"x" * 100])), \
            mock.patch.object(attachments, "extract_any", return_value=("본문 " * 40, "olefile")), \
            mock.patch("builtins.open", side_effect=failing_open):
        text_b, notes_b = attachments.process_notice_attachments(dict(row_s, 공고번호="N5"), os.path.join(tmp, "files5"), os.path.join(tmp, "text5"))
    assert text_b and notes_b[0]["추출성공"] == "Y" and "저장 실패" in notes_b[0]["오류"], notes_b
    assert attachments._short_name("a" * 200 + ".hwp").endswith(".hwp") and len(attachments._short_name("a" * 200 + ".hwp")) < 80
    row3 = dict(row2, 공고번호="N7")
    with mock.patch("requests.get", side_effect=lambda *a, **k: R([b"%PDF-1.4 " + b"x" * 100])), \
            mock.patch.object(attachments, "extract_any", return_value=("본문 " * 40, "pdfplumber")):
        text, notes3 = attachments.process_notice_attachments(row3, os.path.join(tmp, "files3"), os.path.join(tmp, "text3"))
    assert [n["파일명"] for n in notes3] == ["공고문.pdf"], "pdf 성공 시 hwp 는 건너뜀"
    # 첨부파일명 'None'/'nan' 은 빈 이름으로 취급되어 매직바이트 확장자로 저장
    row = {"공고번호": "N9", "첨부URL1": "http://x/f?fileSeq=9", "첨부파일명1": "None", "첨부URL2": "nan", "첨부파일명2": ""}
    with mock.patch("requests.get", return_value=R([b"%PDF-1.4 body of pdf " * 20])):
        with mock.patch.object(attachments, "extract_any", return_value=("본문 " * 40, "pdfplumber")):
            text, notes = attachments.process_notice_attachments(row, os.path.join(tmp, "files"), os.path.join(tmp, "text"))
    assert len(notes) == 1 and notes[0]["다운로드"] == "Y" and notes[0]["형식"] == ".pdf" and text


def run():
    # Windows 에서는 열려 있는 파일(수집 캐시 cache.sqlite 등)이 든 임시 폴더를 지우지 못한다 — 정리 실패로 뒤의 시험(test_api_review·test_servc_link)이 돌지 않는 일이 없게 한다
    with tempfile.TemporaryDirectory(ignore_cleanup_errors=True) as tmp:
        test_api_client(tmp)
        test_api_response_shapes_and_windows(tmp)
        test_collector_review(tmp)
        test_classify()
        test_discover_and_dedup(tmp)
        test_discover_keeps_review(tmp)
        test_discover_keeps_merged_row_order()
        test_discover_new_row_hints()
        test_audit_1121_fixes(tmp)
        test_institution_prefix()
        test_research_fallback_and_merge()
        test_research_other_orders()
        test_research_other_orders_more()
        test_supplement_marker_decisions()
        test_attach_merge_and_redo()
        test_output_schema_limits()
        test_error_summary()
        test_awards()
        test_extract_and_verify()
        test_handoff_and_batch(tmp)
        test_included_trades()
        test_trade_second_pass(tmp)
        test_trade_review_fixes(tmp)
        test_trade_allocation_formulas(tmp)
        test_amount_structure(tmp)
        test_request_review_fixes(tmp)
        test_work_split()
        test_work_split_review(tmp)
        test_work_split_review2(tmp)
        test_work_split_review3(tmp)
        test_trade_run_feedback(tmp)
        test_attachments(tmp)
    from tests import test_api_review          # 수집기·응답 해석·check-raw·probe 독립 검토의 확인 항목(가짜 응답만 사용)
    test_api_review.run()
    from tests import test_servc_link          # 용역 공고 연결(link-servc)
    test_servc_link.run()
    print("OK: 회귀 테스트 통과")


if __name__ == "__main__":
    run()
