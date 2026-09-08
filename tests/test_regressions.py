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
    assert "nan" not in set(std["주공종명"]) and "None" not in set(std["첨부파일명1"])
    agg = discover.discover_candidates(std)
    # 같은 이름이라도 수요기관이 다르면 별도 시설
    assert len(agg) == 3 and set(agg["수요기관"]) == {"가상군", "다른군", "셋째군"}
    # 검수 내용 이어받기(시설ID·검수 컬럼)
    prev = agg.copy()
    prev.loc[prev["수요기관"] == "가상군", "검수_별칭(;구분)"] = "가상군문예회관"
    prev.loc[prev["수요기관"] == "가상군", "시설ID"] = "F0007"
    prev = prev[prev["수요기관"] != "셋째군"]            # 셋째군은 이전 파일에 없던 새 시설 → 이어지는 번호(F0008)
    agg2 = discover.discover_candidates(std, previous_review=prev)
    r = agg2[agg2["수요기관"] == "가상군"].iloc[0]
    assert r["시설ID"] == "F0007" and r["검수_별칭(;구분)"] == "가상군문예회관"
    assert agg2["시설ID"].is_unique and agg2[agg2["수요기관"] == "셋째군"]["시설ID"].iloc[0] == "F0008"
    # 검수 파일 왕복: 빈 검수_시설명 → 시설명_후보 대체, 수요기관 필터
    path = os.path.join(tmp, "facility_candidates.xlsx")
    agg2.loc[agg2["수요기관"] == "가상군", "검수_시설명"] = None
    discover.write_review_workbook(agg2, path)
    reviewed = discover.read_reviewed(path)
    assert set(reviewed["검수_시설명"]) == {"가상군 문화예술회관", "다른군 문화예술회관", "문화예술회관"}
    hits = discover.research_by_facility(std, reviewed)
    # '문화예술회관'(셋째군)은 이름만으로는 전국 매칭이지만 수요기관 필터로 셋째군 공고만
    third = hits[hits["시설명"] == "문화예술회관"]
    assert set(third["공고번호"]) == {"C1", "C2", "C3"}
    assert set(hits[hits["시설명"] == "가상군 문화예술회관"]["공고번호"]) == {"A1", "A2", "A3"}
    hist, latest, logs = dedup.dedup_latest(hits, 0.30, 0.30)
    rep = latest.set_index("프로젝트키")["공고번호"].to_dict()
    fid = reviewed.set_index("검수_시설명")["시설ID"]
    # 소액 설치공사(4억)가 본공사(300억) 대표를 대체하지 않음
    assert rep[f"{fid['가상군 문화예술회관']}-N|건축|"] == "A1", rep
    assert any(l["항목"] == "소액 공고 분리" for l in logs)
    # 단계 토큰: 괄호 안 (1단계)/(2단계)는 별도 프로젝트키, '2차 재공고'는 회차이므로 토큰 없음 → C2 를 C3 가 대체(날짜 없는 C3 는 가장 오래된 것으로 취급되어 C2 유지)
    f3 = fid["문화예술회관"]
    assert rep[f"{f3}-N|건축|1단계"] == "C1" and rep[f"{f3}-N|건축|2단계"] == "C2"
    assert rep[f"{f3}-N|건축|"] == "C3" and hist[hist["공고번호"] == "C3"]["최신여부"].iloc[0]
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
    with tempfile.TemporaryDirectory() as tmp:
        test_api_client(tmp)
        test_classify()
        test_discover_and_dedup(tmp)
        test_extract_and_verify()
        test_attachments(tmp)
    print("OK: 회귀 테스트 통과")


if __name__ == "__main__":
    run()
