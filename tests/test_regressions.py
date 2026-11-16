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
    agg2 = discover.discover_candidates(std, previous_review=prev)
    r = agg2[agg2["수요기관"] == "가상군"].iloc[0]
    assert r["검수_포함여부"] == "포함", "기본값 그대로였던 '제외'는 이어받지 않고 새 규칙(신축 → 포함)으로"
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
    with tempfile.TemporaryDirectory() as tmp:
        test_api_client(tmp)
        test_classify()
        test_discover_and_dedup(tmp)
        test_institution_prefix()
        test_research_fallback_and_merge()
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
        test_trade_run_feedback(tmp)
        test_attachments(tmp)
    print("OK: 회귀 테스트 통과")


if __name__ == "__main__":
    run()
