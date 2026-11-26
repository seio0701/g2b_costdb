"""3단계(b) 용역 공고 연결(link-servc)의 회귀 테스트(네트워크·API 키 불필요). 실행: python -m tests.test_servc_link
회귀 테스트(tests.test_regressions)가 끝에서 함께 실행한다. ※ 모든 공고·기관·시설은 가상의 예시."""
from __future__ import annotations

import contextlib
import io
import os
import shutil
import subprocess
import sys
import tempfile
import time

import pandas as pd

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
from g2b_costdb import build_excel, dedup, discover, pipeline, servc  # noqa: E402
from g2b_costdb.classify import load_config  # noqa: E402


def test_kinds():
    k = servc.classify_kind
    ok = lambda t, kind, sub=None: (k(t)[0] == kind and not k(t)[2] and (sub is None or sub in k(t)[1])) or (_ for _ in ()).throw(AssertionError((t, k(t))))
    no = lambda t, part="": (bool(k(t)[2]) and part in k(t)[2] + k(t)[1]) or (_ for _ in ()).throw(AssertionError((t, k(t))))
    # 설계: 공모·용역·분리발주·재설계. '공모' 가 없는 제목도 설계
    ok("◇ 건립 설계공모", "설계", "설계공모") and ok("◇ 신축공사 설계용역 제안공모 공고", "설계", "설계공모") and ok("◇ 건축설계경기", "설계", "설계공모")
    ok("◇ 증축공사 기본 및 실시설계용역", "설계", "설계용역") and ok("◇ 신축 전기, 통신, 전기소방 설계용역", "설계", "분리발주") and ok("전기설계용역(◇ 신축사업)", "설계", "분리발주")
    ok("◇ 타워 (재)실시설계용역", "설계", "재설계") and ok("◇ 건립 국제지명 설계공모", "설계")
    # 설계비가 아닌 부대 용역·전시 제작·일괄 계약은 '기타'(08 에만)
    for t in ("◇ 건립 국제설계공모 관리용역", "◇ 건립공사 설계의 경제성(VE)등 검토용역", "◇ 건립사업 설계의도 구현 용역", "◇ 건립 타당성조사 및 기본계획 수립 용역",
              "◇ 실시설계 안전성검토 용역", "◇ 전시 기획·설계 및 전시물 제작·설치 용역", "◇ 조성 설계 및 감리용역", "◇ 증축공사 전력시설물 설계감리용역"):
        assert k(t)[0] == "기타" and k(t)[2], (t, k(t))
    assert k("◇ 위탁운영 수탁기관 선정 제안공모")[0] == "기타", "'제안공모' 만으로는 설계가 아니다"
    # 건설사업관리·감리. 공종은 세부에. 설계단계가 든 건설사업관리는 표시만 하고 더한다
    ok("◇ 신축공사 감독권한대행 등 건설사업관리용역(계속비)", "건설사업관리·감리", "건설사업관리") and ok("◇ 건립 건설사업관리(CM)용역", "건설사업관리·감리")
    ok("◇ 신축 전기공사 감리용역", "건설사업관리·감리", "전기") and ok("◇ 건립공사 소방감리용역(긴급)", "건설사업관리·감리", "소방")
    ok("◇ 증축 기타 정보통신공사 감리용역", "건설사업관리·감리", "정보통신") and ok("◇ 신축공사 건축사(상주) 감리용역", "건설사업관리·감리", "건축")
    assert "설계단계" in k("◇ 건립사업 실시설계에 대한 통합건설사업관리 및 설계VE 용역")[1] and not k("◇ 건립사업 실시설계에 대한 통합건설사업관리 및 설계VE 용역")[2]
    assert k("◇ 건립공사 CMS 구축 용역")[0] == "기타", "'CM' 은 앞뒤에 영문자가 없을 때만"
    no("◇ 석면해체 제거공사 감리용역", "석면") and no("(구)◇ 해체공사 감리용역", "석면·해체")
    # 건축 사업의 용역이 아닌 것
    for t in ("차세대 ◇ 통합의료정보시스템 구축 감리 용역", "◇ 유무선 네트워크 개선 사업 정보통신공사 감리용역", "2026년 제2차 정보자원 통합구축 감리용역"):
        assert k(t)[0] == "기타", (t, k(t))
    no("○○지구 주택건설공사 감리자 지정 입찰공고", "주택법") and ok("◇ 통합공공임대주택 건설공사 감리자(건축) 모집 공고", "건설사업관리·감리")
    no("2025년 소규모 시설공사 설계용역 연간단가계약", "단가")
    # 폐기물: 공사 폐기물만 '폐기물처리'. 운영 폐기물·폐기물 시설의 설계·감리는 기타, 연간·단가 계약은 더하지 않는다
    ok("◇ 신축공사 건설폐기물 처리용역 입찰 공고", "폐기물처리", "건설폐기물") and ok("◇ 건립공사 폐기물 운반 및 처리 용역", "폐기물처리")
    ok("◇ 신축공사 지정폐기물(석면) 처리용역", "폐기물처리", "지정폐기물") and ok("◇ 건립공사 건설폐기물(폐보드) 처리용역", "폐기물처리")
    for t in ("◇ 생활폐기물 수집·운반 대행용역", "2026년 음식물류폐기물 수집·운반 대행 용역(1구역)", "폐기물처리시설 운영·관리 위탁 용역(◇)", "◇ 자원순환시설 설치사업 건설사업관리용역"):
        assert k(t)[0] == "기타", (t, k(t))
    assert k("청사 청소 위탁", "생활폐기물수집운반서비스")[0] == "기타" and k("◇ 건립공사 처리용역", "건설폐기물처리서비스")[0] == "폐기물처리", "공고명에 없으면 조달분류로"
    no("2026년 연간단가 폐기물 처리용역(1구역)", "단가") and no("2026년 상반기 ◇ 폐기물 처리용역", "단가") and no("2026년 적환장 건설폐기물 처리용역", "단가")
    # 묶음 표기
    for t in ("◇ 외 1교(△△) 신축공사 통합건설사업관리용역", "◇ 등 2교 다목적강당 증축 감리용역", "◇ 외 2개소 신축공사 감리용역", "◇ 등 3개 블록 아파트 건설공사 건설사업관리용역"):
        assert servc.BUNDLE.search(t), t
    assert not servc.BUNDLE.search("◇ 체육관 및 급식소 증축 공사 감리용역") and not servc.BUNDLE.search("◇ 관람석 등 개보수공사 감리용역")
    # 재공고 열쇠: 재공고·정정·PQ 안내·'공사/사업' 차이는 같은 용역, 차수·폐기물 종류·(재)설계는 다른 용역
    rk = servc.rebid_key
    assert rk("◇ 청사 신축공사 지정폐기물(석면) 처리용역 입찰공고") == rk("◇ 청사 신축 지정폐기물(석면) 처리용역(재공고)")
    assert rk("◇ 건립 건설사업관리용역 사업수행능력평가서 제출 안내 공고(긴급)") == rk("◇ 건립 건설사업관리용역")
    assert rk("◇ 신축공사 건설폐기물처리용역") != rk("◇ 신축공사 임목폐기물처리용역") and rk("◇ 보수공사(2차분) 감리용역") != rk("◇ 보수공사(3차분) 감리용역")
    assert rk("◇ 타워 실시설계용역") != rk("◇ 타워 (재)실시설계용역")


def test_institutions():
    same, rel = servc.inst_same, servc.inst_related
    assert same("경상북도 경주시", "경상북도 경주시 시립도서관") and same("경상남도 거창군 체육시설사업소", "경상남도 거창군 인구교육과"), "같은 시·군의 사업소·부서"
    assert same("강원도 강릉시 농업기술센터", "강원특별자치도 강릉시") and same("전라북도 군산시", "전북특별자치도 군산시"), "광역 이름이 바뀌어도 같은 시·군"
    assert same("전라남도 화순군", "전남광주통합특별시 화순군") and same("광주광역시", "전남광주통합특별시 건설본부")
    assert same("서울특별시 한강사업본부", "서울특별시 도시기반시설본부") and same("천안시", "충청남도 천안시")
    assert not same("부산광역시 강서구", "서울특별시 강서구") and not same("경상남도 고성군", "강원특별자치도 고성군"), "이름이 같은 다른 시·군·구"
    assert not same("서울특별시", "서울특별시 송파구") and not same("경기도 수원시", "경기도 성남시") and not same("경기도", "경기도 수원시")
    assert same("경상남도교육청 경상남도창원교육지원청", "경상남도교육청 함양제일고등학교") and same("전라북도교육청", "전북특별자치도교육청 전주교육지원청")
    assert not same("경상남도교육청", "경상북도교육청") and not same("경상남도교육청", "경상남도 창원시")
    assert same("한국토지주택공사 경기남부지역본부", "한국토지주택공사") and same("(재)세종문화회관", "재단법인세종문화회관") and not same("부산교통공사", "서울교통공사")
    assert not same("", "경기도 수원시") and not same("조달청", "경기도 수원시")
    assert rel("경기도 성남시", "성남도시개발공사") and rel("과천도시공사", "경기도 과천시") and rel("서울특별시 송파구", "서울특별시 도시기반시설본부")
    assert not rel("경기도 성남시", "용인도시공사") and not rel("경기도 성남시", "경기도 성남시 수정구") and not rel("경기도 성남시", "성남교육지원청")


def _world(cfg):
    """가상 시설 둘(가상시립미술관: 신축·리모델링 두 프로젝트 / 예시군 문화예술회관: 신축)을 실제 단계 함수로 만든다."""
    import tests.test_dedup_and_excel as T
    std = discover.standardize(pd.DataFrame(T.RAW), cfg, pd.DataFrame(T.BSIS))
    agg = discover.discover_candidates(std)
    reviewed = agg[agg["검수_포함여부"] == "포함"].copy()
    for c in ("검수_별칭(;구분)", "검수_수요기관"):
        reviewed[c] = ""
    with contextlib.redirect_stderr(io.StringIO()):
        hits = discover.research_by_facility(std, reviewed)
    hist, latest, logs = dedup.dedup_latest(hits, cfg["dedup"]["price_change_warn_ratio"])
    return reviewed, hist, latest, logs


def _rows(cfg, items):
    f = cfg["fields"]
    out = []
    for it in items:
        d = {f["bid_no"]: it["no"], f["bid_ord"]: it.get("ord", "000"), f["name"]: it["name"], f["kind"]: it.get("kind", "등록공고"), f["re_notice"]: it.get("re", "N"),
             f["reg_type"]: "", f["notice_dt"]: it["dt"], f["notice_inst"]: it.get("ntc", "조달청"), f["demand_inst"]: it.get("inst", "가상광역시"),
             f["presmpt_price"]: str(it.get("price", "")), f["servc_budget"]: str(int(it["price"] * 1.1)) if it.get("price") else "", f["servc_div"]: "기술용역",
             f["servc_class"]: it.get("cls", ""), f["detail_url"]: "https://example.invalid/s", f["prev_bid_no"]: it.get("prev", ""), f["demand_inst_cd"]: it.get("code", "")}
        out.append(d)
    return pd.DataFrame(out)


def test_link():
    cfg = load_config()
    reviewed, hist, latest, _ = _world(cfg)
    mid = reviewed.loc[reviewed["검수_시설명"] == "가상시립미술관", "시설ID"].iloc[0]
    hid = reviewed.loc[reviewed["검수_시설명"] == "예시군 문화예술회관", "시설ID"].iloc[0]
    N, R, H = f"{mid}-N", f"{mid}-R", f"{hid}-N"
    raw = _rows(cfg, [
        dict(no="S01", name="가상시립미술관 건립공사 건설사업관리용역", dt="2024-01-20 10:00:00", price=900_000_000),
        dict(no="S02", name="가상시립미술관 건립 설계공모", dt="2022-03-02 10:00:00", price=800_000_000),
        dict(no="S03", name="가상시립미술관 건립공사 건설폐기물 처리용역", dt="2024-09-02 10:00:00", price=50_000_000),
        dict(no="S04", name="가상시립미술관 건립 건설폐기물 처리용역(재공고)", dt="2024-09-20 10:00:00", price=52_000_000, re="Y"),          # 이름이 조금 달라도 재공고
        dict(no="S05", name="가상시립미술관 생활폐기물 수집운반 대행용역", dt="2024-05-02 10:00:00", price=30_000_000),
        dict(no="S06", name="가상시립미술관 노후시설 개보수공사 전기공사 감리용역", dt="2026-05-20 10:00:00", price=40_000_000),
        dict(no="S07", name="가상시립미술관 건립공사 건설사업관리용역", dt="2024-01-25 10:00:00", price=900_000_000, inst="다른광역시"),           # 다른 기관 → 잇지 않음
        dict(no="S08", name="가상시립미술관 건립 전기공사 감리용역", dt="2024-03-20 10:00:00", price=150_000_000, ord="000"),
        dict(no="S08", name="가상시립미술관 건립 전기공사 감리용역(변경)", dt="2024-03-25 10:00:00", price=160_000_000, ord="001", kind="변경공고"),  # 최신 차수의 금액
        dict(no="S09", name="가상시립미술관 건립 소방공사 감리용역", dt="2024-04-01 10:00:00", price=70_000_000, ord="000"),
        dict(no="S09", name="가상시립미술관 건립 소방공사 감리용역", dt="2024-04-05 10:00:00", price=70_000_000, ord="001", kind="취소공고"),      # 취소
        dict(no="S10", name="가상시립미술관 외 2개소 건립 정보통신공사 감리용역", dt="2024-03-22 10:00:00", price=90_000_000),                 # 묶음
        dict(no="S11", name="가상시립미술관 LED 조명 교체공사 감리용역", dt="2024-06-01 10:00:00", price=20_000_000),                      # 유지보수
        dict(no="S12", name="가상시립미술관 건립 타당성조사 용역", dt="2021-03-01 10:00:00", price=80_000_000, cls=""),                    # 수집 규칙에는 안 걸리지만 들어와도 기타
        dict(no="S13", name="가상시립미술관 감리용역", dt="2010-01-01 10:00:00", price=10_000_000),                                 # 시기가 맞는 프로젝트 없음
        dict(no="S14", name="가상시립 미술관 건립공사 임목폐기물 처리용역", dt="2024-09-03 10:00:00", price=8_000_000),                     # 폐기물 종류가 다르면 다른 용역(둘 다 더함)
        dict(no="S15", name="가상시립미술관 건립공사 건설사업관리용역", dt="2024-01-22 10:00:00", price=950_000_000, inst="가상광역시 건설본부", prev="S01"),  # 같은 광역의 건설본부, 이전공고번호
        dict(no="S16", name="가상시립미술관 폐기물 처리용역", dt="2024-10-02 10:00:00", price=5_000_000),                               # 어느 공사의 폐기물인지 모름
        dict(no="S17", name="예시군 문화예술회관 건립 건설사업관리용역", dt="2025-01-15 10:00:00", price=1_500_000_000, inst="예시군 문화시설사업소"),
        dict(no="S18", name="예시군 문화예술회관 건립공사 설계용역", dt="2023-02-01 10:00:00", price=1_200_000_000, inst="예시도시공사"),          # 대행 기관 → 추정
        dict(no="S19", name="예시군 문화예술회관 및 가상시립미술관 건립 설계공모 관리용역", dt="2023-02-01 10:00:00", price=30_000_000, inst="예시군"),
        dict(no="S20", name="가상시립미술관 리모델링 실시설계용역", dt="2025-06-01 10:00:00", price=200_000_000),                          # 사업 이름은 다르지만 사업유형으로 R
        dict(no="S21", name="가상시립미술관 건립공사 건설폐기물 처리용역(2차)", dt="2026-03-02 10:00:00", price=90_000_000),                  # 550일 안이지만 금액이 다름 → 다음 차수
    ])
    out = servc.link(raw, reviewed, hist, latest, cfg)
    assert list(out.columns) == servc.LINK_COLS == build_excel.SERVC_COLS and servc.KINDS == tuple(build_excel.SERVC_KINDS)
    g = {(r["공고번호"], r["시설ID"]): r for _, r in out.iterrows()}
    r = lambda no, fid=mid: g[(no, fid)]
    assert ("S07", mid) not in g and len(out[out["공고번호"] == "S08"]) == 1, "다른 기관의 같은 이름 공고는 잇지 않는다 / 공고번호마다 한 줄"
    # 프로젝트 배정과 합산
    assert r("S01")["프로젝트ID"] == N and r("S01")["상태"] == "재공고 이전" and r("S15")["합산여부"] == "Y" and r("S15")["추정가격"] == 950_000_000, "이전공고번호로 대체"
    assert r("S02")["프로젝트ID"] == N and r("S02")["합산여부"] == "Y" and r("S02")["용역세부"] == "설계공모" and r("S02")["공사대비_개월"] < -20
    assert r("S03")["상태"] == "재공고 이전" and r("S03")["합산여부"] == "N" and r("S04")["합산여부"] == "Y" and r("S04")["프로젝트ID"] == N
    assert r("S14")["합산여부"] == "Y" and r("S14")["프로젝트ID"] == N, "띄어쓰기가 다른 시설 이름 / 폐기물 종류가 다른 공고는 재공고가 아니다"
    assert r("S21")["합산여부"] == "Y" and r("S04")["상태"] == "대표", "같은 이름이라도 멀리 떨어지고 금액이 다르면 다음 차수의 용역"
    assert r("S06")["프로젝트ID"] == R and r("S06")["합산여부"] == "Y" and r("S06")["용역세부"] == "감리(전기)"
    assert r("S20")["프로젝트ID"] == R and r("S20")["합산여부"] == "Y", "공고명의 사업유형(리모델링)으로 프로젝트를 고른다"
    assert r("S08")["추정가격"] == 160_000_000 and r("S08")["공고차수"] == "001" and r("S08")["합산여부"] == "Y", "변경공고는 최신 차수의 금액"
    assert r("S09")["상태"] == "취소" and r("S09")["합산여부"] == "N" and r("S09")["합산제외_까닭"] == "취소 공고"
    assert r("S05")["용역종류"] == "기타" and r("S05")["합산여부"] == "N"
    assert r("S10")["합산여부"] == "N" and "묶은" in r("S10")["합산제외_까닭"]
    assert r("S11")["합산여부"] == "N" and "유지보수" in r("S11")["합산제외_까닭"]
    assert r("S12")["용역종류"] == "기타" and r("S13")["프로젝트ID"] == "" and "시점이 맞는 프로젝트가 없음" in r("S13")["합산제외_까닭"]
    assert r("S16")["합산여부"] == "N" and "어느 공사" in r("S16")["합산제외_까닭"]
    assert r("S17", hid)["프로젝트ID"] == H and r("S17", hid)["합산여부"] == "Y", "같은 군의 사업소"
    # 대행 기관이 낸 공고는 '추정' 이지만, 공고명에 그 프로젝트의 사업 이름('예시군 문화예술회관 건립')이 통째로 있으면 더한다(2026-11 실자료 점검. 사업 이름이 없으면 N — test_field_audit ⑥)
    assert r("S18", hid)["연결강도"] == "추정" and r("S18", hid)["합산여부"] == "Y" and "대행 기관" in r("S18", hid)["연결근거"]
    assert ("S19", hid) in g and ("S19", mid) not in g and r("S19", hid)["용역종류"] == "기타", "기관이 다른 시설에는 잇지 않는다"
    # 05 에 더해지는 금액: 프로젝트·종류별 합
    y = out[out["합산여부"] == "Y"].groupby(["프로젝트ID", "용역종류"])["추정가격"].sum().to_dict()
    assert y == {(N, "건설사업관리·감리"): 950_000_000 + 160_000_000, (N, "설계"): 800_000_000, (N, "폐기물처리"): 52_000_000 + 8_000_000 + 90_000_000,
                 (R, "건설사업관리·감리"): 40_000_000, (R, "설계"): 200_000_000, (H, "건설사업관리·감리"): 1_500_000_000, (H, "설계"): 1_200_000_000}, y
    # 비율 점검: 지나치게 큰 용역비만 경고
    big = servc.link(_rows(cfg, [dict(no="B1", name="가상시립미술관 노후시설 개보수공사 설계용역", dt="2025-10-01 10:00:00", price=2_000_000_000)]), reviewed, hist, latest, cfg)
    lg = servc.ratio_logs(big, latest)
    assert len(lg) == 1 and lg[0]["공고번호"] == R and lg[0]["판정"] == "경고" and "용역비 비율(설계)" == lg[0]["항목"] and not servc.ratio_logs(out, latest), lg
    # 요약·진단 글: 건수가 맞고, 공고명은 싣지 않는다(진단)
    txt = servc.summary(out, len(raw), raw[cfg["fields"]["bid_no"]].nunique(), latest)
    assert "합산여부 Y" in txt and "설계: 3건" in txt and "폐기물처리: 3건" in txt and "더하지 않는 줄" in txt, txt
    prof = servc.profile(raw, cfg)
    assert "공고번호 기준 21건" in prof and "가상시립미술관" not in prof and "묶음 표기" in prof, prof
    assert servc.link(raw.iloc[0:0], reviewed, hist, latest, cfg).empty and servc.summary(out.iloc[0:0], 0, 0, latest).count("\n") == 1
    return reviewed, hist, latest, out


def test_name_and_project_rules():
    cfg = load_config()
    f = cfg["fields"]
    base = dict(공고차수="000", 공고종류="일반", 추정가격=1_000_000_000, 공종="건축", 사업유형="신축", 수요기관="경상남도 고성군", 세부사업명="")
    hist = pd.DataFrame([
        dict(base, 시설ID="F1", 시설명="고성군 수장고", 프로젝트ID="F1-N", 공고번호="C1", 공고명="수장고 건립공사", 공고일시="2024-03-01 10:00:00", 수요기관코드="5350000"),
        dict(base, 시설ID="F2", 시설명="가람 교양센터", 프로젝트ID="F2-N-01", 공고번호="C2", 공고명="가람근린공원 교양(국민체육)센터 건립공사", 공고일시="2024-03-01 10:00:00",
             수요기관="강원도 강릉시", 세부사업명="교양(국민체육)센터 건립"),
        dict(base, 시설ID="F2", 시설명="가람 교양센터", 프로젝트ID="F2-N-02", 공고번호="C3", 공고명="가람근린공원 교양(국민체육)센터 건립사업 경관조명 조성공사",
             공고일시="2025-03-01 10:00:00", 수요기관="강원특별자치도 강릉시", 세부사업명="경관조명 조성", 추정가격=200_000_000),
        dict(base, 시설ID="F3", 시설명="나래 체육관", 프로젝트ID="F3-R-01", 공고번호="C4", 공고명="나래 체육관 시설개선공사", 공고일시="2023-04-01 10:00:00", 사업유형="리모델링",
             수요기관="경기도 가온시"),
        dict(base, 시설ID="F3", 시설명="나래 체육관", 프로젝트ID="F3-R-02", 공고번호="C5", 공고명="나래 체육관 시설개선공사", 공고일시="2025-04-01 10:00:00", 사업유형="리모델링",
             수요기관="경기도 가온시"),
        dict(base, 시설ID="F4", 시설명="다온 문화 회관", 프로젝트ID="F4-N", 공고번호="C6", 공고명="다온 복합 문화 체육 회관 건립공사", 공고일시="2024-05-01 10:00:00",
             수요기관="충청북도 다온군"),
        dict(base, 시설ID="F5", 시설명="라온 임대주택", 프로젝트ID="F5-R", 공고번호="C7", 공고명="2025년 상반기 라온 임대주택 리모델링공사(1구역)", 공고일시="2025-03-01 10:00:00",
             사업유형="리모델링", 수요기관="경기도 가온시"),
    ])
    latest = hist.copy()
    reviewed = pd.DataFrame([
        {"시설ID": "F1", "검수_시설명": "고성군 수장고", "검수_별칭(;구분)": "", "검수_수요기관": "", "수요기관": "경상남도 고성군", "검색어": "수장고"},
        {"시설ID": "F2", "검수_시설명": "가람 교양센터", "검수_별칭(;구분)": "가람근린공원 교양센터", "검수_수요기관": "", "수요기관": "강원도 강릉시", "검색어": "체육센터"},
        {"시설ID": "F3", "검수_시설명": "나래 체육관", "검수_별칭(;구분)": "", "검수_수요기관": "", "수요기관": "경기도 가온시", "검색어": "체육관"},
        {"시설ID": "F4", "검수_시설명": "다온 문화 회관", "검수_별칭(;구분)": "", "검수_수요기관": "", "수요기관": "충청북도 다온군", "검색어": "문화회관"},
        {"시설ID": "F5", "검수_시설명": "라온 임대주택", "검수_별칭(;구분)": "", "검수_수요기관": "", "수요기관": "경기도 가온시", "검색어": "임대주택"},
        {"시설ID": "F9", "검수_시설명": "공고가 없는 시설", "검수_별칭(;구분)": "", "검수_수요기관": "", "수요기관": "경기도 가온시", "검색어": "체육관"},
    ])
    keys = servc.facility_keys(reviewed, hist)
    assert set(keys) == {"F1", "F2", "F3", "F4", "F5"} and "수장고" in keys["F1"]["pats"] and keys["F1"]["codes"] == {"5350000"}, keys
    mk = lambda no, name, dt, inst, **kw: dict(no=no, name=name, dt=dt, inst=inst, price=kw.pop("price", 50_000_000), **kw)
    raw = _rows(cfg, [
        mk("A1", "수장고 건립공사 설계용역", "2023-01-10 10:00:00", "경상남도 고성군"),
        mk("A2", "학용정수장 고도정수처리시설 설치사업 건설사업관리용역", "2024-01-10 10:00:00", "경상남도 고성군"),      # '수장고' 가 띄어쓰기를 가로질러 걸림 → 잇지 않음
        mk("A3", "수장고 건립공사 건설사업관리용역", "2024-02-10 10:00:00", "강원특별자치도 고성군"),                # 이름이 같은 다른 군
        mk("A4", "수장고 건립공사 감리용역", "2024-02-12 10:00:00", "", code="5350000", ntc="조달청"),              # 수요기관 이름이 없어도 코드가 같으면
        mk("B1", "가람근린공원 교양(국민체육)센터 건립사업 건설사업관리용역", "2024-02-01 10:00:00", "강원도 강릉시"),           # 괄호 속을 빼면 별칭과 같다
        mk("B2", "가람근린공원 교양(국민체육)센터 건립사업 경관조명 조성공사 실시설계용역", "2024-08-25 10:00:00", "강원특별자치도 강릉시 체육시설사업소"),   # 두 프로젝트 모두 넓은 창 안
        mk("B3", "가람 다목적 교양센터 건립 폐기물 처리용역", "2024-06-01 10:00:00", "강원특별자치도 강릉시"),              # 정확한 이름으로 찾은 공고가 있는 시설은 느슨한 꼴을 보지 않는다
        mk("D1", "다온 복합 문화 체육 회관 건립 건설사업관리용역", "2024-04-01 10:00:00", "충청북도 다온군"),             # 정확한 이름으로는 없고 어절 사이에 다른 말 → 추정
        mk("C1", "나래 체육관 시설개선공사 감리용역", "2023-04-10 10:00:00", "경기도 가온시"),                       # 같은 이름의 프로젝트 둘: 시점이 맞는 하나
        mk("C2", "나래 체육관 시설개선공사 설계용역", "2022-12-01 10:00:00", "경기도 가온시"),                       # 설계는 공사보다 먼저: 2023 공사 것(2025 공사는 28개월 뒤 — 좁은 창 안)
        mk("C3", "나래 체육관 시설개선공사 설계용역(2024)", "2024-06-01 10:00:00", "경기도 가온시"),
        # '상반기'·'1구역' 은 그 공사 이름에 든 말 → 연간·단가 계약으로 보지 않는다. 공사 이름에 없는 '(연간단가)' 가 더 있으면 본다
        mk("E1", "2025년 상반기 라온 임대주택 리모델링공사(1구역) 건설사업관리용역", "2025-02-10 10:00:00", "경기도 가온시"),
        mk("E2", "2025년 상반기 라온 임대주택 리모델링공사(1구역) 건설폐기물 처리용역", "2025-06-10 10:00:00", "경기도 가온시", cls="건설폐기물처리서비스"),
        mk("E3", "2025년 상반기 라온 임대주택 리모델링공사(1구역) 감리용역(연간단가)", "2025-02-12 10:00:00", "경기도 가온시"),
        mk("E4", "라온 임대주택 2구역 리모델링 감리용역", "2025-02-14 10:00:00", "경기도 가온시"),
    ])
    out = servc.link(raw, reviewed, hist, latest, cfg)
    g = {r["공고번호"]: r for _, r in out.iterrows()}
    assert "A2" not in g and "A3" not in g and g["A1"]["프로젝트ID"] == "F1-N" and g["A1"]["합산여부"] == "Y" and g["A4"]["프로젝트ID"] == "F1-N", sorted(g)
    assert g["B1"]["프로젝트ID"] == "F2-N-01" and g["B1"]["연결강도"] == "확실", g["B1"].to_dict()
    assert g["B2"]["프로젝트ID"] == "F2-N-02" and "사업 이름" in g["B2"]["연결근거"], "두 프로젝트의 사업 이름이 다 들어 있으면 더 긴(구체적인) 쪽"
    # 어절 사이에 다른 말이 낀 이름은 '추정' 이지만, 공고명에 그 프로젝트의 사업 이름('다온 복합 문화 체육 회관 건립')이 통째로 있으면 더한다(2026-11 실자료 점검. 사업 이름이 없으면 N — test_field_audit ⑥)
    assert "B3" not in g and g["D1"]["연결강도"] == "추정" and g["D1"]["합산여부"] == "Y" and g["D1"]["프로젝트ID"] == "F4-N", g["D1"].to_dict()
    assert g["C1"]["프로젝트ID"] == "F3-R-01" and g["C3"]["프로젝트ID"] == "F3-R-02", (g["C1"]["연결근거"], g["C3"]["연결근거"])
    assert g["C2"]["프로젝트ID"] == "" and g["C2"]["프로젝트_후보"] == "F3-R-01, F3-R-02" and g["C2"]["합산여부"] == "N", "좁은 창에 둘 다 들면 정하지 않는다"
    assert g["E1"]["합산여부"] == g["E2"]["합산여부"] == "Y" and g["E1"]["프로젝트ID"] == g["E2"]["프로젝트ID"] == "F5-R", (g["E1"].to_dict(), g["E2"].to_dict())
    assert g["E3"]["합산여부"] == "N" and g["E3"]["합산제외_까닭"].startswith("연간·단가"), g["E3"].to_dict()
    assert g["E4"]["합산여부"] == "N" and g["E4"]["합산제외_까닭"].startswith("연간·단가"), "공사 이름이 통째로 들어 있지 않은 공고의 '2구역'"
    assert servc._unit_words_from("상반기 가온 리모델링공사 감리", ["2025년상반기가온리모델링공사"]) and not servc._unit_words_from("가온 리모델링공사 감리", ["x"])
    # assign_project 를 직접: 사업유형이 다르면 하나뿐인 프로젝트에도 붙이지 않는다 / 시설 이름뿐인 사업 이름은 근거가 아니다
    ts = pd.Timestamp
    p_new = {"pid": "P-N", "type": "신축", "first": ts("2024-03-01"), "last": ts("2024-05-01"), "cores": ["가온회관건립"], "amt": 1e9, "keys": {"X"}}
    p_rem = {"pid": "P-R", "type": "리모델링", "first": ts("2026-03-01"), "last": ts("2026-03-01"), "cores": ["가온회관객석개보수"], "amt": 1e8, "keys": {"Y"}}
    ap = servc.assign_project
    assert ap("설계", "가온회관 리모델링 설계용역", ts("2023-01-01"), [p_new], "리모델링")[0] == "" and ap("설계", "가온회관 건립 설계공모", ts("2022-01-01"), [p_new], "신축")[0] == "P-N"
    assert ap("설계", "가온회관 설계공모", ts("2022-06-01"), [p_new, p_rem], "미분류")[:1] == ("P-N",), "사업유형 말이 없으면 좁은 창에 드는 하나(2026 공사는 45개월 뒤)"
    assert ap("건설사업관리·감리", "가온회관 건립공사 건설사업관리용역", ts("2024-02-01"), [p_new, p_rem], "신축")[3] > 3
    assert ap("폐기물처리", "가온회관 건립공사 폐기물 처리용역", ts("2030-01-01"), [p_new, p_rem], "신축")[0] == "", "창 밖"
    assert servc.project_core("가온회관 건립공사(건축) 긴급 입찰공고") == servc.project_core("(가칭)가온회관 건립 전기공사 감리용역") == "가온회관건립"


def _recalc(path: str):
    soffice = shutil.which("soffice") or shutil.which("libreoffice")
    if not soffice:
        return None
    outdir = os.path.join(os.path.dirname(path), "calc")
    subprocess.run([soffice, "--headless", "--convert-to", "xlsx", "--outdir", outdir, os.path.abspath(path)], capture_output=True, timeout=300)
    out = os.path.join(outdir, os.path.basename(path))
    return out if os.path.exists(out) else None


def _eval_done(wb, ws5, formula, rows=None) -> str:
    """05 '발주완료_폐기물용역' 수식을 직접 계산한다 — LibreOffice 가 없는 PC 에서도 값이 맞는지 보려고. 수식의 꼴(기본 = 프로젝트ID·용역종류·합산여부 조건):
    =IF(COUNTIFS(기본)=0,"",IF(COUNTIFS(기본,일수 범위,"<183")=COUNTIFS(기본),"발주완료(최근)","발주완료"))
    rows 를 주면 조건 범위가 08 의 자료 줄 전체(첫 줄, 마지막 줄)인지도 본다."""
    import re
    m = re.fullmatch(r'=IF\(COUNTIFS\((.+?)\)=0,"",IF\(COUNTIFS\((.+?)\)=COUNTIFS\((.+?)\),"발주완료\(최근\)","발주완료"\)\)', str(formula))
    assert m and m.group(1) == m.group(3) and m.group(2).startswith(m.group(1) + ","), ("세 COUNTIFS 의 기본 조건이 같아야 한다", formula)
    age = re.fullmatch(r"'([^']+)'!\$([A-Z]+)\$\d+:\$[A-Z]+\$\d+,\"<(\d+)\"", m.group(2)[len(m.group(1)) + 1:])
    assert age and int(age.group(3)) == build_excel.SERVC_RECENT_DAYS and wb[age.group(1)][f"{age.group(2)}1"].value == build_excel.SERVC_AGE_COL, \
        ("'최근' 조건은 08 의 일수 열이 기준 날수 미만인지 하나뿐이어야 한다", formula)
    txt = lambda v: "" if v is None else str(v).casefold()
    spans = set()

    def count(text) -> int:
        args = re.findall(r"""'[^']+'!\$[A-Z]+\$\d+:\$[A-Z]+\$\d+|\$[A-Z]+\d+|"[^"]*\"""", text)
        assert len(args) % 2 == 0 and ",".join(args) == text, (args, formula)
        conds = []
        for rng, crit in zip(args[0::2], args[1::2]):
            sheet, c1, r1, c2, r2 = re.fullmatch(r"'([^']+)'!\$([A-Z]+)\$(\d+):\$([A-Z]+)\$(\d+)", rng).groups()
            assert c1 == c2, formula
            spans.add((int(r1), int(r2)))
            want = crit[1:-1] if crit.startswith('"') else ws5[crit.replace("$", "")].value
            # Excel 의 COUNTIFS 와 같게: 빈 조건 칸은 빈 줄과 맞지 않고, 글자 조건은 대소문자를 가리지 않는다. '<숫자' 조건은 숫자 칸만 센다(빈칸·글자 칸은 세지 않는다).
            # 와일드카드와 그 밖의 비교 연산자 조건은 이 도우미가 계산하지 않는다
            assert crit.startswith('"') or want not in (None, ""), ("조건 칸이 비어 있다", crit, formula)
            less = re.fullmatch(r"<(\d+)", str(want))
            assert less or not re.search(r"[*?~]|^[<>=]", str(want)), ("와일드카드·비교 연산자 조건은 계산하지 못한다", crit, formula)
            if less:
                vals = [wb[sheet][f"{c1}{r}"].value for r in range(int(r1), int(r2) + 1)]
                assert all(v is None or type(v) in (int, float) or (isinstance(v, str) and not v.startswith("=")) for v in vals), \
                    ("일수 칸의 수식·날짜 값은 이 도우미가 계산하지 못한다(Excel 은 그 결과를 숫자로 센다)", formula)
                conds.append(lambda r, ws=wb[sheet], c=c1, lim=int(less.group(1)): type(ws[f"{c}{r}"].value) in (int, float) and ws[f"{c}{r}"].value < lim)
            else:
                conds.append(lambda r, ws=wb[sheet], c=c1, w="" if want is None else str(want).casefold(): txt(ws[f"{c}{r}"].value) == w)
        assert len(spans) == 1, ("조건 범위의 행이 서로 같아야 한다", spans, formula)
        (a, b), = spans
        return sum(all(f(r) for f in conds) for r in range(a, b + 1))
    n, recent = count(m.group(1)), count(m.group(2))
    assert rows is None or spans == {rows}, ("조건 범위는 08 의 자료 줄 전체(2행~마지막 자료 줄)여야 한다", spans, rows, formula)
    return "" if n == 0 else ("발주완료(최근)" if recent == n else "발주완료")


def test_excel_and_stage(tmp):
    import openpyxl
    import tests.test_dedup_and_excel as T
    cfg = dict(load_config())
    reviewed, hist, latest, linked = test_link()
    fac = pipeline.enrich_facility(dedup.facility_summary(latest, hist), latest, T.DOCS)
    trade = pipeline.assemble_trade_table(latest, T.DOCS)
    n_formula = lambda wb: sum(1 for ws in wb.worksheets for row in ws.iter_rows() for c in row if isinstance(c.value, str) and c.value.startswith("="))
    p0, p1 = os.path.join(tmp, "sv0.xlsx"), os.path.join(tmp, "sv1.xlsx")
    build_excel.build_workbook(p0, fac, hist, latest, trade, pd.DataFrame(), pd.DataFrame(), meta={"period": "x"})
    build_excel.build_workbook(p1, fac, hist, latest, trade, pd.DataFrame(), pd.DataFrame(), meta={"period": "x"}, servc=linked)
    w0, w1 = openpyxl.load_workbook(p0), openpyxl.load_workbook(p1)
    assert build_excel.S_SERVC not in w0.sheetnames and w1.sheetnames[-1] == build_excel.S_SERVC and w1.sheetnames[:-1] == w0.sheetnames, "용역 자료가 없으면 예전 그대로"
    h0, h1 = [c.value for c in w0[build_excel.S_FAC][1]], [c.value for c in w1[build_excel.S_FAC][1]]
    assert not any(str(h).startswith("용역_") for h in h0) and [h for h in h1 if h not in h0] == build_excel.SERVC5_COLS
    assert h1.index("문서배분_중복경고(수식)") < h1.index("용역_설계_추정가격(수식)") < h1.index("시설간_중복") and h1[-1] == "메모"
    assert n_formula(w1) - n_formula(w0) == len(fac) * (len(build_excel.SERVC5_COLS) + 1) and len(build_excel.SERVC5_COLS) == 5, \
        "05 에 프로젝트마다 용역 수식(금액 3 + 공고수 + 발주완료), 그리고 실제 누락 수별 시트(05-k)의 그 프로젝트 줄에 발주완료를 읽는 칸 하나만 늘어난다"
    miss_names = [build_excel.S_MISS_SUM] + [build_excel.S_MISS_N.format(k=k) for k in range(5)]
    assert all(build_excel.SERVC_DONE_COL in [c.value for c in w1[s][1]] and build_excel.SERVC_DONE_COL not in [c.value for c in w0[s][1]] for s in miss_names[1:])
    ws8 = w1[build_excel.S_SERVC]
    assert [c.value for c in ws8[1]] == build_excel.SERVC8_COLS and ws8.max_row >= len(linked) + 1
    i8 = build_excel.SERVC8_COLS.index(build_excel.SERVC_AGE_COL)
    assert build_excel.SERVC8_COLS[i8 - 1] == "공사대비_개월" and build_excel.SERVC8_COLS[:i8] + build_excel.SERVC8_COLS[i8 + 1:] == build_excel.SERVC_COLS, "08 = 연결 표의 열 + 일수 열 하나"
    # 공사 공고 수집 종료일 = 워크북에 실린 공사 공고(02 시트 = hist)의 가장 늦은 공고일. config 의 수집 기간이나 오늘 날짜가 아니다
    day = lambda v: pd.Timestamp(str(v)[:10])
    end0 = max(day(v) for v in hist["공고일시"])
    assert [ws8.cell(r, i8 + 1).value for r in range(2, len(linked) + 2)] == [int((end0 - day(v)).days) for v in linked["공고일시"]], "08 의 일수 = 실린 공사 공고의 가장 늦은 공고일 − 용역 공고일"
    note_of = lambda ws, start: next(str(r[0].value) for r in ws.iter_rows(min_row=start) if str(r[0].value or "").startswith("주:"))
    texts_of = lambda wb: (" ".join(str(r[1].value) for r in wb["README"].iter_rows() if r[0].value == "용역(08·05)"), note_of(wb[build_excel.S_FAC], 3), note_of(wb[build_excel.S_SERVC], len(linked) + 2))
    assert all(f"공사 공고 수집 종료일({end0.date().isoformat()})" in t and "공사 공고 수집 종료일 = 이 워크북에 실린 공사 공고 가운데 가장 늦은 공고일(02 시트)" in t and "읽지 못해" not in t
               for t in texts_of(w1)), [t[-500:] for t in texts_of(w1)]
    assert [r[0].value for r in w1["README"].iter_rows() if r[0].value in (build_excel.S_SERVC, "용역(08·05)")] == [build_excel.S_SERVC, "용역(08·05)"]
    r5 = {ws.cell(r, 1).value: r for ws in [w1[build_excel.S_FAC]] for r in range(2, ws.max_row + 1)}
    cell = w1[build_excel.S_FAC].cell(r5[linked["프로젝트ID"].iloc[-1] or list(r5)[0]], h1.index("용역_설계_추정가격(수식)") + 1).value
    assert cell.startswith("=IF(COUNTIFS('08_용역공고(참고)'!$A$2:$A$") and '"설계"' in cell and ',"Y")=0,"",SUMIFS(' in cell, cell
    # 발주완료_폐기물용역(보조지표): 더한(합산여부 Y) 폐기물처리 용역이 있는 프로젝트만 '발주완료', 나머지는 빈칸('미완료' 라고 적지 않는다)
    pids = set(fac["프로젝트ID"].astype(str))

    def done_values(path):
        wb = openpyxl.load_workbook(path)
        ws5, ws8 = wb[build_excel.S_FAC], wb[build_excel.S_SERVC]
        ci = [c.value for c in ws5[1]].index(build_excel.SERVC_DONE_COL) + 1
        no = [c.value for c in ws8[1]].index("공고번호") + 1
        last = max(r for r in range(2, ws8.max_row + 1) if ws8.cell(r, no).value)          # 표 아래 주석 줄에는 공고번호가 없다
        return {ws5.cell(r, 1).value: _eval_done(wb, ws5, ws5.cell(r, ci).value, rows=(2, last)) for r in range(2, ws5.max_row + 1) if ws5.cell(r, 1).value in pids}
    assert h1.index(build_excel.SERVC_DONE_COL) == h1.index("용역_공고수(수식)") + 1
    yes = linked[linked["합산여부"] == "Y"]
    waste_y = set(yes[yes["용역종류"] == "폐기물처리"]["프로젝트ID"]) - {""}
    other_y = (set(yes[yes["용역종류"] != "폐기물처리"]["프로젝트ID"]) - {""} - waste_y) & pids
    got = done_values(p1)
    assert set(got) == pids and waste_y and waste_y <= pids and {p for p, v in got.items() if v} == waste_y, (got, waste_y)
    assert other_y and all(got[p] == "" for p in other_y), "설계·건설사업관리 용역만 더해진 프로젝트는 발주완료가 아니다"
    # 발주완료(최근): 그 프로젝트의 폐기물처리 Y 줄이 모두 공사 공고 수집 종료일에서 6개월(183일) 안쪽이거나 그 뒤일 때만. 가장 이른 Y 줄이 183일 이상 앞이면 '발주완료'
    wy = linked[(linked["용역종류"] == "폐기물처리") & (linked["합산여부"] == "Y") & linked["프로젝트ID"].isin(waste_y)]
    wp = sorted(waste_y)[0]
    wdays = sorted(day(v) for v in wy[wy["프로젝트ID"] == wp]["공고일시"])
    assert len(wdays) >= 2 and wdays[0] < wdays[-1] and build_excel.SERVC_RECENT_DAYS == 183 and build_excel.SERVC_DONE_RECENT == "발주완료(최근)"
    assert (end0 - wdays[0]).days >= 183 and got[wp] == "발주완료", ("가장 이른 폐기물 Y 줄이 마지막 공사 공고일의 183일 넘게 앞", got)

    def hist_until(end):
        """공사 공고의 가장 늦은 공고일이 end 인 이력: 그 뒤의 공고일은 end 로 당기고, 한 줄은 그 날 밤(23:30)으로 둔다 — 기준일은 시각을 버린 날짜다. (이 표는 02 시트와 기준일에만 쓰인다)"""
        h = hist.copy()
        cut = pd.Timestamp(end) + pd.Timedelta(hours=9)
        h["공고일시"] = [min(pd.Timestamp(str(v)), cut).strftime("%Y-%m-%d %H:%M") for v in h["공고일시"]]
        h.loc[h.index[-1], "공고일시"] = (pd.Timestamp(end) + pd.Timedelta(hours=23, minutes=30)).strftime("%Y-%m-%d %H:%M")
        return h
    with_end = lambda name, end, sv=linked: build_excel.build_workbook(os.path.join(tmp, name), fac, hist_until(end), latest, trade, pd.DataFrame(), pd.DataFrame(),
                                                                       meta={"period": "x"}, servc=sv) or os.path.join(tmp, name)
    iso = lambda t: t.date().isoformat()
    p_old = with_end("sv5.xlsx", iso(wdays[0] + pd.Timedelta(days=183)))          # 가장 이른 줄이 꼭 183일 앞(나머지 줄은 그보다 가깝거나 종료일 뒤) → 최근이 아니다
    p_new = with_end("sv6.xlsx", iso(wdays[0] + pd.Timedelta(days=182)))          # 가장 이른 줄도 182일 앞 → 모두 최근
    p_aft = with_end("sv7.xlsx", iso(wdays[0] - pd.Timedelta(days=30)))           # 폐기물 용역 공고가 모두 공사 공고 수집 종료일 뒤(음수) → 최근
    assert done_values(p_old)[wp] == "발주완료" and done_values(p_new)[wp] == "발주완료(최근)" and done_values(p_aft)[wp] == "발주완료(최근)", \
        (done_values(p_old), done_values(p_new), done_values(p_aft))
    assert all({p for p, v in done_values(p).items() if v} == waste_y for p in (p_old, p_new, p_aft)), "'최근' 을 가려도 표시가 붙는 프로젝트는 그대로다"
    wbo = openpyxl.load_workbook(p_old)
    ws8o, ws5o = wbo[build_excel.S_SERVC], wbo[build_excel.S_FAC]
    h8o = [c.value for c in ws8o[1]]
    c8 = lambda name: h8o.index(name) + 1
    end_old = wdays[0] + pd.Timedelta(days=183)
    assert all(f"공사 공고 수집 종료일({iso(end_old)})" in t for t in texts_of(wbo)), "가장 늦은 공사 공고가 밤(23:30)에 났어도 설명에 찍히는 날짜는 그 날(02 시트의 가장 늦은 공고일)"
    ages = [ws8o.cell(r, c8(build_excel.SERVC_AGE_COL)).value for r in range(2, len(linked) + 2)]
    assert ages == [int((end_old - day(v)).days) for v in linked["공고일시"]] and all(type(a) is int for a in ages), "일수 = 종료일 − 용역 공고일(날짜끼리, 모든 줄에 숫자로)"
    assert min(ages) < 0 and 183 in ages, "종료일 뒤에 난 공고는 음수, 가장 이른 폐기물 Y 줄은 꼭 183"
    assert ws8o.cell(2, c8("공사비대비_비율")).number_format == "0.0%" and ws8o.cell(2, c8(build_excel.SERVC_AGE_COL)).number_format != "0.0%", "일수 열이 끼어도 비율 서식은 비율 열에"
    age2 = ws8o.cell(2, c8(build_excel.SERVC_AGE_COL))
    assert age2.font.color is not None and str(age2.font.color.rgb).endswith("0000FF") and age2.number_format == "General" \
        and not str(age2.fill.fgColor.rgb).endswith("FFFF00"), "일수 열도 값 열이다: 파란 글자·일반 서식, 노란(고치는) 칸이 아니다"
    h2o = [c.value for c in wbo["02_공고목록_전체"][1]]
    assert max(day(r[h2o.index("공고일시")].value) for r in wbo["02_공고목록_전체"].iter_rows(min_row=2) if r[h2o.index("공고일시")].value) == end_old, "기준일은 02 시트에서 확인할 수 있다"
    cio = [c.value for c in ws5o[1]].index(build_excel.SERVC_DONE_COL) + 1
    evo = lambda: {ws5o.cell(r, 1).value: _eval_done(wbo, ws5o, ws5o.cell(r, cio).value) for r in range(2, ws5o.max_row + 1) if ws5o.cell(r, 1).value in pids}[wp]
    row_of = lambda no: next(r for r in range(2, len(linked) + 2) if ws8o.cell(r, c8("공고번호")).value == no)
    first_no = wy[(wy["프로젝트ID"] == wp) & (wy["공고일시"].map(day) == wdays[0])]["공고번호"].iloc[0]
    assert evo() == "발주완료"
    ws8o.cell(row_of(first_no), c8("합산여부")).value = "N"                      # 08 에서 183일 앞의 줄을 N 으로 바꾸면 남은 Y 줄은 모두 최근
    assert evo() == "발주완료(최근)"
    ws8o.cell(row_of(first_no), c8("합산여부")).value = "Y"
    ws8o.cell(row_of(first_no), c8(build_excel.SERVC_AGE_COL)).value = 182     # 경계: 182 는 최근, 183 은 아니다
    assert evo() == "발주완료(최근)"
    ws8o.cell(row_of(first_no), c8(build_excel.SERVC_AGE_COL)).value = None    # 일수 칸이 빈 줄(끼워 넣은 줄)은 '최근' 으로 치지 않는다
    assert evo() == "발주완료"
    ws8o.cell(row_of(first_no), c8(build_excel.SERVC_AGE_COL)).value = "10"    # 글자로 적힌 숫자도 Excel 의 '<183' 조건에 세어지지 않는다
    assert evo() == "발주완료"
    bad = linked.copy()
    bad.loc[bad["공고번호"] == first_no, "공고일시"] = "알 수 없음"                  # 공고일시를 읽지 못한 줄은 일수가 비고, 그 줄 때문에 '최근' 이 되지는 않는다
    assert done_values(with_end("sv8.xlsx", iso(wdays[0] - pd.Timedelta(days=30)), bad))[wp] == "발주완료"
    ws8b = openpyxl.load_workbook(os.path.join(tmp, "sv8.xlsx"))[build_excel.S_SERVC]
    assert [ws8b.cell(r, c8(build_excel.SERVC_AGE_COL)).value for r in range(2, len(bad) + 2)] == \
        [None if no == first_no else int((wdays[0] - pd.Timedelta(days=30) - day(v)).days) for no, v in zip(bad["공고번호"], bad["공고일시"])], \
        "공고일시를 읽지 못한 줄의 일수 칸만 빈칸(다른 값으로 채우지 않는다), 나머지 줄은 그대로"
    mix = linked.copy()
    mix.index = range(100, 100 + len(mix))                                         # 인덱스가 0 부터가 아닌 표여도 일수가 제 줄에 들어간다
    mix.loc[mix["공고번호"] == first_no, "공고일시"] = wdays[0].strftime("%Y%m%d") + "2359"   # 원문 형식이 섞여 있어도(dedup._parse_dt), 밤늦게 난 공고여도 같은 날수(날짜끼리 뺀다 — 반올림하지 않는다)
    ws8m = openpyxl.load_workbook(with_end("sv9.xlsx", iso(end_old), mix))[build_excel.S_SERVC]
    assert [ws8m.cell(r, c8(build_excel.SERVC_AGE_COL)).value for r in range(2, len(mix) + 2)] == ages, "형식이 다른 공고일시·오후에 난 공고·0 부터가 아닌 인덱스"
    # 실린 공사 공고에서 공고일을 하나도 읽지 못하면(또는 공사 공고가 없으면) 일수 칸을 비우고 '최근' 을 가리지 않는다 — 설명도 날짜 없는 글로('종료일()' 같은 빈 괄호가 찍히지 않는다)
    # 가장 늦은 공고일은 글자 순서나 줄 위치가 아니라 읽은 날짜로 고른다: 형식이 섞이고(가장 늦은 줄만 원래 형식) 줄 순서가 뒤집힌 이력에서도 같은 기준일
    hmix = hist.assign(공고일시=[v if day(v) == end0 else pd.Timestamp(str(v)).strftime("%Y%m%d%H%M") for v in hist["공고일시"]]).iloc[::-1]
    build_excel.build_workbook(os.path.join(tmp, "sv13.xlsx"), fac, hmix, latest, trade, pd.DataFrame(), pd.DataFrame(), meta={"period": "x"}, servc=linked)
    ws13 = openpyxl.load_workbook(os.path.join(tmp, "sv13.xlsx"))[build_excel.S_SERVC]
    assert [ws13.cell(r, i8 + 1).value for r in range(2, len(linked) + 2)] == [int((end0 - day(v)).days) for v in linked["공고일시"]], "형식이 섞이고 순서가 뒤집힌 이력"
    for name, h in (("sv10.xlsx", hist.assign(공고일시="")), ("sv11.xlsx", hist.iloc[0:0]), ("sv12.xlsx", hist.drop(columns=["공고일시"]))):
        pn = os.path.join(tmp, name)
        build_excel.build_workbook(pn, fac, h, latest, trade, pd.DataFrame(), pd.DataFrame(), meta={"period": "x"}, servc=linked)
        wn = openpyxl.load_workbook(pn)
        assert {wn[build_excel.S_SERVC].cell(r, i8 + 1).value for r in range(2, len(linked) + 2)} == {None}, name
        gn = done_values(pn)
        assert {p for p, v in gn.items() if v} == waste_y and set(gn.values()) == {"", "발주완료"}, (name, gn)
        tn = texts_of(wn)
        assert all("종료일()" not in t for t in tn) and all("읽지 못해" in t and "'발주완료(최근)'" in t for t in tn[1:]) and "공고가 더 나올 수 있다" in tn[0], (name, [t[-400:] for t in tn])
    # 폐기물 용역이 N 뿐이면(취소·재공고 이전·출처 모름) 빈칸이고, '기타' 용역을 Y 로 바꿔도 빈칸
    alt = linked.copy()
    alt.loc[alt["용역종류"] == "폐기물처리", "합산여부"] = "N"
    alt.loc[(alt["용역종류"] == "기타") & alt["프로젝트ID"].isin(waste_y), "합산여부"] = "Y"
    assert ((alt["용역종류"] == "기타") & (alt["합산여부"] == "Y")).any() and ((alt["용역종류"] == "폐기물처리") & alt["프로젝트ID"].isin(waste_y)).any()
    p2 = os.path.join(tmp, "sv2.xlsx")
    build_excel.build_workbook(p2, fac, hist, latest, trade, pd.DataFrame(), pd.DataFrame(), meta={"period": "x"}, servc=alt)
    assert set(done_values(p2).values()) == {""}, done_values(p2)
    # 폐기물처리 줄의 프로젝트ID 가 다른 프로젝트로 바뀌고 Y 가 되면 그 프로젝트가 발주완료로 바뀐다
    alt2 = alt.copy()
    tgt = sorted(other_y)[0]
    i = alt2.index[alt2["용역종류"] == "폐기물처리"][0]
    alt2.loc[i, ["프로젝트ID", "합산여부"]] = [tgt, "Y"]
    p3 = os.path.join(tmp, "sv3.xlsx")
    build_excel.build_workbook(p3, fac, hist, latest, trade, pd.DataFrame(), pd.DataFrame(), meta={"period": "x"}, servc=alt2)
    assert {p for p, v in done_values(p3).items() if v} == {tgt}, done_values(p3)
    # 08 을 손으로 고치는 경우를 그대로: 프로젝트ID 가 빈(시설까지만 이어진) 폐기물처리 줄 — 합산여부만 Y 로 바꿔서는 안 되고 프로젝트ID 도 적어야 그 프로젝트가 발주완료
    alt3 = alt.copy()
    alt3.loc[i, ["프로젝트ID", "합산여부"]] = ["", "N"]
    p4 = os.path.join(tmp, "sv4.xlsx")
    build_excel.build_workbook(p4, fac, hist, latest, trade, pd.DataFrame(), pd.DataFrame(), meta={"period": "x"}, servc=alt3)
    wb4 = openpyxl.load_workbook(p4)
    ws8h, ws5h = wb4[build_excel.S_SERVC], wb4[build_excel.S_FAC]
    h8h, row8 = [c.value for c in ws8h[1]], 2 + list(alt3.index).index(i)
    assert ws8h.cell(row8, h8h.index("공고번호") + 1).value == alt3.loc[i, "공고번호"] and not ws8h.cell(row8, h8h.index("프로젝트ID") + 1).value
    cih = [c.value for c in ws5h[1]].index(build_excel.SERVC_DONE_COL) + 1
    ev = lambda: {ws5h.cell(r, 1).value: _eval_done(wb4, ws5h, ws5h.cell(r, cih).value) for r in range(2, ws5h.max_row + 1) if ws5h.cell(r, 1).value in pids}
    assert set(ev().values()) == {""}, ev()
    ws8h.cell(row8, h8h.index("합산여부") + 1).value = "Y"                      # 합산여부만 Y: 아직 어느 프로젝트의 것도 아니다
    assert set(ev().values()) == {""}, ev()
    ws8h.cell(row8, h8h.index("프로젝트ID") + 1).value = tgt                    # 프로젝트ID 를 적으면 그 프로젝트만 발주완료
    assert {p for p, v in ev().items() if v} == {tgt}, ev()
    ws8h.cell(row8, h8h.index("합산여부") + 1).value = "y"                      # 소문자로 적어도 Excel 은 센다
    assert {p for p, v in ev().items() if v} == {tgt}, ev()
    calc = _recalc(p1)
    if calc:                                           # LibreOffice 가 있으면 다시 계산한 값을 판다스 합과 견준다
        wc = openpyxl.load_workbook(calc, data_only=True)
        errs = [(ws.title, c.coordinate) for ws in wc.worksheets for row in ws.iter_rows() for c in row if isinstance(c.value, str) and c.value.startswith("#")]
        assert not errs, errs[:5]
        ws = wc[build_excel.S_FAC]
        y = linked[linked["합산여부"] == "Y"]
        want = y.groupby(["프로젝트ID", "용역종류"])["추정가격"].sum().to_dict()
        tot0 = {}
        c0 = _recalc(p0)
        if c0:
            w0c = openpyxl.load_workbook(c0, data_only=True)[build_excel.S_FAC]
            tot0 = {w0c.cell(r, 1).value: w0c.cell(r, h0.index("총공사비합계_기초금액기준(수식)") + 1).value for r in range(2, w0c.max_row + 1)}
        for r in range(2, ws.max_row + 1):
            pid = ws.cell(r, 1).value
            if pid not in set(fac["프로젝트ID"].astype(str)):       # 표 아래의 주석 행
                continue
            for k in build_excel.SERVC_KINDS:
                v = ws.cell(r, h1.index(f"용역_{k}_추정가격(수식)") + 1).value
                assert (v in (None, "")) if (pid, k) not in want else float(v) == float(want[(pid, k)]), (pid, k, v, want.get((pid, k)))
            assert int(ws.cell(r, h1.index("용역_공고수(수식)") + 1).value or 0) == int((y["프로젝트ID"] == pid).sum())
            assert (ws.cell(r, h1.index(build_excel.SERVC_DONE_COL) + 1).value or "") == ("발주완료" if (pid, "폐기물처리") in want else ""), pid
            if tot0:
                assert ws.cell(r, h1.index("총공사비합계_기초금액기준(수식)") + 1).value == tot0[pid], "용역 금액은 총공사비 합계에 들어가지 않는다"
        for pth in (p_new, p_aft):                     # '발주완료(최근)' 도 다시 계산한 값이 위의 직접 계산과 같아야 한다(p_old 는 위에서 08 을 고치지 않은 파일 그대로)
            cq = _recalc(pth)
            if cq:
                wq = openpyxl.load_workbook(cq, data_only=True)[build_excel.S_FAC]
                gotq = {wq.cell(r, 1).value: (wq.cell(r, h1.index(build_excel.SERVC_DONE_COL) + 1).value or "") for r in range(2, wq.max_row + 1) if wq.cell(r, 1).value in pids}
                assert gotq == done_values(pth), (gotq, done_values(pth))
        print("용역 참고 열 값 검증 OK(LibreOffice 재계산)")
    else:
        print("(LibreOffice 없음 → 용역 참고 열의 값 검증 생략)")

    # ── 단계 실행: link-servc → excel ──
    cfg["paths"] = dict(cfg["paths"], data_dir=os.path.join(tmp, "sv", "data"), out_dir=os.path.join(tmp, "sv", "out"))
    os.makedirs(cfg["paths"]["data_dir"], exist_ok=True)
    os.makedirs(cfg["paths"]["out_dir"], exist_ok=True)
    data = lambda n: os.path.join(cfg["paths"]["data_dir"], n)

    def run(fn):
        buf = io.StringIO()
        with contextlib.redirect_stdout(buf):
            try:
                fn(cfg)
            except SystemExit as e:
                buf.write(f"\n[SystemExit] {e}")
        return buf.getvalue()
    out = run(pipeline.stage_link_servc)
    assert "collect --servc" in out and "[SystemExit]" in out, "용역 자료가 없으면 무엇을 먼저 할지 알린다"
    raw = _rows(cfg, [dict(no="S01", name="가상시립미술관 건립공사 건설사업관리용역", dt="2024-01-20 10:00:00", price=900_000_000),
                      dict(no="S05", name="가상시립미술관 생활폐기물 수집운반 대행용역", dt="2024-05-02 10:00:00", price=30_000_000)])
    raw["쓰지않는열"] = "x"
    raw.to_parquet(data("servc_all.parquet"), index=False)
    assert "dedup 을 먼저" in run(pipeline.stage_link_servc)
    hist.to_parquet(data("notices_hist.parquet"), index=False)
    latest.to_parquet(data("notices_latest.parquet"), index=False)
    assert "검수 파일이 비어 있거나 없습니다" in run(pipeline.stage_link_servc)
    agg = discover.discover_candidates(discover.standardize(pd.DataFrame(T.RAW), cfg, pd.DataFrame(T.BSIS)))
    discover.write_review_workbook(agg, os.path.join(cfg["paths"]["out_dir"], "facility_candidates.xlsx"))
    out = run(pipeline.stage_link_servc)
    sl = pd.read_parquet(data("servc_linked.parquet"))
    assert len(sl) == 2 and list(sl.columns) == servc.LINK_COLS and set(sl["합산여부"]) == {"Y", "N"}, (out, sl)
    assert "[진단] 용역 공고 2줄" in out and "합산여부 Y): 1줄" in out and "python -m g2b_costdb.pipeline excel" in out, out
    assert open(os.path.join(cfg["paths"]["out_dir"], "servc_link_summary.txt"), encoding="utf-8-sig").read().strip() in out
    assert pd.read_parquet(data("logs_servc.parquet")).empty and "—" not in out, "화면 글에 cp949 로 찍히지 않는 '—' 가 없어야 한다"
    out = run(pipeline.stage_excel)
    wb = openpyxl.load_workbook(os.path.join(cfg["paths"]["out_dir"], cfg["paths"]["excel_name"]))
    assert build_excel.S_SERVC in wb.sheetnames and wb[build_excel.S_SERVC].max_row >= 3 and "용역 연결(08) 2줄" in out and "[주의] dedup 을 link-servc 보다" not in out, out
    # dedup 을 다시 돌린 뒤 link-servc 를 잊으면 알린다 / 연결 결과가 없으면 안내만 하고 예전 그대로 만든다
    future = time.time() + 120
    os.utime(data("notices_latest.parquet"), (future, future))
    assert "[주의] dedup 을 link-servc 보다 나중에 실행했습니다" in run(pipeline.stage_excel)
    os.remove(data("servc_linked.parquet"))
    out = run(pipeline.stage_excel)
    wb = openpyxl.load_workbook(os.path.join(cfg["paths"]["out_dir"], cfg["paths"]["excel_name"]))
    assert build_excel.S_SERVC not in wb.sheetnames and "[안내] 용역 공고를 수집했지만 아직 연결하지 않았습니다" in out and "link-servc" in out, out


def test_more_rules():
    """변이 시험에서 살아남은 규칙들(창·재공고 문턱·여러 시설에 걸친 공고·제외 규칙·표준화)을 하나씩 고정한다. 모든 이름은 가상."""
    cfg = load_config()
    f = cfg["fields"]
    k, same, ap, ts = servc.classify_kind, servc.inst_same, servc.assign_project, pd.Timestamp
    # ── 종류 판정의 남은 가지 ──
    assert k("◇ 청사 철거 및 폐기물 처리공사 감리용역")[0] == "건설사업관리·감리", "폐기물 말이 있어도 마지막 용역 이름이 감리면 감리"
    assert k("◇ 건립공사 폐기물 용역", "건설폐기물처리서비스")[0] == "폐기물처리" and k("◇ 건립공사 폐기물 용역")[0] == "기타", "처리·운반 말이 없으면 조달분류로"
    assert k("◇ 건립사업 건설사업관리 및 타당성 검토 용역")[0] == "기타" and k("◇ 건립공사 감리 적정성 검토 용역")[0] == "기타", "건설사업관리·감리의 부대 용역"
    assert k("◇ 건립 설계 제안공모")[:2] == ("설계", "설계공모") and k("◇ 신축공사 석면 폐기물 처리용역")[1] == "지정폐기물(석면)"
    assert k("◇ 건립 제안공모")[0] == "기타", "'제안공모' 만으로는 설계가 아니다(협상계약 방식의 다른 용역도 제안공모다)"
    # ── 기관: 첫 어절만 같은 경우 / 한쪽 이름이 다른 쪽에 든 경우 ──
    assert same("한국토지주택공사 경기남부지역본부", "한국토지주택공사 서울지역본부") and same("국립가온박물관", "문화체육관광부 국립가온박물관")
    assert not same("가온시설공단", "나래시설공단")
    # ── 넓은 창(WINDOWS): 종류마다 앞·뒤 경계의 안과 밖 ──
    p = {"pid": "P-N", "type": "신축", "first": ts("2024-03-01"), "last": ts("2024-09-01"), "cores": ["가온회관건립"], "amt": 1e9, "keys": {"X"}}
    one = lambda kind, dt, pp=p: ap(kind, "가온회관 용역", ts(dt), [pp], "")[0]
    assert one("설계", "2019-03-01") == "P-N" and one("설계", "2017-01-01") == "" and one("설계", "2025-02-01") == "P-N" and one("설계", "2025-09-01") == "", "설계: 84개월 전~6개월 뒤"
    assert one("건설사업관리·감리", "2021-03-01") == "P-N" and one("건설사업관리·감리", "2020-01-01") == "", "감리: 48개월 전"
    assert one("건설사업관리·감리", "2026-09-01") == "P-N" and one("건설사업관리·감리", "2027-12-01") == "", "감리: 공사 마지막 공고 36개월 뒤"
    assert one("폐기물처리", "2023-09-01") == "P-N" and one("폐기물처리", "2022-03-01") == "", "폐기물: 12개월 전"
    assert one("기타", "2026-09-01") == "P-N" and one("기타", "2018-03-01") == "P-N" and one("기타", "2029-01-01") == ""
    long_p = dict(p, first=ts("2020-01-01"), last=ts("2024-01-01"))
    assert one("건설사업관리·감리", "2025-06-01", long_p) == "P-N", "뒤쪽 경계는 공사 '마지막' 공고부터 잰다(첫 공고부터면 65개월)"
    nodate = dict(p, first=pd.NaT, last=pd.NaT)
    assert one("설계", "2010-01-01", nodate) == "P-N", "공사 공고일을 모르면 시기로 거르지 않는다"
    # ── 좁은 창(TIGHT): 사업 이름·사업유형으로 가릴 수 없는 프로젝트 둘 ──
    mk = lambda pid, first, last=None, cores=("x",), typ="신축": {"pid": pid, "type": typ, "first": ts(first), "last": ts(last or first), "cores": list(cores), "amt": 1e9, "keys": {pid}}
    two = lambda kind, dt, a, b, title="가온회관 용역", wt="": ap(kind, title, ts(dt), [a, b], wt)[0]
    assert two("설계", "2024-03-01", mk("A", "2023-12-01"), mk("B", "2024-06-01")) == "B", "설계: 공사 뒤에 난 공고는 그 공사의 설계가 아니다(좁은 창의 뒤쪽 0)"
    assert two("폐기물처리", "2024-07-01", mk("A", "2022-01-01"), mk("B", "2024-01-01")) == "B", "폐기물: 18개월 뒤까지"
    assert two("폐기물처리", "2023-06-01", mk("A", "2022-01-01"), mk("B", "2024-01-01")) == "A", "폐기물: 공사 1개월 전부터"
    assert two("폐기물처리", "2024-06-01", mk("A", "2022-01-01", "2023-06-01"), mk("B", "2025-03-01")) == "A", "폐기물의 좁은 창은 공사 마지막 공고부터 잰다"
    assert two("건설사업관리·감리", "2024-03-10", mk("A", "2024-03-01"), mk("B", "2024-03-05")) == "", "좁은 창에 둘이면 정하지 않는다"
    # ── ② 사업 이름이 통째로 든 프로젝트가 여럿: 긴 쪽 / 같은 이름이면 사업유형 ──
    assert two("설계", "2023-03-01", mk("A", "2024-03-01", cores=["가온회관건립"]), mk("B", "2024-03-01", cores=["가온회관건립경관조명조성"]),
               "가온회관 건립 경관조명 조성공사 실시설계용역") == "B", "두 사업 이름이 다 들어 있으면 더 긴(구체적인) 쪽"
    assert two("건설사업관리·감리", "2023-03-01", mk("A", "2024-03-01", cores=["가온회관시설개선"]), mk("B", "2024-06-01", cores=["가온회관시설개선"], typ="리모델링"),
               "가온회관 시설개선공사 건설사업관리용역", "리모델링") == "B", "같은 사업 이름이면 공고명의 사업유형으로(좁은 창 밖이어도)"
    assert ap("설계", "가온회관 증축 및 리모델링 설계용역", ts("2023-06-01"), [mk("R", "2024-03-01", typ="리모델링")], "증축·리모델링")[0] == "R", "'증축·리모델링' 은 둘 중 하나만 맞아도"
    # ── ④ 공고명 유사도: 0.75 이상이고 다음과 0.15 이상 벌어질 때만 ──
    big, small, park, other = (mk(i, "2024-03-01", cores=[c]) for i, c in (("BIG", "가온회관대공연장객석개보수"), ("SMALL", "가온회관소공연장객석개보수"),
                                                                        ("PARK", "가온회관주차장증설"), ("ETC", "무대기계교체")))
    r4 = ap("건설사업관리·감리", "가온회관 대공연장 객석 개선공사 감리용역", ts("2024-03-10"), [big, park], "")
    assert r4[0] == "BIG" and r4[3] == 1 and "가장 가까운" in r4[1], r4
    assert two("건설사업관리·감리", "2024-03-10", big, small, "가온회관 공연장 객석 개선 감리용역") == "", "비슷한 이름이 둘이면 정하지 않는다"
    assert two("건설사업관리·감리", "2024-03-10", big, other, "가온회관 대공연장 음향설비 감리용역") == "", "가장 가까워도 유사도 0.75 미만이면 정하지 않는다"

    # ── 연결 표: 가상 시설 여럿 ──
    base = dict(공고차수="000", 공고종류="일반", 추정가격=1_000_000_000, 공종="건축", 사업유형="신축", 수요기관="경기도 가온시", 세부사업명="", 수요기관코드="")
    H = lambda fid, name, pid, no, title, dt, **kw: dict(base, 시설ID=fid, 시설명=name, 프로젝트ID=pid, 공고번호=no, 공고명=title, 공고일시=dt + " 10:00:00", **kw)
    hist = pd.DataFrame([
        H("G1", "가온 체육관", "G1-N", "C01", "가온 체육관 건립공사", "2024-03-01"),
        H("G1", "가온 체육관", "G1-N", "C02", "가온 체육관 진입도로 개설공사", "2024-06-01", 추정가격=200_000_000, 공종="토목"),
        H("G2", "가온 체육관 수영장", "G2-N", "C03", "가온 체육관 수영장 건립공사", "2024-03-01"),
        H("G3", "다온 수영장", "G3-N", "C04", "다온 수영장 건립공사", "2024-04-01"),
        H("G4", "나래 문화센터", "G4-N", "C05", "나래 문화센터 공연장 건립공사", "2024-05-01"),          # 같은 공사 공고가 두 시설 후보에(05 '시설간_중복')
        H("G5", "나래 문화센터 공연장", "G5-N", "C05", "나래 문화센터 공연장 건립공사", "2024-05-01"),
        H("G6", "마루 복지관", "G6-R", "C06", "마루 복지관 LED 조명 교체공사", "2024-03-01", 사업유형="리모델링"),
        H("R1", "라온 도서관", "R1-N", "C07", "라온 도서관 건립공사", "2024-03-01", 추정가격=2_000_000_000, 수요기관코드="5350000"),
        H("I1", "가온 정보화도서관", "I1-N", "C08", "가온 정보화도서관 건립공사", "2024-03-01"),
        H("B1", "다온 국민체육센터", "B1-N", "C09", "다온 국민체육센터 건립공사", "2024-03-01"),
        H("Q1", "가온시 청사", "Q1-N", "C10", "청사 건립공사", "2024-03-01"),
        H("M1", "미르 문화회관", "M1-N", "C11", "미르 문화회관 건립공사", "2024-03-01"),
        H("U1", "라온 임대주택", "U1-R", "C12", "2025년 상반기 라온 임대주택 리모델링공사(1구역)", "2025-03-01", 사업유형="리모델링"),
        H("K1", "한울 회관", "K1-N-01", "C13", "한울 회관 건립공사", "2021-03-01"),
        H("K1", "한울 회관", "K1-N-02", "C14", "한울 회관 건립공사", "2024-03-01"),
    ])
    latest = hist.copy()
    rv = lambda fid, name, **kw: {"시설ID": fid, "검수_시설명": name, "검수_별칭(;구분)": "", "검수_수요기관": kw.get("rinst", ""), "수요기관": "경기도 가온시", "검색어": ""}
    reviewed = pd.DataFrame([rv(fid, name) for fid, name in hist.drop_duplicates("시설ID")[["시설ID", "시설명"]].itertuples(index=False) if fid != "M1"]
                            + [rv("M1", "미르 문화회관", rinst="가온문화재단")])
    assert "청사" not in servc.facility_keys(reviewed, hist)["Q1"]["pats"], "두 글자 이름은 검색 이름으로 쓰지 않는다"
    mk = lambda no, name, dt, **kw: dict(no=no, name=name, dt=(dt + " 10:00:00") if dt else "", inst=kw.pop("inst", "경기도 가온시"), **kw)
    raw = _rows(cfg, [
        # 한 공고가 시설 둘에: 서로 다른 사업 → 모두 N / 한쪽에 더 정확히 맞음 → 그쪽만 Y / 같은 공사 공고를 함께 가진 두 후보 → 모두 Y
        mk("MF1", "가온 체육관 다온 수영장 통합 설계공모", "2023-01-10", price=900_000_000),
        mk("MF2", "가온 체육관 수영장 건립공사 설계용역", "2023-02-10", price=300_000_000),
        mk("MF3", "나래 문화센터 공연장 건립공사 건설사업관리용역", "2024-04-10", price=500_000_000),
        # 제외 규칙과 그 예외(공고명에 DB 프로젝트의 사업 이름이 통째로 있으면 보지 않는다)
        mk("NB1", "가온 체육관 신축부지 관로 이설 설계용역", "2023-06-01", price=30_000_000),
        mk("NB2", "가온 체육관 진입도로 개설공사 감리용역", "2024-06-10", price=24_000_000),
        mk("MT1", "마루 복지관 LED 조명 교체공사 감리용역", "2024-03-10", price=10_000_000),
        mk("MT2", "마루 복지관 LED 조명 교체공사 폐기물 처리용역", "2024-04-10", price=5_000_000),
        mk("UN1", "2025년 상반기 라온 임대주택 감리용역", "2025-02-10", price=50_000_000),                      # 공사 이름이 통째로 있지 않으면 '상반기' 를 그대로 본다
        # 시설 이름 속의 말('정보화')은 종류 판정에서 가린다 — 띄어쓰기가 달라도
        mk("IT1", "가온 정보화도서관 건립공사 전기공사 감리용역", "2024-03-10", price=40_000_000),
        mk("IT2", "가온 정보화 도서관 건립 소방공사 감리용역", "2024-03-12", price=20_000_000),
        mk("BR1", "다온(국민체육)센터 건립공사 전기공사 감리용역", "2024-03-10", price=40_000_000),              # 괄호 기호만 빼면 시설 이름
        mk("QQ1", "신청사 건립공사 설계용역", "2023-03-10", price=40_000_000),
        mk("MR1", "미르 문화회관 건립공사 설계용역", "2023-03-10", price=40_000_000, inst="가온문화재단"),          # 검수_수요기관에 적은 기관은 '같은 발주 기관'
        mk("CD1", "라온 도서관 건립공사 방송설비 감리용역", "2024-03-11", price=15_000_000, inst="", code="5350000.0"),
        mk("KA1", "한울 회관 리모델링 설계용역", "2023-12-01", price=10_000_000),                               # 프로젝트 미정 → 가장 가까운 프로젝트 기준의 참고 열
        # 재공고 판정(라온 도서관 한 프로젝트, 공종 말로 이름을 달리함)
        mk("RB1", "라온 도서관 건립공사 전기공사 감리용역", "2024-02-01", price=100_000_000),
        mk("RB2", "라온 도서관 건립공사 전기공사 감리용역", "2024-04-01", price=150_000_000),                   # 60일 뒤, 금액이 달라도 재공고
        mk("RC1", "라온 도서관 건립공사 소방공사 감리용역", "2024-01-10", price=50_000_000),
        mk("RC2", "라온 도서관 건립공사 소방공사 감리용역", "2024-11-05", price=52_000_000),                    # 300일 뒤, 금액이 거의 같음 → 재공고
        mk("RD1", "라온 도서관 건립공사 정보통신공사 감리용역", "2023-06-01", price=30_000_000),
        mk("RD2", "라온 도서관 건립공사 정보통신공사 감리용역", "2025-05-01", price=30_000_000),                 # 700일 뒤 → 다른 용역(둘 다 더함)
        mk("RE1", "라온 도서관 건립 건설사업관리용역", "2023-12-01", price=800_000_000),
        mk("RE2", "라온 도서관 건립 건설사업관리용역", "2024-09-01", price=1_000_000_000, re="Y"),               # 275일 뒤, 금액은 다르지만 '재공고' 로 표시
        mk("RF1", "라온 도서관 건립 기본 및 실시설계용역", "2022-01-10"),
        mk("RF2", "라온 도서관 건립 기본 및 실시설계용역", "2022-10-01", price=400_000_000),                    # 앞 공고에 금액이 없으면 같은 금액으로 본다
        mk("RG1", "라온 도서관 건립 조경공사 감리용역", "", price=20_000_000),                                 # 공고일이 없는 쌍 → 한쪽만 남긴다
        mk("RG2", "라온 도서관 건립 조경공사 감리용역", "2024-05-01", price=20_000_000),
        mk("PV1", "라온 도서관 건립공사 기계설비 감리용역", "2024-05-01", price=20_000_000),
        mk("PV2", "라온 도서관 기계설비공사 감리 용역(명칭 조정)", "2024-06-10", price=21_000_000, prev="PV1"),  # 이름이 달라도 이전공고번호로 이어진 공고
        mk("TK1", "라온 도서관 건립공사 폐기물 용역", "2024-07-01", price=10_000_000, cls="건설폐기물처리서비스"),
        mk("TK2", "라온 도서관 건립공사 폐기물 용역", "2024-07-20", price=10_000_000),                         # 같은 이름이라도 종류가 다르게 읽힌 공고와는 묶지 않는다
        mk("XC1", "라온 도서관 건립 무대장치 감리용역", "2024-03-10", price=25_000_000),
        mk("XC2", "라온 도서관 건립 무대장치 감리용역", "2024-03-20", price=25_000_000, kind="취소공고"),          # 뒤 공고가 취소면 앞 공고가 그대로 대표
        mk("ZP1", "라온 도서관 건립공사 승강기 감리용역", "2024-03-15"),                                       # 추정가격 없음
        # 표준화: 차수는 숫자로 견주고(입력 순서·날짜와 무관), 같은 차수가 두 번이면 늦은 쪽, 차수는 세 자리로, 공고번호가 없는 줄은 버린다
        mk("OD1", "라온 도서관 건립 가구 설계용역", "2023-03-20", price=77_000_000, ord="001"),
        mk("OD1", "라온 도서관 건립 가구 설계용역", "2023-03-20", price=70_000_000, ord="000"),
        mk("OD2", "라온 도서관 건립 사인물 설계용역", "2023-04-20", price=33_000_000, ord="1"),
        mk("OD2", "라온 도서관 건립 사인물 설계용역", "2023-04-10", price=30_000_000, ord="1"),
        mk("", "라온 도서관 건립 번호없는 설계용역", "2023-05-10", price=1_000_000),
    ])
    raw[f["servc_class_lrg"]] = ["폐기물 처리 및 재활용서비스" if n == "TK1" else "" for n in raw[f["bid_no"]]]     # 대분류의 '재활용' 은 보지 않는다
    out = servc.link(raw, reviewed, hist, latest, cfg)
    g = {(r["공고번호"], r["시설ID"]): r for _, r in out.iterrows()}
    r = lambda no, fid: g[(no, fid)]
    yn = lambda no, fid: (r(no, fid)["합산여부"], r(no, fid)["프로젝트ID"])
    assert list(out["시설ID"]) == sorted(out["시설ID"]) and "" not in set(out["공고번호"]), "시설ID 순으로 싣는다 / 공고번호 없는 줄은 버린다"
    # 여러 시설에 걸친 공고
    assert yn("MF1", "G1") == ("N", "G1-N") and yn("MF1", "G3") == ("N", "G3-N") and "서로 다른 사업의 시설 2개" in r("MF1", "G1")["합산제외_까닭"], r("MF1", "G1").to_dict()
    assert yn("MF2", "G2") == ("Y", "G2-N") and yn("MF2", "G1") == ("N", "G1-N") and "더 정확히 맞음" in r("MF2", "G1")["합산제외_까닭"], r("MF2", "G1").to_dict()
    assert yn("MF3", "G4") == ("Y", "G4-N") and yn("MF3", "G5") == ("Y", "G5-N") and r("MF3", "G4")["상태"] == r("MF3", "G5")["상태"] == "대표", "같은 공사 공고를 함께 가진 두 후보"
    assert r("MF2", "G2")["프로젝트_후보"] == "" and r("MF2", "G2")["공사비대비_비율"] == 0.3
    # 제외 규칙과 예외
    assert yn("NB1", "G1") == ("N", "G1-N") and "건축 사업의 용역이 아닌" in r("NB1", "G1")["합산제외_까닭"], r("NB1", "G1").to_dict()
    assert yn("NB2", "G1") == ("Y", "G1-N") and r("NB2", "G1")["공사비대비_비율"] == 0.02, "프로젝트에 든 토목 공고의 용역 / 비율의 분모는 그 프로젝트 공사 추정가격의 합(12억)"
    assert yn("MT1", "G6") == ("Y", "G6-R") and yn("MT2", "G6") == ("Y", "G6-R"), "DB 에 실린 공사 이름이 통째로 있으면 유지보수·폐기물 출처 규칙을 보지 않는다"
    assert yn("UN1", "U1") == ("N", "U1-R") and r("UN1", "U1")["합산제외_까닭"].startswith("연간·단가"), r("UN1", "U1").to_dict()
    assert yn("IT1", "I1") == ("Y", "I1-N") and yn("IT2", "I1") == ("Y", "I1-N") and r("IT2", "I1")["용역세부"] == "감리(소방)", r("IT2", "I1").to_dict()
    assert r("BR1", "B1")["연결강도"] == "확실" and ("QQ1", "Q1") not in g and r("MR1", "M1")["연결강도"] == "확실" and yn("MR1", "M1") == ("Y", "M1-N")
    assert yn("CD1", "R1") == ("Y", "R1-N"), "수요기관코드는 '.0' 이 붙어 읽혀도 맞춘다"
    assert yn("KA1", "K1") == ("N", "") and r("KA1", "K1")["프로젝트_후보"] == "K1-N-02" and r("KA1", "K1")["공사_최초공고일"] == "2024-03-01", r("KA1", "K1").to_dict()
    # 재공고
    st = lambda no: (r(no, "R1")["상태"], r(no, "R1")["합산여부"])
    assert st("RB1") == ("재공고 이전", "N") and st("RB2") == ("대표", "Y"), "120일 안에 같은 이름으로 다시 난 공고"
    assert st("RC1") == ("재공고 이전", "N") and st("RC2") == ("대표", "Y"), "550일 안이고 금액이 ±15% 안"
    assert st("RD1") == st("RD2") == ("대표", "Y"), "550일을 넘으면 금액이 같아도 다른 용역"
    assert st("RE1") == ("재공고 이전", "N") and st("RE2") == ("대표", "Y") and r("RE2", "R1")["공고종류"] == "재공고", "재공고여부 Y"
    assert st("RF1")[0] == "재공고 이전" and st("RF2") == ("대표", "Y")
    assert sorted([st("RG1")[1], st("RG2")[1]]) == ["N", "Y"], "공고일이 없어도 같은 이름의 쌍은 한 번만 더한다"
    assert st("PV1") == ("재공고 이전", "N") and st("PV2") == ("대표", "Y"), "이전공고번호"
    assert st("TK1") == ("대표", "Y") and r("TK1", "R1")["용역종류"] == "폐기물처리" and r("TK2", "R1")["용역종류"] == "기타", (r("TK1", "R1").to_dict(), r("TK2", "R1").to_dict())
    assert st("XC1") == ("대표", "Y") and st("XC2") == ("취소", "N")
    assert st("ZP1") == ("대표", "N") and r("ZP1", "R1")["합산제외_까닭"] == "추정가격 없음"
    assert r("OD1", "R1")["추정가격"] == 77_000_000 and r("OD1", "R1")["공고차수"] == "001" and r("OD2", "R1")["추정가격"] == 33_000_000 and r("OD2", "R1")["공고차수"] == "001"

    # ── 비율 점검: Y 줄만, 분모는 공사 추정가격의 합, 종류마다 제 문턱 ──
    L = lambda pid, kind, yn_, price: {"프로젝트ID": pid, "용역종류": kind, "합산여부": yn_, "추정가격": price, "공고번호": "L" + kind[:2] + yn_}
    lk = pd.DataFrame([L("G1-N", "설계", "Y", 130_000_000), L("G1-N", "설계", "N", 900_000_000),                   # 1.3억/12억 = 10.8% → 경고 아님(N 줄은 세지 않는다)
                       L("R1-N", "건설사업관리·감리", "Y", 320_000_000), L("R1-N", "폐기물처리", "Y", 210_000_000),       # 16%·10.5% → 경고
                       L("G3-N", "건설사업관리·감리", "Y", 140_000_000), L("G3-N", "폐기물처리", "Y", 90_000_000)])      # 14%·9% → 경고 아님
    lg = servc.ratio_logs(lk, latest)
    assert sorted((x["공고번호"], x["항목"]) for x in lg) == [("R1-N", "용역비 비율(건설사업관리·감리)"), ("R1-N", "용역비 비율(폐기물처리)")], lg

    # ── 화면 요약·진단의 숫자 ──
    txt = servc.summary(out, len(raw), 99, latest, pd.DataFrame({"프로젝트ID": ["G2-N", "G6-R", "G3-N", "R1-N"], "연면적_m2": [None, None, None, 1200.0]}))
    y = out[out["합산여부"] == "Y"]
    assert f"합산여부 Y): {len(y):,}줄, 프로젝트 {y['프로젝트ID'].nunique():,}개 / 전체 {latest['프로젝트ID'].nunique():,}개" in txt and len(y) > y["프로젝트ID"].nunique(), txt
    d = y[y["용역종류"] == "설계"]
    assert f"설계: {d['공고번호'].nunique()}건({len(d)}줄), 프로젝트 {d['프로젝트ID'].nunique()}개, 추정가격 합 {d.drop_duplicates('공고번호')['추정가격'].sum() / 1e8:,.1f}억" in txt \
        and d["추정가격"].sum() == 850_000_000, txt
    # 연면적이 빈 G2-N(설계 Y 줄 있음)만 센다 — G6-R(전기 감리뿐)·G3-N(설계 줄이 N)은 과업지시서로 연면적을 보완할 후보가 아니다 / 연면적이 있는 R1-N 은 세지 않는다
    assert "연면적이 빈 프로젝트 3개 가운데 설계·건설사업관리 용역(합산여부 Y)이 이어진 프로젝트: 1개" in txt, txt
    only_waste = out[(out["공고번호"] == "MT2") | (out["공고번호"] == "RB1")]
    assert "이 이어진 프로젝트: 0개" in servc.summary(only_waste, 2, 2, latest, pd.DataFrame({"프로젝트ID": ["G6-R", "R1-N"], "연면적_m2": [None, None]})), \
        "폐기물 용역만 이어졌거나 재공고 이전인 줄만 있는 프로젝트는 연면적 보완 후보가 아니다"
    assert "중앙값 0개월" in txt and "-0개월" not in txt, "시점 표시에 '-0' 이 나오지 않는다"
    prof = servc.profile(raw, cfg)
    n_priced = int((pd.to_numeric(servc.latest_orders(servc.standardize(raw, cfg))["추정가격"], errors="coerce") > 0).sum())
    assert f"1.09~1.11 {n_priced:,}건, 1 미만 0건" in prof and "건설사업관리·감리·감리(전기)" in prof and "설계·설계용역" in prof, prof


def test_stage_more(tmp):
    """단계 실행에서 시험되지 않던 것: 필요한 열만 읽기, 요약의 건수, 06 의 '용역비 비율' 경고가 실리고 빠지는 조건, 프로젝트 표는 대표 공고(latest)로, 08·05 의 서식."""
    import openpyxl
    import tests.test_dedup_and_excel as T
    cfg = dict(load_config())
    # 사용자의 config 수집 기간에 기대지 않는다: '발주완료(최근)' 기준일이 config 의 공사·용역 수집 기간이나 오늘이 아니라 실린 공사 공고에서 온다는 것을 가릴 수 있게 서로 다른 달로 둔다
    cfg["period"], cfg["servc"] = dict(cfg["period"], end="2025-08"), dict(cfg.get("servc") or {}, period_end="2026-03")
    reviewed, hist, latest, _ = _world(cfg)
    mid = reviewed.loc[reviewed["검수_시설명"] == "가상시립미술관", "시설ID"].iloc[0]
    cfg["paths"] = dict(cfg["paths"], data_dir=os.path.join(tmp, "sm", "data"), out_dir=os.path.join(tmp, "sm", "out"))
    os.makedirs(cfg["paths"]["data_dir"], exist_ok=True)
    os.makedirs(cfg["paths"]["out_dir"], exist_ok=True)
    data = lambda n: os.path.join(cfg["paths"]["data_dir"], n)

    def run(fn):
        buf = io.StringIO()
        with contextlib.redirect_stdout(buf):
            fn(cfg)
        return buf.getvalue()
    raw = _rows(cfg, [dict(no="S01", name="가상시립미술관 건립공사 건설사업관리용역", dt="2024-01-20 10:00:00", price=900_000_000),
                      dict(no="S08", name="가상시립미술관 건립 전기공사 감리용역", dt="2024-03-20 10:00:00", price=150_000_000, ord="000"),
                      dict(no="S08", name="가상시립미술관 건립 전기공사 감리용역(변경)", dt="2024-03-25 10:00:00", price=160_000_000, ord="001", kind="변경공고"),
                      dict(no="B1", name="가상시립미술관 노후시설 개보수공사 설계용역", dt="2025-10-01 10:00:00", price=2_000_000_000)])
    raw["쓰지않는열"] = "x"
    raw.to_parquet(data("servc_all.parquet"), index=False)
    hist.to_parquet(data("notices_hist.parquet"), index=False)
    latest.to_parquet(data("notices_latest.parquet"), index=False)
    agg = discover.discover_candidates(discover.standardize(pd.DataFrame(T.RAW), cfg, pd.DataFrame(T.BSIS)))
    discover.write_review_workbook(agg, os.path.join(cfg["paths"]["out_dir"], "facility_candidates.xlsx"))
    seen, real_link = {}, servc.link

    def spy(raw_, *a, **kw):
        seen["cols"] = list(raw_.columns)
        return real_link(raw_, *a, **kw)
    servc.link = spy
    try:
        out = run(pipeline.stage_link_servc)
    finally:
        servc.link = real_link
    assert "쓰지않는열" not in seen["cols"] and cfg["fields"]["name"] in seen["cols"], "연결에 쓰는 열만 읽는다(원문은 수십 개 열 × 수십만 행)"
    assert "용역 공고 4줄(공고번호 기준 3건)" in out and "[진단] 용역 공고 4줄 → 공고번호 기준 3건" in out, out
    assert "용역비가 공사비에 견주어 큰 프로젝트 1개" in out and f"{mid}-R" in out.split("용역비가 공사비에 견주어 큰 프로젝트")[1], out
    lg = pd.read_parquet(data("logs_servc.parquet"))
    assert list(lg["항목"]) == ["용역비 비율(설계)"] and list(lg["공고번호"]) == [f"{mid}-R"], lg
    sl = pd.read_parquet(data("servc_linked.parquet")).set_index("공고번호")
    base = float(pd.to_numeric(latest.loc[latest["프로젝트ID"] == f"{mid}-N", "추정가격"]).sum())
    assert base < float(pd.to_numeric(hist.loc[hist["프로젝트ID"] == f"{mid}-N", "추정가격"]).sum()), "이 가상 자료에서는 이력(재공고 이전 포함)의 합이 대표 공고의 합보다 크다"
    assert sl.at["S01", "공사비대비_비율"] == round(900_000_000 / base, 4), "공사비대비_비율의 분모는 대표 공고(latest)의 합"
    # excel: 06 에 '용역비 비율' 경고가 실리고, 08·05 의 서식이 맞다
    six = lambda wb: [r[1].value for r in wb["06_검증로그"].iter_rows(min_row=2) if r[1].value]
    run(pipeline.stage_excel)
    wb = openpyxl.load_workbook(os.path.join(cfg["paths"]["out_dir"], cfg["paths"]["excel_name"]))
    assert [c.value for c in wb["06_검증로그"][1]][1] == "항목" and six(wb).count("용역비 비율(설계)") == 1, six(wb)
    ws8, ws5 = wb[build_excel.S_SERVC], wb[build_excel.S_FAC]
    h8, h5 = [c.value for c in ws8[1]], [c.value for c in ws5[1]]
    assert ws8.cell(2, h8.index("합산여부") + 1).fill.fgColor.rgb.endswith("FFFF00") and not str(ws8.cell(2, h8.index("공고명") + 1).fill.fgColor.rgb).endswith("FFFF00"), "고칠 수 있는 칸만 노랗게"
    assert ws8.cell(2, h8.index("프로젝트ID") + 1).fill.fgColor.rgb.endswith("FFFF00"), "프로젝트ID 가 빈 줄은 프로젝트ID 도 적어야 05 에 더해진다 → 고치는 칸으로 보이게"
    note8 = next(str(r[0].value) for r in ws8.iter_rows(min_row=len(sl) + 2) if str(r[0].value or "").startswith("주:"))
    assert "프로젝트ID(노란 칸)도 적어야" in note8 and "표 아래에 새로 적은 줄은 더해지지 않는다" in note8, note8
    note5 = next(str(r[0].value) for r in ws5.iter_rows(min_row=3) if str(r[0].value or "").startswith("주:"))
    assert "더한(합산여부 Y) 공고 없음" in note5 and "이어진 공고 없음" not in note5, "빈칸 = 합산여부 Y 인 공고가 없음(이어진 공고가 없다는 뜻이 아니다)"
    assert ws8.cell(2, h8.index("공사비대비_비율") + 1).number_format == "0.0%" and ws8.cell(2, h8.index("추정가격") + 1).number_format == build_excel.MONEY
    assert ws5.cell(2, h5.index("용역_설계_추정가격(수식)") + 1).number_format == build_excel.MONEY and ws5.cell(2, h5.index("용역_공고수(수식)") + 1).number_format != build_excel.MONEY
    # 05 수식의 조건과 범위를 글자 그대로 고정한다(LibreOffice 가 없는 PC 에서는 값 검증이 건너뛰어지므로)
    L8 = {c: openpyxl.utils.get_column_letter(i + 1) for i, c in enumerate(h8)}
    rng = lambda c: f"'{build_excel.S_SERVC}'!${L8[c]}$2:${L8[c]}${len(sl) + 1}"
    for kind in build_excel.SERVC_KINDS:
        cond = f'{rng("프로젝트ID")},$A2,{rng("용역종류")},"{kind}",{rng("합산여부")},"Y"'
        assert ws5.cell(2, h5.index(f"용역_{kind}_추정가격(수식)") + 1).value == f'=IF(COUNTIFS({cond})=0,"",SUMIFS({rng("추정가격")},{cond}))', kind
        assert ws5.cell(2, h5.index(f"용역_{kind}_추정가격(수식)") + 1).number_format == build_excel.MONEY, kind
    assert ws5.cell(2, h5.index("용역_공고수(수식)") + 1).value == f'=COUNTIFS({rng("프로젝트ID")},$A2,{rng("합산여부")},"Y")' and ws8.max_row >= len(sl) + 1
    done = f'{rng("프로젝트ID")},$A2,{rng("용역종류")},"폐기물처리",{rng("합산여부")},"Y"'
    assert ws5.cell(2, h5.index(build_excel.SERVC_DONE_COL) + 1).value == \
        f'=IF(COUNTIFS({done})=0,"",IF(COUNTIFS({done},{rng("공사수집종료까지_일수")},"<183")=COUNTIFS({done}),"발주완료(최근)","발주완료"))', \
        "발주완료 = 더한(Y) 폐기물처리 용역이 있음. 그 줄이 모두 공사 공고 수집 종료일에서 6개월(183일) 안쪽·그 뒤면 발주완료(최근)"
    assert ws5.cell(2, h5.index(build_excel.SERVC_DONE_COL) + 1).number_format != build_excel.MONEY and build_excel.SERVC_DONE_KIND in build_excel.SERVC_KINDS
    assert "발주완료_폐기물용역(수식) =" in note5 and "'미완료' 가 아니다" in note5 and "'발주완료_폐기물용역' 이 '발주완료' 가 된다" in note8, (note5[-400:], note8[-300:])
    # 보조지표의 한계를 주석에 적는다: 최근 공고는 아직 발주 중일 수 있음 / 공사비가 다 들어 있다는 뜻이 아님 / 빈칸은 미완료가 아님(README 시트에도)
    assert "08 의 공고일시를 확인할 것" in note5 and "공종누락경고는 따로 본다" in note5 and "공사비가 모두 들어 있다는 뜻이 아니다" in note5, note5[-800:]
    readme = " ".join(str(r[1].value) for r in wb["README"].iter_rows() if r[0].value == "용역(08·05)")
    assert "'발주완료_폐기물용역' =" in readme and "'미완료' 가 아니라 '판단할 근거 없음'" in readme and "공고가 더 나올 수 있다" in readme and "공종누락경고를 함께 볼 것" in readme, readme
    assert build_excel.SERVC_DONE_COL == "발주완료_폐기물용역(수식)", "05·08 의 주석과 README 가 이 열 이름을 글자 그대로 적는다"
    # 발주완료(최근)의 기준일 = 공사 공고 수집 종료일 = 이 워크북에 실린 공사 공고(notices_hist → 02 시트)의 가장 늦은 공고일. config 의 공사·용역 수집 기간이나 오늘 날짜가 아니다
    # (수집 기간만 늘리고 excel 을 다시 돌려도 공사 공고가 그대로면 기준일도 그대로여야 한다). 그 날이 08 의 일수·세 곳의 설명에 그대로 실린다
    import datetime as _dt
    end = dedup._parse_dt(pd.read_parquet(data("notices_hist.parquet"))["공고일시"]).max().date().isoformat()
    assert end == "2026-05-03" and _dt.date.today().isoformat() > end and cfg["period"]["end"] == "2025-08" and cfg["servc"]["period_end"] == "2026-03", (end, cfg["period"], cfg["servc"])
    h2 = [c.value for c in wb["02_공고목록_전체"][1]]
    assert max(str(r[h2.index("공고일시")].value)[:10] for r in wb["02_공고목록_전체"].iter_rows(min_row=2) if r[h2.index("공고일시")].value) == end, "기준일은 02 시트의 가장 늦은 공고일"
    for txt_, where in ((note5, "05 주석"), (note8, "08 주석"), (readme, "README 시트")):
        assert f"공사 공고 수집 종료일({end})" in txt_ and "'발주완료(최근)'" in txt_ and "읽지 못해" not in txt_, (where, txt_[-900:])
        assert "공사 공고 수집 종료일 = 이 워크북에 실린 공사 공고 가운데 가장 늦은 공고일(02 시트)" in txt_, (where, txt_[-900:])
    assert "6개월(183일) 안쪽이거나 그 뒤면" in note5 and "'공사수집종료까지_일수' 가 모두 183 미만" in note5 and "발주가 끝난 사업을 고를 때는 '발주완료' 만 고르고" in note5 and "음수 = 그날 뒤에 난 용역 공고" in note8
    assert "6개월(183일) 안쪽이거나 그 뒤면" in readme and "'공사수집종료까지_일수' 가 모두 183 미만이면" in note8, "README·08 주석도 같은 기준 날수(183)를 적는다"
    # '발주완료' 인 행도 공사 공고가 최근이면 아직 발주 중일 수 있다는 한계(05 주석·README), 08 에 폐기물 줄을 끼워 넣으면 일수 칸도 적어야 한다는 안내(08 주석)
    assert "01 의 '최종공고일' 이 공사 공고 수집 종료일에서 6개월 안쪽" in note5 and "01 의 '최종공고일'" in readme and "폐기물처리 줄을 끼워 넣었으면 이 칸도 적을 것" in note8, (note5[-900:], note8[-600:])
    assert "최종공고일" in [c.value for c in wb["01_시설마스터"][1]], "주석이 가리키는 열이 01 에 있다"
    assert build_excel.SERVC_AGE_COL == "공사수집종료까지_일수" and [ws8.cell(r, h8.index(build_excel.SERVC_AGE_COL) + 1).value for r in range(2, len(sl) + 2)] == \
        [int((pd.Timestamp(end) - pd.Timestamp(str(ws8.cell(r, h8.index("공고일시") + 1).value)[:10])).days) for r in range(2, len(sl) + 2)], "08 의 일수 = 종료일 − 용역 공고일"
    # 연결 결과를 지우면(= link-servc 를 돌리지 않은 상태) 남아 있는 logs_servc.parquet 의 경고도 06 에 싣지 않는다
    os.remove(data("servc_linked.parquet"))
    run(pipeline.stage_excel)
    wb = openpyxl.load_workbook(os.path.join(cfg["paths"]["out_dir"], cfg["paths"]["excel_name"]))
    assert os.path.exists(data("logs_servc.parquet")) and not any(str(v).startswith("용역비 비율") for v in six(wb)), six(wb)


def _mini(hist_items, reviewed_items, raw_items, cfg):
    """작은 가상 세계 하나를 돌려 {(공고번호, 시설ID): 행} 을 돌려준다. hist: (시설ID, 시설명, 프로젝트ID, 공고번호, 공고명, 날짜, 수요기관[, 사업유형[, 추정가격]])"""
    base = dict(공고차수="000", 공고종류="일반", 공종="건축", 세부사업명="", 수요기관코드="")
    hist = pd.DataFrame([dict(base, 시설ID=h[0], 시설명=h[1], 프로젝트ID=h[2], 공고번호=h[3], 공고명=h[4], 공고일시=h[5] + " 10:00:00", 수요기관=h[6],
                              사업유형=(h[7] if len(h) > 7 else "신축"), 추정가격=(h[8] if len(h) > 8 else 1_000_000_000)) for h in hist_items])
    reviewed = pd.DataFrame([{"시설ID": r[0], "검수_시설명": r[1], "검수_별칭(;구분)": "", "검수_수요기관": (r[3] if len(r) > 3 else ""), "수요기관": r[2], "검색어": ""} for r in reviewed_items])
    raw = _rows(cfg, [dict(no=x[0], name=x[1], dt=x[2] + " 10:00:00", inst=x[3], price=x[4], **(x[5] if len(x) > 5 else {})) for x in raw_items])
    out = servc.link(raw, reviewed, hist, hist.copy(), cfg)
    return {(r["공고번호"], r["시설ID"]): r for _, r in out.iterrows()}, out, hist


def test_review_fixes():
    """독립 검토(2026-11-23)에서 재현된 결함을 고친 자리. 모든 이름은 가상."""
    cfg = load_config()
    k, I = servc.classify_kind, "경기도 가온시"
    # ── 종류 판정 ──
    assert k("◇ 건립공사 미술작품 제작·설치 제안공모")[0] == k("◇ 신축 구내식당 위탁운영 사업자 선정 제안공모")[0] == "기타", "'제안공모' + 건립·신축 말만으로는 설계가 아니다"
    assert k("◇ 건립공사 건축설계 제안공모")[:2] == ("설계", "설계공모")
    for t in ("◇ 철거 및 신축공사 감리용역", "(구)◇ 철거 후 신축공사 건축감리용역", "◇ 리모델링공사(철거 포함) 감리용역"):
        assert k(t)[0] == "건설사업관리·감리" and k(t)[2] == "" and "석면" not in k(t)[1], (t, k(t))
    for t in ("◇ 해체공사 감리용역", "◇ 리모델링공사 석면해체 제거공사 감리용역", "◇ 교사동 개축 및 철거공사 감리용역"):
        assert k(t)[1] == "감리(석면·해체)" and k(t)[2], (t, k(t))
    for t, kind in (("도시계획시설(문화시설) ◇ 건립 기본 및 실시설계용역", "설계"), ("◇ 신축공사 설계용역(측량 및 지반조사 포함)", "설계"),
                    ("◇ 건립공사 감리용역(BF인증 포함)", "건설사업관리·감리"), ("도시계획시설(공공청사) ◇ 건립공사 건설사업관리용역", "건설사업관리·감리")):
        assert k(t)[0] == kind and k(t)[2] == "", (t, k(t))
    assert k("◇ 건립 설계공모 관리용역")[0] == k("◇ 건립부지 지반조사 용역")[0] == k("◇ 녹색건축 인증 용역(설계단계)")[0] == "기타", "부대 용역은 그대로 기타"
    for t in ("◇ 건립공사 건설폐기물(건설오니) 처리용역", "◇ 건립공사 건설오니 처리용역", "◇ 건립공사 건설폐기물 처리용역(재활용)", "◇ 건립공사 폐아스콘 재활용 처리",
              "◇ 건립공사 건설폐기물 처리(민간위탁)", "◇ 건립공사 건설폐기물 처리 장기계속계약"):
        assert k(t)[:2] == ("폐기물처리", "건설폐기물") and k(t)[2] == "", (t, k(t))
    assert k("◇ 음식물류폐기물 재활용 처리용역")[0] == k("◇ 폐기물 처리장 설계용역")[0] == k("하수슬러지 처리용역", "폐기물처리서비스")[0] == "기타"
    assert k("◇ 건립공사 관급자재 설계용역")[1] == "설계용역" and k("◇ 건립공사 재설계용역")[1] == "설계용역(재설계·보완)"
    # ── 기관: 같은 계열의 다른 기관(inst_same 이지만 inst_narrow 는 아님) ──
    nr = servc.inst_narrow
    assert servc.inst_same("국가보훈처 국립대전현충원", "국가보훈처 국립이천호국원") and not nr("국가보훈처 국립대전현충원", "국가보훈처 국립이천호국원")
    assert not nr("경기도교육청 수원농생명과학고등학교", "경기도교육청 여주자영농업고등학교") and not nr("한국농어촌공사 경기지역본부", "한국농어촌공사 충남지역본부 부여지사")
    assert nr("경기도교육청", "경기도교육청 수원농생명과학고등학교") and nr("한국농어촌공사", "한국농어촌공사 경기지역본부") and nr("경기도 가온시", "경기도 가온시 체육시설사업소")
    assert nr("강원도", "강원특별자치도 가온사업소") and not nr("경기도 가온시", "경기도 나래시")
    g, out, _ = _mini(
        [("F1", "국가보훈처 봉안당", "F1-N", "C1", "국립대전현충원 봉안당 신축공사", "2022-03-02", "국가보훈처 국립대전현충원")],
        [("F1", "국가보훈처 봉안당", "국가보훈처 국립대전현충원", "국가보훈처 국립대전현충원")],
        [("S1", "국립이천호국원 봉안당 신축공사 기본 및 실시설계용역", "2021-02-01", "국가보훈처 국립이천호국원", 310_000_000),
         ("S2", "국립괴산호국원 봉안당 건립공사 건설사업관리용역", "2022-05-01", "국가보훈처 국립괴산호국원", 450_000_000),
         ("S9", "국립대전현충원 봉안당 신축공사 설계용역", "2021-01-01", "국가보훈처 국립대전현충원", 200_000_000)], cfg)
    assert g[("S9", "F1")]["합산여부"] == "Y" and g[("S1", "F1")]["합산여부"] == g[("S2", "F1")]["합산여부"] == "N", out[["공고번호", "합산여부", "합산제외_까닭"]].to_dict("records")
    assert "같은 계열의 다른 기관" in g[("S1", "F1")]["연결근거"] and "같은 계열의 다른 기관" in g[("S1", "F1")]["합산제외_까닭"] and g[("S1", "F1")]["연결강도"] == "확실"
    # ── 사업유형 판정에서 용역 이름의 '건설' 을 보지 않는다: 유지보수 공사의 '건설사업관리'·'건설폐기물' 용역 ──
    g, out, _ = _mini([("F1", "가온회관", "F1-N", "C1", "가온회관 건립공사", "2023-03-02", I)], [("F1", "가온회관", I)],
                      [("M1", "가온회관 외벽 도장공사 감리용역", "2025-04-01", I, 11_000_000), ("M2", "가온회관 외벽 도장공사 건설사업관리용역", "2025-04-02", I, 12_000_000),
                       ("M3", "가온회관 외벽 도장공사 건설폐기물 처리용역", "2025-06-01", I, 13_000_000), ("M4", "가온회관 승강기 교체공사 건설사업관리용역", "2026-02-01", I, 14_000_000),
                       ("M5", "가온회관 폐기물 처리용역", "2024-01-10", I, 15_000_000), ("M6", "가온회관 건설폐기물 처리용역", "2024-01-11", I, 16_000_000),
                       ("M7", "가온회관 건립공사 건설폐기물 처리용역", "2023-09-01", I, 17_000_000), ("M8", "가온회관 건립공사 건설사업관리용역", "2023-02-01", I, 18_000_000)], cfg)
    # (M6 '가온회관 건설폐기물 처리용역' 은 2026-11 실자료 점검 뒤 Y — 프로젝트가 하나뿐인 시설에서 공사 공고 10개월 뒤에 난 '건설폐기물' 공고. test_field_audit ⑩)
    assert [g[(n, "F1")]["합산여부"] for n in ("M1", "M2", "M3", "M4", "M5", "M6", "M7", "M8")] == ["N"] * 5 + ["Y", "Y", "Y"], out[["공고번호", "합산여부", "합산제외_까닭"]].to_dict("records")
    assert g[("M2", "F1")]["합산제외_까닭"].startswith("유지보수") and g[("M5", "F1")]["합산제외_까닭"].startswith("어느 공사에서 나온 폐기물인지")
    assert g[("M8", "F1")]["합산제외_까닭"] == "" and "일반적인 시설 이름" not in g[("M8", "F1")]["연결근거"], "검수한 이름이 원래 네 글자인 시설은 '일반적인 이름' 규칙에 걸리지 않는다"
    # ── 한 공고가 두 시설에: 가리키는 시설에서 프로젝트를 못 정했다고 이름이 짧은 다른 시설이 가져가지 않는다 ──
    g, out, _ = _mini([("FA", "가온 체육관", "FA-E", "C1", "가온 체육관 증축공사", "2024-06-03", I, "증축"),
                       ("FB", "가온 체육관 수영장", "FB-E-01", "C2", "가온 체육관 수영장 증축공사", "2024-03-04", I, "증축"),
                       ("FB", "가온 체육관 수영장", "FB-E-02", "C3", "가온 체육관 수영장 증축공사", "2024-10-07", I, "증축")],
                      [("FA", "가온 체육관", I), ("FB", "가온 체육관 수영장", I)],
                      [("S1", "가온 체육관 수영장 증축공사 기본 및 실시설계용역", "2023-02-01", I, 400_000_000), ("S2", "가온 체육관 수영장 증축공사 건설사업관리용역", "2024-06-20", I, 700_000_000),
                       ("S3", "가온 체육관 증축공사 건설사업관리용역", "2024-06-21", I, 500_000_000)], cfg)
    assert g[("S1", "FA")]["합산여부"] == g[("S2", "FA")]["합산여부"] == "N" and "다른 시설(FB)에도" in g[("S2", "FA")]["합산제외_까닭"], g[("S2", "FA")].to_dict()
    assert g[("S3", "FA")]["합산여부"] == "Y" and ("S3", "FB") not in g, "제 시설 이름만 든 공고는 그대로 더한다"
    # ── 묶음 발주: 서로 다른 시설 이름이 '및'·'·' 로 이어져 있으면 사업 이름이 통째로 맞는 쪽이 있어도 모두 N ──
    g, out, _ = _mini([("F1", "가온 체육관", "F1-N", "C1", "가온 체육관 건립공사", "2024-03-01", I), ("F2", "다온 수영장", "F2-N", "C2", "다온 수영장 건립공사", "2024-04-01", I),
                       ("F3", "나래 도서관", "F3-N-01", "C3", "나래 도서관 건립공사", "2021-03-01", I), ("F3", "나래 도서관", "F3-N-02", "C4", "나래 도서관 건립공사", "2024-03-01", I),
                       ("F4", "가온시 주차타워", "F4-N", "C5", "주차타워 건립공사", "2024-03-01", I)],
                      [("F1", "가온 체육관", I), ("F2", "다온 수영장", I), ("F3", "나래 도서관", I), ("F4", "가온시 주차타워", I)],
                      [("B1", "가온 체육관 및 다온 수영장 건립 설계용역", "2023-01-10", I, 90_000_000), ("B2", "가온 체육관·다온 수영장 건립공사 통합건설사업관리용역", "2024-02-10", I, 70_000_000),
                       ("B3", "가온 체육관 및 나래 도서관 건립공사 건설폐기물 처리용역", "2024-09-10", I, 8_000_000),
                       ("G1", "가온역 환승 주차타워 건립공사 건설사업관리용역", "2026-01-10", I, 50_000_000), ("G2", "주차타워 건립공사 건설사업관리용역", "2024-02-10", I, 50_000_000),
                       ("G3", "가온시 주차타워 건립공사 설계용역", "2020-01-10", I, 50_000_000), ("D1", "나래 도서관 리모델링 건설사업관리용역", "2024-03-01", I, 50_000_000)], cfg)
    for no, fs in (("B1", ("F1", "F2")), ("B2", ("F1", "F2")), ("B3", ("F1", "F3"))):
        assert all(g[(no, f)]["합산여부"] == "N" for f in fs) and "묶음 발주" in g[(no, fs[0])]["합산제외_까닭"], (no, [g[(no, f)].to_dict() for f in fs])
    # 접두를 뗀 짧은 이름('주차타워')으로만 이어진 공고: 공사와 시점이 가까울 때만 더하고 근거에 적는다. 검수한 이름이 그대로 있으면 해당하지 않는다
    assert g[("G1", "F4")]["합산여부"] == "N" and "일반적인 시설 이름" in g[("G1", "F4")]["합산제외_까닭"] and "일반적인 시설 이름(4글자)" in g[("G1", "F4")]["연결근거"]
    assert g[("G2", "F4")]["합산여부"] == "Y" and "일반적인 시설 이름(4글자)" in g[("G2", "F4")]["연결근거"]
    assert g[("G3", "F4")]["합산여부"] == "Y" and "일반적인 시설 이름" not in g[("G3", "F4")]["연결근거"], "검수한 이름이 통째로 있으면 넓은 창(설계 84개월 전)을 그대로 쓴다"
    assert g[("D1", "F3")]["프로젝트ID"] == "" and g[("D1", "F3")]["공사_최초공고일"] == "2024-03-01" and g[("D1", "F3")]["공사대비_개월"] == 0.0, "같은 날의 프로젝트가 '가장 가까운' 프로젝트"
    # ── 같은 용역이 두 번 더해지지 않게: 재재공고·수정공고·사전공고·PQ 서류 안내, 금액을 고쳐 다시 낸 설계·건설사업관리 / 폐기물 차수분은 둘 다 ──
    T, D, Wn = "가온회관 건립공사 건설사업관리용역", "가온회관 건립공사 기본 및 실시설계용역", "가온회관 건립공사 건설폐기물 처리용역"

    def ys(*items):
        g_, out_, _ = _mini([("F1", "가온회관", "F1-N", "C1", "가온회관 건립공사", "2024-09-02", I, "신축", 30_000_000_000)], [("F1", "가온회관", I)],
                            [(no, name, dt, I, price) for no, name, dt, price in items], cfg)
        return sorted(no for (no, _), r in g_.items() if r["합산여부"] == "Y"), out_
    for items, want in (
            ([("A1", T, "2024-01-10", 9e8), ("A2", T + "(재공고)", "2024-02-20", 9e8), ("A3", T + "(재재공고)", "2024-03-25", 9e8)], ["A3"]),
            ([("A2", T + " 재공고", "2024-02-20", 9e8), ("A3", T + " 재재공고", "2024-03-25", 9e8)], ["A3"]),
            ([("A1", T, "2024-01-10", 9e8), ("A3", T + " 수정공고", "2024-01-25", 9e8)], ["A3"]),
            ([("P1", T + " 사전공고", "2024-01-10", 9e8), ("P2", T, "2024-02-20", 9e8)], ["P2"]),
            ([("P1", T + " 사업수행능력평가(PQ) 서류 제출 안내 공고", "2024-01-10", 9e8), ("P2", T, "2024-02-20", 9e8)], ["P2"]),
            ([("D1", D, "2022-03-10", 10e8), ("D2", D, "2022-10-12", 12e8)], ["D2"]),                     # 7개월 뒤 금액을 20% 올려 다시 낸 설계
            ([("A1", T, "2024-01-10", 9e8), ("A2", T, "2024-06-12", 10.8e8)], ["A2"]),                    # 5개월 뒤 금액을 고쳐 다시 낸 건설사업관리
            ([("W1", Wn + "(1차)", "2024-11-10", 5e7), ("W2", Wn + "(2차)", "2025-02-10", 8e7)], ["W1", "W2"]),   # 폐기물 차수분: 차수 표기와 금액이 모두 다르다
            ([("W1", Wn + "(1차)", "2024-11-10", 5e7), ("W2", Wn + "(2차)", "2024-12-10", 5e7)], ["W2"]),         # 금액이 같으면 다시 낸 공고
            ([("W1", Wn, "2024-11-10", 5e7), ("W2", Wn, "2024-11-25", 8e7)], ["W2"]),                             # 차수 표기가 없으면 120일 규칙 그대로
            ([("W1", Wn, "2024-11-10", 5e7), ("W2", Wn, "2025-06-10", 8e7)], ["W1", "W2"])):                      # 7개월 뒤 금액이 다른 폐기물 공고는 다른 용역
        got, out_ = ys(*items)
        assert got == want, (items, out_[["공고번호", "상태", "합산여부", "합산제외_까닭"]].to_dict("records"))
    assert servc.rebid_key(T + "(재재공고)") == servc.rebid_key(T + " 수정공고") == servc.rebid_key(T + " 사업수행능력평가(PQ) 서류 제출 안내 공고") == servc.rebid_key(T)
    # ── 유사도는 사업을 가리는 말이 양쪽에 있을 때만 ──
    g, out, _ = _mini([("F1", "가온회관", "F1-N-01", "C1", "가온회관 건립공사", "2021-03-02", I), ("F1", "가온회관", "F1-N-02", "C2", "가온회관 전기공사", "2023-10-04", I, "신축", 3e8),
                       ("F1", "가온회관", "F1-N-02", "C3", "가온회관 소방공사", "2023-10-04", I, "신축", 1e8)], [("F1", "가온회관", I)],
                      [("S2", "가온회관 소방공사 감리용역", "2023-02-01", I, 9_000_000), ("S3", "가온회관 전기공사 감리용역", "2023-09-01", I, 20_000_000)], cfg)
    assert g[("S2", "F1")]["프로젝트ID"] == "" and "가릴 수 없음" in g[("S2", "F1")]["연결근거"] and g[("S3", "F1")]["프로젝트ID"] == "F1-N-02", out[["공고번호", "프로젝트ID", "연결근거"]].to_dict("records")
    g, out, _ = _mini([("F2", "나래회관", "F2-N", "D1", "나래회관 건립공사", "2021-03-02", I), ("F2", "나래회관", "F2-R", "D2", "나래회관 객석 개보수공사", "2026-03-02", I, "리모델링")],
                      [("F2", "나래회관", I)], [("T1", "나래회관 감리용역", "2023-01-10", I, 30_000_000)], cfg)
    assert g[("T1", "F2")]["프로젝트ID"] == "" and g[("T1", "F2")]["합산여부"] == "N", "공고명에 시설 이름뿐이면 유사도로 정하지 않는다"
    # ── 시설 이름에 기호가 있어도 '시설 이름 + 공종' 뿐인 사업 이름은 근거가 아니다 ──
    J = "경상남도 가온군"
    g, out, _ = _mini([("F1", "가온4·26추모공원", "F1-N-01", "C1", "가온 4·26 추모공원 조성사업", "2022-03-02", J, "신축", 9e9),
                       ("F1", "가온4·26추모공원", "F1-N-02", "C2", "가온 4·26 추모공원 조경 공사", "2024-09-02", J, "신축", 3e8)], [("F1", "가온4·26추모공원", J)],
                      [("S1", "가온 4·26 추모공원 건립 설계공모", "2020-05-01", J, 500_000_000), ("S2", "가온 4·26 추모공원 건설사업관리용역", "2022-02-01", J, 400_000_000)], cfg)
    assert g[("S1", "F1")]["프로젝트ID"] == g[("S2", "F1")]["프로젝트ID"] == "F1-N-01", out[["공고번호", "프로젝트ID", "연결근거"]].to_dict("records")


def test_review_round2():
    """2차 검토(2026-11-23)의 변이 시험에서 살아남은 자리 — 고친 규칙의 조건 하나하나가 실제로 듣는지. 모든 이름은 가상."""
    cfg = load_config()
    k, nr, I = servc.classify_kind, servc.inst_narrow, "경기도 가온시"
    # ── 종류 판정 ──
    assert k("◇ 철거 및 신축공사 감리용역")[1] == "감리", "본 공사의 감리로 읽은 공고의 세부에 '해체' 표지를 남기지 않는다"
    assert k("◇ 건립공사 건설폐기물 처리용역", "건설폐기물 재활용서비스")[:2] == ("폐기물처리", "건설폐기물"), "건설폐기물 공고는 조달분류의 '재활용' 으로 운영 폐기물이 되지 않는다"
    assert k("◇ 건립공사 건설오니 위탁 처리")[0] == "폐기물처리", "'건설오니 … 위탁' 도 폐기물 처리 용역의 이름이다"
    # ── 기관: 같은 시·군·구의 구청·사업소끼리(이름이 서로 들어 있지 않아도) ──
    assert nr("경기도 가온시 나래구", "경기도 가온시 체육시설사업소") and not nr("경기도 가온시 나래구", "경기도 나래시 가온구")
    # ── 같은 계열의 다른 기관(소속기관이 다름): 공고명에 사업 이름이 통째로 있으면 더한다 / 수요기관코드가 같으면 '바로 그 기관' ──
    base = dict(공고차수="000", 공고종류="일반", 공종="건축", 세부사업명="", 사업유형="신축", 추정가격=1_000_000_000)
    hist = pd.DataFrame([dict(base, 시설ID="F1", 시설명="국가보훈처 봉안당", 프로젝트ID="F1-N", 공고번호="C1", 공고명="국립대전현충원 봉안당 신축공사", 공고일시="2022-03-02 10:00:00",
                              수요기관="국가보훈처 국립대전현충원", 수요기관코드="1230001")])
    reviewed = pd.DataFrame([{"시설ID": "F1", "검수_시설명": "국가보훈처 봉안당", "검수_별칭(;구분)": "", "검수_수요기관": "", "수요기관": "국가보훈처 국립대전현충원", "검색어": ""}])
    raw = _rows(cfg, [dict(no="S8", name="국립대전현충원 봉안당 신축공사 감리용역", dt="2022-03-10 10:00:00", inst="국가보훈처 국립이천호국원", price=50_000_000),
                      dict(no="K1", name="국립대전현충원 봉안당 설계용역", dt="2021-06-01 10:00:00", inst="국가보훈부 국립대전현충원", price=200_000_000, code="1230001"),
                      dict(no="K2", name="국립대전현충원 봉안당 설계용역(조경)", dt="2021-06-02 10:00:00", inst="국가보훈부 국립대전현충원", price=20_000_000, code="9999999")])
    out = servc.link(raw, reviewed, hist, hist.copy(), cfg)
    g = {(r["공고번호"], r["시설ID"]): r for _, r in out.iterrows()}
    assert g[("S8", "F1")]["합산여부"] == "Y" and "같은 계열의 다른 기관" in g[("S8", "F1")]["연결근거"], "다른 소속기관이 낸 공고도 이 프로젝트의 사업 이름이 통째로 있으면 더한다"
    assert g[("K1", "F1")]["합산여부"] == "Y" and "같은 발주 기관" in g[("K1", "F1")]["연결근거"], "수요기관코드가 같으면 이름이 바뀌어도 바로 그 기관"
    assert ("K2", "F1") not in g, "이름도 코드도 다른 기관의 공고는 잇지 않는다"
    # ── 사업유형·일반 이름·대행 기관 ──
    g, out, _ = _mini([("F1", "가온회관", "F1-N", "C1", "가온회관 건립공사", "2023-03-02", I), ("F4", "가온시 국민체육센터", "F4-N", "C5", "국민체육센터 건립공사", "2024-03-01", I)],
                      [("F1", "가온회관", I), ("F4", "가온시 국민체육센터", I)],
                      [("M9", "가온회관 외벽 보수공사 건설기술용역(감리)", "2025-04-03", I, 9_000_000), ("G5", "국민체육센터 건립공사 건설사업관리용역", "2025-01-10", I, 50_000_000),
                       ("R1", "가온회관 감리용역", "2023-03-10", "가온도시공사", 30_000_000)], cfg)
    assert g[("M9", "F1")]["합산여부"] == "N" and g[("M9", "F1")]["합산제외_까닭"].startswith("유지보수"), "'건설기술용역' 의 '건설' 도 신축 어휘가 아니다"
    assert g[("G5", "F4")]["합산여부"] == "Y" and "일반적인 시설 이름" not in g[("G5", "F4")]["연결근거"], "접두를 뗀 이름이 다섯 글자 이상이면 '일반적인 이름' 규칙의 대상이 아니다"
    assert g[("R1", "F1")]["연결강도"] == "추정" and g[("R1", "F1")]["합산제외_까닭"].startswith("연결이 추정임"), "대행 기관의 공고는 '같은 계열의 다른 기관' 이 아니라 '추정' 으로 적는다"
    # ── 같은 공사 공고를 함께 가진 시설 후보끼리는 이름이 '및' 으로 이어져 있어도 묶음이 아니다(한 사업) ──
    g, out, _ = _mini([("F1", "가온 도서관", "F1-N", "C1", "가온 도서관 및 가온 생활문화센터 건립공사", "2023-03-02", I), ("F2", "가온 생활문화센터", "F2-N", "C1", "가온 도서관 및 가온 생활문화센터 건립공사", "2023-03-02", I)],
                      [("F1", "가온 도서관", I), ("F2", "가온 생활문화센터", I)], [("S1", "가온 도서관 및 가온 생활문화센터 건립공사 건설사업관리용역", "2023-02-01", I, 500_000_000)], cfg)
    assert g[("S1", "F1")]["합산여부"] == g[("S1", "F2")]["합산여부"] == "Y", out[["공고번호", "시설ID", "합산여부", "합산제외_까닭"]].to_dict("records")
    # ── 이름이 겹친 두 시설: 공고명에 이름이 두 번 나오면 뒤의 자리도 본다('○○ 체육관 수영장 및 ○○ 체육관 주차장') ──
    g, out, _ = _mini([("FA", "가온 체육관", "FA-E", "C1", "가온 체육관 증축공사", "2024-06-03", I, "증축"), ("FB", "가온 체육관 수영장", "FB-N", "C2", "가온 체육관 수영장 건립공사", "2024-03-04", I)],
                      [("FA", "가온 체육관", I), ("FB", "가온 체육관 수영장", I)], [("S2", "가온 체육관 수영장 및 가온 체육관 주차장 건립 설계용역", "2023-03-01", I, 900_000_000)], cfg)
    assert g[("S2", "FB")]["합산여부"] == "N" and "묶음 발주" in g[("S2", "FB")]["합산제외_까닭"], g[("S2", "FB")].to_dict()
    # ── 한 번짜리 용역 규칙은 같은 프로젝트일 때만: 같은 이름의 두 프로젝트(2022, 2025)의 설계가 16개월 간격으로 났으면 둘 다 남긴다 ──
    D = "가온회관 건립공사 실시설계용역"
    g, out, _ = _mini([("F1", "가온회관", "F1-N-01", "C1", "가온회관 건립공사", "2022-03-02", I), ("F1", "가온회관", "F1-N-02", "C2", "가온회관 건립공사", "2025-03-03", I)], [("F1", "가온회관", I)],
                      [("D1", D, "2021-06-01", I, 900_000_000), ("D2", D, "2022-10-04", I, 600_000_000)], cfg)   # (금액이 공사 추정가격 합 10억과 같으면 '공사비를 적은 공고' 로 빠진다 — test_field_audit ⑦)
    assert [(g[(n, "F1")]["프로젝트ID"], g[(n, "F1")]["합산여부"]) for n in ("D1", "D2")] == [("F1-N-01", "Y"), ("F1-N-02", "Y")], out[["공고번호", "프로젝트ID", "상태", "합산여부", "합산제외_까닭"]].to_dict("records")
    # ── 유사도: 사업 이름이 비워진 프로젝트(시설 이름 + 공종뿐)가 후보에 있으면, 공고명에 사업을 가리는 말이 있어도 유사도로 정하지 않는다 ──
    g, out, _ = _mini([("F1", "가온회관", "F1-E-01", "C1", "가온회관 별관 증축공사", "2021-03-02", I, "증축"), ("F1", "가온회관", "F1-E-02", "C2", "가온회관 전기공사", "2023-10-04", I, "증축", 3e8)],
                      [("F1", "가온회관", I)], [("S1", "가온회관 별관 감리용역", "2022-06-01", I, 30_000_000)], cfg)
    assert g[("S1", "F1")]["프로젝트ID"] == "" and "가릴 수 없음" in g[("S1", "F1")]["연결근거"], out[["공고번호", "프로젝트ID", "연결근거"]].to_dict("records")


def test_merge_rules():
    """검토 반영을 합친 뒤(2026-11-23) 시험이 잡지 못하던 자리: 폐기물 재공고의 금액·표시 조건, 프로젝트가 다른 같은 이름의 설계, 일반 이름 규칙의 길이,
    조달분류의 '재활용', 같은 계열 기관의 '사업 이름이 통째로 있으면 더함', 요약의 건수(공고마다 한 번). 모든 이름은 가상."""
    cfg = load_config()
    k, I = servc.classify_kind, "경기도 가온시"
    Wn = "가온회관 건립공사 건설폐기물 처리용역"

    def ys(items, hist=None, rev=None):
        g_, out_, _ = _mini(hist or [("F1", "가온회관", "F1-N", "C1", "가온회관 건립공사", "2024-09-02", I, "신축", 30_000_000_000)], rev or [("F1", "가온회관", I)],
                            [(x[0], x[1], x[2], I, x[3]) + ((x[4],) if len(x) > 4 else ()) for x in items], cfg)
        return sorted(no for (no, _), r in g_.items() if r["합산여부"] == "Y"), g_, out_
    # 폐기물 처리(한 번 하는 용역이 아님): 120일~550일 사이는 금액이 ±15% 이내이거나, 한쪽 금액이 없거나, 뒤 공고가 재공고로 표시됐을 때만 다시 낸 공고
    for items, want in (
            ([("W1", Wn, "2024-11-10", 5e7), ("W2", Wn, "2025-06-10", 5.5e7)], ["W2"]),                       # 10% 차이 → 같은 금액으로 본다
            ([("W1", Wn, "2024-11-10", 5e7), ("W2", Wn, "2025-06-10", 9e7, dict(re="Y"))], ["W2"]),           # 금액이 달라도 재공고로 표시
            ([("W1", Wn, "2024-11-10", 5e7), ("W2", Wn, "2025-06-10", 9e7)], ["W1", "W2"]),                   # 표시도 없고 금액도 다르면 다른 용역
            ([("W1", Wn, "2024-11-10", 5e7), ("W2", Wn + "(2차 공고)", "2024-12-10", 8e7)], ["W2"])):         # '(2차 공고)' 는 차수분이 아니라 다시 낸 공고
        got, g_, out_ = ys(items)
        assert got == want, (items, out_[["공고번호", "상태", "합산여부", "합산제외_까닭"]].to_dict("records"))
    got, g_, out_ = ys([("W1", Wn, "2024-11-10", ""), ("W2", Wn, "2025-06-10", 8e7)])
    assert got == ["W2"] and g_[("W1", "F1")]["상태"] == "재공고 이전", "금액이 한쪽에 없으면 같은 금액으로 본다"
    # 설계·건설사업관리의 '금액이 달라도 같은 용역' 은 같은 프로젝트일 때만: 같은 이름의 두 프로젝트(다시 지은 시설)에 하나씩 이어진 설계는 둘 다 남는다
    D = "나래 도서관 건립공사 기본 및 실시설계용역"
    got, g_, out_ = ys([("D1", D, "2020-06-01", 9e8), ("D2", D, "2021-09-01", 14e8)],
                       hist=[("F3", "나래 도서관", "F3-N-01", "C3", "나래 도서관 건립공사", "2021-03-02", I), ("F3", "나래 도서관", "F3-N-02", "C4", "나래 도서관 건립공사", "2023-06-01", I)],
                       rev=[("F3", "나래 도서관", I)])
    assert got == ["D1", "D2"] and g_[("D1", "F3")]["프로젝트ID"] == "F3-N-01" and g_[("D2", "F3")]["프로젝트ID"] == "F3-N-02", out_[["공고번호", "프로젝트ID", "상태"]].to_dict("records")
    # 일반 이름 규칙은 4글자 이하에만: 접두를 뗀 이름이 6글자('시민체육센터')면 넓은 창 그대로
    got, g_, out_ = ys([("G5", "시민체육센터 건립공사 건설사업관리용역", "2026-01-10", 5e7), ("G6", "주차타워 건립공사 건설사업관리용역", "2026-01-10", 5e7)],
                       hist=[("F5", "가온시 시민체육센터", "F5-N", "C5", "시민체육센터 건립공사", "2024-03-01", I), ("F4", "가온시 주차타워", "F4-N", "C6", "주차타워 건립공사", "2024-03-01", I)],
                       rev=[("F5", "가온시 시민체육센터", I), ("F4", "가온시 주차타워", I)])
    assert got == ["G5"] and "일반적인 시설 이름" not in g_[("G5", "F5")]["연결근거"] and "일반적인 시설 이름(4글자)" in g_[("G6", "F4")]["연결근거"], out_[["공고번호", "합산여부", "연결근거"]].to_dict("records")
    # 조달분류의 '재활용': 건설폐기물 말이 있는 공고에서만 눈감는다 / 건설폐기물 말이 없으면 '재활용'·'민간위탁' 은 운영 폐기물
    assert k("◇ 건립공사 건설폐기물 처리용역", "건설폐기물재활용서비스")[:2] == ("폐기물처리", "건설폐기물") and k("◇ 폐기물 처리용역", "생활폐기물재활용서비스")[0] == "기타"
    assert k("◇ 폐기물 재활용 처리용역")[0] == k("◇ 폐기물 수집운반 민간위탁 용역")[0] == "기타" and k("◇ 건립공사 건설폐기물 수집운반 민간위탁 용역")[0] == "폐기물처리"
    # 같은 계열의 다른 기관이 낸 공고: 공고명에 그 프로젝트의 사업 이름이 통째로 있으면 더한다(없으면 N — test_review_fixes)
    g, out, _ = _mini([("F1", "국가보훈처 봉안당", "F1-N", "C1", "국립대전현충원 봉안당 신축공사", "2022-03-02", "국가보훈처 국립대전현충원")],
                      [("F1", "국가보훈처 봉안당", "국가보훈처 국립대전현충원", "국가보훈처 국립대전현충원")],
                      [("S3", "국립대전현충원 봉안당 신축공사 건설사업관리용역", "2022-02-01", "국가보훈처 국립이천호국원", 300_000_000)], cfg)
    assert g[("S3", "F1")]["합산여부"] == "Y" and "같은 계열의 다른 기관" in g[("S3", "F1")]["연결근거"], g[("S3", "F1")].to_dict()
    # 요약: 같은 공고가 시설 후보 둘에 함께 Y 여도 건수·금액은 공고마다 한 번(줄 수는 괄호 안에)
    L = lambda pid, fid, no, price: dict({c: None for c in servc.LINK_COLS}, 프로젝트ID=pid, 시설ID=fid, 시설명=fid, 용역종류="설계", 용역세부="설계용역", 합산여부="Y",
                                         합산제외_까닭="", 공고번호=no, 추정가격=price, 공사대비_개월=-0.2)
    text = servc.summary(pd.DataFrame([L("A-N", "A", "S1", 300_000_000), L("B-N", "B", "S1", 300_000_000), L("A-N", "A", "S2", 100_000_000)]), 2, 2,
                         pd.DataFrame({"프로젝트ID": ["A-N", "B-N"]}))
    assert "· 설계: 2건(3줄), 프로젝트 2개, 추정가격 합 4.0억" in text and "중앙값 0개월" in text and "-0" not in text, text


def test_review_stage(tmp):
    """dedup 보다 오래된 연결 결과는 싣지 않는다 / 수집이 덜 됐거나 연결 뒤에 다시 수집했으면 알린다 / 요약의 건수는 [진단] 과 같은 식 / 요약 사본은 BOM 있는 UTF-8."""
    import openpyxl
    import time
    import tests.test_dedup_and_excel as T
    cfg = dict(load_config())
    reviewed, hist, latest, _ = _world(cfg)
    cfg["paths"] = dict(cfg["paths"], data_dir=os.path.join(tmp, "rv", "data"), out_dir=os.path.join(tmp, "rv", "out"), raw_dir=os.path.join(tmp, "rv", "raw"))
    for d in ("data_dir", "out_dir", "raw_dir"):
        os.makedirs(cfg["paths"][d], exist_ok=True)
    cfg["servc"] = dict(cfg.get("servc") or {}, period_start="2024-01", period_end="2024-03")
    data = lambda n: os.path.join(cfg["paths"]["data_dir"], n)
    raw = _rows(cfg, [dict(no="S01", name="가상시립미술관 건립공사 건설사업관리용역", dt="2024-01-20 10:00:00", price=900_000_000),
                      dict(no="", name="번호 없는 줄", dt="2024-01-20 10:00:00", price=1)])
    raw.to_parquet(data("servc_all.parquet"), index=False)
    hist.to_parquet(data("notices_hist.parquet"), index=False)
    latest.to_parquet(data("notices_latest.parquet"), index=False)
    agg = discover.discover_candidates(discover.standardize(pd.DataFrame(T.RAW), cfg, pd.DataFrame(T.BSIS)))
    discover.write_review_workbook(agg, os.path.join(cfg["paths"]["out_dir"], "facility_candidates.xlsx"))

    def run(fn):
        buf = io.StringIO()
        with contextlib.redirect_stdout(buf):
            fn(cfg)
        return buf.getvalue()
    with open(os.path.join(cfg["paths"]["raw_dir"], "servc_done.txt"), "w", encoding="utf-8") as fh:
        fh.write("202401\n")
    out = run(pipeline.stage_link_servc)
    assert "용역 공고 2줄(공고번호 기준 1건)" in out and "[진단] 용역 공고 2줄 → 공고번호 기준 1건" in out, out
    assert "[경고] 용역 공고 수집 미완료 월 2개 (202402, 202403)" in out and "collect --servc" in out, out
    with open(os.path.join(cfg["paths"]["out_dir"], "servc_link_summary.txt"), "rb") as fh:
        head = fh.read(3)
    assert head == b"\xef\xbb\xbf", "PowerShell 5.1 의 Get-Content 가 한글을 읽도록 BOM 을 붙인다"
    with open(os.path.join(cfg["paths"]["out_dir"], "servc_link_summary.txt"), encoding="utf-8-sig") as fh:
        assert "[경고] 용역 공고 수집 미완료 월 2개 (202402, 202403)" in fh.read(), "사용자가 보내 주는 요약 사본에도 미완료 경고가 있어야 한다"
    with open(os.path.join(cfg["paths"]["raw_dir"], "servc_done.txt"), "w", encoding="utf-8") as fh:
        fh.write("202401\n202402\n202403\n")
    assert "[경고] 용역 공고 수집 미완료" not in run(pipeline.stage_link_servc)
    book = lambda: openpyxl.load_workbook(os.path.join(cfg["paths"]["out_dir"], cfg["paths"]["excel_name"]))
    out = run(pipeline.stage_excel)
    assert build_excel.S_SERVC in book().sheetnames and "[주의]" not in out and "다시 수집했습니다" not in out, out
    fut = time.time() + 120
    os.utime(data("servc_all.parquet"), (fut, fut))                 # 연결 뒤에 용역 공고를 다시 수집함 → 싣기는 하되 알린다
    out = run(pipeline.stage_excel)
    assert "[안내] link-servc 뒤에 용역 공고를 다시 수집했습니다" in out and build_excel.S_SERVC in book().sheetnames, out
    os.utime(data("notices_latest.parquet"), (fut, fut))            # 연결 뒤에 dedup 을 다시 돌림 → 프로젝트ID 가 다른 사업을 가리킬 수 있으므로 싣지 않는다
    out = run(pipeline.stage_excel)
    wb = book()
    assert "[주의] dedup 을 link-servc 보다 나중에 실행했습니다" in out and "싣지 않습니다" in out and build_excel.S_SERVC not in wb.sheetnames, out
    assert "python -m g2b_costdb.pipeline link-servc 를 실행한 뒤 excel 을 다시" in out, "무엇을 실행할지 알린다"
    assert not any("용역_" in str(c.value) or c.value in build_excel.SERVC5_COLS for c in wb[build_excel.S_FAC][1]), "05 의 용역 참고 열(발주완료 포함)도 만들지 않는다"


# ── 2차 검토(2026-11-23)에서 고친 자리: 고치기 전 코드에서는 실패한다. 모든 공고·기관·시설은 가상의 예시 ──
_R2 = {"cfg": None}
I = "경기도 가온시"


def _ys(g):
    return sorted(no for (no, _), r in g.items() if r["합산여부"] == "Y")


def _test_R1_wide_weak():
    """같은 계열의 다른 기관: 사업 이름이 '접두를 뗀 시설 이름 + 신축' 뿐이면 통째로 있어도 더하지 않는다 / 그런 줄은 제 기관의 공고를 '재공고 이전' 으로 만들지 않는다."""
    E1, E2 = "경상북도교육청 가온고등학교", "경상북도교육청 나래고등학교"
    g, out, _ = _mini([("F1", "경상북도교육청 스마트팜온실", "F1-N", "C1", "스마트팜 온실 신축공사", "2023-06-01", E1)], [("F1", "경상북도교육청 스마트팜온실", E1)],
                         [("S1", "스마트팜 온실 신축공사 설계용역", "2022-09-01", E2, 80_000_000), ("S2", "스마트팜 온실 신축공사 건설폐기물 처리용역", "2024-01-10", E2, 9_000_000),
                          ("S3", "나래고 스마트팜 온실 신축공사 감리용역", "2023-05-10", E2, 30_000_000), ("S4", "스마트팜 온실 신축공사 설계용역", "2022-08-01", E1, 70_000_000)], _R2["cfg"])
    assert _ys(g) == ["S4"], out[["공고번호", "수요기관", "상태", "합산여부", "합산제외_까닭"]].to_dict("records")
    assert g[("S1", "F1")]["상태"] == g[("S4", "F1")]["상태"] == "대표" and "같은 계열의 다른 기관" in g[("S1", "F1")]["합산제외_까닭"]
    E1, E2 = "경기도교육청 가온고등학교", "경기도교육청 나래고등학교"
    g, out, _ = _mini([("F1", "경기도교육청 유리온실", "F1-N", "C1", "가온고 유리온실 신축공사", "2023-06-01", E1)], [("F1", "경기도교육청 유리온실", E1)],
                         [("S1", "유리온실 신축 설계용역", "2022-08-01", E1, 70_000_000), ("S2", "유리온실 신축 설계용역", "2023-04-01", E2, 40_000_000),
                          ("A1", "유리온실 신축공사 감리용역", "2023-06-10", E1, 30_000_000), ("A2", "유리온실 신축공사 감리용역", "2024-09-10", E2, 20_000_000)], _R2["cfg"])
    assert _ys(g) == ["A1", "S1"], out[["공고번호", "수요기관", "상태", "합산여부", "합산제외_까닭"]].to_dict("records")
    # 여러 지사가 같은 이름의 사업을 하는 경우(검수한 이름이 그대로 있고 사업 이름도 통째로 있지만 공고명에 다른 현장 이름 '한빛지구' 가 있다)
    A, B, HQ = "한국가온공사 충남지역본부 나래지사", "한국가온공사 경남지역본부 다온지사", "한국가온공사"
    g, out, _ = _mini([("F1", "수출전문 스마트팜 온실", "F1-N", "C1", "수출전문 스마트팜 온실 신축공사", "2023-06-01", A, "신축", 5_000_000_000)], [("F1", "수출전문 스마트팜 온실", A)],
                         [("S1", "한빛지구 수출전문 스마트팜 온실 신축공사 건설사업관리용역", "2023-05-01", B, 300_000_000), ("S2", "한빛지구 수출전문 스마트팜 온실 신축공사 기본 및 실시설계용역", "2021-12-01", B, 200_000_000),
                          ("S4", "수출전문 스마트팜 온실 신축공사 건설사업관리용역", "2023-05-10", A, 250_000_000), ("S5", "수출전문 스마트팜 온실 신축공사 실시설계용역", "2021-11-01", HQ, 180_000_000)], _R2["cfg"])
    assert _ys(g) == ["S4", "S5"] and "가릴 수 없음" in g[("S1", "F1")]["합산제외_까닭"], out[["공고번호", "수요기관", "합산여부", "합산제외_까닭"]].to_dict("records")
    # 다른 소속기관이 낸 공고라도 이 프로젝트의 (현장을 가리는) 사업 이름이 통째로 있으면 그대로 더한다
    g, out, _ = _mini([("F1", "국가보훈처 봉안당", "F1-N", "C1", "국립대전현충원 봉안당 신축공사", "2022-03-02", "국가보훈처 국립대전현충원")], [("F1", "국가보훈처 봉안당", "국가보훈처 국립대전현충원")],
                         [("S8", "국립대전현충원 봉안당 신축공사 감리용역", "2022-03-10", "국가보훈처 국립이천호국원", 50_000_000)], _R2["cfg"])
    assert _ys(g) == ["S8"], out[["공고번호", "수요기관", "합산여부", "합산제외_까닭"]].to_dict("records")


def _test_R2_wtsvc():
    """'건설혼합폐기물'·'건설오니' 의 '건설' 도 신축 어휘가 아니다."""
    items = [("M1", "가온회관 외벽 도장공사 건설혼합폐기물 처리용역", "2025-04-01"), ("M2", "가온회관 건설혼합폐기물 처리용역", "2025-04-08"), ("M3", "가온회관 건설오니 처리용역", "2025-04-15"),
             ("M4", "가온회관 건립공사 건설혼합폐기물 처리용역", "2023-09-01"), ("M5", "가온회관 건립공사 건설오니 처리용역", "2023-09-08")]
    g, out, _ = _mini([("F1", "가온회관", "F1-N", "C1", "가온회관 건립공사", "2023-03-02", I)], [("F1", "가온회관", I)], [(no, t, dt, I, 10_000_000 + i) for i, (no, t, dt) in enumerate(items)], _R2["cfg"])
    assert _ys(g) == ["M4", "M5"], out[["공고번호", "공고명", "합산여부", "합산제외_까닭"]].to_dict("records")
    assert g[("M1", "F1")]["합산제외_까닭"].startswith("유지보수") and g[("M3", "F1")]["합산제외_까닭"].startswith("어느 공사에서 나온 폐기물인지")


def _test_R3_lots():
    """폐기물 차수분 사이에 낀 재공고는 제 차수의 앞 공고를 대체한다 / '(2차 공고)' 는 다시 낸 공고 / 프로젝트를 정하지 못한 줄에는 한 번짜리 용역 규칙을 쓰지 않는다."""
    W = "가온회관 건립공사 건설폐기물 처리용역"
    H, R = [("F1", "가온회관", "F1-N", "C1", "가온회관 건립공사", "2024-09-02", I, "신축", 30_000_000_000)], [("F1", "가온회관", I)]
    for items, want in (
            ([("W1", W + "(1차)", "2024-11-10", 8e7), ("W2", W + "(2차)", "2024-12-10", 5e7), ("W3", W + "(1차)", "2024-12-20", 8e7, dict(kind="재공고"))], ["W2", "W3"]),
            ([("W1", W + "(1차)", "2024-11-10", 8e7), ("W2", W + "(2차)", "2024-12-10", 5e7), ("W3", W + "(1차)", "2024-12-20", 8e7)], ["W2", "W3"]),
            ([("W1", W + "(1차)", "2024-11-10", 5e7), ("W2", W + "(2차)", "2024-12-10", 5e7), ("W3", W + "(1차)", "2024-12-20", 5e7)], ["W3"]),     # 금액이 모두 같으면 예전처럼 마지막 한 건
            ([("W1", W + "(1차)", "2024-11-10", 5e7), ("W2", W + "(2차 공고)", "2024-12-10", 6.5e7)], ["W2"]),
            ([("W1", W + "(1차)", "2024-11-10", 5e7), ("W2", W + "(2차)", "2025-01-10", 8e7), ("W3", W + "(3차)", "2025-03-10", 3e7)], ["W1", "W2", "W3"])):
        g, out, _ = _mini(H, R, [(x[0], x[1], x[2], I, x[3], (x[4] if len(x) > 4 else {})) for x in items], _R2["cfg"])
        assert _ys(g) == want, (items, out[["공고번호", "상태", "합산여부", "합산제외_까닭"]].to_dict("records"))
    g, out, _ = _mini([("F1", "가온회관", "F1-N-01", "C1", "가온회관 건립공사", "2024-03-02", I), ("F1", "가온회관", "F1-N-02", "C2", "가온회관 건립공사", "2024-05-02", I)], R,
                         [("D1", "가온회관 건립공사 실시설계용역", "2022-03-10", I, 3e8), ("D2", "가온회관 건립공사 실시설계용역", "2022-11-12", I, 6e8)], _R2["cfg"])
    assert g[("D1", "F1")]["상태"] == "대표" and g[("D1", "F1")]["합산제외_까닭"].startswith("프로젝트를 정하지 못함"), g[("D1", "F1")].to_dict()


def _test_R4_rename():
    """광역 이름이 바뀐 같은 교육지원청·도 사업소는 '바로 그 기관' 이다."""
    nr = servc.inst_narrow
    assert nr("강원도교육청 강원도가온교육지원청", "강원특별자치도교육청 강원특별자치도가온교육지원청") and nr("전라북도 농업기술원", "전북특별자치도 농업기술원")
    assert not nr("강원도교육청 강원도가온교육지원청", "강원특별자치도교육청 강원특별자치도나래교육지원청") and not nr("서울특별시 상수도사업본부", "서울특별시 한강사업본부")
    assert not nr("경기도교육청 경기도가온교육지원청", "경기도교육청 가온고등학교") and nr("경기도교육청", "경기도교육청 가온고등학교")
    P_OLD, P_NEW = "전라북도 농업기술원", "전북특별자치도 농업기술원"
    g, out, _ = _mini([("F2", "전북 스마트팜 혁신밸리 온실", "F2-N", "C2", "스마트팜 혁신밸리 온실 신축공사", "2024-06-04", P_NEW)], [("F2", "전북 스마트팜 혁신밸리 온실", P_NEW)],
                         [("T1", "전북 스마트팜 혁신밸리 온실 실시설계용역", "2023-02-01", P_OLD, 200_000_000)], _R2["cfg"])
    assert g[("T1", "F2")]["합산여부"] == "Y" and "같은 발주 기관" in g[("T1", "F2")]["연결근거"], g[("T1", "F2")].to_dict()


def _test_R5_spans():
    """묶음 규칙: 괄호 속을 빼야 이름이 맞는 시설·느슨한 꼴로 맞은 시설이 낀 묶음 공고도 모두 N."""
    g, out, _ = _mini([("F1", "가온센터", "F1-N", "C1", "가온(국민체육)센터 건립공사", "2024-03-01", I), ("F2", "나래 도서관", "F2-N", "C2", "나래 도서관 건립공사", "2024-04-01", I),
                          ("F3", "다온 국민체육센터", "F3-N", "C3", "다온 국민체육센터 건립공사", "2024-03-01", I)],
                         [("F1", "가온센터", I), ("F2", "나래 도서관", I), ("F3", "다온 국민체육센터", I)],
                         [("B1", "가온(국민체육)센터 및 나래 도서관 건립 설계용역", "2023-01-10", I, 9e8), ("B2", "나래 도서관 및 가온(국민체육)센터 건립공사 건설사업관리용역", "2024-02-10", I, 7e8),
                          ("B3", "다온 반다비 국민체육센터 및 나래 도서관 건립 설계용역", "2023-01-12", I, 9e8), ("S1", "가온(국민체육)센터 건립공사 설계용역", "2023-01-20", I, 4e8)], _R2["cfg"])
    assert _ys(g) == ["S1"], out[["공고번호", "시설ID", "연결강도", "합산여부", "합산제외_까닭"]].to_dict("records")
    assert all("묶음 발주" in g[k]["합산제외_까닭"] for k in (("B1", "F2"), ("B2", "F1"), ("B3", "F2")))


def _test_R6_demolition():
    k = servc.classify_kind
    for t in ("◇ 석면해체·제거공사(리모델링 구간) 감리용역", "◇ 철거공사(신축공사 사전) 감리용역", "◇ 해체공사 및 건립부지 조성공사 감리용역"):
        assert k(t)[1] == "감리(석면·해체)" and k(t)[2], (t, k(t))
    for t in ("◇ 철거 및 신축공사 감리용역", "◇ 리모델링공사(철거 포함) 감리용역", "(구)가온청사 철거부지 ◇ 건립공사 감리용역", "◇ 개축공사 감리용역 (철거포함)", "◇ 신축공사 감리용역(해체공사 포함)"):
        assert k(t)[1].startswith("감리") and "석면" not in k(t)[1] and k(t)[2] == "", (t, k(t))


def test_review_round2_fixes():
    _R2["cfg"] = load_config()
    _test_R1_wide_weak()
    _test_R2_wtsvc()
    _test_R3_lots()
    _test_R4_rename()
    _test_R5_spans()
    _test_R6_demolition()


def test_field_audit():
    """실제 용역 공고 41만 건으로 한 점검(2026-11-24~25, 연결 3,771줄 가운데 2,400여 줄을 읽어 판정)에서 나온 결함의 재현 사례. 모든 이름은 가상."""
    cfg = load_config()
    k, rk, I = servc.classify_kind, servc.rebid_key, "경기도 가온시"
    yn = lambda g, nos, f="F1": [g[(n, f)]["합산여부"] for n in nos]
    dump = lambda out: out[["공고번호", "프로젝트ID", "합산여부", "합산제외_까닭", "상태", "연결근거"]].to_dict("records")

    # ── ① 평가 안내(PQ) 공고와 본 공고(가격입찰·지명경쟁), 수의시담 공고가 같은 프로젝트에 둘 다 더해짐(실자료 110쌍·1,292억) ──
    base = rk("◇ 건립공사 건설사업관리용역")
    for t in ("◇ 건립공사 감독권한대행 등 건설사업관리용역 사업수행능력평가서 제출 안내 공고", "◇ 건립공사 건설사업관리용역(PQ후 가격입찰)", "[P.Q 후 가격입찰]◇ 건립공사 건설사업관리용역",
              "◇ 건립공사 건설사업관리용역 전자 수의시담 공고", "용역사업 집행계획 및 사업수행능력평가서 제출 안내 공고(긴급)[◇ 건립공사 건설사업관리용역]", "[PQ제출안내]◇ 건립공사 건설사업관리용역(긴급)",
              "◇ 건립공사 건설사업관리용역(P. Q, 후 가격입찰)", "◇ 건립공사 건설사업관리용역[지명경쟁]", "◇ 건립공사 건설사업관리용역 (사업수행능력 평가자료 제출)"):
        assert rk(t) == base, (t, rk(t), base)
    assert rk("◇ 건립공사 전기공사 감리용역(PQ)") != rk("◇ 건립공사 소방공사 감리용역(PQ)") and rk("◇ 건립공사 설계용역(PQ)") != rk("◇ 증축공사 설계용역(PQ)")
    g, out, _ = _mini([("F1", "가온미술관", "F1-N", "C1", "가온미술관 건립공사", "2023-03-02", I, "신축", 20_000_000_000)], [("F1", "가온미술관", I)],
                      [("P1", "가온미술관 건립공사 감독권한대행 등 건설사업관리용역 사업수행능력평가서 제출 안내 공고", "2023-01-05", I, 1_680_000_000),
                       ("P2", "가온미술관 건립공사 감독권한대행 등 건설사업관리용역(PQ후 가격입찰)", "2023-02-01", I, 1_680_000_000),
                       ("P3", "가온미술관 건립공사 건설사업관리용역 전자 수의시담 공고", "2023-02-20", I, 1_680_000_000),
                       ("P4", "[P.Q제출안내] 가온미술관 건립 기본 및 실시설계용역", "2021-03-02", I, 500_000_000),
                       ("P5", "가온미술관 건립 기본 및 실시설계용역[지명경쟁]", "2021-04-01", I, 500_000_000),
                       ("P6", "용역사업 집행계획 및 사업수행능력평가서 제출 안내 공고(긴급)[가온미술관 건립공사 전기공사 감리용역]", "2023-02-10", I, 90_000_000),
                       ("P7", "가온미술관 건립공사 전기공사 감리용역(긴급)", "2023-03-15", I, 90_000_000),
                       ("P8", "가온미술관 건립공사 소방공사 감리용역", "2023-03-16", I, 90_000_000),                     # 금액이 같아도 공종이 다르면 다른 용역
                       ("Q1", "가온미술관 본관 건립사업 건축감리 용역", "2023-03-20", I, 300_000_000),                   # 이름을 조금 바꿔 다시 낸 공고(같은 금액·20일 뒤)
                       ("Q2", "가온미술관 건립공사 건축 감리용역", "2023-04-09", I, 300_000_000),
                       ("Q3", "가온미술관 건립공사(1단계) 통신공사 감리용역", "2023-05-01", I, 40_000_000),             # 단계가 다르면 금액이 같아도 다른 용역
                       ("Q4", "가온미술관 건립공사(2단계) 통신공사 감리용역", "2023-05-20", I, 40_000_000)], cfg)
    assert yn(g, ("P1", "P2", "P3", "P4", "P5", "P6", "P7", "P8", "Q1", "Q2", "Q3", "Q4")) == ["N", "N", "Y", "N", "Y", "N", "Y", "Y", "N", "Y", "Y", "Y"], dump(out)
    assert all(g[(n, "F1")]["상태"] == "재공고 이전" for n in ("P1", "P2", "P4", "P6", "Q1")), dump(out)
    y = out[out["합산여부"] == "Y"].groupby("용역종류")["추정가격"].sum().to_dict()
    assert y == {"건설사업관리·감리": 1_680_000_000 + 90_000_000 * 2 + 300_000_000 + 40_000_000 * 2, "설계": 500_000_000}, y

    # ── ② 지자체·기관 접두를 뗀 짧은 이름('공공주택')으로만 이어진 공고: 공고명에 그 프로젝트의 사업 이름이 통째로 있을 때만 더한다 ──
    #    (실자료: 사업 이름 없이 시점만으로 더한 17줄 가운데 10줄이 다른 단지·다른 사업, 사업 이름이 있는데 시점이 멀다고 뺀 13줄 가운데 11줄은 맞는 용역)
    S = "가온주택도시공사"
    g, out, _ = _mini([("F1", "가온주택도시공사 공공주택", "F1-N-01", "C1", "나래지구 3단지 공공주택 건설공사", "2023-03-02", S, "신축", 30_000_000_000),
                       ("F1", "가온주택도시공사 공공주택", "F1-N-02", "C2", "다솜지구 공공주택 건설공사", "2020-03-02", S, "신축", 30_000_000_000)],
                      [("F1", "가온주택도시공사 공공주택", S)],
                      [("G1", "마루지구 9단지 공공주택 건설공사 건설폐기물 처리용역", "2023-04-01", S, 30_000_000),          # 다른 단지 — 시점만 맞음
                       ("G2", "나래지구 3단지 공공주택 건설공사 건설폐기물 처리용역", "2023-05-01", S, 40_000_000),
                       ("G3", "다솜지구 공공주택 건설공사 감독권한대행 등 건설사업관리용역", "2022-06-01", S, 900_000_000)], cfg)   # 사업 이름이 통째로 — 공사 공고 27개월 뒤여도
    assert yn(g, ("G1", "G2", "G3")) == ["N", "Y", "Y"], dump(out)
    assert "일반적인 시설 이름" in g[("G1", "F1")]["합산제외_까닭"] and g[("G2", "F1")]["프로젝트ID"] == "F1-N-01" and g[("G3", "F1")]["프로젝트ID"] == "F1-N-02", dump(out)

    # ── ③ 'A 및 B … 통합 건설사업관리·감리': 묶인 상대(A)가 이 프로젝트의 공사 이름에 없으면 금액을 나눌 수 없다 ──
    g, out, _ = _mini([("F1", "가온동 공영주차장", "F1-N", "C1", "가온동 공영주차장 건립공사", "2023-03-02", I, "신축", 10_000_000_000),
                       ("F2", "나래체육관", "F2-E", "C2", "나래체육관 및 급식소 증축공사", "2023-03-02", I, "증축", 5_000_000_000)],
                      [("F1", "가온동 공영주차장", I), ("F2", "나래체육관", I)],
                      [("B1", "마루공원 및 가온동 공영주차장 건립공사 통합건설사업관리용역", "2023-02-01", I, 2_000_000_000),
                       ("B2", "가온동 공영주차장 건립공사 통합건설사업관리용역(전기·통신·소방)", "2023-02-10", I, 300_000_000),     # '통합' 만으로는 묶음이 아니다
                       ("B3", "다솜동·가온동 공영주차장 전기공사 통합감리용역", "2023-04-01", I, 100_000_000),
                       ("B4", "나래체육관 및 급식소 증축공사 통합건설사업관리용역", "2023-02-01", I, 400_000_000)], cfg)          # 공사 이름 자체가 '○○ 및 △△'
    assert [g[("B1", "F1")]["합산여부"], g[("B2", "F1")]["합산여부"], g[("B3", "F1")]["합산여부"], g[("B4", "F2")]["합산여부"]] == ["N", "Y", "N", "Y"], dump(out)
    assert "묶" in g[("B1", "F1")]["합산제외_까닭"] and "묶" in g[("B3", "F1")]["합산제외_까닭"], dump(out)

    # ── ④ 조달분류가 '폐기물재활용서비스'·'생활폐기물처리서비스' 로 붙은 공사 폐기물이 운영 폐기물로 읽혀 빠짐 ──
    for t, c in (("◇ 리모델링사업(건축) 폐기물처리 용역", "폐기물재활용서비스"), ("◇ 개보수 공사 가연성폐기물 처리용역", "생활폐기물처리서비스"),
                 ("◇ 조성사업 폐목재재활용 처리용역", "폐기물재활용서비스"), ("◇ 건립공사 폐판넬 처리용역", "폐기물재활용서비스"),
                 ("◇ 건립사업 매립폐기물 처리용역", "생활폐기물처리서비스"), ("◇ 조성사업 인조잔디재활용처리 용역", "폐기물재활용서비스")):
        assert k(t, c)[0] == "폐기물처리" and not k(t, c)[2], (t, c, k(t, c))
    for t, c in (("2026년 ◇ 폐기물 수집·운반·처리 용역", "생활폐기물처리서비스"), ("◇ 사업장일반폐기물 수거 용역", "생활폐기물처리서비스"),
                 ("◇ 인조잔디 분리회수 및 재활용 처리용역", "폐기물재활용서비스"), ("◇ 신축공사 현장 생활폐기물 처리용역", "생활폐기물처리서비스")):
        assert k(t, c)[0] == "기타", (t, c, k(t, c))

    # ── ⑤ 설계 용역에 딸린 일(측량·기본계획·컨설팅)이 함께 적힌 설계, 조달분류가 설계인 제안공모가 기타로 빠짐 ──
    for t in ("◇ 조성사업 조사측량 및 실시설계 용역", "◇ 건립공사 기본계획 및 실시설계 용역", "◇ 조성 공간컨설팅, 실시설계 및 용도변경 용역"):
        assert k(t)[0] == "설계" and not k(t)[2], (t, k(t))
    for t in ("◇ 건립 기본계획 수립 용역", "◇ 실시설계 안전성검토 용역", "◇ 건립사업 설계VE 용역", "◇ 건립 타당성조사 및 기본계획 수립 용역", "◇ 건립공사 기본계획 및 실시설계 VE 용역"):
        assert k(t)[0] == "기타", (t, k(t))
    assert k("◇ 리모델링 제안공모 공고", "건축설계용역")[:2] == ("설계", "설계공모") and k("◇ 위탁운영 수탁기관 선정 제안공모", "기타사업지원서비스")[0] == "기타"
    # 부대 용역 낱말이 시설 이름·지명 속의 말('홍보관'·'시험인증센터'·'적산가옥'·'재해예방센터')이나 '(공모 대행)' 머리말에 걸려 설계·감리가 기타로 빠짐
    for t, kind in (("가온 홍보관 건립공사 설계용역", "설계"), ("가온시험인증센터 신축공사 감독권한대행 등 건설사업관리용역", "건설사업관리·감리"), ("가온 적산가옥 보수복원공사 실시설계용역", "설계"),
                    ("(공모 대행) ◇ 증축공사 건축설계 제안공모", "설계"), ("◇ 다목적체육관 건축기획 설계용역", "설계"), ("가온 산림재해예방센터 신축공사 감리용역", "건설사업관리·감리")):
        assert k(t)[0] == kind and not k(t)[2], (t, k(t))
    for t in ("◇ 건립 홍보영상 제작 및 설계 용역", "◇ 녹색건축 인증 용역(설계단계)", "◇ 건립 설계공모 관리용역", "◇ 신축공사 적산 용역(설계)", "◇ 건립 설계공모 대행 용역", "◇ 신축공사 감리용역(BF인증)"):
        assert k(t)[0] == "기타", (t, k(t))
    # 기관: '○○광역시 동구청' 은 자치구다(광역 사업소가 아니다)
    assert servc.inst_same("인천광역시 동구청", "인천광역시 동구") and not servc.inst_same("인천광역시 문화예술회관", "인천광역시 동구청") and servc.inst_same("수원시청", "경기도 수원시")

    # ── ⑥ '추정' 연결(어절 사이에 다른 말이 낀 이름 / 그 지자체의 대행 기관)이라도 공고명에 프로젝트의 사업 이름이 통째로 있으면 더한다 ──
    g, out, _ = _mini([("F1", "가온 센터 공영주차장", "F1-N", "C1", "가온혁신센터 공영주차장 건립사업", "2023-03-02", I, "신축", 3_000_000_000),
                       ("F2", "나래동 지하 공영주차장", "F2-N", "C2", "나래동 지하 공영주차장 건립공사", "2023-03-02", "가온도시공사", "신축", 9_000_000_000)],
                      [("F1", "가온 센터 공영주차장", I), ("F2", "나래동 지하 공영주차장", "가온도시공사")],
                      [("E1", "가온혁신센터 공영주차장 건립사업 건설사업관리용역", "2023-04-01", I, 170_000_000),
                       ("E2", "가온나눔센터 뒤편 공영주차장 조성 실시설계용역", "2022-04-01", I, 20_000_000),                # 느슨한 이름뿐 — 다른 주차장일 수 있다
                       ("E3", "나래동 지하 공영주차장 건립 실시설계용역", "2021-06-01", I, 430_000_000),                    # 대행 관계의 기관이 낸, 사업 이름이 같은 공고
                       ("E4", "나래동 지하 공영주차장 주변 보도 정비 실시설계용역", "2021-07-01", I, 30_000_000)], cfg)
    assert [g[("E1", "F1")]["합산여부"], g[("E2", "F1")]["합산여부"], g[("E3", "F2")]["합산여부"], g[("E4", "F2")]["합산여부"]] == ["Y", "N", "Y", "N"], dump(out)
    assert g[("E1", "F1")]["연결강도"] == "추정" and g[("E3", "F2")]["연결강도"] == "추정" and "추정" in g[("E2", "F1")]["합산제외_까닭"], dump(out)

    # ── ⑦ 용역 추정가격이 그 프로젝트의 공사 추정가격 합과 같다(공사비를 잘못 적은 공고) ──
    g, out, _ = _mini([("F1", "가온복합청사", "F1-N", "C1", "가온복합청사 건립공사", "2023-12-01", I, "신축", 15_191_000_000)], [("F1", "가온복합청사", I)],
                      [("A1", "가온복합청사 건립사업 설계용역", "2021-05-31", I, 15_192_000_000), ("A2", "가온복합청사 건립공사 건설사업관리용역", "2023-11-01", I, 1_200_000_000)], cfg)
    assert yn(g, ("A1", "A2")) == ["N", "Y"] and "공사 추정가격" in g[("A1", "F1")]["합산제외_까닭"], dump(out)

    # ── ⑧ 공고명에 사업 이름이 없는 설계 공고가 그 공사 공고의 직전·뒤에 났다 — 그 공사의 설계일 수 없다(다음 단계·다른 사업의 설계) ──
    g, out, _ = _mini([("F1", "가온체육공원", "F1-R", "C1", "가온체육공원 롤러경기장 개보수공사", "2024-03-17", I, "리모델링", 800_000_000)], [("F1", "가온체육공원", I)],
                      [("T1", "가온체육공원 시설물 개보수사업 실시설계용역", "2024-02-20", I, 52_000_000),                  # 공사 25일 전
                       ("T2", "가온체육공원 롤러경기장 개보수공사 실시설계용역", "2023-08-01", I, 40_000_000),              # 사업 이름이 통째로
                       ("T3", "가온체육공원 개보수 설계용역", "2023-06-01", I, 30_000_000),                             # 사업 이름은 없지만 공사 9개월 전 — 프로젝트 하나
                       ("T4", "가온체육공원 관리동 개보수 실시설계용역", "2024-06-01", I, 20_000_000)], cfg)               # 공사 뒤
    assert yn(g, ("T1", "T2", "T3", "T4")) == ["N", "Y", "Y", "N"] and "설계 공고" in g[("T1", "F1")]["합산제외_까닭"], dump(out)

    # ── ⑨ 공고명의 사업유형 말('건립' = 신축)이 공사 쪽 분류(증축)와 달라 프로젝트를 정하지 못함: 후보가 하나뿐이고 사업 이름이 거의 같으면 그 프로젝트 ──
    g, out, _ = _mini([("F1", "가온 시민주차타워", "F1-E", "C1", "가온 시민주차타워 확장 건립사업 건축공사", "2023-03-02", I, "증축", 5_000_000_000)], [("F1", "가온 시민주차타워", I)],
                      [("X1", "가온 시민주차타워 건립사업 전기공사 감리용역", "2023-03-10", I, 33_000_000),
                       ("X2", "가온 시민주차타워 리모델링 실시설계용역", "2022-06-01", I, 20_000_000)], cfg)                   # 사업 이름이 다르면(유사도 0.8 미만) 그대로 정하지 않는다
    assert yn(g, ("X1", "X2")) == ["Y", "N"] and g[("X1", "F1")]["프로젝트ID"] == "F1-E" and "유사도" in g[("X1", "F1")]["연결근거"] and g[("X2", "F1")]["프로젝트ID"] == "", dump(out)

    # ── ⑩ 공사 이름 없는 '건설폐기물' 공고: 그 시설의 프로젝트가 하나뿐이고 공사 공고 무렵이면 그 공사의 폐기물(실자료: 이 까닭으로 뺀 줄의 22% 가 맞는 용역) ──
    g, out, _ = _mini([("F1", "가온회관", "F1-N", "C1", "가온회관 건립공사", "2023-03-02", I, "신축", 8_000_000_000)], [("F1", "가온회관", I)],
                      [("W1", "가온회관 건설폐기물 처리용역", "2023-06-01", I, 30_000_000), ("W2", "가온회관 폐기물 처리용역", "2023-06-02", I, 15_000_000),          # 무슨 폐기물인지 모름
                       ("W3", "가온회관 폐아스콘 처리용역", "2025-12-01", I, 16_000_000)], cfg)                                                           # 공사 공고 33개월 뒤
    assert yn(g, ("W1", "W2", "W3")) == ["Y", "N", "N"] and g[("W2", "F1")]["합산제외_까닭"].startswith("어느 공사에서 나온 폐기물인지"), dump(out)
    g, out, _ = _mini([("F1", "가온회관", "F1-N", "C1", "가온회관 건립공사", "2023-03-02", I), ("F1", "가온회관", "F1-R", "C2", "가온회관 리모델링공사", "2023-09-01", I, "리모델링")], [("F1", "가온회관", I)],
                      [("W1", "가온회관 건설폐기물 처리용역", "2023-06-01", I, 30_000_000)], cfg)                                                         # 프로젝트가 둘이면 가릴 수 없다
    assert yn(g, ("W1",)) == ["N"], dump(out)


def test_field_audit_edges():
    """test_field_audit 의 규칙들이 놓인 테두리(변이 시험에서 살아남은 변이를 잡는 사례 — 2026-11-25 독립 검토). 모든 이름은 가상."""
    cfg = load_config()
    k, rk, I, S = servc.classify_kind, servc.rebid_key, "경기도 가온시", "가온주택도시공사"
    yn = lambda g, nos, f="F1": [g[(n, f)]["합산여부"] for n in nos]
    C = {}
    ART = ([("F1", "가온미술관", "F1-N", "C1", "가온미술관 건립공사", "2023-03-02", I, "신축", 20_000_000_000)], [("F1", "가온미술관", I)])
    TOWER = ([("F4", "가온시 주차타워", "F4-N", "C5", "주차타워 건립공사", "2024-03-01", I)], [("F4", "가온시 주차타워", I)])
    LOT = ([("F1", "가온동 공영주차장", "F1-N", "C1", "가온동 공영주차장 건립공사", "2023-03-02", I, "신축", 10_000_000_000)], [("F1", "가온동 공영주차장", I)])
    HALL = ([("F1", "가온복합청사", "F1-N", "C1", "가온복합청사 건립공사", "2023-12-01", I, "신축", 15_191_000_000)], [("F1", "가온복합청사", I)])
    PARK = ([("F1", "가온체육공원", "F1-R", "C1", "가온체육공원 롤러경기장 개보수공사", "2024-03-17", I, "리모델링", 800_000_000)], [("F1", "가온체육공원", I)])
    CLUB = ([("F1", "가온회관", "F1-N", "C1", "가온회관 건립공사", "2023-03-02", I, "신축", 8_000_000_000)], [("F1", "가온회관", I)])
    PT = ([("F1", "가온 시민주차타워", "F1-E", "C1", "가온 시민주차타워 확장 건립사업 건축공사", "2023-03-02", I, "증축", 5_000_000_000)], [("F1", "가온 시민주차타워", I)])

    # ── ① rebid_key: 입찰 단계 말의 나머지 갈래 ──
    base = rk("◇ 건립공사 건설사업관리용역")
    same = lambda *ts: all(rk(t) == base for t in ts)
    C["F1f"] = lambda: same("◇ 건립공사 건설사업관리용역(가격입찰)", "◇ 건립공사 건설사업관리용역 (평가 후 가격입찰)")
    C["F1g"] = lambda: same("◇ 건립공사 건설사업관리용역 평가자료 제출 안내 공고", "◇ 건립공사 건설사업관리용역 (평가자료(PQ) 제출)")
    C["F1h"] = lambda: same("◇ 건립공사 건설사업관리용역(지명)", "◇ 건립공사 건설사업관리용역 지명공고", "[지명] ◇ 건립공사 건설사업관리용역")
    C["F1j"] = lambda: same("◇ 건립공사 건설사업관리용역(PQ 후 입찰)", "◇ 건립공사 건설사업관리용역(PQ, 가격)", "◇ 건립공사 건설사업관리용역(PQ 입찰)")


    def _pq_pair():      # 금액을 고쳐(1% 넘게) 다시 낸 본 공고: 이름 열쇠가 같아야만 앞 공고가 '재공고 이전' 이 된다(같은 금액 규칙으로는 못 잡는다)
        g, out, _ = _mini(*ART, [("P1", "가온미술관 건립공사 건설사업관리용역 사업수행능력평가서 제출 안내 공고", "2023-01-05", I, 1_680_000_000),
                                 ("P2", "가온미술관 건립공사 건설사업관리용역(가격입찰)", "2023-02-01", I, 1_720_000_000)], cfg)
        return yn(g, ("P1", "P2")) == ["N", "Y"] and g[("P1", "F1")]["상태"] == "재공고 이전"


    C["F1f_e2e"] = _pq_pair


    # ── ② _mark_status 의 '같은 금액·120일' 규칙의 테두리 ──
    def two(raw, world=ART, f="F1"):
        g, out, _ = _mini(*world, raw, cfg)
        return [(g[(x[0], f)]["합산여부"], g[(x[0], f)]["상태"]) for x in raw]


    YY, NY = [("Y", "대표"), ("Y", "대표")], [("N", "재공고 이전"), ("Y", "대표")]
    C["F2b"] = lambda: two([("D1", "가온미술관 건립 기본 및 실시설계용역", "2021-03-02", I, 500_000_000), ("D2", "가온미술관 건립 수장고 실시설계용역", "2021-04-01", I, 450_000_000)]) == YY      # 금액이 10% 다르면 다른 용역
    C["F2c"] = lambda: two([("Q1", "가온미술관 본관 건립사업 건축감리 용역", "2023-03-20", I, 300_000_000), ("Q2", "가온미술관 건립공사 건축 감리용역", "2023-10-20", I, 300_000_000)]) == YY   # 214일 뒤(120일 밖)
    C["F2d"] = lambda: two([("K1", "가온미술관 건립공사 건설사업관리용역", "2023-02-01", I, 300_000_000), ("K2", "가온미술관 건립공사 감리용역", "2023-02-20", I, 300_000_000)]) == YY           # 건설사업관리 ↔ 감리
    C["F2f"] = lambda: two([("A1", "가온미술관 건립공사 전기공사 감리용역", "2023-03-15", I, 90_000_000), ("A2", "가온미술관 외 2개소 건립공사 전기공사 감리용역", "2023-04-01", I, 90_000_000)]) \
        == [("Y", "대표"), ("N", "대표")]                                                                                                                                              # 묶음 공고(N)가 멀쩡한 공고를 대체하지 않는다
    C["F2i"] = lambda: two([("D1", "가온미술관 건립공사 실시설계용역", "2021-03-02", I, 40_000_000), ("D2", "가온미술관 건립공사 조경 실시설계용역", "2021-03-22", I, 40_000_000)]) == YY          # 조경 설계는 한쪽에만 적혀 있어도 다른 용역
    C["F2k"] = lambda: two([("D1", "가온회관 건립공사 실시설계용역", "2022-03-01", I, 50_000_000), ("D2", "가온회관 별관 리모델링공사 실시설계용역", "2022-04-01", I, 50_000_000)],
                           world=([("F1", "가온회관", "F1-N", "C1", "가온회관 건립공사", "2023-03-02", I, "신축", 8_000_000_000),
                                   ("F1", "가온회관", "F1-R", "C2", "가온회관 별관 리모델링공사", "2023-06-01", I, "리모델링", 2_000_000_000)], [("F1", "가온회관", I)])) == YY      # 프로젝트가 다르면 같은 금액이어도 따로
    C["F2l"] = lambda: two([("Q1", "가온미술관 본관 건립사업 감리 용역", "2023-03-20", I, 300_000_000), ("Q2", "가온미술관 건립공사 건축 감리용역", "2023-04-09", I, 300_000_000)]) == NY         # 공종이 한쪽에만 적혀 있으면 같은 용역


    # ── ③ 짧은 일반 이름(_site_fit) ──
    def one(raw, world, f):
        g, out, _ = _mini(*world, [raw], cfg)
        r = g[(raw[0], f)]
        return r["합산여부"], r["프로젝트ID"], r["합산제외_까닭"]


    C["F3d"] = lambda: (lambda y, p, w: y == "N" and "다른 현장" in w)(*one(("G1", "가온역 환승 주차타워 건립공사 건설사업관리용역", "2024-02-10", I, 50_000_000), TOWER, "F4"))      # 시점이 가까워도 다른 현장
    C["F3e"] = lambda: (lambda y, p, w: y == "N" and "다른 현장" in w)(*one(
        ("G1", "나래지구 공공주택 9단지 아파트 건설공사 건설사업관리용역", "2023-03-10", S, 900_000_000),
        ([("F1", "가온주택도시공사 공공주택", "F1-N", "C1", "나래지구 공공주택 조경공사", "2023-03-02", S, "신축", 3_000_000_000)], [("F1", "가온주택도시공사 공공주택", S)]), "F1"))
    C["F3j"] = lambda: (lambda y, p, w: y == "N" and "사업 이름이 없음" in w)(*one(
        ("G1", "가온 다솜지구 주차타워 조성 실시설계용역", "2023-03-01", I, 50_000_000),
        ([("F4", "가온시 주차타워", "F4-N", "C5", "가온 주차타워 건립공사", "2024-03-01", I)], [("F4", "가온시 주차타워", I)]), "F4"))                                     # 지자체 이름('가온')만 같은 것은 현장 말이 아니다
    C["F3k"] = lambda: one(("G1", "나래동 주차타워 건립공사 건설사업관리용역", "2024-02-10", I, 50_000_000), TOWER, "F4")[0] == "Y"                                    # 읍·면·동 이름은 위치 표시
    C["F3q"] = lambda: one(("G1", "가온 주차타워 건립공사 건설사업관리용역", "2024-02-10", I, 50_000_000), TOWER, "F4")[0] == "Y"                                      # 제 기관 이름('가온')은 다른 현장이 아니다
    C["F3o"] = lambda: one(("G1", "용역사업 집행계획 및 사업수행능력평가서 제출 안내 공고[주차타워 건립공사 건설사업관리용역]", "2024-02-10", I, 50_000_000), TOWER, "F4")[0] == "Y"   # 입찰 단계 말은 현장 이름이 아니다
    C["F3p"] = lambda: one(("G1", "주차타워 신축 임목폐기물 처리용역", "2024-05-10", I, 5_000_000), TOWER, "F4")[0] == "Y"                                              # 폐기물 종류 말은 현장 이름이 아니다

    # ── ④ _joint_partner ──
    C["F4b"] = lambda: one(("B6", "가온동 공영주차장 건립 기본 및 실시설계 통합용역", "2021-04-01", I, 300_000_000), LOT, "F1")[0] == "Y"                               # '기본 및 실시설계' 는 현장 이름이 아니다
    C["F4g"] = lambda: one(("B5", "가온동 공영주차장 건립공사 전기·통신 통합감리용역", "2023-04-01", I, 100_000_000), LOT, "F1")[0] == "Y"                             # '전기·통신 통합감리' 는 묶음이 아니다
    C["F4e"] = lambda: one(("B7", "가온동 공영주차장 건립공사 기본 및 실시설계용역(건축·구조)", "2021-06-01", I, 300_000_000), LOT, "F1")[0] == "Y"                     # '통합' 이 없으면 보지 않는다
    C["F4f"] = lambda: one(("B8", "가온동 공영주차장 건립공사 통합건설사업관리용역(건축·구조·토목)", "2023-02-10", I, 300_000_000), LOT, "F1")[0] == "Y"                # '통합' 뒤의 어절은 보지 않는다
    C["F4h"] = lambda: (lambda y, p, w: y == "N" and "묶" in w)(*one(("B9", "마루공원 및 가온동 공영주차장 건립공사 통합 실시설계용역", "2021-06-01", I, 300_000_000), LOT, "F1"))   # 설계도 묶음 발주면 N
    C["F4i"] = lambda: one(("B1", "가온 문화·체육센터 건립 통합건설사업관리용역", "2023-02-10", I, 300_000_000),
                           ([("F1", "가온 문화·체육센터", "F1-N", "C1", "가온복합센터 건립공사", "2023-03-02", I, "신축", 10_000_000_000)], [("F1", "가온 문화·체육센터", I)]), "F1")[0] == "Y"   # 시설 이름 속의 '·'

    # ── ⑤ 공사비와 같은 금액 / 설계 공고 시점 ──
    C["F5c"] = lambda: (lambda y, p, w: y == "N" and "공사 추정가격" in w)(*one(("A3", "가온복합청사 건립사업 설계용역", "2021-05-31", I, 15_200_000_000), HALL, "F1"))          # 0.06% 차이
    C["F5b"] = lambda: one(("A5", "가온복합청사 건립공사 건설사업관리용역", "2023-11-01", I, 15_000_000_000), HALL, "F1")[0] == "Y"                                      # 1.3% 차이(테두리 고정용 — 본문 참고)
    C["F5h"] = lambda: (lambda y, p, w: y == "N" and "공사 추정가격" in w)(*one(("A6", "가온복합청사 건립공사 건설사업관리용역", "2023-11-01", I, 15_191_000_000), HALL, "F1"))   # 건설사업관리에도 적용
    C["F5e"] = lambda: (lambda y, p, w: y == "N" and "설계 공고" in w)(*one(("T5", "가온체육공원 시설물 개보수사업 실시설계용역", "2024-02-01", I, 52_000_000), PARK, "F1"))     # 공사 45일 전
    C["F5f"] = lambda: one(("T6", "가온체육공원 개보수 설계용역", "2023-12-01", I, 30_000_000), PARK, "F1")[0] == "Y"                                                   # 공사 3.5개월 전
    C["F5g"] = lambda: one(("T7", "가온체육공원 롤러경기장 개보수공사 실시설계용역", "2024-03-01", I, 40_000_000), PARK, "F1")[0] == "Y"                                 # 사업 이름이 통째로 있으면 공사 직전이어도

    # ── ⑥ 공사 이름 없는 건설폐기물 공고의 시점 테두리 ──
    C["F6d"] = lambda: (lambda y, p, w: y == "N" and w.startswith("어느 공사"))(*one(("W4", "가온회관 건설폐기물 처리용역", "2024-06-01", I, 30_000_000), CLUB, "F1"))          # 15개월 뒤
    C["F6e"] = lambda: (lambda y, p, w: y == "N" and w.startswith("어느 공사"))(*one(("W5", "가온회관 건설폐기물 처리용역", "2022-12-01", I, 30_000_000), CLUB, "F1"))          # 3개월 전
    C["F6h"] = C["F6e"]

    # ── ⑧ 사업유형 불일치 구제 ──
    C["F8c"] = lambda: (lambda y, p, w: y == "N" and p == "")(*one(
        ("X3", "가온 시민주차타워 건립사업 전기공사 감리용역", "2023-03-10", I, 33_000_000),
        (PT[0] + [("F1", "가온 시민주차타워", "F1-E-02", "C2", "가온 시민주차타워 별관 증축공사", "2023-09-01", I, "증축", 2_000_000_000)], PT[1]), "F1"))               # 후보가 둘이면 정하지 않는다
    C["F8e"] = lambda: (lambda y, p, w: y == "N" and p == "F1-E" and "설계 공고" in w)(*one(("X4", "가온 시민주차타워 건립사업 실시설계용역", "2023-03-10", I, 20_000_000), PT, "F1"))   # 구제된 줄도 사업 이름이 통째로 있는 줄(lvl 3)은 아니다
    C["F8f"] = lambda: (lambda y, p, w: y == "N" and p == "")(*one(("X5", "가온 시민주차타워 부설 관리동 건립 감리용역", "2023-03-10", I, 33_000_000), PT, "F1"))               # 유사도 0.74
    C["F8g"] = lambda: (lambda y, p, w: y == "N" and p == "")(*one(
        ("X6", "가온 시민주차타워 건립 감리용역", "2023-03-10", I, 33_000_000),
        ([("F1", "가온 시민주차타워", "F1-E", "C1", "가온 시민주차타워 증축공사", "2023-03-02", I, "증축", 5_000_000_000)], PT[1]), "F1"))                                  # 유사도는 0.80 이지만 사업유형 말만 다른 이름('건립' ↔ '증축')은 다른 사업

    # ── ⑨ 종류 판정의 나머지 갈래 ──
    is_ = lambda kind, *a: k(*a)[0] == kind and (kind == "기타" or not k(*a)[2])
    C["F9c"] = lambda: is_("폐기물처리", "◇ 철거 폐기물 처리용역", "폐기물재활용서비스") and is_("폐기물처리", "◇ 해체에 따른 폐기물 처리용역", "생활폐기물처리서비스")
    C["F9f"] = lambda: is_("폐기물처리", "◇ 내진보강공사 폐기물 처리용역", "폐기물재활용서비스") and is_("폐기물처리", "◇ 외벽 보수공사 폐기물처리 용역", "생활폐기물처리서비스")
    C["F9k"] = lambda: is_("설계", "◇ 건립공사 지반조사 및 실시설계 용역") and is_("설계", "◇ 건립공사 지질조사 및 기본설계 용역")
    C["F9m"] = lambda: is_("설계", "◇ 조성사업 기본구상 및 기본설계 용역")
    C["F9n"] = lambda: is_("설계", "◇ 건립 기본계획 수립 및 실시설계 용역")
    C["F9o"] = lambda: is_("설계", "◇ 조성사업 측량 및 기본 및 실시설계 용역")
    C["F9p"] = lambda: is_("설계", "◇ 건립공사 기본계획 및 기본설계 용역")
    C["F9v"] = lambda: is_("설계", "◇ 다목적체육관 건축기획 및 기본설계 용역") and is_("설계", "◇ 다목적체육관 건축기획 및 실시설계 용역")
    C["F9w"] = lambda: is_("설계", "◇ 시험인증 지원동 신축공사 설계용역") and is_("건설사업관리·감리", "가온 시험인증 클러스터 건립공사 감리용역")
    C["F9x"] = lambda: is_("설계", "가온 탄소인증센터 신축공사 설계용역") and is_("건설사업관리·감리", "가온 인증지원센터 건립공사 감리용역") and is_("설계", "가온 종자인증동 증축공사 설계용역")
    C["F9z"] = lambda: (is_("설계", "가온 농산물 홍보판매장 건립공사 설계용역") and is_("설계", "가온 안전 홍보체험관 건립공사 설계용역")
                        and is_("건설사업관리·감리", "가온 홍보전시관 리모델링공사 감리용역") and is_("설계", "가온홍보원 청사 신축공사 설계용역"))
    C["F9B"] = lambda: is_("기타", "◇ 건립 홍보관리 및 설계 용역") and is_("기타", "◇ 건립사업 홍보관련 영상 설계 용역")

    # ── ⑩ 기관: '○군청' ──
    C["F10b"] = lambda: (servc.inst_same("전라남도 가온군청", "전라남도 가온군") and servc.inst_same("가온군청", "전라남도 가온군")
                         and not servc.inst_same("전라남도 문화예술회관", "전라남도 가온군청"))
    bad = [lab for lab, fn in C.items() if fn() is not True]
    assert not bad, bad


def test_field_audit_review():
    """독립 검토(2026-11-25)가 실행으로 재현한 반례 — test_field_audit 의 규칙을 처음 넣었을 때의 결함. 모든 이름은 가상."""
    cfg = load_config()
    k, I, S, G = servc.classify_kind, "경기도 가온시", "가온주택도시공사", "서울특별시 가온구"
    dump = lambda out: out[["공고번호", "시설ID", "프로젝트ID", "합산여부", "합산제외_까닭", "상태"]].to_dict("records")
    yn = lambda g, nos, f="F1": [g[(n, f)]["합산여부"] for n in nos]
    # ── ① '추정'(대행 기관·느슨한 이름) + 사업 이름 통째: 접두를 뗀 일반 이름 + '건립' 뿐인 사업 이름('국민체육센터 건립')은 광역 건설본부가 낸 다른 자치구의 같은 이름 사업에도 들어 있다 ──
    g, out, _ = _mini([("F1", "가온구 국민체육센터", "F1-N", "C1", "가온구 국민체육센터 건립공사", "2023-03-02", G, "신축", 10_000_000_000),
                       ("F1", "가온구 국민체육센터", "F1-N", "C2", "국민체육센터 건립 전기공사", "2023-03-10", G, "신축", 1_000_000_000)], [("F1", "가온구 국민체육센터", G)],
                      [("E1", "나래구 국민체육센터 건립공사 건설사업관리용역", "2023-02-01", "서울특별시 도시기반시설본부", 900_000_000),
                       ("E2", "서남권 국민체육센터 건립 기본 및 실시설계용역", "2021-02-01", "서울특별시 도시기반시설본부", 700_000_000),
                       ("E3", "가온구 국민체육센터 건립공사 건설사업관리용역", "2023-02-05", "서울특별시 도시기반시설본부", 950_000_000)], cfg)   # 검수한 이름이 그대로 — 더한다
    assert yn(g, ("E1", "E2", "E3")) == ["N", "N", "Y"] and "추정" in g[("E1", "F1")]["합산제외_까닭"], dump(out)
    g, out, _ = _mini([("F1", "가온 센터 공영주차장", "F1-N", "C1", "가온혁신센터 공영주차장 건립사업", "2023-03-02", I, "신축", 3_000_000_000),
                       ("F1", "가온 센터 공영주차장", "F1-N", "C2", "공영주차장 조성 전기공사", "2023-03-12", I, "신축", 300_000_000)], [("F1", "가온 센터 공영주차장", I)],
                      [("E2", "가온나눔센터 뒤편 공영주차장 조성 실시설계용역", "2022-04-01", I, 20_000_000)], cfg)             # 짧은 공종 공고명('공영주차장 조성')이 통째로 들어 있을 뿐
    assert yn(g, ("E2",)) == ["N"], dump(out)
    # ── ② 짧은 일반 이름: '○○시장'·'○○지구' 는 위치 표시가 아니라 다른 현장 이름 / 꾸밈말('공영')뿐인 사업 이름은 현장을 가리지 못한다 / 사업 이름 뒤의 다른 대상 ──
    T = ([("F1", "가온시 주차타워", "F1-N", "C1", "가온시 주차타워 건립공사", "2023-03-02", I, "신축", 8_000_000_000),
          ("F1", "가온시 주차타워", "F1-N", "C2", "주차타워 건립 전기공사", "2023-03-12", I, "신축", 500_000_000)], [("F1", "가온시 주차타워", I)])
    g, out, _ = _mini(*T, [("A2", "중앙시장 주차타워 건립공사 건설사업관리용역", "2023-02-02", I, 410_000_000), ("A3", "나래지구 주차타워 건립공사 감리용역", "2023-02-03", I, 420_000_000),
                           ("A4", "원도심 주차타워 건립 기본 및 실시설계용역", "2021-06-03", I, 430_000_000), ("A6", "나래동 주차타워 건립공사 전기공사 감리용역", "2023-03-03", I, 45_000_000),
                           ("A9", "문화의거리 주차타워 건립공사 통신공사 감리용역", "2023-03-05", I, 47_000_000)], cfg)
    assert yn(g, ("A2", "A3", "A4", "A6", "A9")) == ["N", "N", "N", "Y", "N"] and "다른 현장" in g[("A2", "F1")]["합산제외_까닭"], dump(out)
    g, out, _ = _mini([("F1", "가온시 주차타워", "F1-N", "C1", "공영 주차타워 건립공사", "2023-03-02", I, "신축", 8_000_000_000)], [("F1", "가온시 주차타워", I)],
                      [("B1", "중앙시장 공영 주차타워 건립공사 건설사업관리용역", "2025-03-02", I, 400_000_000), ("B3", "공영 주차타워 건립공사 건설사업관리용역", "2023-02-02", I, 380_000_000),
                       ("B4", "공영 주차타워 건립공사 건설폐기물 처리용역", "2026-03-02", I, 30_000_000)], cfg)                   # 사업 이름이 '공영 주차타워 건립' 뿐 — 시점까지 맞아야(36개월 뒤는 N)
    assert yn(g, ("B1", "B3", "B4")) == ["N", "Y", "N"], dump(out)
    g, out, _ = _mini([("F1", "가온주택도시공사 공공주택", "F1-N", "C1", "나래지구 공공주택 전기공사", "2023-03-02", S, "신축", 3_000_000_000)], [("F1", "가온주택도시공사 공공주택", S)],
                      [("S1", "나래지구 공공주택(A1) 공동주택 전기 설계용역", "2022-03-02", S, 300_000_000),                     # 그 지구의 다른 대상(A1 공동주택)
                       ("S2", "나래지구 공공주택 건설폐기물(소각) 처리용역", "2023-06-02", S, 40_000_000)], cfg)
    assert yn(g, ("S1", "S2")) == ["N", "Y"] and "다른 현장" in g[("S1", "F1")]["합산제외_까닭"], dump(out)
    # ── ③ 통합 용역의 묶음 판정: 사업 이름 자리(시설 이름 앞, 또는 시설 이름과 짝을 이뤄 첫 공사 말 앞)에 놓인 어절만 상대 현장으로 본다 ──
    M = ([("F1", "가온미술관", "F1-N", "C1", "가온미술관 건립공사", "2023-03-02", I, "신축", 20_000_000_000)], [("F1", "가온미술관", I)])
    g, out, _ = _mini(*M, [("J1", "가온미술관 건립공사 전력시설물·정보통신공사 통합감리용역", "2023-03-10", I, 90_000_000), ("J4", "가온미술관 건립공사 건축·태양광 통합감리용역", "2023-03-12", I, 70_000_000),
                           ("J8", "가온미술관 전시동·수장고동 건립공사 통합건설사업관리용역", "2023-02-01", I, 800_000_000),
                           ("J9", "가온미술관 건립공사 통합건설사업관리용역(PQ, SOQ 제출 안내)", "2023-01-10", I, 810_000_000),
                           ("K1", "가온미술관 및 나래문학관 건립공사 통합건설사업관리용역", "2023-02-02", I, 1_500_000_000),
                           ("K2", "가온미술관 및 주차복합빌딩 건립사업 통합 건설사업관리용역", "2023-02-03", I, 1_600_000_000)], cfg)
    assert yn(g, ("J1", "J4", "J8", "J9", "K1", "K2")) == ["Y", "Y", "Y", "Y", "N", "N"], dump(out)
    g, out, _ = _mini([("F1", "가온 통합수장고", "F1-N", "C1", "가온 통합수장고 건립공사", "2023-03-02", I, "신축", 9_000_000_000)], [("F1", "가온 통합수장고", I)],
                      [("Q2", "가온 통합수장고 건립공사 전기·통신 감리용역", "2023-03-10", I, 80_000_000)], cfg)                     # 시설 이름 속의 '통합'
    assert yn(g, ("Q2",)) == ["Y"], dump(out)
    # ── ④ '같은 금액' 규칙의 공종 비교는 시설 이름을 가린 공고명으로('○○소방서' 의 '소방', '주차전용건축물' 의 '건축') ──
    for fac in ("가온소방서", "가온 정보통신센터", "가온동 주차전용건축물"):
        g, out, _ = _mini([("F1", fac, "F1-N", "C1", fac + " 신축공사", "2023-03-02", I, "신축", 8_000_000_000)], [("F1", fac, I)],
                          [("D1", fac + " 신축공사 전기공사 감리용역", "2023-03-10", I, 22_000_000), ("D2", fac + " 신축공사 소방공사 감리용역", "2023-03-20", I, 22_000_000),
                           ("D3", fac + " 신축공사 정보통신공사 감리용역", "2023-03-25", I, 22_000_000)], cfg)
        assert yn(g, ("D1", "D2", "D3")) == ["Y", "Y", "Y"], (fac, dump(out))
    # ── ⑤ 공사 이름 없는 '건설폐기물' 공고: 폐기물 말 앞에 다른 시설·다른 공사 말이 있으면 그 시설의 하나뿐인 프로젝트에 더하지 않는다 ──
    g, out, _ = _mini([("F1", "가온회관", "F1-N", "C1", "가온회관 건립공사", "2023-03-02", I, "신축", 8_000_000_000)], [("F1", "가온회관", I)],
                      [("W1", "가온회관 및 나래복지관 건설폐기물 처리용역", "2023-06-01", I, 30_000_000), ("W6", "가온회관 사무실 칸막이 공사 건설폐기물 처리용역", "2023-06-06", I, 35_000_000),
                       ("W7", "가온회관 등 건설폐기물 처리용역", "2023-06-07", I, 36_000_000), ("W8", "가온회관, 나래도서관 건설폐기물 처리용역", "2023-06-08", I, 37_000_000),
                       ("WB", "가온회관 건설폐기물(폐콘크리트) 처리용역", "2023-06-11", I, 40_000_000)], cfg)
    assert yn(g, ("W1", "W6", "W7", "W8", "WB")) == ["N", "N", "N", "N", "Y"], dump(out)
    # ── ⑥ 종류 판정: 공사 말과 폐기물 말이 떨어져 있으면 그 공사의 폐기물이 아니다 / 기관 이름 속의 '공사' / 설계 뒤의 지반조사 / 건축 설계 분류가 아닌 제안공모 ──
    for t, c in (("◇ 신축 이전에 따른 구청사 사무집기 폐기물 처리용역", "생활폐기물처리서비스"), ("◇ 증축공사 준공 청소 및 폐기물 처리용역", "건물청소서비스"),
                 ("가온도시공사 ◇ 사업장폐기물 처리용역", "생활폐기물처리서비스"), ("가온주택도시공사 ◇ 입주민 폐기물 수거 용역", "생활폐기물처리서비스"), ("폐목재 재활용시설 운영 용역", "폐기물재활용서비스"),
                 ("◇ 리모델링 실감콘텐츠 제안공모", "소프트웨어설계서비스"), ("◇ 운영 수탁자 제안공모", "건축설계서비스"), ("◇ 건립공사 녹색건축 인증 시설 설계용역", "")):
        assert k(t, c)[0] == "기타", (t, c, k(t, c))
    for t, c, kind in (("◇ 시설물 해체 폐기물 처리용역", "생활폐기물처리서비스", "폐기물처리"), ("◇조성사업혼합(가연성)폐기물처리용역(긴급)", "폐기물재활용서비스", "폐기물처리"),
                       ("◇ 공공주택지구(△△ 시행구역) 폐목재재활용 처리용역", "폐기물재활용서비스", "폐기물처리"), ("◇ 건립공사 실시설계 및 지반조사 용역", "", "설계")):
        assert k(t, c)[0] == kind and not k(t, c)[2], (t, c, k(t, c))
    # ── ⑦ 설계 공고 시점 규칙은 철거·해체 공고를 기준에서 뺀다: 철거 공고만 먼저 난 사업의 설계를 잘못 빼지 않게 ──
    g, out, _ = _mini([("F1", "가온문화센터", "F1-N", "C1", "가온문화센터 건립부지 기존 건축물 철거공사", "2023-03-02", I, "신축", 300_000_000)], [("F1", "가온문화센터", I)],
                      [("T1", "가온문화센터 건립 기본 및 실시설계용역", "2023-02-10", I, 500_000_000)], cfg)
    assert yn(g, ("T1",)) == ["Y"], dump(out)
    g, out, _ = _mini([("F1", "가온문화센터", "F1-N", "C1", "가온문화센터 건립부지 기존 건축물 철거공사", "2023-03-02", I, "신축", 300_000_000),
                       ("F1", "가온문화센터", "F1-N", "C2", "가온문화센터 신축공사", "2023-04-02", I, "신축", 9_000_000_000)], [("F1", "가온문화센터", I)],
                      [("T1", "가온문화센터 건립 기본 및 실시설계용역", "2023-03-20", I, 500_000_000)], cfg)                       # 본 공사 공고 13일 전
    assert yn(g, ("T1",)) == ["N"] and "설계 공고" in g[("T1", "F1")]["합산제외_까닭"], dump(out)


def run():
    test_kinds()
    test_institutions()
    test_name_and_project_rules()
    test_more_rules()
    test_review_fixes()
    test_review_round2()
    test_merge_rules()
    test_review_round2_fixes()
    test_field_audit()
    test_field_audit_edges()
    test_field_audit_review()
    with tempfile.TemporaryDirectory() as tmp:
        test_excel_and_stage(tmp)          # test_link 를 안에서 함께 실행
        test_stage_more(tmp)
        test_review_stage(tmp)
    print("OK: 용역 연결 테스트 통과")


if __name__ == "__main__":
    run()
