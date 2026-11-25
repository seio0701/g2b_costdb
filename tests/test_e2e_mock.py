"""가짜 나라장터 서버로 전 단계 통합 실행: probe → collect → discover → research → dedup → attach → extract(견적/--yes 모의) → excel.
실행: python -m tests.test_e2e_mock   (네트워크·API 키 불필요, 임시 폴더에서 동작)
※ 모든 데이터는 가상의 예시. LLM 호출은 텍스트에서 정규식으로 값을 뽑는 가짜 함수로 대체한다.
"""
from __future__ import annotations

import contextlib
import io
import json
import os
import re
import sys
import tempfile
from unittest import mock

import pandas as pd
import yaml

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
from g2b_costdb import extract_llm, pipeline  # noqa: E402
from g2b_costdb.classify import ROOT, load_config  # noqa: E402
from tests.mock_g2b import KEY, MockG2B  # noqa: E402


def _run(argv):
    """pipeline.main 을 실행하고 표준출력을 돌려준다."""
    buf = io.StringIO()
    with contextlib.redirect_stdout(buf):
        pipeline.main(argv)
    out = buf.getvalue()
    print(out[-1500:] if len(out) > 1500 else out)
    return out


def _fake_extra(text: str, api_hint: dict, cfg_llm: dict, max_tokens=None) -> dict:
    """보강 추출 가짜 모델: 건축 공고에는 공사비 정보광장 항목, 전기 공고에는 기타금액(한전불입금)."""
    bid = str(api_hint.get("공고번호"))
    if bid == "R24010001":
        raw = {"대지면적_m2": "30000", "건축면적_m2": "4500", "건폐율_pct": "15", "용적률_pct": "45.5", "최고높이_m": "24.5", "주차대수": "210",
               "건축법상용도": "문화 및 집회시설(공연장)", "설계특이사항": "본동 1개동, 무대기계 포함", "신뢰도": "high",
               "근거문구": [{"항목": "대지면적", "원문": "대지면적 30,000㎡"}]}
    elif bid == "R24010002":
        raw = {"기타금액_원": "38000000", "기타금액_항목": "한전불입금 38,000,000", "기타금액_위치": "별도", "신뢰도": "high", "근거문구": []}
    else:
        raw = {"신뢰도": "low", "근거문구": []}
    d = extract_llm.normalize_extra_doc(raw)
    d["_model"], d["_usage_in"], d["_usage_out"] = "fake", 100, 30
    return d


def _fake_trades(snippet: str, api_hint: dict, cfg_llm: dict, max_tokens=None) -> dict:
    """2차 추출 가짜 모델: 집계표 발췌에서 '공종행<TAB>…<TAB>금액' 을 읽어 공종별 금액 JSON 을 만든다(모델이 할 일을 규칙으로 흉내)."""
    assert "총괄집계표" in snippet and "기계설비공사" in snippet and "세부항목" not in snippet, snippet[:300]
    label = {"건축공사": "건축", "기계설비공사": "기계설비", "토목공사": "토목", "조경공사": "조경", "간접노무비": "간접비", "일반관리비": "간접비",
             "이윤": "간접비", "부가가치세": "부가세"}
    items, total = [], None
    for line in snippet.splitlines():
        cells = line.split("\t")
        if cells[0] in label:
            items.append({"공종": label[cells[0]], "항목명_원문": cells[0], "금액_원": cells[-1]})
        if cells[0] == "총공사비" and len(cells) > 1 and cells[-1].replace(",", "").isdigit():   # 현장설명서의 '총공사비' 머리글 줄은 제외
            total = cells[-1]
    return extract_llm.normalize_trade_doc({"금액기준": "부가세포함", "관급자재_포함여부": "제외", "단위": "원", "합계_원": total or "",
                                           "공종별금액": items, "출처": "총괄집계표", "신뢰도": "high", "근거문구": "총괄집계표 공종 행"})


def _fake_extract(text: str, api_hint: dict, cfg_llm: dict, max_tokens=None) -> dict:
    """LLM 대신 정규식으로 공고문 텍스트에서 값을 뽑는 모의 추출기."""
    def amt(label):
        m = re.search(label + r"[^\d]{0,20}([\d,]{5,})", text)
        return int(m.group(1).replace(",", "")) if m else None
    area = re.search(r"연면적\s*([\d,\.]+)\s*㎡", text)
    floors = re.search(r"지상\s*(\d+)\s*층", text)
    days = re.search(r"(\d[\d,]*)\s*일", text)
    doc = {"공사명": api_hint.get("공고명"), "사업유형": "신축", "공종": "건축",
           "연면적_m2": float(area.group(1).replace(",", "")) if area else None,
           "지상층수": int(floors.group(1)) if floors else None,
           "공사기간_일": int(days.group(1).replace(",", "")) if days else None,
           "추정가격_원": amt("추정가격"), "기초금액_원": amt("기초금액"), "관급자관급액_원": amt("관급자관급액"),
           "총공사금액_원": amt("총공사금액"), "신뢰도": "high",
           "근거문구": [{"항목": "연면적_m2", "원문": area.group(0) if area else ""}]}
    out = extract_llm.normalize_doc(doc)
    out.update({"_model": "mock", "_usage_in": 1000, "_usage_out": 300})
    return out


def run():
    server = MockG2B()
    base = server.start()
    os.environ["NO_PROXY"] = "127.0.0.1,localhost"
    tmp = tempfile.mkdtemp(prefix="g2b_e2e_")
    cfg = yaml.safe_load(open(os.path.join(ROOT, "config.yaml"), encoding="utf-8"))
    cfg["api"]["base_url"] = base + "/BidPublicInfoService"
    cfg["api"]["sleep_between_calls_sec"] = 0
    cfg["api"]["max_retries"] = 2
    cfg["api"]["use_license_limit"] = True                          # 가짜 서버로 면허제한 경로도 검증
    cfg["api"]["use_awards"] = True                                 # 낙찰정보서비스(별도 End Point) 경로도 검증
    cfg["api"]["award_base_url"] = base + "/ScsbidInfoService"
    cfg["period"] = {"start": "2024-01", "end": "2024-04"}          # 2024-04 는 데이터 없음(03) 경로
    cfg["servc"]["period_end"] = "2024-04"                           # 용역 수집 기간(비우면 오늘이 속한 달까지 — 테스트에서는 고정)
    cfg["paths"] = {"data_dir": "data", "raw_dir": "data/raw", "cache_db": "data/cache.sqlite", "text_dir": "data/text",
                    "files_dir": "data/files", "out_dir": "output", "excel_name": "공사비DB.xlsx"}
    cfg_path = os.path.join(tmp, "config.yaml")
    with open(cfg_path, "w", encoding="utf-8") as f:
        yaml.safe_dump(cfg, f, allow_unicode=True)
    data = lambda *p: os.path.join(tmp, "data", *p)
    try:
        # 0) 잘못된 키 → probe 가 30 오류를 한국어로 안내하고 죽지 않음, 재시도 없음
        os.environ["G2B_SERVICE_KEY"] = "wrong-key"
        out = _run(["probe", "--ym", "2024-01", "--config", cfg_path])
        assert "API 오류 30" in out and "G2B_SERVICE_KEY" in out, out
        assert server.calls["getBidPblancListInfoCnstwk"] == 1, "키 오류는 1회만 호출"
        # 1) probe — 매핑 OK + 면허제한 필드 후보 안내
        os.environ["G2B_SERVICE_KEY"] = KEY
        out = _run(["probe", "--ym", "2024-01", "--config", cfg_path])
        assert out.count("[매핑 OK]") == 2 and "후보 lcnsLmtNm 발견" in out, out
        assert "[award_amt] OK" in out and "[award_rate] OK" in out and "[낙찰정보] 기간 조회 파라미터 {'inqryDiv': '1'}" in out, out
        # 1-2) 용역 공고(같은 서비스의 용역 오퍼레이션): probe --servc → collect --servc. 폐기물·건설사업관리만 보존, 무관한 청소 용역은 버림
        out = _run(["probe", "--servc", "--ym", "2024-03", "--config", cfg_path])
        # 표본은 그 달의 첫 구간(7일)으로 받는다 — 한 달 전체 조회는 서버가 07(입력범위값 초과)로 거절하며, probe 는 그 답을 그대로 보여 준다
        assert "getBidPblancListInfoServc (20240301~20240307, 7일) totalCount=2" in out and "[매핑 OK]" in out and "표본 2건 중 1건이 보존 대상" in out, out
        assert "[조회 기간] 한 달 전체(20240301~20240331) 조회" in out and "코드 07" in out and "입력범위값 초과 에러" in out and "5구간으로 나눠" in out, out
        assert "공고명 '폐기물' (20240301~20240307) totalCount=0" in out and "[호출 수 추정]" in out and "한 달(5구간)" in out, out
        calls0 = server.calls["getBidPblancListInfoServc"]
        out = _run(["collect", "--servc", "--config", cfg_path])
        sv = pd.read_parquet(data("servc_all.parquet"))
        assert sorted(sv["bidNtceNo"]) == ["R24S00001", "R24S00002", "R24S00004"], sv["bidNtceNm"].tolist()
        assert set(open(data("raw", "servc_done.txt")).read().split()) == {"202401", "202402", "202403", "202404"}
        assert server.calls["getBidPblancListInfoServc"] - calls0 == 20, "한 달을 7일 이하 5구간으로 나눠 조회 × 4개월"
        assert not os.path.exists(data("notices_all.parquet")), "collect --servc 는 공사 수집을 하지 않는다"
        import sqlite3
        con = sqlite3.connect(data("cache.sqlite"))
        assert con.execute("SELECT COUNT(*) FROM cache WHERE op='getBidPblancListInfoServc' AND params LIKE '%inqryBgnDt%'").fetchone()[0] == 0, "완료한 달의 용역 응답 캐시는 비운다"
        con.close()
        calls1 = server.calls["getBidPblancListInfoServc"]
        _run(["collect", "--servc", "--config", cfg_path])
        assert server.calls["getBidPblancListInfoServc"] == calls1, "완료한 달은 다시 호출하지 않는다"
        # 2) collect — 4개월(1개월은 NODATA), done.txt, parquet
        out = _run(["collect", "--config", cfg_path])
        notices = pd.read_parquet(data("notices_all.parquet"))
        assert len(notices) == len(server.data["notices"]), len(notices)
        assert set(open(data("raw", "notices_done.txt")).read().split()) == {"202401", "202402", "202403", "202404"}
        assert os.path.exists(data("bsis_all.parquet"))
        awards = pd.read_parquet(data("award_all.parquet"))
        assert len(awards) == len(server.data["awards"]) and awards["_award_amt"].notna().all(), awards
        assert set(open(data("raw", "award_done.txt")).read().split()) == {"202401", "202402", "202403", "202404"}
        calls_after_collect = (server.calls["getBidPblancListInfoCnstwk"], server.calls["getScsbidListSttusCnstwk"])
        # 2-2) check-raw — 받은 원문 점검(API 호출 없음): 서버가 알려 준 건수와 저장된 줄 수가 맞는다
        out = _run(["check-raw", "--config", cfg_path])
        assert "다시 받아야 할 달이 없습니다" in out and "notices_*.jsonl" in out and "다시 받아야 할 달 0개" in out, out
        assert (server.calls["getBidPblancListInfoCnstwk"], server.calls["getScsbidListSttusCnstwk"]) == calls_after_collect, "check-raw 는 API 를 부르지 않는다"
        _run(["collect", "--config", cfg_path])                        # 재실행: 완료된 달은 호출 없음(낙찰 목록 포함)
        assert (server.calls["getBidPblancListInfoCnstwk"], server.calls["getScsbidListSttusCnstwk"]) == calls_after_collect
        # 3) discover — 시설 3개(가상군/다른군 문화예술회관, 가상시 야외공연장), 유지보수는 '제외' 기본값, 잡음 없음
        out = _run(["discover", "--config", cfg_path])
        assert "낙찰정보 연결: 공고 4건" in out, out            # R24010001(000·001), R24010002, R24020006
        from g2b_costdb.discover import read_review_file
        rev = read_review_file(os.path.join(tmp, "output", "facility_candidates.xlsx"))
        assert len(rev) == 3, rev[["시설ID", "시설명_후보", "수요기관", "사업유형", "검수_포함여부"]]
        assert set(rev["수요기관"]) == {"가상군", "다른군", "가상시"} and (rev["검수_포함여부"] == "포함").all()
        assert list(rev[rev["수요기관"] == "가상시"]["검색어"]) == ["야외공연장"], "두 분류에 걸려도 구체적 검색어 1건"
        ids_before = dict(zip(rev["수요기관"], rev["시설ID"]))
        _run(["discover", "--config", cfg_path])                       # 재실행: 시설ID 유지
        rev2 = read_review_file(os.path.join(tmp, "output", "facility_candidates.xlsx"))
        assert dict(zip(rev2["수요기관"], rev2["시설ID"])) == ids_before
        out = _run(["discover", "--fresh", "--config", cfg_path])       # 검수 칸 초기화, 시설ID 는 유지
        rev3 = pd.read_excel(os.path.join(tmp, "output", "facility_candidates.xlsx"), sheet_name="시설후보_검수")
        assert "[--fresh]" in out and dict(zip(rev3["수요기관"], rev3["시설ID"])) == ids_before and (rev3["검수_포함여부"] == "포함").all()
        # 4) research — 면허제한 필드 자동 탐색(lcnsLmtNm), 방수공사(유지보수) 제외, 다른군은 자기 공고만
        out = _run(["research", "--config", cfg_path])
        assert "'lcnsLmtNm'" in out, out
        rs = pd.read_parquet(data("notices_research.parquet"))
        assert "R24010008" not in set(rs["공고번호"]) and set(rs[rs["시설명"] == "다른군 문화예술회관"]["공고번호"]) == {"R24020007"}
        assert rs[rs["공고번호"] == "R24010001"]["공종"].iloc[0] == "건축", "토목건축공사업 → 건축"
        assert rs[rs["공고번호"] == "R24020004"]["공고종류"].iloc[0] == "재공고", "'(취소 후 재공고)' 는 재공고"
        # 5) dedup — 변경차수 채택, 소방은 재공고가 대표, 본공사 두 달 뒤의 관리사무소 건립공사(4억)는 같은 세부사업의 건축 공고로 합산,
        #    기초금액 없는 다른군은 대표 유지
        out = _run(["dedup", "--config", cfg_path])
        latest = pd.read_parquet(data("notices_latest.parquet"))
        gid = ids_before["가상군"]
        by_no = latest.set_index("공고번호")
        assert by_no.at["R24010001", "프로젝트ID"] == f"{gid}-N" and by_no.at["R24010001", "공고차수"] == "001"
        assert by_no.at["R24020004", "프로젝트키"] == f"{gid}-N|소방|" and "R24010003" not in by_no.index
        assert by_no.at["R24030005", "프로젝트ID"] == f"{gid}-N" and by_no.at["R24030005", "대표선정사유"].startswith("합산(같은 공종 2건)"), by_no["대표선정사유"].to_dict()
        assert "같은 공종 합산" in out and "소액 공고" not in out
        assert len(latest) == 6 and latest["프로젝트ID"].nunique() == 3, latest[["프로젝트키", "공고번호"]]
        # 4-2) link-servc — 용역 공고를 시설·프로젝트에 연결(API 호출 없음). 건설폐기물 처리용역·건설사업관리용역은 문화예술회관 신축 프로젝트로,
        #      생활폐기물(조달분류) 공고는 이름이 맞지 않아 연결되지 않는다
        calls_sv = sum(server.calls.values())
        out = _run(["link-servc", "--config", cfg_path])
        assert sum(server.calls.values()) == calls_sv, "link-servc 는 API 를 호출하지 않는다"
        sl = pd.read_parquet(data("servc_linked.parquet")).set_index("공고번호")
        assert set(sl.index) == {"R24S00001", "R24S00002"} and set(sl["프로젝트ID"]) == {f"{gid}-N"} and set(sl["합산여부"]) == {"Y"}, sl.to_dict("index")
        assert sl.at["R24S00001", "용역종류"] == "폐기물처리" and sl.at["R24S00002", "용역종류"] == "건설사업관리·감리" and sl.at["R24S00002", "추정가격"] == 1500000000
        assert "[진단] 용역 공고 3줄" in out and "합산여부 Y): 2줄" in out and os.path.exists(os.path.join(tmp, "output", "servc_link_summary.txt")), out
        # 6) attach — HWPX/DOCX/cp949 TXT 추출, HTML 응답은 실패 기록, 재실행 시 파일 재요청 없음
        out = _run(["attach", "--config", cfg_path])
        texts = json.load(open(data("texts.json"), encoding="utf-8"))
        notes = pd.read_parquet(data("notes_attach.parquet"))
        assert "연면적 15,200㎡" in texts["R24010001-001"] and "관급자관급액" in texts["R24010001-001"] and "건축면적 6,300㎡" in texts["R24010001-001"], \
            "대표(변경 차수 001)에는 변경공고문만 있어도 000 차수의 입찰공고문·현장설명서를 합쳐 처리"
        assert "정정합니다" in texts["R24010001-001"] and "다른 차수의 첨부를 합쳐 처리한 공고 1건" in out, out
        assert "추정가격 2,600,000,000원" in texts["R24010002-000"], "cp949 텍스트 복원"
        html_note = notes[notes["URL"].str.endswith("login.html")].iloc[0]
        assert html_note["다운로드"] == "N", html_note.to_dict()
        assert "객석 1,200석" in texts["R24020006-000"]
        n_file_calls = sum(v for k, v in server.calls.items() if k.startswith("file:"))
        _run(["attach", "--config", cfg_path])
        assert sum(v for k, v in server.calls.items() if k.startswith("file:")) == n_file_calls, "이미 처리한 공고는 재요청 없음"
        n_notes = len(pd.read_parquet(data("notes_attach.parquet")))
        out = _run(["attach", "--retry-failed", "--config", cfg_path])          # 실패 파일(login.html)이 있는 공고만 다시 처리, 기록은 중복되지 않음
        notes2 = pd.read_parquet(data("notes_attach.parquet"))
        assert "[--retry-failed]" in out and len(notes2) == n_notes and not notes2.duplicated(["공고번호", "URL"]).any(), (n_notes, len(notes2))
        assert sum(v for k, v in server.calls.items() if k.startswith("file:")) == n_file_calls + 1, "실패했던 파일 1개만 다시 요청"
        assert json.load(open(data("texts.json"), encoding="utf-8")).keys() == texts.keys()
        out = _run(["attach", "--report", "--config", cfg_path])
        assert "텍스트 없는 공고 3건의 원인" in out and "다운로드 실패" in out and os.path.exists(data("notes_notext.parquet")), out
        # 7) extract — 견적만 / --yes 는 가짜 추출기로
        out = _run(["extract", "--config", cfg_path])
        assert "extract --yes" in out and not os.path.exists(data("llm_docs.json"))
        os.environ["ANTHROPIC_API_KEY"] = "dummy-for-test"
        with mock.patch.object(extract_llm, "extract_with_claude", side_effect=_fake_extract):
            out = _run(["extract", "--yes", "--config", cfg_path])
        docs = json.load(open(data("llm_docs.json"), encoding="utf-8"))
        assert docs["R24010001-001"]["연면적_m2"] == 15200.0 and docs["R24010001-001"]["관급자관급액_원"] == 2100000000
        # 파일 인수인계: 1건을 실패 상태로 만들고 export → (Cowork 역할) JSON 작성 → import
        docs["R24010002-000"] = {"_error": "모의 실패"}
        with open(data("llm_docs.json"), "w", encoding="utf-8") as f:
            json.dump(docs, f, ensure_ascii=False)
        out = _run(["extract", "--export", "--limit", "5", "--config", cfg_path])
        assert os.path.exists(data("llm_in", "R24010002-000.txt")) and "내보내기 완료: 1건" in out and "(--limit 5)" in out, out
        exported = open(data("llm_in", "R24010002-000.txt"), encoding="utf-8").read()
        fake = _fake_extract(exported.split("[공고문 텍스트]")[1], {"공고명": "가상군 문화예술회관 건립 전기공사"}, cfg["llm"])
        with open(data("llm_out", "R24010002-000.json"), "w", encoding="utf-8") as f:
            json.dump({k: v for k, v in fake.items() if not k.startswith("_")}, f, ensure_ascii=False)
        out = _run(["extract", "--import", "--config", cfg_path])
        assert "JSON 1건 반영" in out and "미추출 0건" in out, out
        assert os.path.exists(data("llm_done", "R24010002-000.json")) and os.path.exists(data("llm_done", "R24010002-000.txt")) \
            and not os.path.exists(data("llm_out", "R24010002-000.json")), "반영한 파일은 llm_done 으로 이동"
        # API(모의) 결과가 있는 R24010001-001 에 사람이 쓴 JSON 이 들어와도 덮어쓰지 않음
        os.makedirs(data("llm_out"), exist_ok=True)
        json.dump({"연면적_m2": 1, "신뢰도": "low"}, open(data("llm_out", "R24010001-001.json"), "w", encoding="utf-8"))
        out = _run(["extract", "--import", "--config", cfg_path])
        docs = json.load(open(data("llm_docs.json"), encoding="utf-8"))
        assert "이미 API 결과가 있어 유지한 공고 1건" in out and docs["R24010001-001"]["연면적_m2"] == 15200.0, out
        docs = json.load(open(data("llm_docs.json"), encoding="utf-8"))
        assert docs["R24010002-000"]["추정가격_원"] == 2600000000 and docs["R24010002-000"]["_model"] == "manual/cowork"
        # 배치: 제출 전 견적(50%)만 출력하고 제출하지 않음
        # 2차 추출(공종별 금액): 견적 → --trades 로 1차와 함께 즉시 실행(가짜 모델) → 04b 시트
        out = _run(["extract-trades", "--config", cfg_path])
        assert "2차 추출 대상" in out and "extract-trades --yes" in out and not os.path.exists(data("llm_trades.json")) or os.path.exists(data("llm_trades.json")), out
        with mock.patch.object(extract_llm, "extract_with_claude", side_effect=_fake_extract), \
                mock.patch.object(extract_llm, "extract_trades_with_claude", side_effect=_fake_trades):
            out = _run(["extract", "--yes", "--trades", "--config", cfg_path])
        assert "── 2차 추출" in out, out
        tdocs = json.load(open(data("llm_trades.json"), encoding="utf-8"))
        t = tdocs["R24010001-001"]
        assert t["_공종합_원"] == 26_000_000_000 and t["합계_원"] == 33_000_000_000 and t["_files"] == ["공내역서_가상군문화예술회관.xlsx"], t
        assert sum(1 for k, v in tdocs.items() if "_error" in v and "집계표 없음" in v["_error"]) >= 1, "내역서 없는 건축 공고는 건너뜀 표시"
        assert os.path.exists(data("logs_trades.parquet"))
        # 2차 배치 경로(가짜 Batches 클라이언트): 1건을 지우고 제출 → 상태 파일에 배치·출처파일 기록 → 수거 → 2차 파서로 해석·출처파일 복원
        from types import SimpleNamespace as NS
        tdocs.pop("R24010001-001")
        json.dump(tdocs, open(data("llm_trades.json"), "w", encoding="utf-8"), ensure_ascii=False)
        seen = {}

        class FakeBatches:
            def create(self, requests):
                seen["reqs"] = requests
                return NS(id="msgbatch_t1", processing_status="in_progress")

            def retrieve(self, bid):
                return NS(processing_status="ended", request_counts=NS(processing=0, succeeded=1, errored=0, canceled=0, expired=0))

            def results(self, bid):
                txt = json.dumps({"금액기준": "부가세포함", "관급자재_포함여부": "제외", "단위": "원", "합계_원": "33000000000",
                                  "공종별금액": [{"공종": "건축", "항목명_원문": "건축공사", "금액_원": "20000000000"}, {"공종": "기계설비", "항목명_원문": "기계설비공사", "금액_원": "3000000000"}],
                                  "출처": "총괄집계표", "신뢰도": "high", "근거문구": "r"}, ensure_ascii=False)
                return [NS(custom_id="R24010001-001", result=NS(type="succeeded", message=NS(content=[NS(type="text", text=txt)], usage=NS(input_tokens=10, output_tokens=5))))]
        fake_client = NS(messages=NS(batches=FakeBatches()))
        import anthropic
        with mock.patch.object(anthropic, "Anthropic", return_value=fake_client):
            out = _run(["extract-trades", "--batch", "--yes", "--config", cfg_path])
        assert "2차 배치 1개 제출 (공고 1건)" in out and seen["reqs"][0]["custom_id"] == "R24010001-001" and seen["reqs"][0]["params"]["system"] == extract_llm.TRADE_SYSTEM, out
        state = json.load(open(data("llm_trades_batch.json"), encoding="utf-8"))
        assert state["batches"][0]["kind"] == "trades" and state["batches"][0]["files"]["R24010001-001"] == ["공내역서_가상군문화예술회관.xlsx"]
        with mock.patch.object(anthropic, "Anthropic", return_value=fake_client):
            out = _run(["extract-trades", "--batch", "--config", cfg_path])
        assert "2차 배치 msgbatch_t1: 완료" in out and "2차 추출 성공 1건" in out, out
        tdocs = json.load(open(data("llm_trades.json"), encoding="utf-8"))
        t = tdocs["R24010001-001"]
        assert t["_files"] == ["공내역서_가상군문화예술회관.xlsx"] and t["공종별금액"][0]["금액_원"] == 20_000_000_000 and t["_공종합_원"] == 23_000_000_000 and t["_usage_in"] == 10, t
        assert json.load(open(data("llm_trades_batch.json"), encoding="utf-8"))["batches"][0]["status"] == "ended"
        out = _run(["extract", "--batch", "--config", cfg_path])
        assert "제출할 공고가 없습니다" in out and "Batches 50% 할인" in out, out
        ver = pd.read_parquet(data("logs_verify.parquet"))
        assert ((ver["공고번호"] == "R24010001") & (ver["항목"] == "추정가격") & (ver["판정"] == "정상")).any(), ver
        assert set(ver["판정"]) <= {"정상", "경고", "오류", "참고", "미확인"}
        # 보강 추출(extract-extra): 견적 → 즉시 실행(가짜 모델) → 1건을 지우고 배치 제출·수거. 1차 결과(llm_docs.json)는 그대로
        out = _run(["extract-extra", "--config", cfg_path])
        assert "보강 추출 대상" in out and "extract-extra --yes" in out and not os.path.exists(data("llm_extra.json")), out
        docs_before = open(data("llm_docs.json"), encoding="utf-8").read()
        with mock.patch.object(extract_llm, "extract_extra_with_claude", side_effect=_fake_extra):
            out = _run(["extract-extra", "--yes", "--config", cfg_path])
        assert "보강 추출 성공 3건 / 실패 0건" in out and "주차대수 1" in out and "기타금액 1" in out, out
        xdocs = json.load(open(data("llm_extra.json"), encoding="utf-8"))
        assert xdocs["R24010001-001"]["주차대수"] == 210 and xdocs["R24010002-000"]["기타금액_원"] == 38000000 and xdocs["R24010002-000"]["기타금액_위치"] == "별도"
        assert open(data("llm_docs.json"), encoding="utf-8").read() == docs_before, "보강 추출은 1차 결과를 건드리지 않는다"
        out = _run(["extract-extra", "--only", "전기", "--config", cfg_path])
        assert "공종 전기" in out and "이번 대상 0건" in out, out
        xdocs.pop("R24010002-000")
        json.dump(xdocs, open(data("llm_extra.json"), "w", encoding="utf-8"), ensure_ascii=False)

        class FakeExtraBatches:
            def create(self, requests):
                seen["xreqs"] = requests
                return NS(id="msgbatch_x1", processing_status="in_progress")

            def retrieve(self, bid):
                return NS(processing_status="ended", request_counts=NS(processing=0, succeeded=1, errored=0, canceled=0, expired=0))

            def results(self, bid):
                txt = json.dumps({"기타금액_원": "38000000", "기타금액_항목": "한전불입금 38,000,000", "기타금액_위치": "별도", "신뢰도": "high",
                                  "근거문구": [{"항목": "기타금액", "원문": "한전불입금 38,000,000원(별도)"}]}, ensure_ascii=False)
                return [NS(custom_id="R24010002-000", result=NS(type="succeeded", message=NS(content=[NS(type="text", text=txt)], usage=NS(input_tokens=7, output_tokens=3))))]
        fake_x = NS(messages=NS(batches=FakeExtraBatches()))
        with mock.patch.object(anthropic, "Anthropic", return_value=fake_x):
            out = _run(["extract-extra", "--batch", "--yes", "--config", cfg_path])
            assert "보강 배치 1개 제출 (공고 1건)" in out and seen["xreqs"][0]["params"]["system"] == extract_llm.EXTRA_SYSTEM, out
            assert "[공고문 텍스트]" in seen["xreqs"][0]["params"]["messages"][0]["content"]
            out = _run(["extract-extra", "--batch", "--config", cfg_path])
        assert "보강 배치 msgbatch_x1: 완료" in out and "보강 추출 성공 1건" in out, out
        xdocs = json.load(open(data("llm_extra.json"), encoding="utf-8"))
        assert xdocs["R24010002-000"]["기타금액_원"] == 38000000 and xdocs["R24010002-000"]["_usage_in"] == 7
        assert json.load(open(data("llm_extra_batch.json"), encoding="utf-8"))["batches"][0]["kind"] == "extra"
        # 8) excel — 생성 + 수식 오류 0. 2차 검증로그는 excel 이 저장된 결과에 현재 규칙을 다시 적용해 새로 만든다(수거 때 로그가 옛 규칙이어도)
        stale = pd.DataFrame([{"공고번호": "STALE", "항목": "집계표합계 기준", "판정": "경고", "비고": "옛 규칙"}])
        stale.to_parquet(data("logs_trades.parquet"), index=False)
        out = _run(["excel", "--config", cfg_path])
        tl = pd.read_parquet(data("logs_trades.parquet"))
        assert "STALE" not in set(tl["공고번호"]) and not tl.empty, "excel 단계가 2차 검증로그를 다시 만들어야 한다"
        xlsx = os.path.join(tmp, "output", "공사비DB.xlsx")
        assert os.path.exists(xlsx)
        import openpyxl
        wb = openpyxl.load_workbook(xlsx)
        assert wb["04_공사비DB_공종별"].max_row - 1 == len(latest) and wb["01_시설마스터"].max_row - 1 == 3
        ws4 = wb["04_공사비DB_공종별"]; h4 = [c.value for c in ws4[1]]
        r4 = next(r for r in range(2, ws4.max_row + 1) if ws4.cell(r, h4.index("공고번호") + 1).value == "R24010001")
        assert ws4.cell(r4, h4.index("관급자관급액_API") + 1).value == 2100000000, "API 관급자재 금액이 04 시트에"
        assert ws4.cell(r4, h4.index("예산금액_API") + 1).value == 35100000000
        assert ws4.cell(r4, h4.index("낙찰금액_API") + 1).value == 28500000000 and ws4.cell(r4, h4.index("낙찰률_API") + 1).value == 86.36, "낙찰 참고 컬럼(차수 001 로 연결)"
        assert ws4.cell(r4, h4.index("낙찰자") + 1).value == "가상건설(주)"
        assert str(ws4.cell(r4, h4.index("낙찰률(수식)") + 1).value).startswith("=IF(")
        inc = ws4.cell(r4, h4.index("포함공종_API") + 1).value
        assert inc == "건축·토목", f"통합발주 공고의 포함 공종(대표 차수 001 의 주공종명+부공종명): {inc!r}"
        doc_inc = (ws4.cell(r4, h4.index("포함공종_문서") + 1).value, ws4.cell(r4, h4.index("포함공종_출처") + 1).value,
                   ws4.cell(r4, h4.index("포함공종(종합)") + 1).value)
        assert doc_inc == ("건축·토목·조경", "규칙", "건축·토목·조경"), f"현장설명서 공사금액 표의 공종 나열(텍스트 규칙, 차수 병합 텍스트): {doc_inc!r}"
        r4e = next(r for r in range(2, ws4.max_row + 1) if ws4.cell(r, h4.index("공고번호") + 1).value == "R24010002")
        assert ws4.cell(r4e, h4.index("포함공종_API") + 1).value == "전기"
        ws5 = wb["05_공사비DB_시설합산"]; h5 = [c.value for c in ws5[1]]
        pid4 = ws4.cell(r4, h4.index("프로젝트ID") + 1).value
        r5 = next(r for r in range(2, ws5.max_row + 1) if ws5.cell(r, h5.index("프로젝트ID") + 1).value == pid4)
        b5 = ws5.cell(r5, h5.index("통합발주_포함공종") + 1).value
        assert b5 == "건축(+토목·조경)", f"05 프로젝트 단위 통합발주 표기(API+문서 합집합): {b5!r}"
        assert "부공종명" in [c.value for c in wb["03_공고목록_최신"][1]]
        h1 = [c.value for c in wb["01_시설마스터"][1]]
        ws1 = wb["01_시설마스터"]
        r1 = next(r for r in range(2, ws1.max_row + 1) if ws1.cell(r, h1.index("프로젝트ID") + 1).value == pid4)
        g1 = lambda name: ws1.cell(r1, h1.index(name) + 1).value
        assert g1("주차대수") == 210 and g1("건폐율_pct") == 15 and g1("용적률_pct") == 45.5 and g1("최고높이_m") == 24.5, "보강 추출 값이 01 공사개요에"
        assert g1("대지면적_m2") == 30000 and g1("설계특이사항") == "본동 1개동, 무대기계 포함" and g1("건축법상용도_문서") == "문화 및 집회시설(공연장)"
        assert g1("공사명") and str(g1("발주년월")) == "2024-01" and g1("발주년월_기준") == "건축 공고" and "건축" in str(g1("개요_출처")), (g1("공사명"), g1("발주년월"), g1("개요_출처"))
        assert ws4.cell(r4e, h4.index("기타금액_문서") + 1).value == 38000000 and str(ws4.cell(r4e, h4.index("기타금액_판정") + 1).value).startswith("가산")
        lg6 = pd.read_parquet(data("logs_extra.parquet"))
        assert ((lg6["항목"] == "기타금액(문서)") & (lg6["공고번호"] == "R24010002")).any() and (lg6["항목"] == "건폐율=건축면적/대지면적").any(), lg6
        assert h1.index("대지면적_m2") == h1.index("연면적_m2") + 1 and h1.index("건축면적_m2") == h1.index("대지면적_m2") + 1, "01 공사개요는 공사비 정보광장 순서(연면적·대지면적·건축면적)"
        assert h1.index("공사명") < h1.index("공사규모(수식)") < h1.index("연면적_m2") < h1.index("현장위치") < h1.index("발주년월") < h1.index("설계특이사항")
        wsd = wb["04b_공종세부내역(문서)"]; hd = [c.value for c in wsd[1]]
        drows = [{hd[i]: wsd.cell(r, i + 1).value for i in range(len(hd))} for r in range(2, wsd.max_row + 1) if wsd.cell(r, 1).value]
        labs = sorted(d["세부공종"] for d in drows if d["공고번호"] == "R24010001")
        assert labs == ["건축", "기계설비"], f"배치 수거로 바뀐 2차 결과(건축·기계설비)가 04b 에: {labs}"
        flags = {wsd.cell(r, hd.index("배분적용") + 1).value for r in range(2, wsd.max_row + 1) if wsd.cell(r, hd.index("공고키") + 1).value == "R24010001-001"}
        assert flags == {"Y"}, f"공종 행이 집계표를 설명하는 문서는 배분적용 Y: {flags}"
        mech = next(d for d in drows if d["공고번호"] == "R24010001" and d["세부공종"] == "기계설비")
        assert mech["금액_원_문서"] == 3_000_000_000 and str(mech["비율(수식)"]).startswith("=IF(OR(") and "SUMIFS" in str(mech["배분금액(수식)"])
        assert "공고키" in h4 and "문서배분_기계설비(수식)" in h5 and "건축_순건축추정(수식)" in h5
        hist_df = pd.read_parquet(data("notices_hist.parquet"))
        e_row = hist_df[hist_df["공고번호"] == "R24010002"].iloc[0]
        assert e_row["낙찰자"] == "가상전기(주)" and e_row["낙찰금액_API"] == 2574000000 and e_row["참가업체수"] == 7, "재개찰 2행 중 최신 개찰일시 행"
        assert hist_df[hist_df["공고번호"] == "R24030005"]["낙찰금액_API"].isna().all(), "낙찰 없는 공고는 공란"
        assert hist_df[hist_df["공고번호"] == "R24020004"]["이전공고번호"].iloc[0] == "R24010003" and \
            set(hist_df[hist_df["공고번호"] == "R24010001"]["부공종명"]) <= {"토목공사업 / 조경공사업", "토목공사업"}
        assert wb["06_검증로그"].max_row > 1 and wb["07_추출노트"].max_row > 1
        assert wb.sheetnames[-1] == "08_용역공고(참고)" and wb["08_용역공고(참고)"].max_row >= 3 and "용역 연결(08) 2줄" in out, (wb.sheetnames, out[-300:])
        h5s = [c.value for c in wb["05_공사비DB_시설합산"][1]]
        assert "용역_설계_추정가격(수식)" in h5s and "용역_폐기물처리_추정가격(수식)" in h5s and h5s.index("용역_공고수(수식)") < h5s.index("시설간_중복")
        assert h5s.index("발주완료_폐기물용역(수식)") == h5s.index("용역_공고수(수식)") + 1
        ws8 = wb["08_용역공고(참고)"]; h8 = [c.value for c in ws8[1]]
        age8 = {ws8.cell(r, h8.index("공고번호") + 1).value: ws8.cell(r, h8.index("공사수집종료까지_일수") + 1).value for r in range(2, 4)}
        last_cn = str(max(pd.read_parquet(data("notices_hist.parquet"))["공고일시"].astype(str)))[:10]
        assert last_cn == "2024-03-05", f"이 가짜 자료에서 워크북에 실리는 공사 공고의 가장 늦은 공고일: {last_cn}"
        # 08 의 일수 = 공사 공고 수집 종료일(= 실린 공사 공고의 가장 늦은 공고일. config 의 수집 종료월 2024-04 말일이 아니다) − 용역 공고일. 음수 = 그 뒤에 난 용역 공고
        assert age8 == {"R24S00001": -15, "R24S00002": 0}, f"08 의 일수: {age8}"
        from tests.check_excel import check
        rc = check(xlsx)
        assert rc in (0, 2), "Excel 수식 오류"          # 2 = 검증 도구 없음(LibreOffice/formulas) → 건너뜀
        # 05 시설합산 값 검증(formulas 패키지가 있을 때): 연면적은 01 시트에서, 총공사비는 04 SUMIFS 로
        try:
            import formulas  # type: ignore
            import logging, warnings
            warnings.filterwarnings("ignore"); logging.disable(logging.CRITICAL)
            sol = formulas.ExcelModel().loads(xlsx).finish().calculate()
            logging.disable(logging.NOTSET)

            def val(sheet, coord):
                for k, v in sol.items():
                    if k.upper().endswith(f"{sheet.upper()}'!{coord}"):
                        x = getattr(v, "value", v); x = x.tolist() if hasattr(x, "tolist") else x
                        while isinstance(x, list) and x:
                            x = x[0]
                        return x
            ws5 = wb["05_공사비DB_시설합산"]; h5 = [c.value for c in ws5[1]]
            col = lambda name: openpyxl.utils.get_column_letter(h5.index(name) + 1)
            rows = {ws5.cell(r, 1).value: r for r in range(2, ws5.max_row + 1)}
            r = rows[f"{gid}-N"]
            area, total, per = val("05_공사비DB_시설합산", f"{col('연면적_m2')}{r}"), val("05_공사비DB_시설합산", f"{col('총공사비합계_기초금액기준(수식)')}{r}"), val("05_공사비DB_시설합산", f"{col('㎡당공사비_기초금액기준_원(수식)')}{r}")
            assert float(area) == 15200.0, area
            # 건축(기초+관급) + 전기(+한전불입금) + 소방 + 같은 세부사업으로 합산된 관리사무소 건립공사(추정가격 4억 + 부가세 10% 가정)
            assert abs(float(total) - (33000000000 + 2100000000 + 2860000000 + 583000000 + 38000000 + 440000000)) < 1, total
            assert abs(float(per) - float(total) / 15200) < 1, per
            total_p = val("05_공사비DB_시설합산", f"{col('총공사비합계_추정가격기준(수식)')}{r}")
            assert abs(float(total_p) - float(total)) < 1, ("과세 공사는 추정가격+부가세 기준과 기초금액 기준이 같다", total_p, total)
            c4 = lambda name: openpyxl.utils.get_column_letter(h4.index(name) + 1)
            v4 = lambda name, rr: val("04_공사비DB_공종별", f"{c4(name)}{rr}")
            assert v4("부가세_출처(수식)", r4) == "API" and float(v4("부가세(수식)", r4)) == 3000000000 and v4("과세구분(수식)", r4) == "과세"
            assert float(v4("추정금액_계ⓐ(수식)", r4)) == 33000000000 and float(v4("관급자관급비ⓑ(수식)", r4)) == 2100000000
            assert float(v4("금액합계_추정가격기준(수식)", r4)) == 35100000000 == float(v4("금액합계_기초금액기준(수식)", r4)), "요청 양식의 금액 합계 = 예산금액(API)"
            r3 = rows[ids_before["다른군"] + "-N"]
            assert abs(float(val("05_공사비DB_시설합산", f"{col('건축(수식)')}{r3}")) - 25000000000 * 1.1) < 1, "기초금액·부가세가 없으면 추정가격 + 10% 가정"
            rate = val("04_공사비DB_공종별", f"{openpyxl.utils.get_column_letter(h4.index('낙찰률(수식)') + 1)}{r4}")
            # 04b 비율·배분과 05 문서배분: 기계설비 3/26, 배분 = 비율 × 04 총공사비(R24010001), 05 건축_순건축추정 = 건축(수식) − 비건축 배분
            dcol = lambda name: openpyxl.utils.get_column_letter(hd.index(name) + 1)
            rd = next(r for r in range(2, wsd.max_row + 1) if wsd.cell(r, hd.index("공고번호") + 1).value == "R24010001"
                      and wsd.cell(r, hd.index("세부공종") + 1).value == "기계설비")
            ratio = float(val("04b_공종세부내역(문서)", f"{dcol('비율(수식)')}{rd}"))
            assert abs(ratio - 3 / 23) < 1e-6, ratio          # 배치 결과: 건축 200억 + 기계설비 30억 → 기계설비 비율 3/23
            tot04 = float(val("04_공사비DB_공종별", f"{openpyxl.utils.get_column_letter(h4.index('금액합계_기초금액기준(수식)') + 1)}{r4}"))
            alloc = float(val("04b_공종세부내역(문서)", f"{dcol('배분금액(수식)')}{rd}"))
            assert abs(alloc - ratio * tot04) < 1, (alloc, ratio, tot04)
            r5a = next(r for r in range(2, ws5.max_row + 1) if ws5.cell(r, h5.index("프로젝트ID") + 1).value == pid4)
            assert abs(float(val("05_공사비DB_시설합산", f"{col('문서배분_기계설비(수식)')}{r5a}")) - alloc) < 1
            net = float(val("05_공사비DB_시설합산", f"{col('건축_순건축추정(수식)')}{r5a}"))
            arch = float(val("05_공사비DB_시설합산", f"{col('건축(수식)')}{r5a}"))
            assert abs(net - (arch - tot04 * (3 / 23))) < 1, (net, arch, tot04)
            assert val("05_공사비DB_시설합산", f"{col('문서배분_중복경고(수식)')}{r5a}") in ("", None), "분리발주 중복 없음"
            assert abs(float(rate) - 28500000000 / 33000000000) < 1e-6, ("낙찰률(수식) = 낙찰금액/기초금액", rate)
            assert abs(float(total) - (33000000000 + 2100000000 + 2860000000 + 583000000 + 38000000 + 440000000)) < 1, "낙찰금액은 총공사비에 쓰이지 않음"
            assert float(v4("기타금액ⓒ(수식)", r4e)) == 38000000 and float(v4("금액합계_기초금액기준(수식)", r4e)) == 2860000000 + 38000000, "전기 공고의 한전불입금이 금액 합계에 가산"
            assert float(val("05_공사비DB_시설합산", f"{col('용역_건설사업관리·감리_추정가격(수식)')}{r}")) == 1500000000
            assert float(val("05_공사비DB_시설합산", f"{col('용역_폐기물처리_추정가격(수식)')}{r}")) == 129834000
            assert val("05_공사비DB_시설합산", f"{col('용역_설계_추정가격(수식)')}{r}") in ("", None) and int(val("05_공사비DB_시설합산", f"{col('용역_공고수(수식)')}{r}")) == 2
            assert val("05_공사비DB_시설합산", f"{col('발주완료_폐기물용역(수식)')}{r}") == "발주완료(최근)", "폐기물처리 용역(합산여부 Y)이 있는 프로젝트 — 공고가 마지막 공사 공고일의 6개월 안쪽·그 뒤라 '최근'"
            assert val("05_공사비DB_시설합산", f"{col('발주완료_폐기물용역(수식)')}{r3}") in ("", None), "용역이 이어지지 않은 프로젝트는 빈칸"
            print("05 시설합산 값 검증 OK")
        except ImportError:
            print("(formulas 패키지 없음 → 05 값 검증 생략)")
        print(f"OK: 통합 실행 통과 (임시 폴더 {tmp}, API 호출 {sum(v for k, v in server.calls.items() if not k.startswith('file:'))}회)")
    finally:
        server.stop()
        load_config.cache_clear()


if __name__ == "__main__":
    run()
