"""합성 데이터로 S2(분류)·S4(중복정리)·S7(Excel) 검증. 실행: python -m tests.test_dedup_and_excel
※ 아래 공고는 모두 가상의 예시(실제 공고 아님). API 응답 형식만 모사.
"""
from __future__ import annotations

import os
import sys

import pandas as pd

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
from g2b_costdb import dedup, discover  # noqa: E402
from g2b_costdb.build_excel import build_workbook  # noqa: E402
from g2b_costdb.classify import (classify_trade, classify_work_type, extract_facility_name, load_config,  # noqa: E402
                                 normalize_facility_key)
from g2b_costdb.pipeline import assemble_trade_table, enrich_facility  # noqa: E402

RAW = [  # API 필드명(bidNtceNo …)으로 모사한 가상 공고
    dict(bidNtceNo="20240100001", bidNtceOrd="00", bidNtceNm="가상시립미술관 건립공사(건축)", ntceKindNm="일반", reNtceYn="N",
         bidNtceDt="2024-03-05 10:00", dminsttNm="가상광역시", ntceInsttNm="가상광역시", presmptPrce="18200000000",
         mainCnsttyNm="건축공사", cnstrtsiteRgnNm="가상광역시 중구", bidNtceDtlUrl="https://www.g2b.go.kr/x/1",
         ntceSpecDocUrl1="https://example.invalid/a.hwp", ntceSpecFileNm1="입찰공고문.hwp"),
    dict(bidNtceNo="20240100001", bidNtceOrd="01", bidNtceNm="가상시립미술관 건립공사(건축) [변경]", ntceKindNm="변경", reNtceYn="N",
         bidNtceDt="2024-03-12 10:00", dminsttNm="가상광역시", ntceInsttNm="가상광역시", presmptPrce="18200000000",
         mainCnsttyNm="건축공사", cnstrtsiteRgnNm="가상광역시 중구", bidNtceDtlUrl="https://www.g2b.go.kr/x/1"),
    dict(bidNtceNo="20240200002", bidNtceOrd="00", bidNtceNm="가상시립미술관 건립공사(건축) 재공고", ntceKindNm="재공고", reNtceYn="Y",
         bidNtceDt="2024-05-20 10:00", dminsttNm="가상광역시", ntceInsttNm="가상광역시", presmptPrce="19100000000",
         mainCnsttyNm="건축공사", cnstrtsiteRgnNm="가상광역시 중구", bidNtceDtlUrl="https://www.g2b.go.kr/x/2"),
    dict(bidNtceNo="20240100003", bidNtceOrd="00", bidNtceNm="가상시립미술관 건립 전기공사", ntceKindNm="일반", reNtceYn="N",
         bidNtceDt="2024-03-06 10:00", dminsttNm="가상광역시", ntceInsttNm="가상광역시", presmptPrce="1450000000",
         mainCnsttyNm="전기공사", cnstrtsiteRgnNm="가상광역시 중구"),
    dict(bidNtceNo="20240100004", bidNtceOrd="00", bidNtceNm="가상시립미술관 건립 정보통신공사", ntceKindNm="일반", reNtceYn="N",
         bidNtceDt="2024-03-06 11:00", dminsttNm="가상광역시", ntceInsttNm="가상광역시", presmptPrce="620000000",
         mainCnsttyNm="정보통신공사", cnstrtsiteRgnNm="가상광역시 중구"),
    dict(bidNtceNo="20240100005", bidNtceOrd="00", bidNtceNm="가상시립미술관 건립 소방시설공사", ntceKindNm="취소", reNtceYn="N",
         bidNtceDt="2024-03-07 10:00", dminsttNm="가상광역시", ntceInsttNm="가상광역시", presmptPrce="510000000",
         mainCnsttyNm="소방공사", cnstrtsiteRgnNm="가상광역시 중구"),
    dict(bidNtceNo="20240300006", bidNtceOrd="00", bidNtceNm="가상시립미술관 건립 소방시설공사(재공고)", ntceKindNm="일반", reNtceYn="Y",
         bidNtceDt="2024-04-15 10:00", dminsttNm="가상광역시", ntceInsttNm="가상광역시", presmptPrce="530000000",
         mainCnsttyNm="소방공사", cnstrtsiteRgnNm="가상광역시 중구"),
    dict(bidNtceNo="20230900007", bidNtceOrd="00", bidNtceNm="가상시립미술관 옥상 방수공사", ntceKindNm="일반", reNtceYn="N",
         bidNtceDt="2023-09-01 10:00", dminsttNm="가상광역시", ntceInsttNm="가상광역시", presmptPrce="80000000",
         mainCnsttyNm="건축공사", cnstrtsiteRgnNm="가상광역시 중구"),
    dict(bidNtceNo="20260500011", bidNtceOrd="00", bidNtceNm="가상시립미술관 노후시설 개보수공사(건축)", ntceKindNm="일반", reNtceYn="N",
         bidNtceDt="2026-05-02 10:00", dminsttNm="가상광역시", ntceInsttNm="가상광역시", presmptPrce="4200000000",
         mainCnsttyNm="건축공사", cnstrtsiteRgnNm="가상광역시 중구"),
    dict(bidNtceNo="20260500012", bidNtceOrd="00", bidNtceNm="가상시립미술관 전기공사", ntceKindNm="일반", reNtceYn="N",
         bidNtceDt="2026-05-03 10:00", dminsttNm="가상광역시", ntceInsttNm="가상광역시", presmptPrce="380000000",
         mainCnsttyNm="전기공사", cnstrtsiteRgnNm="가상광역시 중구"),
    dict(bidNtceNo="20250100008", bidNtceOrd="00", bidNtceNm="예시군 문화예술회관 건립공사", ntceKindNm="일반", reNtceYn="N",
         bidNtceDt="2025-02-10 10:00", dminsttNm="예시군", ntceInsttNm="예시군", presmptPrce="32000000000",
         mainCnsttyNm="건축공사", cnstrtsiteRgnNm="전라남도 예시군"),
    dict(bidNtceNo="20250100009", bidNtceOrd="00", bidNtceNm="예시군 문화예술회관 건립 전기공사", ntceKindNm="일반", reNtceYn="N",
         bidNtceDt="2025-02-10 11:00", dminsttNm="예시군", ntceInsttNm="예시군", presmptPrce="2600000000",
         mainCnsttyNm="전기공사", cnstrtsiteRgnNm="전라남도 예시군"),
    dict(bidNtceNo="20250100010", bidNtceOrd="00", bidNtceNm="샘플시 공영주차장 노면 도색공사", ntceKindNm="일반", reNtceYn="N",
         bidNtceDt="2025-03-01 10:00", dminsttNm="샘플시", ntceInsttNm="샘플시", presmptPrce="30000000",
         mainCnsttyNm="건축공사", cnstrtsiteRgnNm="경기도 샘플시"),
]
BSIS = [dict(bidNtceNo=r["bidNtceNo"], bidNtceOrd=r["bidNtceOrd"], bssamt=str(int(int(r["presmptPrce"]) * 1.1))) for r in RAW]
DOCS = {  # LLM 추출 결과 모사 (가상)
    "20260500011-00": {"연면적_m2": 6100.0, "지하층수": 1, "지상층수": 4, "구조": "철근콘크리트조", "용도": "미술관", "사업유형": "리모델링",
                       "공사기간_일": 420, "추정가격_원": 4200000000, "기초금액_원": 4620000000, "관급자관급액_원": 300000000,
                       "신뢰도": "medium", "근거문구": {"연면적_m2": "리모델링 대상면적 6,100㎡"}},
    "20240200002-00": {"연면적_m2": 9850.0, "건축면적_m2": 3200.0, "지하층수": 2, "지상층수": 4, "구조": "철근콘크리트조",
                       "용도": "미술관", "공사기간_일": 900, "추정가격_원": 19100000000, "기초금액_원": 21010000000,
                       "도급자관급액_원": 850000000, "관급자관급액_원": 1200000000, "신뢰도": "high",
                       "근거문구": {"연면적_m2": "연면적 9,850㎡", "추정가격_원": "추정가격 19,100,000,000원"}},
    "20250100008-00": {"연면적_m2": 15200.0, "지하층수": 1, "지상층수": 3, "구조": "철골철근콘크리트조", "용도": "공연장",
                       "공사기간_일": 1080, "추정가격_원": 32000000000, "기초금액_원": 35200000000, "관급자관급액_원": 2100000000,
                       "신뢰도": "high", "근거문구": {"연면적_m2": "연면적 15,200㎡"}},
}


def run(out_path: str):
    cfg = load_config()
    # 분류 단위 점검
    assert extract_facility_name("가상시립미술관 건립 전기공사", "미술관") == "가상시립미술관"
    assert extract_facility_name("예시군 문화예술회관 건립공사(건축)", "문화예술회관") == "예시군 문화예술회관"
    assert classify_work_type("가상시립미술관 옥상 방수공사") == "유지보수"
    assert classify_work_type("가상시립미술관 노후시설 개보수공사(건축)") == "리모델링"
    assert classify_work_type("OO미술관 증축 및 리모델링공사") == "증축·리모델링"
    assert classify_work_type("예시군 문화예술회관 건립공사") == "신축"
    assert classify_trade("가상시립미술관 건립 정보통신공사", "정보통신공사", "")[0] == "정보통신"
    assert classify_trade("가상시립미술관 건립공사(건축)", "", "건축공사업")[0] == "건축"
    assert normalize_facility_key("가상시립미술관 건립공사 (재공고)") == normalize_facility_key("가상시립미술관 건립 전기공사")

    std = discover.standardize(pd.DataFrame(RAW), cfg, pd.DataFrame(BSIS))
    agg = discover.discover_candidates(std)
    assert set(agg["시설명_후보"]) >= {"가상시립미술관", "예시군 문화예술회관"}
    reviewed = agg[agg["검수_포함여부"] == "포함"].copy()
    hits = discover.research_by_facility(std, reviewed)
    assert "20230900007" not in set(hits["공고번호"]), "유지보수 공고는 재검색에서 제외되어야 함"

    hist, latest, logs = dedup.dedup_latest(hits, cfg["dedup"]["price_change_warn_ratio"])
    rep = latest.set_index("프로젝트키")
    mus_id = reviewed.loc[reviewed["검수_시설명"] == "가상시립미술관", "시설ID"].iloc[0]
    assert rep.loc[f"{mus_id}-N|건축|", "공고번호"] == "20240200002", "신축 프로젝트: 재공고(최신)가 대표여야 함"
    assert rep.loc[f"{mus_id}-R|건축|", "공고번호"] == "20260500011", "리모델링은 신축과 별도 프로젝트여야 함"
    assert rep.loc[f"{mus_id}-R|전기|", "공고번호"] == "20260500012", "미분류 전기공사는 가까운 건축공고(리모델링) 유형을 상속"
    assert rep.loc[f"{mus_id}-N|전기|", "공고번호"] == "20240100003"
    fire = [k for k in rep.index if "|소방|" in k]
    assert rep.loc[fire[0], "공고번호"] == "20240300006", "취소공고 제외 후 재공고가 대표여야 함"
    assert (hist["공고종류"] == "취소").sum() == 1 and (hist["최신여부"]).sum() == len(latest)
    assert set(dedup.facility_summary(latest, hist)["프로젝트ID"]) >= {f"{mus_id}-N", f"{mus_id}-R"}

    fac = enrich_facility(dedup.facility_summary(latest, hist), latest, DOCS)
    trade = assemble_trade_table(latest, DOCS)
    notes = pd.DataFrame([{"공고번호": "20240200002", "파일명": "입찰공고문.hwp", "형식": ".hwp", "우선순위": 10, "다운로드": "Y",
                           "추출성공": "Y", "추출글자수": 18234, "파서": "olefile", "LLM호출": "Y", "오류": "", "URL": "https://example.invalid/a.hwp"}])
    os.makedirs(os.path.dirname(out_path), exist_ok=True)
    build_workbook(out_path, fac, hist, latest, trade, pd.DataFrame(logs), notes, meta={"period": "합성 데이터(예시)"})
    print(f"OK: 전체 {len(hist)}건 / 대표 {len(latest)}건 / 시설 {len(fac)}개 / 경고 {len(logs)}건 → {out_path}")
    return hist, latest, fac, trade


if __name__ == "__main__":
    run(os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), "sample_output", "공사비DB_샘플.xlsx"))
    from tests.test_regressions import run as run_regressions  # noqa: E402
    run_regressions()
