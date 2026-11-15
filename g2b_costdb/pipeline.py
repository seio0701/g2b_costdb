"""파이프라인 CLI. g2b_costdb 폴더에서 실행한다.

python -m g2b_costdb.pipeline doctor                     # 환경 점검(파이썬·패키지·키 존재 여부·네트워크)
python -m g2b_costdb.pipeline probe --ym 2026-08          # 필드명 확인 + config.yaml 매핑 대조
python -m g2b_costdb.pipeline collect                     # S1 전량수집(공고+기초금액)
python -m g2b_costdb.pipeline discover                    # S2 후보 시설 검수용 Excel 생성
#  → output/facility_candidates.xlsx 검수(노란 셀) 후
python -m g2b_costdb.pipeline research                    # S3 시설명 재검색 + 면허제한 조회
python -m g2b_costdb.pipeline dedup                       # S4 최신 공고 선별
python -m g2b_costdb.pipeline attach                      # S5 첨부 다운로드·텍스트 추출
python -m g2b_costdb.pipeline extract                     # S6 비용 견적만 출력 (--yes: API 호출 / --batch --yes: 배치 제출 후 --batch 로 수거 / --export → --import: 파일 인수인계)
python -m g2b_costdb.pipeline extract-trades              # (선택) 2차 추출: 내역서 집계표의 공종별 금액 → 04b (실행 방식은 extract 와 같음)
python -m g2b_costdb.pipeline extract-extra               # (선택) 보강 추출: 기타금액(한전불입금 등)·면세 문구·공사개요 세부 항목 → 01·04 (--only 전기,건축 / --limit N)
python -m g2b_costdb.pipeline excel                       # S7 Excel DB (검증로그는 저장된 추출 결과에 현재 규칙을 다시 적용해 새로 만든다)
python -m g2b_costdb.pipeline probe --servc --ym 2026-08  # 용역 입찰공고 오퍼레이션 시험 호출(필드명·분류명 확인)
python -m g2b_costdb.pipeline collect --servc             # 용역 입찰공고 수집(설계·건설사업관리·폐기물처리만 보존)
"""
from __future__ import annotations

import argparse
import importlib
import importlib.util
import json
import logging
import os
import platform
import re
import shutil
import sys

from typing import Dict, List, Optional, Tuple

import pandas as pd

from . import attachments, collect, dedup, discover, extract_llm, classify
from .api_client import ApiError, DailyBudgetExceeded
from .build_excel import build_workbook
from .classify import ROOT, load_config

log = logging.getLogger("g2b_costdb")


def _p(cfg, *parts):
    return os.path.join(cfg["paths"]["data_dir"], *parts)


def _read(cfg, name):
    path = _p(cfg, name)
    if not os.path.exists(path):
        raise SystemExit(f"중간 산출물이 없습니다: {path} (이전 단계를 먼저 실행)")
    return pd.read_parquet(path)


def _read_raw_notices(cfg):
    """notices_all.parquet 에서 standardize 가 쓰는 컬럼만 읽는다(전량 145개 컬럼 × 백만 행이면 메모리 수 GB 절약)."""
    path = _p(cfg, "notices_all.parquet")
    if not os.path.exists(path):
        raise SystemExit(f"중간 산출물이 없습니다: {path} (이전 단계를 먼저 실행)")
    try:
        import pyarrow.parquet as pq
        have = set(pq.read_schema(path).names)
    except Exception:
        return pd.read_parquet(path)
    cols = [c for c in discover.raw_columns_needed(cfg) if c in have]
    return pd.read_parquet(path, columns=cols) if cols else pd.read_parquet(path)


LOG_COLS = ["공고번호", "항목", "API값", "문서값", "차이", "차이율", "판정", "비고"]


def _logs_frame(logs) -> pd.DataFrame:
    """검증로그 rows → DataFrame. API값/문서값/차이/차이율은 숫자 컬럼으로 통일하고(parquet 혼합타입 방지),
    문자열 값(예: 사업유형 비교)은 비고에 옮겨 적는다."""
    rows = []
    for l in logs or []:
        r = {c: l.get(c) for c in LOG_COLS}
        extra = []
        for c in ("API값", "문서값", "차이", "차이율"):
            v = r.get(c)
            if v is None or isinstance(v, bool):
                r[c] = None
                continue
            try:
                x = float(v)
                r[c] = None if x != x else x
            except (TypeError, ValueError):
                extra.append(f"{c}={v}")
                r[c] = None
        if extra:
            r["비고"] = "; ".join(extra + ([str(r["비고"])] if r.get("비고") else []))
        r["공고번호"] = "" if r.get("공고번호") is None else str(r["공고번호"])
        rows.append(r)
    df = pd.DataFrame(rows, columns=LOG_COLS)
    for c in ("API값", "문서값", "차이", "차이율"):
        df[c] = pd.to_numeric(df[c], errors="coerce").astype(float)
    for c in ("공고번호", "항목", "판정", "비고"):
        df[c] = df[c].map(lambda v: "" if v is None or (isinstance(v, float) and v != v) else str(v)).astype(object)
    return df


def _excel_busy(path: str) -> str:
    return (f"[중단] '{path}' 파일을 열 수 없습니다. Excel 에서 이 파일이 열려 있으면 닫은 뒤 같은 명령을 다시 실행하세요.")


# ── doctor ──────────────────────────────────────────────────────
def stage_doctor(cfg):
    print(f"Python {platform.python_version()} ({sys.executable}) / OS {platform.system()} {platform.release()}")
    ok = tuple(int(x) for x in platform.python_version_tuple()[:2]) >= (3, 10)
    print(f"  파이썬 3.10 이상: {'OK' if ok else '아님 → 3.10+ 설치 필요'}")
    print("필수 패키지:")
    for mod, pkg in [("requests", "requests"), ("pandas", "pandas"), ("pyarrow", "pyarrow"), ("openpyxl", "openpyxl"),
                     ("yaml", "PyYAML"), ("olefile", "olefile"), ("pdfplumber", "pdfplumber"), ("docx", "python-docx"),
                     ("anthropic", "anthropic")]:
        try:
            m = importlib.import_module(mod)
            print(f"  {pkg:12s} OK {getattr(m, '__version__', '')}")
        except BaseException as e:  # noqa: BLE001 — 네이티브 확장 오류(PanicException 등)도 보고
            print(f"  {pkg:12s} 문제 → pip install --upgrade {pkg}  ({type(e).__name__}: {str(e)[:60]})")
    print("선택 패키지(폴백용):")
    for mod, pkg, use in [("pytesseract", "pytesseract", "스캔 PDF OCR"), ("pdf2image", "pdf2image", "스캔 PDF OCR"),
                          ("win32com", "pywin32", "Windows 한컴 COM 폴백")]:
        try:
            importlib.import_module(mod)
            print(f"  {pkg:12s} OK ({use})")
        except BaseException:  # noqa: BLE001
            print(f"  {pkg:12s} 없음 ({use}; 필요 시 pip install {pkg})")
    hwp5 = shutil.which("hwp5txt") or (importlib.util.find_spec("hwp5") is not None)
    print(f"  pyhwp        {'OK (hwp5txt 사용 가능)' if hwp5 else '없음 (HWP 폴백; pip install pyhwp)'}")
    print("환경변수(값은 출력하지 않음):")
    for env in (cfg["api"]["service_key_env"], cfg["llm"].get("api_key_env", "ANTHROPIC_API_KEY")):
        v = os.environ.get(env, "")
        state = "설정됨" if v.strip() else "없음"
        extra = ""
        if v and "%" in v:
            extra = " (URL 인코딩된 키로 보임 → 자동 변환하지만 가급적 Decoding 키를 사용)"
        if v and (v.strip() != v or v.strip('"') != v):
            extra += " (앞뒤 공백/따옴표 포함 → 제거 권장)"
        print(f"  {env}: {state}{extra}")
    print("경로:")
    for k in ("data_dir", "raw_dir", "text_dir", "files_dir", "out_dir"):
        p = cfg["paths"][k]
        try:
            os.makedirs(p, exist_ok=True)
            print(f"  {k:10s} {p} (쓰기 가능)")
        except OSError as e:
            print(f"  {k:10s} {p} 생성 실패: {e}")
    print(f"수집 기간: {cfg['period']['start']} ~ {cfg['period']['end']} "
          f"({len(collect.G2BClient.month_ranges(cfg['period']['start'], cfg['period']['end']))}개월)")
    import requests
    base = cfg["api"]["base_url"]
    try:
        r = requests.get(base + "/" + cfg["api"]["ops"]["cnstwk_list"], params={"serviceKey": "test", "numOfRows": "1", "pageNo": "1",
                         "inqryDiv": "1", "inqryBgnDt": "202601010000", "inqryEndDt": "202601312359", "type": "json"}, timeout=15)
        from .api_client import explain_code, parse_portal_error
        e = parse_portal_error(r.text)
        if e and e["code"] in ("30", "20", "33"):
            note = f", 포털 응답 코드 {e['code']} — 키 없이 시험 호출했으므로 정상"
        elif e:
            note = f", 포털 응답 코드 {e['code']}: {explain_code(e['code'])}"
        else:
            note = ""
        print(f"네트워크: {base} 연결 OK (HTTP {r.status_code}{note})")
    except Exception as e:  # noqa: BLE001
        print(f"네트워크: {base} 연결 실패 → {type(e).__name__}: {str(e)[:120]}")
    cache = cfg["paths"]["cache_db"]
    if os.path.exists(cache):
        import sqlite3
        from datetime import date
        con = sqlite3.connect(cache)
        row = con.execute("SELECT n FROM calls WHERE day=?", (date.today().isoformat(),)).fetchone()
        n_cache = con.execute("SELECT COUNT(*) FROM cache").fetchone()[0]
        print(f"API 캐시: 응답 {n_cache}건 저장됨, 오늘 호출 {row[0] if row else 0}회 / 예산 {cfg['api']['daily_call_budget']}")
    for name in ("notices_all.parquet", "bsis_all.parquet", "award_all.parquet", "notices_std.parquet", "notices_research.parquet",
                 "notices_latest.parquet", "texts.json", "llm_docs.json"):
        p = _p(cfg, name)
        if os.path.exists(p):
            print(f"산출물: {name} ({os.path.getsize(p) // 1024} KB)")


# ── stages ──────────────────────────────────────────────────────
def stage_collect(cfg):
    collect.collect_notices(cfg)
    collect.collect_bsis_amounts(cfg)
    if cfg["api"].get("use_awards"):
        try:
            collect.collect_awards(cfg)
        except ApiError as e:
            # 공고·기초금액은 이미 저장됨. 낙찰정보서비스만 미승인/End Point 오류인 경우 여기서 멈추지 않고 안내만 한다
            if e.code in ("20", "30", "31", "32", "12"):
                print(f"[낙찰정보] 수집 실패: API 오류 {e.code} ({e}) — 「조달청_나라장터 낙찰정보서비스」 활용신청 승인 여부와 "
                      f"config.yaml api.award_base_url(End Point)을 확인. 확인 전까지는 use_awards: false 로 두면 이 안내 없이 진행됩니다")
            else:
                raise
    else:
        print("[안내] 낙찰정보(낙찰금액·낙찰률)는 config.yaml api.use_awards: true 로 켜면 함께 수집됩니다(낙찰정보서비스 활용신청 필요)")


def stage_discover(cfg, fresh: bool = False):
    for prefix, label in (("notices_", "공고"), ("bsis_", "기초금액")) + ((("award_", "낙찰"),) if cfg["api"].get("use_awards") else ()):
        miss = collect.incomplete_months(cfg, prefix)
        if miss:
            print(f"[경고] {label} 수집 미완료 월 {len(miss)}개 ({', '.join(miss[:8])}{' …' if len(miss) > 8 else ''}) — "
                  f"부분 데이터로 후보를 만듭니다. 전체 결과가 필요하면 collect 를 먼저 완료하세요.")
    raw = _read_raw_notices(cfg)
    bsis_path = _p(cfg, "bsis_all.parquet")
    bsis = pd.read_parquet(bsis_path) if os.path.exists(bsis_path) else None
    award_path = _p(cfg, "award_all.parquet")
    awards = pd.read_parquet(award_path) if os.path.exists(award_path) else None
    std = discover.standardize(raw, cfg, bsis, awards=awards)
    if awards is not None:
        print(f"낙찰정보 연결: 공고 {int(std['낙찰금액_API'].notna().sum())}건에 낙찰금액 부여 (낙찰 목록 {len(awards)}건)")
    std.to_parquet(_p(cfg, "notices_std.parquet"), index=False)
    out = os.path.join(cfg["paths"]["out_dir"], "facility_candidates.xlsx")
    try:
        prev = discover.read_review_file(out)
    except PermissionError:
        raise SystemExit(_excel_busy(out))
    if fresh and prev is not None and not prev.empty:
        print("[--fresh] 이전 검수 파일에서 시설ID만 이어받고 검수 칸은 모두 새 기본값으로 채웁니다")
    agg = discover.discover_candidates(std, previous_review=prev, ids_only=fresh)
    if agg.empty:
        raise SystemExit("후보 시설 0건 — keywords.yaml 의 검색어(include)를 넓히거나 수집 기간을 확인하세요.")
    try:
        discover.write_review_workbook(agg, out)
    except PermissionError:
        raise SystemExit(_excel_busy(out))
    print(f"공고 {len(std)}건 중 후보 시설 {len(agg)}개 → {out} 에서 검수(노란 셀) 후 `research` 실행")
    print("표2 대분류별 시설 수:\n" + agg.groupby("표2_대분류").size().to_string())
    print("사업유형별 시설 수:\n" + agg.groupby("사업유형").size().to_string())


def stage_research(cfg):
    std = _read(cfg, "notices_std.parquet")
    review_path = os.path.join(cfg["paths"]["out_dir"], "facility_candidates.xlsx")
    try:
        reviewed = discover.read_reviewed(review_path)
    except PermissionError:
        raise SystemExit(_excel_busy(review_path))
    if reviewed.empty:
        raise SystemExit("검수_포함여부가 '포함'인 시설이 없습니다. facility_candidates.xlsx 를 검수한 뒤 저장했는지 확인하세요.")
    hits = discover.research_by_facility(std, reviewed)
    if hits.empty:
        raise SystemExit("재검색 결과 없음")
    # 면허제한: config api.use_license_limit 가 true 일 때만 공고별 조회(호출량 큼). 기본은 목록 응답의 주공종명·부공종명(업종명) 사용
    lic, col = None, None
    if cfg["api"].get("use_license_limit", False):
        lic = collect.collect_license_limits(cfg, sorted(hits["공고번호"].astype(str).unique()))
        col = collect.license_name_column(cfg, lic)
    if not col and "부공종명" in hits.columns:
        from .classify import classify_trade
        # 주공종명이 건축이면 주공종명만 본다('건축공사업 / 조경공사업' 이 부공종 때문에 조경이 되던 문제). 그 밖에는 주공종명 + 부공종명
        is_arch = lambda mc: classify.match_trade_label(mc or "", classify.load_keywords()["trade_rules"]) == "건축"
        tr = [classify_trade(n, (mc or "") if is_arch(mc) else (mc or "") + (" / " + sub if sub else ""), l)
              for n, mc, sub, l in zip(hits["공고명"], hits["주공종명"], hits["부공종명"], hits["면허제한업종"])]
        hits["공종"], hits["공종근거"] = [t[0] for t in tr], [t[1] for t in tr]
    if col:
        m = lic.groupby("_bid_no")[col].apply(lambda s: " / ".join(sorted(set(str(x) for x in s if x))))
        hits["면허제한업종"] = hits["공고번호"].astype(str).map(m).fillna("")
        from .classify import classify_trade
        tr = [classify_trade(n, mc, l) for n, mc, l in zip(hits["공고명"], hits["주공종명"], hits["면허제한업종"])]
        hits["공종"], hits["공종근거"] = [t[0] for t in tr], [t[1] for t in tr]
        if col != cfg["fields"].get("license_name"):
            print(f"[안내] 면허제한 업종명 필드로 '{col}' 를 사용했습니다. config.yaml fields.license_name 을 이 값으로 바꾸세요.")
    elif lic is not None and not lic.empty:
        print(f"[경고] 면허제한 응답에서 업종명 필드를 찾지 못했습니다. 응답 컬럼: {list(lic.columns)[:15]} → config.yaml fields.license_name 확인")
    elif cfg["api"].get("use_license_limit", False):
        print("[안내] 면허제한 정보 없음(조회 실패 또는 예산 소진) → 공종은 주공종명·부공종명·공고명 규칙으로 분류")
    else:
        print("[안내] 공종은 목록 응답의 주공종명·부공종명(업종명)과 공고명 규칙으로 분류 (면허제한 API 미사용: config api.use_license_limit)")
    hits.to_parquet(_p(cfg, "notices_research.parquet"), index=False)
    print(f"재검색 공고 {len(hits)}건 저장 (시설 {hits['시설ID'].nunique()}개)")
    piv = hits.pivot_table(index=["시설ID", "시설명"], columns="공종", values="공고번호", aggfunc="nunique", fill_value=0)
    print("시설별 공종 분포(공고 수):\n" + piv.head(40).to_string())


def stage_dedup(cfg):
    hits = _read(cfg, "notices_research.parquet")
    hist, latest, logs = dedup.dedup_latest(hits, cfg["dedup"]["price_change_warn_ratio"])
    hist.to_parquet(_p(cfg, "notices_hist.parquet"), index=False)
    latest.to_parquet(_p(cfg, "notices_latest.parquet"), index=False)
    # 같은 공고가 여러 시설의 대표로 잡힌 경우(복합 발주: '충남미술관 및 공영주차장 건립공사' 등) → 05 시트에서 시설마다 집계되므로 검수자가 판단
    multi = []
    if len(latest) and "시설ID" in latest.columns:
        g = latest.groupby("공고키")["시설ID"].agg(lambda x: sorted(set(map(str, x))))
        for key, ids in g[g.map(len) > 1].items():
            r = latest[latest["공고키"] == key].iloc[0]
            multi.append({"공고번호": str(r["공고번호"]), "프로젝트키": "/".join(ids), "항목": "복수 시설 배정",
                          "비고": f"'{r['공고명']}' 이 시설 {len(ids)}개({', '.join(ids)})의 대표 공고 — 05 시트에 시설마다 집계됨. 한 시설만 남기려면 검수에서 다른 시설을 제외"})
    dup = dedup.cross_facility_duplicates(latest)
    n_drop = int((dup["중복_제외권장"] == "Y").sum()) if len(dup) else 0
    dup_log = [{"공고번호": "", "항목": "시설 간 중복 프로젝트", "판정": "참고",
                "비고": f"프로젝트 {len(dup)}개가 다른 시설 후보의 프로젝트와 대표 공고를 함께 가짐(같은 사업이 두 번 이상 집계). "
                        f"05 시트 '중복_제외권장' = Y 인 {n_drop}개 행을 빼면 같은 공고가 두 번 합산되지 않음 — 관계는 05 '시설간_중복' 열"}] if len(dup) else []
    _logs_frame(logs + multi + dup_log).to_parquet(_p(cfg, "logs_dedup.parquet"), index=False)
    n_split = int((hist.drop_duplicates("프로젝트ID")["세부사업수"] > 1).sum()) if len(hist) else 0
    n_base = hist.loc[hist["세부사업수"] > 1, "시설ID"].astype(str).add(hist.loc[hist["세부사업수"] > 1, "사업유형"].astype(str)).nunique() if len(hist) else 0
    warn = [l for l in logs if l.get("판정") == "경고"]
    print(f"전체 {len(hist)}건 → 대표 {len(latest)}건 (고유 공고 {latest['공고키'].nunique() if len(latest) else 0}건, "
          f"프로젝트 {latest['프로젝트ID'].nunique() if len(latest) else 0}개, 경고 {len(warn)}건, 복수 시설 배정 {len(multi)}건)")
    if n_split:
        print(f"  - 세부사업 분리: 시설 후보 {n_base}개를 세부사업 {n_split}개로 나눔(프로젝트ID 끝 -01, -02 …). 06_검증로그 '세부사업 분리'·'같은 공종 합산' 에서 확인")
    from collections import Counter
    for name, n in Counter(str(l.get("항목")) for l in logs).most_common():
        print(f"  - {name}: {n}건")
    for l in warn[:10]:
        print(f"    · [{l.get('항목')}] {str(l.get('비고'))[:150]}")
    if multi:
        print(f"  - 복수 시설 배정 {len(multi)}건 (06_검증로그 '복수 시설 배정' 항목, 예: {multi[0]['비고'][:80]}…)")
    if len(dup):
        print(f"  - 시설 간 중복 프로젝트 {len(dup)}개: 05 시트 '중복_제외권장' = Y 인 {n_drop}개 행을 빼면 같은 공고가 두 번 합산되지 않습니다('시설간_중복' 열에 관계 표시)")


def _merge_attachment_slots(rep: dict, others: list, attach_max: int, spt_max: int = 5):
    """대표 공고(최신 차수)의 첨부에 같은 공고번호의 다른 차수 첨부를 URL 기준으로 이어 붙인다.
    정정·변경공고 차수에는 '정정공고서' 한 장만 붙어 있고 원 공고문·현장설명서는 첫 차수에 있는 경우가 많다."""
    row = dict(rep)
    seen, slots, spts = set(), [], []
    for r in [rep] + list(others):
        for i in range(1, attach_max + 1):
            u, nm = str(r.get(f"첨부URL{i}") or "").strip(), str(r.get(f"첨부파일명{i}") or "").strip()
            if len(u) >= 5 and u.lower() not in ("nan", "none") and u not in seen:
                seen.add(u); slots.append((u, nm))
        for i in range(1, spt_max + 1):
            u = str(r.get(f"현장설명서URL{i}") or "").strip()
            if len(u) >= 5 and u.lower() not in ("nan", "none") and u not in seen:
                seen.add(u); spts.append(u)
    n = max(attach_max, len(slots))
    for i in range(1, n + 1):
        row[f"첨부URL{i}"], row[f"첨부파일명{i}"] = (slots[i - 1] if i <= len(slots) else ("", ""))
    for i in range(1, spt_max + 1):
        row[f"현장설명서URL{i}"] = spts[i - 1] if i <= len(spts) else ""
    return row, n


def stage_attach(cfg, retry_failed: bool = False, report_only: bool = False):
    """대표 공고의 첨부를 내려받아 텍스트 추출. 텍스트가 없는 공고는 매 실행마다 다시 시도(내려받은 파일은 재사용).
    retry_failed=True 면 텍스트가 있어도 추출 실패 파일이 하나라도 있는 공고를 다시 처리(HWP 백엔드 설치·파서 개선 뒤 일괄 재시도)."""
    latest = _read(cfg, "notices_latest.parquet")
    notes, texts = [], {}
    texts_path = _p(cfg, "texts.json")
    if os.path.exists(texts_path):
        with open(texts_path, encoding="utf-8") as f:
            texts = json.load(f)
    prev_notes = _p(cfg, "notes_attach.parquet")
    prev_df = pd.read_parquet(prev_notes) if os.path.exists(prev_notes) else pd.DataFrame()
    if report_only:
        _attach_report(cfg, latest, texts, prev_df)
        return
    # 같은 공고번호의 다른 차수(원 공고·변경공고) 첨부를 합쳐서 처리 — 대표가 정정공고 차수면 원 공고문이 이전 차수에만 있다
    hist_path = _p(cfg, "notices_hist.parquet")
    others_by_no = {}
    if os.path.exists(hist_path):
        h = pd.read_parquet(hist_path)
        h = h[[c for c in h.columns if c.startswith("첨부URL") or c.startswith("첨부파일명") or c.startswith("현장설명서URL") or c in ("공고번호", "공고키", "공고차수")]]
        h = h.drop_duplicates("공고키").sort_values("공고차수", ascending=False)
        for no, g in h.groupby(h["공고번호"].astype(str)):
            if len(g) > 1:
                others_by_no[no] = [r for r in g.to_dict("records")]
    attach_max = int(cfg["fields"]["attach_max"])
    retry_nos = set()
    if retry_failed and not prev_df.empty and "추출성공" in prev_df.columns:
        err = prev_df["오류"].fillna("").astype(str).str.strip() != "" if "오류" in prev_df.columns else False
        retry_nos = set(prev_df.loc[(prev_df["추출성공"] != "Y") | err, "공고번호"].astype(str))
        # 다른 차수의 첨부가 아직 반영되지 않은 공고도 다시 — 단, 대표 텍스트가 짧거나(정정공고서뿐) 다른 차수에 공고문·현장설명서류가 있을 때만
        seen_urls = prev_df.groupby(prev_df["공고번호"].astype(str))["URL"].apply(set).to_dict() if "URL" in prev_df.columns else {}
        rep_len = {str(r["공고번호"]): len(texts.get(r["공고키"], "")) for _, r in latest.iterrows()}
        key_doc = re.compile(r"공고문|공고서|입찰공고|현장설명|설계설명|입찰안내|과업")
        n_merge = 0
        for no, rows in others_by_no.items():
            if no not in rep_len:
                continue
            unseen = []
            for r in rows:
                for i in range(1, attach_max + 1):
                    u, nm = str(r.get(f"첨부URL{i}") or "").strip(), str(r.get(f"첨부파일명{i}") or "")
                    if len(u) >= 5 and u.lower() not in ("nan", "none") and u not in seen_urls.get(no, set()):
                        unseen.append(nm)
                for i in range(1, 6):
                    u = str(r.get(f"현장설명서URL{i}") or "").strip()
                    if len(u) >= 5 and u.lower() not in ("nan", "none") and u not in seen_urls.get(no, set()):
                        unseen.append("현장설명서")
            if unseen and (rep_len.get(no, 0) < 5000 or any(key_doc.search(nm) for nm in unseen)):
                if no not in retry_nos:
                    n_merge += 1
                retry_nos.add(no)
        print(f"[--retry-failed] 추출 실패·오류 기록이 있는 공고와 다른 차수의 공고문·현장설명서가 아직 반영되지 않은 공고 {len(retry_nos)}건"
              f"(그중 차수 병합 {n_merge}건)을 다시 처리합니다(내려받은 파일은 재사용)")
    processed_nos, n_merged = set(), 0
    for i, (_, row) in enumerate(latest.iterrows(), 1):
        if row["공고키"] in texts and str(row["공고번호"]) not in retry_nos:
            continue
        processed_nos.add(str(row["공고번호"]))
        rd = row.to_dict()
        others = [r for r in others_by_no.get(str(row["공고번호"]), []) if r.get("공고키") != row["공고키"]]
        if others:
            rd, amax = _merge_attachment_slots(rd, others, attach_max)
            n_merged += 1
        else:
            amax = attach_max
        try:
            text, n = attachments.process_notice_attachments(rd, cfg["paths"]["files_dir"], cfg["paths"]["text_dir"], attach_max=amax)
        except Exception as e:  # noqa: BLE001 — 한 공고의 오류가 전체를 멈추지 않도록
            log.warning("첨부 처리 실패 %s: %s", row["공고번호"], e)
            text, n = "", [{"공고번호": str(row["공고번호"]), "파일명": "", "우선순위": None, "URL": "", "다운로드": "N",
                            "추출성공": "N", "추출글자수": 0, "파서": "", "오류": f"공고 처리 오류: {str(e)[:150]}"}]
        notes.extend(n)
        if text:
            texts[row["공고키"]] = text
        if i % 25 == 0:
            with open(texts_path, "w", encoding="utf-8") as f:
                json.dump(texts, f, ensure_ascii=False)
            log.info("첨부 처리 %d/%d", i, len(latest))
    notes_df = pd.DataFrame(notes)
    if n_merged:
        print(f"다른 차수의 첨부를 합쳐 처리한 공고 {n_merged}건")
    # 이번에 다시 처리한 공고의 옛 기록은 버리고 새 기록으로 대체(재시도할 때마다 07 추출노트가 중복되지 않도록)
    if not prev_df.empty and processed_nos and "공고번호" in prev_df.columns:
        prev_df = prev_df[~prev_df["공고번호"].astype(str).isin(processed_nos)]
    notes_df = pd.concat([d for d in (prev_df, notes_df) if not d.empty], ignore_index=True) if (not prev_df.empty or not notes_df.empty) else notes_df
    notes_df.to_parquet(prev_notes, index=False)
    with open(texts_path, "w", encoding="utf-8") as f:
        json.dump(texts, f, ensure_ascii=False)
    _attach_report(cfg, latest, texts, notes_df)


def _attach_report(cfg, latest, texts, notes_df):
    """첨부 처리 요약 + 텍스트를 못 얻은 공고의 원인 분해(첨부 없음 / 다운로드 실패 / 미지원 형식만 / 추출 실패) → data/notes_notext.parquet"""
    have = set(latest["공고키"])
    n_rows, n_uniq = len(latest), len(have)
    dup_note = f" (대표 공고 행 {n_rows}개 중 같은 공고가 여러 시설에 배정된 것 {n_rows - n_uniq}개)" if n_rows != n_uniq else ""
    print(f"첨부 처리: 고유 공고 {n_uniq}건{dup_note}, 파일 {len(notes_df)}개, 텍스트 확보 {len([k for k in texts if k in have])}/{n_uniq}건")
    if not notes_df.empty and "형식" in notes_df.columns:
        s = notes_df.groupby(notes_df["형식"].fillna("(다운로드 실패)")).agg(파일수=("파일명", "size"), 추출성공=("추출성공", lambda x: int((x == "Y").sum())))
        print("형식별 추출 결과:\n" + s.to_string())
    missing = latest[~latest["공고키"].isin(set(texts))].drop_duplicates("공고키").copy()
    if missing.empty:
        return
    by_no = {}
    if not notes_df.empty:
        for no, g in notes_df.groupby(notes_df["공고번호"].astype(str)):
            by_no[no] = g
    unsupported = {".xlsb", ".doc", ".pptx", ".7z", ".egg", ".cell", ".pme", ".dwg", ".jpg", ".png"}      # .xls 는 xlrd 로 지원

    def cause(row):
        g = by_no.get(str(row["공고번호"]))
        if int(row.get("첨부파일수") or 0) == 0 and (g is None or g.empty):
            return "첨부 없음"
        if g is None or g.empty:
            return "첨부 없음"
        if (g["다운로드"] != "Y").all():
            return "다운로드 실패(로그인 페이지 등)"
        forms = set(g.loc[g["다운로드"] == "Y", "형식"].fillna("").astype(str).str.lower())
        if forms and forms <= unsupported:
            return "미지원 형식만(" + "/".join(sorted(forms)) + ")"
        ok = g[g["추출성공"] == "Y"]
        if not ok.empty:
            chars = int(pd.to_numeric(ok.get("추출글자수"), errors="coerce").fillna(0).sum())
            return "추출은 됐으나 텍스트 없음(글자수 0 — 표지·이미지뿐)" if chars == 0 else "추출은 됐으나 텍스트가 저장되지 않음 → attach 재실행으로 복구"
        if forms == {".pdf"}:
            return "PDF 추출 실패(스캔)"
        return "추출 실패(" + "/".join(sorted(forms)[:3]) + ")"
    missing["텍스트없음_원인"] = [cause(r) for _, r in missing.iterrows()]
    missing["_amt"] = pd.to_numeric(missing["추정가격"], errors="coerce").fillna(0)
    cnt = missing["텍스트없음_원인"].value_counts()
    print(f"텍스트 없는 공고 {len(missing)}건의 원인:\n" + cnt.to_string())
    if not notes_df.empty and "오류" in notes_df.columns:
        rows = notes_df[notes_df["공고번호"].astype(str).isin(set(missing["공고번호"].astype(str)))]
        errs = rows["오류"].fillna("").astype(str).str.slice(0, 70).replace("", "(오류 메시지 없음)").value_counts().head(8)
        if len(errs):
            print("해당 공고 파일들의 오류 메시지 상위(추출 성공 기록 포함):\n" + errs.to_string())
    top = missing.sort_values("_amt", ascending=False).head(10)
    print("그중 추정가격 상위 10건:\n" + top[["공고번호", "공고명", "공종", "추정가격", "텍스트없음_원인"]].to_string(index=False))
    out = missing.drop(columns=["_amt"])[[c for c in ("공고키", "공고번호", "공고명", "시설명", "공종", "추정가격", "첨부파일수", "상세URL", "텍스트없음_원인") if c in missing.columns]]
    out.to_parquet(_p(cfg, "notes_notext.parquet"), index=False)
    print(f"→ 목록 저장: {_p(cfg, 'notes_notext.parquet')} (상세URL 로 나라장터에서 직접 확인 가능)")


def _redo_keys(latest: pd.DataFrame, docs: dict) -> list:
    """다시 추출할 공고: 신뢰도 low, 또는 건축 공종인데 연면적이 비어 있는 추출 결과(원 공고문이 빠졌던 경우 등)."""
    out = []
    trade = dict(zip(latest["공고키"], latest["공종"].astype(str))) if "공종" in latest.columns else {}
    for k, d in docs.items():
        if not isinstance(d, dict) or "_error" in d:
            continue
        low = str(d.get("신뢰도") or "").lower() == "low"
        no_area = d.get("연면적_m2") in (None, "", 0) and trade.get(k) == "건축"
        if low or no_area:
            out.append(k)
    return out


def _extract_context(cfg, limit: Optional[int] = None, redo_low: bool = False):
    latest = _read(cfg, "notices_latest.parquet")
    texts_path = _p(cfg, "texts.json")
    if not os.path.exists(texts_path):
        raise SystemExit("texts.json 이 없습니다 (attach 를 먼저 실행)")
    with open(texts_path, encoding="utf-8") as f:
        texts = json.load(f)
    cache_path = _p(cfg, "llm_docs.json")
    docs = {}
    if os.path.exists(cache_path):
        with open(cache_path, encoding="utf-8") as f:
            docs = json.load(f)
    # 같은 공고가 여러 시설의 대표로 잡혀 latest 에 두 번 나올 수 있다 → 공고키 기준으로 한 번만 추출(비용 중복 방지)
    todo = list(dict.fromkeys(k for k in latest["공고키"] if k in texts and not (k in docs and "_error" not in docs[k])))
    redo = []
    if redo_low:
        redo = [k for k in _redo_keys(latest, docs) if k in texts]
        todo = list(dict.fromkeys(redo + todo))          # 다시 할 것을 맨 앞에 → --limit 과 함께 쓰면 그것부터
        print(f"(--redo-low) 신뢰도 low 또는 건축인데 연면적이 없는 기존 추출 {len(redo)}건을 다시 대상에 포함(맨 앞)")
    if limit:
        # 시범 추출용: 건축 공종을 먼저, 그 안에서 추정가격이 큰 순(연면적·규모가 있는 본공사 공고문이 앞에 오도록)
        order = latest.assign(_amt=pd.to_numeric(latest["추정가격"], errors="coerce").fillna(0),
                              _arch=(latest["공종"].astype(str) == "건축").astype(int) if "공종" in latest.columns else 0)
        rank = {k: i for i, k in enumerate(order.sort_values(["_arch", "_amt"], ascending=[False, False])["공고키"])}
        redo_set = set(redo)
        todo = sorted(todo, key=lambda k: (0 if k in redo_set else 1, rank.get(k, 10**9)))[:int(limit)]
        print(f"(--limit {limit}) {'다시 할 것 → ' if redo else ''}건축 공종·추정가격 큰 순으로 {len(todo)}건만 대상")
    hints = {r["공고키"]: {"공고번호": r["공고번호"], "공고명": r["공고명"], "수요기관": r["수요기관"]} for _, r in latest.iterrows()}
    return latest, texts, docs, todo, hints, cache_path


def _save_docs(cache_path, docs):
    with open(cache_path, "w", encoding="utf-8") as f:
        json.dump(docs, f, ensure_ascii=False, default=str)


def _error_summary(docs: dict, top: int = 8) -> str:
    """추출 실패(_error) 원인을 앞 70자 기준으로 묶어 상위 top 개를 표로."""
    from collections import Counter
    errs = Counter()
    for d in docs.values():
        if isinstance(d, dict) and "_error" in d:
            msg = str(d["_error"])
            msg = re.sub(r"\| 응답: .*$", "", msg).strip()          # 응답 앞부분(공고마다 다름)은 묶을 때 제외
            errs[msg[:70]] += 1
    if not errs:
        return ""
    lines = [f"  {n:5d}  {m}" for m, n in errs.most_common(top)]
    return f"추출 실패 {sum(errs.values())}건의 원인 상위:\n" + "\n".join(lines)


def _verify_logs(cfg, latest, docs) -> List[Dict]:
    """1차 추출 검증로그를 현재 규칙·현재 대표 공고로 만들어 logs_verify.parquet 에 저장(수거 직후와 excel 단계 양쪽에서 호출, 재과금 없음).
    excel 때 다시 만들지 않으면 규칙이 바뀌거나(부가세 '참고' 규칙) 대표 공고가 바뀐 뒤에도 옛 로그가 06 시트에 그대로 실린다."""
    logs: List[Dict] = []
    seen = set()
    for _, row in latest.iterrows():
        k = row["공고키"]
        d = docs.get(k, {})
        if isinstance(d, dict) and d and "_error" not in d and k not in seen:
            seen.add(k)                       # 복수 시설 배정 공고는 한 번만
            logs.extend(extract_llm.cross_verify(row["공고번호"], row.to_dict(), d,
                                                 cfg["verify"]["price_diff_warn_ratio"], cfg["verify"]["vat_ratio"]))
    _logs_frame(logs).to_parquet(_p(cfg, "logs_verify.parquet"), index=False)
    return logs


def _pending_batch_keys(cfg, state_name: str) -> set:
    """아직 수거하지 않은 배치에 들어 있는 공고키. 즉시 호출(--yes)이 같은 공고를 한 번 더 부르지 않게(이중 과금 방지) 대상에서 뺀다."""
    path = _p(cfg, state_name)
    if not os.path.exists(path):
        return set()
    try:
        state = json.load(open(path, encoding="utf-8"))
    except (OSError, ValueError):
        return set()
    return {k for b in state.get("batches", []) if b.get("status") != "ended" for k in b.get("keys", [])}


def _skip_pending(cfg, todo: List[str], state_name: str, cmd: str) -> List[str]:
    pend = _pending_batch_keys(cfg, state_name)
    hit = [k for k in todo if k in pend]
    if hit:
        print(f"대기 중인 배치에 들어 있는 공고 {len(hit)}건은 이번 대상에서 제외합니다 — 먼저 수거:  python -m g2b_costdb.pipeline {cmd} --batch")
    return [k for k in todo if k not in pend]


def _verify_and_report(cfg, latest, docs, n_ok=None, n_err=None):
    logs = _verify_logs(cfg, latest, docs)
    done = len([k for k in latest["공고키"] if k in docs and "_error" not in docs[k]])
    usage_in = sum((d.get("_usage_in") or 0) for d in docs.values() if isinstance(d, dict))
    usage_out = sum((d.get("_usage_out") or 0) for d in docs.values() if isinstance(d, dict))
    es = _error_summary(docs)
    if es:
        print(es)
    head = f"LLM 추출 성공 {n_ok}건 / 실패 {n_err}건, " if n_ok is not None else ""
    print(f"{head}대표 공고 {len(latest)}건 중 추출 완료 {done}건 (누적 토큰 입력 {usage_in:,} 출력 {usage_out:,}), 검증로그 {len(logs)}건")
    if logs:
        print(pd.DataFrame(logs)["판정"].value_counts().to_string())


def _truncated_retry_tokens(cfg, docs: dict, todo: List[str]) -> Dict[str, int]:
    """이전 시도가 max_tokens 에서 잘린 공고 → 이번엔 설정값의 2배로. (사고 토큰이 한도를 다 써 JSON 이 나오기 전에 끊긴 경우)"""
    base = int(cfg["llm"].get("max_tokens", extract_llm.DEFAULT_MAX_TOKENS))
    out = {k: base * 2 for k in todo if "stop=max_tokens" in str((docs.get(k) or {}).get("_error", ""))}
    if out:
        print(f"이전에 max_tokens({base:,})에서 잘린 {len(out)}건은 max_tokens {base * 2:,} 로 다시 요청합니다")
    return out


def _print_estimate(cfg, texts, todo, docs, latest, batch=False):
    est = extract_llm.estimate_cost(texts, todo, cfg["llm"], float(cfg["llm"].get("usd_krw", 1400)))
    if batch:
        est["usd"], est["krw"] = round(est["usd"] / 2, 2), est["krw"] // 2
    es = _error_summary(docs)
    if es:
        print(es + "\n  → 실패 건은 다음 extract 실행 때 자동으로 다시 대상에 포함됩니다")
    print(f"LLM 추출 대상: {len(todo)}건 (이미 완료 {len([k for k in docs if '_error' not in docs[k]])}건, 텍스트 없음 "
          f"{len([k for k in latest['공고키'] if k not in texts])}건)")
    print(f"예상 입력 {est['input_tokens']:,} 토큰 / 출력 {est['output_tokens']:,} 토큰, 모델 {est['model']}"
          f"{' (Batches 50% 할인 적용)' if batch else ''} → 약 ${est['usd']} (≈{est['krw']:,}원, 환율 {cfg['llm'].get('usd_krw', 1400)})")
    return est


def stage_extract(cfg, yes: bool = False, batch: bool = False, export: bool = False, import_: bool = False, limit: Optional[int] = None,
                  redo_low: bool = False, refetch: bool = False):
    """LLM 추출. 세 가지 방식:
      (기본)   extract            → 견적만 / extract --yes → API 로 한 건씩 호출
      --batch  extract --batch    → 대기 중인 배치가 있으면 결과 수거, 없으면 견적 / --batch --yes → Message Batches 제출(50% 할인, 최대 24h)
      --export extract --export   → data/llm_in/<공고키>.txt 로 내보내기(Cowork·사람이 채움) → extract --import 로 data/llm_out/*.json 수거"""
    latest, texts, docs, todo, hints, cache_path = _extract_context(cfg, limit=limit, redo_low=redo_low)
    in_dir, out_dir = _p(cfg, "llm_in"), _p(cfg, "llm_out")

    if export:
        n = extract_llm.export_prompts(todo, texts, hints, cfg["llm"], in_dir, out_dir)
        print(f"내보내기 완료: {n}건 → {os.path.join(in_dir, '<공고키>.txt')}  (지시문: README_지시문.md)")
        print(f"채운 JSON 은 {os.path.join(out_dir, '<공고키>.json')} 으로 저장한 뒤  python -m g2b_costdb.pipeline extract --import")
        return

    if import_:
        found, errors = extract_llm.import_results(out_dir, set(latest["공고키"]))
        done_dir = _p(cfg, "llm_done")
        os.makedirs(done_dir, exist_ok=True)
        for k in found:                      # 반영한 작업지시(txt)·결과(json)는 llm_done 으로 옮겨 llm_in/llm_out 에는 남은 일만 보이게
            for d_, ext in ((in_dir, ".txt"), (out_dir, ".json")):
                src = os.path.join(d_, k + ext)
                if os.path.exists(src):
                    shutil.move(src, os.path.join(done_dir, k + ext))
        kept = []
        for k, d in found.items():
            old_doc = docs.get(k)
            # 이미 API 결과가 있는 공고에 사람이 쓴 옛 JSON 이 다시 들어오면 덮어쓰지 않는다(예: 시범 40건 JSON 이 llm_out 에 남아 있다가 재반영)
            if isinstance(old_doc, dict) and "_error" not in old_doc and not str(old_doc.get("_model", "")).startswith("manual") \
                    and str(d.get("_model", "")).startswith("manual"):
                kept.append(k)
                continue
            docs[k] = d
        _save_docs(cache_path, docs)
        print(f"가져오기: JSON {len(found) - len(kept)}건 반영(파일은 {done_dir} 로 이동), 문제 파일 {len(errors)}건"
              + (f", 이미 API 결과가 있어 유지한 공고 {len(kept)}건" if kept else ""))
        for fn, why in errors[:20]:
            print(f"  - {fn}: {why}")
        remaining = [k for k in latest["공고키"] if k in texts and not (k in docs and "_error" not in docs[k])]
        print(f"아직 미추출 {len(remaining)}건" + (f" (예: {', '.join(remaining[:5])})" if remaining else ""))
        _verify_and_report(cfg, latest, docs)
        return

    env = cfg["llm"].get("api_key_env", "ANTHROPIC_API_KEY")
    if batch:
        state_path = _p(cfg, "llm_batch.json")
        state = json.load(open(state_path, encoding="utf-8")) if os.path.exists(state_path) else {"batches": []}
        if refetch:
            # 끝난 배치의 결과를 다시 내려받아(결과는 29일 보관, 재과금 없음) 실패로 남은 공고만 다시 해석 — 파서를 고친 뒤 복구용
            if not os.environ.get(env, "").strip():
                raise SystemExit(f"환경변수 {env} 가 없습니다.")
            import anthropic  # type: ignore
            client = anthropic.Anthropic(api_key=os.environ.get(env) or None)
            n_fix = n_still = 0
            for b in state["batches"]:
                status, got, counts = extract_llm.fetch_batch(client, b["id"], cfg["llm"].get("model", ""))
                if status != "ended":
                    print(f"배치 {b['id']}: {status} — 아직 끝나지 않음")
                    continue
                for k in b["keys"]:
                    if k in docs and "_error" not in docs[k]:
                        continue
                    d = got.get(k, {"_error": "배치 결과에 없음"})
                    docs[k] = d
                    n_fix += 0 if "_error" in d else 1
                    n_still += 1 if "_error" in d else 0
            _save_docs(cache_path, docs)
            print(f"[--refetch] 실패 건 다시 해석: 복구 {n_fix}건, 여전히 실패 {n_still}건")
            _verify_and_report(cfg, latest, docs)
            return
        pending = [b for b in state["batches"] if b.get("status") != "ended"]
        if pending:
            if not os.environ.get(env, "").strip():
                raise SystemExit(f"환경변수 {env} 가 없습니다.")
            import anthropic  # type: ignore
            client = anthropic.Anthropic(api_key=os.environ.get(env) or None)
            n_ok = n_err = 0
            for b in pending:
                status, got, counts = extract_llm.fetch_batch(client, b["id"], cfg["llm"].get("model", ""))
                if status != "ended":
                    print(f"배치 {b['id']}: {status} (처리중 {counts.get('processing', 0)}, 성공 {counts.get('succeeded', 0)}, 오류 {counts.get('errored', 0)}) — 나중에 같은 명령으로 다시 확인")
                    continue
                for k in b["keys"]:
                    d = got.get(k, {"_error": "배치 결과에 없음"})
                    docs[k] = d
                    n_ok += 0 if "_error" in d else 1
                    n_err += 1 if "_error" in d else 0
                b["status"] = "ended"
                print(f"배치 {b['id']}: 완료 (성공 {counts.get('succeeded', 0)}, 오류 {counts.get('errored', 0)}, 만료 {counts.get('expired', 0)})")
            _save_docs(cache_path, docs)
            with open(state_path, "w", encoding="utf-8") as f:
                json.dump(state, f, ensure_ascii=False, indent=1)
            if any(b.get("status") != "ended" for b in state["batches"]):
                return
            _verify_and_report(cfg, latest, docs, n_ok, n_err)
            return
        _print_estimate(cfg, texts, todo, docs, latest, batch=True)
        if not todo:
            print("제출할 공고가 없습니다.")
            return
        if not yes:
            print("배치 제출을 실행하려면:  python -m g2b_costdb.pipeline extract --batch --yes   (결과 수거도 같은 명령 extract --batch)")
            return
        if not os.environ.get(env, "").strip():
            raise SystemExit(f"환경변수 {env} 가 없습니다. PowerShell 에서 setx {env} \"키\" 로 설정한 뒤 새 터미널에서 다시 실행하세요.")
        import anthropic  # type: ignore
        client = anthropic.Anthropic(api_key=os.environ.get(env) or None)
        retry_mt = _truncated_retry_tokens(cfg, docs, todo)

        def _keep(b):                                    # 청크(500건)마다 바로 기록 → 뒤 청크가 실패해도 앞 배치를 잃지 않는다(이중 과금 방지)
            state["batches"].append(b)
            with open(state_path, "w", encoding="utf-8") as f:
                json.dump(state, f, ensure_ascii=False, indent=1)
        try:
            submitted = extract_llm.submit_batches(client, [(k, texts[k], hints.get(k, {})) for k in todo], cfg["llm"],
                                                   max_tokens_by_key=retry_mt, on_submitted=_keep)
        except (anthropic.AuthenticationError, anthropic.PermissionDeniedError) as e:
            raise SystemExit(f"[중단] Claude API 인증 실패: {e}. {env} 값을 확인하세요.")
        except Exception as e:  # noqa: BLE001
            kept = getattr(e, "submitted", None) or []
            raise SystemExit(f"[중단] 배치 제출 실패: {e}. 이미 제출된 배치 {len(kept)}개는 기록했으니 extract --batch 로 수거한 뒤 다시 --yes 하세요.")
        print(f"배치 {len(submitted)}개 제출 (공고 {len(todo)}건). 대개 1시간 안에, 늦어도 24시간 안에 끝납니다.")
        print("결과 수거:  python -m g2b_costdb.pipeline extract --batch")
        return

    todo = _skip_pending(cfg, todo, "llm_batch.json", "extract")
    _print_estimate(cfg, texts, todo, docs, latest)
    if not yes:
        print("실제 추출을 실행하려면:  python -m g2b_costdb.pipeline extract --yes   (50% 저렴한 배치: extract --batch --yes / 파일 인수인계: extract --export)")
        return
    if not os.environ.get(env, "").strip():
        raise SystemExit(f"환경변수 {env} 가 없습니다. PowerShell 에서 setx {env} \"키\" 로 설정한 뒤 새 터미널에서 다시 실행하세요.")
    import anthropic  # type: ignore
    n_ok = n_err = 0
    retry_mt = _truncated_retry_tokens(cfg, docs, todo)
    for k in todo:
        try:
            docs[k] = extract_llm.extract_with_claude(texts[k], hints.get(k, {}), cfg["llm"], max_tokens=retry_mt.get(k))
            n_ok += 1
        except (anthropic.AuthenticationError, anthropic.PermissionDeniedError) as e:
            raise SystemExit(f"[중단] Claude API 인증 실패: {e}. {env} 값을 확인하세요.")
        except anthropic.RateLimitError as e:
            log.warning("속도 제한(429)으로 중단 — 잠시 후 extract --yes 를 다시 실행하면 이어서 진행: %s", e)
            break
        except Exception as e:  # noqa: BLE001
            log.warning("LLM 추출 실패 %s: %s", k, e)
            docs[k] = {"_error": str(e)[:300]}
            n_err += 1
        _save_docs(cache_path, docs)
    _verify_and_report(cfg, latest, docs, n_ok, n_err)


# ── 2차 추출: 내역서 집계표의 공종별 금액(건축 공고·통합발주 공고만) ─────────────────────────────
TRADE_INPUT_CHARS = 20000
# 본문(공고문·현장설명서)만으로 2차 대상에 넣는 단서: 공종별 금액 표가 있을 법한 표현만('원가계산에 의한 예정가격' 같은 상투어는 제외)
_TRADE_TEXT_HINT = re.compile(r"공종별\s*(금액|공사비|내역|추정가격|예정가격)|집계표|총괄\s*(내역|집계)|원가계산서"
                              r"|(건축|토목|기계|조경|전기|정보통신|통신|소방|철거)[^\n]{0,20}\d{1,3}(,\d{3}){2,}")
_TEXT_BLOCK = re.compile(r"^===== \[(.+?)\] =====$", re.M)


def _non_excel_text(text: str) -> str:
    """texts.json 의 공고 텍스트에서 Excel 첨부 블록(===== [파일.xlsx] =====)을 뺀 나머지(공고문·현장설명서 등)만.
    Excel 내역서는 cost_sheet_snippets 가 파일에서 직접 추리므로 본문 창에 중복으로 들어가지 않게 한다."""
    if not text:
        return ""
    parts = _TEXT_BLOCK.split(text)          # [앞머리, 이름1, 본문1, 이름2, 본문2, …]
    out = [parts[0]]
    for name, body in zip(parts[1::2], parts[2::2]):
        if not re.search(r"\.xls[xm]?$", name.strip(), re.I):
            out.append(body)
    return "\n".join(out)
DETAIL_COLS = ["프로젝트ID", "시설ID", "시설명", "사업유형", "상위공종", "공고번호", "공고차수", "공고키", "세부공종", "항목명_원문",
               "금액_원_문서", "배분적용", "금액기준", "관급자재_포함", "단위", "집계표합계_원_문서", "출처파일", "신뢰도", "근거문구"]


def _trade_targets(latest: pd.DataFrame, docs: dict, texts: dict) -> List[str]:
    """2차 추출 대상 공고키: 공종이 건축이거나, 포함공종(API∪문서)이 2개 이상인 공고(중복 제거)."""
    keys: List[str] = []
    for _, r in latest.iterrows():
        k = r["공고키"]
        if k in keys:
            continue
        inc_api = classify.included_trades(r.get("주공종명"), r.get("부공종명"), r["공종"])
        d = docs.get(k) or {}
        if "_error" in d:
            d = {}
        inc_doc = classify.normalize_trade_list(d.get("포함공종_문서") or "") or \
            classify.included_trades_from_text(texts.get(k, ""), assigned=str(r["공종"]))
        inc = classify.merge_trade_lists(inc_api, inc_doc, assigned=str(r["공종"]))
        if str(r["공종"]) == "건축" or len([p for p in inc.split("·") if p]) >= 2:
            keys.append(k)
    return keys


def _trade_input(cfg, bid_no: str, text: str, max_chars: int) -> Tuple[str, List[str]]:
    """2차 추출 입력: 이미 내려받은 내역서 Excel 의 집계표 후보 행(파일당 최대 12,000자) + 공고문·현장설명서의 공사금액·집계 키워드 창.
    내역서 파일이 없고 본문에 공종별·집계 류 단서도 없으면 빈 문자열(대상에서 제외)."""
    files = attachments.find_cost_sheet_files(cfg["paths"]["files_dir"], str(bid_no))
    parts, used, budget = [], [], int(max_chars)
    for f in files:
        snip = attachments.cost_sheet_snippets(f, max_chars=min(12000, budget))
        if snip:
            parts.append(snip)
            used.append(os.path.basename(f))
            budget -= len(snip)
        if budget <= 2000:
            break
    body = _non_excel_text(text or "")
    if parts or _TRADE_TEXT_HINT.search(body):
        win = extract_llm.keyword_windows(body, extract_llm._TRADE_KEYS, budget=min(8000, max(0, budget)))
        if win:
            parts.append("## [공고문·현장설명서 발췌]\n" + win)
    return "\n\n".join(parts), used


def _trade_context(cfg, limit: Optional[int] = None):
    latest = _read(cfg, "notices_latest.parquet")
    texts_path = _p(cfg, "texts.json")
    if not os.path.exists(texts_path):
        raise SystemExit("texts.json 이 없습니다 (attach 를 먼저 실행)")
    texts = json.load(open(texts_path, encoding="utf-8"))
    dpath = _p(cfg, "llm_docs.json")
    docs = json.load(open(dpath, encoding="utf-8")) if os.path.exists(dpath) else {}
    cache_path = _p(cfg, "llm_trades.json")
    tdocs = json.load(open(cache_path, encoding="utf-8")) if os.path.exists(cache_path) else {}
    targets = _trade_targets(latest, docs, texts)
    todo = [k for k in targets if not (k in tdocs and "_error" not in tdocs[k])]
    max_chars = int(cfg["llm"].get("trades_max_input_chars", TRADE_INPUT_CHARS))
    bid_of = dict(zip(latest["공고키"], latest["공고번호"].astype(str)))
    inputs, skipped = {}, []
    for n, k in enumerate(todo, 1):
        if len(todo) > 200 and n % 200 == 0:
            print(f"  집계표 발췌 중 … {n}/{len(todo)}")
        try:
            snip, used = _trade_input(cfg, bid_of.get(k, ""), texts.get(k, ""), max_chars)
        except BaseException as e:  # noqa: BLE001 — 손상 파일(러스트 panic 포함) 하나 때문에 전체가 멈추지 않게
            if not attachments._catchable(e):
                raise
            log.warning("2차 입력 준비 실패 %s: %s", k, str(e)[:120])
            snip, used = "", []
        if snip:
            inputs[k] = (snip, used)
        else:
            skipped.append(k)
            tdocs[k] = {"_error": "건너뜀: 집계표 없음(첨부에 내역서 Excel 없고 본문에 공종별 금액 표 단서 없음)"}
    todo = [k for k in todo if k in inputs]
    if limit:
        amt = dict(zip(latest["공고키"], pd.to_numeric(latest["추정가격"], errors="coerce").fillna(0)))
        todo = sorted(todo, key=lambda k: -amt.get(k, 0))[:int(limit)]
        print(f"(--limit {limit}) 추정가격 큰 순으로 {len(todo)}건만 대상")
    hints = {r["공고키"]: {"공고번호": r["공고번호"], "공고명": r["공고명"], "수요기관": r["수요기관"], "공종": r["공종"]}
             for _, r in latest.iterrows()}
    done = len([k for k in targets if k in tdocs and "_error" not in tdocs[k]])
    n_xl = sum(1 for k in todo if inputs[k][1])
    print(f"2차 추출 대상 {len(targets)}건(건축 공고 + 통합발주 공고, 고유 공고키) 중 완료 {done}건, 이번 대상 {len(todo)}건"
          f"(내역서 Excel 있음 {n_xl}건 / 본문 표만 {len(todo) - n_xl}건), 집계표 없음 {len(skipped)}건(제외)")
    return latest, texts, tdocs, todo, inputs, hints, cache_path, max_chars


def _trade_dup_logs(latest: pd.DataFrame, tdocs: dict) -> List[Dict]:
    """같은 프로젝트에 전기·정보통신·소방 공고가 따로 있는데 건축 공고 집계표에도 그 공종 금액이 있으면 '중복 가능' 경고."""
    logs: List[Dict] = []
    sep = latest[latest["공종"].astype(str) != "건축"].groupby("프로젝트ID")["공종"].agg(lambda s: set(map(str, s))).to_dict()
    for _, r in latest.iterrows():
        t = tdocs.get(r["공고키"])
        if not t or "_error" in t or str(r["공종"]) != "건축":
            continue
        have = sep.get(r["프로젝트ID"], set())
        for it in t.get("공종별금액") or []:
            if it["공종"] in ("전기", "정보통신", "소방") and it["공종"] in have:
                logs.append({"공고번호": r["공고번호"], "항목": "공종 중복 가능(문서)", "문서값": it["금액_원"], "판정": "경고",
                             "비고": f"{it['공종']}: 분리발주 공고가 따로 있는데 건축 집계표에도 금액 있음 → 05 '문서배분_중복경고' 참고, 합산 금지"})
    return logs


def _trade_logs(cfg, latest, tdocs) -> List[Dict]:
    """2차 결과의 검증로그를 현재 규칙으로 다시 만들어 logs_trades.parquet 에 저장(재과금 없음). 수거 직후와 excel 단계 양쪽에서 부른다 —
    excel 이 옛 수거 때 저장한 로그를 그대로 싣지 않도록."""
    logs: List[Dict] = []
    seen = set()
    tdocs = _renormalize_trades(tdocs)
    for _, row in latest.iterrows():
        t = tdocs.get(row["공고키"])
        if not t or "_error" in t or row["공고키"] in seen:
            continue
        seen.add(row["공고키"])
        logs.extend(extract_llm.cross_verify_trades(row["공고번호"], row.to_dict(), t, cfg.get("verify", {}).get("vat_ratio", 0.10)))
    logs.extend(_trade_dup_logs(latest, tdocs))
    _logs_frame(logs).to_parquet(_p(cfg, "logs_trades.parquet"), index=False)
    return logs


def _verify_trades_and_report(cfg, latest, tdocs, n_ok=None, n_err=None):
    logs = _trade_logs(cfg, latest, tdocs)
    done = len({k for k in latest["공고키"] if k in tdocs and "_error" not in tdocs[k]})
    with_items = len({k for k in latest["공고키"] if k in tdocs and "_error" not in tdocs[k] and (tdocs[k].get("공종별금액") or [])})
    usage_in = sum((d.get("_usage_in") or 0) for d in tdocs.values() if isinstance(d, dict))
    usage_out = sum((d.get("_usage_out") or 0) for d in tdocs.values() if isinstance(d, dict))
    is_skip = lambda d: isinstance(d, dict) and str(d.get("_error", "")).startswith("건너뜀")
    n_skip = sum(1 for d in tdocs.values() if is_skip(d))
    es = _error_summary({k: d for k, d in tdocs.items() if not is_skip(d)})      # 건너뜀(집계표 없음)은 실패가 아니므로 따로 센다
    if es:
        print(es)
    head = f"2차 추출 성공 {n_ok}건 / 실패 {n_err}건, " if n_ok is not None else ""
    print(f"{head}완료 {done}건(공종별 금액이 나온 공고 {with_items}건, 누적 토큰 입력 {usage_in:,} 출력 {usage_out:,}), "
          f"집계표 없음(건너뜀) {n_skip}건, 검증로그 {len(logs)}건")
    if logs:
        print(pd.DataFrame(logs)["판정"].value_counts().to_string())


def stage_extract_trades(cfg, yes: bool = False, batch: bool = False, limit: Optional[int] = None, refetch: bool = False):
    """2차 추출: 건축 공고(와 통합발주 공고)의 내역서 집계표에서 공종별 금액을 읽어 data/llm_trades.json 에 저장 → excel 의 04b 시트.
    실행 방식은 extract 와 같다(견적 → --yes 즉시 / --batch --yes 제출 → --batch 수거 / --refetch). 1차 결과는 건드리지 않는다."""
    latest, texts, tdocs, todo, inputs, hints, cache_path, max_chars = _trade_context(cfg, limit=limit)
    env = cfg["llm"].get("api_key_env", "ANTHROPIC_API_KEY")
    model = cfg["llm"].get("model", "")

    def _estimate(is_batch):
        est = extract_llm.estimate_cost({k: inputs[k][0] for k in todo}, todo, cfg["llm"], float(cfg["llm"].get("usd_krw", 1400)),
                                        max_chars=max_chars, out_per_doc=1500)
        if is_batch:
            est["usd"], est["krw"] = round(est["usd"] / 2, 2), est["krw"] // 2
        print(f"2차 추출 예상 입력 {est['input_tokens']:,} 토큰 / 출력 {est['output_tokens']:,} 토큰, 모델 {est['model']}"
              f"{' (Batches 50% 할인 적용)' if is_batch else ''} → 약 ${est['usd']} (≈{est['krw']:,}원)")
        return est

    def _client():
        if not os.environ.get(env, "").strip():
            raise SystemExit(f"환경변수 {env} 가 없습니다.")
        import anthropic  # type: ignore
        return anthropic.Anthropic(api_key=os.environ.get(env) or None)

    def _attach_files(k, d, files_map=None):
        """04b 출처파일: 제출 때 상태 파일에 적어 둔 파일명(files_map) 우선, 없으면 이번 실행의 입력."""
        if isinstance(d, dict) and "_error" not in d:
            d["_files"] = (files_map or {}).get(k) or inputs.get(k, ("", []))[1] or d.get("_files") or []
        return d

    if batch:
        state_path = _p(cfg, "llm_trades_batch.json")
        state = json.load(open(state_path, encoding="utf-8")) if os.path.exists(state_path) else {"batches": []}
        if refetch:
            client = _client()
            n_fix = n_still = 0
            for b in state["batches"]:
                status, got, counts = extract_llm.fetch_batch(client, b["id"], model, parser=extract_llm.parse_trade_response_text)
                if status != "ended":
                    print(f"배치 {b['id']}: {status} — 아직 끝나지 않음")
                    continue
                for k in b["keys"]:
                    if k in tdocs and "_error" not in tdocs[k]:
                        continue
                    d = got.get(k, {"_error": "배치 결과에 없음"})
                    tdocs[k] = _attach_files(k, d, b.get("files"))
                    n_fix += 0 if "_error" in d else 1
                    n_still += 1 if "_error" in d else 0
            _save_docs(cache_path, tdocs)
            print(f"[--refetch] 2차 추출 실패 건 다시 해석: 복구 {n_fix}건, 여전히 실패 {n_still}건")
            _verify_trades_and_report(cfg, latest, tdocs)
            return
        pending = [b for b in state["batches"] if b.get("status") != "ended"]
        if pending:
            client = _client()
            n_ok = n_err = 0
            for b in pending:
                status, got, counts = extract_llm.fetch_batch(client, b["id"], model, parser=extract_llm.parse_trade_response_text)
                if status != "ended":
                    print(f"2차 배치 {b['id']}: {status} (처리중 {counts.get('processing', 0)}, 성공 {counts.get('succeeded', 0)}, "
                          f"오류 {counts.get('errored', 0)}) — 나중에 같은 명령으로 다시 확인")
                    continue
                for k in b["keys"]:
                    d = got.get(k, {"_error": "배치 결과에 없음"})
                    tdocs[k] = _attach_files(k, d, b.get("files"))
                    n_ok += 0 if "_error" in d else 1
                    n_err += 1 if "_error" in d else 0
                b["status"] = "ended"
                print(f"2차 배치 {b['id']}: 완료 (성공 {counts.get('succeeded', 0)}, 오류 {counts.get('errored', 0)}, 만료 {counts.get('expired', 0)})")
            _save_docs(cache_path, tdocs)
            with open(state_path, "w", encoding="utf-8") as f:
                json.dump(state, f, ensure_ascii=False, indent=1)
            if any(b.get("status") != "ended" for b in state["batches"]):
                return
            _verify_trades_and_report(cfg, latest, tdocs, n_ok, n_err)
            new_todo = [k for k in todo if not (k in tdocs and "_error" not in tdocs[k])]
            if new_todo:
                print(f"[안내] 아직 제출되지 않은 2차 대상 {len(new_todo)}건(실패 재대상 또는 1차 결과로 새로 대상이 된 공고) → extract-trades --batch --yes 로 제출")
            return
        _estimate(True)
        _save_docs(cache_path, tdocs)                    # 건너뜀 표시 저장
        if not todo:
            print("2차 추출: 제출할 공고가 없습니다.")
            return
        if not yes:
            print("2차 배치 제출을 실행하려면:  python -m g2b_costdb.pipeline extract-trades --batch --yes   (또는 extract --batch --yes --trades)")
            return
        client = _client()
        retry_mt = _truncated_retry_tokens(cfg, tdocs, todo)

        def _keep(b):                                    # 청크마다 바로 기록(배치 ID + 공고별 출처파일) → 뒤 청크 실패 시에도 수거 가능, 이중 과금 방지
            b["kind"] = "trades"
            b["files"] = {k: inputs[k][1] for k in b["keys"] if k in inputs}
            state["batches"].append(b)
            with open(state_path, "w", encoding="utf-8") as f:
                json.dump(state, f, ensure_ascii=False, indent=1)
        try:
            submitted = extract_llm.submit_batches(client, [(k, inputs[k][0], hints.get(k, {})) for k in todo], cfg["llm"],
                                                   max_tokens_by_key=retry_mt, params_builder=extract_llm.build_trade_request_params,
                                                   on_submitted=_keep)
        except Exception as e:  # noqa: BLE001
            kept = getattr(e, "submitted", None) or []
            raise SystemExit(f"[중단] 2차 배치 제출 실패: {e}. 이미 제출된 배치 {len(kept)}개는 기록했으니 extract --batch --trades 로 수거한 뒤 다시 --yes 하세요.")
        print(f"2차 배치 {len(submitted)}개 제출 (공고 {len(todo)}건). 결과 수거:  python -m g2b_costdb.pipeline extract-trades --batch   (또는 extract --batch --trades)")
        return

    todo[:] = _skip_pending(cfg, todo, "llm_trades_batch.json", "extract-trades")
    _estimate(False)
    _save_docs(cache_path, tdocs)
    if not todo:
        print("2차 추출: 대상이 없습니다.")
        return
    if not yes:
        print("2차 추출을 실행하려면:  python -m g2b_costdb.pipeline extract-trades --yes   (배치: extract-trades --batch --yes)")
        return
    _client()
    import anthropic  # type: ignore
    n_ok = n_err = 0
    retry_mt = _truncated_retry_tokens(cfg, tdocs, todo)
    for k in todo:
        try:
            tdocs[k] = _attach_files(k, extract_llm.extract_trades_with_claude(inputs[k][0], hints.get(k, {}), cfg["llm"], max_tokens=retry_mt.get(k)))
            n_ok += 1
        except (anthropic.AuthenticationError, anthropic.PermissionDeniedError) as e:
            raise SystemExit(f"[중단] Claude API 인증 실패: {e}. {env} 값을 확인하세요.")
        except anthropic.RateLimitError as e:
            log.warning("속도 제한(429)으로 중단 — 잠시 후 다시 실행하면 이어서 진행: %s", e)
            break
        except Exception as e:  # noqa: BLE001
            log.warning("2차 추출 실패 %s: %s", k, e)
            tdocs[k] = {"_error": str(e)[:300]}
            n_err += 1
        _save_docs(cache_path, tdocs)
    _verify_trades_and_report(cfg, latest, tdocs, n_ok, n_err)


# ── 보강 추출(3차): 기타금액·면세 문구·공사비 정보광장 공사개요 항목 → data/llm_extra.json (1차 결과는 건드리지 않음) ─────────────────
def _extra_logs(cfg, latest, docs: dict, xdocs: dict) -> List[Dict]:
    """보강 추출 검증로그를 현재 규칙으로 만들어 logs_extra.parquet 에 저장(수거 직후와 excel 단계 양쪽에서 호출, 재과금 없음)."""
    logs: List[Dict] = []
    seen = set()
    xdocs = _renormalize_extra(xdocs)
    for _, row in latest.iterrows():
        k = row["공고키"]
        x = xdocs.get(k)
        if not isinstance(x, dict) or "_error" in x or k in seen:
            continue
        seen.add(k)
        d = docs.get(k) or {}
        logs.extend(extract_llm.cross_verify_extra(row["공고번호"], row.to_dict(), {} if "_error" in d else d, x))
    for (pid, k), keep in _etc_duplicates(latest, docs, xdocs).items():
        no = str(latest.loc[latest["공고키"] == k, "공고번호"].iloc[0])
        logs.append({"공고번호": no, "항목": "기타금액 중복 가능", "문서값": extract_llm._num(xdocs[k].get("기타금액_원")), "판정": "참고",
                     "비고": f"{pid}: 같은 프로젝트의 공고 {keep} 와 같은 기타금액 → 이 공고의 금액은 합계에서 제외(이중계산 방지). 서로 다른 항목이면 04 '기타금액_판정' 확인"})
    _logs_frame(logs).to_parquet(_p(cfg, "logs_extra.parquet"), index=False)
    return logs


def _extra_report(cfg, latest, docs, xdocs, n_ok=None, n_err=None):
    logs = _extra_logs(cfg, latest, docs, xdocs)
    ok = {k: x for k, x in xdocs.items() if isinstance(x, dict) and "_error" not in x and k in set(latest["공고키"])}
    cnt = lambda f: sum(1 for x in ok.values() if x.get(f) not in (None, ""))
    es = _error_summary(xdocs)
    if es:
        print(es)
    head = f"보강 추출 성공 {n_ok}건 / 실패 {n_err}건, " if n_ok is not None else ""
    print(f"{head}완료 {len(ok)}건 — 대지면적 {cnt('대지면적_m2')}, 건폐율 {cnt('건폐율_pct')}, 용적률 {cnt('용적률_pct')}, 최고높이 {cnt('최고높이_m')}, "
          f"주차대수 {cnt('주차대수')}, 설계특이사항 {cnt('설계특이사항')}, 기타금액 {cnt('기타금액_원')}, 면세 문구 {cnt('면세_문구')} (검증로그 {len(logs)}건)")


def stage_extract_extra(cfg, yes: bool = False, batch: bool = False, limit: Optional[int] = None, refetch: bool = False,
                        only: Optional[str] = None):
    """보강 추출: 대표 공고의 공고문에서 1차 추출에 없는 항목(기타금액[한전불입금 등]·면세 문구·조경면적·건폐율·용적률·기준층층고·최고높이·주차대수·세대수·
    건축법상용도·설계특이사항, 그리고 1차에서 빈 대지면적·건축면적)을 읽어 data/llm_extra.json 에 저장 → excel 의 01·04 시트.
    실행 방식은 extract 와 같다(견적 → --yes 즉시 / --batch --yes 제출 → --batch 수거 / --refetch). --only 전기,건축 처럼 공종을 한정할 수 있다."""
    latest = _read(cfg, "notices_latest.parquet")
    texts_path = _p(cfg, "texts.json")
    if not os.path.exists(texts_path):
        raise SystemExit("texts.json 이 없습니다 (attach 를 먼저 실행)")
    texts = json.load(open(texts_path, encoding="utf-8"))
    docs_path, cache_path = _p(cfg, "llm_docs.json"), _p(cfg, "llm_extra.json")
    docs = json.load(open(docs_path, encoding="utf-8")) if os.path.exists(docs_path) else {}
    xdocs = json.load(open(cache_path, encoding="utf-8")) if os.path.exists(cache_path) else {}
    env, model = cfg["llm"].get("api_key_env", "ANTHROPIC_API_KEY"), cfg["llm"].get("model", "")
    max_chars = int(cfg["llm"].get("extra_max_input_chars", extract_llm.EXTRA_INPUT_CHARS))
    alias = {"통신": "정보통신", "기계": "기계설비", "설비": "기계설비"}
    trades = list(dict.fromkeys(alias.get(t.strip(), t.strip()) for t in re.split(r"[,·ㆍ/\s]+", str(only or "")) if t.strip()))
    known = ["건축", "전기", "정보통신", "소방", "기계설비", "토목", "조경", "기타"]
    bad = [t for t in trades if t not in known]
    if bad:
        raise SystemExit(f"--only 의 공종 이름을 알 수 없습니다: {', '.join(bad)}. 쓸 수 있는 이름: {', '.join(known)} (쉼표로 여러 개, 예: --only 전기,건축)")
    pool = latest[latest["공종"].astype(str).isin(trades)] if trades else latest
    # 건축 공종을 먼저, 그 안에서 추정가격 큰 순(--limit 시범 추출이 공사개요가 있는 본공사부터 되도록)
    order = pool.assign(_amt=pd.to_numeric(pool["추정가격"], errors="coerce").fillna(0), _arch=(pool["공종"].astype(str) == "건축").astype(int))
    keys = list(dict.fromkeys(order.sort_values(["_arch", "_amt"], ascending=[False, False])["공고키"]))
    todo = [k for k in keys if texts.get(k) and not (k in xdocs and "_error" not in xdocs[k])]
    if not batch:                                     # 즉시 호출: 수거 전 배치에 든 공고는 빼고(이중 과금 방지) 그다음에 --limit 을 적용
        todo = _skip_pending(cfg, todo, "llm_extra_batch.json", "extract-extra")
    if limit:
        todo = todo[:int(limit)]
    hints = {r["공고키"]: {"공고번호": r["공고번호"], "공고명": r["공고명"], "수요기관": r["수요기관"], "공종": r["공종"]} for _, r in latest.iterrows()}
    done = len([k for k in keys if k in xdocs and "_error" not in xdocs[k]])
    print(f"보강 추출 대상 {len(keys)}건({'공종 ' + '·'.join(trades) if trades else '대표 공고 전체'}, 고유 공고키) 중 완료 {done}건, 이번 대상 {len(todo)}건"
          f"{f' (--limit {limit})' if limit else ''}, 텍스트 없음 {len([k for k in keys if not texts.get(k)])}건")

    def _estimate(is_batch):
        est = extract_llm.estimate_cost(texts, todo, cfg["llm"], float(cfg["llm"].get("usd_krw", 1400)), max_chars=max_chars, out_per_doc=1800)
        if is_batch:
            est["usd"], est["krw"] = round(est["usd"] / 2, 2), est["krw"] // 2
        print(f"보강 추출 예상 입력 {est['input_tokens']:,} 토큰 / 출력 {est['output_tokens']:,} 토큰, 모델 {est['model']}"
              f"{' (Batches 50% 할인 적용)' if is_batch else ''} → 약 ${est['usd']} (≈{est['krw']:,}원)")

    def _client():
        if not os.environ.get(env, "").strip():
            raise SystemExit(f"환경변수 {env} 가 없습니다. PowerShell 에서 setx {env} \"키\" 로 설정한 뒤 새 터미널에서 다시 실행하세요.")
        import anthropic  # type: ignore
        return anthropic.Anthropic(api_key=os.environ.get(env) or None)

    parser = extract_llm.parse_extra_response_text
    if batch:
        state_path = _p(cfg, "llm_extra_batch.json")
        state = json.load(open(state_path, encoding="utf-8")) if os.path.exists(state_path) else {"batches": []}
        pending = [b for b in state["batches"] if b.get("status") != "ended"]
        if refetch or pending:
            client = _client()
            n_ok = n_err = 0
            for b in (state["batches"] if refetch else pending):
                status, got, counts = extract_llm.fetch_batch(client, b["id"], model, parser=parser)
                if status != "ended":
                    print(f"보강 배치 {b['id']}: {status} (처리중 {counts.get('processing', 0)}, 성공 {counts.get('succeeded', 0)}, "
                          f"오류 {counts.get('errored', 0)}) — 나중에 같은 명령으로 다시 확인")
                    continue
                for k in b["keys"]:
                    if k in xdocs and "_error" not in xdocs[k]:
                        continue                              # 이미 성공한 결과는 덮어쓰지 않는다(--refetch 는 실패 건만 다시 해석)
                    d = got.get(k, {"_error": "배치 결과에 없음"})
                    xdocs[k] = d
                    n_ok += 0 if "_error" in d else 1
                    n_err += 1 if "_error" in d else 0
                if not refetch:
                    b["status"] = "ended"
                    print(f"보강 배치 {b['id']}: 완료 (성공 {counts.get('succeeded', 0)}, 오류 {counts.get('errored', 0)}, 만료 {counts.get('expired', 0)})")
            _save_docs(cache_path, xdocs)
            with open(state_path, "w", encoding="utf-8") as f:
                json.dump(state, f, ensure_ascii=False, indent=1)
            if not refetch and any(b.get("status") != "ended" for b in state["batches"]):
                return
            _extra_report(cfg, latest, docs, xdocs, n_ok, n_err)
            return
        _estimate(True)
        if not todo:
            print("보강 추출: 제출할 공고가 없습니다.")
            return
        if not yes:
            print("보강 배치 제출을 실행하려면:  python -m g2b_costdb.pipeline extract-extra --batch --yes   (수거도 같은 명령 extract-extra --batch)")
            return
        client = _client()
        retry_mt = _truncated_retry_tokens(cfg, xdocs, todo)

        def _keep(b):                                    # 청크마다 바로 기록 → 뒤 청크가 실패해도 앞 배치를 잃지 않는다(이중 과금 방지)
            b["kind"] = "extra"
            state["batches"].append(b)
            with open(state_path, "w", encoding="utf-8") as f:
                json.dump(state, f, ensure_ascii=False, indent=1)
        try:
            submitted = extract_llm.submit_batches(client, [(k, texts[k], hints.get(k, {})) for k in todo], cfg["llm"], max_tokens_by_key=retry_mt,
                                                   params_builder=extract_llm.build_extra_request_params, on_submitted=_keep)
        except Exception as e:  # noqa: BLE001
            kept = getattr(e, "submitted", None) or []
            raise SystemExit(f"[중단] 보강 배치 제출 실패: {e}. 이미 제출된 배치 {len(kept)}개는 기록했으니 extract-extra --batch 로 수거한 뒤 다시 --yes 하세요.")
        print(f"보강 배치 {len(submitted)}개 제출 (공고 {len(todo)}건). 결과 수거:  python -m g2b_costdb.pipeline extract-extra --batch")
        return

    _estimate(False)
    if not todo:
        print("보강 추출: 대상이 없습니다.")
        return
    if not yes:
        print("보강 추출을 실행하려면:  python -m g2b_costdb.pipeline extract-extra --yes   (50% 저렴한 배치: extract-extra --batch --yes, 시범: --limit 30)")
        return
    _client()
    import anthropic  # type: ignore
    n_ok = n_err = 0
    retry_mt = _truncated_retry_tokens(cfg, xdocs, todo)
    for k in todo:
        try:
            xdocs[k] = extract_llm.extract_extra_with_claude(texts[k], hints.get(k, {}), cfg["llm"], max_tokens=retry_mt.get(k))
            n_ok += 1
        except (anthropic.AuthenticationError, anthropic.PermissionDeniedError) as e:
            raise SystemExit(f"[중단] Claude API 인증 실패: {e}. {env} 값을 확인하세요.")
        except anthropic.RateLimitError as e:
            log.warning("속도 제한(429)으로 중단 — 잠시 후 다시 실행하면 이어서 진행: %s", e)
            break
        except Exception as e:  # noqa: BLE001
            log.warning("보강 추출 실패 %s: %s", k, e)
            xdocs[k] = {"_error": str(e)[:300]}
            n_err += 1
        _save_docs(cache_path, xdocs)
    _extra_report(cfg, latest, docs, xdocs, n_ok, n_err)


def _renormalize_extra(xdocs: dict) -> dict:
    """저장된 보강 추출 결과에 현재 정규화 규칙을 다시 적용(재과금 없음): 단위가 붙어 비워졌던 숫자('_raw')를 다시 해석하고 위치 표기를 정리한다."""
    out = {}
    for k, x in (xdocs or {}).items():
        out[k] = extract_llm.normalize_extra_doc(x) if (isinstance(x, dict) and "_error" not in x) else x
    return out


def _renormalize_trades(tdocs: dict) -> dict:
    """저장된 2차 결과에 현재 라벨 규칙을 다시 적용(재과금 없음): '도급비(건축, 토목, 기계)' 같은 합산 행은 '합계' 로, 라벨 밖 공종명은 매핑."""
    out = {}
    for k, t in tdocs.items():
        if isinstance(t, dict) and "_error" not in t:
            keep = {kk: vv for kk, vv in t.items() if kk.startswith("_")}
            t = extract_llm.normalize_trade_doc(t)
            t.update({kk: vv for kk, vv in keep.items() if kk != "_공종합_원"})
        out[k] = t
    return out


def _alloc_decision(api_row: Dict, t: Dict) -> str:
    """04b 비율 배분을 적용할지('Y' / 'N(사유)'). 공종 행 합이 집계표(합계 − 간접비·관급·부가세·배분제외)의 30~125%(집계표 합계가 없으면 API 기초금액의 30~160%)여야 한다.
    모델이 일부 행만 읽은 문서(건축 행 누락, 기타 행만 남음)는 작은 행 하나가 100% 가 되어 05 문서배분을 왜곡하므로 막고,
    집계표 합계가 API 금액과 맞지 않는 문서(차수·다른 공사 집계표 의심)도 배분하지 않는다."""
    items = t.get("공종별금액") or []
    alloc = [i["금액_원"] for i in items if i["공종"] in extract_llm.TRADE_ALLOC_LABELS]
    if not alloc:
        return "N(공종 행 없음)"
    asum = sum(alloc)
    if asum <= 0:
        return "N(공종 행 금액 0)"
    total = extract_llm._num(t.get("합계_원"))
    # 집계표 합계(없으면 읽은 행들의 합)가 API 추정가격·기초금액(관급·부가세 보정 포함)과 맞지 않으면 차수·다른 공사 집계표 또는 일부 행만 읽은 문서 → 배분하지 않음
    basis = next((l for l in extract_llm.cross_verify_trades(str(api_row.get("공고번호", "")), api_row, t) if l["항목"] == "집계표합계 기준"), None)
    if basis and basis["판정"] == "경고":
        return "N(집계표 합계가 API 금액과 불일치)" if total else "N(읽은 행의 합이 API 금액과 불일치)"
    if total:
        ref = total - sum(i["금액_원"] for i in items if i["공종"] in ("공통가설", "간접비", "관급자재", "부가세", "배분제외"))
        if ref <= 0:
            ref = total
        src, hi = "집계표 공종분", 1.25
    else:
        ref = extract_llm._num(api_row.get("기초금액")) or (extract_llm._num(api_row.get("추정가격")) or 0) * 1.1
        if not ref:
            return "Y"
        src, hi = "API 기초금액", 1.6
    cov = asum / ref
    if cov < 0.3:
        return f"N(공종 행 합이 {src}의 {cov:.0%} — 일부 행만 읽힘)"
    if cov > hi:
        return f"N(공종 행 합이 {src}의 {cov:.0%} — 중복·다른 집계표 의심)"
    return "Y"


def assemble_trade_detail_table(latest: pd.DataFrame, tdocs: dict) -> pd.DataFrame:
    """04b 행: 대표 공고 × 집계표 항목(공종별 금액). 복수 시설 배정 공고는 04 와 같이 프로젝트마다 나타난다."""
    rows = []
    tdocs = _renormalize_trades(tdocs)
    decisions: Dict[str, str] = {}
    for _, r in latest.iterrows():
        t = tdocs.get(r["공고키"])
        if not t or "_error" in t:
            continue
        if r["공고키"] not in decisions:
            decisions[r["공고키"]] = _alloc_decision(r.to_dict(), t)
        for it in t.get("공종별금액") or []:
            rows.append({"프로젝트ID": r["프로젝트ID"], "시설ID": r["시설ID"], "시설명": r["시설명"], "사업유형": r["사업유형"],
                         "상위공종": r["공종"], "공고번호": r["공고번호"], "공고차수": r["공고차수"], "공고키": r["공고키"],
                         "세부공종": it.get("공종"), "항목명_원문": it.get("항목명_원문"), "금액_원_문서": it.get("금액_원"), "배분적용": decisions[r["공고키"]],
                         "금액기준": t.get("금액기준"), "관급자재_포함": t.get("관급자재_포함여부"), "단위": t.get("단위"),
                         "집계표합계_원_문서": t.get("합계_원"), "출처파일": ", ".join(t.get("_files") or []) or "공고문·현장설명서(본문)",
                         "신뢰도": t.get("신뢰도"), "근거문구": t.get("근거문구")})
    return pd.DataFrame(rows, columns=DETAIL_COLS)


def _etc_decision(etc, where: str, api_row: Dict, doc: Dict) -> str:
    """기타금액(한전불입금·분담금 등)을 금액 합계에 더할지. 문서가 도급·관급과 '별도' 로 적었으면 가산, 관급·추정가격에 포함이라 적었거나
    관급 금액과 같은 값이면(발주기관이 한전불입금을 관급 칸에 넣은 사례) 제외 — 이중계산 방지.
    추정가격·기초금액·부가세·총공사금액·예산과 같은 값이거나 기초금액의 30% 를 넘는 값은 다른 금액을 잘못 읽은 것으로 보고 더하지 않는다(확인 필요)."""
    num = extract_llm._num
    v = num(etc)
    if not v:
        return ""
    same = lambda g: bool(g) and abs(g - v) <= max(1.0, v * 0.001)
    gov = [num(api_row.get(k)) for k in ("도급자관급액_API", "관급자관급액_API", "관급자재_API")] + \
          [num(doc.get(k)) for k in ("도급자관급액_원", "관급자관급액_원")]
    if any(same(g) for g in gov):
        return "제외(관급 금액과 같은 값 — 이중계산 방지)"
    main = {"추정가격": [num(api_row.get("추정가격")), num(doc.get("추정가격_원"))], "기초금액": [num(api_row.get("기초금액")), num(doc.get("기초금액_원"))],
            "부가세": [num(api_row.get("부가세_API")), num(doc.get("부가가치세_원"))], "총공사금액": [num(doc.get("총공사금액_원")), num(doc.get("설계금액_원"))],
            "예산금액": [num(api_row.get("예산금액"))]}
    for name, vals in main.items():
        if any(same(g) for g in vals):
            return f"제외(확인 필요: {name}과 같은 값 — 오독 의심)"
    base = num(api_row.get("기초금액")) or num(doc.get("기초금액_원")) or (num(api_row.get("추정가격")) or num(doc.get("추정가격_원")) or 0) * 1.1
    if base and v > base * extract_llm.ETC_MAX_SHARE:
        return f"제외(확인 필요: 기초금액의 {extract_llm.ETC_MAX_SHARE:.0%} 초과 — 오독 의심)"
    w = extract_llm._etc_where(where)                 # 표기 변형('관급 포함', '도급액에 포함', '도급액 별도')을 뜻에 맞게 정리
    if w == "별도":
        return "가산(문서: 별도 금액)"
    if w == "관급에포함":
        return "제외(문서: 관급에 포함)"
    if w == "추정가격에포함":
        return "제외(문서: 추정가격에 포함)"
    return "가산(위치 불명 — 확인 필요)"


def _etc_duplicates(latest: pd.DataFrame, docs: dict, extra: dict) -> Dict[Tuple[str, str], str]:
    """한 프로젝트의 여러 공고(예: 건축 공고문의 사업비 표와 전기 공고문)가 같은 기타금액을 읽은 경우 하나만 더한다.
    반환 {(프로젝트ID, 제외할 공고키): 남기는 공고키}. 남기는 공고 = 항목에 '한전'이 있으면 전기 공고, 아니면 건축 공고, 그다음 목록 순서."""
    out: Dict[Tuple[str, str], str] = {}
    if latest is None or latest.empty or not extra:
        return out
    for pid, g in latest.groupby("프로젝트ID"):
        cand = []
        for _, r in g.drop_duplicates("공고키").iterrows():
            k = r["공고키"]
            x, d = extra.get(k) or {}, docs.get(k) or {}
            if not isinstance(x, dict) or "_error" in x:
                continue
            d = {} if (not isinstance(d, dict) or "_error" in d) else d
            v = extract_llm._num(x.get("기타금액_원"))
            if v and _etc_decision(v, x.get("기타금액_위치"), r.to_dict(), d).startswith("가산"):
                cand.append((k, v, str(r["공종"]), str(x.get("기타금액_항목") or "")))
        used = set()
        for i, (k, v, trade, item) in enumerate(cand):
            if k in used:
                continue
            grp = [c for c in cand[i:] if c[0] not in used and abs(c[1] - v) <= max(1.0, v * 0.001)]
            used.update(c[0] for c in grp)
            if len(grp) < 2:
                continue
            want = "전기" if any("한전" in c[3] for c in grp) else "건축"
            keep = next((c for c in grp if c[2] == want), None) or next((c for c in grp if c[2] == "건축"), None) or grp[0]
            for c in grp:
                if c[0] != keep[0]:
                    out[(str(pid), c[0])] = keep[0]
    return out


def assemble_trade_table(latest: pd.DataFrame, docs: dict, texts: Optional[dict] = None, extra: Optional[dict] = None) -> pd.DataFrame:
    """04 시트 행. 포함공종: API(주공종명·부공종명) / 문서(LLM 필드 '포함공종_문서', 없으면 텍스트 규칙) / 종합(합집합).
    extra = 보강 추출(llm_extra.json): 기타금액·면세 문구."""
    rows = []
    texts, extra = texts or {}, extra or {}
    dup = _etc_duplicates(latest, docs, extra)
    for _, r in latest.iterrows():
        d = docs.get(r["공고키"], {}) or {}
        if not isinstance(d, dict) or "_error" in d:
            d = {}
        x = extra.get(r["공고키"], {}) or {}
        if not isinstance(x, dict) or "_error" in x:
            x = {}
        raw_ev, raw_xev = d.get("근거문구"), x.get("근거문구")       # 옛 추출 결과는 근거문구가 문자열·배열일 수 있다 → 그대로 '원문' 으로
        ev = dict(raw_ev) if isinstance(raw_ev, dict) else ({"원문": raw_ev} if raw_ev else {})
        ev.update({f"[보강]{k}": v for k, v in (raw_xev.items() if isinstance(raw_xev, dict) else ([("원문", raw_xev)] if raw_xev else []))})
        inc_api = classify.included_trades(r.get("주공종명"), r.get("부공종명"), r["공종"])
        inc_doc, inc_src = classify.normalize_trade_list(d.get("포함공종_문서") or ""), "LLM"
        if not inc_doc:
            inc_doc, inc_src = classify.included_trades_from_text(texts.get(r["공고키"], ""), assigned=str(r["공종"])), "규칙"
        if not inc_doc:
            inc_src = ""
        rows.append({
            "프로젝트ID": r["프로젝트ID"], "시설ID": r["시설ID"], "시설명": r["시설명"], "세부사업명": r.get("세부사업명") or "", "사업유형": r["사업유형"],
            "공종": r["공종"],
            "공종_비고": classify.trade_note(r["공종"], r.get("공종근거"), r["공고명"], r.get("주공종명"), r.get("부공종명")),
            "포함공종_API": inc_api, "포함공종_문서": inc_doc, "포함공종_출처": inc_src,
            "포함공종(종합)": classify.merge_trade_lists(inc_api, inc_doc, assigned=str(r["공종"])),
            "분리발주_언급공종_문서": d.get("분리발주_언급공종") or "",
            "공고번호": r["공고번호"], "공고차수": r["공고차수"], "공고키": r["공고키"],
            "공고명": r["공고명"], "공고일시": r["공고일시"], "수요기관": r["수요기관"],
            "추정가격_API": r.get("추정가격"), "기초금액_API": r.get("기초금액"), "부가세_API": r.get("부가세_API"),
            "관급자재_API": r.get("관급자재_API"), "도급자관급액_API": r.get("도급자관급액_API"), "관급자관급액_API": r.get("관급자관급액_API"),
            "예산금액_API": r.get("예산금액"),
            "낙찰금액_API": r.get("낙찰금액_API"), "낙찰률_API": r.get("낙찰률_API"), "낙찰자": r.get("낙찰자"), "낙찰하한율": r.get("낙찰하한율"),
            "부가세_문서": d.get("부가가치세_원"),
            "도급자관급액_문서": d.get("도급자관급액_원"), "관급자관급액_문서": d.get("관급자관급액_원"),
            "기타금액_문서": x.get("기타금액_원"), "기타금액_항목_문서": x.get("기타금액_항목") or "",
            "기타금액_판정": (f"제외(같은 프로젝트의 공고 {dup[(str(r['프로젝트ID']), r['공고키'])]} 와 같은 금액 — 이중계산 방지)"
                              if (str(r["프로젝트ID"]), r["공고키"]) in dup
                              else _etc_decision(x.get("기타금액_원"), x.get("기타금액_위치"), r.to_dict(), d)),
            "면세_문구_문서": x.get("면세_문구") or "",
            "공사기간_일_문서": d.get("공사기간_일"), "추정가격_문서": d.get("추정가격_원"), "기초금액_문서": d.get("기초금액_원"),
            "신뢰도": d.get("신뢰도"), "근거문구": json.dumps(ev, ensure_ascii=False) if ev else "",
            "출처파일": d.get("_source_files", ""), "상세URL": r.get("상세URL"),
        })
    return pd.DataFrame(rows)


FAC_BASE_COLS = ["연면적_m2", "건축면적_m2", "대지면적_m2", "지하층수", "지상층수", "구조", "용도", "공사기간_일"]
# 보강 추출(llm_extra.json)로 채우는 공사비 정보광장 공사개요 항목: 01 열 이름 → 보강 추출 필드 이름
FAC_EXTRA_COLS = {"조경면적_m2": "조경면적_m2", "건폐율_pct": "건폐율_pct", "용적률_pct": "용적률_pct", "기준층층고_m": "기준층층고_m",
                  "최고높이_m": "최고높이_m", "주차대수": "주차대수", "세대수": "세대수", "건축법상용도_문서": "건축법상용도", "설계특이사항": "설계특이사항"}


def _empty(v) -> bool:
    return v is None or v == "" or (isinstance(v, float) and v != v)


# 0 이 '값 없음' 인 열(면적·비율·높이·대수·기간). 지하층수·지상층수는 0 이 실제 값일 수 있어 제외
_ZERO_EMPTY_COLS = {"연면적_m2", "건축면적_m2", "대지면적_m2", "조경면적_m2", "건폐율_pct", "용적률_pct", "기준층층고_m", "최고높이_m",
                    "주차대수", "세대수", "공사기간_일"}


def _blank(c: str, v) -> bool:
    """01 공사개요 칸을 비어 있다고 볼지: None·''·NaN, 그리고 면적·대수 열의 0(1차 추출이 0 으로 적은 값이 보강 값을 막지 않게)."""
    if _empty(v):
        return True
    return c in _ZERO_EMPTY_COLS and extract_llm._num(v) == 0


def enrich_facility(fac: pd.DataFrame, latest: pd.DataFrame, docs: dict, extra: Optional[dict] = None) -> pd.DataFrame:
    """프로젝트(시설×사업유형) 공사개요를 조달청 '공사비 정보광장' 항목 틀로 채운다. 값은 건축 공종 문서 → 없으면 다른 공종 문서 순,
    같은 공고 안에서는 1차 추출(llm_docs) → 보강 추출(llm_extra) 순. 어느 공종 공고에서 왔는지 '개요_출처' 에 적는다.
    공사명 = 건축 대표 공고명(추정가격 최대), 현장위치 = 문서의 공사위치(없으면 API 공사현장지역), 발주년월 = 건축 대표 공고의 최초 공고월."""
    fac = fac.copy()
    extra = extra or {}
    cols = FAC_BASE_COLS + list(FAC_EXTRA_COLS)
    for c in cols + ["공사명", "현장위치", "발주년월", "발주년월_기준", "개요_출처"]:
        fac[c] = None
    by_pid = {pid: g for pid, g in latest.groupby("프로젝트ID")} if len(latest) else {}
    for i, f in fac.iterrows():
        g = by_pid.get(f["프로젝트ID"])
        if g is None or g.empty:
            continue
        g = g.assign(_pri=g["공종"].map(lambda t: 0 if t == "건축" else 1), _amt=pd.to_numeric(g["추정가격"], errors="coerce").fillna(0))
        g = g.sort_values(["_pri", "_amt"], ascending=[True, False])
        has_arch = bool((g["공종"] == "건축").any())
        sources: List[str] = []
        for _, r in g.iterrows():
            d = docs.get(r["공고키"], {}) or {}
            x = extra.get(r["공고키"], {}) or {}
            d = {} if "_error" in d else d
            x = {} if "_error" in x else x
            used = False
            for c in cols:
                if not _blank(c, fac.at[i, c]):
                    continue
                v = d.get(c) if c in FAC_BASE_COLS else None
                if _blank(c, v):
                    v = x.get(FAC_EXTRA_COLS.get(c, c))
                if not _blank(c, v):
                    fac.at[i, c] = v
                    used = True
            if _empty(fac.at[i, "현장위치"]) and not _empty(d.get("공사위치")):
                fac.at[i, "현장위치"] = d.get("공사위치")
                used = True
            if used and str(r["공종"]) not in sources:
                sources.append(str(r["공종"]))
        top = g.iloc[0]
        fac.at[i, "공사명"] = top["공고명"]
        if _empty(fac.at[i, "현장위치"]):
            fac.at[i, "현장위치"] = f.get("공사현장지역") if "공사현장지역" in fac.columns else None
        base = g[g["공종"] == "건축"] if has_arch else g
        dts = sorted(str(v)[:7] for v in base["공고일시"] if str(v)[:4].isdigit())
        if dts:
            fac.at[i, "발주년월"] = dts[0]
            fac.at[i, "발주년월_기준"] = "건축 공고" if has_arch else f"{top['공종']} 공고(건축 공고 없음)"
        if sources:
            fac.at[i, "개요_출처"] = "·".join(sources) + ("" if has_arch else " (건축 공고 없음)")
    return fac


def stage_excel(cfg):
    hist, latest = _read(cfg, "notices_hist.parquet"), _read(cfg, "notices_latest.parquet")
    docs_path = _p(cfg, "llm_docs.json")
    docs = json.load(open(docs_path, encoding="utf-8")) if os.path.exists(docs_path) else {}
    tpath = _p(cfg, "llm_trades.json")
    tdocs = json.load(open(tpath, encoding="utf-8")) if os.path.exists(tpath) else {}
    if docs:
        _verify_logs(cfg, latest, docs)        # 1차 검증로그도 현재 규칙·현재 대표 공고 기준으로 새로 만든다(규칙 변경·대표 교체 뒤 옛 로그가 남지 않게)
    if tdocs:
        _trade_logs(cfg, latest, tdocs)        # 2차 검증로그는 저장된 결과에 현재 규칙을 다시 적용해 새로 만든다(수거 때 로그가 옛 규칙일 수 있음)
    xpath = _p(cfg, "llm_extra.json")
    extra = _renormalize_extra(json.load(open(xpath, encoding="utf-8")) if os.path.exists(xpath) else {})
    if extra:
        _extra_logs(cfg, latest, docs, extra)  # 보강 추출 검증로그도 현재 규칙으로 다시 만든다
    log_names = ("logs_dedup.parquet", "logs_verify.parquet") + (("logs_trades.parquet",) if tdocs else ()) \
        + (("logs_extra.parquet",) if extra else ())   # 결과 파일이 없으면 남아 있는 옛 로그도 싣지 않는다
    frames = [pd.read_parquet(_p(cfg, n)) for n in log_names if os.path.exists(_p(cfg, n))]
    frames = [f for f in frames if not f.empty]
    logs = pd.concat(frames, ignore_index=True) if frames else pd.DataFrame()
    notes = pd.read_parquet(_p(cfg, "notes_attach.parquet")) if os.path.exists(_p(cfg, "notes_attach.parquet")) else pd.DataFrame()
    fac = enrich_facility(dedup.facility_summary(latest, hist), latest, docs, extra)
    texts_path = _p(cfg, "texts.json")
    texts = json.load(open(texts_path, encoding="utf-8")) if os.path.exists(texts_path) else {}
    trade = assemble_trade_table(latest, docs, texts, extra)
    detail = assemble_trade_detail_table(latest, tdocs)
    os.makedirs(cfg["paths"]["out_dir"], exist_ok=True)
    out = os.path.join(cfg["paths"]["out_dir"], cfg["paths"]["excel_name"])
    try:
        build_workbook(out, fac, hist, latest, trade, logs, notes, meta={"period": f"{cfg['period']['start']}~{cfg['period']['end']}"},
                       trade_detail=detail)
    except PermissionError:
        raise SystemExit(_excel_busy(out))
    print(f"Excel DB 저장: {out} (프로젝트 {len(fac)}, 대표공고 {len(latest)}, 이력 {len(hist)}, 검증로그 {len(logs)}, 추출노트 {len(notes)}, "
          f"공종세부내역(04b) {len(detail)}행)")
    if not logs.empty and "판정" in logs.columns:
        print("검증로그 판정별 건수:\n" + logs["판정"].value_counts().to_string())


def main(argv=None):
    for stream in (sys.stdout, sys.stderr):  # Windows cp949 콘솔에서 특수문자로 죽지 않도록
        try:
            stream.reconfigure(errors="replace")
        except (AttributeError, ValueError):
            pass
    ap = argparse.ArgumentParser(description="나라장터 공사비 DB 파이프라인 (g2b_costdb 폴더에서 실행)")
    ap.add_argument("stage", choices=["doctor", "probe", "collect", "discover", "research", "dedup", "attach", "extract", "extract-trades", "extract-extra", "excel"])
    ap.add_argument("--config", default=None)
    ap.add_argument("--ym", default="2026-08", help="probe 대상 월(YYYY-MM)")
    ap.add_argument("--fresh", action="store_true", help="discover: 이전 검수 파일의 시설ID 만 이어받고 검수 칸은 새 기본값으로(검수 시작 전 규칙이 바뀌었을 때)")
    ap.add_argument("--retry-failed", dest="retry_failed", action="store_true", help="attach: 추출 실패 파일이 있는 공고를 텍스트가 있어도 다시 처리(HWP 백엔드 설치 뒤 일괄 재시도)")
    ap.add_argument("--report", action="store_true", help="attach: 처리 없이 첨부 결과 요약과 텍스트 없는 공고의 원인만 출력")
    ap.add_argument("--limit", type=int, default=None, help="extract: 건축 공종·추정가격 큰 순으로 N건만(시범 추출용)")
    ap.add_argument("--redo-low", dest="redo_low", action="store_true", help="extract: 신뢰도 low 또는 건축인데 연면적이 없는 기존 추출 결과를 다시 대상에 포함")
    ap.add_argument("--refetch", action="store_true", help="extract --batch: 끝난 배치 결과를 다시 내려받아 실패 건만 다시 해석(재과금 없음)")
    ap.add_argument("--yes", action="store_true", help="extract: 비용 견적 확인 후 실제 LLM 호출(또는 배치 제출) 실행")
    ap.add_argument("--batch", action="store_true", help="extract: Message Batches 로 제출/수거 (50%% 할인, 최대 24시간)")
    ap.add_argument("--export", action="store_true", help="extract: 공고별 작업지시 txt 를 data/llm_in 에 내보내기 (Cowork·사람이 채움)")
    ap.add_argument("--import", dest="import_", action="store_true", help="extract: data/llm_out 의 JSON 을 가져와 반영")
    ap.add_argument("--servc", action="store_true", help="probe/collect: 용역 입찰공고(폐기물처리 용역·건설사업관리·설계용역) 시험 호출·수집")
    ap.add_argument("--only", default=None, help="extract-extra: 공종을 한정(쉼표 구분, 예: 전기,건축). 생략하면 대표 공고 전체")
    ap.add_argument("--trades", action="store_true", help="extract 와 함께 2차 추출(내역서 집계표의 공종별 금액)도 같은 방식(--yes/--batch/--limit)으로 실행")
    a = ap.parse_args(argv)
    logging.basicConfig(level=logging.INFO, format="%(asctime)s %(levelname)s %(name)s: %(message)s")
    cfg = load_config(a.config)
    for k in ("data_dir", "raw_dir", "text_dir", "files_dir", "out_dir"):
        os.makedirs(cfg["paths"][k], exist_ok=True)
    if a.stage == "doctor":
        stage_doctor(cfg)
    elif a.stage == "probe":
        try:
            collect.probe_servc(cfg, a.ym) if a.servc else collect.probe(cfg, a.ym)
        except DailyBudgetExceeded as e:
            print(f"[중단] 오늘 API 호출 예산을 다 썼습니다({e}). 내일 같은 명령으로 다시 실행하세요.")
        except RuntimeError as e:
            print(f"[중단] API 에 연결하지 못했습니다: {str(e)[:300]}\n  네트워크 상태를 확인하고 잠시 뒤 다시 실행하세요(python -m g2b_costdb.pipeline doctor 로 점검).")
    elif a.stage == "collect" and a.servc:
        collect.collect_servc(cfg)
    elif a.stage == "extract":
        stage_extract(cfg, yes=a.yes, batch=a.batch, export=a.export, import_=a.import_, limit=a.limit, redo_low=a.redo_low, refetch=a.refetch)
        if a.trades and not (a.export or a.import_):
            print("\n── 2차 추출: 내역서 집계표의 공종별 금액 ──")
            stage_extract_trades(cfg, yes=a.yes, batch=a.batch, limit=a.limit, refetch=a.refetch)
    elif a.stage == "extract-trades":
        stage_extract_trades(cfg, yes=a.yes, batch=a.batch, limit=a.limit, refetch=a.refetch)
    elif a.stage == "extract-extra":
        stage_extract_extra(cfg, yes=a.yes, batch=a.batch, limit=a.limit, refetch=a.refetch, only=a.only)
    elif a.stage == "discover":
        stage_discover(cfg, fresh=a.fresh)
    elif a.stage == "attach":
        stage_attach(cfg, retry_failed=a.retry_failed, report_only=a.report)
    else:
        {"collect": stage_collect, "discover": stage_discover, "research": stage_research, "dedup": stage_dedup,
         "attach": stage_attach, "excel": stage_excel}[a.stage](cfg)


if __name__ == "__main__":
    sys.exit(main())
