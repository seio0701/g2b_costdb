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
python -m g2b_costdb.pipeline excel                       # S7 Excel DB
"""
from __future__ import annotations

import argparse
import importlib
import importlib.util
import json
import logging
import os
import platform
import shutil
import sys

from typing import Optional

import pandas as pd

from . import attachments, collect, dedup, discover, extract_llm
from .api_client import ApiError
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
        tr = [classify_trade(n, (mc or "") + (" / " + sub if sub else ""), l)
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
    hist, latest, logs = dedup.dedup_latest(hits, cfg["dedup"]["price_change_warn_ratio"],
                                            float(cfg["dedup"].get("minor_notice_ratio", 0.30)))
    hist.to_parquet(_p(cfg, "notices_hist.parquet"), index=False)
    latest.to_parquet(_p(cfg, "notices_latest.parquet"), index=False)
    _logs_frame(logs).to_parquet(_p(cfg, "logs_dedup.parquet"), index=False)
    print(f"전체 {len(hist)}건 → 대표 {len(latest)}건 (프로젝트 {latest['프로젝트ID'].nunique() if len(latest) else 0}개, 경고 {len(logs)}건)")
    for l in logs[:20]:
        print(f"  - {l.get('항목')}: {l.get('비고')}")


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
    retry_nos = set()
    if retry_failed and not prev_df.empty and "추출성공" in prev_df.columns:
        retry_nos = set(prev_df.loc[prev_df["추출성공"] != "Y", "공고번호"].astype(str))
        print(f"[--retry-failed] 추출 실패 파일이 있는 공고 {len(retry_nos)}건을 다시 처리합니다(내려받은 파일은 재사용)")
    processed_nos = set()
    for i, (_, row) in enumerate(latest.iterrows(), 1):
        if row["공고키"] in texts and str(row["공고번호"]) not in retry_nos:
            continue
        processed_nos.add(str(row["공고번호"]))
        try:
            text, n = attachments.process_notice_attachments(row.to_dict(), cfg["paths"]["files_dir"], cfg["paths"]["text_dir"],
                                                             attach_max=int(cfg["fields"]["attach_max"]))
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
    print(f"첨부 처리: 공고 {len(latest)}건, 파일 {len(notes_df)}개, 텍스트 확보 {len([k for k in texts if k in have])}/{len(latest)}건")
    if not notes_df.empty and "형식" in notes_df.columns:
        s = notes_df.groupby(notes_df["형식"].fillna("(다운로드 실패)")).agg(파일수=("파일명", "size"), 추출성공=("추출성공", lambda x: int((x == "Y").sum())))
        print("형식별 추출 결과:\n" + s.to_string())
    missing = latest[~latest["공고키"].isin(set(texts))].copy()
    if missing.empty:
        return
    by_no = {}
    if not notes_df.empty:
        for no, g in notes_df.groupby(notes_df["공고번호"].astype(str)):
            by_no[no] = g
    unsupported = {".xls", ".xlsb", ".doc", ".pptx", ".7z", ".egg", ".cell", ".pme", ".dwg", ".jpg", ".png"}

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
        if forms == {".pdf"}:
            return "PDF 추출 실패(스캔)"
        return "추출 실패(" + "/".join(sorted(forms)[:3]) + ")"
    missing["텍스트없음_원인"] = [cause(r) for _, r in missing.iterrows()]
    missing["_amt"] = pd.to_numeric(missing["추정가격"], errors="coerce").fillna(0)
    cnt = missing["텍스트없음_원인"].value_counts()
    print(f"텍스트 없는 공고 {len(missing)}건의 원인:\n" + cnt.to_string())
    top = missing.sort_values("_amt", ascending=False).head(10)
    print("그중 추정가격 상위 10건:\n" + top[["공고번호", "공고명", "공종", "추정가격", "텍스트없음_원인"]].to_string(index=False))
    out = missing.drop(columns=["_amt"])[[c for c in ("공고키", "공고번호", "공고명", "시설명", "공종", "추정가격", "첨부파일수", "상세URL", "텍스트없음_원인") if c in missing.columns]]
    out.to_parquet(_p(cfg, "notes_notext.parquet"), index=False)
    print(f"→ 목록 저장: {_p(cfg, 'notes_notext.parquet')} (상세URL 로 나라장터에서 직접 확인 가능)")


def _extract_context(cfg, limit: Optional[int] = None):
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
    todo = [k for k in latest["공고키"] if k in texts and not (k in docs and "_error" not in docs[k])]
    if limit:
        # 시범 추출용: 건축 공종을 먼저, 그 안에서 추정가격이 큰 순(연면적·규모가 있는 본공사 공고문이 앞에 오도록)
        order = latest.assign(_amt=pd.to_numeric(latest["추정가격"], errors="coerce").fillna(0),
                              _arch=(latest["공종"].astype(str) == "건축").astype(int) if "공종" in latest.columns else 0)
        rank = {k: i for i, k in enumerate(order.sort_values(["_arch", "_amt"], ascending=[False, False])["공고키"])}
        todo = sorted(todo, key=lambda k: rank.get(k, 10**9))[:int(limit)]
        print(f"(--limit {limit}) 건축 공종·추정가격 큰 순으로 {len(todo)}건만 대상")
    hints = {r["공고키"]: {"공고번호": r["공고번호"], "공고명": r["공고명"], "수요기관": r["수요기관"]} for _, r in latest.iterrows()}
    return latest, texts, docs, todo, hints, cache_path


def _save_docs(cache_path, docs):
    with open(cache_path, "w", encoding="utf-8") as f:
        json.dump(docs, f, ensure_ascii=False, default=str)


def _verify_and_report(cfg, latest, docs, n_ok=None, n_err=None):
    logs = []
    for _, row in latest.iterrows():
        d = docs.get(row["공고키"], {})
        if d and "_error" not in d:
            logs.extend(extract_llm.cross_verify(row["공고번호"], row.to_dict(), d,
                                                 cfg["verify"]["price_diff_warn_ratio"], cfg["verify"]["vat_ratio"]))
    _logs_frame(logs).to_parquet(_p(cfg, "logs_verify.parquet"), index=False)
    done = len([k for k in latest["공고키"] if k in docs and "_error" not in docs[k]])
    usage_in = sum((d.get("_usage_in") or 0) for d in docs.values() if isinstance(d, dict))
    usage_out = sum((d.get("_usage_out") or 0) for d in docs.values() if isinstance(d, dict))
    head = f"LLM 추출 성공 {n_ok}건 / 실패 {n_err}건, " if n_ok is not None else ""
    print(f"{head}대표 공고 {len(latest)}건 중 추출 완료 {done}건 (누적 토큰 입력 {usage_in:,} 출력 {usage_out:,}), 검증로그 {len(logs)}건")
    if logs:
        print(pd.DataFrame(logs)["판정"].value_counts().to_string())


def _print_estimate(cfg, texts, todo, docs, latest, batch=False):
    est = extract_llm.estimate_cost(texts, todo, cfg["llm"], float(cfg["llm"].get("usd_krw", 1400)))
    if batch:
        est["usd"], est["krw"] = round(est["usd"] / 2, 2), est["krw"] // 2
    print(f"LLM 추출 대상: {len(todo)}건 (이미 완료 {len([k for k in docs if '_error' not in docs[k]])}건, 텍스트 없음 "
          f"{len([k for k in latest['공고키'] if k not in texts])}건)")
    print(f"예상 입력 {est['input_tokens']:,} 토큰 / 출력 {est['output_tokens']:,} 토큰, 모델 {est['model']}"
          f"{' (Batches 50% 할인 적용)' if batch else ''} → 약 ${est['usd']} (≈{est['krw']:,}원, 환율 {cfg['llm'].get('usd_krw', 1400)})")
    return est


def stage_extract(cfg, yes: bool = False, batch: bool = False, export: bool = False, import_: bool = False, limit: Optional[int] = None):
    """LLM 추출. 세 가지 방식:
      (기본)   extract            → 견적만 / extract --yes → API 로 한 건씩 호출
      --batch  extract --batch    → 대기 중인 배치가 있으면 결과 수거, 없으면 견적 / --batch --yes → Message Batches 제출(50% 할인, 최대 24h)
      --export extract --export   → data/llm_in/<공고키>.txt 로 내보내기(Cowork·사람이 채움) → extract --import 로 data/llm_out/*.json 수거"""
    latest, texts, docs, todo, hints, cache_path = _extract_context(cfg, limit=limit)
    in_dir, out_dir = _p(cfg, "llm_in"), _p(cfg, "llm_out")

    if export:
        n = extract_llm.export_prompts(todo, texts, hints, cfg["llm"], in_dir, out_dir)
        print(f"내보내기 완료: {n}건 → {os.path.join(in_dir, '<공고키>.txt')}  (지시문: README_지시문.md)")
        print(f"채운 JSON 은 {os.path.join(out_dir, '<공고키>.json')} 으로 저장한 뒤  python -m g2b_costdb.pipeline extract --import")
        return

    if import_:
        found, errors = extract_llm.import_results(out_dir, set(latest["공고키"]))
        for k, d in found.items():
            docs[k] = d
        _save_docs(cache_path, docs)
        print(f"가져오기: JSON {len(found)}건 반영, 문제 파일 {len(errors)}건")
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
        try:
            submitted = extract_llm.submit_batches(client, [(k, texts[k], hints.get(k, {})) for k in todo], cfg["llm"])
        except (anthropic.AuthenticationError, anthropic.PermissionDeniedError) as e:
            raise SystemExit(f"[중단] Claude API 인증 실패: {e}. {env} 값을 확인하세요.")
        state["batches"].extend(submitted)
        with open(state_path, "w", encoding="utf-8") as f:
            json.dump(state, f, ensure_ascii=False, indent=1)
        print(f"배치 {len(submitted)}개 제출 (공고 {len(todo)}건). 대개 1시간 안에, 늦어도 24시간 안에 끝납니다.")
        print("결과 수거:  python -m g2b_costdb.pipeline extract --batch")
        return

    _print_estimate(cfg, texts, todo, docs, latest)
    if not yes:
        print("실제 추출을 실행하려면:  python -m g2b_costdb.pipeline extract --yes   (50% 저렴한 배치: extract --batch --yes / 파일 인수인계: extract --export)")
        return
    if not os.environ.get(env, "").strip():
        raise SystemExit(f"환경변수 {env} 가 없습니다. PowerShell 에서 setx {env} \"키\" 로 설정한 뒤 새 터미널에서 다시 실행하세요.")
    import anthropic  # type: ignore
    n_ok = n_err = 0
    for k in todo:
        try:
            docs[k] = extract_llm.extract_with_claude(texts[k], hints.get(k, {}), cfg["llm"])
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


def assemble_trade_table(latest: pd.DataFrame, docs: dict) -> pd.DataFrame:
    rows = []
    for _, r in latest.iterrows():
        d = docs.get(r["공고키"], {}) or {}
        if "_error" in d:
            d = {}
        ev = d.get("근거문구") or {}
        rows.append({
            "프로젝트ID": r["프로젝트ID"], "시설ID": r["시설ID"], "시설명": r["시설명"], "사업유형": r["사업유형"],
            "공종": r["공종"], "공고번호": r["공고번호"], "공고차수": r["공고차수"],
            "공고명": r["공고명"], "공고일시": r["공고일시"], "수요기관": r["수요기관"],
            "추정가격_API": r.get("추정가격"), "기초금액_API": r.get("기초금액"),
            "관급자재_API": r.get("관급자재_API"), "도급자관급액_API": r.get("도급자관급액_API"), "관급자관급액_API": r.get("관급자관급액_API"),
            "예산금액_API": r.get("예산금액"),
            "낙찰금액_API": r.get("낙찰금액_API"), "낙찰률_API": r.get("낙찰률_API"), "낙찰자": r.get("낙찰자"), "낙찰하한율": r.get("낙찰하한율"),
            "도급자관급액_문서": d.get("도급자관급액_원"), "관급자관급액_문서": d.get("관급자관급액_원"),
            "공사기간_일_문서": d.get("공사기간_일"), "추정가격_문서": d.get("추정가격_원"), "기초금액_문서": d.get("기초금액_원"),
            "신뢰도": d.get("신뢰도"), "근거문구": json.dumps(ev, ensure_ascii=False) if ev else "",
            "출처파일": d.get("_source_files", ""), "상세URL": r.get("상세URL"),
        })
    return pd.DataFrame(rows)


def enrich_facility(fac: pd.DataFrame, latest: pd.DataFrame, docs: dict) -> pd.DataFrame:
    """프로젝트(시설×사업유형) 개요(연면적·층수·구조·용도·공사기간)는 건축 공종 문서 → 없으면 다른 공종 문서 순으로 채움."""
    fac = fac.copy()
    cols = ["연면적_m2", "건축면적_m2", "지하층수", "지상층수", "구조", "용도", "공사기간_일"]
    for c in cols:
        fac[c] = None
    for i, f in fac.iterrows():
        g = latest[latest["프로젝트ID"] == f["프로젝트ID"]].copy()
        g["_pri"] = g["공종"].map(lambda t: 0 if t == "건축" else 1)
        for _, r in g.sort_values("_pri").iterrows():
            d = docs.get(r["공고키"], {}) or {}
            if "_error" in d:
                continue
            for c in cols:
                cur = fac.at[i, c]
                if (cur is None or cur == "" or (isinstance(cur, float) and cur != cur)) and d.get(c) not in (None, ""):
                    fac.at[i, c] = d.get(c)
    return fac


def stage_excel(cfg):
    hist, latest = _read(cfg, "notices_hist.parquet"), _read(cfg, "notices_latest.parquet")
    docs_path = _p(cfg, "llm_docs.json")
    docs = json.load(open(docs_path, encoding="utf-8")) if os.path.exists(docs_path) else {}
    frames = [pd.read_parquet(_p(cfg, n)) for n in ("logs_dedup.parquet", "logs_verify.parquet") if os.path.exists(_p(cfg, n))]
    frames = [f for f in frames if not f.empty]
    logs = pd.concat(frames, ignore_index=True) if frames else pd.DataFrame()
    notes = pd.read_parquet(_p(cfg, "notes_attach.parquet")) if os.path.exists(_p(cfg, "notes_attach.parquet")) else pd.DataFrame()
    fac = enrich_facility(dedup.facility_summary(latest, hist), latest, docs)
    trade = assemble_trade_table(latest, docs)
    os.makedirs(cfg["paths"]["out_dir"], exist_ok=True)
    out = os.path.join(cfg["paths"]["out_dir"], cfg["paths"]["excel_name"])
    try:
        build_workbook(out, fac, hist, latest, trade, logs, notes, meta={"period": f"{cfg['period']['start']}~{cfg['period']['end']}"})
    except PermissionError:
        raise SystemExit(_excel_busy(out))
    print(f"Excel DB 저장: {out} (프로젝트 {len(fac)}, 대표공고 {len(latest)}, 이력 {len(hist)}, 검증로그 {len(logs)}, 추출노트 {len(notes)})")
    if not logs.empty and "판정" in logs.columns:
        print("검증로그 판정별 건수:\n" + logs["판정"].value_counts().to_string())


def main(argv=None):
    for stream in (sys.stdout, sys.stderr):  # Windows cp949 콘솔에서 특수문자로 죽지 않도록
        try:
            stream.reconfigure(errors="replace")
        except (AttributeError, ValueError):
            pass
    ap = argparse.ArgumentParser(description="나라장터 공사비 DB 파이프라인 (g2b_costdb 폴더에서 실행)")
    ap.add_argument("stage", choices=["doctor", "probe", "collect", "discover", "research", "dedup", "attach", "extract", "excel"])
    ap.add_argument("--config", default=None)
    ap.add_argument("--ym", default="2026-08", help="probe 대상 월(YYYY-MM)")
    ap.add_argument("--fresh", action="store_true", help="discover: 이전 검수 파일의 시설ID 만 이어받고 검수 칸은 새 기본값으로(검수 시작 전 규칙이 바뀌었을 때)")
    ap.add_argument("--retry-failed", dest="retry_failed", action="store_true", help="attach: 추출 실패 파일이 있는 공고를 텍스트가 있어도 다시 처리(HWP 백엔드 설치 뒤 일괄 재시도)")
    ap.add_argument("--report", action="store_true", help="attach: 처리 없이 첨부 결과 요약과 텍스트 없는 공고의 원인만 출력")
    ap.add_argument("--limit", type=int, default=None, help="extract: 건축 공종·추정가격 큰 순으로 N건만(시범 추출용)")
    ap.add_argument("--yes", action="store_true", help="extract: 비용 견적 확인 후 실제 LLM 호출(또는 배치 제출) 실행")
    ap.add_argument("--batch", action="store_true", help="extract: Message Batches 로 제출/수거 (50%% 할인, 최대 24시간)")
    ap.add_argument("--export", action="store_true", help="extract: 공고별 작업지시 txt 를 data/llm_in 에 내보내기 (Cowork·사람이 채움)")
    ap.add_argument("--import", dest="import_", action="store_true", help="extract: data/llm_out 의 JSON 을 가져와 반영")
    a = ap.parse_args(argv)
    logging.basicConfig(level=logging.INFO, format="%(asctime)s %(levelname)s %(name)s: %(message)s")
    cfg = load_config(a.config)
    for k in ("data_dir", "raw_dir", "text_dir", "files_dir", "out_dir"):
        os.makedirs(cfg["paths"][k], exist_ok=True)
    if a.stage == "doctor":
        stage_doctor(cfg)
    elif a.stage == "probe":
        collect.probe(cfg, a.ym)
    elif a.stage == "extract":
        stage_extract(cfg, yes=a.yes, batch=a.batch, export=a.export, import_=a.import_, limit=a.limit)
    elif a.stage == "discover":
        stage_discover(cfg, fresh=a.fresh)
    elif a.stage == "attach":
        stage_attach(cfg, retry_failed=a.retry_failed, report_only=a.report)
    else:
        {"collect": stage_collect, "discover": stage_discover, "research": stage_research, "dedup": stage_dedup,
         "attach": stage_attach, "excel": stage_excel}[a.stage](cfg)


if __name__ == "__main__":
    sys.exit(main())
