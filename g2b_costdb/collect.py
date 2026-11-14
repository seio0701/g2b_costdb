"""S1 전량 수집 / S3 면허제한 조회 / probe.

- collect_notices(): 공사 공고목록을 월 단위로 전량 수집 → data/raw/notices_YYYYMM.jsonl + data/notices_all.parquet
- collect_bsis_amounts(): 공사 기초금액을 월 단위로 전량 수집 → data/bsis_all.parquet
- collect_license_limits(bid_nos): 공고번호별 면허제한 업종 → data/license_all.parquet
- probe(): 1개월 표본으로 실제 응답 필드명 확인 + config.yaml fields 매핑 자동 대조

진행 중인 달(오늘이 속한 달)은 완료 표시를 하지 않고 매 실행 시 다시 받는다(캐시 키에 날짜를 섞음).
미래 달은 건너뛴다. 원문 JSONL은 절대 덮어쓰지 않는 것이 원칙이지만, 완료되지 않은 달의 파일은 재수집 시 교체한다.
"""
from __future__ import annotations

import json
import logging
import math
import os
from datetime import date
from typing import Dict, Iterable, List, Optional

import pandas as pd

from .api_client import ApiConfig, ApiError, DailyBudgetExceeded, G2BClient, explain_code
from .classify import load_config, parse_amount

log = logging.getLogger(__name__)

# 면허제한 응답에서 업종명이 들어 있을 수 있는 필드 후보(활용가이드 버전에 따라 다름) — probe 결과로 config.fields.license_name 확정
LICENSE_NAME_CANDIDATES = ["lcnsLmtNm", "indstrytyNm", "permsnIndstrytyList", "lmtNm", "indstrytyLmtNm"]


AWARD_FIELD_CANDIDATES = {
    "award_amt": ["sucsfbidAmt", "scsbidAmt", "sucsfbidPrce", "bidprcAmt"],
    "award_rate": ["sucsfbidRate", "scsbidRate", "sucsfbidRt", "bidprcRate"],
    "award_bidder": ["bidwinnrNm", "scsbidCorpNm", "bidwinnrCorpNm", "prcbdrNm"],
    "award_bidder_bizno": ["bidwinnrBizno", "scsbidBizno"],
    "award_prtcpt_cnt": ["prtcptCnum", "bidPrtcptCnum", "prtcptCnt"],
    "award_open_dt": ["opengDt", "rlOpengDt", "sucsfbidDt"],
}


def make_client(cfg: dict, base_url: str = None) -> G2BClient:
    key = (os.environ.get(cfg["api"]["service_key_env"], "") or "").strip().strip('"').strip("'")
    if not key:
        raise SystemExit(
            f"환경변수 {cfg['api']['service_key_env']} 에 공공데이터포털 '일반 인증키(Decoding)'를 설정하세요.\n"
            f"  PowerShell(영구): setx {cfg['api']['service_key_env']} \"발급받은키\"  → 새 터미널을 열어야 적용됩니다.\n"
            f"  (키를 채팅이나 파일에 붙여 넣지 마세요)")
    a = cfg["api"]
    return G2BClient(
        ApiConfig(base_url=base_url or a["base_url"], service_key=key, num_of_rows=int(a["num_of_rows"]),
                  timeout_sec=int(a["timeout_sec"]), max_retries=int(a["max_retries"]),
                  sleep_between_calls_sec=float(a["sleep_between_calls_sec"]),
                  daily_call_budget=int(a["daily_call_budget"])),
        cache_db=cfg["paths"]["cache_db"],
    )


def _append_jsonl(path: str, rows: Iterable[dict]) -> int:
    os.makedirs(os.path.dirname(path), exist_ok=True)
    n = 0
    with open(path, "a", encoding="utf-8") as f:
        for r in rows:
            f.write(json.dumps(r, ensure_ascii=False) + "\n")
            n += 1
    return n


def _load_jsonl_dir(raw_dir: str, prefix: str) -> pd.DataFrame:
    frames = []
    for fn in sorted(os.listdir(raw_dir)) if os.path.isdir(raw_dir) else []:
        if fn.startswith(prefix) and fn.endswith(".jsonl"):
            with open(os.path.join(raw_dir, fn), encoding="utf-8") as f:
                rows = [json.loads(l) for l in f if l.strip()]
            if rows:
                frames.append(pd.DataFrame(rows))
    return pd.concat(frames, ignore_index=True) if frames else pd.DataFrame()


def _stringify(df: pd.DataFrame) -> pd.DataFrame:
    """월별 응답의 타입이 섞여도(숫자/문자/None) parquet 저장이 실패하지 않도록 모든 원문 컬럼을 문자열(또는 None)로 통일."""
    out = df.copy()
    for c in out.columns:
        out[c] = out[c].map(lambda v: None if v is None or (isinstance(v, float) and math.isnan(v)) else
                            (json.dumps(v, ensure_ascii=False) if isinstance(v, (dict, list)) else str(v)))
        out[c] = out[c].astype(object)
    return out


def _read_done(path: str) -> set:
    return set(open(path, encoding="utf-8").read().split()) if os.path.exists(path) else set()


def _write_done(path: str, done: set) -> None:
    os.makedirs(os.path.dirname(path), exist_ok=True)
    with open(path, "w", encoding="utf-8") as f:
        f.write("\n".join(sorted(done)))


def _collect_monthly(cfg: dict, op_key: str, prefix: str, start: Optional[str], end: Optional[str],
                     base_url: Optional[str] = None, op_name: Optional[str] = None, extra: Optional[dict] = None,
                     row_filter=None, purge_cache: bool = False) -> pd.DataFrame:
    """월 단위 전량 수집 공통 루틴. 반환: 누적 원문 DataFrame(문자열화).
    row_filter: 저장할 행만 고르는 함수(용역 공고처럼 대부분이 무관한 경우). purge_cache: 완료한 달의 응답 캐시를 비움."""
    client = make_client(cfg, base_url)
    op = op_name or cfg["api"]["ops"][op_key]
    raw_dir = cfg["paths"]["raw_dir"]
    start = start or cfg["period"]["start"]
    end = end or cfg["period"]["end"]
    done_marker = os.path.join(raw_dir, f"{prefix}done.txt")
    done = _read_done(done_marker)
    today_ym = date.today().strftime("%Y%m")
    skipped_future, collected, failed = [], [], []
    try:
        for bgn, endd in G2BClient.month_ranges(start, end):
            ym = bgn[:6]
            if ym > today_ym:
                skipped_future.append(ym)
                continue
            is_current = ym == today_ym
            if ym in done and not is_current:
                continue
            path = os.path.join(raw_dir, f"{prefix}{ym}.jsonl")
            # 진행 중인 달, 또는 done.txt 에서 지워 재수집을 요청한 달(파일은 있음)은 캐시를 우회해 새로 받는다
            refresh = is_current or os.path.exists(path)
            try:
                rows = list(client.iter_all(op, {**(extra or {"inqryDiv": "1"}), "inqryBgnDt": bgn, "inqryEndDt": endd},
                                            cache_salt=(date.today().isoformat() if refresh else None)))
            except ApiError as e:
                if e.code in ("10", "11", "12", "20", "30", "31", "32", "33"):
                    raise
                log.warning("월 %s 수집 실패(건너뜀): %s", ym, e)
                failed.append(ym)
                continue
            except RuntimeError as e:
                log.warning("월 %s 수집 실패(건너뜀): %s", ym, e)
                failed.append(ym)
                continue
            n_all = len(rows)
            if row_filter is not None:
                rows = [r for r in rows if row_filter(r)]
            if os.path.exists(path):
                os.remove(path)
            n = _append_jsonl(path, rows)
            log.info("%s %s: %s%d건%s", op, ym, f"전체 {n_all:,}건 중 보존 " if row_filter is not None else "", n,
                     " (진행 중인 달 — 완료 표시 안 함)" if is_current else "")
            collected.append(ym)
            if not is_current:
                done.add(ym)
                _write_done(done_marker, done)
                if purge_cache:
                    client.purge_cache(op, f'"inqryBgnDt": "{bgn}"')
    except DailyBudgetExceeded as e:
        log.warning("%s → 내일 같은 명령으로 재개하면 완료된 월은 건너뜁니다. (오늘 호출 %d회)", e, client.calls_today())
    except ApiError as e:
        log.error("%s", e)
        raise SystemExit(f"[중단] {e}")
    if skipped_future:
        log.info("미래 달 %s 은 건너뜀", ", ".join(skipped_future))
    remaining = [b[:6] for b, _ in G2BClient.month_ranges(start, end) if b[:6] < today_ym and b[:6] not in done]
    if failed:
        log.warning("%s 수집 실패 월: %s (다음 실행 때 다시 시도)", op, ", ".join(failed))
    if remaining:
        log.warning("[미완료] %s: 완료 %d개월, 미완료 %d개월 → %s%s  (내일 같은 명령으로 재개)", op, len(done), len(remaining),
                    ", ".join(remaining[:12]), " …" if len(remaining) > 12 else "")
    else:
        log.info("[완료] %s: %d개월 수집 완료", op, len(done))
    df = _load_jsonl_dir(raw_dir, prefix)
    return _stringify(df).copy() if not df.empty else df


def incomplete_months(cfg: dict, prefix: str) -> List[str]:
    """done.txt 기준으로 아직 수집되지 않은(오늘 이전) 월 목록 — discover 단계에서 경고용."""
    done = _read_done(os.path.join(cfg["paths"]["raw_dir"], f"{prefix}done.txt"))
    today_ym = date.today().strftime("%Y%m")
    return [b[:6] for b, _ in G2BClient.month_ranges(cfg["period"]["start"], cfg["period"]["end"]) if b[:6] < today_ym and b[:6] not in done]


def collect_notices(cfg: dict, start: str = None, end: str = None) -> pd.DataFrame:
    df = _collect_monthly(cfg, "cnstwk_list", "notices_", start, end)
    os.makedirs(cfg["paths"]["data_dir"], exist_ok=True)
    if not df.empty:
        f = cfg["fields"]
        df["_presmpt_price"] = (df[f["presmpt_price"]] if f["presmpt_price"] in df.columns
                                else pd.Series([None] * len(df), index=df.index)).map(parse_amount)
        df.to_parquet(os.path.join(cfg["paths"]["data_dir"], "notices_all.parquet"), index=False)
    log.info("누적 공사 공고 %d건 → notices_all.parquet", len(df))
    return df


def collect_bsis_amounts(cfg: dict, start: str = None, end: str = None) -> pd.DataFrame:
    df = _collect_monthly(cfg, "cnstwk_bsis_amount", "bsis_", start, end)
    if not df.empty:
        f = cfg["fields"]
        df["_bsis_amount"] = (df[f["bsis_amount"]] if f["bsis_amount"] in df.columns
                              else pd.Series([None] * len(df), index=df.index)).map(parse_amount)
        df.to_parquet(os.path.join(cfg["paths"]["data_dir"], "bsis_all.parquet"), index=False)
    log.info("누적 기초금액 %d건 → bsis_all.parquet", len(df))
    return df


# ── 용역 입찰공고(폐기물처리 용역·건설사업관리·설계용역) ─────────────────────────────────────────────
SERVC_KEEP_NAME = r"폐기물|건설사업관리|감리|설계\s*용역|설계\s*공모|실시\s*설계|기본\s*설계|기본\s*및\s*실시|CM\s*용역"
SERVC_KEEP_CLASS = r"폐기물"


def servc_period(cfg: dict) -> tuple:
    """용역 수집 기간: config servc.period_start(없으면 공사 수집 시작월) ~ servc.period_end(없으면 오늘이 속한 달 — 폐기물처리 용역은 공사 뒤에 발주)."""
    sv = cfg.get("servc") or {}
    return (str(sv.get("period_start") or cfg["period"]["start"]), str(sv.get("period_end") or date.today().strftime("%Y-%m")))


def servc_keep_filter(cfg: dict):
    """용역 공고 중 보존할 행: 공고명이 keep_name_regex 에 걸리거나 공공조달분류명이 keep_class_regex 에 걸리는 공고.
    용역 공고는 월 1.5만~2.5만 건이고 대부분(청소·경비·전산 등)이 무관하므로 관련 공고만 원문으로 남긴다."""
    import re
    sv = cfg.get("servc") or {}
    f = cfg["fields"]
    name_re = re.compile(sv.get("keep_name_regex") or SERVC_KEEP_NAME)
    cls_re = re.compile(sv.get("keep_class_regex") or SERVC_KEEP_CLASS)
    name_k = f["name"]
    cls_keys = [f.get(k) for k in ("servc_class", "servc_class_mid", "servc_class_lrg") if f.get(k)]

    def keep(row: dict) -> bool:
        if name_re.search(str(row.get(name_k) or "")):
            return True
        return any(cls_re.search(str(row.get(k) or "")) for k in cls_keys)
    return keep


def collect_servc(cfg: dict, start: str = None, end: str = None) -> pd.DataFrame:
    """용역 입찰공고 목록을 월 단위로 받아 관련 공고(폐기물·건설사업관리·설계)만 data/raw/servc_YYYYMM.jsonl + data/servc_all.parquet 에 저장.
    공사 수집과 같은 서비스·같은 인증키. 호출량이 크므로(월 15~26회, 전 기간 약 1,300~2,400회) 일일 예산 안에서 여러 날에 걸쳐 이어받는다."""
    if "servc_list" not in (cfg["api"].get("ops") or {}):
        raise SystemExit("config.yaml api.ops.servc_list 가 없습니다(용역 입찰공고 목록 오퍼레이션명).")
    s0, e0 = servc_period(cfg)
    df = _collect_monthly(cfg, "servc_list", "servc_", start or s0, end or e0, row_filter=servc_keep_filter(cfg), purge_cache=True)
    os.makedirs(cfg["paths"]["data_dir"], exist_ok=True)
    if not df.empty:
        f = cfg["fields"]
        df["_presmpt_price"] = (df[f["presmpt_price"]] if f["presmpt_price"] in df.columns
                                else pd.Series([None] * len(df), index=df.index)).map(parse_amount)
        df.to_parquet(os.path.join(cfg["paths"]["data_dir"], "servc_all.parquet"), index=False)
        name = df[f["name"]].astype(str) if f["name"] in df.columns else pd.Series([""] * len(df))
        log.info("누적 용역 공고(보존분) %d건 → servc_all.parquet (공고명 '폐기물' %d건, '건설사업관리·감리' %d건, '설계' %d건)", len(df),
                 int(name.str.contains("폐기물").sum()), int(name.str.contains("건설사업관리|감리").sum()), int(name.str.contains("설계").sum()))
    else:
        log.info("누적 용역 공고(보존분) 0건")
    return df


def probe_servc(cfg: dict, ym: str) -> None:
    """용역 입찰공고 오퍼레이션 시험 호출(2~3회): 응답 필드·표본, config 매핑 대조, 한 달 건수로 전 기간 호출 수 추정, 공고명 검색(PPSSrch) 동작 확인."""
    client = make_client(cfg)
    bgn, end = G2BClient.month_ranges(ym, ym)[0]
    f = cfg["fields"]
    ops = cfg["api"].get("ops") or {}
    op = ops.get("servc_list")
    if not op:
        print("config.yaml api.ops.servc_list 가 없습니다.")
        return
    which = ["bid_no", "bid_ord", "name", "kind", "notice_dt", "notice_inst", "demand_inst", "demand_inst_cd", "presmpt_price", "vat",
             "servc_budget", "servc_div", "servc_class", "prespec_no", "detail_url", "attach_url_prefix", "attach_name_prefix"]
    try:
        body = client.call(op, {"inqryDiv": "1", "inqryBgnDt": bgn, "inqryEndDt": end, "pageNo": 1, "numOfRows": 50}, use_cache=False)
    except ApiError as e:
        print(f"\n=== {op} ({ym}) 실패: API 오류 {e.code} ({explain_code(e.code)}) — 공공데이터포털 마이페이지에서 「조달청_나라장터 입찰공고정보서비스」 활용신청의 "
              f"상세기능에 용역 오퍼레이션이 포함돼 있는지, 오퍼레이션명(config api.ops.servc_list)이 맞는지 확인")
        return
    items = client._items(body)
    total = body.get("totalCount")
    print(f"\n=== {op} ({ym}) totalCount={total} 표본 {len(items)}건 ===")
    if not items:
        print("항목 없음 — 다른 달(--ym)로 다시 시도하세요.")
        return
    keys = sorted(set().union(*[it.keys() for it in items]))
    print("필드:", keys)
    print("표본:", json.dumps(items[0], ensure_ascii=False, indent=1)[:2500])
    missing = _check_mapping(f, keys, [k for k in which if k in f])
    if missing:
        print("[매핑 확인 필요] config.yaml fields 중 용역 응답에 없는 항목:", {k: f[k] for k in missing})
    else:
        print("[매핑 OK] config.yaml fields 의 용역 관련 항목이 모두 응답에 존재합니다.")
    keep = servc_keep_filter(cfg)
    kept = [it for it in items if keep(it)]
    print(f"[보존 규칙] 표본 {len(items)}건 중 {len(kept)}건이 보존 대상(공고명·조달분류 기준): "
          + "; ".join(str(it.get(f['name']))[:40] for it in kept[:5]))
    try:
        n = int(float(total))
        pages = math.ceil(n / int(cfg["api"]["num_of_rows"]))
        s0, e0 = servc_period(cfg)
        months = len(G2BClient.month_ranges(s0, e0))
        print(f"[호출 수 추정] {ym} 한 달 {n:,}건 = {pages}회. 수집 기간 {s0}~{e0}({months}개월)을 같은 규모로 보면 약 {pages * months:,}회 "
              f"→ 일 {cfg['api']['daily_call_budget']}회 예산으로 약 {math.ceil(pages * months / int(cfg['api']['daily_call_budget']))}일(공사 수집과 예산을 함께 씀)")
    except (TypeError, ValueError):
        print("[호출 수 추정] totalCount 가 없어 추정하지 못했습니다.")
    sop = ops.get("servc_search")
    if sop:
        try:
            b2 = client.call(sop, {"inqryDiv": "1", "inqryBgnDt": bgn, "inqryEndDt": end, f["name"]: "폐기물", "pageNo": 1, "numOfRows": 5}, use_cache=False)
            it2 = client._items(b2)
            print(f"\n=== {sop} 공고명 '폐기물' ({ym}) totalCount={b2.get('totalCount')} 표본 {len(it2)}건 ===")
            for it in it2[:5]:
                print("  -", str(it.get(f["name"]))[:70], "|", it.get(f["demand_inst"]), "| 추정가격", it.get(f["presmpt_price"]),
                      "|", it.get(f.get("servc_class", ""), ""))
            print("[검색] 공고명 검색이 동작하면 totalCount 가 위 전체 건수보다 훨씬 작습니다(같으면 검색 조건이 무시된 것).")
        except ApiError as e:
            print(f"\n=== {sop} 실패: API 오류 {e.code} ({explain_code(e.code)}) — 검색 오퍼레이션은 참고용이라 수집에는 영향 없음")
    print(f"\n오늘 API 호출 {client.calls_today()}회 (예산 {cfg['api']['daily_call_budget']})")
    print("다음: python -m g2b_costdb.pipeline collect --servc   (여러 날에 걸쳐 같은 명령으로 이어받기)")


def collect_awards(cfg: dict, start: str = None, end: str = None) -> pd.DataFrame:
    """공사 낙찰 목록(낙찰정보서비스)을 월 단위로 전량 수집 → data/raw/award_YYYYMM.jsonl + data/award_all.parquet"""
    a = cfg["api"]
    df = _collect_monthly(cfg, "", "award_", start, end, base_url=a.get("award_base_url"),
                          op_name=(a.get("award_ops") or {}).get("cnstwk_list", "getScsbidListSttusCnstwk"),
                          extra=dict(a.get("award_query") or {"inqryDiv": "1"}))
    if not df.empty:
        amt_col = award_column(cfg, df, "award_amt")
        rate_col = award_column(cfg, df, "award_rate")
        df["_award_amt"] = df[amt_col].map(parse_amount) if amt_col else None
        df["_award_rate"] = pd.to_numeric(df[rate_col].astype(str).str.replace("%", "").str.replace(",", ""), errors="coerce") if rate_col else None
        df.to_parquet(os.path.join(cfg["paths"]["data_dir"], "award_all.parquet"), index=False)
    log.info("누적 낙찰 %d건 → award_all.parquet", len(df))
    return df


def award_column(cfg: dict, df: pd.DataFrame, key: str) -> Optional[str]:
    """낙찰 응답에서 key(award_amt 등)에 해당하는 컬럼명(config 매핑 → 후보 순)."""
    if df is None or df.empty:
        return None
    for c in [cfg["fields"].get(key, "")] + AWARD_FIELD_CANDIDATES.get(key, []):
        if c and c in df.columns:
            return c
    return None


def license_query_params(cfg: dict, bid_no: str) -> Dict[str, str]:
    """면허제한 조회 파라미터. config.api.license_query 로 조정(probe 결과에 따라 inqryDiv 값이 다를 수 있음)."""
    lq = dict((cfg["api"].get("license_query") or {"inqryDiv": "2"}))
    lq[cfg["fields"]["bid_no"]] = bid_no
    lq.setdefault("pageNo", "1")
    lq.setdefault("numOfRows", "100")
    return lq


def collect_license_limits(cfg: dict, bid_nos: List[str]) -> pd.DataFrame:
    """공고번호별 면허제한(업종) 조회 — 재검색된 공고에 대해서만 호출하여 호출량 절약. 캐시된 공고는 호출하지 않음."""
    client = make_client(cfg)
    op = cfg["api"]["ops"]["license_limit"]
    out_path = os.path.join(cfg["paths"]["data_dir"], "license_all.parquet")
    rows: List[dict] = []
    n_ok = 0
    try:
        for i, no in enumerate(bid_nos):
            try:
                body = client.call(op, license_query_params(cfg, str(no)))
            except ApiError as e:
                if e.code in ("10", "11", "12") and n_ok == 0:
                    log.error("면허제한 조회 파라미터가 맞지 않는 것으로 보입니다(%s). config.yaml의 api.license_query 를 probe 결과에 맞게 조정하세요.", e)
                    break
                log.warning("면허제한 조회 실패 %s: %s", no, e)
                continue
            n_ok += 1
            for it in client._items(body):
                it["_bid_no"] = str(no)
                rows.append(it)
            if (i + 1) % 200 == 0:
                log.info("면허제한 조회 %d/%d (오늘 호출 %d회)", i + 1, len(bid_nos), client.calls_today())
    except DailyBudgetExceeded as e:
        log.warning("%s — 지금까지 조회된 %d건은 캐시에 남아 있으므로 내일 research 를 다시 실행하면 이어서 진행됩니다.", e, n_ok)
    df = _stringify(pd.DataFrame(rows)) if rows else pd.DataFrame()
    if not df.empty:
        df.to_parquet(out_path, index=False)
    log.info("면허제한: 공고 %d건 조회 → 업종 행 %d건", n_ok, len(df))
    return df


def license_name_column(cfg: dict, lic: pd.DataFrame) -> Optional[str]:
    """면허제한 DataFrame에서 업종명 컬럼을 찾는다(config 매핑 → 후보 목록 순)."""
    if lic is None or lic.empty:
        return None
    for c in [cfg["fields"].get("license_name", "")] + LICENSE_NAME_CANDIDATES:
        if c and c in lic.columns:
            return c
    return None


# ── probe ──────────────────────────────────────────────────────
def _check_mapping(fields: Dict[str, str], keys: List[str], which: List[str]) -> List[str]:
    """config.fields 중 which 에 해당하는 매핑이 응답 필드에 있는지 확인 → 없는 내부명 목록."""
    missing = []
    for k in which:
        v = fields.get(k)
        if k in ("attach_url_prefix", "attach_name_prefix"):
            if not any(x.startswith(str(v)) for x in keys):
                missing.append(k)
        elif isinstance(v, str) and v not in keys:
            missing.append(k)
    return missing


def probe(cfg: dict, ym: str) -> None:
    """한 달치 첫 페이지만 호출하여 필드명·표본을 출력하고 config.yaml fields 매핑을 대조."""
    client = make_client(cfg)
    bgn, end = G2BClient.month_ranges(ym, ym)[0]
    f = cfg["fields"]
    first_bid_no = None
    plan = [
        ("cnstwk_list", ["bid_no", "bid_ord", "name", "kind", "re_notice", "reg_type", "notice_dt", "close_dt", "open_dt",
                         "notice_inst", "demand_inst", "presmpt_price", "main_cnstty", "site_region", "detail_url",
                         "ref_no", "prespec_no", "attach_url_prefix", "attach_name_prefix"]),
        ("cnstwk_bsis_amount", ["bid_no", "bid_ord", "bsis_amount"]),
    ]
    for op_key, which in plan:
        op = cfg["api"]["ops"][op_key]
        try:
            body = client.call(op, {"inqryDiv": "1", "inqryBgnDt": bgn, "inqryEndDt": end, "pageNo": 1, "numOfRows": 5},
                               use_cache=False)
        except ApiError as e:
            print(f"\n=== {op} ({ym}) 실패: {e}")
            continue
        items = client._items(body)
        print(f"\n=== {op} ({ym}) totalCount={body.get('totalCount')} 표본 {len(items)}건 ===")
        if not items:
            print("항목 없음 — 다른 달(--ym)로 다시 시도하세요.")
            continue
        keys = sorted(set().union(*[it.keys() for it in items]))
        print("필드:", keys)
        print("표본:", json.dumps(items[0], ensure_ascii=False, indent=1)[:2500])
        missing = _check_mapping(f, keys, which)
        if missing:
            print("[매핑 확인 필요] config.yaml fields 중 응답에 없는 항목:", {k: f[k] for k in missing})
        else:
            print("[매핑 OK] config.yaml fields 의 관련 항목이 모두 응답에 존재합니다.")
        if op_key == "cnstwk_list" and f["bid_no"] in items[0]:
            first_bid_no = str(items[0][f["bid_no"]])
    # 면허제한: 조회구분(inqryDiv) 변형을 차례로 시험하여 결과가 나오는 파라미터를 안내 (config api.use_license_limit 가 false 면 참고용)
    if first_bid_no:
        op = cfg["api"]["ops"]["license_limit"]
        variants = [dict(cfg["api"].get("license_query") or {"inqryDiv": "2"}),
                    {"inqryDiv": "3"}, {"inqryDiv": "1", "inqryBgnDt": bgn, "inqryEndDt": end}]
        seen_ok = None
        for v in variants:
            params = dict(v); params[f["bid_no"]] = first_bid_no; params.setdefault("pageNo", "1"); params.setdefault("numOfRows", "100")
            try:
                body = client.call(op, params, use_cache=False)
            except ApiError as e:
                print(f"\n=== {op} params={v} 실패: API 오류 {e.code} ({explain_code(e.code)})")
                continue
            items = client._items(body)
            print(f"\n=== {op} params={v} (공고번호 {first_bid_no}) totalCount={body.get('totalCount')} 항목 {len(items)}건 ===")
            if items:
                keys = sorted(set().union(*[it.keys() for it in items]))
                print("필드:", keys)
                print("표본:", json.dumps(items[0], ensure_ascii=False, indent=1)[:1200])
                col = next((c for c in [f.get("license_name", "")] + LICENSE_NAME_CANDIDATES if c and c in keys), None)
                print(f"[업종명 필드] {'config 매핑 ' + f.get('license_name', '') + ' 존재' if f.get('license_name') in keys else ('후보 ' + col + ' 발견 → config.yaml fields.license_name 을 이 값으로 수정' if col else '후보 없음 → 표본에서 업종명 필드를 찾아 config.yaml fields.license_name 수정')}")
                seen_ok = v
                break
        if seen_ok:
            print(f"[면허제한] 결과가 나온 파라미터: {seen_ok} → 사용하려면 config.yaml api.license_query 를 이 값으로, api.use_license_limit 를 true 로")
        else:
            print("[면허제한] 어떤 조회구분으로도 항목이 없음 — 이 공고에 면허제한이 없을 수 있음. 공종 분류는 주공종명·부공종명으로 충분하므로 use_license_limit 는 false 유지")
    if cfg["api"].get("use_awards"):
        probe_awards(cfg, ym, first_bid_no)
    print(f"\n오늘 API 호출 {client.calls_today()}회 (예산 {cfg['api']['daily_call_budget']})")


def probe_awards(cfg: dict, ym: str, sample_bid_no: Optional[str] = None) -> None:
    """낙찰정보서비스 표본 조회: 기간 조회 변형을 시험하고 낙찰 필드 매핑을 대조."""
    a = cfg["api"]
    base = a.get("award_base_url")
    op = (a.get("award_ops") or {}).get("cnstwk_list", "getScsbidListSttusCnstwk")
    f = cfg["fields"]
    if not base:
        print("\n[낙찰정보] config.yaml api.award_base_url 이 없습니다.")
        return
    client = make_client(cfg, base)
    bgn, end = G2BClient.month_ranges(ym, ym)[0]
    variants = [dict(a.get("award_query") or {"inqryDiv": "1"}), {"inqryDiv": "2"}, {"inqryDiv": "3"}]
    seen = []
    ok = None
    for v in variants:
        if v in seen:
            continue
        seen.append(v)
        params = dict(v); params.update({"inqryBgnDt": bgn, "inqryEndDt": end, "pageNo": 1, "numOfRows": 5})
        try:
            body = client.call(op, params, use_cache=False)
        except ApiError as e:
            print(f"\n=== {op} params={v} 실패: API 오류 {e.code} ({explain_code(e.code)})")
            if e.code in ("20", "30", "31", "32"):
                print("→ 「조달청_나라장터 낙찰정보서비스」 활용신청이 승인되었는지, api.award_base_url 이 포털 페이지의 End Point 와 같은지 확인")
                return
            continue
        items = client._items(body)
        print(f"\n=== {op} params={v} ({ym}) totalCount={body.get('totalCount')} 표본 {len(items)}건 ===")
        if not items:
            continue
        keys = sorted(set().union(*[it.keys() for it in items]))
        print("필드:", keys)
        print("표본:", json.dumps(items[0], ensure_ascii=False, indent=1)[:2000])
        for k in ("award_amt", "award_rate", "award_bidder", "award_prtcpt_cnt", "award_open_dt"):
            col = next((c for c in [f.get(k, "")] + AWARD_FIELD_CANDIDATES.get(k, []) if c and c in keys), None)
            state = "OK" if f.get(k) in keys else (f"후보 {col} 발견 → config.yaml fields.{k} 를 이 값으로" if col else "없음 → 표본에서 찾아 config.yaml fields 수정")
            print(f"  [{k}] {state}")
        if f["bid_no"] not in keys:
            print(f"  [경고] 공고번호 필드 {f['bid_no']} 가 없어 공고와 연결할 수 없음 → 표본에서 공고번호 필드를 확인")
        ok = v
        break
    if ok:
        print(f"[낙찰정보] 기간 조회 파라미터 {ok} 로 결과 확인 → config.yaml api.award_query 를 이 값으로 두고 collect 를 실행하면 낙찰 목록도 수집됩니다")
    else:
        print("[낙찰정보] 어떤 조회구분으로도 항목이 없음 — 활용가이드의 기간 조회 파라미터를 확인")
