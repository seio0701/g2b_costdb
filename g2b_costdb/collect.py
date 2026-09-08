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


def make_client(cfg: dict) -> G2BClient:
    key = (os.environ.get(cfg["api"]["service_key_env"], "") or "").strip().strip('"').strip("'")
    if not key:
        raise SystemExit(
            f"환경변수 {cfg['api']['service_key_env']} 에 공공데이터포털 '일반 인증키(Decoding)'를 설정하세요.\n"
            f"  PowerShell(영구): setx {cfg['api']['service_key_env']} \"발급받은키\"  → 새 터미널을 열어야 적용됩니다.\n"
            f"  (키를 채팅이나 파일에 붙여 넣지 마세요)")
    a = cfg["api"]
    return G2BClient(
        ApiConfig(base_url=a["base_url"], service_key=key, num_of_rows=int(a["num_of_rows"]),
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


def _collect_monthly(cfg: dict, op_key: str, prefix: str, start: Optional[str], end: Optional[str]) -> pd.DataFrame:
    """월 단위 전량 수집 공통 루틴. 반환: 누적 원문 DataFrame(문자열화)."""
    client = make_client(cfg)
    op = cfg["api"]["ops"][op_key]
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
                rows = list(client.iter_all(op, {"inqryDiv": "1", "inqryBgnDt": bgn, "inqryEndDt": endd},
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
            if os.path.exists(path):
                os.remove(path)
            n = _append_jsonl(path, rows)
            log.info("%s %s: %d건%s", op, ym, n, " (진행 중인 달 — 완료 표시 안 함)" if is_current else "")
            collected.append(ym)
            if not is_current:
                done.add(ym)
                _write_done(done_marker, done)
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
    print(f"\n오늘 API 호출 {client.calls_today()}회 (예산 {cfg['api']['daily_call_budget']})")
