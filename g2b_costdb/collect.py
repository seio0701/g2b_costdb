"""S1 전량 수집 / S3 면허제한 조회 / probe.

- collect_notices(): 공사 공고목록을 월 단위로 전량 수집 → data/raw/notices_YYYYMM.jsonl + data/notices_all.parquet
- collect_bsis_amounts(): 공사 기초금액을 월 단위로 전량 수집 → data/bsis_all.parquet
- collect_license_limits(bid_nos): 공고번호별 면허제한 업종 → data/license_all.parquet
- probe(): 그 달의 첫 구간(api.max_window_days, 기본 7일) 표본으로 실제 응답 필드명 확인 + config.yaml fields 매핑 자동 대조

진행 중인 달(오늘이 속한 달)은 완료 표시를 하지 않고 매 실행 시 다시 받는다(캐시 키에 날짜를 섞음).
미래 달은 건너뛴다. 원문 JSONL은 절대 덮어쓰지 않는 것이 원칙이지만, 완료되지 않은 달의 파일은 재수집 시 교체한다.
"""
from __future__ import annotations

import hashlib
import json
import logging
import math
import os
import re
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
                  daily_call_budget=int(a["daily_call_budget"]),
                  max_window_days=7 if a.get("max_window_days") in (None, "") else int(a["max_window_days"])),
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


def _count_lines(path: str) -> int:
    with open(path, encoding="utf-8") as f:
        return sum(1 for l in f if l.strip())


def _record_month(raw_dir: str, prefix: str, ym: str, info: dict) -> None:
    """달마다 서버 totalCount·받은 건수·저장 건수·쪽수·받은 날을 <prefix>months.json 에 남긴다(나중에 API 없이 점검할 근거).
    용역처럼 캐시를 비우고 걸러서 저장하는 수집은 이 기록이 없으면 '짧게 받은 달'을 사후에 알 길이 없다."""
    path = os.path.join(raw_dir, f"{prefix}months.json")
    try:
        meta = json.load(open(path, encoding="utf-8")) if os.path.exists(path) else {}
    except ValueError:
        meta = {}
    if not isinstance(meta, dict):
        meta = {}
    meta[ym] = info
    with open(path + ".tmp", "w", encoding="utf-8") as f:
        json.dump(meta, f, ensure_ascii=False, indent=0, sort_keys=True)
    os.replace(path + ".tmp", path)


def api_failure_text(e: ApiError) -> str:
    """probe 용 실패 안내: 코드 + 설명 + 서버가 준 문구(resultMsg/응답 앞부분)를 모두 보여 준다."""
    return f"API 오류 {e.code} ({explain_code(e.code)}) — 서버 응답: {str(e.msg)[:300] or '(문구 없음)'}"


def _sample(client: G2BClient, op: str, params: dict, wins: List[tuple], rows: int = 5):
    """probe 용 표본 호출: 그 달의 구간(wins)을 차례로 조회해 항목이 나오는 첫 구간의 응답을 돌려준다 → (body, (bgn, end), 오류 또는 None).
    구간이 정말 0건이면 다음 구간으로 넘어간다. 오류(ApiError)는 그 구간에서 멈추고 돌려준다."""
    body, win = {}, wins[0]
    for win in wins:
        try:
            body = client.call(op, {**params, "inqryBgnDt": win[0], "inqryEndDt": win[1], "pageNo": 1, "numOfRows": rows}, use_cache=False)
        except ApiError as e:
            return {}, win, e
        if client._items(body):
            break
    return body, win, None


def _read_done(path: str) -> set:
    """완료한 달(YYYYMM) 집합. 손으로 고친 파일에 BOM·UTF-16·메모가 섞여 있어도 달만 읽는다
    (메모장·PowerShell 로 저장하면 BOM 이 붙는다 — 그대로 읽으면 첫 달을 완료하지 않은 달로 보고 다시 받게 된다)."""
    if not os.path.exists(path):
        return set()
    raw = open(path, "rb").read()
    text = raw.decode("utf-16", "replace") if raw[:2] in (b"\xff\xfe", b"\xfe\xff") else raw.decode("utf-8-sig", "replace")
    return {t for t in text.replace("\ufeff", " ").split() if re.fullmatch(r"(?:19|20)\d{2}(?:0[1-9]|1[0-2])", t)}


def _write_done(path: str, done: set) -> None:
    os.makedirs(os.path.dirname(path), exist_ok=True)
    with open(path, "w", encoding="utf-8") as f:
        f.write("\n".join(sorted(done)))


def _collect_monthly(cfg: dict, op_key: str, prefix: str, start: Optional[str], end: Optional[str],
                     base_url: Optional[str] = None, op_name: Optional[str] = None, extra: Optional[dict] = None,
                     row_filter=None, purge_cache: bool = False) -> pd.DataFrame:
    """월 단위 전량 수집 공통 루틴. 반환: 누적 원문 DataFrame(문자열화).
    한 달을 api.max_window_days(기본 7일) 이하의 구간으로 나눠 조회한다(2026-11-10 부터 서버가 7일 넘는 기간을 07 오류로 거절).
    구간 하나라도 실패하면 그 달은 실패 — 파일을 바꾸지 않고 완료 표시도 하지 않는다(받은 구간은 캐시에 남아 다음 실행 때 이어받는다).
    row_filter: 저장할 행만 고르는 함수(용역 공고처럼 대부분이 무관한 경우). purge_cache: 완료한 달의 응답 캐시를 비움."""
    client = make_client(cfg, base_url)
    op = op_name or cfg["api"]["ops"][op_key]
    raw_dir = cfg["paths"]["raw_dir"]
    start = start or cfg["period"]["start"]
    end = end or cfg["period"]["end"]
    done_marker = os.path.join(raw_dir, f"{prefix}done.txt")
    done = _read_done(done_marker)
    today_ym = date.today().strftime("%Y%m")
    today_ymd = date.today().strftime("%Y%m%d")
    skipped_future, collected, failed = [], [], []
    consecutive_fail = 0
    stopped = False
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
            if refresh and purge_cache:
                # 다시 받는 달은 전날까지 받아 둔 사본을 먼저 지운다(진행 중인 달을 날마다 받으면 날짜별 사본이 쌓임). 오늘 받은 페이지는 남겨 이어받기에 쓴다
                client.purge_cache(op, f'"inqryBgnDt": "{ym}', before=date.today().isoformat())
            stats: dict = {"total": 0, "pages": 0, "windows": 0}
            rows: list = []
            n_all = n_dup = 0
            prev_keys: dict = {}
            try:
                # require_total: totalCount 없는 응답은 오류. 받은 건수 < totalCount 이면 IncompleteResult(RuntimeError) → 아래에서 '실패한 달'
                for wb, we in G2BClient.split_window(bgn, endd, client.cfg.max_window_days):
                    if wb[:8] > today_ymd:              # 진행 중인 달의 아직 오지 않은 구간
                        continue
                    st: dict = {}
                    got = list(client.iter_all(op, {**(extra or {"inqryDiv": "1"}), "inqryBgnDt": wb, "inqryEndDt": we},
                                               cache_salt=(date.today().isoformat() if refresh else None), require_total=True, stats=st))
                    # 구간 경계의 1분은 앞 구간과 이 구간에 걸쳐 조회하므로 같은 행이 두 번 올 수 있다 → 바로 앞 구간이 이미 준 행(똑같은 행)만 뺀다.
                    # 한 구간 안에서 서버가 똑같은 행을 두 번 준 것은 서버 건수 그대로 남긴다(월 단위로 받던 때와 같게)
                    keys = [hashlib.sha1(json.dumps(r, sort_keys=True, ensure_ascii=False).encode("utf-8")).digest() for r in got]
                    left = dict(prev_keys)
                    for r, k in zip(got, keys):
                        if left.get(k, 0) > 0:
                            left[k] -= 1
                            n_dup += 1
                        else:
                            rows.append(r)
                    prev_keys = {}
                    for k in keys:
                        prev_keys[k] = prev_keys.get(k, 0) + 1
                    n_all += len(got)
                    stats["total"] += st.get("total") or 0
                    stats["pages"] += st.get("pages") or 0
                    stats["windows"] += 1
            except (ApiError, RuntimeError) as e:
                if isinstance(e, DailyBudgetExceeded) or (isinstance(e, ApiError) and e.code in ("10", "11", "12", "20", "30", "31", "32", "33")):
                    raise
                log.warning("월 %s 수집 실패(건너뜀, 완료 표시 안 함, 기존 파일 유지): %s", ym, e)
                failed.append(ym)
                consecutive_fail += 1
                if consecutive_fail >= 3:
                    # 같은 종류의 실패가 이어지면 남은 달마다 호출을 태우지 않고 멈춘다(일일 예산 보호)
                    log.error("%s: 연속 %d개월 수집 실패 → 중단합니다(그 뒤의 달은 시도하지 않았습니다). 위 오류 문구를 확인한 뒤 같은 명령으로 재개하세요.", op, consecutive_fail)
                    stopped = True
                    break
                continue
            consecutive_fail = 0
            if row_filter is not None:
                rows = [r for r in rows if row_filter(r)]
            n_old = _count_lines(path) if os.path.exists(path) else None
            tmp_path = path + ".tmp"
            if os.path.exists(tmp_path):
                os.remove(tmp_path)
            n = _append_jsonl(tmp_path, rows)
            if n_old is not None and n < n_old:
                # 다시 받은 결과가 기존보다 적으면 기존 파일을 .prev 로 남긴다(원문을 조용히 잃지 않게)
                prev, k = path + ".prev", 1
                while os.path.exists(prev):          # 앞서 남긴 .prev 를 덮어쓰지 않는다(두 번째로 줄어들 때 더 큰 원본이 사라지지 않게) → .prev2, .prev3 …
                    k += 1
                    prev = f"{path}.prev{k}"
                os.replace(path, prev)
                log.warning("%s %s: 다시 받은 결과(%d건)가 기존 파일(%d건)보다 적습니다 → 기존 파일을 %s 로 남김", op, ym, n, n_old, os.path.basename(prev))
            os.replace(tmp_path, path)
            # total = 구간별 서버 건수의 합(경계에서 겹친 행은 두 번 셈), received = 받은 행 수, dup = 그중 겹쳐서 뺀 행 수, saved = 파일에 쓴 행 수
            _record_month(raw_dir, prefix, ym, {"total": stats.get("total"), "received": n_all, "dup": n_dup, "saved": n, "pages": stats.get("pages"),
                                                "windows": stats.get("windows"), "fetched": date.today().isoformat(), "current": bool(is_current)})
            if n_all == 0 and not is_current:
                log.warning("%s %s: 서버가 0건(자료 없음)이라고 답했습니다 — 이 달에 정말 공고가 없는지 확인하세요(공사 공고는 매달 수천 건)", op, ym)
            log.info("%s %s: %s%d건%s%s", op, ym, f"전체 {n_all - n_dup:,}건 중 보존 " if row_filter is not None else "", n,
                     f" (구간 경계에서 겹친 {n_dup}건은 한 번만 저장)" if n_dup else "", " (진행 중인 달 — 완료 표시 안 함)" if is_current else "")
            collected.append(ym)
            if not is_current:
                done.add(ym)
                _write_done(done_marker, done)
                if purge_cache:
                    client.purge_cache(op, f'"inqryBgnDt": "{ym}')
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
        log.warning("[미완료] %s: 완료 %d개월, 미완료 %d개월 → %s%s  (%s)", op, len(done), len(remaining),
                    ", ".join(remaining[:12]), " …" if len(remaining) > 12 else "",
                    "연속 실패로 중단 — 위 오류 문구를 해결한 뒤 같은 명령으로 재개" if stopped else "내일 같은 명령으로 재개")
    else:
        log.info("[완료] %s: %d개월 수집 완료", op, len(done))
    df = _load_jsonl_dir(raw_dir, prefix)
    return _stringify(df).copy() if not df.empty else df


def incomplete_months(cfg: dict, prefix: str, period: Optional[tuple] = None) -> List[str]:
    """done.txt 기준으로 아직 수집되지 않은(오늘 이전) 월 목록 — discover·link-servc 단계에서 경고용. period = (시작월, 끝월). 없으면 공사 수집 기간."""
    done = _read_done(os.path.join(cfg["paths"]["raw_dir"], f"{prefix}done.txt"))
    today_ym = date.today().strftime("%Y%m")
    start, end = period or (cfg["period"]["start"], cfg["period"]["end"])
    return [b[:6] for b, _ in G2BClient.month_ranges(start, end) if b[:6] < today_ym and b[:6] not in done]


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
SERVC_KEEP_NAME = (r"폐기물|건설사업관리|감리|설계\s*용역|설계\s*(?:제안\s*)?공모|제안\s*공모|설계안\s*공모|설계\s*경기|현상\s*설계|건축\s*설계|기획\s*설계|"
                   r"실시\s*설계|기본\s*설계|기본\s*및\s*실시|CM\s*용역")
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
    공사 수집과 같은 서비스·같은 인증키. 호출량이 크므로(월 약 20~30회, 전 기간 약 2,000~2,900회) 일일 예산 안에서 여러 날에 걸쳐 이어받는다."""
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


_STATUS_KO = {"ok": "정상", "nodata": "정상(0건)", "error": "서버가 요청을 거절", "unreadable": "읽을 수 없는 응답", "temporary": "일시 오류",
              "quota": "일일 한도", "network": "연결 실패"}


def _diag_line(label: str, d: dict) -> str:
    """진단 호출 1회의 결과를 한 줄로: HTTP 상태, 판정, 코드·서버 문구, totalCount, 표본 수(읽을 수 없는 응답이면 본문 앞부분도)."""
    code = d.get("code")
    line = f"  {label}: HTTP {d.get('http') or '-'} | {_STATUS_KO.get(d.get('status'), d.get('status'))}"
    if code and code != "00":
        line += f" | 코드 {code} ({explain_code(code)})"
    if d.get("msg"):
        line += f" | {'서버 문구' if d.get('status') == 'error' else '사유'}: {d['msg']}"
    if d.get("status") in ("ok", "nodata"):
        line += f" | totalCount={d.get('total')} | 표본 {d.get('n')}건"
    if d.get("status") in ("unreadable", "temporary") and d.get("snippet"):
        line += f"\n      최상위 키 {d.get('keys')} | 본문 앞부분: {d.get('snippet')}"
    return line


def _probe_servc_diagnosis(cfg: dict, client: G2BClient, op: str, base: dict, first: dict, ym: str, month: tuple) -> None:
    """용역 목록 조회가 표본을 돌려주지 않았을 때: 조회 조건을 바꿔 가며(기간·건수·다른 오퍼레이션) 서버의 답을 그대로 출력한다(5회 안팎 호출)."""
    f = cfg["fields"]
    ops = cfg["api"].get("ops") or {}
    bgn, end = base["inqryBgnDt"], base["inqryEndDt"]
    days = int(end[6:8]) - int(bgn[6:8]) + 1
    if first.get("status") == "quota":                  # 일일 한도에 걸렸으면 조건을 바꿔 봐야 호출만 쓴다
        print(f"\n=== {op} ({ym}) 오늘의 호출 한도에 도달했습니다({first.get('msg')}) → 내일 같은 명령으로 다시 실행하세요.")
        return
    if first.get("status") == "nodata":                 # 그 달의 모든 구간이 0건: 조건을 바꿔 물어도 호출만 쓴다(답은 이미 나왔다)
        n_all = len(G2BClient.split_window(month[0], month[1], client.cfg.max_window_days, overlap=False))
        print(f"\n=== {op} ({ym}) 표본이 없습니다 — {month[0][:8]}~{month[1][:8]} 을 {n_all}개 구간으로 모두 조회했습니다 ===")
        print(_diag_line(f"마지막 구간 {bgn[:8]}~{end[:8]}({days}일)", first))
        print("[판단]\n  서버가 이 기간은 0건이라고 답했습니다(오류가 아님) → 다른 달(--ym)로 다시 시도하세요.")
        print(f"\n오늘 API 호출 {client.calls_today()}회 (예산 {cfg['api']['daily_call_budget']})")
        return
    print(f"\n=== {op} ({ym}) 표본을 받지 못했습니다 — 조회 조건을 바꿔 가며 서버의 답을 확인합니다 ===")
    print(_diag_line(f"① {bgn[:8]}~{end[:8]}({days}일), 50건", first))
    tries = [("② 첫 1일, 50건", op, {**base, "inqryEndDt": bgn[:8] + "2359"}),
             ("③ 한 달 전체, 50건", op, {**base, "inqryBgnDt": month[0], "inqryEndDt": month[1]}),
             (f"④ {days}일, 10건", op, {**base, "numOfRows": 10})]
    if ops.get("servc_search"):
        tries.append((f"⑤ 용역 검색(공고명 '폐기물'), {days}일", ops["servc_search"], {**base, f["name"]: "폐기물", "numOfRows": 5}))
    if ops.get("cnstwk_list"):
        tries.append((f"⑥ 대조: 공사 공고 목록, {days}일", ops["cnstwk_list"], {**base, "numOfRows": 5}))
    got = {}
    for label, o, params in tries:
        got[label[0]] = d = client.diagnose(o, params)
        print(_diag_line(label, d))

    def accepted(k) -> bool:                             # 서버가 요청을 받아 준 경우(항목이 있든 0건이든) — 거절·읽을 수 없는 응답과 구별
        return (got.get(k) or {}).get("status") in ("ok", "nodata")
    rejected = first.get("status") == "error"           # 서버가 코드를 붙여 거절했을 때만 '조회 조건' 문제로 본다(일시 오류·읽을 수 없는 응답은 조건 탓이 아닐 수 있다)
    quota = any(d.get("status") == "quota" for d in got.values())
    print("[판단]")
    if first.get("status") == "nodata":
        print("  서버가 이 기간은 0건이라고 답했습니다(오류가 아님) → 다른 달(--ym)로 다시 시도하세요.")
    elif rejected and accepted("②") and not accepted("④"):
        print(f"  하루는 조회되는데 {days}일은 안 됩니다 → 한 번에 조회할 수 있는 기간이 {days}일보다 짧습니다. config.yaml api.max_window_days 를 줄여야 합니다(이 출력을 알려 주세요).")
    elif rejected and accepted("④"):
        print("  건수를 줄이면 조회됩니다 → 한 번에 받는 건수(numOfRows) 제한 문제입니다(이 출력을 알려 주세요).")
    elif quota:
        print("  조건을 바꿔 보는 도중 오늘의 호출 한도에 도달해 원인을 가리지 못했습니다 → 내일 같은 명령으로 다시 실행하세요.")
    elif accepted("④"):
        print(f"  같은 {days}일을 다시 물으니 응답이 왔습니다 → 처음 조회의 실패는 일시적인 것으로 보입니다(조회 기간 문제가 아님). 잠시 뒤 같은 명령을 다시 실행하세요.")
    elif accepted("②"):
        print(f"  하루는 조회되는데 {days}일 조회는 서버가 코드 없이 실패했습니다(일시 오류·읽을 수 없는 응답) → 잠시 뒤 같은 명령을 다시 실행해 보고, 되풀이되면 이 출력을 알려 주세요.")
    elif accepted("⑥"):
        print("  공사 공고는 조회되는데 용역 공고만 안 됩니다 → 위 ①~⑤ 의 코드·서버 문구가 원인입니다. 코드 12·20·30 이면 공공데이터포털 마이페이지에서 "
              "「조달청_나라장터 입찰공고정보서비스」 활용신청 상세기능에 용역 조회가 들어 있는지, 오퍼레이션명(config api.ops.servc_list)이 맞는지 확인하고, "
              "그 밖의 코드면 이 출력을 그대로 알려 주세요.")
    else:
        print("  공사 공고(대조)도 조회되지 않습니다 → 서비스·네트워크·인증키 쪽 문제이거나 조회 조건이 또 바뀐 것입니다. "
              "python -m g2b_costdb.pipeline doctor 를 실행해 보고, 이 출력을 그대로 알려 주세요.")
    print(f"\n오늘 API 호출 {client.calls_today()}회 (예산 {cfg['api']['daily_call_budget']})")


def probe_servc(cfg: dict, ym: str) -> None:
    """용역 입찰공고 오퍼레이션 시험 호출(3~4회): 응답 필드·표본, config 매핑 대조, 표본 기간의 건수로 전 기간 호출 수 추정, 공고명 검색(PPSSrch) 동작 확인.
    표본은 그 달의 첫 구간(api.max_window_days, 기본 7일)으로 받는다 — 2026-11-10 부터 서버가 7일 넘는 기간을 07 오류로 거절한다.
    표본을 받지 못하면 조회 조건을 바꿔 가며 서버의 답(코드·문구)을 그대로 보여 준다('항목 없음' 으로 뭉뚱그리지 않는다)."""
    client = make_client(cfg)
    month = G2BClient.month_ranges(ym, ym)[0]
    wins = G2BClient.split_window(month[0], month[1], client.cfg.max_window_days, overlap=False)   # 표본용(겹치지 않는 구간)
    n_win = len(G2BClient.split_window(month[0], month[1], client.cfg.max_window_days))            # 수집이 실제로 쓰는 구간 수
    bgn, end = wins[0]
    days = int(end[6:8]) - int(bgn[6:8]) + 1
    f = cfg["fields"]
    ops = cfg["api"].get("ops") or {}
    op = ops.get("servc_list")
    if not op:
        print("config.yaml api.ops.servc_list 가 없습니다.")
        return
    which = ["bid_no", "bid_ord", "name", "kind", "notice_dt", "notice_inst", "demand_inst", "demand_inst_cd", "presmpt_price", "vat",
             "servc_budget", "servc_div", "servc_class", "servc_class_mid", "servc_class_lrg", "prespec_no", "detail_url",
             "attach_url_prefix", "attach_name_prefix"]
    base = {"inqryDiv": "1", "inqryBgnDt": bgn, "inqryEndDt": end, "pageNo": 1, "numOfRows": 50}
    first = client.diagnose(op, base)
    for wb, we in wins[1:]:                             # 첫 구간이 정말 0건이면(작은 기관·옛 자료) 그 달의 다음 구간에서 표본을 찾는다
        if first.get("status") != "nodata":
            break
        bgn, end = wb, we
        base = {**base, "inqryBgnDt": bgn, "inqryEndDt": end}
        first = client.diagnose(op, base)
    days = int(end[6:8]) - int(bgn[6:8]) + 1
    if first.get("status") != "ok":
        _probe_servc_diagnosis(cfg, client, op, base, first, ym, month)
        return
    body = first["body"]
    items = client._items(body)
    total = body.get("totalCount")
    print(f"\n=== {op} ({bgn[:8]}~{end[:8]}, {days}일) totalCount={total} 표본 {len(items)}건 ===")
    if n_win > 1:                                       # 한 달 전체 조회가 지금도 되는지(안 되면 왜인지) 한 줄로 남긴다
        whole = client.diagnose(op, {**base, "inqryBgnDt": month[0], "inqryEndDt": month[1], "numOfRows": 1})
        print(_diag_line(f"[조회 기간] 한 달 전체({month[0][:8]}~{month[1][:8]}) 조회", whole))
        print(f"  → 수집은 한 달을 {client.cfg.max_window_days}일 이하 {n_win}구간으로 나눠 조회합니다(config.yaml api.max_window_days).")
        cw = G2BClient.split_window(month[0], month[1], client.cfg.max_window_days)[0]       # 수집이 실제로 쓰는 구간(끝 = 다음 구간의 시작 분)도 서버가 받는지
        if cw != (bgn, end):
            print(_diag_line(f"[수집 구간] 수집에 쓰는 첫 구간({cw[0][:8]} {cw[0][8:10]}:{cw[0][10:]}~{cw[1][:8]} {cw[1][8:10]}:{cw[1][10:]}) 조회",
                             client.diagnose(op, {**base, "inqryBgnDt": cw[0], "inqryEndDt": cw[1], "numOfRows": 1})))
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
    # 보존 규칙을 넓힐지 판단할 자료: 표본의 조달분류명 분포와, 버려지는 공고 중 '설계'·'관리' 가 든 이름
    from collections import Counter
    ck = f.get("servc_class")
    if ck:
        top = Counter(str(it.get(ck) or "(없음)") for it in items).most_common(8)
        print("[조달분류 분포] " + ", ".join(f"{n} {c}" for n, c in top))
    lost = [str(it.get(f["name"]))[:50] for it in items if not keep(it) and re.search(r"설계|사업관리|감독|해체|철거", str(it.get(f["name"]) or ""))]
    if lost:
        print(f"[버려지는 공고 중 확인할 이름 {len(lost)}건] " + "; ".join(lost[:8]))
    try:
        n = int(float(total))
        pages = math.ceil(n / int(cfg["api"]["num_of_rows"]))
        s0, e0 = servc_period(cfg)
        months = len(G2BClient.month_ranges(s0, e0))
        per_month = pages * n_win
        print(f"[호출 수 추정] {bgn[:8]}~{end[:8]}({days}일) {n:,}건 = {pages}회 → 한 달({n_win}구간) 약 {per_month}회. "
              f"수집 기간 {s0}~{e0}({months}개월)을 같은 규모로 보면 약 {per_month * months:,}회 "
              f"→ 일 {cfg['api']['daily_call_budget']}회 예산으로 약 {math.ceil(per_month * months / int(cfg['api']['daily_call_budget']))}일(공사 수집과 예산을 함께 씀)")
    except (TypeError, ValueError):
        print("[호출 수 추정] totalCount 가 없어 추정하지 못했습니다.")
    sop = ops.get("servc_search")
    if sop:
        try:
            b2 = client.call(sop, {"inqryDiv": "1", "inqryBgnDt": bgn, "inqryEndDt": end, f["name"]: "폐기물", "pageNo": 1, "numOfRows": 5}, use_cache=False)
            it2 = client._items(b2)
            print(f"\n=== {sop} 공고명 '폐기물' ({bgn[:8]}~{end[:8]}) totalCount={b2.get('totalCount')} 표본 {len(it2)}건 ===")
            for it in it2[:5]:
                print("  -", str(it.get(f["name"]))[:70], "|", it.get(f["demand_inst"]), "| 추정가격", it.get(f["presmpt_price"]),
                      "|", it.get(f.get("servc_class", ""), ""))
            print("[검색] 공고명 검색이 동작하면 totalCount 가 위 전체 건수보다 훨씬 작습니다(같으면 검색 조건이 무시된 것).")
        except ApiError as e:
            print(f"\n=== {sop} 실패: {api_failure_text(e)} — 검색 오퍼레이션은 참고용이라 수집에는 영향 없음")
        except RuntimeError as e:
            print(f"\n=== {sop} 실패: {client._scrub(e)[:300]} — 검색 오퍼레이션은 참고용이라 수집에는 영향 없음")
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
    n_ok = n_err = 0
    try:
        for i, no in enumerate(bid_nos):
            try:
                body = client.call(op, license_query_params(cfg, str(no)))
            except ApiError as e:
                n_err += 1
                # 06·07·08 은 조달청 서비스의 요청 조건 오류(10·11·12 와 같은 성격). 그 밖의 오류도 한 건도 성공하지 못한 채 10건 이어지면 멈춘다(앞의 몇 건만 읽지 못한 경우에 뒤의 정상 공고까지 버리지 않게)
                # (오류 응답은 캐시에 넣지 않으므로, 멈추지 않으면 실행할 때마다 공고 수만큼 호출을 쓴다)
                if n_ok == 0 and (e.code in ("06", "07", "08", "10", "11", "12") or n_err >= 10):
                    log.error("면허제한 조회 파라미터가 맞지 않는 것으로 보입니다(%s). config.yaml의 api.license_query 를 probe 결과에 맞게 조정하세요.", e)
                    break
                log.warning("면허제한 조회 실패 %s: %s", no, e)
                continue
            except RuntimeError as e:                   # 재시도 끝에 실패한 한 건 때문에 전체를 멈추지 않는다(DailyBudgetExceeded 는 바깥에서 처리)
                if isinstance(e, DailyBudgetExceeded):
                    raise
                n_err += 1
                if n_ok == 0 and n_err >= 3:            # 한 건도 성공하지 못한 채 재시도 끝 실패가 이어지면 서버·연결 문제 — 공고 수만큼 예산을 태우지 않는다
                    log.error("면허제한 조회가 계속 실패합니다(%s). 서버·네트워크 상태를 확인한 뒤 research 를 다시 실행하세요.", client._scrub(e)[:200])
                    break
                log.warning("면허제한 조회 실패 %s: %s", no, client._scrub(e)[:200])
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
    """그 달 첫 구간(api.max_window_days, 기본 7일)의 첫 페이지만 호출하여 필드명·표본을 출력하고 config.yaml fields 매핑을 대조.
    (2026-11-10 부터 서버가 7일 넘는 기간을 07 오류로 거절하므로 한 달 전체로는 조회하지 않는다)"""
    client = make_client(cfg)
    month = G2BClient.month_ranges(ym, ym)[0]
    wins = G2BClient.split_window(month[0], month[1], client.cfg.max_window_days, overlap=False)
    bgn, end = wins[0]
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
        body, (bgn, end), err = _sample(client, op, {"inqryDiv": "1"}, wins)
        span = f"{bgn[:8]}~{end[:8]}"
        if err is not None:
            print(f"\n=== {op} ({span}) 실패: {api_failure_text(err)}")
            if err.code in ("20", "30", "31", "32"):
                print(f"  → {explain_code(err.code)}")
            elif err.code in ("06", "07", "08"):
                print(f"  → 조회 조건을 서버가 거절했습니다. python -m g2b_costdb.pipeline probe --servc --ym {ym} 가 조건을 바꿔 가며 원인을 보여 줍니다.")
            continue
        items = client._items(body)
        print(f"\n=== {op} ({span}) totalCount={body.get('totalCount')} 표본 {len(items)}건 ===")
        if not items:
            print(f"서버가 {ym} 의 모든 구간이 0건이라고 답했습니다(totalCount=0) — 다른 달(--ym)로 다시 시도하세요.")
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
        lic_errors = 0
        for v in variants:
            params = dict(v); params[f["bid_no"]] = first_bid_no; params.setdefault("pageNo", "1"); params.setdefault("numOfRows", "100")
            try:
                body = client.call(op, params, use_cache=False)
            except ApiError as e:
                print(f"\n=== {op} params={v} 실패: {api_failure_text(e)}")
                lic_errors += 1
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
        elif lic_errors == len(variants):
            print("[면허제한] 모든 조회구분이 오류로 거절됨('면허제한 없음'이 아님) — 위 서버 문구로 필수 파라미터를 확인. 확인 전까지 use_license_limit 는 false 유지")
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
    month = G2BClient.month_ranges(ym, ym)[0]
    wins = G2BClient.split_window(month[0], month[1], client.cfg.max_window_days, overlap=False)             # 조회 기간 제한(기본 7일) 안에서 시험
    variants = [dict(a.get("award_query") or {"inqryDiv": "1"}), {"inqryDiv": "1"}, {"inqryDiv": "2"}, {"inqryDiv": "3"}]
    seen = []
    ok = None
    n_err = 0
    for v in variants:
        if v in seen:
            continue
        seen.append(v)
        body, (bgn, end), e = _sample(client, op, dict(v), wins)
        if e is not None:
            print(f"\n=== {op} params={v} 실패: {api_failure_text(e)}")
            n_err += 1
            if e.code in ("12", "20", "30", "31", "32"):
                print("→ 「조달청_나라장터 낙찰정보서비스」 활용신청이 승인되었는지, api.award_base_url 이 포털 페이지의 End Point 와 같은지 확인")
                return
            continue
        items = client._items(body)
        print(f"\n=== {op} params={v} ({bgn[:8]}~{end[:8]}) totalCount={body.get('totalCount')} 표본 {len(items)}건 ===")
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
    elif n_err == len(seen):
        print("[낙찰정보] 모든 조회구분이 오류로 거절됨('자료 없음'이 아님) — 위 서버 문구와 활용가이드의 기간 조회 파라미터를 확인")
    else:
        print("[낙찰정보] 어떤 조회구분으로도 항목이 없음 — 활용가이드의 기간 조회 파라미터를 확인")
