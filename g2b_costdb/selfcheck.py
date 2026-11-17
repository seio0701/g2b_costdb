"""이미 수집한 원문에 '비었거나 짧게 저장되고 완료 표시된 달'이 있는지 점검한다 (API 호출 없음, 인증키를 읽지 않음).

    python -m g2b_costdb.pipeline check-raw            # 점검만
    python -m g2b_costdb.pipeline check-raw --reopen   # 다시 받아야 할 달을 done.txt 에서 빼 둔다(원문 파일은 그대로) → 이어서 collect

배경: 2026-11-17 이전 버전의 수집기는 서버의 오류 응답(HTTP 200 에 'response' 가 없는 모양, 예: 조회 기간 초과 07)이나
잘린 쪽(totalCount>0 인데 항목 0건)을 '자료 없음'으로 읽어, 그 달을 0행 또는 쪽 경계까지만 저장하고 완료 표시했다.

보는 것(완료 표시된 달에 대해)
  1) data/raw/<prefix>YYYYMM.jsonl 의 줄 수(= 저장된 행 수)와 <prefix>done.txt
  2) 새 수집기가 달마다 남기는 기록 <prefix>months.json (서버 건수·받은 건수·저장 건수)
  3) data/cache.sqlite 의 응답 캐시 중 수집에 쓴 쪽 크기(config api.num_of_rows)의 행: 달·쪽별 totalCount / 항목 수 / 읽지 못한 본문({} 등)
     — 한 달을 통째로 조회하던 예전 수집분은 이것으로 서버 건수를 대조한다
표시
  EMPTY_DONE         완료 표시된 달인데 0행(서버가 0건이라고 답한 기록이 있으면 표시하지 않음)
  SHORT_VS_TOTAL     저장 행 수(또는 받은 건수) < 서버가 알려 준 건수
  SHORT_VS_SAVED     저장 행 수 < 수집 때 저장했다고 기록한 건수(파일이 줄어듦)
  BAD_CACHE_BODY     서버 건수와 맞는지 확인되지 않은 달의 캐시에 읽지 못한 본문({} 등)·잘린 쪽이 있음
  CACHE_SUM_LT_TOTAL 서버 건수와 맞는지 확인되지 않은 달의 캐시에 든 쪽들의 항목 합 < totalCount (쪽이 빠짐)
  PAGE_MULTIPLE      저장 행 수가 쪽 크기의 정확한 배수 — 쪽 경계에서 끊긴 흔적(서버 건수와 맞는 것이 확인되면 표시 안 함)
  DONE_NO_FILE       완료 표시인데 파일이 없음
  LOW_VS_NEIGHBOURS  앞뒤 6개월 중앙값의 절반 미만                  RATIO_OUTLIER  (기초금액÷공고) 비율이 전체 중앙값에서 크게 벗어남
  NOT_DONE           파일은 있는데 완료 표시가 없음(진행 중이던 달)    NEVER_COLLECTED 기간 안인데 파일도 완료 표시도 없음
앞의 일곱(HARD)은 다시 받아야 하는 달, 뒤의 넷(SOFT)은 눈으로 확인할 달. 서버 건수와 맞는 것이 확인된 달에는 LOW_VS_NEIGHBOURS·RATIO_OUTLIER 를 붙이지 않는다.
"""
from __future__ import annotations

import calendar
import json
import os
import re
import sqlite3
import statistics
from collections import defaultdict
from datetime import date
from typing import Dict, List, Optional, Tuple

DEFAULT_OPS = {"notices_": "getBidPblancListInfoCnstwk", "bsis_": "getBidPblancListInfoCnstwkBsisAmount",
               "award_": "getScsbidListSttusCnstwk", "servc_": "getBidPblancListInfoServc"}
FILTERED = {"servc_"}            # 보존 규칙으로 걸러 저장하는 수집(줄 수 ≠ totalCount 가 정상)
COLLECT_CMD = {"servc_": "collect --servc"}        # 그 밖의 원문은 collect
HARD = ("EMPTY_DONE", "SHORT_VS_TOTAL", "SHORT_VS_SAVED", "BAD_CACHE_BODY", "CACHE_SUM_LT_TOTAL", "PAGE_MULTIPLE", "DONE_NO_FILE")
SOFT = ("LOW_VS_NEIGHBOURS", "RATIO_OUTLIER", "NOT_DONE", "NEVER_COLLECTED")
FLAG_KO = {"EMPTY_DONE": "빈 달인데 완료 표시", "SHORT_VS_TOTAL": "서버 건수보다 적게 저장", "SHORT_VS_SAVED": "수집 때 저장한 건수보다 줄 수가 적음",
           "BAD_CACHE_BODY": "읽지 못한 응답이 저장돼 있음", "CACHE_SUM_LT_TOTAL": "쪽이 빠짐", "PAGE_MULTIPLE": "쪽 경계에서 끊긴 흔적",
           "DONE_NO_FILE": "완료 표시인데 파일 없음", "LOW_VS_NEIGHBOURS": "앞뒤 달보다 크게 적음", "RATIO_OUTLIER": "기초금액÷공고 비율이 다른 달과 다름",
           "NOT_DONE": "완료 표시 없음(부분 수집)", "NEVER_COLLECTED": "받은 적 없음"}
_BAD_KINDS = ("NON_DICT", "EMPTY_NO_TOTAL", "TRUNCATED")
_YM = re.compile(r"(?:19|20)\d{2}(?:0[1-9]|1[0-2])")


def read_done(path: str) -> Tuple[set, List[str]]:
    """<prefix>done.txt → (달 집합, 달이 아닌 토막). 손으로 고친 파일(메모장·PowerShell 이 붙인 BOM, UTF-16, 메모 글자)에도 멈추거나 죽지 않는다."""
    if not os.path.exists(path):
        return set(), []
    raw = open(path, "rb").read()
    text = raw.decode("utf-16", "replace") if raw[:2] in (b"\xff\xfe", b"\xfe\xff") else raw.decode("utf-8-sig", "replace")
    toks = text.replace("﻿", " ").split()
    good = {t for t in toks if _YM.fullmatch(t)}
    return good, sorted(set(toks) - good)


def count_lines(path: str) -> int:
    n, last = 0, b"\n"
    with open(path, "rb") as f:
        while True:
            chunk = f.read(4 << 20)
            if not chunk:
                break
            n += chunk.count(b"\n")
            last = chunk[-1:]
    return n + (0 if last == b"\n" else 1)


def month_seq(a: str, b: str) -> List[str]:
    if not (_YM.fullmatch(a or "") and _YM.fullmatch(b or "")):       # 달(YYYYMM)이 아닌 값이 끝점이면 끝없이 돌게 된다
        return []
    y, m = int(a[:4]), int(a[4:6])
    out = []
    while f"{y:04d}{m:02d}" <= b:
        out.append(f"{y:04d}{m:02d}")
        m += 1
        if m > 12:
            y, m = y + 1, 1
    return out


def scan_files(raw_dir: str, prefix: str):
    files = {}
    if os.path.isdir(raw_dir):
        for fn in os.listdir(raw_dir):
            m = re.fullmatch(re.escape(prefix) + r"(\d{6})\.jsonl", fn)
            if m and _YM.fullmatch(m.group(1)):
                files[m.group(1)] = count_lines(os.path.join(raw_dir, fn))
    done, junk = read_done(os.path.join(raw_dir, f"{prefix}done.txt"))
    mp = os.path.join(raw_dir, f"{prefix}months.json")
    try:
        meta = json.load(open(mp, encoding="utf-8")) if os.path.exists(mp) else {}
    except (ValueError, OSError):
        meta = {}
    return files, done, (meta if isinstance(meta, dict) else {}), junk


def _kind(is_dict: bool, total, n: int, page_no: int, rows: int) -> Tuple[str, Optional[int], int]:
    """캐시 본문 요약 → (종류, totalCount 또는 None, 항목 수)."""
    if not is_dict:
        return "NON_DICT", None, 0
    try:
        total = int(float(total))
    except (TypeError, ValueError, OverflowError):
        total = None
    if n:
        return ("DATA" if total is not None else "DATA_NO_TOTAL"), total, n
    if total is None or total < 0:
        return "EMPTY_NO_TOTAL", None, 0          # {} 등: 해석하지 못한 응답이 '자료 없음'으로 저장된 것
    if total == 0:
        return "NODATA", 0, 0                     # 서버가 0건이라고 명시(03 또는 totalCount 0)
    if rows and (page_no - 1) * rows >= total:
        return "PAST_END", total, 0               # 마지막 쪽 다음 쪽(정상)
    return "TRUNCATED", total, 0                  # totalCount>0 인데 항목 0건


def body_kind(body, page_no: int, rows: int) -> Tuple[str, Optional[int], int]:
    """파이썬으로 해석한 본문 → (종류, totalCount, 항목 수)."""
    if not isinstance(body, dict):
        return _kind(False, None, 0, page_no, rows)
    items = body.get("items")
    if isinstance(items, dict):
        items = items.get("item", [])
        if isinstance(items, dict):
            items = [items]
    n = len(items) if isinstance(items, list) else 0
    return _kind(True, body.get("totalCount"), n, page_no, rows)


_FAST_SQL = """
SELECT op, params, fetched_at,
       CASE WHEN json_valid(body) THEN json_type(body) END,
       CASE WHEN json_valid(body) AND json_type(body) = 'object' THEN json_extract(body, '$.totalCount') END,
       CASE WHEN json_valid(body) AND json_type(body) = 'object' THEN
            CASE json_type(body, '$.items')
                 WHEN 'array' THEN json_array_length(body, '$.items')
                 WHEN 'object' THEN CASE json_type(body, '$.items.item') WHEN 'array' THEN json_array_length(body, '$.items.item')
                                                                         WHEN 'object' THEN 1 ELSE 0 END
                 ELSE 0 END
            ELSE 0 END,
       substr(body, 1, 80),
       CASE WHEN json_valid(body) THEN NULL ELSE body END
FROM cache"""


def _py_row(body):
    """본문 한 건을 파이썬으로 해석 → (객체인지, totalCount, 항목 수, 본문 앞부분)."""
    try:
        b = json.loads(body)
    except (TypeError, ValueError, RecursionError):
        b = None
    _, _, n = body_kind(b, 1, 0)
    head = body.decode("utf-8", "replace") if isinstance(body, bytes) else (body or "")
    return isinstance(b, dict), (b.get("totalCount") if isinstance(b, dict) else None), n, str(head)[:80]


def _cache_rows(con):
    """캐시의 모든 행 → (op, params, fetched_at, 본문이 객체인지, totalCount, 항목 수, 본문 앞부분). 본문은 SQLite 안에서 요약해 읽는다(쪽 하나가 수 MB)."""
    try:
        for op, params, fetched, typ, total, n, head, raw in con.execute(_FAST_SQL):
            if raw is not None:                   # SQLite 가 JSON 으로 인정하지 않는 본문 → 아래 느린 길과 같은 방식으로 판정(두 길의 결과가 같게)
                yield (op, params, fetched) + _py_row(raw)
            else:
                yield op, params, fetched, typ == "object", total, int(n or 0), head or ""
        return
    except sqlite3.OperationalError as e:         # JSON 함수가 없는 SQLite → 파이썬으로 해석(느림). 표가 없거나 잠긴 경우는 그대로 올린다
        if "no such function" not in str(e) and "json" not in str(e).lower():
            raise
    for op, params, fetched, body in con.execute("SELECT op, params, fetched_at, body FROM cache"):
        yield (op, params, fetched) + _py_row(body)


def _open_ro(cache_db: str):
    import pathlib
    try:
        return sqlite3.connect(pathlib.Path(cache_db).resolve().as_uri() + "?mode=ro", uri=True, timeout=5)     # 읽기 전용
    except (sqlite3.Error, ValueError):
        return sqlite3.connect(cache_db, timeout=5)


def scan_cache(cache_db: str, ops, page_size: Optional[int] = None) -> Tuple[Dict, List[dict], int]:
    """→ (months[(op, ym)] = {...}, bad_rows, 읽은 행 수). 달 판정에는 수집에 쓴 쪽 크기(page_size)의 행만 쓴다
    (probe 가 남긴 5건·50건짜리 행이나 다른 쪽 크기로 받다 만 행이 섞이지 않게). 한 달이 여러 조회 구간으로 나뉘어 있어도 달 단위로 모은다."""
    con = _open_ro(cache_db)
    bad_rows: List[dict] = []
    wins = defaultdict(lambda: defaultdict(dict))        # (op, ym) -> {(bgn, end): {pageNo: (fetched, kind, total, n)}}
    n_rows = 0
    try:
        for op, params, fetched, is_dict, total, n, head in _cache_rows(con):
            n_rows += 1
            try:
                p = json.loads(params)
            except (TypeError, ValueError):
                p = {}
            if not isinstance(p, dict):
                p = {}
            bgn, end = str(p.get("inqryBgnDt") or ""), str(p.get("inqryEndDt") or "")
            ym = bgn[:6]
            try:
                page_no, rows = int(p.get("pageNo") or 1), int(p.get("numOfRows") or 0)
            except (TypeError, ValueError):
                page_no, rows = 1, 0
            kind, total, n = _kind(is_dict, total, n, page_no, rows)
            if kind in _BAD_KINDS:
                bad_rows.append({"op": op, "ym": ym, "bgn": bgn[:8], "pageNo": page_no, "numOfRows": rows, "kind": kind, "fetched_at": fetched,
                                 "params": {k: v for k, v in p.items() if k not in ("inqryBgnDt", "inqryEndDt", "pageNo", "numOfRows")},
                                 "body": head})
            if (op in ops and _YM.fullmatch(ym) and re.fullmatch(r"\d{12}", bgn) and re.fullmatch(r"\d{12}", end) and end[:6] == ym
                    and (not page_size or rows == page_size)):
                cur = wins[(op, ym)][(bgn, end)].get(page_no)
                if cur is None or str(fetched) > str(cur[0]):          # 같은 쪽을 여러 날 받았으면 가장 최근 것
                    wins[(op, ym)][(bgn, end)][page_no] = (fetched, kind, total, n)
    finally:
        con.close()
    months = {}
    for (op, ym), w in wins.items():
        last_day = calendar.monthrange(int(ym[:4]), int(ym[4:6]))[1]
        whole = {k: v for k, v in w.items() if k[0][6:8] == "01" and k[1][6:8] == f"{last_day:02d}"}
        part = {k: v for k, v in w.items() if k not in whole}
        if whole and part:                                         # 한 달 통째로 받은 것과 구간으로 나눠 받은 것이 함께 있으면 나중에 받은 쪽
            newest = lambda d: max(str(v[0]) for pg in d.values() for v in pg.values())
            use = part if newest(part) > newest(whole) else whole
        else:
            use = whole or part
        total, tot_known, s, bad, short, n_pages, fetch = 0, True, 0, [], [], 0, ""
        for (b, e), pg in sorted(use.items()):
            p1 = pg.get(1)
            totals = [v[2] for v in pg.values() if v[2] is not None]
            wt = p1[2] if p1 and p1[2] is not None else (max(totals) if totals else None)
            ws = sum(v[3] for v in pg.values())
            if wt is None:
                tot_known = False
            else:
                total += wt
                if ws < wt:
                    short.append(b[6:8])
            s += ws
            bad += [f"{b[6:8]}일~:{k}쪽" for k, v in sorted(pg.items()) if v[1] in _BAD_KINDS]
            n_pages += len(pg)
            fetch = max([fetch] + [str(v[0]) for v in pg.values()])
        is_whole = use is whole and bool(whole)
        # total: 한 달을 통째로 조회한 예전 수집분에서만 '그 달의 서버 건수' 로 쓸 수 있다(구간으로 나눈 수집은 경계의 행을 두 번 세므로 months.json 을 본다)
        months[(op, ym)] = {"whole": is_whole, "windows": len(use), "pages": n_pages, "total": total if (is_whole and tot_known) else None,
                            "sum": s, "bad": bad, "short": short, "last_fetch": fetch[:10]}
    return months, bad_rows, n_rows


def _int(v) -> Optional[int]:
    return v if isinstance(v, int) and not isinstance(v, bool) else None


def analyse(data_dir: str, page_size: int, use_cache: bool, ops: dict, period=None, low_ratio: float = 0.5, today_ym: Optional[str] = None,
            raw_dir: Optional[str] = None, cache_db: Optional[str] = None, periods: Optional[dict] = None):
    """periods: {prefix: (YYYYMM, YYYYMM)} — 원문 종류마다 수집 기간이 다를 때(용역). 없으면 period 를 쓴다."""
    raw_dir = raw_dir or os.path.join(data_dir, "raw")
    cache_db = cache_db or os.path.join(data_dir, "cache.sqlite")
    today_ym = today_ym or date.today().strftime("%Y%m")
    cache_months, bad_rows, n_cache, cache_error = {}, [], 0, ""
    have_cache = use_cache and os.path.exists(cache_db)
    if have_cache:
        try:
            cache_months, bad_rows, n_cache = scan_cache(cache_db, set(ops.values()), page_size)
        except sqlite3.Error as e:                # 빈 파일·깨진 파일·다른 프로그램이 잠근 파일 → 캐시 없이 점검
            have_cache, cache_error = False, f"{type(e).__name__}: {e}"
    report = {}
    for prefix, op in ops.items():
        files, done, meta, junk = scan_files(raw_dir, prefix)
        if not files and not done:
            continue
        known = sorted(set(files) | done)
        lo, hi = known[0], known[-1]
        per = (periods or {}).get(prefix) or period
        if per and _YM.fullmatch(per[0]) and _YM.fullmatch(per[1]):
            lo, hi = min(lo, per[0]), max(hi, min(per[1], today_ym))
        filtered = prefix in FILTERED
        rows = []
        for ym in [m for m in month_seq(lo, hi) if m <= today_ym]:
            c = cache_months.get((op, ym))
            mj = meta.get(ym) if isinstance(meta.get(ym), dict) else None
            rec = mj if (mj and not mj.get("current")) else None               # 진행 중일 때의 기록은 완료 판정에 쓰지 않는다
            r = {"ym": ym, "lines": files.get(ym), "done": ym in done, "current": ym == today_ym, "flags": [], "confirmed": False,
                 "server": None, "cache_sum": c["sum"] if c else None, "cache_pages": c["pages"] if c else None, "cache_bad": c["bad"] if c else []}
            f = r["flags"]
            if r["lines"] is None:
                f.append("DONE_NO_FILE" if r["done"] else "NEVER_COLLECTED")
            elif not r["done"]:
                if not r["current"]:
                    f.append("NOT_DONE")
            else:                                                             # 완료 표시된 달만 서버 건수·캐시와 대조한다(받는 중인 달의 캐시는 미완성이 정상)
                confirmed = empty_ok = False
                tot, rcv, sav = (_int(rec.get("total")), _int(rec.get("received")), _int(rec.get("saved"))) if rec else (None, None, None)
                if rec and tot is not None and rcv is not None and sav is not None:
                    r["server"] = tot
                    if rcv < tot:
                        f.append("SHORT_VS_TOTAL")
                    if r["lines"] < sav:
                        f.append("SHORT_VS_SAVED")
                    confirmed = rcv >= tot and r["lines"] >= sav
                    empty_ok = sav == 0
                elif c and c["total"] is not None and not filtered:           # 한 달을 통째로 조회한 예전 수집분: 캐시의 totalCount 가 그 달의 서버 건수
                    r["server"] = c["total"]
                    if r["lines"] < c["total"]:
                        f.append("SHORT_VS_TOTAL")
                    else:
                        confirmed = True
                    empty_ok = c["total"] == 0 and not c["bad"]
                if c and not confirmed:
                    if c["bad"]:
                        f.append("BAD_CACHE_BODY")
                    if c["short"]:
                        f.append("CACHE_SUM_LT_TOTAL")
                if r["lines"] == 0 and not empty_ok and not confirmed:
                    f.append("EMPTY_DONE")
                if not filtered and page_size and r["lines"] > 0 and r["lines"] % page_size == 0 and not confirmed:
                    f.append("PAGE_MULTIPLE")
                r["confirmed"] = confirmed
            rows.append(r)
        fin = [r for r in rows if r["lines"] is not None and r["done"]]      # 앞뒤 6개월 중앙값 비교(완료된 달끼리). 서버 건수와 맞는 달은 적어도 정상
        for i, r in enumerate(fin):
            nb = [x["lines"] for x in fin[max(0, i - 6): i] + fin[i + 1: i + 7]]
            if len(nb) >= 4:
                med = statistics.median(nb)
                r["nb_median"] = med
                if med > 0 and r["lines"] < low_ratio * med and "EMPTY_DONE" not in r["flags"] and not r["confirmed"]:
                    r["flags"].append("LOW_VS_NEIGHBOURS")
        report[prefix] = {"op": op, "rows": rows, "done": len(done), "files": len(files), "junk": junk, "filtered": filtered}
    if "notices_" in report and "bsis_" in report:                           # 기초금액 ÷ 공고 비율(같은 달): 한쪽만 짧으면 비율이 튄다
        a = {r["ym"]: r for r in report["notices_"]["rows"]}
        b = {r["ym"]: r for r in report["bsis_"]["rows"]}
        ratios = {ym: b[ym]["lines"] / a[ym]["lines"] for ym in a if ym in b and a[ym]["lines"] and b[ym]["lines"] and a[ym]["done"] and b[ym]["done"]}
        if len(ratios) >= 6:
            med = statistics.median(ratios.values())
            for ym, q in ratios.items():
                if (q < 0.6 * med or q > 1.6 * med) and not (a[ym]["confirmed"] and b[ym]["confirmed"]):
                    (b[ym] if q < med else a[ym])["flags"].append("RATIO_OUTLIER")
            report["_ratio_median"] = round(med, 3)
    return report, bad_rows, {"have_cache": have_cache, "n_cache": n_cache, "cache_db": cache_db, "raw_dir": raw_dir, "cache_error": cache_error}


def hard_months(report) -> Dict[str, List[str]]:
    return {p: [r["ym"] for r in rep["rows"] if any(f in HARD for f in r["flags"])] for p, rep in report.items() if not p.startswith("_")}


def _next_cmds(prefixes) -> str:
    cmds = sorted({COLLECT_CMD.get(p, "collect") for p in prefixes})
    return " 그리고 ".join(f"python -m g2b_costdb.pipeline {c}" for c in cmds)


def print_report(report, bad_rows, meta, verbose: bool = False) -> int:
    def num(v) -> str:
        return "-" if v is None else format(v, ",")
    print(f"원문 폴더: {meta['raw_dir']}")
    if meta["have_cache"]:
        print(f"응답 캐시: {meta['cache_db']} (행 {meta['n_cache']:,}개 읽음)")
    elif meta.get("cache_error"):
        print(f"응답 캐시: {meta['cache_db']} 를 읽지 못했습니다({meta['cache_error']}) → 캐시 없이 점검합니다. 수집(collect)이 실행 중이면 끝난 뒤 다시 실행하세요.")
    else:
        print(f"응답 캐시: {meta['cache_db']} (없음 → 서버 건수 대조는 수집 기록(months.json)이 있는 달만, 나머지는 줄 수·쪽 배수·이웃 달 비교)")
    hard_total, n_done, n_confirmed, prefixes_hard = 0, 0, 0, []
    for prefix, rep in report.items():
        if prefix.startswith("_"):
            continue
        rows = rep["rows"]
        lines = [r["lines"] for r in rows if r["lines"] is not None]
        conf = sum(1 for r in rows if r["confirmed"])
        done_rows = sum(1 for r in rows if r["done"])
        n_done, n_confirmed = n_done + done_rows, n_confirmed + conf
        note = " (걸러서 저장하는 수집 — 저장 줄 수는 수집 때의 기록과 대조)" if rep["filtered"] else ""
        print(f"\n== {prefix}*.jsonl ({rep['op']}){note} — 파일 {rep['files']}개, 완료 표시 {rep['done']}개월, 합계 {sum(lines):,}행, "
              f"월 중앙값 {int(statistics.median(lines)) if lines else 0:,}행, 서버 건수와 맞는 것이 확인된 달 {conf}개 ==")
        if rep["junk"]:
            print(f"   [주의] {prefix}done.txt 에 달(YYYYMM)이 아닌 글자가 있어 무시했습니다: {', '.join(repr(x) for x in rep['junk'][:5])}")
        shown = [r for r in rows if verbose or r["flags"]]
        if shown:
            print(f"   {'월':6s} {'저장행':>9s} {'완료':>4s} {'서버건수':>9s} {'캐시합':>9s} {'쪽':>4s} {'이웃중앙':>8s}  표시")
        for r in shown:
            txt = ", ".join(f"{FLAG_KO.get(x, x)}[{x}]" for x in r["flags"]) + (f" (문제 쪽: {', '.join(r['cache_bad'][:6])})" if r["cache_bad"] and r["flags"] else "")
            print(f"   {r['ym']:6s} {num(r['lines']):>9s} {('Y' if r['done'] else ('·' if r['current'] else 'N')):>4s} {num(r['server']):>9s} "
                  f"{num(r['cache_sum']):>9s} {num(r['cache_pages']):>4s} {('-' if 'nb_median' not in r else format(int(r['nb_median']), ',')):>8s}  {txt or 'ok'}")
        hard = [r["ym"] for r in rows if any(f in HARD for f in r["flags"])]
        soft = [r["ym"] for r in rows if r["flags"] and r["ym"] not in hard]
        hard_total += len(hard)
        if hard:
            prefixes_hard.append(prefix)
        print(f"   → 다시 받아야 할 달 {len(hard)}개: {', '.join(hard) or '없음'}")
        print(f"   → 눈으로 확인할 달 {len(soft)}개: {', '.join(soft) or '없음'}")
    if "_ratio_median" in report:
        print(f"\n기초금액÷공고 비율 중앙값: {report['_ratio_median']}")
    if meta["have_cache"] and bad_rows:
        print(f"\n== 참고: 캐시에 남아 있는 '읽지 못한 응답' 행 {len(bad_rows)}개 ==")
        print("   예전 버전이 서버의 오류 응답을 '자료 없음'으로 저장해 둔 흔적입니다. 지금 버전은 이 행을 쓰지 않고 서버에 다시 묻습니다(위의 달 판정과는 별개).")
        for b in sorted(bad_rows, key=lambda x: (x["op"], x["bgn"], x["pageNo"]))[:60]:
            print(f"   {b['op']} 조회 시작일={b['bgn'] or '-'} {b['pageNo']}쪽 numOfRows={b['numOfRows']} 종류={b['kind']} 받은 날={str(b['fetched_at'])[:10]} "
                  f"기타 조건={json.dumps(b['params'], ensure_ascii=False)} 본문={b['body']!r}")
        if len(bad_rows) > 60:
            print(f"   … 외 {len(bad_rows) - 60}개")
    if hard_total:
        print("\n[판정] '다시 받아야 할 달'은 서버가 알려 준 건수보다 적게 저장됐거나, 읽지 못한 응답이 '자료 없음'으로 저장된 달입니다.\n"
              f"  python -m g2b_costdb.pipeline check-raw --reopen  으로 그 달들의 완료 표시를 지운 뒤(원문 파일은 그대로 둠) {_next_cmds(prefixes_hard)} 를 실행하세요.")
    elif not n_done and not [p for p in report if not p.startswith("_")]:
        print("\n[판정] 수집된 원문 파일이 없습니다(data/raw 가 비어 있음) — 점검할 것이 없습니다.")
    elif n_confirmed:
        print(f"\n[판정] 다시 받아야 할 달이 없습니다 — 완료 표시된 {n_done}개 달·종류 가운데 {n_confirmed}개는 저장된 줄 수가 서버가 알려 준 건수와 맞는 것을 확인했고"
              + ("," if n_confirmed < n_done else ".") + (f" 나머지 {n_done - n_confirmed}개는 대조할 기록이 없어 줄 수만 살폈습니다." if n_confirmed < n_done else ""))
    else:
        print("\n[판정] 다시 받아야 할 달이 없습니다 — 다만 서버 건수와 대조할 기록(응답 캐시·수집 기록)이 없어 줄 수·쪽 배수·이웃 달 비교만 했습니다.")
    return hard_total


def reopen(raw_dir: str, report) -> Dict[str, List[str]]:
    """다시 받아야 할 달(HARD)을 <prefix>done.txt 에서 뺀다. 원문 jsonl 은 건드리지 않는다(다시 받은 결과가 더 적으면 수집기가 .prev 로 남긴다)."""
    out = {}
    for prefix, months in hard_months(report).items():
        dp = os.path.join(raw_dir, f"{prefix}done.txt")
        if not months or not os.path.exists(dp):
            continue
        done, _ = read_done(dp)
        drop = sorted(done & set(months))
        if drop:
            with open(dp, "w", encoding="utf-8") as f:
                f.write("\n".join(sorted(done - set(drop))))
            out[prefix] = drop
    return out


def run(cfg: dict, do_reopen: bool = False, verbose: bool = False) -> int:
    """pipeline check-raw: 점검 결과를 출력하고 '다시 받아야 할 달' 수를 돌려준다."""
    p = cfg["paths"]
    ym = lambda v: str(v or "").replace("-", "")[:6]
    period = (ym(cfg["period"]["start"]), ym(cfg["period"]["end"]))
    sv = cfg.get("servc") or {}
    periods = {"servc_": (ym(sv.get("period_start")) or period[0], ym(sv.get("period_end")) or date.today().strftime("%Y%m"))}
    ops = dict(DEFAULT_OPS)
    a = cfg.get("api") or {}
    for prefix, key in (("notices_", "cnstwk_list"), ("bsis_", "cnstwk_bsis_amount"), ("servc_", "servc_list")):
        if (a.get("ops") or {}).get(key):
            ops[prefix] = a["ops"][key]
    if (a.get("award_ops") or {}).get("cnstwk_list"):
        ops["award_"] = a["award_ops"]["cnstwk_list"]
    if os.path.exists(p["cache_db"]):
        print(f"응답 캐시를 읽는 중입니다({os.path.getsize(p['cache_db']) / 1e6:,.0f} MB — 크면 1~2분 걸릴 수 있습니다. API 는 호출하지 않습니다) …", flush=True)
    report, bad_rows, meta = analyse(p["data_dir"], int(a.get("num_of_rows") or 999), True, ops, period, raw_dir=p["raw_dir"], cache_db=p["cache_db"],
                                     periods=periods)
    n = print_report(report, bad_rows, meta, verbose)
    if do_reopen:
        changed = reopen(p["raw_dir"], report)
        for prefix, months in changed.items():
            print(f"[완료 표시 삭제] {prefix}done.txt 에서 {len(months)}개월: {', '.join(months)}")
        if changed:
            print(f"다음: {_next_cmds(changed)}   (완료 표시가 없는 달 — 방금 지운 달과 아직 덜 받은 달 — 을 받습니다)")
        else:
            print("[완료 표시 삭제] 지울 달이 없습니다(다시 받아야 할 달이 없거나 이미 지웠습니다).")
    return n
