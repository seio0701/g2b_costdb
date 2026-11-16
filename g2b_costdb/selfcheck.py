"""이미 수집한 원문에 '비었거나 짧게 저장되고 완료 표시된 달'이 있는지 점검한다 (API 호출 없음, 인증키를 읽지 않음).

    python -m g2b_costdb.pipeline check-raw            # 점검만
    python -m g2b_costdb.pipeline check-raw --reopen   # 다시 받아야 할 달을 done.txt 에서 빼 둔다(원문 파일은 그대로) → 이어서 collect

배경: 2026-11-17 이전 버전의 수집기는 서버의 오류 응답(HTTP 200 에 'response' 가 없는 모양, 예: 조회 기간 초과 07)이나
잘린 쪽(totalCount>0 인데 항목 0건)을 '자료 없음'으로 읽어, 그 달을 0행 또는 쪽 경계까지만 저장하고 완료 표시했다.

보는 것
  1) data/raw/<prefix>YYYYMM.jsonl 의 줄 수(= 저장된 행 수)와 <prefix>done.txt, (새 수집기가 남기는) <prefix>months.json
  2) data/cache.sqlite 의 응답 캐시: 달·조회 구간·쪽별 totalCount / 항목 수 / 본문이 비어 있는 행({} 등)
표시(달마다)
  EMPTY_DONE         완료 표시된 달인데 0행
  SHORT_VS_TOTAL     저장 행 수 < 서버가 알려 준 건수(캐시의 totalCount 또는 months.json)
  BAD_CACHE_BODY     그 달의 캐시에 totalCount 도 항목도 없는 본문({} 등) 또는 totalCount>0 인데 항목 0건인 쪽이 있음
  CACHE_SUM_LT_TOTAL 캐시에 든 쪽들의 항목 합 < totalCount (쪽이 빠짐)
  PAGE_MULTIPLE      저장 행 수가 쪽 크기의 정확한 배수 — 쪽 경계에서 끊긴 흔적(캐시로 건수 일치가 확인되면 표시 안 함)
  DONE_NO_FILE       완료 표시인데 파일이 없음
  LOW_VS_NEIGHBOURS  앞뒤 6개월 중앙값의 절반 미만                  RATIO_OUTLIER  (기초금액÷공고) 비율이 전체 중앙값에서 크게 벗어남
  NOT_DONE           파일은 있는데 완료 표시가 없음(진행 중이던 달)    NEVER_COLLECTED 기간 안인데 파일도 완료 표시도 없음
앞의 여섯(HARD)은 다시 받아야 하는 달, 뒤의 넷(SOFT)은 눈으로 확인할 달.
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
HARD = ("EMPTY_DONE", "SHORT_VS_TOTAL", "BAD_CACHE_BODY", "CACHE_SUM_LT_TOTAL", "PAGE_MULTIPLE", "DONE_NO_FILE")
SOFT = ("LOW_VS_NEIGHBOURS", "RATIO_OUTLIER", "NOT_DONE", "NEVER_COLLECTED")
FLAG_KO = {"EMPTY_DONE": "빈 달인데 완료 표시", "SHORT_VS_TOTAL": "서버 건수보다 적게 저장", "BAD_CACHE_BODY": "읽지 못한 응답이 저장돼 있음",
           "CACHE_SUM_LT_TOTAL": "쪽이 빠짐", "PAGE_MULTIPLE": "쪽 경계에서 끊긴 흔적", "DONE_NO_FILE": "완료 표시인데 파일 없음",
           "LOW_VS_NEIGHBOURS": "앞뒤 달보다 크게 적음", "RATIO_OUTLIER": "기초금액÷공고 비율이 다른 달과 다름",
           "NOT_DONE": "완료 표시 없음(부분 수집)", "NEVER_COLLECTED": "받은 적 없음"}
_BAD_KINDS = ("NON_DICT", "EMPTY_NO_TOTAL", "TRUNCATED")


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
            if m:
                files[m.group(1)] = count_lines(os.path.join(raw_dir, fn))
    dp = os.path.join(raw_dir, f"{prefix}done.txt")
    done = set(open(dp, encoding="utf-8").read().split()) if os.path.exists(dp) else set()
    mp = os.path.join(raw_dir, f"{prefix}months.json")
    try:
        meta = json.load(open(mp, encoding="utf-8")) if os.path.exists(mp) else {}
    except ValueError:
        meta = {}
    return files, done, (meta if isinstance(meta, dict) else {})


def _kind(is_dict: bool, total, n: int, page_no: int, rows: int) -> Tuple[str, Optional[int], int]:
    """캐시 본문 요약 → (종류, totalCount 또는 None, 항목 수)."""
    if not is_dict:
        return "NON_DICT", None, 0
    try:
        total = int(float(total))
    except (TypeError, ValueError):
        total = None
    if n:
        return ("DATA" if total is not None else "DATA_NO_TOTAL"), total, n
    if total is None:
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
       substr(body, 1, 80)
FROM cache"""


def _cache_rows(con):
    """캐시의 모든 행 → (op, params, fetched_at, 본문이 객체인지, totalCount, 항목 수, 본문 앞부분). 본문은 SQLite 안에서 요약해 읽는다(쪽 하나가 수 MB)."""
    try:
        for op, params, fetched, typ, total, n, head in con.execute(_FAST_SQL):
            yield op, params, fetched, typ == "object", total, int(n or 0), head
        return
    except sqlite3.OperationalError:              # JSON 함수가 없는 SQLite → 파이썬으로 해석(느림)
        pass
    for op, params, fetched, body in con.execute("SELECT op, params, fetched_at, body FROM cache"):
        try:
            b = json.loads(body)
        except (TypeError, ValueError):
            b = None
        kind, total, n = body_kind(b, 1, 0)
        yield op, params, fetched, isinstance(b, dict), (b.get("totalCount") if isinstance(b, dict) else None), n, (body or "")[:80]


def _open_ro(cache_db: str):
    import pathlib
    try:
        return sqlite3.connect(pathlib.Path(cache_db).resolve().as_uri() + "?mode=ro", uri=True)     # 읽기 전용
    except sqlite3.Error:
        return sqlite3.connect(cache_db)


def scan_cache(cache_db: str, ops) -> Tuple[Dict, List[dict], int]:
    """→ (months[(op, ym)] = {...}, bad_rows, 읽은 행 수). 한 달이 여러 조회 구간(7일 단위)으로 나뉘어 있어도 달 단위로 모은다."""
    con = _open_ro(cache_db)
    bad_rows: List[dict] = []
    wins = defaultdict(lambda: defaultdict(dict))        # (op, ym, rows) -> {(bgn, end): {pageNo: (fetched, kind, total, n)}}
    n_rows = 0
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
        except ValueError:
            page_no, rows = 1, 0
        kind, total, n = _kind(is_dict, total, n, page_no, rows)
        if kind in _BAD_KINDS:
            bad_rows.append({"op": op, "ym": ym, "bgn": bgn[:8], "pageNo": page_no, "numOfRows": rows, "kind": kind, "fetched_at": fetched,
                             "params": {k: v for k, v in p.items() if k not in ("inqryBgnDt", "inqryEndDt", "pageNo", "numOfRows")},
                             "body": head})
        if op in ops and ym:
            cur = wins[(op, ym, rows)][(bgn, end)].get(page_no)
            if cur is None or str(fetched) > str(cur[0]):          # 같은 쪽을 여러 날 받았으면 가장 최근 것
                wins[(op, ym, rows)][(bgn, end)][page_no] = (fetched, kind, total, n)
    con.close()
    months = {}
    best = {}
    for (op, ym, rows), w in wins.items():                         # 달마다: 쪽 수가 가장 많은(=수집에 쓴) numOfRows 묶음
        n_pages = sum(len(pg) for pg in w.values())
        if (op, ym) not in best or (n_pages, rows) > best[(op, ym)][:2]:
            best[(op, ym)] = (n_pages, rows, w)
    for (op, ym), (_, rows, w) in best.items():
        last_day = calendar.monthrange(int(ym[:4]), int(ym[4:6]))[1]
        whole = {k: v for k, v in w.items() if k[0][6:8] == "01" and k[1][6:8] == f"{last_day:02d}"}
        part = {k: v for k, v in w.items() if k not in whole}
        if whole and part:                                         # 한 달 통째로 받은 것과 구간으로 나눠 받은 것이 함께 있으면 나중에 받은 쪽
            newest = lambda d: max(str(v[0]) for pg in d.values() for v in pg.values())
            use = part if newest(part) > newest(whole) else whole
        else:
            use = whole or part
        covered, day = bool(whole) and use is whole, 1
        if not covered:                                            # 구간들이 1일부터 말일까지 빈틈없이 이어지는지
            for (b, e) in sorted(use):
                if int(b[6:8]) != day:
                    break
                day = int(e[6:8]) + 1
            covered = day == last_day + 1
        total, tot_known, s, bad, short, n_pages, psize, fetch = 0, True, 0, [], [], 0, 0, ""
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
            psize = max([psize] + [v[3] for v in pg.values()])
            fetch = max([fetch] + [str(v[0]) for v in pg.values()])
        months[(op, ym)] = {"rows": rows, "windows": len(use), "pages": n_pages, "total": total if (covered and tot_known) else None,
                            "sum": s, "bad": bad, "short": short, "page_size": psize, "last_fetch": fetch[:10], "covered": covered}
    return months, bad_rows, n_rows


def analyse(data_dir: str, page_size: int, use_cache: bool, ops: dict, period=None, low_ratio: float = 0.5, today_ym: Optional[str] = None,
            raw_dir: Optional[str] = None, cache_db: Optional[str] = None):
    raw_dir = raw_dir or os.path.join(data_dir, "raw")
    cache_db = cache_db or os.path.join(data_dir, "cache.sqlite")
    today_ym = today_ym or date.today().strftime("%Y%m")
    cache_months, bad_rows, n_cache = ({}, [], 0)
    have_cache = use_cache and os.path.exists(cache_db)
    if have_cache:
        cache_months, bad_rows, n_cache = scan_cache(cache_db, set(ops.values()))
    report = {}
    for prefix, op in ops.items():
        files, done, meta = scan_files(raw_dir, prefix)
        if not files and not done:
            continue
        known = sorted(set(files) | done)
        lo, hi = known[0], known[-1]
        if period:
            lo, hi = min(lo, period[0]), max(hi, min(period[1], today_ym))
        rows = []
        for ym in [m for m in month_seq(lo, hi) if m <= today_ym]:
            c = cache_months.get((op, ym))
            mj = meta.get(ym) if isinstance(meta.get(ym), dict) else None
            r = {"ym": ym, "lines": files.get(ym), "done": ym in done, "current": ym == today_ym, "flags": [],
                 "cache_total": c["total"] if c else None, "cache_sum": c["sum"] if c else None,
                 "cache_pages": c["pages"] if c else None, "cache_bad": c["bad"] if c else [], "cache_fetch": c["last_fetch"] if c else "",
                 "rec_total": mj.get("total") if mj else None, "rec_received": mj.get("received") if mj else None}
            f = r["flags"]
            if r["lines"] is None:
                f.append("DONE_NO_FILE" if r["done"] else "NEVER_COLLECTED")
            else:
                if not r["done"] and not r["current"]:
                    f.append("NOT_DONE")
                if r["done"] and r["lines"] == 0 and not (c and c["total"] == 0 and not c["bad"]) and not (mj and mj.get("total") == 0):
                    f.append("EMPTY_DONE")                         # 서버가 0건이라고 답한 기록이 있으면(캐시 totalCount 0 / months.json) 빈 달이 맞다
                total_ok = False
                if prefix not in FILTERED and not r["current"]:
                    server = c["total"] if (c and c["total"] is not None) else None
                    if mj and isinstance(mj.get("total"), int) and isinstance(mj.get("received"), int) and not mj.get("current"):
                        if mj["received"] < mj["total"]:
                            f.append("SHORT_VS_TOTAL")
                        elif server is None and r["lines"] >= mj["total"]:
                            total_ok = True
                    if server is not None:
                        if r["lines"] < server and "SHORT_VS_TOTAL" not in f:
                            f.append("SHORT_VS_TOTAL")
                        elif r["lines"] >= server:
                            total_ok = True
                ps = page_size if prefix not in FILTERED else 0
                if ps and r["lines"] > 0 and r["lines"] % ps == 0 and not total_ok and r["done"]:
                    f.append("PAGE_MULTIPLE")
            if c and not r["current"]:
                if c["bad"]:
                    f.append("BAD_CACHE_BODY")
                if c["short"]:
                    f.append("CACHE_SUM_LT_TOTAL")
            rows.append(r)
        fin = [r for r in rows if r["lines"] is not None and r["done"]]      # 앞뒤 6개월 중앙값 비교(완료된 달끼리)
        for i, r in enumerate(fin):
            nb = [x["lines"] for x in fin[max(0, i - 6): i] + fin[i + 1: i + 7]]
            if len(nb) >= 4:
                med = statistics.median(nb)
                r["nb_median"] = med
                if med > 0 and r["lines"] < low_ratio * med and "EMPTY_DONE" not in r["flags"]:
                    r["flags"].append("LOW_VS_NEIGHBOURS")
        report[prefix] = {"op": op, "rows": rows, "done": len(done), "files": len(files)}
    if "notices_" in report and "bsis_" in report:                           # 기초금액 ÷ 공고 비율(같은 달): 한쪽만 짧으면 비율이 튄다
        a = {r["ym"]: r for r in report["notices_"]["rows"]}
        b = {r["ym"]: r for r in report["bsis_"]["rows"]}
        ratios = {ym: b[ym]["lines"] / a[ym]["lines"] for ym in a if ym in b and a[ym]["lines"] and b[ym]["lines"] and a[ym]["done"] and b[ym]["done"]}
        if len(ratios) >= 6:
            med = statistics.median(ratios.values())
            for ym, q in ratios.items():
                if q < 0.6 * med or q > 1.6 * med:
                    (b[ym] if q < med else a[ym])["flags"].append("RATIO_OUTLIER")
            report["_ratio_median"] = round(med, 3)
    return report, bad_rows, {"have_cache": have_cache, "n_cache": n_cache, "cache_db": cache_db, "raw_dir": raw_dir}


def hard_months(report) -> Dict[str, List[str]]:
    return {p: [r["ym"] for r in rep["rows"] if any(f in HARD for f in r["flags"])] for p, rep in report.items() if not p.startswith("_")}


def print_report(report, bad_rows, meta, verbose: bool = False) -> int:
    def num(v) -> str:
        return "-" if v is None else format(v, ",")
    print(f"원문 폴더: {meta['raw_dir']}")
    print(f"응답 캐시: {meta['cache_db']} " + (f"(행 {meta['n_cache']:,}개 읽음)" if meta["have_cache"]
                                              else "(없음 → 서버 건수 대조 생략, 줄 수·쪽 배수·이웃 달 비교만)"))
    hard_total = 0
    for prefix, rep in report.items():
        if prefix.startswith("_"):
            continue
        rows = rep["rows"]
        lines = [r["lines"] for r in rows if r["lines"] is not None]
        n_checked = sum(1 for r in rows if r["cache_total"] is not None or r["rec_total"] is not None)
        print(f"\n== {prefix}*.jsonl ({rep['op']}) — 파일 {rep['files']}개, 완료 표시 {rep['done']}개월, 합계 {sum(lines):,}행, "
              f"월 중앙값 {int(statistics.median(lines)) if lines else 0:,}행, 서버 건수와 대조한 달 {n_checked}개 ==")
        shown = [r for r in rows if verbose or r["flags"]]
        if shown:
            print(f"   {'월':6s} {'저장행':>9s} {'완료':>4s} {'서버건수':>9s} {'캐시합':>9s} {'쪽':>4s} {'이웃중앙':>8s}  표시")
        for r in shown:
            server = r["cache_total"] if r["cache_total"] is not None else r["rec_total"]
            txt = ", ".join(f"{FLAG_KO.get(x, x)}[{x}]" for x in r["flags"]) + (f" (문제 쪽: {', '.join(r['cache_bad'][:6])})" if r["cache_bad"] else "")
            print(f"   {r['ym']:6s} {num(r['lines']):>9s} {('Y' if r['done'] else ('·' if r['current'] else 'N')):>4s} {num(server):>9s} "
                  f"{num(r['cache_sum']):>9s} {num(r['cache_pages']):>4s} {('-' if 'nb_median' not in r else format(int(r['nb_median']), ',')):>8s}  {txt or 'ok'}")
        hard = [r["ym"] for r in rows if any(f in HARD for f in r["flags"])]
        soft = [r["ym"] for r in rows if r["flags"] and r["ym"] not in hard]
        hard_total += len(hard)
        print(f"   → 다시 받아야 할 달 {len(hard)}개: {', '.join(hard) or '없음'}")
        print(f"   → 눈으로 확인할 달 {len(soft)}개: {', '.join(soft) or '없음'}")
    if "_ratio_median" in report:
        print(f"\n기초금액÷공고 비율 중앙값: {report['_ratio_median']}")
    if meta["have_cache"]:
        print(f"\n== 캐시에 저장된 '읽지 못한 응답' 행 {len(bad_rows)}개 (모든 오퍼레이션, probe 호출 포함) ==")
        for b in sorted(bad_rows, key=lambda x: (x["op"], x["bgn"], x["pageNo"]))[:60]:
            print(f"   {b['op']} 조회 시작일={b['bgn'] or '-'} {b['pageNo']}쪽 numOfRows={b['numOfRows']} 종류={b['kind']} 받은 날={str(b['fetched_at'])[:10]} "
                  f"기타 조건={json.dumps(b['params'], ensure_ascii=False)} 본문={b['body']!r}")
        if len(bad_rows) > 60:
            print(f"   … 외 {len(bad_rows) - 60}개")
        if not bad_rows:
            print("   없음")
    if not [p for p in report if not p.startswith("_")]:
        print("\n수집된 원문 파일이 없습니다(data/raw 가 비어 있음).")
    print("\n[판정] " + ("'다시 받아야 할 달'은 서버가 알려 준 건수보다 적게 저장됐거나, 읽지 못한 응답이 '자료 없음'으로 저장된 달입니다.\n"
                        "  python -m g2b_costdb.pipeline check-raw --reopen  으로 그 달들의 완료 표시를 지운 뒤(원문 파일은 그대로 둠) collect 를 다시 실행하세요."
                        if hard_total else "다시 받아야 할 달이 없습니다 — 저장된 줄 수가 서버가 알려 준 건수와 맞고, 읽지 못한 응답도 저장돼 있지 않습니다(대조할 기록이 있는 달에 한함)."))
    return hard_total


def reopen(raw_dir: str, report) -> Dict[str, List[str]]:
    """다시 받아야 할 달(HARD)을 <prefix>done.txt 에서 뺀다. 원문 jsonl 은 건드리지 않는다(다시 받은 결과가 더 적으면 수집기가 .prev 로 남긴다)."""
    out = {}
    for prefix, months in hard_months(report).items():
        dp = os.path.join(raw_dir, f"{prefix}done.txt")
        if not months or not os.path.exists(dp):
            continue
        done = set(open(dp, encoding="utf-8").read().split())
        drop = sorted(done & set(months))
        if drop:
            with open(dp, "w", encoding="utf-8") as f:
                f.write("\n".join(sorted(done - set(drop))))
            out[prefix] = drop
    return out


def run(cfg: dict, do_reopen: bool = False, verbose: bool = False) -> int:
    """pipeline check-raw: 점검 결과를 출력하고 '다시 받아야 할 달' 수를 돌려준다."""
    p = cfg["paths"]
    period = (str(cfg["period"]["start"]).replace("-", ""), str(cfg["period"]["end"]).replace("-", ""))
    ops = dict(DEFAULT_OPS)
    a = cfg.get("api") or {}
    for prefix, key in (("notices_", "cnstwk_list"), ("bsis_", "cnstwk_bsis_amount"), ("servc_", "servc_list")):
        if (a.get("ops") or {}).get(key):
            ops[prefix] = a["ops"][key]
    if (a.get("award_ops") or {}).get("cnstwk_list"):
        ops["award_"] = a["award_ops"]["cnstwk_list"]
    if os.path.exists(p["cache_db"]):
        print(f"응답 캐시를 읽는 중입니다({os.path.getsize(p['cache_db']) / 1e6:,.0f} MB — 크면 1~2분 걸릴 수 있습니다. API 는 호출하지 않습니다) …", flush=True)
    report, bad_rows, meta = analyse(p["data_dir"], int(a.get("num_of_rows", 999)), True, ops, period, raw_dir=p["raw_dir"], cache_db=p["cache_db"])
    n = print_report(report, bad_rows, meta, verbose)
    if do_reopen:
        changed = reopen(p["raw_dir"], report)
        for prefix, months in changed.items():
            print(f"[완료 표시 삭제] {prefix}done.txt 에서 {len(months)}개월: {', '.join(months)}")
        print("다음: python -m g2b_costdb.pipeline collect   (지운 달만 다시 받습니다)" if changed else "[완료 표시 삭제] 지울 달이 없습니다.")
    return n
