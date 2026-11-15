"""S4 중복·재발주 정리.

프로젝트 = (시설ID, 사업유형[, 세부사업]) — 같은 시설의 신축과 리모델링은 별도 프로젝트, 한 시설 후보에 여러 사업이 묶였으면 세부사업으로 나눈다
 0) 공종 공고의 사업유형이 '미분류'(예: 'OO미술관 전기공사')면 같은 시설에서 공고일이 가장 가까운 건축 공고의 사업유형을 상속
 1) 동일 (시설, 공고번호) 내 최대 차수(변경공고) 채택 — 같은 공고가 두 시설에 걸려도 서로 영향 없음
 2) 취소공고 제외 후, 시설ID×사업유형 안의 공고를 세부사업으로 나눈다(works.assign_works: 공고명의 현장·건물 이름, 식별 번호, 시기).
    세부사업이 둘 이상이면 프로젝트ID 에 '-01', '-02' … 를 붙인다(최초 공고일 순)
 3) 세부사업 안에서는 재공고·재발주로 대체된 공고만 빼고 모두 대표(합산 대상). 같은 공종의 공고가 여럿이면(기계소방/전기소방, 부대공사 등) 모두 합산
 4) 대체된 공고에는 '대체됨' + 대체공고번호 기록, 금액 급변(±30%) 시 경고
"""
from __future__ import annotations

import logging
import re
from typing import List, Optional, Tuple

import pandas as pd

from . import works
from .classify import WORK_TYPE_CODE, clean_notice_name

# 분할발주 토큰: N단계 / N공구 / N차분 / 골조 / 마감. 'N차 재공고·공고·입찰·변경'(공고 회차)과 '제N차'는 제외
_RENOTICE_ORD = re.compile(r"(제\s*)?\d+\s*차\s*(재공고|공고|입찰|재입찰|변경|정정|수정)")
_PHASE = re.compile(r"(\d+\s*단계|\d+\s*공구|\d+\s*차분|골조|마감)")
log = logging.getLogger(__name__)


def _phase_token(name: str) -> str:
    """괄호 안 표기('(1단계)', '[2공구]')도 살리기 위해 괄호를 지우기 전의 원문에서 추출한다."""
    raw = re.sub(_RENOTICE_ORD, " ", str(name or ""))
    return "".join(p.replace(" ", "") for p in _PHASE.findall(raw))


def _num(v) -> Optional[float]:
    try:
        if v is None:
            return None
        x = float(v)
        return None if x != x else x
    except (TypeError, ValueError):
        return None


def _mode(s: pd.Series, default: str = "") -> str:
    """가장 많이 나온 값. 같은 횟수면 가나다순 첫 값(입력 행 순서에 따라 결과가 달라지지 않게)."""
    vc = s[s.map(lambda v: v is not None and str(v).strip() not in ("", "nan", "None"))].astype(str).value_counts()
    return sorted(vc[vc == vc.max()].index)[0] if len(vc) else default


def inherit_work_type(d: pd.DataFrame, max_days: int = 730) -> pd.Series:
    """'미분류' 공고에 같은 시설의 가장 가까운(±max_days) 건축 공고 사업유형을 상속. 없으면 시설 내 최빈 유형, 그것도 없으면 '신축'."""
    out = d["사업유형"].copy()
    dt = pd.to_datetime(d["공고일시"].astype(str).str[:16], errors="coerce")
    for fid, g in d.groupby("시설ID"):
        typed = g[~g["사업유형"].isin(["미분류"])]
        anchors = typed[typed["공종"] == "건축"] if (typed["공종"] == "건축").any() else typed
        for i in g.index[g["사업유형"] == "미분류"]:
            if anchors.empty:
                out.at[i] = "신축"
                continue
            gap = (dt.loc[anchors.index] - dt.at[i]).abs().dt.days
            # 가장 가까운 건축 공고. 간격이 같으면 공고번호가 앞선 것(입력 행 순서와 무관하게 같은 결과)
            order = sorted(anchors.index, key=lambda k: (gap.at[k] if pd.notna(gap.at[k]) else float("inf"), str(anchors.at[k, "공고번호"]), str(anchors.at[k, "공고차수"])))
            j = order[0]
            out.at[i] = anchors.at[j, "사업유형"] if (pd.isna(gap.at[j]) or gap.at[j] <= max_days) \
                else _mode(typed["사업유형"], "신축")
    return out


def dedup_latest(df: pd.DataFrame, price_change_warn_ratio: float = 0.30,
                 minor_ratio: Optional[float] = None) -> Tuple[pd.DataFrame, pd.DataFrame, List[dict]]:
    """반환: (전체 이력 df[프로젝트ID·세부사업·최신여부·대체공고번호 부여], 대표(합산 대상) 공고 df, 검증로그 rows).
    minor_ratio 는 예전 '소액 공고는 대표 제외' 규칙의 인수로, 지금은 쓰지 않는다(이름이 다른 소액 공고는 같은 세부사업의 부대공사로 합산하거나 별도 세부사업이 된다)."""
    if df.empty:
        return df, df, []
    d = df.copy().reset_index(drop=True)              # 아래에서 행 라벨로 값을 쓰므로 라벨이 겹치지 않게
    d["공종"] = [t if (t is not None and str(t).strip() not in ("", "nan", "None")) else "기타" for t in d["공종"]]   # 공종이 빈 공고가 집계에서 빠지지 않게
    d["단계토큰"] = d["공고명"].map(_phase_token)
    d["사업유형_원분류"] = d["사업유형"]
    d["사업유형"] = inherit_work_type(d)
    d["_ord"] = pd.to_numeric(d["공고차수"], errors="coerce").fillna(0)
    d["_dt"] = pd.to_datetime(d["공고일시"].astype(str).str.strip().str[:19], errors="coerce")
    if d["_dt"].isna().any():
        n_bad = int(d["_dt"].isna().sum())
        log.warning("공고일시를 해석할 수 없는 공고 %d건 — 가장 오래된 것으로 취급", n_bad)
    d["_dt"] = d["_dt"].fillna(pd.Timestamp("1900-01-01"))
    d["_price"] = pd.to_numeric(d["추정가격"], errors="coerce")
    if "이전공고번호" not in d.columns:
        d["이전공고번호"] = ""
    # 이전공고번호는 차수마다 채워지지 않을 수 있다(변경 차수에서 빠짐) → 같은 (시설, 공고번호)의 어느 차수에든 있으면 모든 차수에 채운다
    pv = d["이전공고번호"].map(lambda v: "" if (v is None or str(v).strip() in ("", "nan", "None", "<NA>")) else str(v).strip())
    d["이전공고번호"] = pv.groupby([d["시설ID"].astype(str), d["공고번호"].astype(str)]).transform(lambda x: next((y for y in x if y), ""))

    # 1) (시설, 공고번호)별 최대 차수. 구차수의 사업유형은 최신 차수를 따른다(공고명 오타를 고친 변경공고 때문에 같은 공고가 두 프로젝트로 갈리지 않게)
    d = d.sort_values(["시설ID", "공고번호", "_ord", "_dt"])
    d["최신차수여부"] = ~d.duplicated(["시설ID", "공고번호"], keep="last")
    last_type = d[d["최신차수여부"]].set_index(["시설ID", "공고번호"])["사업유형"].to_dict()
    d["사업유형"] = [last_type.get((f, n), t) for f, n, t in zip(d["시설ID"], d["공고번호"], d["사업유형"])]
    d["_base"] = d["시설ID"].astype(str) + "-" + d["사업유형"].map(lambda t: WORK_TYPE_CODE.get(t, "U"))

    # 2) 세부사업 배정(유효 공고 = 최신 차수이고 취소가 아닌 공고)
    d["최신여부"] = False
    d["대체공고번호"] = ""
    d["대표선정사유"] = ""
    d["세부사업번호"] = 1
    d["세부사업명"] = ""
    logs: List[dict] = []
    valid = d["최신차수여부"] & (d["공고종류"] != "취소")
    w = works.assign_works(d[valid], "_base")
    d.loc[valid, "세부사업번호"] = w["세부사업번호"].astype(int)
    d.loc[valid, "세부사업명"] = w["세부사업명"]
    role = w["역할"].to_dict()
    n_works = d[valid].groupby("_base")["세부사업번호"].max().to_dict()
    # 유효하지 않은 행(구차수·취소공고)은 같은 공고번호의 유효 행을 따르고, 없으면 공고명이 가장 비슷한 유효 공고의 세부사업을 따른다
    by_notice = {(b, str(n)): k for b, n, k in zip(d.loc[valid, "_base"], d.loc[valid, "공고번호"], d.loc[valid, "세부사업번호"])}
    valid_by_base = {b: g for b, g in d[valid].groupby("_base")}
    for i in d.index[~valid]:
        b = d.at[i, "_base"]
        k = by_notice.get((b, str(d.at[i, "공고번호"])))
        if k is None and b in valid_by_base:
            g = valid_by_base[b]
            key = works.title_key(d.at[i, "공고명"])
            k = int(g.at[max(g.index, key=lambda j: works.similarity(key, works.title_key(g.at[j, "공고명"]))), "세부사업번호"])
        d.at[i, "세부사업번호"] = int(k or 1)
    name_of = {(b, k): n for b, k, n in zip(d.loc[valid, "_base"], d.loc[valid, "세부사업번호"], d.loc[valid, "세부사업명"])}
    d["세부사업수"] = d["_base"].map(lambda b: int(n_works.get(b, 1)))
    # 세부사업명은 시설 후보가 둘 이상으로 나뉜 경우에만 적는다(나뉘지 않았으면 시설명으로 충분)
    d["세부사업명"] = [name_of.get((b, k), "") if n_works.get(b, 1) > 1 else "" for b, k in zip(d["_base"], d["세부사업번호"])]
    d["프로젝트ID"] = [b if n_works.get(b, 1) <= 1 else f"{b}-{int(k):02d}" for b, k in zip(d["_base"], d["세부사업번호"])]
    d["프로젝트키"] = d["프로젝트ID"] + "|" + d["공종"].astype(str) + "|" + d["단계토큰"]
    for b, n in n_works.items():
        if n > 1:
            g = d[valid & (d["_base"] == b)]
            names = [f"{k:02d} {name_of.get((b, k), '')}" for k in sorted(set(g["세부사업번호"]))]
            logs.append({"공고번호": "", "항목": "세부사업 분리", "판정": "참고",
                         "비고": f"{b}({g['시설명'].iloc[0] if '시설명' in g.columns else ''}): 공고 {len(g)}건을 세부사업 {n}개로 나눔 — " + "; ".join(names)[:600]})

    # 3) 세부사업·공종별 대표(합산 대상): 재공고로 대체된 공고만 제외
    no_valid = set(d.loc[d["최신차수여부"], "_base"]) - set(d.loc[valid, "_base"])
    for b in sorted(no_valid):
        g = d[d["최신차수여부"] & (d["_base"] == b)]
        d.loc[g.index, "대표선정사유"] = "전부 취소공고"
        logs.append({"공고번호": ", ".join(g["공고번호"].astype(str)), "항목": "대표선정", "판정": "경고", "비고": f"{b}: 유효 공고 없음(모두 취소)"})
    for (pid, trade), g in d[d["최신차수여부"]].groupby(["프로젝트ID", "공종"], sort=False):
        if g["_base"].iloc[0] in no_valid:
            continue
        ok = g[g["공고종류"] != "취소"].sort_values("_dt")
        cnt = [i for i in ok.index if role.get(i) != "superseded"]
        sup = [i for i in ok.index if role.get(i) == "superseded"]
        n_cancel = int((g["공고종류"] == "취소").sum())
        for i in cnt:
            d.at[i, "최신여부"] = True
            n_rep = sum(1 for j in sup if str(w.at[j, "대체공고번호"]) == str(d.at[i, "공고번호"]))
            reason = "단독 공고" if len(cnt) == 1 else f"합산(같은 공종 {len(cnt)}건)"
            d.at[i, "대표선정사유"] = reason + (f", 재공고로 {n_rep}건 대체" if n_rep else "") + (f", 취소 {n_cancel}건 제외" if n_cancel else "")
        if len(cnt) > 1:
            parts = "; ".join(f"{d.at[i, '공고번호']} '{str(d.at[i, '공고명'])[:40]}' {(_num(d.at[i, '추정가격']) or 0) / 1e8:,.1f}억" for i in cnt)
            logs.append({"공고번호": str(d.at[cnt[-1], "공고번호"]), "항목": "같은 공종 합산", "판정": "참고",
                         "비고": f"{pid}|{trade}: 같은 세부사업·같은 공종의 공고 {len(cnt)}건을 모두 합산(분야별·차수별 발주, 부대공사로 봄) — {parts}"[:700]
                                 + " → 재공고가 섞여 이중계산이면 검수에서 확인"})
        for i in sup:
            rep_no = str(w.at[i, "대체공고번호"])
            d.at[i, "대체공고번호"] = rep_no
            d.at[i, "대표선정사유"] = "대체됨(이후 재공고/재발주)"
            rep = ok[ok["공고번호"].astype(str) == rep_no]
            p0, p1 = _num(d.at[i, "추정가격"]), (_num(rep["추정가격"].iloc[0]) if len(rep) else None)
            if p0 and p1 and p0 > 0:
                ratio = (p1 - p0) / p0
                if abs(ratio) > price_change_warn_ratio:
                    logs.append({"공고번호": rep_no, "항목": "재발주 금액변동", "API값": p0, "문서값": p1,
                                 "차이": p1 - p0, "차이율": round(ratio, 4), "판정": "경고",
                                 "비고": f"{pid}|{trade}: 선행 {d.at[i, '공고번호']} 대비 {ratio:+.1%} → 설계변경 의심"})
        for i in g.index[g["공고종류"] == "취소"]:
            d.at[i, "대표선정사유"] = "취소공고"
            if cnt:
                key = works.title_key(d.at[i, "공고명"])
                d.at[i, "대체공고번호"] = str(d.at[max(cnt, key=lambda j: works.similarity(key, works.title_key(d.at[j, "공고명"]))), "공고번호"])
    # 차수 갱신된 구공고 표시 (최신 차수가 취소공고면 '취소됨'으로 표시)
    old = ~d["최신차수여부"]
    last_kind = d[d["최신차수여부"]].set_index(["시설ID", "공고번호"])["공고종류"]
    for i in d.index[old]:
        k = (d.at[i, "시설ID"], d.at[i, "공고번호"])
        if last_kind.get(k) == "취소":
            d.at[i, "대표선정사유"] = "취소됨(이후 차수가 취소공고)"
        else:
            d.at[i, "대표선정사유"] = "구차수(이후 차수로 대체)"
            d.at[i, "대체공고번호"] = d.at[i, "공고번호"]

    # 5) API 의 이전공고번호(befBidBbancNo) 연결 가운데 세부사업 배정으로 처리되지 않은 것(사업유형이 달라 다른 시설×유형에 놓인 재공고):
    #    유효한(취소되지 않은 최신 차수) 공고 B 가 A 를 이전 공고로 가리키면 A 는 대표에서 뺀다. 취소된 재공고는 원 공고를 대체하지 못한다
    prev_rows = d[valid & d["이전공고번호"].map(bool)]
    for _, b in prev_rows.iterrows():
        targets = d.index[(d["시설ID"] == b["시설ID"]) & (d["공고번호"].astype(str) == str(b["이전공고번호"])) & d["최신여부"]]
        for i in targets:
            if i == b.name:
                continue
            d.at[i, "최신여부"] = False
            d.at[i, "대체공고번호"] = b["공고번호"]
            d.at[i, "대표선정사유"] = "대체됨(API 이전공고번호 연결)"
            logs.append({"공고번호": d.at[i, "공고번호"], "항목": "이전공고번호 연결", "판정": "참고",
                         "비고": f"{b['공고번호']} 의 이전 공고로 지정되어 대표에서 제외"})

    latest = d[d["최신여부"]].copy()
    d = d.drop(columns=["_ord", "_dt", "_price", "_base"])
    latest = latest.drop(columns=["_ord", "_dt", "_price", "_base"])
    return d, latest, logs


def facility_summary(latest: pd.DataFrame, all_hist: pd.DataFrame) -> pd.DataFrame:
    """프로젝트(시설×사업유형) 단위 요약: 공종 커버리지·공고건수·기간. 신축과 리모델링은 별도 행."""
    trades = ["건축", "전기", "정보통신", "소방", "조경", "기계설비", "토목", "기타"]
    rows = []
    if latest.empty:
        return pd.DataFrame(columns=["프로젝트ID", "시설ID", "시설명", "세부사업명", "세부사업수", "사업유형", "표2_대분류", "표2_중분류", "검색어", "수요기관",
                                     "공사현장지역", "공고건수_전체", "공고건수_최신", "최초공고일", "최종공고일"] + [f"공종_{t}" for t in trades])
    for pid, g in latest.groupby("프로젝트ID"):
        h = all_hist[all_hist["프로젝트ID"] == pid]
        row = {
            "프로젝트ID": pid, "시설ID": g["시설ID"].iloc[0], "시설명": g["시설명"].iloc[0],
            # 세부사업명: 시설 후보가 여러 사업으로 나뉜 경우 그 사업의 이름(대표 건축 공고명). 나뉘지 않았으면 빈칸
            "세부사업명": g["세부사업명"].iloc[0] if "세부사업명" in g.columns else "",
            "세부사업수": int(g["세부사업수"].iloc[0] or 1) if "세부사업수" in g.columns else 1,
            "사업유형": g["사업유형"].iloc[0],
            "표2_대분류": g["표2_대분류"].iloc[0], "표2_중분류": g["표2_중분류"].iloc[0], "검색어": g["검색어"].iloc[0],
            "수요기관": _mode(g["수요기관"]),
            "공사현장지역": _mode(g["공사현장지역"]) if "공사현장지역" in g.columns else "",
            "공고건수_전체": h["공고번호"].nunique(), "공고건수_최신": len(g),
            "최초공고일": h["공고일시"].min(), "최종공고일": h["공고일시"].max(),
        }
        for t in trades:
            row[f"공종_{t}"] = "O" if (g["공종"] == t).any() else "X"
        rows.append(row)
    return pd.DataFrame(rows)
