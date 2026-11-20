"""S4 중복·재발주 정리.

프로젝트 = (시설ID, 사업유형[, 세부사업]) — 같은 시설의 신축과 리모델링은 별도 프로젝트, 한 시설 후보에 여러 사업이 묶였으면 세부사업으로 나눈다
 0) 공종 공고의 사업유형이 '미분류'(예: 'OO미술관 전기공사')면 같은 시설에서 공고일이 가장 가까운 건축 공고의 사업유형을 상속
 1) 동일 (시설, 공고번호) 내 최대 차수(변경공고) 채택 — 같은 공고가 두 시설에 걸려도 서로 영향 없음
 2) 취소공고 제외 후, 시설ID×사업유형 안의 공고를 세부사업으로 나눈다(works.assign_works: 공고명의 현장·건물 이름, 식별 번호, 시기).
    세부사업이 둘 이상이면 프로젝트ID 에 '-01', '-02' … 를 붙인다(최초 공고일 순)
 3) 세부사업 안에서는 재공고·재발주로 대체된 공고만 빼고 모두 대표(합산 대상). 같은 공종의 공고가 여럿이면(기계소방/전기소방, 부대공사 등) 모두 합산
 4) 대체된 공고에는 '대체됨(사유)' + 대체공고번호 기록, 금액 급변(±30%) 시 경고
 5) 공법·자재 선정용 제안서 제출 안내 공고(공사 발주 아님)는 대표에서 제외
 6) 같은 공고가 여러 시설 후보에 걸려 같은 사업이 두 프로젝트로 집계되면 cross_facility_duplicates 가 표시(05 시트 '시설간_중복'·'중복_제외권장')
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
PROPOSAL_REASON = "제안 안내 공고(공사 발주 아님)"


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


def _parse_dt(col: pd.Series) -> pd.Series:
    """공고일시 → Timestamp. 한 열에 형식이 섞여 있어도('2024-01-10 10:00:00', '2024-01-10T10:00', '202401101000') 값마다 다시 읽는다."""
    raw = col.map(lambda v: "" if (v is None or v is pd.NaT or (isinstance(v, float) and v != v)) else str(v).strip()).astype(object)
    out = pd.to_datetime(raw.map(lambda v: v[:19]), errors="coerce")
    bad = out.isna() & ~raw.isin(["", "nan", "None", "NaT", "<NA>"])
    for i in out.index[bad]:
        v = raw.at[i]
        t = pd.to_datetime(v[:19], errors="coerce")
        if pd.isna(t) and re.fullmatch(r"\d{8,14}", v):
            t = pd.to_datetime(v[:12].ljust(12, "0"), format="%Y%m%d%H%M", errors="coerce")
        out.at[i] = t
    return out


def _no(v) -> str:
    """공고번호·이전공고번호 비교용 문자열: 공백 제거, 빈 값은 '', 숫자로 읽혀 붙은 '.0' 제거."""
    if v is None:
        return ""
    t = str(v).strip()
    if t in ("", "nan", "None", "<NA>", "NaT"):
        return ""
    return t[:-2] if re.fullmatch(r"\d+\.0", t) else t


def inherit_work_type(d: pd.DataFrame, max_days: int = 730) -> pd.Series:
    """'미분류' 공고에 같은 시설의 가장 가까운(±max_days) 건축 공고 사업유형을 상속. 없으면 시설 내 최빈 유형, 그것도 없으면 '신축'."""
    out = d["사업유형"].copy()
    dt = _parse_dt(d["공고일시"])
    # research 가 공고번호로 보충한 차수(차수보충 = Y)는 기준에서 뺀다 — 차수가 늘었다고 다른 공고가 물려받는 유형(가까운 건축 공고·최빈 유형)이 달라지지 않게
    extra = (d["차수보충"].astype(str) == "Y") if "차수보충" in d.columns else pd.Series(False, index=d.index)
    for fid, g in d.groupby("시설ID"):
        typed = g[~g["사업유형"].isin(["미분류"]) & ~extra.loc[g.index]]
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
    d["_ord"] = pd.to_numeric(d["공고차수"], errors="coerce").fillna(0)
    d["_dt"] = _parse_dt(d["공고일시"])
    if d["_dt"].isna().any():
        n_bad = int(d["_dt"].isna().sum())
        log.warning("공고일시를 해석할 수 없는 공고 %d건 — 가장 오래된 것으로 취급", n_bad)
    d["_dt"] = d["_dt"].fillna(pd.Timestamp("1900-01-01"))
    # research 가 공고번호로 보충한 차수(차수보충 = Y: 시설 이름 검색에 잡히지 않은 차수 — 변경공고에서 공고명이 바뀐 경우)의 공고명은 시설 이름과 맞지 않는다.
    # 그 공고명으로 사업을 묶으면 같은 시설의 다른 공고와 갈리므로, 이 함수 안의 모든 판정(단계·세부사업·재공고·취소공고 배치)에는 같은 공고에서 검색으로 찾은
    # 가장 늦은 차수의 공고명을 쓴다 → 보충은 금액·일정·취소 여부만 바꾸고 사업 묶음은 보충 전과 같다. 표시 공고명은 끝에서 되돌린다.
    # (검색으로 찾은 행의 공고명은 건드리지 않는다 — 별칭으로 잡힌 새 이름의 공고는 그 이름으로 묶여야 한다)
    d["_shown"] = d["공고명"]
    sup = (d["차수보충"].astype(str) == "Y") if "차수보충" in d.columns else pd.Series(False, index=d.index)
    sup_base: dict = {}
    if sup.any():
        m = d[~sup].sort_values(["_ord", "_dt"], kind="stable").drop_duplicates(["시설ID", "공고번호"], keep="last")
        found = {(str(f), str(n)): j for f, n, j in zip(m["시설ID"], m["공고번호"], m.index)}
        sup_base = {i: found[(str(f), str(n))] for i, f, n in zip(d.index[sup], d.loc[sup, "시설ID"], d.loc[sup, "공고번호"]) if (str(f), str(n)) in found}
        d["공고명"] = [d.at[sup_base[i], "_shown"] if i in sup_base else t for i, t in zip(d.index, d["_shown"])]
        # 보충 행의 공종이 자기 공고명에서 나온 것이면(주공종명·면허제한 근거가 아님) 검색으로 찾은 차수의 공종을 따른다 —
        # 바뀐 공고명에서 '(전기)' 꼬리표가 빠졌다고 그 공고의 공종이 달라져 재공고 판정·같은 공종 합산이 달라지지 않게
        if "공종근거" in d.columns:
            for i, j in sup_base.items():
                if str(d.at[i, "공종근거"]) not in ("주공종명", "면허제한"):
                    d.at[i, "공종"], d.at[i, "공종근거"] = d.at[j, "공종"], d.at[j, "공종근거"]
    d["단계토큰"] = d["공고명"].map(_phase_token)
    d["사업유형_원분류"] = d["사업유형"]
    d["사업유형"] = inherit_work_type(d)
    for i, j in sup_base.items():                     # 보충 행의 사업유형은 검색으로 찾은 차수가 물려받은 것 그대로(보충 차수의 공고일시로 다시 고르지 않는다)
        if d.at[i, "사업유형_원분류"] == "미분류" and d.at[j, "사업유형_원분류"] == "미분류":
            d.at[i, "사업유형"] = d.at[j, "사업유형"]
    d["_price"] = pd.to_numeric(d["추정가격"], errors="coerce")
    if "이전공고번호" not in d.columns:
        d["이전공고번호"] = ""
    # 이전공고번호는 차수마다 채워지지 않을 수 있다(변경 차수에서 빠짐) → 같은 (시설, 공고번호)의 어느 차수에든 있으면 모든 차수에 채운다
    pv = d["이전공고번호"].map(_no)
    # 차수마다 값이 다르면 판정에는 가장 늦은 차수의 값을 쓴다(입력 행 순서와 무관). 보이는 열은 빈 차수만 채우고 API 가 준 값은 그대로 둔다
    _o = d.assign(_pv=pv)[(pv != "") & (pv != d["공고번호"].map(_no))].sort_values(["_ord", "_dt", "_pv"])
    _last = _o.groupby([_o["시설ID"].astype(str), _o["공고번호"].astype(str)])["_pv"].last().to_dict()
    d["_pv_eff"] = [_last.get((str(f), str(n)), "") for f, n in zip(d["시설ID"], d["공고번호"])]
    d["이전공고번호"] = pv.where(pv != "", d["_pv_eff"])

    # 1) (시설, 공고번호)별 최대 차수. 구차수의 사업유형은 최신 차수를 따른다(공고명 오타를 고친 변경공고 때문에 같은 공고가 두 프로젝트로 갈리지 않게)
    d = d.sort_values(["시설ID", "공고번호", "_ord", "_dt"])
    d["최신차수여부"] = ~d.duplicated(["시설ID", "공고번호"], keep="last")
    last_type = d[d["최신차수여부"]].set_index(["시설ID", "공고번호"])["사업유형"].to_dict()
    d["사업유형"] = [last_type.get((f, n), t) for f, n, t in zip(d["시설ID"], d["공고번호"], d["사업유형"])]
    d["_base"] = d["시설ID"].astype(str) + "-" + d["사업유형"].map(lambda t: WORK_TYPE_CODE.get(t, "U"))

    # 2) 세부사업 배정(유효 공고 = 최신 차수이고 취소·제안 안내 공고가 아닌 공고)
    d["최신여부"] = False
    d["대체공고번호"] = ""
    d["대표선정사유"] = ""
    d["세부사업번호"] = 1
    d["세부사업명"] = ""
    logs: List[dict] = []
    d["_prop"] = d["공고명"].map(works.is_proposal)     # 공법·자재 선정용 제안서 제출 안내 공고: 공사 발주가 아니므로 합산하지 않는다
    valid = d["최신차수여부"] & (d["공고종류"] != "취소") & ~d["_prop"]
    for i in d.index[d["최신차수여부"] & sup.reindex(d.index)]:      # 어느 공고의 최신 차수가 보충으로 바뀌었는지 06 에 남긴다(금액·일정이 바뀌었는지 볼 단서)
        j = sup_base.get(i)
        base_txt = f"검색으로 찾은 가장 늦은 차수는 {d.at[j, '공고차수']} '{str(d.at[j, '_shown'])[:60]}'" if j is not None else "검색으로 찾은 차수 없음"
        tail = " 이 차수가 취소공고라 이 공고는 합산하지 않음" if d.at[i, "공고종류"] == "취소" else \
            (" 검색으로 찾은 차수의 공고명이 제안 안내 공고라 합산하지 않음" if d.at[i, "_prop"] else " 변경공고에서 금액이 바뀌었는지 확인")
        logs.append({"공고번호": str(d.at[i, "공고번호"]), "항목": "최신 차수 보충", "판정": "참고",
                     "비고": f"{d.at[i, '시설ID']}: 최신 차수 {d.at[i, '공고차수']} '{str(d.at[i, '_shown'])[:60]}' 은 시설 이름 검색에 잡히지 않아 공고번호로 보충한 차수({base_txt}). "
                             f"사업 묶음은 보충 전과 같게 두고 금액·일정·표시 공고명만 최신 차수 것을 씀 —{tail}"})
    # 이전공고번호가 유효하지 않은 공고(취소된 재공고 등)를 가리키면 그 공고의 이전공고번호를 따라 올라간다: A ← B(취소) ← C 이면 C 가 대체하는 공고는 A
    d["_no"] = d["공고번호"].map(_no)
    valid_nos = set(zip(d.loc[valid, "시설ID"].astype(str), d.loc[valid, "_no"]))
    prev_of = {(str(f), n): p for f, n, p in zip(d["시설ID"], d["_no"], d["_pv_eff"]) if p}

    def _valid_prev(f, p):
        seen = set()
        while p and (f, p) not in valid_nos and (f, p) in prev_of and p not in seen:
            seen.add(p)
            p = prev_of[(f, p)]
        return p
    d["_prev"] = [_valid_prev(str(f), p) if p else "" for f, p in zip(d["시설ID"], d["_pv_eff"])]
    # works 에도 정규화한 공고번호를 넘긴다(이전공고번호와 같은 꼴이어야 연결을 찾는다 — 숫자로 읽혀 '.0' 이 붙은 공고번호)
    w = works.assign_works(d[valid].assign(이전공고번호=d.loc[valid, "_prev"], 공고번호=d.loc[valid, "_no"]), "_base")
    d.loc[valid, "세부사업번호"] = w["세부사업번호"].astype(int)
    d.loc[valid, "세부사업명"] = w["세부사업명"]
    role = w["역할"].to_dict()
    n_works = d[valid].groupby("_base")["세부사업번호"].max().to_dict()
    # 유효하지 않은 행(구차수·취소공고·제안 안내 공고)은 같은 공고번호의 유효 행을 따른다. 유효 행이 없는 공고는 그 공고의 최신 차수 공고명으로
    # 세부사업을 고른다(현장 이름[구분어]이 같은 유효 공고 가운데 공고명이 가장 비슷한 것 → 없으면 전체에서 가장 비슷한 것). 한 공고의 모든 차수는 같은 세부사업
    by_notice = {(b, str(n)): k for b, n, k in zip(d.loc[valid, "_base"], d.loc[valid, "공고번호"], d.loc[valid, "세부사업번호"])}
    valid_by_base = {b: g for b, g in d[valid].groupby("_base")}
    orphan: dict = {}
    for i in d.index[~valid]:
        b, no = d.at[i, "_base"], str(d.at[i, "공고번호"])
        k = by_notice.get((b, no))
        if k is None:
            orphan.setdefault((b, no), []).append(i)
        else:
            d.at[i, "세부사업번호"] = int(k)
    disc_cache: dict = {}
    for (b, no), rows in orphan.items():
        k = 1
        if b in valid_by_base:
            g = valid_by_base[b]
            ref = max(rows, key=lambda i: (d.at[i, "_ord"], d.at[i, "_dt"]))
            fac_name = d.at[ref, "시설명"] if "시설명" in d.columns else ""
            if b not in disc_cache:
                disc_cache[b] = {j: works.features(g.at[j, "공고명"], fac_name)["disc"] for j in g.index}
            disc, key, trade = works.features(d.at[ref, "공고명"], fac_name)["disc"], works.title_key(d.at[ref, "공고명"]), d.at[ref, "공종"]
            gap = {j: abs((g.at[j, "_dt"] - d.at[ref, "_dt"]).days) for j in g.index}
            near = works.PARAMS["rebid_max_days"]
            same = [j for j in g.index if g.at[j, "공종"] == trade and gap[j] <= near and works.title_key(g.at[j, "공고명"]) == key]
            if same:                                   # 같은 공고명·같은 공종으로 가까운 때 다시 난 공고(취소 뒤 재공고)가 있으면 그 사업
                best = min(same, key=lambda j: (gap[j], str(g.at[j, "공고번호"])))
            else:                                      # 없으면 현장 이름이 같은 공고 가운데 가까운 때의 것, 이름이 비슷한 것, 같은 공종 순
                cand = [j for j in g.index if disc_cache[b][j] == disc] or list(g.index)
                best = max(cand, key=lambda j: (gap[j] <= near, works.similarity(key, works.title_key(g.at[j, "공고명"])), g.at[j, "공종"] == trade, -gap[j]))
            k = int(g.at[best, "세부사업번호"])
        for i in rows:
            d.at[i, "세부사업번호"] = k
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
    first_dt = d.groupby([d["시설ID"].astype(str), d["_no"]])["_dt"].min().to_dict()
    no_valid = set(d.loc[d["최신차수여부"], "_base"]) - set(d.loc[valid, "_base"])
    for b in sorted(no_valid):
        g = d[d["최신차수여부"] & (d["_base"] == b)]
        only_cancel = bool((g["공고종류"] == "취소").all())
        d.loc[g.index, "대표선정사유"] = ["전부 취소공고" if c else PROPOSAL_REASON for c in (g["공고종류"] == "취소")]
        logs.append({"공고번호": ", ".join(g["공고번호"].astype(str)), "항목": "대표선정", "판정": "경고",
                     "비고": f"{b}: 유효 공고 없음({'모두 취소' if only_cancel else '취소 또는 제안 안내 공고뿐'})"})
    for (pid, trade), g in d[d["최신차수여부"]].groupby(["프로젝트ID", "공종"], sort=False):
        if g["_base"].iloc[0] in no_valid:
            continue
        ok = g[(g["공고종류"] != "취소") & ~g["_prop"]].sort_values("_dt")
        for i in g.index[g["_prop"] & (g["공고종류"] != "취소")]:
            d.at[i, "대표선정사유"] = PROPOSAL_REASON
            amt = _num(d.at[i, "추정가격"]) or 0
            logs.append({"공고번호": str(d.at[i, "공고번호"]), "항목": "제안 안내 공고 제외", "판정": "참고",
                         "비고": f"{pid}|{trade}: '{str(d.at[i, '공고명'])[:60]}' 추정가격 {amt / 1e8:,.2f}억 — 공법·자재 선정용 제안서 제출 안내(공사 발주 아님)로 보아 합산하지 않음. "
                                 "이 금액이 본 공사 공고에 들어 있지 않은 별도 계약이면 검수에서 확인"})
        cnt = [i for i in ok.index if role.get(i) != "superseded"]
        sup = [i for i in ok.index if role.get(i) == "superseded"]
        n_cancel = int((g["공고종류"] == "취소").sum())
        for i in cnt:
            d.at[i, "최신여부"] = True
            n_rep = sum(1 for j in sup if str(w.at[j, "대체공고번호"]) == d.at[i, "_no"])
            reason = "단독 공고" if len(cnt) == 1 else f"합산(같은 공종 {len(cnt)}건)"
            d.at[i, "대표선정사유"] = reason + (f", 재공고로 {n_rep}건 대체" if n_rep else "") + (f", 취소 {n_cancel}건 제외" if n_cancel else "")
        if len(cnt) > 1:
            parts = "; ".join(f"{d.at[i, '공고번호']} '{str(d.at[i, '_shown'])[:40]}' {(_num(d.at[i, '추정가격']) or 0) / 1e8:,.1f}억" for i in cnt)
            logs.append({"공고번호": str(d.at[cnt[-1], "공고번호"]), "항목": "같은 공종 합산", "판정": "참고",
                         "비고": f"{pid}|{trade}: 같은 세부사업·같은 공종의 공고 {len(cnt)}건을 모두 합산(분야별·차수별 발주, 부대공사로 봄) — {parts}"[:700]
                                 + " → 재공고가 섞여 이중계산이면 검수에서 확인"})
        for i in sup:
            rep_no = str(w.at[i, "대체공고번호"])
            d.at[i, "대체공고번호"] = rep_no
            d.at[i, "대표선정사유"] = f"대체됨({str(w.at[i, '사유']) or '이후 재공고/재발주'})"
            rep = ok[ok["_no"] == rep_no]
            p0, p1 = _num(d.at[i, "추정가격"]), (_num(rep["추정가격"].iloc[0]) if len(rep) else None)
            if p0 and p1 and p0 > 0:
                ratio = (p1 - p0) / p0
                if abs(ratio) > price_change_warn_ratio:
                    logs.append({"공고번호": rep_no, "항목": "재발주 금액변동", "API값": p0, "문서값": p1,
                                 "차이": p1 - p0, "차이율": round(ratio, 4), "판정": "경고",
                                 "비고": f"{pid}|{trade}: 선행 {d.at[i, '공고번호']} 대비 {ratio:+.1%} → 설계변경 의심"})
        for i in g.index[g["공고종류"] == "취소"]:
            d.at[i, "대표선정사유"] = "취소공고"
            # 취소공고가 처음 난 날보다 먼저 난 공고는 그 공고를 '대체한 공고'가 될 수 없다(이전공고번호로 이어진 원 공고는 예외) → 없으면 비워 둔다
            first = first_dt.get((str(d.at[i, "시설ID"]), d.at[i, "_no"]), d.at[i, "_dt"])
            later = [j for j in cnt if first_dt.get((str(d.at[j, "시설ID"]), d.at[j, "_no"]), d.at[j, "_dt"]) >= first or d.at[j, "_no"] == d.at[i, "_prev"]]
            if later:
                key = works.title_key(d.at[i, "공고명"])
                d.at[i, "대체공고번호"] = str(d.at[max(later, key=lambda j: works.similarity(key, works.title_key(d.at[j, "공고명"]))), "공고번호"])
    # 차수 갱신된 구공고 표시 (최신 차수가 취소공고면 '취소됨'으로 표시)
    old = ~d["최신차수여부"]
    last_kind = d[d["최신차수여부"]].set_index(["시설ID", "공고번호"])["공고종류"]
    for i in d.index[old]:
        k = (d.at[i, "시설ID"], d.at[i, "공고번호"])
        if last_kind.get(k) == "취소":
            d.at[i, "대표선정사유"] = "취소됨(이후 차수가 취소공고)"
        else:
            d.at[i, "대표선정사유"] = "구차수(이후 차수로 대체)"
            d.at[i, "대체공고번호"] = str(d.at[i, "공고번호"])

    # 5) API 의 이전공고번호(befBidBbancNo) 연결 가운데 세부사업 배정으로 처리되지 않은 것(사업유형이 달라 다른 시설×유형에 놓인 재공고):
    #    유효한(취소되지 않은 최신 차수) 공고 B 가 A 를 이전 공고로 가리키면 A 는 대표에서 뺀다. 취소된 재공고는 원 공고를 대체하지 못한다
    prev_rows = d[valid & d["_prev"].map(bool)].sort_values(["_dt", "_no"], ascending=False)      # 늦게 난 공고부터(잘못된 순환 연결이면 늦은 공고가 남는다)
    row_of = {(f, n): i for i, f, n in zip(d.index[d["최신차수여부"]], d.loc[d["최신차수여부"], "시설ID"].astype(str), d.loc[d["최신차수여부"], "_no"])}
    demoted = set()                                    # 여기서 대표가 빠진 (프로젝트ID, 공종)
    for _, b in prev_rows.iterrows():
        f = str(b["시설ID"])
        i = row_of.get((f, b["_prev"]))
        if i is None or i == b.name or not d.at[i, "최신여부"]:
            continue
        t, seen = b.name, set()                       # b 가 (여러 단계를 거쳐) 이미 i 로 대체되어 있으면 서로를 가리키는 연결 → 건너뛴다
        while t is not None and t != i and t not in seen and not d.at[t, "최신여부"]:
            seen.add(t)
            t = row_of.get((f, _no(d.at[t, "대체공고번호"])))
        if t == i:
            continue
        d.at[i, "최신여부"] = False
        d.at[i, "대체공고번호"] = str(b["공고번호"])
        d.at[i, "대표선정사유"] = "대체됨(API 이전공고번호 연결)"
        demoted.add((d.at[i, "프로젝트ID"], d.at[i, "공종"]))
        logs.append({"공고번호": d.at[i, "공고번호"], "항목": "이전공고번호 연결", "판정": "참고",
                     "비고": f"{b['공고번호']} 의 이전 공고로 지정되어 대표에서 제외"})

    # 위 5)에서 대표가 바뀌면 대체공고번호가 대표에서 빠진 공고를 가리킬 수 있다 → 끝까지 남은 공고를 가리키게 한다(대체됨·취소공고 행)
    for i in d.index[d["최신차수여부"] & d["대체공고번호"].map(bool) & ~d["최신여부"]]:
        f = str(d.at[i, "시설ID"])
        t, seen = row_of.get((f, _no(d.at[i, "대체공고번호"]))), {i}
        while t is not None and not d.at[t, "최신여부"] and t not in seen:
            seen.add(t)
            t = row_of.get((f, _no(d.at[t, "대체공고번호"])))
        if t is not None and d.at[t, "최신여부"]:
            d.at[i, "대체공고번호"] = str(d.at[t, "공고번호"]).strip()
    # 대표가 빠진 프로젝트·공종에 남은 대표 공고의 사유('합산(같은 공종 N건)'/'단독 공고')를 남은 건수에 맞춘다
    for pid, trade in sorted(demoted):
        left = d.index[d["최신여부"] & (d["프로젝트ID"] == pid) & (d["공종"] == trade)]
        head = "단독 공고" if len(left) == 1 else f"합산(같은 공종 {len(left)}건)"
        for j in left:
            d.at[j, "대표선정사유"] = re.sub(r"^(?:단독 공고|합산\(같은 공종 \d+건\))", head, str(d.at[j, "대표선정사유"]))
        k = next((n for n, l in enumerate(logs) if l.get("항목") == "같은 공종 합산" and str(l.get("비고")).startswith(f"{pid}|{trade}: ")), None)
        if k is not None:
            if len(left) > 1:
                keep = sorted(left, key=lambda j: (d.at[j, "_dt"], d.at[j, "_no"]))
                parts = "; ".join(f"{d.at[j, '공고번호']} '{str(d.at[j, '_shown'])[:40]}' {(_num(d.at[j, '추정가격']) or 0) / 1e8:,.1f}억" for j in keep)
                logs[k] = {"공고번호": str(d.at[keep[-1], "공고번호"]), "항목": "같은 공종 합산", "판정": "참고",
                           "비고": f"{pid}|{trade}: 같은 세부사업·같은 공종의 공고 {len(keep)}건을 모두 합산(분야별·차수별 발주, 부대공사로 봄) — {parts}"[:700]
                                   + " → 재공고가 섞여 이중계산이면 검수에서 확인"}
            else:
                del logs[k]
    d["공고명"] = d["_shown"]                          # 묶음에 쓴 공고명 → 그 차수의 공고명으로 되돌린다
    latest = d[d["최신여부"]].copy()
    tmp_cols = ["_ord", "_dt", "_price", "_base", "_prop", "_no", "_prev", "_pv_eff", "_shown"]
    d = d.drop(columns=tmp_cols)
    latest = latest.drop(columns=tmp_cols)
    return d, latest, logs


def facility_summary(latest: pd.DataFrame, all_hist: pd.DataFrame) -> pd.DataFrame:
    """프로젝트(시설×사업유형) 단위 요약: 공종 커버리지·공고건수·기간. 신축과 리모델링은 별도 행."""
    trades = ["건축", "전기", "정보통신", "소방", "조경", "기계설비", "토목", "기타"]
    rows = []
    if latest.empty:
        return pd.DataFrame(columns=["프로젝트ID", "시설ID", "시설명", "세부사업명", "세부사업수", "사업유형", "표2_대분류", "표2_중분류", "검색어", "수요기관",
                                     "공사현장지역", "공고건수_전체", "공고건수_최신", "최초공고일", "최종공고일"] + [f"공종_{t}" for t in trades]
                                    + ["시설간_중복", "중복_제외권장"])
    dup = cross_facility_duplicates(latest)
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
            # 최종공고일(05 '공고연도'의 근거)은 합산 대상 공고 기준 — 유효 차수가 없는 취소공고가 다른 해에 붙어 있어도 연도가 밀리지 않게
            # 빈 값(None)은 빼고 견준다 — 문자열과 결측이 섞인 열의 min/max 는 판다스 2.x 에서 TypeError 가 난다
            "최초공고일": h["공고일시"].dropna().min(), "최종공고일": g["공고일시"].dropna().max() if g["공고일시"].notna().any() else h["공고일시"].dropna().max(),
        }
        for t in trades:
            row[f"공종_{t}"] = "O" if (g["공종"] == t).any() else "X"
        row["시설간_중복"] = dup.at[pid, "시설간_중복"] if pid in dup.index else ""
        row["중복_제외권장"] = dup.at[pid, "중복_제외권장"] if pid in dup.index else ""
        rows.append(row)
    return pd.DataFrame(rows)


def cross_facility_duplicates(latest: pd.DataFrame) -> pd.DataFrame:
    """같은 공고가 여러 시설 후보의 대표로 잡혀 같은 사업이 05 시트에 두 번 이상 집계되는 경우를 찾는다
    (예: '○○공사 공공주택' 이라는 묶음 시설의 세부사업과, 그 단지 이름으로 따로 잡힌 시설의 프로젝트가 같은 공고들로 이루어진 경우).
    반환: index = 프로젝트ID(다른 시설의 프로젝트와 공고가 겹치는 것만), 열
      시설간_중복   — 겹치는 프로젝트와 관계: '같은 사업: X'(대표 공고가 모두 같음) / 'X 의 일부'(이 프로젝트의 공고가 모두 X 에 있음) /
                      'X 를 포함(n/m건)' / 'X 와 일부 겹침(n/m건)'  (m = 이 프로젝트의 대표 공고 수, n = 그중 X 에도 있는 수) /
                      'X + Y 에 나뉘어 있음'(이 프로젝트의 공고가 다른 시설 후보 하나의 세부사업 X, Y 에 나뉘어 모두 들어 있음)
      중복_제외권장 — 'Y' = 이 행을 빼도 그 공고가 남는 프로젝트가 있다(다른 프로젝트의 일부이거나, 같은 사업 가운데 남길 한 행이 아닌 경우). Y 행을 모두 빼면
                      같은 공고가 두 번 합산되지 않는다 — 단, 서로 일부만 겹치는 프로젝트('와 일부 겹침')는 그것만으로는 Y 가 붙지 않으므로,
                      다른 관계로 한쪽이 Y 가 되지 않으면 그 공고는 두 번 남는다(pipeline.stage_dedup 이 Y 행을 뺀 뒤에 남는 건수를 세어 적는다).
                      같은 사업끼리는 세부사업으로 나뉘지 않은(더 구체적인) 시설의 프로젝트를 남기고, 같으면 프로젝트ID 가 앞선 것을 남긴다.
                      일부만 겹치는 경우는 어느 쪽도 권장하지 않는다(검수에서 판단)."""
    cols = ["시설간_중복", "중복_제외권장"]
    if latest is None or latest.empty or "프로젝트ID" not in latest.columns:
        return pd.DataFrame(columns=cols)
    sets = {str(p): frozenset(g["공고번호"].map(_no)) for p, g in latest.groupby("프로젝트ID")}
    fac = {str(p): str(f) for p, f in zip(latest["프로젝트ID"], latest["시설ID"])}
    n_works = {str(p): (int(k) if (k is not None and k == k) else 1) for p, k in zip(latest["프로젝트ID"], latest["세부사업수"])} if "세부사업수" in latest.columns else {}
    by_no: dict = {}
    for p, nos in sets.items():
        for n in nos:
            by_no.setdefault(n, set()).add(p)
    share: dict = {}
    for ps in by_no.values():
        if len(ps) < 2:
            continue
        for a in ps:
            for b in ps:
                if a != b and fac[a] != fac[b]:
                    share.setdefault(a, {})
                    share[a][b] = share[a].get(b, 0) + 1
    keep_key = lambda p: (n_works.get(p, 1), p)
    # 한 프로젝트의 공고가 다른 시설 후보 하나의 여러 세부사업에 나뉘어 모두 들어 있고 그쪽에 공고가 더 있으면(그쪽이 더 완전한 묶음) 이 프로젝트를 뺀다.
    # 이때 그 세부사업들은 '이 프로젝트의 일부'라는 이유로 빼지 않는다(둘 다 빼면 공고가 사라진다)
    spread: dict = {}
    for a in sorted(share, key=keep_key, reverse=True):
        by_f: dict = {}
        for b in share[a]:
            by_f.setdefault(fac[b], []).append(b)
        for f in sorted(by_f):
            bs = sorted(by_f[f])
            union = frozenset().union(*[sets[b] for b in bs])
            if len(bs) >= 2 and sets[a] < union and not any(sets[a] <= sets[b] for b in bs) and not any(b in spread for b in bs):
                spread[a] = bs
                break
    rows = {}
    for a in sorted(share):
        notes, drop = [], a in spread
        if a in spread:
            notes.append(" + ".join(spread[a]) + " 에 나뉘어 있음")
        for b in sorted(share[a], key=lambda x: (-share[a][x], x)):
            n, m = share[a][b], len(sets[a])
            if sets[a] == sets[b]:
                notes.append(f"같은 사업: {b}")
                drop = drop or (keep_key(b) < keep_key(a) and b not in spread)
            elif sets[a] < sets[b]:
                notes.append(f"{b} 의 일부")
                drop = drop or b not in spread
            elif sets[a] > sets[b]:
                notes.append(f"{b} 를 포함({n}/{m}건)")
            else:
                notes.append(f"{b} 와 일부 겹침({n}/{m}건)")
        rows[a] = {"시설간_중복": "; ".join(notes), "중복_제외권장": "Y" if drop else ""}
    # 안전장치: Y 행을 모두 뺐을 때 어느 프로젝트에도 남지 않는 공고가 생기면 그 프로젝트의 권장을 거둔다(권장을 따라도 공고가 사라지지 않게).
    # 지금의 권장 규칙에서는 걸리는 경우가 없다(무작위 구조 12만 개와 실제 자료에서 0건) — 권장 규칙을 바꿀 때를 위한 보호 장치
    dropped = {a for a, r in rows.items() if r["중복_제외권장"] == "Y"}
    while True:
        kept_nos = set().union(*[sets[q] for q in sets if q not in dropped]) if len(dropped) < len(sets) else set()
        back = next((a for a in sorted(dropped) if not sets[a] <= kept_nos), None)
        if back is None:
            break
        dropped.discard(back)
        rows[back]["중복_제외권장"] = ""
    return pd.DataFrame.from_dict(rows, orient="index", columns=cols)
