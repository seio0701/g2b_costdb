"""S4 중복·재발주 정리.

프로젝트키 = (시설ID, 사업유형, 공종[, 단계토큰]) — 같은 시설의 신축과 리모델링은 별도 프로젝트
 0) 공종 공고의 사업유형이 '미분류'(예: 'OO미술관 전기공사')면 같은 시설에서 공고일이 가장 가까운 건축 공고의 사업유형을 상속
 1) 동일 (시설, 공고번호) 내 최대 차수(변경공고) 채택 — 같은 공고가 두 시설에 걸려도 서로 영향 없음
 2) 취소공고 제외 후, 동일 프로젝트키에서 공고일시 최신 1건을 대표로 채택 (재공고·재발주 처리)
    단, 추정가격이 프로젝트키 내 최대치의 minor_ratio 미만인 공고(부대·소규모 공사 의심)는 대표가 되지 않는다
 3) 대체된 공고에는 '대체됨' + 대체공고번호 기록, 금액 급변(±30%) 시 경고
"""
from __future__ import annotations

import logging
import re
from typing import List, Optional, Tuple

import pandas as pd

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
    vc = s[s.map(lambda v: v is not None and str(v).strip() not in ("", "nan", "None"))].value_counts()
    return str(vc.index[0]) if len(vc) else default


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
            j = gap.idxmin() if gap.notna().any() else anchors.index[0]
            out.at[i] = anchors.at[j, "사업유형"] if (pd.isna(gap.at[j]) or gap.at[j] <= max_days) \
                else _mode(typed["사업유형"], "신축")
    return out


def dedup_latest(df: pd.DataFrame, price_change_warn_ratio: float = 0.30,
                 minor_ratio: float = 0.30) -> Tuple[pd.DataFrame, pd.DataFrame, List[dict]]:
    """반환: (전체 이력 df[최신여부·대체공고번호 부여], 최신만 df, 검증로그 rows)"""
    if df.empty:
        return df, df, []
    d = df.copy()
    d["단계토큰"] = d["공고명"].map(_phase_token)
    d["사업유형_원분류"] = d["사업유형"]
    d["사업유형"] = inherit_work_type(d)
    d["프로젝트ID"] = d["시설ID"].astype(str) + "-" + d["사업유형"].map(lambda t: WORK_TYPE_CODE.get(t, "U"))
    d["프로젝트키"] = d["프로젝트ID"] + "|" + d["공종"].astype(str) + "|" + d["단계토큰"]
    d["_ord"] = pd.to_numeric(d["공고차수"], errors="coerce").fillna(0)
    d["_dt"] = pd.to_datetime(d["공고일시"].astype(str).str.strip().str[:19], errors="coerce")
    if d["_dt"].isna().any():
        n_bad = int(d["_dt"].isna().sum())
        log.warning("공고일시를 해석할 수 없는 공고 %d건 — 대표 선정 시 가장 오래된 것으로 취급", n_bad)
    d["_dt"] = d["_dt"].fillna(pd.Timestamp("1900-01-01"))
    d["_price"] = pd.to_numeric(d["추정가격"], errors="coerce")

    # 1) (시설, 공고번호)별 최대 차수
    d = d.sort_values(["시설ID", "공고번호", "_ord", "_dt"])
    d["최신차수여부"] = ~d.duplicated(["시설ID", "공고번호"], keep="last")

    # 2) 프로젝트키별 최신(취소 제외)
    d["최신여부"] = False
    d["대체공고번호"] = ""
    d["대표선정사유"] = ""
    logs: List[dict] = []
    for key, g in d[d["최신차수여부"]].groupby("프로젝트키", sort=False):
        valid = g[g["공고종류"] != "취소"].sort_values("_dt")
        if valid.empty:
            for i in g.index:
                d.at[i, "대표선정사유"] = "전부 취소공고"
            logs.append({"공고번호": ", ".join(g["공고번호"].astype(str)), "항목": "대표선정", "판정": "경고",
                         "비고": f"{key}: 유효 공고 없음(모두 취소)"})
            continue
        # 본공사 보호: 같은 프로젝트키 안에서 추정가격이 최대치의 minor_ratio 미만인 공고(부대·소규모 공사 의심)는 대표가 되지 않는다
        pmax = valid["_price"].max()
        minor = valid[valid["_price"].notna() & (valid["_price"] < pmax * minor_ratio)] if pd.notna(pmax) and pmax > 0 else valid.iloc[0:0]
        main = valid.drop(minor.index)
        rep = main.iloc[-1]
        d.at[rep.name, "최신여부"] = True
        n_valid, n_cancel = len(main), int((g["공고종류"] == "취소").sum())
        reason = "단독 공고" if n_valid == 1 else f"최신 공고 채택({n_valid}건 중)"
        d.at[rep.name, "대표선정사유"] = reason + (f", 취소 {n_cancel}건 제외" if n_cancel else "") + \
            (f", 소액 {len(minor)}건 별도" if len(minor) else "")
        for i, row in g.iterrows():
            if i == rep.name:
                continue
            d.at[i, "대체공고번호"] = rep["공고번호"]
            if row["공고종류"] == "취소":
                d.at[i, "대표선정사유"] = "취소공고"
            elif i in minor.index:
                d.at[i, "대표선정사유"] = "소액 공고(부대공사 의심) — 대표 제외, 검토 필요"
                logs.append({"공고번호": row["공고번호"], "항목": "소액 공고 분리", "API값": _num(row.get("추정가격")), "문서값": _num(rep.get("추정가격")),
                             "판정": "경고", "비고": f"{key}: '{row['공고명']}' 추정가격이 대표({rep['공고번호']})의 {minor_ratio:.0%} 미만 → 부대공사로 보고 대표에서 제외. "
                                                "본공사가 맞으면 keywords.yaml/검수 결과를 확인"})
                continue
            else:
                d.at[i, "대표선정사유"] = "대체됨(이후 재공고/재발주)"
            # 금액 급변 경고
            p0, p1 = _num(row.get("추정가격")), _num(rep.get("추정가격"))
            if p0 and p1 and p0 > 0:
                ratio = (p1 - p0) / p0
                if abs(ratio) > price_change_warn_ratio:
                    logs.append({"공고번호": rep["공고번호"], "항목": "재발주 금액변동", "API값": p0, "문서값": p1,
                                 "차이": p1 - p0, "차이율": round(ratio, 4), "판정": "경고",
                                 "비고": f"{key}: 선행 {row['공고번호']} 대비 {ratio:+.1%} → 설계변경 의심"})
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

    latest = d[d["최신여부"]].copy()
    d = d.drop(columns=["_ord", "_dt", "_price"])
    latest = latest.drop(columns=["_ord", "_dt", "_price"])
    return d, latest, logs


def facility_summary(latest: pd.DataFrame, all_hist: pd.DataFrame) -> pd.DataFrame:
    """프로젝트(시설×사업유형) 단위 요약: 공종 커버리지·공고건수·기간. 신축과 리모델링은 별도 행."""
    trades = ["건축", "전기", "정보통신", "소방", "조경", "기계설비", "토목", "기타"]
    rows = []
    if latest.empty:
        return pd.DataFrame(columns=["프로젝트ID", "시설ID", "시설명", "사업유형", "표2_대분류", "표2_중분류", "검색어", "수요기관",
                                     "공사현장지역", "공고건수_전체", "공고건수_최신", "최초공고일", "최종공고일"] + [f"공종_{t}" for t in trades])
    for pid, g in latest.groupby("프로젝트ID"):
        h = all_hist[all_hist["프로젝트ID"] == pid]
        row = {
            "프로젝트ID": pid, "시설ID": g["시설ID"].iloc[0], "시설명": g["시설명"].iloc[0], "사업유형": g["사업유형"].iloc[0],
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
