"""S2 광역 탐색(로컬) + 시설명 추출 + 검수용 Excel / S3 시설명 재검색.

전량 수집된 notices_all.parquet 위에서만 동작하므로 API 호출이 없다.
- 시설 후보 단위 = (표2 분류, 검색어, 시설키, 수요기관). 같은 이름의 시설이 여러 지자체에 있어도 섞이지 않도록 수요기관을 키에 포함한다.
- discover 를 다시 실행해도 기존 facility_candidates.xlsx 의 검수 내용(노란 셀)과 시설ID는 (시설키, 수요기관) 기준으로 유지된다.
"""
from __future__ import annotations

import difflib
import logging
import os
import re
from typing import Dict, List, Optional

import pandas as pd
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, PatternFill
from openpyxl.utils import get_column_letter

from .classify import (classify_notice_kind, classify_trade, classify_work_type, extract_facility_name, load_keywords,
                       match_categories, normalize_facility_key, parse_amount, prefix_institution, strip_institution_prefix,
                       work_type_for)

log = logging.getLogger(__name__)

YELLOW = PatternFill("solid", fgColor="FFFF00")
HEADER_FILL = PatternFill("solid", fgColor="DDEBF7")
FONT = Font(name="Arial", size=10)
BOLD = Font(name="Arial", size=10, bold=True)
REVIEW_COLS = ["검수_포함여부", "검수_시설명", "검수_별칭(;구분)", "검수_수요기관", "검수_메모"]
STATUS_COL = "갱신_표시"            # discover 를 다시 돌렸을 때 이전 검수 파일과 달라진 점(신규·확인·후보 이름 바뀜·이번 후보에 없음). 검수 칸이 아니다
PREV_MARK = "지난번: "               # 이전 실행 때 붙은 표시가 아직 확인되지 않은 것으로 보여 그대로 넘긴 것
_EMPTY_STRS = {"", "nan", "none", "null", "nat", "<na>"}


def _s(series: pd.Series) -> pd.Series:
    """원문 값 → 문자열('nan'/'None'/공백은 공란). 한 번의 map 으로 처리(129만 행 규모에서 두 번 map 은 비용이 큼)."""
    def conv(v):
        if v is None or (isinstance(v, float) and v != v):
            return ""
        t = str(v)
        return "" if t.strip().lower() in _EMPTY_STRS else t
    return series.map(conv)


def _blank(v) -> bool:
    return v is None or (isinstance(v, float) and v != v) or str(v).strip().lower() in _EMPTY_STRS


def raw_columns_needed(cfg: dict) -> List[str]:
    """standardize 가 실제로 읽는 원문 컬럼(전량 parquet 는 145개 컬럼 × 백만 행이므로 필요한 것만 읽어 메모리를 줄인다)."""
    f = cfg["fields"]
    keys = ["bid_no", "bid_ord", "name", "kind", "re_notice", "reg_type", "notice_dt", "close_dt", "open_dt", "notice_inst",
            "demand_inst", "presmpt_price", "main_cnstty", "site_region", "detail_url", "ref_no", "prespec_no",
            "demand_inst_cd", "bdgt_amt", "vat", "govsply_amt", "contractor_govsply", "gov_govsply", "prev_bid_no", "chg_reason",
            "sucsfbid_rate", "std_notice_url"]
    cols = [f[k] for k in keys if isinstance(f.get(k), str) and f[k]]
    cols += [f"{f.get('sub_cnstty_prefix', '')}{i}" for i in range(1, int(f.get("sub_cnstty_max", 0) or 0) + 1)]
    cols += [f"{f.get('spt_url_prefix', '')}{i}" for i in range(1, int(f.get("spt_max", 0) or 0) + 1)]
    n = int(f["attach_max"])
    cols += [f"{f['attach_url_prefix']}{i}" for i in range(1, n + 1)] + [f"{f['attach_name_prefix']}{i}" for i in range(1, n + 1)]
    return list(dict.fromkeys(cols))


def standardize(df_raw: pd.DataFrame, cfg: dict, bsis: Optional[pd.DataFrame] = None,
                license_df: Optional[pd.DataFrame] = None, awards: Optional[pd.DataFrame] = None) -> pd.DataFrame:
    """API 원문 → 내부 표준 컬럼. 미존재 필드는 공란."""
    f = cfg["fields"]
    empty = pd.Series([""] * len(df_raw), index=df_raw.index, dtype=object)
    g = lambda k: _s(df_raw[f[k]]) if f[k] in df_raw.columns else empty
    df = pd.DataFrame({
        "공고번호": g("bid_no"),
        "공고차수": g("bid_ord"),
        "공고명": g("name"),
        "공고종류_원문": g("kind"),
        "재공고여부": g("re_notice"),
        "등록유형": g("reg_type"),
        "공고일시": g("notice_dt"),
        "입찰마감일시": g("close_dt"),
        "개찰일시": g("open_dt"),
        "공고기관": g("notice_inst"),
        "수요기관": g("demand_inst"),
        "추정가격": g("presmpt_price").map(parse_amount),
        "주공종명": g("main_cnstty"),
        "공사현장지역": g("site_region"),
        "상세URL": g("detail_url"),
        "참조번호": g("ref_no"),
        "사전규격번호": g("prespec_no"),
    })
    # 실제 응답에서 확인된 추가 필드(없으면 공란)
    opt = lambda k: _s(df_raw[f[k]]) if (k in f and f[k] in df_raw.columns) else empty
    df["수요기관코드"] = opt("demand_inst_cd")
    df["예산금액"] = opt("bdgt_amt").map(parse_amount)
    df["부가세_API"] = opt("vat").map(parse_amount)
    df["관급자재_API"] = opt("govsply_amt").map(parse_amount)
    df["도급자관급액_API"] = opt("contractor_govsply").map(parse_amount)
    df["관급자관급액_API"] = opt("gov_govsply").map(parse_amount)
    df["이전공고번호"] = opt("prev_bid_no")
    df["변경사유"] = opt("chg_reason")
    df["낙찰하한율"] = opt("sucsfbid_rate")
    df["표준공고문URL"] = opt("std_notice_url")
    subs = []
    for i in range(1, int(f.get("sub_cnstty_max", 0) or 0) + 1):
        c = f"{f.get('sub_cnstty_prefix', '')}{i}"
        if c in df_raw.columns:
            subs.append(_s(df_raw[c]))
    df["부공종명"] = [" / ".join(x for x in row if x) for row in zip(*subs)] if subs else ""
    for i in range(1, int(f.get("spt_max", 0) or 0) + 1):
        c = f"{f.get('spt_url_prefix', '')}{i}"
        df[f"현장설명서URL{i}"] = _s(df_raw[c]) if c in df_raw.columns else ""
    # 첨부파일
    for i in range(1, int(f["attach_max"]) + 1):
        u, nm = f"{f['attach_url_prefix']}{i}", f"{f['attach_name_prefix']}{i}"
        df[f"첨부URL{i}"] = _s(df_raw[u]) if u in df_raw.columns else ""
        df[f"첨부파일명{i}"] = _s(df_raw[nm]) if nm in df_raw.columns else ""
    df["첨부파일수"] = sum((df[f"첨부URL{i}"].str.len() > 4).astype(int) for i in range(1, int(f["attach_max"]) + 1))
    df["공고차수"] = df["공고차수"].map(lambda s: s.zfill(3) if s.isdigit() else (s or "000"))   # 나라장터 차수는 '000' 3자리
    df["공고키"] = df["공고번호"] + "-" + df["공고차수"]

    # 기초금액 병합(공고번호+차수 기준, 없으면 공고번호)
    df["기초금액"] = None
    if bsis is not None and not bsis.empty and f["bid_no"] in bsis.columns:
        b = bsis.copy()
        if "_bsis_amount" not in b.columns:
            b["_bsis_amount"] = b[f["bsis_amount"]].map(parse_amount) if f["bsis_amount"] in b.columns else None
        ords = _s(b[f["bid_ord"]]) if f["bid_ord"] in b.columns else pd.Series(["000"] * len(b), index=b.index)
        b["_k"] = _s(b[f["bid_no"]]) + "-" + ords.map(lambda s: s.zfill(3) if s.isdigit() else (s or "000"))
        b = b[b["_bsis_amount"].notna()]
        m = b.drop_duplicates("_k", keep="last").set_index("_k")["_bsis_amount"]
        df["기초금액"] = df["공고키"].map(m)
        m2 = b.assign(_no=_s(b[f["bid_no"]])).drop_duplicates("_no", keep="last").set_index("_no")["_bsis_amount"]
        df["기초금액"] = df["기초금액"].fillna(df["공고번호"].map(m2))

    # 낙찰정보 병합(참고 컬럼: 공고번호+차수 → 없으면 공고번호). 같은 공고에 여러 행이면 개찰일시가 최신인 행
    for c in ("낙찰금액_API", "낙찰률_API", "낙찰자", "참가업체수", "낙찰개찰일시"):
        df[c] = None
    if awards is not None and not awards.empty and f["bid_no"] in awards.columns:
        from .collect import award_column
        w = awards.copy()
        amt_c, rate_c = award_column(cfg, w, "award_amt"), award_column(cfg, w, "award_rate")
        bid_c, cnt_c, dt_c = award_column(cfg, w, "award_bidder"), award_column(cfg, w, "award_prtcpt_cnt"), award_column(cfg, w, "award_open_dt")
        w["_amt"] = w["_award_amt"] if "_award_amt" in w.columns else (w[amt_c].map(parse_amount) if amt_c else None)
        w["_rate"] = w["_award_rate"] if "_award_rate" in w.columns else \
            (pd.to_numeric(_s(w[rate_c]).str.replace("%", "").str.replace(",", ""), errors="coerce") if rate_c else None)
        w["_bidder"] = _s(w[bid_c]) if bid_c else ""
        w["_cnt"] = pd.to_numeric(_s(w[cnt_c]), errors="coerce") if cnt_c else None
        w["_dt"] = _s(w[dt_c]) if dt_c else ""
        ords = _s(w[f["bid_ord"]]) if f["bid_ord"] in w.columns else pd.Series(["000"] * len(w), index=w.index)
        w["_no"] = _s(w[f["bid_no"]])
        w["_k"] = w["_no"] + "-" + ords.map(lambda s: s.zfill(3) if s.isdigit() else (s or "000"))
        w = w[w["_amt"].notna() | w["_rate"].notna()].sort_values("_dt")
        for key_col, src in (("_k", df["공고키"]), ("_no", df["공고번호"])):
            m = w.drop_duplicates(key_col, keep="last").set_index(key_col)
            for c, wc in (("낙찰금액_API", "_amt"), ("낙찰률_API", "_rate"), ("낙찰자", "_bidder"), ("참가업체수", "_cnt"), ("낙찰개찰일시", "_dt")):
                filled = src.map(m[wc])
                df[c] = df[c].where(df[c].notna(), filled)

    # 면허제한 업종(복수 → ' / ' 결합)
    df["면허제한업종"] = ""
    if license_df is not None and not license_df.empty and f["license_name"] in license_df.columns:
        lic = license_df.groupby("_bid_no")[f["license_name"]].apply(lambda s: " / ".join(sorted(set(map(str, s)))))
        df["면허제한업종"] = df["공고번호"].map(lic).fillna("")

    df["공고종류"] = [classify_notice_kind(k, r, t, n) for k, r, t, n in
                   zip(df["공고종류_원문"], df["재공고여부"], df["등록유형"], df["공고명"])]
    tr = [classify_trade(n, m, l) for n, m, l in zip(df["공고명"], df["주공종명"], df["면허제한업종"])]
    df["공종"] = [t[0] for t in tr]
    df["공종근거"] = [t[1] for t in tr]
    df["사업유형"] = df["공고명"].map(classify_work_type)
    return df


def _mode(s: pd.Series, default: str = "") -> str:
    vc = s[s.map(lambda v: not _blank(v))].value_counts()
    return str(vc.index[0]) if len(vc) else default


_PROJECT_TYPE_PRIORITY = ("신축", "증축·리모델링", "증축", "리모델링")


def _facility_work_type(s: pd.Series) -> str:
    """시설의 사업유형: 공고 하나라도 신축·증축·리모델링이면 그 유형(우선순위 순). 유지보수 공고가 아무리 많아도 본공사 1건을 묻지 않는다.
    프로젝트 유형이 없으면 미분류/유지보수 중 최빈값."""
    present = set(s.dropna().astype(str))
    for t in _PROJECT_TYPE_PRIORITY:
        if t in present:
            return t
    return _mode(s, "미분류")


def discover_candidates(std: pd.DataFrame, previous_review: Optional[pd.DataFrame] = None, ids_only: bool = False,
                        recalc: bool = False) -> pd.DataFrame:
    """공고명에 표2 검색어가 포함된 공고 → 시설명 후보 + 분류 (검수용).
    previous_review: 이전 facility_candidates.xlsx 내용. 있으면 기본은 **검수 유지**(_keep_review) — 이전 행의 시설ID·검수 칸을 그대로 두고
    새 시설만 더하며, 달라진 점은 `갱신_표시` 열에 적는다(검수를 마친 뒤 공고가 늘어 다시 돌릴 때).
    recalc: 손대지 않은 검수 칸은 새 규칙의 기본값으로 다시 계산(`discover --recalc` — 검수를 시작하기 전 규칙이 바뀌었을 때).
    ids_only: 시설ID만 이어받고 검수 컬럼은 모두 새 기본값으로(`discover --fresh`)."""
    kw = load_keywords()
    rows = []
    # 공고명만 순회하고(iterrows 는 행마다 Series 를 만들어 백만 행 규모에서 느림) 매칭된 행만 전체 컬럼을 읽는다
    for i, name in enumerate(std["공고명"].tolist()):
        matches = match_categories(name, kw)
        if not matches:
            continue
        base = std.iloc[i].to_dict()
        for major, minor, w in matches:
            fac = prefix_institution(extract_facility_name(name, w), w, str(base.get("수요기관") or ""))   # 단지명 없으면 지자체명 접두
            rows.append({**base, "표2_대분류": major, "표2_중분류": minor, "검색어": w,
                         "시설명_후보": fac, "시설키": normalize_facility_key(fac),
                         "사업유형": work_type_for(name, w, fac)})   # 시설 자체가 주차장 등이면 부속어에서 제외
    cand = pd.DataFrame(rows)
    if cand.empty:
        return cand
    cand["수요기관"] = _s(cand["수요기관"])
    # 시설 단위 요약 (검수 대상은 공고가 아니라 시설). 같은 이름이라도 수요기관이 다르면 다른 시설로 본다.
    agg = (cand.groupby(["표2_대분류", "표2_중분류", "검색어", "시설키", "수요기관"], as_index=False)
           .agg(시설명_후보=("시설명_후보", lambda s: _mode(s)),
                공고건수=("공고번호", "nunique"),
                최초공고일=("공고일시", "min"), 최종공고일=("공고일시", "max"),
                사업유형=("사업유형", _facility_work_type),
                사업유형_분포=("사업유형", lambda s: ", ".join(f"{k} {v}" for k, v in s.value_counts().items())),
                공종목록=("공종", lambda s: "/".join(sorted(set(s)))),
                추정가격_합계=("추정가격", lambda s: pd.to_numeric(s, errors="coerce").fillna(0).sum()),
                공고명_예시=("공고명", "first")))
    agg = agg[["표2_대분류", "표2_중분류", "검색어", "시설키", "시설명_후보", "수요기관", "공고건수", "최초공고일", "최종공고일",
               "사업유형", "사업유형_분포", "공종목록", "추정가격_합계", "공고명_예시"]]

    agg["검수_포함여부"] = agg["사업유형"].map(default_include)
    agg["검수_시설명"] = agg["시설명_후보"]
    agg["검수_별칭(;구분)"] = ""
    agg["검수_수요기관"] = agg["수요기관"]
    agg["검수_메모"] = ""
    own_cols = list(agg.columns)
    usable = previous_review is not None and not previous_review.empty and {"시설키", "시설ID"} <= set(previous_review.columns)
    if ids_only or recalc or not usable:      # 시설키·시설ID 열이 없는 파일은 맞출 기준이 없다 → 새로 만든다(pipeline 은 그 전에 review_file_problem 으로 멈춘다)
        if review_file_problem(previous_review):
            log.warning("이전 검수 파일의 %s → 시설ID·검수 내용을 이어받지 못합니다", review_file_problem(previous_review))
        agg = _carry_over_review(agg, previous_review, ids_only=ids_only)
        agg[STATUS_COL] = ""
    else:
        agg = _keep_review(agg, previous_review)
    head = [c for c in own_cols if c not in ["시설ID", STATUS_COL] + REVIEW_COLS]
    extra = [c for c in agg.columns if c not in own_cols + ["시설ID", STATUS_COL]]          # 검수자가 더한 열은 노란 칸 뒤(원래 있던 자리)에
    cols = ["시설ID"] + head + [STATUS_COL] + REVIEW_COLS + extra
    return agg[cols].reset_index(drop=True)


def review_file_problem(prev: Optional[pd.DataFrame]) -> str:
    """이전 검수 파일을 이어받을 수 없는 까닭(이전 파일이 없거나 이어받을 수 있으면 '').
    첫 줄(머리글)에서 맞추는 기준(시설ID·시설키·수요기관)이나 검수 열을 찾지 못하면 그 내용을 이어받을 수 없다 — 열을 지웠거나 이름을 고쳤을 때,
    첫 줄 위에 줄을 넣었을 때, 정렬하다 머리글 줄이 자료 사이로 들어갔을 때. 같은 이름의 열이 둘이면(열을 복사) 어느 쪽이 검수 값인지 알 수 없다.
    검수자가 더한 열은 이름이 있어야 옮길 수 있으므로, 이름 없는 열에 글이 있을 때도 알린다(말없이 빼지 않는다).
    그대로 다시 만들면 검수 칸이 기본값으로 돌아가므로 부르는 쪽(pipeline.stage_discover)이 덮어쓰기 전에 멈춘다."""
    if prev is None or prev.empty:
        return ""
    names = [str(c) for c in prev.columns]
    miss = [c for c in ["시설ID", "시설키", "수요기관"] + REVIEW_COLS if c not in names]
    dup = sorted({c for c in names if c != "None" and names.count(c) > 1})
    why = []
    if miss:
        why.append("첫 줄(머리글)에 " + ", ".join(f"'{c}'" for c in miss) + " 열이 없습니다")
    if dup:
        why.append("첫 줄(머리글)에 " + ", ".join(f"'{c}'" for c in dup) + " 열이 두 번 있습니다")
    orphan = sum(1 for k, c in enumerate(names) if c == "None" for v in prev.iloc[:, k] if not _blank(v))
    if orphan and not miss:                          # 이름 없는 열은 새 파일로 옮길 수 없다(머리글 줄 자체가 밀린 경우는 위에서 이미 알렸다)
        why.append(f"첫 줄(머리글)에 이름이 없는 열에 글이 {orphan}칸 있습니다(그 열의 첫 줄에 이름을 적어 주세요. 이름이 있어야 새 파일로 옮겨집니다)")
    return " / ".join(why)


def default_include(work_type: str) -> str:
    """사업유형 → 검수_포함여부 기본값."""
    if work_type in ("신축", "증축", "리모델링", "증축·리모델링"):
        return "포함"
    return "제외" if work_type == "유지보수" else "검토필요"


def _user_edited(prev_row, col: str) -> bool:
    """이전 파일의 검수 값이 사용자가 고친 것인지(기본값 그대로면 False).
    기본값이었다면 새 규칙으로 다시 계산한 값을 쓰고, 고친 값만 이어받는다."""
    v = prev_row.get(col)
    if _blank(v):
        return False
    v = str(v).strip()
    if col == "검수_포함여부" and not _blank(prev_row.get("사업유형")):
        return v != default_include(str(prev_row.get("사업유형")).strip())
    if col == "검수_시설명" and not _blank(prev_row.get("시설명_후보")):
        return v != str(prev_row.get("시설명_후보")).strip()
    if col == "검수_수요기관" and not _blank(prev_row.get("수요기관")):
        return v != str(prev_row.get("수요기관")).strip()
    return True          # 별칭·메모(기본값 공란)와, 비교 기준 컬럼이 없는 옛 형식 파일


def _pending_notes(old_flag: str, untouched: bool, include_now: str) -> List[str]:
    """이전 파일의 갱신_표시 가운데 **아직 확인하지 않았을 수 있는** 표시만 PREV_MARK 를 붙여 돌려준다(discover 를 한 번 더 돌려도 볼 행이 사라지지 않게).
    · '신규…'(+ '비슷한 기존 시설'·'이름이 같은 기존 시설' 단서): 검수 칸이 기본값 그대로일 때만(하나라도 고쳤으면 본 것이다)
    · '확인: …(지금 'X' 그대로 둠)': 검수_포함여부가 아직 그 X 일 때만
    · '후보 이름 바뀜'·'시설키 바뀜'·'이번 후보에 없음': 결정할 것이 없거나 매번 새로 붙으므로 넘기지 않는다"""
    parts = [x.strip() for x in str(old_flag or "").split(" / ") if x.strip()]
    parts = [x[len(PREV_MARK):] if x.startswith(PREV_MARK) else x for x in parts]
    was_new = any(x.startswith("신규") for x in parts)
    out = []
    for x in parts:
        if x.startswith("신규") or (x.startswith(("비슷한 기존 시설", "이름이 같은 기존 시설")) and was_new):
            if untouched:
                out.append(PREV_MARK + x)
        elif x.startswith("확인:"):
            m = re.search(r"\(지금 '(.*?)' 그대로 둠\)", x)
            if m and (m.group(1) if m.group(1) != "빈칸" else "") == include_now:
                out.append(PREV_MARK + x)
    return out


def _carry_over_review(agg: pd.DataFrame, prev: Optional[pd.DataFrame], ids_only: bool = False) -> pd.DataFrame:
    """이전 검수 파일의 시설ID와 **사용자가 고친** 검수_* 값을 (시설키, 수요기관) 기준으로 이어받고, 새 시설에는 이어지는 번호를 부여.
    기본값 그대로였던 칸(예: 사업유형 최빈값에 따른 '제외')은 이어받지 않고 새 규칙의 기본값으로 다시 채운다."""
    agg = agg.copy()
    agg["시설ID"] = ""
    if prev is not None and not prev.empty and "시설키" in prev.columns and "시설ID" in prev.columns:
        p = prev.copy()
        p["수요기관"] = _s(p["수요기관"]) if "수요기관" in p.columns else ""
        p = p.drop_duplicates(["시설키", "수요기관"], keep="first").set_index(["시설키", "수요기관"])
        kept, edited = 0, 0
        for i, r in agg.iterrows():
            k = (r["시설키"], r["수요기관"])
            if k in p.index:
                old = p.loc[k]
                agg.at[i, "시설ID"] = str(old["시설ID"])
                any_edit = False
                for c in ([] if ids_only else REVIEW_COLS):
                    if c in p.columns and _user_edited(old, c):
                        agg.at[i, c] = str(old[c]).strip()
                        any_edit = True
                kept += 1
                edited += int(any_edit)
        log.info("이전 검수 내용 이어받음: %d개 시설(그중 사용자가 고친 값이 있는 시설 %d개, 나머지는 새 기본값)", kept, edited)
    used = set(agg.loc[agg["시설ID"] != "", "시설ID"])
    nums = [int(re.sub(r"\D", "", x) or 0) for x in used]
    nxt = (max(nums) if nums else 0) + 1
    for i in agg.index[agg["시설ID"] == ""]:
        agg.at[i, "시설ID"] = f"F{str(nxt).zfill(4)}"
        nxt += 1
    return agg


def _keep_review(agg: pd.DataFrame, prev: pd.DataFrame) -> pd.DataFrame:
    """검수를 마친 파일을 지키는 이어받기(discover 의 기본 동작). 이전 파일에 있던 행은
      · 시설ID 와 검수 칸(REVIEW_COLS)을 **적혀 있던 그대로** 둔다(기본값과 같아 보여도 다시 계산하지 않는다 — 검수자가 확인하고 둔 값일 수 있다. 빈칸도 빈칸 그대로)
      · 이번 후보에 없어도 지우지 않고 끝에 그대로 남긴다(재검색은 이 파일의 '포함' 행을 읽으므로, 지우면 그 시설이 말없이 DB 에서 빠진다)
      · 규칙이 바뀌어 시설키가 달라졌으면 이전 행의 공고명_예시로 지금의 시설키를 다시 내어 맞춘다(같은 시설이 새 번호의 새 행으로 한 번 더 생기지 않게)
    새 시설에는 이어지는 번호와 기본값을 준다. 달라진 점은 STATUS_COL 에 적는다: '신규' / '확인: …'(사업유형이 달라져 기본값이 지금 값과 다름) /
    '후보 이름 바뀜: …' / '이번 후보에 없음'. 나머지 열(공고건수·사업유형·예시 등)은 새로 계산한 값이다.
    맞추는 순서: ① (표2 분류, 검색어, 시설키, 수요기관) ② (시설키, 수요기관) ③ 공고명_예시로 다시 낸 시설키 — 이전 행 하나는 새 행 하나에만 쓴다.
    그래도 남은 새 행 가운데 (시설키, 수요기관)이 같은 이전 행이 있는 것은 그 시설의 다른 검색어 행이므로 시설ID·검수 값을 그 행과 같게 둔다
    (기본값 '포함' 이 붙어, 검수자가 제외한 시설이 다시 들어가지 않게)."""
    agg = agg.copy().reset_index(drop=True)
    agg["시설ID"] = ""
    agg[STATUS_COL] = ""
    txt = lambda v: "" if _blank(v) else str(v).strip()
    p = prev.copy().reset_index(drop=True)
    have = [c for c in REVIEW_COLS if c in p.columns]                 # 이전 파일에 없는 검수 열(옛 형식)은 새 기본값을 쓴다
    for c in ("시설ID", "시설키", "수요기관", "표2_대분류", "표2_중분류", "검색어", "시설명_후보", "사업유형", "공고명_예시"):
        if c not in p.columns:
            p[c] = None
    p_org = list(_s(p["수요기관"]))
    p_key = [txt(v) for v in p["시설키"]]
    p_id = [txt(v) for v in p["시설ID"]]
    p5 = [(txt(a), txt(b), txt(w), k, o) for a, b, w, k, o in zip(p["표2_대분류"], p["표2_중분류"], p["검색어"], p_key, p_org)]
    live = [j for j in range(len(p)) if p_id[j] or p_key[j] or any(txt(p.at[j, c]) for c in have + ["시설명_후보"])]   # 통째로 빈 행(엑셀 끝의 서식만 있는 줄)은 버린다
    by5: Dict[tuple, list] = {}
    by2: Dict[tuple, list] = {}
    for j in live:
        by5.setdefault(p5[j], []).append(j)
        by2.setdefault((p_key[j], p_org[j]), []).append(j)
    a_org = list(agg["수요기관"])
    a_key = [txt(v) for v in agg["시설키"]]
    a5 = [(txt(a), txt(b), txt(w), k, o) for a, b, w, k, o in zip(agg["표2_대분류"], agg["표2_중분류"], agg["검색어"], a_key, a_org)]
    match: Dict[int, int] = {}          # 새 행 → 이전 행
    how: Dict[int, str] = {}
    used: set = set()

    def take(i: int, cands: list, kind: str) -> bool:
        for j in cands:
            if j not in used:
                match[i], how[i] = j, kind
                used.add(j)
                return True
        return False
    for i in range(len(agg)):                                           # ① 묶음 열쇠가 모두 같은 행
        take(i, by5.get(a5[i], []), "same")
    for i in range(len(agg)):                                           # ② (시설키, 수요기관)이 같은 행(표2 분류·검색어만 달라진 경우)
        if i not in match:
            take(i, by2.get((a_key[i], a_org[i]), []), "key")
    free: Dict[tuple, list] = {}                                        # ③ 규칙이 바뀌어 시설키가 달라진 이전 행 → 예시 공고명으로 지금의 시설키를 다시 낸다
    for i in range(len(agg)):
        if i not in match:
            free.setdefault((a_key[i], a_org[i]), []).append(i)
    for j in live:
        if j in used or not txt(p.at[j, "공고명_예시"]) or not txt(p.at[j, "검색어"]):
            continue
        w = txt(p.at[j, "검색어"])
        try:
            k2 = normalize_facility_key(prefix_institution(extract_facility_name(txt(p.at[j, "공고명_예시"]), w), w, p_org[j]))
        except Exception:  # noqa: BLE001 — 사람이 고친 셀에서 나온 값이라 무엇이든 올 수 있다. 맞추지 못하면 그 행은 그대로 남긴다
            continue
        cands = [i for i in free.get((k2, p_org[j]), []) if i not in match]
        cands = [i for i in cands if a5[i][2] == w] or cands
        if cands and k2 != p_key[j]:
            match[cands[0]], how[cands[0]] = j, "rekey"
            used.add(j)
    nums = [int(d) for d in (re.sub(r"\D", "", x) for x in p_id) if d and len(d) <= 6]   # 번호 칸에 날짜 같은 긴 숫자가 들어가 있어도 다음 번호가 그 뒤로 튀지 않게
    nxt = (max(nums) if nums else 0) + 1                                 # 이전 파일의 모든 번호 뒤부터(남겨 둔 행의 번호를 새 시설에 다시 쓰지 않는다)
    prev_last = max((txt(v) for v in p["최종공고일"]), default="") if "최종공고일" in p.columns else ""   # 이전 파일을 만들 때 가장 늦은 공고일
    by_org: Dict[str, list] = {}
    by_key: Dict[str, list] = {}
    for j in live:
        if p_key[j]:
            by_org.setdefault(p_org[j], []).append(j)
            by_key.setdefault(p_key[j], []).append(j)

    def similar(i: int) -> str:
        """새 행과 이름이 비슷한 기존 시설(같은 수요기관), 이름이 같은데 수요기관 이름만 다른 기존 시설 — 같은 시설이 한 번 더 잡힌 것인지 볼 단서. 없으면 ''."""
        best, bj = 0.0, None
        for j in by_org.get(a_org[i], []):
            r = 1.0 if (a_key[i] in p_key[j] or p_key[j] in a_key[i]) and min(len(a_key[i]), len(p_key[j])) >= 3 \
                else difflib.SequenceMatcher(None, a_key[i], p_key[j]).ratio()
            if r > best:
                best, bj = r, j
        hints = [f"비슷한 기존 시설: {p_id[bj] or '?'} '{txt(p.at[bj, '시설명_후보']) or p_key[bj]}'"] if bj is not None and best >= 0.6 else []
        other = [j for j in by_key.get(a_key[i], []) if p_org[j] != a_org[i]]   # 기관 이름이 바뀌었거나(강원도→강원특별자치도), 지자체·공사로 나뉜 같은 시설
        if other:
            hints.append(f"이름이 같은 기존 시설: {p_id[other[0]] or '?'} (수요기관 '{p_org[other[0]] or '빈칸'}')")
        return ", ".join(hints)
    owner: Dict[tuple, int] = {}                                         # 지금의 (시설키, 수요기관) → 그 시설의 이전 행(먼저 맞춘 것)
    for i in sorted(match):
        owner.setdefault((a_key[i], a_org[i]), match[i])
    n_same = n_rekey = n_new = n_share = n_check = n_old_new = 0
    for i in range(len(agg)):
        j = match.get(i)
        share = None
        if j is None:
            share = owner.get((a_key[i], a_org[i]))                       # 같은 시설의 다른 검색어 행
        src = j if j is not None else share
        notes = []
        if src is None:
            agg.at[i, "시설ID"] = f"F{str(nxt).zfill(4)}"
            nxt += 1
            n_new += 1
            old_notice = bool(prev_last) and bool(txt(agg.at[i, "최초공고일"])) and txt(agg.at[i, "최초공고일"]) <= prev_last
            n_old_new += old_notice
            notes.append("신규(예전 공고가 있음: 뒤늦게 수집된 공고이거나, 지웠던 줄이거나, 이름 정리 규칙이 바뀌어 새로 잡힌 후보일 수 있음)" if old_notice else "신규")
            hint = similar(i)
            if hint:
                notes.append(hint + " → 같은 시설이면 '제외' 하거나 시설ID 를 그 번호로")
        else:
            if p_id[src]:
                agg.at[i, "시설ID"] = p_id[src]
            else:
                agg.at[i, "시설ID"] = f"F{str(nxt).zfill(4)}"
                nxt += 1
            old_name, new_name = txt(p.at[src, "시설명_후보"]), txt(agg.at[i, "시설명_후보"])
            for c in have:
                agg.at[i, c] = txt(p.at[src, c])
            if "검수_시설명" in have and not txt(p.at[src, "검수_시설명"]) and old_name and old_name != new_name:
                agg.at[i, "검수_시설명"] = old_name                       # 빈칸 = '시설명_후보를 쓴다' 였으므로, 후보 이름이 달라졌으면 쓰이던 이름을 적어 둔다(새 검색어 행도 그 시설의 이름으로)
            if j is None:
                n_share += 1
                notes.append(f"신규 행(같은 시설 {p_id[src] or '?'} 의 다른 검색어: 검수 값은 그 행과 같게 둠)")
            else:
                n_same += how[i] != "rekey"
                n_rekey += how[i] == "rekey"
                if how[i] == "rekey":
                    notes.append(f"후보 이름 바뀜: '{old_name}' → '{new_name}'(이전 줄의 공고명_예시로 다시 맞춤. 다른 시설로 보이면 검수 칸 확인)" if old_name != new_name
                                 else "시설키 바뀜(이름 정리 규칙)")
                old_t, new_t = txt(p.at[src, "사업유형"]), txt(agg.at[i, "사업유형"])
                if "검수_포함여부" in have and old_t and default_include(new_t) != default_include(old_t) \
                        and txt(p.at[src, "검수_포함여부"]) != default_include(new_t):
                    n_check += 1
                    notes.append(f"확인: 사업유형 {old_t}→{new_t}, 새 기본값 '{default_include(new_t)}'(지금 '{txt(p.at[src, '검수_포함여부']) or '빈칸'}' 그대로 둠)")
                if STATUS_COL in p.columns and txt(p.at[src, STATUS_COL]):        # 지난 실행의 표시 가운데 아직 확인하지 않았을 수 있는 것은 남긴다
                    notes += [x for x in _pending_notes(txt(p.at[src, STATUS_COL]), not any(_user_edited(p.loc[src], c) for c in have),
                                                        txt(p.at[src, "검수_포함여부"]) if "검수_포함여부" in have else "")
                              if x[len(PREV_MARK):].split(":")[0] not in {n.split(":")[0] for n in notes}]
        agg.at[i, STATUS_COL] = " / ".join(notes)
    gone = [j for j in live if j not in used]
    if gone:
        rest = pd.DataFrame({c: ([p.at[j, c] for j in gone] if c in p.columns else [""] * len(gone)) for c in agg.columns})
        rest["시설ID"], rest["시설키"], rest["수요기관"] = [p_id[j] for j in gone], [p_key[j] for j in gone], [p_org[j] for j in gone]
        for c in REVIEW_COLS:
            rest[c] = [txt(p.at[j, c]) for j in gone] if c in have else ""
        rest[STATUS_COL] = "이번 후보에 없음(이전 내용 그대로 둠)"
        agg = pd.concat([agg.astype(object), rest.astype(object)], ignore_index=True)
    src_of = [match.get(i) for i in range(len(a_key))] + gone             # 출력 행 → 이전 파일의 행(새로 생긴 행은 None)
    for c in [c for c in prev.columns if c not in agg.columns and str(c) != "None" and list(prev.columns).count(c) == 1]:
        agg[c] = [prev[c].iloc[j] if j is not None else None for j in src_of]   # 검수자가 직접 더한 열(확인 표시·비고 등)도 그 행에 그대로 둔다
    # 같은 시설ID 를 가진 행들(검수자가 합친 시설, 검색어가 여럿인 시설)은 묶음의 첫 행 자리에 모아 이전 파일에 있던 순서대로 둔다. 재검색은 묶음의 첫 '포함' 행에서
    # 대표 이름·표2 분류·검색어를 읽으므로, 줄 순서가 달라지면(정렬해 둔 파일, 끝으로 옮긴 '이번 후보에 없음' 행, 앞에 끼어든 새 검색어 행) 손대지 않은 시설의 이름·분류가 바뀐다
    ids = list(agg["시설ID"])
    first: Dict[str, int] = {}
    for r, fid in enumerate(ids):
        if fid:
            first.setdefault(fid, r)
    order = sorted(range(len(agg)), key=lambda r: (first.get(ids[r], r), src_of[r] is None, src_of[r] or 0, r))
    if order != list(range(len(agg))):
        agg = agg.iloc[order].reset_index(drop=True)
    log.info("이전 검수 파일 %d행을 그대로 이어받음(시설ID·검수 칸은 적혀 있던 그대로): 같은 시설키 %d, 시설키가 바뀌어 예시 공고명으로 맞춘 행 %d, 이번 후보에 없어 남겨 둔 행 %d. "
             "새 시설 %d개(그중 예전 공고로 새로 잡힌 후보 %d개. 같은 시설의 다른 검색어 행 %d개 별도), 사업유형이 달라져 확인할 행 %d개 → '%s' 열",
             len(live), n_same, n_rekey, len(gone), n_new, n_old_new, n_share, n_check, STATUS_COL)
    return agg


def write_review_workbook(agg: pd.DataFrame, path: str) -> None:
    os.makedirs(os.path.dirname(path) or ".", exist_ok=True)
    wb = Workbook()
    ws = wb.active
    ws.title = "시설후보_검수"
    cols = list(agg.columns)
    ws.append(cols)
    for c in range(1, len(cols) + 1):
        ws.cell(row=1, column=c).font = BOLD
        ws.cell(row=1, column=c).fill = HEADER_FILL
    for row in agg.itertuples(index=False):
        ws.append([None if _blank(v) else (v.item() if hasattr(v, "item") else v) for v in row])      # NaN 도 빈 셀로(openpyxl 은 NaN 을 값이 빈 숫자 셀로 쓴다)
    edit_cols = [i + 1 for i, c in enumerate(cols) if c.startswith("검수_")]
    for r in range(2, ws.max_row + 1):
        for c in range(1, len(cols) + 1):
            ws.cell(row=r, column=c).font = FONT
        for c in edit_cols:
            ws.cell(row=r, column=c).fill = YELLOW
    for i, c in enumerate(cols, 1):
        ws.column_dimensions[get_column_letter(i)].width = min(45, max(12, len(str(c)) + 4))
    ws.freeze_panes = "B2"
    ws.auto_filter.ref = ws.dimensions
    legend = wb.create_sheet("범례")
    for line in ["노란색 셀만 수정하세요. (다른 열은 수정해도 반영되지 않습니다. 예외: 시설ID 는 두 줄을 한 시설로 합칠 때 같은 번호로 고칩니다 — 맨 아래 줄 참고)",
                 "첫 줄(머리글)의 이름과 시트 이름은 바꾸지 말고, 열·줄을 지우지 마세요(빼려는 시설은 줄을 지우지 말고 '제외' — 지운 줄은 다음 discover 때 '신규' 로 다시 생깁니다). 정렬·필터·숨기기는 괜찮습니다.",
                 "검수_포함여부: '포함' → DB 수집 대상 / '제외' → 제외 / '검토필요' → 판단 후 둘 중 하나로 바꾸기",
                 "기본값: 신축·증축·리모델링(증축·리모델링 포함)은 '포함', 유지보수(방수·도색·교체 등)는 '제외', 미분류는 '검토필요'",
                 "같은 시설에 신축과 리모델링이 모두 있으면 한 줄로 보입니다(사업유형_분포 참고). '포함'이면 두 사업이 각각 별도 프로젝트로 DB에 들어갑니다.",
                 "검수_시설명: 재검색에 사용할 정확한 시설명 (예: 수원시립미술관). 비워 두면 시설명_후보를 사용합니다.",
                 "검수_별칭(;구분): 재검색 보조어를 ';'로 구분 (예: 수원시립미술관;수원시 미술관;수원미술관)",
                 "검수_수요기관: 재검색 시 이 수요기관의 공고만 찾습니다(같은 이름의 시설이 다른 지자체에 있어도 섞이지 않게). 비우면 이름만으로 찾습니다.",
                 "사업유형이 '유지보수'로 잡혔지만 실제 리모델링·신축이면 '포함'으로 바꾸고 메모를 남기세요.",
                 "discover 를 다시 실행해도 이 파일에 있던 행의 시설ID 와 노란 칸은 적혀 있던 그대로 남습니다(덮어쓰기 전 사본도 같은 폴더에 남깁니다). "
                 "이번 후보에 없는 행도 지우지 않습니다.",
                 f"{STATUS_COL}: 다시 실행했을 때 이전 파일과 달라진 점 — '신규'(새 시설: 기본값을 확인. 비슷한 이름이나 같은 이름의 기존 시설이 있으면 그 번호를 함께 적음) / "
                 "'확인: …'(사업유형이 달라져 기본값이 지금 값과 다름. 값은 바꾸지 않았음) / '후보 이름 바뀜'·'시설키 바뀜'(이름 정리 규칙이 바뀜. 검수_시설명은 그대로) / "
                 "'신규 행(같은 시설의 다른 검색어)'(검수 값은 그 시설과 같게 둠) / '이번 후보에 없음'(이전 내용 그대로 둠). 빈칸이면 달라진 것이 없는 행입니다.",
                 f"'{PREV_MARK.strip()} …' 로 시작하는 표시는 지난 실행 때 붙었는데 그 뒤 노란 칸이 그대로인 행입니다(이미 확인했으면 그냥 두거나 {STATUS_COL} 칸의 글을 지우세요 — 지워도 결과는 같습니다).",
                 "같은 시설이 두 줄로 잡혔으면 시설ID 칸에 같은 번호를 적어 한 시설로 합칠 수 있습니다(재검색이 한 시설로 묶고, 그 가운데 파일에서 가장 위에 있는 '포함' 줄의 "
                 "시설명·검색어·분류를 대표로 씁니다. 다시 실행해도 합친 줄끼리의 위아래 순서는 그대로 둡니다).",
                 "같은 폴더의 'facility_candidates.이전_날짜_시각.xlsx' 는 덮어쓰기 전 사본입니다. 검수는 사본이 아니라 facility_candidates.xlsx 에 하세요(사본을 고쳐도 반영되지 않습니다)."]:
        legend.append([line])
    tmp = path + ".tmp"                          # 임시 파일에 다 쓴 뒤 바꿔치기 → 쓰다가 멈춰도 검수 파일이 깨지지 않는다
    try:
        wb.save(tmp)
        os.replace(tmp, path)
    except OSError:
        try:
            os.remove(tmp)
        except OSError:
            pass
        raise
    log.info("검수용 워크북 저장: %s (%d개 시설)", path, len(agg))


def read_review_file(path: str) -> pd.DataFrame:
    """facility_candidates.xlsx 전체 행(필터 없음). 파일이 없으면 빈 DataFrame."""
    if not os.path.exists(path):
        return pd.DataFrame()
    wb = load_workbook(path, data_only=True, read_only=True)
    if "시설후보_검수" not in wb.sheetnames:
        names = ", ".join(wb.sheetnames)
        wb.close()
        raise SystemExit(f"[중단] '{path}' 에 '시설후보_검수' 시트가 없습니다(있는 시트: {names}). 시트 이름을 바꿨다면 '시설후보_검수' 로 되돌린 뒤 다시 실행하세요.")
    ws = wb["시설후보_검수"]
    rows = list(ws.values)
    wb.close()
    if not rows:
        return pd.DataFrame()
    return pd.DataFrame(rows[1:], columns=[str(c) for c in rows[0]])


def read_reviewed(path: str) -> pd.DataFrame:
    """검수 완료 행('포함')만. 검수_시설명이 비어 있으면 시설명_후보로 대체, 수요기관 컬럼이 없으면 빈 값."""
    df = read_review_file(path)
    if df.empty:
        raise SystemExit(f"검수 파일이 비어 있거나 없습니다: {path} (discover 를 먼저 실행)")
    for c in REVIEW_COLS + ["시설명_후보", "수요기관", "시설ID", "표2_대분류", "표2_중분류", "검색어"]:
        if c not in df.columns:
            df[c] = ""
    df = df[df["검수_포함여부"].map(lambda v: str(v).strip() == "포함")].copy()
    df["검수_시설명"] = [str(n).strip() if not _blank(n) else str(f).strip() for n, f in zip(df["검수_시설명"], df["시설명_후보"])]
    df["검수_별칭(;구분)"] = df["검수_별칭(;구분)"].map(lambda v: "" if _blank(v) else str(v))
    df["검수_수요기관"] = df["검수_수요기관"].map(lambda v: "" if _blank(v) else str(v).strip())
    df["시설ID"] = df["시설ID"].map(lambda v: str(v).strip())
    return df


def _gap_pattern(name: str) -> Optional[str]:
    """어절 사이 간격을 허용하는 예비 검색 정규식. '송파 창의 공공주택' → '송파.{0,15}창의.{0,15}공공주택'
    (시설명은 괄호·상태어를 지운 이름이라 '송파 창의혁신 공공주택 건설사업' 같은 원 공고명의 부분 문자열이 아닐 수 있다)."""
    toks = [t.replace(" ", "") for t in re.split(r"\s+", str(name or "")) if len(re.sub(r"[^가-힣A-Za-z0-9]", "", t)) >= 2]
    if len(toks) >= 2:
        return ".{0,15}".join(re.escape(t) for t in toks)
    if len(toks) == 1 and len(toks[0]) >= 4:          # '구 종합운동장공원' → 한 글자 어절을 뺀 '종합운동장공원' 만으로
        return re.escape(toks[0])
    return None


def research_by_facility(std: pd.DataFrame, reviewed: pd.DataFrame, max_hits_warn: int = 300) -> pd.DataFrame:
    """검수된 시설명(+별칭)으로 전량 데이터를 재검색 → 모든 공종 공고 수집, 시설ID 부여.
    - 검수_수요기관이 있으면 수요기관이 일치(공백 무시, 부분 포함)하는 공고만 채택.
    - 검수 행 여러 개가 같은 시설ID를 가지면 한 시설로 병합(시설명·별칭·수요기관을 합집합으로 검색) — 지자체와 도시공사로 나뉜 같은 시설을 합칠 때.
    - 정확한 이름으로 결과가 없으면 어절 사이 간격을 허용한 예비 검색."""
    out = []
    std_nospace = std["공고명"].astype(str).str.replace(" ", "", regex=False)
    inst_nospace = std["수요기관"].astype(str).str.replace(" ", "", regex=False)
    inst_values = [v for v in inst_nospace.unique().tolist() if v]      # 수요기관 판정은 고유값(수천 개)에만 하고 isin 으로 확장
    n_fac, n_loose = 0, 0
    for fid, grp in reviewed.groupby(reviewed["시설ID"].astype(str).str.strip(), sort=False):
        fac = grp.iloc[0]
        names, insts = [], []
        for _, r in grp.iterrows():
            nm = [r["검수_시설명"]] + [a.strip() for a in str(r.get("검수_별칭(;구분)") or "").split(";") if a.strip()]
            inst_src = str(r.get("검수_수요기관") or r.get("수요기관") or "")
            # '구례군 국민임대아파트' 처럼 지자체명을 접두한 시설명은 공고명에 지자체명이 없으므로 접두어를 뗀 패턴도 함께 검색(수요기관 필터가 범위를 좁힌다)
            names += [n for n in nm if n and not _blank(n)]
            names += [strip_institution_prefix(n, inst_src) for n in nm if n and not _blank(n)]
            inst = str(r.get("검수_수요기관") or "").replace(" ", "")
            if inst:
                insts.append(inst)
        if len(grp) > 1:
            log.info("시설 %s: 검수 행 %d개가 같은 시설ID → 한 시설로 병합(%s)", fid, len(grp), " / ".join(dict.fromkeys(str(x) for x in grp["검수_시설명"])))
        pats = list(dict.fromkeys(n.replace(" ", "") for n in names if n and not _blank(n)))
        if not pats:
            log.warning("시설 %s: 검수_시설명이 비어 있어 건너뜀", fid)
            continue
        if min(len(p) for p in pats) <= 2 and not insts:
            log.warning("시설 %s '%s': 검색어가 너무 짧아(2자 이하) 엉뚱한 공고가 대량 포함될 수 있음", fid, "/".join(pats))
        inst_mask = None
        if insts:
            inst_mask = inst_nospace.isin({v for v in inst_values if any(i in v or v in i for i in insts)})
        mask = pd.Series(False, index=std.index)
        for p in pats:
            mask |= std_nospace.str.contains(re.escape(p), na=False)
        if inst_mask is not None:
            mask &= inst_mask
        if not mask.any():
            loose = [g for g in (_gap_pattern(n) for n in names) if g]
            if loose:
                for g in dict.fromkeys(loose):
                    mask |= std_nospace.str.contains(g, na=False, regex=True)
                if inst_mask is not None:
                    mask &= inst_mask
                if mask.any():
                    n_loose += 1
                    log.info("시설 %s '%s': 정확한 이름으로는 없음 → 어절 간격을 허용한 예비 검색으로 %d건", fid, fac["검수_시설명"], int(mask.sum()))
        hit = std[mask].copy()
        if hit.empty:
            log.info("시설 %s '%s': 재검색 결과 없음", fid, fac["검수_시설명"])
            continue
        if len(hit) > max_hits_warn:
            log.warning("시설 %s '%s': %d건 매칭 — 검수_시설명이 너무 일반적인지 확인", fid, fac["검수_시설명"], len(hit))
        hit["시설ID"] = fid
        hit["시설명"] = fac["검수_시설명"]
        hit["사업유형"] = [work_type_for(n, str(fac.get("검색어") or ""), fac["검수_시설명"]) for n in hit["공고명"]]
        hit["표2_대분류"] = fac["표2_대분류"]
        hit["표2_중분류"] = fac["표2_중분류"]
        hit["검색어"] = fac["검색어"]
        hit["수요기관_대표"] = fac.get("검수_수요기관") or fac.get("수요기관", "")
        out.append(hit)
        n_fac += 1
    df = pd.concat(out, ignore_index=True) if out else pd.DataFrame()
    if not df.empty:
        matched = set(zip(df["시설ID"].astype(str), df["공고키"].astype(str)))     # 이름으로 찾은 행(아래에서 유지보수로 빠지는 행 포함)
        # 유지보수 성격 공고는 시설 재검색에서도 제외 (신축·증축·리모델링 프로젝트 공사비만)
        df = df[~df["사업유형"].isin(["유지보수"])].copy()
        df = df.drop_duplicates(["시설ID", "공고키"])
        n_found = len(df)
        df, n_sib, newer = add_other_orders(df, std, matched)
        # 보충한 행은 뒤에 붙는다. dedup 이 사업 묶음에는 검색으로 찾은 차수의 공고명을 쓰도록 표시해 둔다(보충은 금액·일정·취소 여부만 바꾼다)
        df["차수보충"] = [""] * n_found + ["Y"] * (len(df) - n_found)
        if n_sib:
            # 건수는 (시설, 공고) 기준이다(한 공고가 세 시설에 걸리면 3건). 목록은 뒤에서 다시 볼 수 없으므로 줄이지 않고 적는다(아주 많을 때만 앞 60건)
            log.info("같은 공고의 다른 차수 %d건 보충(시설 이름 검색에 잡히지 않은 차수 — 변경공고에서 공고명이 바뀐 경우 등). "
                     "그중 더 늦은 차수가 있어 최신 차수가 바뀌는 공고 %d건(시설별로 셈)%s",
                     n_sib, len(newer), (": " + ", ".join(newer[:60]) + (f" … 외 {len(newer) - 60}건" if len(newer) > 60 else "")) if newer else "")
    log.info("시설 재검색: %d개 시설 → %d건 공고 (예비 검색으로 찾은 시설 %d개)",
             reviewed["시설ID"].nunique() if not reviewed.empty else 0, len(df), n_loose)
    return df


FACILITY_COLS = ["시설ID", "시설명", "사업유형", "표2_대분류", "표2_중분류", "검색어", "수요기관_대표"]
_BLANK_NO = {"", "nan", "NaN", "None", "<NA>", "NaT"}


def add_other_orders(df: pd.DataFrame, std: pd.DataFrame, matched: Optional[set] = None):
    """같은 공고번호의 다른 차수 가운데 시설 이름으로 찾지 못한 행을 보충한다. 반환 (df, 보충한 행 수, 최신 차수가 바뀐 공고 설명 목록).
    변경공고에서 공고명이 바뀌면(오타 수정·사업명 변경) 그 차수만 이름 검색에서 빠진다. 뒤 차수가 빠지면 옛 차수가 최신으로 남아 고치기 전 금액·일정이 실리고,
    취소공고 차수가 빠지면 취소된 공고가 합산된다. 앞 차수가 빠지면 원 공고문 첨부(변경공고에는 정정 내용만 있는 경우)를 쓰지 못한다.
    - 보충한 행의 시설 정보와 사업유형은 같은 (시설, 공고번호)에서 이름으로 찾아 **남은**(유지보수로 빠지지 않은) 가장 늦은 차수의 것을 그대로 쓴다 — 어느 시설·사업유형에 넣을지는 바꾸지 않는다.
      (이름으로 찾은 가장 늦은 차수가 유지보수라 빠졌으면 그보다 앞의 남은 차수를 따른다. 보충 행 자기 공고명이 유지보수 꼴이어도 물려받은 사업유형으로 들어간다.)
    - matched = 이름으로 찾았던 (시설ID, 공고키). 유지보수라서 뺀 차수를 다시 넣지 않기 위한 것이다."""
    if df.empty or std.empty or "공고번호" not in std.columns or "공고키" not in std.columns:
        return df, 0, []
    text = lambda s: s.astype(str).fillna("")                # 빈 값은 판다스 판에 따라 'nan' 문자열 또는 결측으로 남는다 → 둘 다 빈 것으로
    nos = set(text(df["공고번호"])) - _BLANK_NO                # 공고번호가 빈 행끼리는 같은 공고가 아니다
    cand = std[text(std["공고번호"]).isin(nos)]
    if cand.empty:
        return df, 0, []
    order = lambda s: pd.to_numeric(s, errors="coerce").fillna(0)
    fac_cols = [c for c in FACILITY_COLS if c in df.columns]
    base = df.assign(_no=text(df["공고번호"]), _base_ord=order(df["공고차수"])).sort_values("_base_ord", kind="stable")
    base = base[base["_no"].isin(nos)].drop_duplicates(["시설ID", "_no"], keep="last")[["_no", "_base_ord"] + fac_cols]
    cand = cand.drop(columns=[c for c in fac_cols if c in cand.columns]).drop_duplicates("공고키")
    add = cand.assign(_no=text(cand["공고번호"])).merge(base, on="_no", how="inner")
    have = set(zip(df["시설ID"].astype(str), df["공고키"].astype(str))) | set(matched or ())
    add = add[[(f, k) not in have for f, k in zip(add["시설ID"].astype(str), add["공고키"].astype(str))]]
    if add.empty:
        return df, 0, []
    add = add.assign(_ord=order(add["공고차수"]))
    top = add[add["_ord"] > add["_base_ord"]].sort_values("_ord").drop_duplicates(["시설ID", "_no"], keep="last")
    num3 = lambda x: f"{int(x):03d}" if abs(x) != float("inf") else str(x)      # 차수가 'inf' 로 읽히는 값이어도 로그 문구 때문에 멈추지 않게
    newer = [f"{n}({f}: {num3(a)}→{num3(b)})" for f, n, a, b in zip(top["시설ID"], top["_no"], top["_base_ord"], top["_ord"])]
    for c in df.columns:                                     # 전량 표에 없는 열(이미 붙어 있는 '차수보충' 등)은 빈 값으로, 열 자료형은 그대로
        if c not in add.columns:
            try:
                add[c] = pd.Series([None] * len(add), index=add.index, dtype=df[c].dtype)
            except (TypeError, ValueError):
                add[c] = None
    out = pd.concat([df, add[list(df.columns)]], ignore_index=True).drop_duplicates(["시설ID", "공고키"])
    return out, len(out) - len(df), sorted(newer)
