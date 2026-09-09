"""가짜 나라장터 API + 첨부파일 서버 (통합 테스트용, 네트워크·키 불필요).

- 공공데이터포털 응답 형식을 모사: JSON 정상 응답, 데이터 없는 달은 resultCode 03, 잘못된 키는 XML 30(HTTP 200)
- 오퍼레이션: getBidPblancListInfoCnstwk / getBidPblancListInfoCnstwkBsisAmount / getBidPblancListInfoLicenseLimit
  (면허제한 업종명 필드는 config 기본값과 다른 'lcnsLmtNm' 으로 내려 후보 자동 탐색을 검증)
- /files/<이름>: HWPX·DOCX·TXT(cp949)·HTML(로그인 페이지) 첨부
※ 모든 공고·기관·금액은 가상의 예시.
"""
from __future__ import annotations

import io
import json
import threading
import zipfile
from collections import Counter
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from typing import Dict, List
from urllib.parse import parse_qs, quote, unquote, urlparse

KEY = "testkey"
_HP = "http://www.hancom.co.kr/hwpml/2011/paragraph"
_HS = "http://www.hancom.co.kr/hwpml/2011/section"


def _hwpx(paragraphs: List[str], table: List[List[str]] | None = None) -> bytes:
    ps = "".join(f"<hp:p><hp:run><hp:t>{p}</hp:t></hp:run></hp:p>" for p in paragraphs)
    tbl = ""
    if table:
        rows = "".join("<hp:tr>" + "".join(f"<hp:tc><hp:subList><hp:p><hp:run><hp:t>{c}</hp:t></hp:run></hp:p></hp:subList></hp:tc>" for c in r) + "</hp:tr>"
                       for r in table)
        tbl = f"<hp:p><hp:run><hp:tbl>{rows}</hp:tbl></hp:run></hp:p>"
    xml = f'<?xml version="1.0" encoding="UTF-8"?><hs:sec xmlns:hs="{_HS}" xmlns:hp="{_HP}">{ps}{tbl}</hs:sec>'
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as z:
        z.writestr("mimetype", "application/hwp+zip")
        z.writestr("Contents/section0.xml", xml)
    return buf.getvalue()


def _docx(paragraphs: List[str]) -> bytes:
    import docx  # python-docx
    d = docx.Document()
    for p in paragraphs:
        d.add_paragraph(p)
    t = d.add_table(rows=2, cols=2)
    t.cell(0, 0).text, t.cell(0, 1).text = "구분", "금액(원)"
    t.cell(1, 0).text, t.cell(1, 1).text = "관급자재(관급자관급액)", "2,100,000,000"
    buf = io.BytesIO()
    d.save(buf)
    return buf.getvalue()


def build_dataset(base: str) -> Dict:
    """가상 공고 데이터셋. base: 파일 서버 URL 접두어."""
    f = lambda name: f"{base}/files/{quote(name)}"
    common = dict(rgstTyNm="조달청 또는 나라장터 자체 공고건", bidClseDt="", opengDt="", refNo="", bfSpecRgstNo="",
                  bidNtceDtlUrl="https://example.invalid/detail", dminsttCd="", bdgtAmt="", VAT="", govsplyAmt="0",
                  contrctrcnstrtnGovsplyMtrlAmt="0", govcnstrtnGovsplyMtrlAmt="0", befBidBbancNo="", chgNtceRsn="",
                  subsiCnsttyNm1="", sptDscrptDocUrl1="", stdNtceDocUrl="", sucsfbidLwltRate="")
    N = [
        # 가상군 문화예술회관 — 건축 본공사(첨부: 공고문.hwpx + 현장설명서.docx)
        dict(common, bidNtceNo="R24010001", bidNtceOrd="000", bidNtceNm="가상군 문화예술회관 건립공사", ntceKindNm="등록공고", reNtceYn="N",
             bidNtceDt="2024-01-10 10:00:00", ntceInsttNm="가상군", dminsttNm="가상군", presmptPrce="30000000000", mainCnsttyNm="건축공사",
             cnstrtsiteRgnNm="전라남도 가상군", ntceSpecDocUrl1=f("입찰공고문.hwpx"), ntceSpecFileNm1="입찰공고문.hwpx",
             sptDscrptDocUrl1=f("현장설명서.docx"), dminsttCd="4790000", bdgtAmt="35100000000", VAT="3000000000",
             govsplyAmt="2100000000", govcnstrtnGovsplyMtrlAmt="2100000000", subsiCnsttyNm1="토목공사업", subsiCnsttyNm2="조경공사업"),
        # 같은 공고의 변경 차수(001)
        dict(common, bidNtceNo="R24010001", bidNtceOrd="001", bidNtceNm="가상군 문화예술회관 건립공사 [변경]", ntceKindNm="변경공고", reNtceYn="N",
             bidNtceDt="2024-01-17 10:00:00", ntceInsttNm="가상군", dminsttNm="가상군", presmptPrce="30000000000", mainCnsttyNm="건축공사",
             cnstrtsiteRgnNm="전라남도 가상군", ntceSpecDocUrl1=f("입찰공고문.hwpx"), ntceSpecFileNm1="입찰공고문.hwpx",
             sptDscrptDocUrl1=f("현장설명서.docx"), dminsttCd="4790000", bdgtAmt="35100000000", VAT="3000000000",
             govsplyAmt="2100000000", govcnstrtnGovsplyMtrlAmt="2100000000", subsiCnsttyNm1="토목공사업", chgNtceRsn="공고 기간 정정"),
        # 전기 — 차수가 정수 0, 추정가격이 숫자형(타입 혼재), 첨부는 cp949 txt
        dict(common, bidNtceNo="R24010002", bidNtceOrd=0, bidNtceNm="가상군 문화예술회관 건립 전기공사", ntceKindNm="일반공고", reNtceYn="N",
             bidNtceDt="2024-01-12 10:00:00", ntceInsttNm="가상군", dminsttNm="가상군", presmptPrce=2600000000, mainCnsttyNm="전기공사",
             cnstrtsiteRgnNm="전라남도 가상군", ntceSpecDocUrl1=f("공고문_전기.txt"), ntceSpecFileNm1=None),
        # 소방 — 원공고(취소됨) + '(취소 후 재공고)'
        dict(common, bidNtceNo="R24010003", bidNtceOrd="00", bidNtceNm="가상군 문화예술회관 건립 소방시설공사", ntceKindNm="취소공고", reNtceYn="N",
             bidNtceDt="2024-01-15 10:00:00", ntceInsttNm="가상군", dminsttNm="가상군", presmptPrce="", mainCnsttyNm="소방공사",
             cnstrtsiteRgnNm="전라남도 가상군"),
        dict(common, bidNtceNo="R24020004", bidNtceOrd="00", bidNtceNm="가상군 문화예술회관 건립 소방시설공사(취소 후 재공고)", ntceKindNm="재공고", reNtceYn="Y", befBidBbancNo="R24010003",
             bidNtceDt="2024-02-20 10:00:00", ntceInsttNm="가상군", dminsttNm="가상군", presmptPrce="530000000", mainCnsttyNm="소방공사",
             cnstrtsiteRgnNm="전라남도 가상군"),
        # 부대공사(소액) — 대표를 대체하면 안 됨
        dict(common, bidNtceNo="R24030005", bidNtceOrd="00", bidNtceNm="가상군 문화예술회관 관리사무소 건립공사", ntceKindNm="일반공고", reNtceYn="N",
             bidNtceDt="2024-03-05 10:00:00", ntceInsttNm="가상군", dminsttNm="가상군", presmptPrce="400000000", mainCnsttyNm="건축공사",
             cnstrtsiteRgnNm="전라남도 가상군"),
        # 야외공연장 — 두 분류(공연장/야외공연장)에 걸림, 첨부 1은 로그인 HTML(실패), 첨부 2는 hwpx
        dict(common, bidNtceNo="R24020006", bidNtceOrd="00", bidNtceNm="가상시 야외공연장 조성공사", ntceKindNm="일반공고", reNtceYn="N",
             bidNtceDt="2024-02-01 10:00:00", ntceInsttNm="가상시", dminsttNm="가상시", presmptPrce="5000000000", mainCnsttyNm="건축공사",
             cnstrtsiteRgnNm="경기도 가상시", ntceSpecDocUrl1=f("login.html"), ntceSpecFileNm1="공고문.hwp",
             ntceSpecDocUrl2=f("야외공연장_공고문.hwpx"), ntceSpecFileNm2="야외공연장_공고문.hwpx"),
        # 다른 군의 같은 이름 시설 — 수요기관으로 분리되어야 함
        dict(common, bidNtceNo="R24020007", bidNtceOrd="00", bidNtceNm="다른군 문화예술회관 건립공사", ntceKindNm="일반공고", reNtceYn="N",
             bidNtceDt="2024-02-15 10:00:00", ntceInsttNm="다른군", dminsttNm="다른군", presmptPrce="25000000000", mainCnsttyNm="건축공사",
             cnstrtsiteRgnNm="충청남도 다른군"),
        # 잡음: 유지보수 / 제외어 / 무관
        dict(common, bidNtceNo="R24010008", bidNtceOrd="00", bidNtceNm="가상군 문화예술회관 옥상 방수공사", ntceKindNm="일반공고", reNtceYn="N",
             bidNtceDt="2024-01-20 10:00:00", ntceInsttNm="가상군", dminsttNm="가상군", presmptPrce="80000000", mainCnsttyNm="건축공사",
             cnstrtsiteRgnNm="전라남도 가상군"),
        dict(common, bidNtceNo="R24030009", bidNtceOrd="00", bidNtceNm="샘플시 공영주차장 정비공사", ntceKindNm="일반공고", reNtceYn="N",
             bidNtceDt="2024-03-02 10:00:00", ntceInsttNm="샘플시", dminsttNm="샘플시", presmptPrce="300000000", mainCnsttyNm="토목공사",
             cnstrtsiteRgnNm="경기도 샘플시"),
        dict(common, bidNtceNo="R24030010", bidNtceOrd="00", bidNtceNm="샘플시 농어촌도로 포장공사", ntceKindNm="일반공고", reNtceYn="N",
             bidNtceDt="2024-03-03 10:00:00", ntceInsttNm="샘플시", dminsttNm="샘플시", presmptPrce="900000000", mainCnsttyNm="토목공사",
             cnstrtsiteRgnNm="경기도 샘플시"),
    ]
    bsis = []
    for n in N:
        p = str(n["presmptPrce"]).strip()
        if p and n["bidNtceNo"] != "R24020007":          # 다른군 건은 기초금액 없음(문서 폴백 검증용)
            bsis.append(dict(bidNtceNo=n["bidNtceNo"], bidNtceOrd=str(n["bidNtceOrd"]).zfill(3), bssamt=str(int(int(p) * 1.1)),
                             bssAmtPurcnstcst=str(int(int(p) * 0.9)), bidNtceDt=n["bidNtceDt"]))
    lic = {
        "R24010001": [{"lcnsLmtNm": "건축공사업"}, {"lcnsLmtNm": "토목건축공사업"}],
        "R24010002": [{"lcnsLmtNm": "전기공사업"}],
        "R24020004": [{"lcnsLmtNm": "소방시설공사업(일반)"}],
        "R24030005": [{"lcnsLmtNm": "건축공사업"}],
        "R24020006": [{"lcnsLmtNm": "건축공사업"}],
        "R24020007": [{"lcnsLmtNm": "건축공사업"}],
    }
    # 낙찰정보서비스(가짜) — 공사 낙찰 목록. R24010001 은 차수 001 로 낙찰, R24010002 는 1월 말 개찰 후 2월 재개찰로 2행(최신 개찰일시 행이 남아야 함)
    awards = [
        dict(bidNtceNo="R24010001", bidNtceOrd="001", bidNtceNm="가상군 문화예술회관 건립공사 [변경]", opengDt="2024-02-05 11:00:00",
             sucsfbidAmt="28500000000", sucsfbidRate="86.36", bidwinnrNm="가상건설(주)", bidwinnrBizno="1234567890", prtcptCnum="12",
             presmptPrce="30000000000", bssamt="33000000000", dminsttNm="가상군"),
        dict(bidNtceNo="R24010002", bidNtceOrd="000", bidNtceNm="가상군 문화예술회관 건립 전기공사", opengDt="2024-01-30 11:00:00",
             sucsfbidAmt="2500000000", sucsfbidRate="87.41", bidwinnrNm="옛날전기(주)", bidwinnrBizno="2234567890", prtcptCnum="5",
             presmptPrce="2600000000", bssamt="2860000000", dminsttNm="가상군"),
        dict(bidNtceNo="R24010002", bidNtceOrd="000", bidNtceNm="가상군 문화예술회관 건립 전기공사", opengDt="2024-02-20 11:00:00",
             sucsfbidAmt="2574000000", sucsfbidRate="90.00", bidwinnrNm="가상전기(주)", bidwinnrBizno="3234567890", prtcptCnum="7",
             presmptPrce="2600000000", bssamt="2860000000", dminsttNm="가상군"),
        dict(bidNtceNo="R24020006", bidNtceOrd="000", bidNtceNm="가상시 야외공연장 조성공사", opengDt="2024-03-11 11:00:00",
             sucsfbidAmt="4800000000", sucsfbidRate="87.27", bidwinnrNm="야외건설(주)", bidwinnrBizno="4234567890", prtcptCnum="9",
             presmptPrce="5000000000", bssamt="5500000000", dminsttNm="가상시"),
    ]
    files = {
        "입찰공고문.hwpx": (_hwpx(["입 찰 공 고 (가상군 문화예술회관 건립공사)", "1. 공사개요",
                                "위치: 전라남도 가상군 가상읍", "규모: 연면적 15,200㎡, 지하1층/지상3층, 철골철근콘크리트조",
                                "공사기간: 착공일로부터 1,080일", "2. 공사금액"],
                               [["구분", "금액"], ["추정가격", "30,000,000,000원"], ["기초금액(부가세 포함)", "33,000,000,000원"],
                                ["관급자관급액", "2,100,000,000원"], ["총공사금액", "35,100,000,000원"]]), "application/octet-stream"),
        "현장설명서.docx": (_docx(["현장설명서 — 가상군 문화예술회관 건립공사", "연면적 15,200㎡, 건축면적 6,300㎡", "용도: 공연장(문화 및 집회시설)"]),
                        "application/octet-stream"),
        "공고문_전기.txt": ("입찰공고문(전기)\n가상군 문화예술회관 건립 전기공사\n추정가격 2,600,000,000원\n기초금액 2,860,000,000원\n공사기간 900일\n".encode("cp949"),
                        "text/plain"),
        "login.html": ("<!DOCTYPE html><html><head><title>login</title></head><body>로그인이 필요합니다</body></html>".encode("utf-8"), "text/html; charset=utf-8"),
        "야외공연장_공고문.hwpx": (_hwpx(["가상시 야외공연장 조성공사 입찰공고", "규모: 객석 1,200석, 연면적 2,400㎡, 지상2층",
                                       "추정가격 5,000,000,000원 / 기초금액 5,500,000,000원", "공사기간 540일"]), "application/octet-stream"),
    }
    return {"notices": N, "bsis": bsis, "license": lic, "files": files, "awards": awards}


def _xml_error(code: str, msg: str) -> bytes:
    return (f"<OpenAPI_ServiceResponse><cmmMsgHeader><errMsg>SERVICE ERROR</errMsg><returnAuthMsg>{msg}</returnAuthMsg>"
            f"<returnReasonCode>{code}</returnReasonCode></cmmMsgHeader></OpenAPI_ServiceResponse>").encode()


class MockG2B:
    def __init__(self):
        self.calls: Counter = Counter()
        self.httpd = None
        self.base = ""
        self.data: Dict = {}

    def start(self) -> str:
        mock = self

        class H(BaseHTTPRequestHandler):
            def log_message(self, *a):  # 콘솔 소음 제거
                pass

            def _send(self, status, body: bytes, ctype="application/json; charset=utf-8", extra=None):
                self.send_response(status)
                self.send_header("Content-Type", ctype)
                self.send_header("Content-Length", str(len(body)))
                for k, v in (extra or {}).items():
                    self.send_header(k, v)
                self.end_headers()
                self.wfile.write(body)

            def do_GET(self):
                u = urlparse(self.path)
                q = {k: v[0] for k, v in parse_qs(u.query).items()}
                parts = u.path.strip("/").split("/")
                if parts[0] == "files":
                    name = unquote(parts[1])
                    mock.calls["file:" + name] += 1
                    if name not in mock.data["files"]:
                        return self._send(404, b"not found", "text/plain")
                    body, ctype = mock.data["files"][name]
                    return self._send(200, body, ctype, {"Content-Disposition": f"attachment; filename*=UTF-8''{quote(name)}"})
                op = parts[-1]
                mock.calls[op] += 1
                if q.get("serviceKey") != KEY:   # 실제 게이트웨이는 type=json 요청에 HTTP 403 + JSON 형식 오류를 돌려준다
                    body = {"cmmMsgHeader": {"errMsg": "SERVICE ERROR", "returnAuthMsg": "SERVICE_KEY_IS_NOT_REGISTERED_ERROR",
                                             "returnReasonCode": "30"}}
                    return self._send(403, json.dumps(body).encode("utf-8"))
                rows = 999 if not q.get("numOfRows") else int(q["numOfRows"])
                page = int(q.get("pageNo", "1"))
                if op == "getBidPblancListInfoCnstwk":
                    items = [n for n in mock.data["notices"] if n["bidNtceDt"][:7].replace("-", "") == q.get("inqryBgnDt", "")[:6]]
                elif op == "getBidPblancListInfoCnstwkBsisAmount":
                    items = [b for b in mock.data["bsis"] if b["bidNtceDt"][:7].replace("-", "") == q.get("inqryBgnDt", "")[:6]]
                elif op == "getScsbidListSttusCnstwk":
                    if parts[0] != "ScsbidInfoService":        # 낙찰정보서비스는 별도 End Point — 입찰공고 경로로 부르면 없는 오퍼레이션
                        return self._send(200, _xml_error("12", "NO_OPENAPI_SERVICE_ERROR"), "application/xml")
                    if q.get("inqryDiv") != "1":
                        return self._send(200, _xml_error("10", "INVALID_REQUEST_PARAMETER_ERROR"), "application/xml")
                    items = [w for w in mock.data["awards"] if w["opengDt"][:7].replace("-", "") == q.get("inqryBgnDt", "")[:6]]
                elif op == "getBidPblancListInfoLicenseLimit":
                    if q.get("inqryDiv") != "2" or not q.get("bidNtceNo"):
                        return self._send(200, _xml_error("10", "INVALID_REQUEST_PARAMETER_ERROR"), "application/xml")
                    items = [dict(it, bidNtceNo=q["bidNtceNo"], bidNtceOrd="00") for it in mock.data["license"].get(q["bidNtceNo"], [])]
                else:
                    return self._send(200, _xml_error("12", "NO_OPENAPI_SERVICE_ERROR"), "application/xml")
                if not items:
                    body = {"response": {"header": {"resultCode": "03", "resultMsg": "NODATA_ERROR"}, "body": {}}}
                else:
                    chunk = items[(page - 1) * rows: page * rows]
                    body = {"response": {"header": {"resultCode": "00", "resultMsg": "정상"},
                                         "body": {"items": chunk, "numOfRows": rows, "pageNo": page, "totalCount": len(items)}}}
                return self._send(200, json.dumps(body, ensure_ascii=False).encode("utf-8"))

        self.httpd = ThreadingHTTPServer(("127.0.0.1", 0), H)
        port = self.httpd.server_address[1]
        self.base = f"http://127.0.0.1:{port}"
        self.data = build_dataset(self.base)
        threading.Thread(target=self.httpd.serve_forever, daemon=True).start()
        return self.base

    def stop(self):
        if self.httpd:
            self.httpd.shutdown()
            self.httpd.server_close()


if __name__ == "__main__":
    m = MockG2B()
    print("mock server:", m.start(), "(Ctrl+C 로 종료)")
    try:
        threading.Event().wait()
    except KeyboardInterrupt:
        m.stop()
