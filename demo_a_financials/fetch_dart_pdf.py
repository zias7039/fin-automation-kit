"""DART 공시의 접수번호(rcpNo)로 보고서 PDF를 내려받는다.

사용법:
    python fetch_dart_pdf.py 20240312000736 -o samples/samsung_2023.pdf
"""
import argparse
import http.cookiejar
import re
import urllib.request

BASE = "https://dart.fss.or.kr"


def fetch(rcp_no, output):
    opener = urllib.request.build_opener(urllib.request.HTTPCookieProcessor(http.cookiejar.CookieJar()))
    opener.addheaders = [("User-Agent", "Mozilla/5.0")]  # 기본 urllib UA로는 빈 응답이 온다
    viewer = opener.open(f"{BASE}/dsaf001/main.do?rcpNo={rcp_no}").read().decode("utf-8", "replace")
    dcm_no = re.search(r"dcmNo'\] = \"(\d+)", viewer).group(1)
    page = f"{BASE}/pdf/download/main.do?rcp_no={rcp_no}&dcm_no={dcm_no}"
    opener.open(page).read()  # PDF 요청 전에 세션 쿠키가 필요하다
    req = urllib.request.Request(f"{BASE}/pdf/download/pdf.do?rcp_no={rcp_no}&dcm_no={dcm_no}", headers={"Referer": page})
    data = opener.open(req).read()
    if not data.startswith(b"%PDF"):
        raise RuntimeError("PDF를 받지 못했습니다. 접수번호를 확인하세요.")
    with open(output, "wb") as f:
        f.write(data)
    print(f"{output} ({len(data):,} bytes)")


if __name__ == "__main__":
    ap = argparse.ArgumentParser()
    ap.add_argument("rcp_no")
    ap.add_argument("-o", "--output", required=True)
    args = ap.parse_args()
    fetch(args.rcp_no, args.output)
