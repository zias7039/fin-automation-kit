"""DART 사업보고서 PDF의 '요약 연결재무정보' 표를 읽어 연도별 엑셀로 정리한다.

- 여러 해의 보고서를 넣으면 연도를 합치고, 겹치는 연도의 숫자가 서로 같은지 검증한다.
- 단위(원/천원/백만원/억원)는 백만원으로 통일한다. 주당이익·회사 수는 그대로 둔다.
- 자동으로 읽지 못한 줄은 '검증' 시트의 '확인 필요'에 남긴다.

사용법:
    python extract_summary.py samples/samsung_2023.pdf samples/samsung_2021.pdf -o out.xlsx
"""
import argparse
import re
import sys

import pdfplumber
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill

TOKEN = r"(?:\(?[△-]?[\d,]+(?:\.\d+)?\)?(?:원/주|개|원)?|-)"
LINE_RE = re.compile(rf"^(?P<label>.*?[^\d,\s(△-])\s+(?P<nums>{TOKEN}(?:\s+{TOKEN})*)$")
YEAR_RE = re.compile(r"(20\d{2})(?:\.\s*12\.\s*31|년\s*12월\s*(?:말|31일))")
UNIT_RE = re.compile(r"단위\s*:\s*(백만원|천원|억원|원)")
UNIT_TO_MILLION = {"원": 1e-6, "천원": 1e-3, "백만원": 1, "억원": 100}
SKIP_WORDS = ("dart.fss.or.kr", "~")


def to_number(token):
    if token == "-":
        return 0.0
    t = re.sub(r"(원/주|개|원)$", "", token)
    negative = t.startswith(("△", "-", "("))
    return -float(t.strip("()△-").replace(",", "")) if negative else float(t.replace(",", ""))


def norm(label):
    return re.sub(r"[\s\[\]ㆍ·]|\(손실\)|\(원\)|의", "", label)


def find_section(pdf):
    """표가 시작되는 페이지 번호와, 다음 페이지까지 이어 붙인 표 본문을 돌려준다."""
    for i, page in enumerate(pdf.pages):
        text = page.extract_text() or ""
        m = re.search(r"요약\s*연결\s*재무\s*정보", text)
        if not m:
            continue
        body = text[m.start():]
        if i + 1 < len(pdf.pages):
            body += "\n" + (pdf.pages[i + 1].extract_text() or "")
        end = re.search(r"^※|^나\.\s", body[1:], re.M)
        body = body[: end.start() + 1 if end else None]
        if "자산총계" in body.replace(" ", "") and YEAR_RE.search(body):  # 목차 등 표가 아닌 언급은 건너뛴다
            return i + 1, body
    raise ValueError("요약 연결재무정보 표를 찾지 못했습니다.")


def parse_section(body):
    """표 본문을 (연도 목록, 행 목록, 확인 필요 줄)로 바꾼다. 행은 (키, 원래 항목명, {연도: 값})."""
    unit = UNIT_RE.search(body)
    scale = UNIT_TO_MILLION[unit.group(1)] if unit else 1
    years = next(
        [int(y) for y in found] for found in (YEAR_RE.findall(line) for line in body.splitlines()) if len(found) >= 2
    )
    rows, unread, seen = [], [], {}
    for raw in body.splitlines():
        line = raw.strip()
        if any(w in line for w in SKIP_WORDS) or line.startswith("(단위") or YEAR_RE.search(line):
            continue
        m = LINE_RE.match(line)
        if not m:
            continue
        label, nums = re.sub(r"ㆍ\s+", "ㆍ", m.group("label").strip()), m.group("nums").split()
        if len(nums) != len(years):
            unread.append(line)
            continue
        key = norm(label)
        seen[key] = seen.get(key, 0) + 1
        if seen[key] > 1:  # '기타', '비지배지분'처럼 여러 번 나오는 항목은 순번으로 구분
            key = f"{key}#{seen[key]}"
        keep_scale = "주당" in label or label.endswith("수") or any(n.endswith(("개", "원/주")) for n in nums)
        values = {y: round(to_number(n) * (1 if keep_scale else scale), 6) for y, n in zip(years, nums)}
        rows.append((key, label, values))
    return years, rows, unread


# 회사마다 항목명이 달라서 비율 계산에 쓰는 표준 항목은 이름 패턴으로 찾는다.
STANDARD = {
    "유동자산": r"^유동자산$",
    "유동부채": r"^유동부채$",
    "자산총계": r"^자산총계$",
    "부채총계": r"^부채총계$",
    "자본총계": r"^자본총계$",
    "지배주주지분": r"^지배기업.*소유",
    "매출액": r"^(매출액|영업수익|매출)$",
    "영업이익": r"^영업이익",
    "당기순이익": r"^(연결총)?(당기)?순이익",
}


def standard_items(keys):
    found = {}
    for name, pattern in STANDARD.items():
        found[name] = next((k for k in keys if re.search(pattern, k)), None)
    # 지배주주 순이익: 당기순이익 줄 바로 다음에 나오는 '지배기업 ... 소유주지분'
    if found["당기순이익"]:
        after = keys[keys.index(found["당기순이익"]) + 1:]
        found["지배주주순이익"] = next((k for k in after[:2] if re.search(r"^지배기업.*소유", k)), None)
    return found


def merge(reports):
    """최신 보고서 값을 우선으로 합치고, 겹치는 연도의 값 비교 결과를 기록한다."""
    merged, labels, order, checks = {}, {}, [], []
    for source, years, rows in sorted(reports, key=lambda r: -max(r[1])):
        for key, label, values in rows:
            if key not in labels:
                labels[key] = label
                order.append(key)
            for year, value in values.items():
                kept = merged.setdefault(year, {}).get(key)
                if kept is None:
                    merged[year][key] = value
                else:
                    checks.append((year, labels[key], kept, value, source, abs(kept - value) < 1e-6))
    return merged, labels, order, checks


def ratios(merged, std):
    years = sorted(merged)
    out = {}
    for prev, year in zip([None] + years, years):
        cur, p = merged[year], merged.get(prev, {})

        def get(d, name):
            return d.get(std.get(name)) if std.get(name) else None

        def div(a, b):
            return a / b if a is not None and b else None

        sales, prev_sales = get(cur, "매출액"), get(p, "매출액")
        eq, prev_eq = get(cur, "지배주주지분"), get(p, "지배주주지분")
        out[year] = {
            "매출액 증가율": sales / prev_sales - 1 if sales is not None and prev_sales else None,
            "영업이익률": div(get(cur, "영업이익"), sales),
            "순이익률": div(get(cur, "당기순이익"), sales),
            "ROE (지배주주, 평균자본)": div(get(cur, "지배주주순이익"), (eq + prev_eq) / 2 if eq and prev_eq else None),
            "부채비율": div(get(cur, "부채총계"), get(cur, "자본총계")),
            "유동비율": div(get(cur, "유동자산"), get(cur, "유동부채")),
        }
    return out


def sum_checks(merged, std):
    """총계가 맞는지 확인해 추출 오류를 잡는다."""
    out = []
    for year in sorted(merged):
        v = merged[year]
        a, l, e = (v.get(std.get(n)) for n in ("자산총계", "부채총계", "자본총계"))
        if None not in (a, l, e):
            out.append((year, "자산총계 = 부채총계 + 자본총계", a - (l + e)))
    return out


def write_excel(path, merged, labels, order, std, checks, sums, unread, sources):
    wb = Workbook()
    years = sorted(merged)

    ws = wb.active
    ws.title = "요약재무"
    ws.append(["항목 (단위: 백만원, 주당이익: 원)"] + years)
    for key in order:
        ws.append([labels[key]] + [merged[y].get(key) for y in years])
        if not labels[key].startswith("ㆍ") and key in std.values():
            ws.cell(ws.max_row, 1).font = Font(bold=True)
    for row in ws.iter_rows(min_row=2, min_col=2):
        for c in row:
            c.number_format = "#,##0;[Red]-#,##0"

    wr = wb.create_sheet("주요비율")
    wr.append(["비율"] + years)
    ratio_rows = ratios(merged, std)
    for name in ratio_rows[years[0]]:
        wr.append([name] + [ratio_rows[y][name] for y in years])
    for row in wr.iter_rows(min_row=2, min_col=2):
        for c in row:
            c.number_format = "0.0%"
    wr.append([])
    wr.append(["비율 계산에 쓴 원본 항목"])
    for name, key in std.items():
        wr.append([name, labels.get(key, "찾지 못함")])

    wc = wb.create_sheet("검증")
    wc.append(["연도", "항목", "채택값(최신 보고서)", "과거 보고서 값", "과거 보고서", "일치"])
    for year, label, kept, other, source, ok in checks:
        wc.append([year, label, kept, other, source, "O" if ok else "X (재작성 여부 확인)"])
    wc.append([])
    wc.append(["연도", "합계 검증", "차이(백만원)"])
    for row in sums:
        wc.append(list(row))
    wc.append([])
    wc.append(["확인 필요 (열 개수가 맞지 않아 자동으로 넣지 않은 줄)"])
    for source, line in unread:
        wc.append([line, source])
    wc.append([])
    wc.append(["원본 파일", "표 시작 페이지"])
    for source, page in sources:
        wc.append([source, page])

    fill = PatternFill("solid", fgColor="DDE7F3")
    for sheet in wb.worksheets:
        for c in sheet[1]:
            c.font = Font(bold=True)
            c.fill = fill
        sheet.column_dimensions["A"].width = 42
        for col in "BCDEFG":
            sheet.column_dimensions[col].width = 20
        sheet.freeze_panes = "B2"
    wb.save(path)


def main(argv=None):
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("pdfs", nargs="+")
    ap.add_argument("-o", "--output", default="요약재무.xlsx")
    args = ap.parse_args(argv)

    reports, sources, unread = [], [], []
    for path in args.pdfs:
        with pdfplumber.open(path) as pdf:
            page, body = find_section(pdf)
        years, rows, skipped = parse_section(body)
        reports.append((path, years, rows))
        sources.append((path, page))
        unread += [(path, line) for line in skipped]
        print(f"{path}: {page}쪽, {min(years)}~{max(years)}년, 항목 {len(rows)}개, 확인 필요 {len(skipped)}줄")

    merged, labels, order, checks = merge(reports)
    std = standard_items(order)
    sums = sum_checks(merged, std)
    write_excel(args.output, merged, labels, order, std, checks, sums, unread, sources)

    missing = [name for name, key in std.items() if key is None]
    print(f"저장: {args.output} ({min(merged)}~{max(merged)}년)")
    print(f"  겹치는 연도 값 {len(checks)}개 중 불일치 {sum(not c[5] for c in checks)}개")
    print(f"  합계 검증 {len(sums)}개 중 불일치 {sum(abs(s[2]) > 1 for s in sums)}개")  # 단위 환산 반올림 1백만원까지 허용
    if missing:
        print(f"  표준 항목을 찾지 못함: {', '.join(missing)}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
