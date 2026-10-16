"""python test_extract.py 로 실행한다. samples/에 삼성전자 2021·2023 사업보고서 PDF가 있어야 한다."""
import os

import pdfplumber

import extract_summary as es

HERE = os.path.dirname(os.path.abspath(__file__))


def load(name):
    with pdfplumber.open(os.path.join(HERE, "samples", name)) as pdf:
        page, body = es.find_section(pdf)
    years, rows, unread = es.parse_section(body)
    return page, years, {k: v for k, _, v in rows}, unread


def test_samsung_2023_values():
    page, years, rows, unread = load("samsung_2023.pdf")
    assert page == 77 and years == [2023, 2022, 2021] and unread == []
    # PDF 원문 숫자와 대조 (단위: 백만원)
    assert rows["매출액"][2023] == 258_935_494
    assert rows["영업이익"][2023] == 6_566_976
    assert rows["자산총계"][2021] == 426_621_158
    assert rows["기타#3"][2021] == -2_128_473  # 지배기업 소유주지분 > 기타, △ 음수
    assert rows["기본주당순이익(단위:원)"][2022] == 8_057


def test_merge_five_years():
    reports = []
    for name in ("samsung_2023.pdf", "samsung_2021.pdf"):
        _, years, rows, _ = load(name)
        reports.append((name, years, [(k, k, v) for k, v in rows.items()]))
    merged, _, order, checks = es.merge(reports)
    std = es.standard_items(order)
    assert sorted(merged) == [2019, 2020, 2021, 2022, 2023]
    assert checks and all(c[5] for c in checks)  # 2021년 값이 두 보고서에서 같다
    assert all(abs(diff) < 1 for _, _, diff in es.sum_checks(merged, std))
    r = es.ratios(merged, std)
    assert abs(r[2023]["영업이익률"] - 6_566_976 / 258_935_494) < 1e-12
    assert r[2019]["매출액 증가율"] is None


def test_number_formats():
    assert es.to_number("△2,128,473") == -2128473
    assert es.to_number("(1,050,414)") == -1050414
    assert es.to_number("(2,315)원/주") == -2315
    assert es.to_number("233개") == 233
    assert es.to_number("-") == 0


if __name__ == "__main__":
    for name, fn in list(globals().items()):
        if name.startswith("test_"):
            fn()
            print("OK", name)
