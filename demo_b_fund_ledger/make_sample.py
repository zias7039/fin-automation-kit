"""데모용 가상 펀드 잔고 엑셀을 만든다 (실제 고객 데이터 아님).

판매사 시스템에서 흔히 받는 형태처럼 '기준일별 잔고' 시트만 있고 거래내역은 없다.
검증용으로 실제 거래(정답)를 따로 CSV로 저장한다.

사용법:
    python make_sample.py -o samples/잔고_가상펀드.xlsx --truth samples/정답_거래내역.csv
"""
import argparse
import csv
import random
from datetime import date

from openpyxl import Workbook

DATES = [date(2025, m, 1) for m in range(1, 13)]
NAMES = ["김민준", "이서연", "박도윤", "최지우", "정하준", "강서윤", "조은우", "윤지아", "장시우", "임하은", "한주원", "오수아"]


def generate(seed=7):
    rng = random.Random(seed)
    nav = 1000.0
    holdings, snapshots, truth = {}, [], []
    investors = [(f"{10000000 + i * 137:08d}", n) for i, n in enumerate(NAMES)]
    for d in DATES:
        if d != DATES[0]:
            nav = round(nav * (1 + rng.gauss(0.006, 0.02)), 2)
        for inv_id, name in investors:
            units = holdings.get(inv_id, 0)
            r = rng.random()
            if units == 0 and (d == DATES[0] and r < 0.6 or r < 0.12):
                kind, delta = "설정", rng.randrange(50, 500) * 1_000_000
            elif units and r < 0.10:
                kind, delta = "추가설정", rng.randrange(10, 200) * 1_000_000
            elif units and r < 0.15:
                kind, delta = "일부환매", -min(units // 2, rng.randrange(10, 150) * 1_000_000)
            elif units and r < 0.18:
                kind, delta = "전액환매", -units
            else:
                continue
            holdings[inv_id] = units + delta
            truth.append((d.isoformat(), inv_id, name, kind, delta))
        rows = [(i, n, holdings[i], nav, round(holdings[i] * nav / 1000)) for i, n in investors if holdings.get(i)]
        snapshots.append((d, rows))
    return snapshots, truth


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("-o", "--output", default="samples/잔고_가상펀드.xlsx")
    ap.add_argument("--truth", default="samples/정답_거래내역.csv")
    args = ap.parse_args()

    snapshots, truth = generate()
    wb = Workbook()
    wb.remove(wb.active)
    for d, rows in snapshots:
        ws = wb.create_sheet(d.strftime("%Y%m%d"))
        ws.append(["가상 IPO 공모주 펀드 수익자 잔고현황", None, f"기준일: {d.isoformat()}"])
        ws.append([])
        ws.append(["수익자번호", "수익자명", "보유좌수", "기준가격", "평가금액"])
        for row in rows:
            ws.append(list(row))
    wb.save(args.output)

    with open(args.truth, "w", newline="", encoding="utf-8-sig") as f:
        w = csv.writer(f)
        w.writerow(["기준일", "수익자번호", "수익자명", "구분", "좌수변동"])
        w.writerows(truth)
    print(f"{args.output}: 시트 {len(snapshots)}개, 정답 거래 {len(truth)}건 -> {args.truth}")


if __name__ == "__main__":
    main()
