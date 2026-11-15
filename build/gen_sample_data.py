# -*- coding: utf-8 -*-
"""架空の月別売上ファイルを作る。

乱数の種を固定し、何度作り直しても同じ内容になるようにしている
（テストの期待値と画像を再現できるようにするため）。
4月・5月は CSV（UTF-8 BOM付き）、6月は Excel にして、形式が混在する現場を再現する。
"""
import csv
import datetime as dt
import random
import sys
from pathlib import Path

import openpyxl

HEADER = ["日付", "顧客名", "商品名", "数量", "単価", "金額"]
CUSTOMERS = ["サンプル商事", "テスト物産", "架空工業", "みほん電機", "ダミー食品"]
# カンマを含む商品名は、CSV のダブルクォート処理を確かめるために入れている
PRODUCTS = {
    "架空ノートA": 320,
    "試作ペンB": 150,
    "見本ファイルC（A4,10冊入）": 1200,
    "サンプル用紙D": 680,
    "デモ封筒E": 95,
}


def make_rows(year, month, n, rng):
    rows = []
    for _ in range(n):
        day = rng.randint(1, 28)
        cust = rng.choice(CUSTOMERS)
        prod = rng.choice(list(PRODUCTS))
        qty = rng.randint(1, 60)
        price = PRODUCTS[prod]
        rows.append([dt.date(year, month, day), cust, prod, qty, price, qty * price])
    rows.sort(key=lambda r: (r[0], r[1], r[2]))
    return rows


def write_csv(path, rows, encoding="utf-8-sig"):
    with open(path, "w", encoding=encoding, newline="") as f:
        w = csv.writer(f, lineterminator="\r\n")
        w.writerow(HEADER)
        for r in rows:
            w.writerow([r[0].strftime("%Y/%m/%d")] + r[1:])


def write_xlsx(path, rows):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "売上"
    ws.append(HEADER)
    for r in rows:
        ws.append(r)
    for c in ws["A"][1:]:
        c.number_format = "yyyy/mm/dd"
    wb.save(path)


def main(out_dir):
    out = Path(out_dir)
    out.mkdir(parents=True, exist_ok=True)
    rng = random.Random(20260401)
    write_csv(out / "売上_2026-04.csv", make_rows(2026, 4, 32, rng))
    write_csv(out / "売上_2026-05.csv", make_rows(2026, 5, 35, rng))
    write_xlsx(out / "売上_2026-06.xlsx", make_rows(2026, 6, 30, rng))


if __name__ == "__main__":
    main(sys.argv[1] if len(sys.argv) > 1 else Path(__file__).resolve().parent.parent / "sample_data")
