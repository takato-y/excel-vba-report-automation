# -*- coding: utf-8 -*-
"""VBA とは別に Python で同じ入力ファイルを集計し、検算用の期待値を作る。

VBA の実装を写さず、標準の csv モジュールと openpyxl だけで独立に計算する。
入力チェックの規則（見出し・日付・必須・数値・金額一致）は要件定義書の記述から実装している。
"""
import csv
import datetime as dt
import math
from collections import defaultdict
from pathlib import Path

import openpyxl

HEADER = ["日付", "顧客名", "商品名", "数量", "単価", "金額"]


def _to_number(v):
    if v is None or str(v).strip() == "":
        return None
    try:
        return float(str(v).replace(",", ""))
    except ValueError:
        return None


def _to_date(v):
    if isinstance(v, dt.datetime):
        return v.date()
    if isinstance(v, dt.date):
        return v
    s = str(v or "").strip()
    for fmt in ("%Y/%m/%d", "%Y-%m-%d"):
        try:
            return dt.datetime.strptime(s, fmt).date()
        except ValueError:
            pass
    return None


def _read(path, encoding):
    if path.suffix.lower() == ".csv":
        with open(path, encoding=encoding, newline="") as f:
            return [r for r in csv.reader(f)]
    ws = openpyxl.load_workbook(path, data_only=True).worksheets[0]
    return [list(r) for r in ws.iter_rows(values_only=True)]


def compute(folder, encoding="utf-8-sig", tax_rate=0.10):
    files = sorted(p for p in Path(folder).iterdir()
                   if p.suffix.lower() in (".csv", ".xlsx", ".xls") and not p.name.startswith("~$"))
    valid, errors = [], []
    for p in files:
        rows = _read(p, encoding)
        if not rows or [str(c or "").strip() for c in rows[0][:6]] != HEADER:
            errors.append((p.name, 1))
            continue
        for i, r in enumerate(rows[1:], start=2):
            r = (list(r) + [None] * 6)[:6]
            if all(str(c or "").strip() == "" for c in r):
                continue
            d = _to_date(r[0])
            cust, prod = str(r[1] or "").strip(), str(r[2] or "").strip()
            q, u, a = (_to_number(x) for x in r[3:6])
            if d is None or not cust or not prod or None in (q, u, a) or abs(q * u - a) > 0.5:
                errors.append((p.name, i))
                continue
            valid.append((d.strftime("%Y-%m"), cust, prod, q, a))

    cust_month = defaultdict(float)
    cust_total = defaultdict(float)
    prod_qty = defaultdict(float)
    prod_amt = defaultdict(float)
    for m, c, pr, q, a in valid:
        cust_month[(c, m)] += a
        cust_total[c] += a
        prod_qty[pr] += q
        prod_amt[pr] += a
    grand = sum(cust_total.values())
    tax = math.floor(grand * tax_rate)
    return {
        "rows": len(valid),
        "errors": errors,
        "months": sorted({v[0] for v in valid}),
        "cust_month": dict(cust_month),
        "cust_order": sorted(cust_total, key=lambda k: (-cust_total[k], k)),
        "cust_total": dict(cust_total),
        "prod_order": sorted(prod_amt, key=lambda k: (-prod_amt[k], k)),
        "prod_qty": dict(prod_qty),
        "prod_amt": dict(prod_amt),
        "grand": grand,
        "tax": tax,
        "total_with_tax": grand + tax,
    }


if __name__ == "__main__":
    import json
    import sys
    e = compute(sys.argv[1] if len(sys.argv) > 1 else Path(__file__).resolve().parent.parent / "sample_data")
    print(json.dumps({k: v for k, v in e.items() if k != "cust_month"}, ensure_ascii=False, indent=1))
