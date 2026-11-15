# -*- coding: utf-8 -*-
"""実機の Excel でツールを動かし、帳票の数値を Python の独立計算と突き合わせる。

使い方:
    python3 tests/run_tests.py            # PDF 以外の全ケース
    python3 tests/run_tests.py --pdf      # PDF 出力のケースも実行（画面なしの Excel では止まることがある）
    python3 tests/run_tests.py --only TC-001

各ケースはブックの写しを Windows の一時フォルダに置いて実行し、元のブックは変更しない。
"""
import argparse
import random
import shutil
import sys
import time
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT / "build"))
sys.path.insert(0, str(ROOT / "tests"))
import excel_com as X        # noqa: E402
import expected as E         # noqa: E402
import gen_sample_data as G  # noqa: E402

# 設定シートに値を入れてからボタンと同じ処理を呼ぶ、テスト専用のモジュール（ブックには保存しない）
DRIVER = '''
Public Function T_Run(inp As Variant, outp As Variant, cs As Variant, tax As Variant, pdf As Variant) As String
    Dim t As Double, st As String
    NameRange(NAME_INPUT).Value = inp
    NameRange(NAME_OUTPUT).Value = outp
    NameRange(NAME_CHARSET).Value = cs
    NameRange(NAME_TAX).Value = tax
    NameRange(NAME_PDF).Value = pdf
    t = Timer
    st = RunReportForTest()
    T_Run = st & "|sec=" & Format$(Timer - t, "0.00")
End Function
'''

RESULTS = []


def check(tc, name, ok, detail=""):
    RESULTS.append((tc, ok))
    print(f"{'PASS' if ok else 'FAIL'} {tc} {name}" + (f"  ({detail})" if detail else ""))


def run_case(book, case, inputs, inp, outp="output", cs="UTF-8", tax=0.1, pdf="しない",
             timeout=180, extra=None):
    work = X.win_temp_dir() / "sales_report_tests" / case
    if work.exists():
        shutil.rmtree(work)
    work.mkdir(parents=True)
    shutil.copy(book, work / "book.xlsm")
    if inputs is not None:
        X.copy_tree(inputs, work / inp)
    if extra:
        extra(work)
    dumps = {s: work / f"{s}.tsv" for s in ("レポート", "明細", "エラー一覧")}
    t0 = time.time()
    r = X.run_job({"open": True, "book": X.winpath(work / "book.xlsm"), "modules": [],
                   "testCode": DRIVER,
                   "call": {"macro": "T_Run", "args": [inp, outp, cs, tax, pdf]},
                   "dump": [{"sheet": s, "out": X.winpath(p)} for s, p in dumps.items()]},
                  work, timeout=timeout)
    r["wall"] = time.time() - t0
    r["work"] = work
    if r.get("ok"):
        r["sheets"] = {s: X.read_tsv(p) for s, p in dumps.items()}
    return r


def parse_report(cells):
    """帳票の見出しの文字を手がかりに表を読み取る（位置を決め打ちしない）。"""
    def row_of(label, start=1):
        for (rr, cc), v in sorted(cells.items()):
            if cc == 1 and v == label and rr >= start:
                return rr
        raise KeyError(label)

    def num(rr, cc):
        return float(cells.get((rr, cc), "0"))

    out = {}
    h = row_of("顧客名")
    months, c = [], 2
    while cells.get((h, c)) not in ("合計", None):
        months.append(cells[(h, c)])
        c += 1
    total_col = c
    out["months"] = months
    out["cust_order"], out["cust_month"], out["cust_total"], shares = [], {}, {}, 0.0
    rr = h + 1
    while cells.get((rr, 1)) != "合計":
        name = cells[(rr, 1)]
        out["cust_order"].append(name)
        for i, m in enumerate(months):
            out["cust_month"][(name, m)] = num(rr, 2 + i)
        out["cust_total"][name] = num(rr, total_col)
        shares += num(rr, total_col + 1)
        rr += 1
    out["matrix_grand"] = num(rr, total_col)
    out["cust_share_sum"] = shares
    out["month_totals"] = {m: num(rr, 2 + i) for i, m in enumerate(months)}

    p = row_of("商品名")
    out["prod_order"], out["prod_qty"], out["prod_amt"] = [], {}, {}
    rr = p + 1
    while cells.get((rr, 1)) != "合計":
        name = cells[(rr, 1)]
        out["prod_order"].append(name)
        out["prod_qty"][name] = num(rr, 2)
        out["prod_amt"][name] = num(rr, 3)
        rr += 1
    out["prod_grand"] = num(rr, 3)
    out["grand"] = num(row_of("売上合計（税抜）"), 2)
    out["tax"] = num(row_of("消費税"), 2)
    out["total_with_tax"] = num(row_of("売上合計（税込）"), 2)
    out["rows"] = num(row_of("取込明細件数"), 2)
    out["errors"] = num(row_of("エラー件数"), 2)
    return out


def close(a, b):
    return abs(a - b) < 0.005


def compare_report(tc, r, exp):
    rep = parse_report(r["sheets"]["レポート"])
    check(tc, "対象月", rep["months"] == exp["months"], ",".join(rep["months"]))
    check(tc, "顧客の並び（売上の多い順）", rep["cust_order"] == exp["cust_order"])
    bad = [k for k in exp["cust_total"] if not close(rep["cust_total"].get(k, -1), exp["cust_total"][k])]
    check(tc, f"顧客別合計 {len(exp['cust_total'])}件", not bad, f"不一致={bad}" if bad else "")
    keys = {(c, m) for c in exp["cust_order"] for m in exp["months"]}
    bad = [k for k in keys if not close(rep["cust_month"].get(k, -1), exp["cust_month"].get(k, 0.0))]
    check(tc, f"顧客×月 {len(keys)}セル", not bad, f"不一致={bad[:3]}" if bad else "")
    mt = {m: sum(v for (c, mm), v in exp["cust_month"].items() if mm == m) for m in exp["months"]}
    check(tc, "月別合計", all(close(rep["month_totals"][m], mt[m]) for m in mt))
    check(tc, "商品の並び（売上の多い順）", rep["prod_order"] == exp["prod_order"])
    bad = [k for k in exp["prod_amt"] if not (close(rep["prod_amt"].get(k, -1), exp["prod_amt"][k])
                                             and close(rep["prod_qty"].get(k, -1), exp["prod_qty"][k]))]
    check(tc, f"商品別 数量・金額 {len(exp['prod_amt'])}件", not bad, f"不一致={bad}" if bad else "")
    check(tc, "売上合計（税抜）", close(rep["grand"], exp["grand"]) and close(rep["matrix_grand"], exp["grand"])
          and close(rep["prod_grand"], exp["grand"]), f"{rep['grand']:,.0f} / 期待 {exp['grand']:,.0f}")
    check(tc, "消費税（切り捨て）", close(rep["tax"], exp["tax"]), f"{rep['tax']:,.0f} / 期待 {exp['tax']:,.0f}")
    check(tc, "売上合計（税込）", close(rep["total_with_tax"], exp["total_with_tax"]),
          f"{rep['total_with_tax']:,.0f} / 期待 {exp['total_with_tax']:,.0f}")
    check(tc, "構成比の合計=100%", close(rep["cust_share_sum"], 1.0))
    check(tc, "取込件数・エラー件数", rep["rows"] == exp["rows"] and rep["errors"] == len(exp["errors"]),
          f"取込 {rep['rows']:.0f} / エラー {rep['errors']:.0f}")
    detail_rows = max(rr for (rr, cc) in r["sheets"]["明細"]) - 1
    check(tc, "明細シートの行数", detail_rows == exp["rows"], f"{detail_rows}行")


def error_list(r):
    cells = r["sheets"]["エラー一覧"]
    rows = sorted({rr for (rr, cc) in cells if rr > 1})
    return [(cells.get((rr, 1)), cells.get((rr, 2)), cells.get((rr, 3), "")) for rr in rows]


def status_of(r):
    return (r.get("status") or "").split("|sec=")[0]


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--book", default=str(ROOT / "sales_report_tool.xlsm"))
    ap.add_argument("--pdf", action="store_true")
    ap.add_argument("--only")
    a = ap.parse_args()
    book = Path(a.book)
    fix = ROOT / "tests" / "fixtures"

    def want(tc):
        return a.only is None or a.only == tc

    if want("TC-001"):
        r = run_case(book, "tc001", ROOT / "sample_data", "sample_data")
        check("TC-001", "正常系の実行", r.get("ok") and status_of(r).startswith("正常終了"),
              status_of(r) or r.get("error"))
        if r.get("ok"):
            compare_report("TC-001", r, E.compute(ROOT / "sample_data"))
            check("TC-001", "エラー一覧が空", error_list(r) == [])

    if want("TC-002"):
        def add_lock(work):
            (work / "in" / "~$売上_一時.csv").write_text("壊れたファイル", encoding="utf-8")
        r = run_case(book, "tc002", fix / "invalid", "in", extra=add_lock)
        check("TC-002", "不正行を含む入力の実行", r.get("ok") and status_of(r).startswith("正常終了"),
              status_of(r) or r.get("error"))
        if r.get("ok"):
            exp = E.compute(fix / "invalid")
            compare_report("TC-002", r, exp)
            got = [(f, int(l) if l not in ("-", None) else 0) for f, l, _ in error_list(r)]
            check("TC-002", "エラー一覧のファイル名・行番号", sorted(got) == sorted(exp["errors"]),
                  f"{got}")
            texts = [t for _, _, t in error_list(r)]
            for kw in ("日付の形式", "顧客名が空欄", "数量が数値ではありません", "数量×単価 と一致しません", "見出し行が想定と異なる"):
                check("TC-002", f"理由「{kw}」が記録される", any(kw in t for t in texts))
            check("TC-002", "~$ で始まる一時ファイルと .txt は読まない",
                  not any(f.startswith("~$") or f.endswith(".txt") for f, _, _ in error_list(r)))

    if want("TC-003"):
        r = run_case(book, "tc003", None, "存在しないフォルダ")
        check("TC-003", "入力フォルダが無い→日本語のエラー", r.get("ok") and
              status_of(r).startswith("エラー：入力フォルダが見つかりません"), status_of(r) or r.get("error"))

    if want("TC-004"):
        r = run_case(book, "tc004", None, "empty", extra=lambda w: (w / "empty").mkdir())
        check("TC-004", "ファイルが無い→日本語のエラー", r.get("ok") and
              status_of(r).startswith("エラー：入力フォルダに CSV / Excel ファイルがありません"),
              status_of(r) or r.get("error"))

    if want("TC-005"):
        r = run_case(book, "tc005", ROOT / "sample_data", "sample_data", tax="abc")
        ok1 = r.get("ok") and status_of(r).startswith("エラー：「消費税率」")
        r2 = run_case(book, "tc005b", ROOT / "sample_data", "sample_data", pdf="たぶん")
        ok2 = r2.get("ok") and status_of(r2).startswith("エラー：「PDF出力」")
        check("TC-005", "設定値の誤り→日本語のエラー", ok1 and ok2, f"{status_of(r)} / {status_of(r2)}")

    if want("TC-006"):
        r = run_case(book, "tc006", fix / "sjis", "in", cs="Shift_JIS")
        check("TC-006", "Shift_JIS の CSV", r.get("ok") and status_of(r).startswith("正常終了"),
              status_of(r) or r.get("error"))
        if r.get("ok"):
            compare_report("TC-006", r, E.compute(fix / "sjis", encoding="cp932"))

    if want("TC-007"):
        r = run_case(book, "tc007", fix / "all_invalid", "in")
        check("TC-007", "取り込める行が0件→日本語のエラー", r.get("ok") and
              status_of(r).startswith("エラー：取り込める明細が1件もありませんでした"),
              status_of(r) or r.get("error"))
        if r.get("ok"):
            check("TC-007", "0件でもエラー一覧に理由が残る", len(error_list(r)) == 1)

    if want("TC-008"):
        perf_src = X.win_temp_dir() / "sales_report_tests" / "perf_input"
        if perf_src.exists():
            shutil.rmtree(perf_src)
        perf_src.mkdir(parents=True)
        rng = random.Random(7)
        for m in range(1, 13):
            G.write_csv(perf_src / f"売上_2025-{m:02d}.csv", G.make_rows(2025, m, 1000, rng))
        r = run_case(book, "tc008", perf_src, "in", timeout=600)
        sec = (r.get("status") or "|sec=?").split("|sec=")[-1]
        check("TC-008", "12ファイル×1,000行（12,000行）", r.get("ok") and status_of(r).startswith("正常終了"),
              f"マクロ処理時間 {sec} 秒 / Excel 起動を含む全体 {r['wall']:.1f} 秒")
        if r.get("ok"):
            compare_report("TC-008", r, E.compute(perf_src))

    if a.pdf and want("TC-009"):
        r = run_case(book, "tc009", ROOT / "sample_data", "sample_data", pdf="する", timeout=150)
        pdfs = list((r["work"] / "output").glob("*.pdf")) if (r["work"] / "output").exists() else []
        ok = r.get("ok") and status_of(r).startswith("正常終了") and pdfs and \
            pdfs[0].read_bytes()[:5] == b"%PDF-"
        check("TC-009", "PDF 出力", bool(ok), status_of(r) or r.get("error"))
        if pdfs:
            print(f"     PDF: {pdfs[0]} ({pdfs[0].stat().st_size:,} bytes)")

    if want("TC-010"):
        # 上の階層が無い出力先。PDF の書き出し前に止まるため、画面なしの Excel でも安全に試せる
        r = run_case(book, "tc010", ROOT / "sample_data", "sample_data", outp="無い階層\\output", pdf="する")
        check("TC-010", "出力フォルダを作れない→日本語のエラー", r.get("ok") and
              status_of(r).startswith("エラー：出力フォルダを作成できませんでした"), status_of(r) or r.get("error"))

    n_ok = sum(1 for _, ok in RESULTS if ok)
    print(f"\n合計 {len(RESULTS)} 件中 PASS {n_ok} / FAIL {len(RESULTS) - n_ok}")
    return 0 if n_ok == len(RESULTS) else 1


if __name__ == "__main__":
    sys.exit(main())
