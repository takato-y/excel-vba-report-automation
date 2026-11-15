# -*- coding: utf-8 -*-
"""配布用のブックと README 用の画像を作る。

1. build.py で src/ からブックを組み立てる
2. サンプルデータでレポートを作成した状態で保存する（開いてすぐ出力例を見られるようにするため）
3. Excel の画面の範囲をそのまま画像にして docs/images/ に保存する
"""
import re
import shutil
import zipfile
from pathlib import Path

import build as B
import excel_com as X

ROOT = Path(__file__).resolve().parent.parent
IMG = ROOT / "docs" / "images"

# 出力例を作る間だけ PDF 出力を止め、保存するブックの設定は既定（する）に戻す
POPULATE = '''
Public Function T_Populate() As String
    Dim st As String
    NameRange(NAME_PDF).Value = "しない"
    st = RunReportForTest()
    NameRange(NAME_PDF).Value = "する"
    T_Populate = st
End Function
'''


def sanitize_ooxml(path, author="takato-y"):
    """保存したブックから、作成者名と保存時の絶対パスを取り除く。

    Excel は保存時に Office のユーザー名を作成者として、保存場所のフルパスを workbook.xml に書き込む。
    公開するブックに作業環境の情報を残さないため、配布前に書き換える（VBA のバイナリには触れない）。
    """
    src = Path(path)
    tmp = src.with_suffix(".sanitize.tmp")
    with zipfile.ZipFile(src) as zin, zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as zout:
        for item in zin.infolist():
            data = zin.read(item.filename)
            if item.filename == "docProps/core.xml":
                text = data.decode("utf-8")
                for tag in ("dc:creator", "cp:lastModifiedBy"):
                    text = re.sub(rf"<{tag}>[^<]*</{tag}>", f"<{tag}>{author}</{tag}>", text)
                data = text.encode("utf-8")
            elif item.filename == "xl/workbook.xml":
                text = data.decode("utf-8")
                text = re.sub(r"<mc:AlternateContent[^>]*>(?:(?!</mc:AlternateContent>).)*?x15ac:absPath.*?</mc:AlternateContent>",
                              "", text, flags=re.S)
                data = text.encode("utf-8")
            zout.writestr(item, data)
    shutil.move(tmp, src)


def main():
    built = B.build()
    work = X.win_temp_dir() / "sales_report_release"
    if work.exists():
        shutil.rmtree(work)
    work.mkdir(parents=True)
    shutil.copy(built, work / "sales_report_tool.xlsm")
    X.copy_tree(ROOT / "sample_data", work / "sample_data")
    IMG.mkdir(parents=True, exist_ok=True)

    final = work / "release.xlsm"
    r = X.run_job({"open": True, "book": X.winpath(work / "sales_report_tool.xlsm"), "modules": [],
                   "testCode": POPULATE, "call": {"macro": "T_Populate", "args": []},
                   "render": [
                       {"sheet": "レポート", "range": "A1:F29", "out": X.winpath(work / "report.png")},
                       {"sheet": "明細", "range": "A1:H16", "out": X.winpath(work / "detail.png")},
                       {"sheet": "設定", "range": "A1:C22", "out": X.winpath(work / "settings.png")},
                   ],
                   "saveAs": X.winpath(final)}, work, timeout=240)
    print("populate:", r.get("ok"), r.get("status"), r.get("error") or "")
    if not r.get("ok") or not str(r.get("status")).startswith("正常終了"):
        raise SystemExit(1)

    # エラー一覧の例は、不正な行を含むテスト用データで作る
    work2 = X.win_temp_dir() / "sales_report_release_err"
    if work2.exists():
        shutil.rmtree(work2)
    work2.mkdir(parents=True)
    shutil.copy(built, work2 / "book.xlsm")
    X.copy_tree(ROOT / "tests" / "fixtures" / "invalid", work2 / "sample_data")
    r2 = X.run_job({"open": True, "book": X.winpath(work2 / "book.xlsm"), "modules": [],
                    "testCode": POPULATE, "call": {"macro": "T_Populate", "args": []},
                    "render": [{"sheet": "エラー一覧", "range": "A1:C6", "out": X.winpath(work2 / "errors.png")},
                               {"sheet": "レポート", "range": "A1:D10", "out": X.winpath(work2 / "summary_err.png")}]},
                   work2, timeout=240)
    print("errors :", r2.get("ok"), r2.get("status"), r2.get("error") or "")

    shutil.copy(final, ROOT / "sales_report_tool.xlsm")
    sanitize_ooxml(ROOT / "sales_report_tool.xlsm")
    for src, name in [(work / "report.png", "report.png"), (work / "detail.png", "detail.png"),
                      (work / "settings.png", "settings.png"), (work2 / "errors.png", "errors.png"),
                      (work2 / "summary_err.png", "summary_with_errors.png")]:
        if src.exists():
            shutil.copy(src, IMG / name)
            print("image  :", IMG / name)


if __name__ == "__main__":
    main()
