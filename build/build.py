# -*- coding: utf-8 -*-
"""src/ の VBA ソースから sales_report_tool.xlsm を組み立てる（ADR-0003）。

手順: 新しいブック → src/ のモジュールを取り込み → modSetup でシートと設定を作成
      → modSetup を取り除いて .xlsm で保存。
"""
import shutil
import sys
from pathlib import Path

import excel_com as X

ROOT = Path(__file__).resolve().parent.parent
MODULES = ["modUtil.bas", "modConfig.bas", "clsSalesRow.cls", "modImport.bas",
           "modAggregate.bas", "modReport.bas", "modExport.bas", "modMain.bas"]


def build(out_path=None):
    work = X.win_temp_dir() / "sales_report_build"
    if work.exists():
        shutil.rmtree(work)
    work.mkdir(parents=True)
    mods = []
    for name in MODULES:
        src = ROOT / "src" / name
        if not src.exists():
            continue                     # 開発途中は存在するモジュールだけで組み立てる
        shutil.copy(src, work / name)
        mods.append({"path": X.winpath(work / name), "type": 2 if name.endswith(".cls") else 1})
    shutil.copy(ROOT / "build" / "modSetup.bas", work / "modSetup.bas")
    mods.append({"path": X.winpath(work / "modSetup.bas"), "type": 1})

    tmp_book = work / "sales_report_tool.xlsm"
    r = X.run_job({"open": False, "modules": mods, "init": "SetupWorkbook",
                   "remove": ["modSetup"], "saveAs": X.winpath(tmp_book)}, work)
    if not r.get("ok"):
        raise SystemExit(f"BUILD NG: {r.get('error')}")
    out = Path(out_path) if out_path else ROOT / "sales_report_tool.xlsm"
    shutil.copy(tmp_book, out)
    print(f"BUILD OK: modules={r.get('modules')} -> {out}")
    return out


if __name__ == "__main__":
    build(sys.argv[1] if len(sys.argv) > 1 else None)
