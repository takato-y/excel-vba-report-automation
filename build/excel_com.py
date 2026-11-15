# -*- coding: utf-8 -*-
"""WSL から Windows の Excel を COM で動かす共通部品。

安全のための決まりごと:
- 利用者が画面で Excel を開いているときは実行しない（作業中のブックに割り込まないため）
- 終了処理では、実行前後の EXCEL.EXE のプロセス番号を比べ、自分が起動したものだけを終了する
- PowerShell 5.1 は BOM の無い .ps1 を Shift_JIS として読むため、実行直前に BOM 付きの写しを作る
"""
import datetime as dt
import json
import os
import shutil
import subprocess
from pathlib import Path

WIN = "/mnt/c/WINDOWS/System32"
PWSH = f"{WIN}/WindowsPowerShell/v1.0/powershell.exe"
TASKLIST = f"{WIN}/tasklist.exe"
TASKKILL = f"{WIN}/taskkill.exe"
HERE = Path(__file__).resolve().parent
JOB_PS1 = HERE / "excel_job.ps1"


def winpath(p):
    return subprocess.run(["wslpath", "-w", str(p)], capture_output=True, text=True).stdout.strip()


def win_temp_dir():
    """Windows 側の一時フォルダ（Excel からは WSL のパスが見えないため）。"""
    if os.environ.get("WIN_WORKDIR"):
        return Path(os.environ["WIN_WORKDIR"])
    out = subprocess.run([f"{WIN}/cmd.exe", "/c", "echo %TEMP%"], capture_output=True,
                         text=True, cwd="/mnt/c").stdout.strip()
    return Path(subprocess.run(["wslpath", "-u", out], capture_output=True, text=True).stdout.strip())


def excel_pids():
    out = subprocess.run([TASKLIST, "/FI", "IMAGENAME eq EXCEL.EXE", "/FO", "CSV", "/NH"],
                         capture_output=True, timeout=30).stdout.decode("utf-8", "replace")
    pids = set()
    for line in out.splitlines():
        cols = [c.strip('"') for c in line.split('","')]
        if len(cols) >= 2 and cols[1].isdigit():
            pids.add(cols[1])
    return pids


def visible_excel_title():
    out = subprocess.run([PWSH, "-NoProfile", "-Command",
                          "Get-Process EXCEL -ErrorAction SilentlyContinue | "
                          "Where-Object { $_.MainWindowTitle -ne '' } | "
                          "Select-Object -First 1 -ExpandProperty MainWindowTitle"],
                         capture_output=True, timeout=30)
    return out.stdout.decode("utf-8", "replace").strip() or None


def run_job(spec, workdir, timeout=180):
    """excel_job.ps1 を1回実行し、結果の dict を返す。"""
    title = visible_excel_title()
    if title:
        return {"ok": False, "error": f"Excel が画面で開かれているため中断しました（{title}）。閉じてから再実行してください。"}
    workdir = Path(workdir)
    workdir.mkdir(parents=True, exist_ok=True)
    stamp = dt.datetime.now().strftime("%H%M%S_%f")
    spec_path = workdir / f"spec_{stamp}.json"
    result_path = workdir / f"result_{stamp}.json"
    spec = dict(spec, result=winpath(result_path))
    spec_path.write_text(json.dumps(spec, ensure_ascii=False), encoding="utf-8")
    ps1 = workdir / f"job_{stamp}.ps1"
    ps1.write_bytes(b"\xef\xbb\xbf" + JOB_PS1.read_bytes())

    before = excel_pids()
    try:
        subprocess.run([PWSH, "-NoProfile", "-ExecutionPolicy", "Bypass", "-File", winpath(ps1),
                        "-Spec", winpath(spec_path)], capture_output=True, timeout=timeout)
        if result_path.exists():
            return json.loads(result_path.read_text(encoding="utf-8-sig"))
        return {"ok": False, "error": "結果ファイルが作られませんでした"}
    except subprocess.TimeoutExpired:
        return {"ok": False, "error": f"timeout ({timeout}s)"}
    finally:
        for pid in excel_pids() - before:
            subprocess.run([TASKKILL, "/F", "/PID", pid], capture_output=True)
        for p in (spec_path, ps1, result_path):
            try:
                p.unlink()
            except OSError:
                pass


def read_tsv(path):
    """excel_job.ps1 が書いた TSV を、シートの行番号・列番号つきの dict に戻す。"""
    lines = Path(path).read_text(encoding="utf-8").split("\n")
    _, r0, c0 = lines[0].split("\t")
    cells = {}
    for i, line in enumerate(lines[1:]):
        if line == "":
            continue
        for j, v in enumerate(line.split("\t")):
            if v != "":
                cells[(int(r0) + i, int(c0) + j)] = v
    return cells


def copy_tree(src, dst):
    if Path(dst).exists():
        shutil.rmtree(dst)
    shutil.copytree(src, dst)
