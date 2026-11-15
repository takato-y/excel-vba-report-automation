# -*- coding: utf-8 -*-
"""設計書のエラーIDと障害対応手順書の対応を検査する。

使い方:
    python3 tools/check_docs_trace.py          # docs/ を検査
    python3 tools/check_docs_trace.py DIR      # 別のフォルダを検査

検査内容:
    1. 詳細設計書（*詳細設計*.md）・フロー設計書（*フロー設計*.md）に出てくる ERR-xxx が、
       障害対応手順書（*_障害対応手順書.md）の 4 章の見出し（### ERR-xxx 名称）と
       7 章の表（エラーID対応表）の両方にあること
    2. 障害対応手順書に、設計書に無い ERR-xxx が出てこないこと

終了コード: 0 = 合格、1 = 不一致あり、2 = 対象の文書が見つからない
Excel を使わずに実行できる（標準ライブラリのみ）。
"""
import re
import sys
from pathlib import Path

ERR_RE = re.compile(r"(?<![A-Za-z0-9_-])ERR-\d{3}(?![A-Za-z0-9_-])")
# 4 章の見出し。ID はインラインコードで書いてもよい（### `ERR-001` 名称）
HEADING_RE = re.compile(r"^###\s+`?(ERR-\d{3})`?\s")
CHAPTER_RE = re.compile(r"^##\s+(\d+)\.")


def chapters(text):
    """「## N.」で始まる章ごとに行を分ける。{章番号: [行, ...]}"""
    out, current = {}, None
    for line in text.splitlines():
        m = CHAPTER_RE.match(line)
        if m:
            current = int(m.group(1))
            out[current] = []
        elif line.startswith("## "):
            current = None
        elif current is not None:
            out[current].append(line)
    return out


def check(docs_dir):
    """不一致の説明を並べたリストを返す（空なら合格）。"""
    docs_dir = Path(docs_dir)
    design_files = sorted(set(docs_dir.glob("*詳細設計*.md")) | set(docs_dir.glob("*フロー設計*.md")))
    runbooks = sorted(docs_dir.glob("*_障害対応手順書.md"))
    if not design_files:
        raise FileNotFoundError(f"詳細設計書・フロー設計書が見つかりません: {docs_dir}")
    if len(runbooks) != 1:
        raise FileNotFoundError(f"障害対応手順書が1つに定まりません（{len(runbooks)} 件）: {docs_dir}")

    design_ids = set()
    for f in design_files:
        design_ids |= set(ERR_RE.findall(f.read_text(encoding="utf-8-sig")))
    if not design_ids:
        raise FileNotFoundError("設計書に ERR-xxx が1件もありません")

    text = runbooks[0].read_text(encoding="utf-8-sig")
    ch = chapters(text)
    heading_ids = {m.group(1) for line in ch.get(4, []) if (m := HEADING_RE.match(line))}
    table_ids = set()
    for line in ch.get(7, []):
        cells = [c.strip() for c in line.strip().strip("|").split("|")]
        if line.lstrip().startswith("|") and cells and ERR_RE.fullmatch(cells[0]):
            table_ids.add(cells[0])
    runbook_ids = set(ERR_RE.findall(text))

    problems = []
    for err in sorted(design_ids - heading_ids):
        problems.append(f"{err}: 設計書にあるが、障害対応手順書の 4 章に見出しが無い")
    for err in sorted(design_ids - table_ids):
        problems.append(f"{err}: 設計書にあるが、障害対応手順書の 7 章の表に無い")
    for err in sorted(runbook_ids - design_ids):
        problems.append(f"{err}: 障害対応手順書にあるが、設計書に無い")
    return problems, design_ids, runbooks[0]


def main(argv):
    docs_dir = Path(argv[1]) if len(argv) > 1 else Path(__file__).resolve().parent.parent / "docs"
    try:
        problems, ids, runbook = check(docs_dir)
    except FileNotFoundError as exc:
        print(f"ERROR {exc}")
        return 2
    for p in problems:
        print(f"FAIL {p}")
    status = "PASS" if not problems else "FAIL"
    print(f"{status} 設計書のエラーID {len(ids)} 件と {runbook.name} の対応（不一致 {len(problems)} 件）")
    return 0 if not problems else 1


if __name__ == "__main__":
    sys.exit(main(sys.argv))
