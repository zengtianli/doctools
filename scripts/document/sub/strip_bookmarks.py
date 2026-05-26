#!/usr/bin/env python3
# distilled from qual-supply/scripts/strip_bookmarks.py (2026-05-25 W1)
# -*- coding: utf-8 -*-
"""
strip_bookmarks.py
==================

单功能描述
----------
从 docx 移除指定前缀的 `<w:bookmarkStart>` 和 `<w:bookmarkEnd>` (按 `w:id` 配对),
默认只删 Word 自动生成的 bookmark (`_Toc` / `_Ref` / `_Hlk` / `_GoBack`), 保留用户
手动命名的 bookmark。

合稿时多份 docx 的 bookmark ID 会冲突 (id="0" 撞 id="0") + `_Toc/_Ref` 已被
`freeze_all_fields` 冻成静态文本,锚点也就失去意义 → 直接清掉避免乱锚。

触发场景
--------
- 合稿后 bookmark ID 冲突 (Word 报"书签未定义")
- freeze_all_fields 已冻 TOC/PAGEREF → `_Toc`/`_Ref` anchor 已无用,清掉减小文件 + 避免 Word `Update Field` 弹框
- 验收前清场, 只留用户手动命名 bookmark (如果有)

CLI
---
    python3 scripts/strip_bookmarks.py <docx> \\
        [--prefixes _Toc,_Ref,_Hlk] [--dry-run] [--no-backup] [--report <json>]

默认行为
--------
- `--prefixes _Toc,_Ref,_Hlk` (`_GoBack` 总是删, Word 默认自动 bookmark)
- `--prefixes all` 删全部 bookmark (包括用户手动命名, 慎用)
- 自动 `--backup` 写 `.bak-N-YYYY-MM-DD.docx`
- 写前 `lsof` 自检 Word/WPS 占用

不许做
------
- 删用户手动命名 bookmark (除非 `--prefixes all` 显式要)
- 改 fldChar/instrText (上轮 freeze_all_fields 已处理)
- 改 hyperlink (那是另一个脚本的活)
- commit / push
- 用 sed/awk
"""
from __future__ import annotations

import argparse
import datetime
import json
import shutil
import subprocess
import sys
from collections import Counter
from pathlib import Path

try:
    from docx import Document
except ImportError:
    print("[ERR] 缺 python-docx: pip install python-docx", file=sys.stderr)
    sys.exit(2)

NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
W = f"{{{NS}}}"

DEFAULT_PREFIXES = ["_Toc", "_Ref", "_Hlk"]


def lsof_check(p: Path) -> None:
    try:
        r = subprocess.run(["lsof", str(p)], capture_output=True, text=True, timeout=5)
        if r.stdout.strip():
            print(f"[ERR] 文件被占用 (Word/WPS 没关?):\n{r.stdout}", file=sys.stderr)
            sys.exit(2)
    except (FileNotFoundError, subprocess.TimeoutExpired):
        pass


def backup_path(p: Path) -> Path:
    today = datetime.date.today().isoformat()
    n = 1
    while True:
        b = p.with_name(f"{p.stem}.bak-{n}-{today}{p.suffix}")
        if not b.exists():
            shutil.copy2(p, b)
            return b
        n += 1


def parse_prefixes(s: str) -> list[str] | None:
    s = (s or "").strip()
    if not s:
        return list(DEFAULT_PREFIXES)
    if s.lower() == "all":
        return None  # None 表示删全部
    return [p.strip() for p in s.split(",") if p.strip()]


def _should_remove(name: str, prefixes: list[str] | None) -> bool:
    """name 是 bookmark name; prefixes=None 表示删全部."""
    if name == "_GoBack":
        # Word 默认自动 bookmark, 任何场景都安全清掉
        return True
    if prefixes is None:
        return True
    for p in prefixes:
        if name.startswith(p):
            return True
    return False


def strip(doc, prefixes: list[str] | None) -> dict:
    """从 doc 移除匹配的 bookmark, 返回 report dict."""
    body = doc.element.body

    # 1. 扫所有 bookmarkStart, 按 id 收集要删的 (按 name prefix 判)
    ids_to_delete: set[str] = set()
    removed_names: list[str] = []
    kept_names: list[str] = []

    for bs in body.iter(f"{W}bookmarkStart"):
        bid = bs.get(f"{W}id")
        name = bs.get(f"{W}name") or ""
        if bid is None:
            continue
        if _should_remove(name, prefixes):
            ids_to_delete.add(bid)
            removed_names.append(name)
        else:
            kept_names.append(name)

    # 2. 删 bookmarkStart
    removed_start = 0
    for bs in list(body.iter(f"{W}bookmarkStart")):
        bid = bs.get(f"{W}id")
        if bid in ids_to_delete:
            parent = bs.getparent()
            if parent is not None:
                parent.remove(bs)
                removed_start += 1

    # 3. 删配对 bookmarkEnd
    removed_end = 0
    for be in list(body.iter(f"{W}bookmarkEnd")):
        bid = be.get(f"{W}id")
        if bid in ids_to_delete:
            parent = be.getparent()
            if parent is not None:
                parent.remove(be)
                removed_end += 1

    by_prefix_removed: Counter = Counter()
    for n in removed_names:
        if n == "_GoBack":
            by_prefix_removed["_GoBack"] += 1
            continue
        matched = False
        for p in ("_Toc", "_Ref", "_Hlk", "_Hyperlink", "OLE_LINK"):
            if n.startswith(p):
                by_prefix_removed[p] += 1
                matched = True
                break
        if not matched:
            by_prefix_removed["user_defined"] += 1

    return {
        "bookmarks_removed": len(ids_to_delete),
        "starts": removed_start,
        "ends": removed_end,
        "removed_by_prefix": dict(by_prefix_removed),
        "kept_count": len(kept_names),
        "kept_sample": kept_names[:20],
    }


# ---------------- pipeline adapter ----------------
def apply(doc, args=None) -> dict:
    """pipeline 调用 — 改 doc 内存对象, 不写盘."""
    prefixes_attr = getattr(args, "bookmark_prefixes", None) if args else None
    if prefixes_attr is None:
        prefixes = list(DEFAULT_PREFIXES)
    elif isinstance(prefixes_attr, str):
        prefixes = parse_prefixes(prefixes_attr)
    else:
        prefixes = list(prefixes_attr) if prefixes_attr else list(DEFAULT_PREFIXES)
    return strip(doc, prefixes)


def main(argv: list[str] | None = None) -> int:
    ap = argparse.ArgumentParser(description="Strip bookmarks (default: _Toc/_Ref/_Hlk/_GoBack) from docx.")
    ap.add_argument("docx", type=Path)
    ap.add_argument("--prefixes", default=",".join(DEFAULT_PREFIXES),
                    help="Comma-separated bookmark name prefixes to strip (e.g. _Toc,_Ref,_Hlk). "
                         "'all' = strip every bookmark. _GoBack is always stripped. "
                         f"Default: {','.join(DEFAULT_PREFIXES)}")
    ap.add_argument("--dry-run", action="store_true", help="Report only, do not modify file")
    ap.add_argument("--no-backup", action="store_true", help="Skip writing .bak-N-<date>.docx")
    ap.add_argument("--report", type=Path, default=None, help="Write JSON report to this path")
    args = ap.parse_args(argv)

    if not args.docx.exists():
        print(f"[ERR] 文件不存在: {args.docx}", file=sys.stderr)
        return 1

    prefixes = parse_prefixes(args.prefixes)
    p = args.docx.resolve()

    if not args.dry_run:
        lsof_check(p)

    doc = Document(str(p))
    report = strip(doc, prefixes)
    report["docx"] = str(p)
    report["dry_run"] = args.dry_run
    report["prefixes"] = "all" if prefixes is None else prefixes

    print(json.dumps(report, ensure_ascii=False, indent=2))

    if args.report:
        args.report.parent.mkdir(parents=True, exist_ok=True)
        with args.report.open("w", encoding="utf-8") as fp:
            json.dump(report, fp, ensure_ascii=False, indent=2)

    if args.dry_run:
        print("[dry-run] no file written")
        return 0

    if not args.no_backup:
        bp = backup_path(p)
        print(f"[backup] {bp.name}")

    doc.save(str(p))
    print(f"[done] wrote {p}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
