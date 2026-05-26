#!/usr/bin/env python3
# distilled from qual-supply/scripts/audit_bookmarks.py (2026-05-25 W1)
# -*- coding: utf-8 -*-
"""
audit_bookmarks.py
==================

单功能描述
----------
只读 audit docx 内全部 `<w:bookmarkStart>` / `<w:bookmarkEnd>`,产出:
- 总数 / 按 prefix 分类计数 (`_Toc`/`_Ref`/`_Hlk`/`_GoBack`/`user_defined`)
- 每个 bookmark 的 id/name/start_para_idx/end_para_idx
- orphan_starts (start 无对应 end) / orphan_ends (end 无对应 start)
- para_idx_for_each (id → start/end para idx)

不改文件; 给 strip_bookmarks.py 做 decision 输入,或验收前看交叉引用 anchor 干净度。

触发场景
--------
- 合稿后 bookmark ID 冲突 / `_Toc`/`_Ref` 失锚根因分析
- 验收前盘点用户手动命名 bookmark vs Word 自动生成 bookmark
- pipeline 前后对比 strip 效果

CLI
---
    python3 scripts/audit_bookmarks.py <docx> [--report <json>]

不许做
------
- 改 doc 内容 (read-only audit)
- 改 fldChar/instrText (那是 audit_word_fields.py 的活)
- commit / push
- 用 sed/awk
"""
from __future__ import annotations

import argparse
import json
import sys
from collections import Counter
from pathlib import Path
from typing import Any

try:
    from docx import Document
except ImportError:
    print("[ERR] 缺 python-docx: pip install python-docx", file=sys.stderr)
    sys.exit(2)

NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
W = f"{{{NS}}}"

# Word 内部自动 bookmark 前缀 (启发集合)
KNOWN_INTERNAL_PREFIXES = ("_Toc", "_Ref", "_Hlk", "_GoBack", "_Hyperlink", "OLE_LINK")


def _classify_name(name: str) -> str:
    """返回 prefix 类别 (用于 bookmark_by_prefix 计数)."""
    if not name:
        return "EMPTY"
    if name == "_GoBack":
        return "_GoBack"
    for p in ("_Toc", "_Ref", "_Hlk", "_Hyperlink", "OLE_LINK"):
        if name.startswith(p):
            return p
    return "user_defined"


def _para_idx_map(body) -> dict:
    """每个 <w:p> 元素 → 物理顺序 idx."""
    return {p: i for i, p in enumerate(body.iter(f"{W}p"))}


def _ancestor_para_idx(elem, para_idx_map) -> int:
    cur = elem
    while cur is not None:
        if cur.tag == f"{W}p":
            return para_idx_map.get(cur, -1)
        cur = cur.getparent()
    return -1


def audit(doc) -> dict:
    """扫 doc body 全部 bookmarkStart/bookmarkEnd, 返回 report dict."""
    body = doc.element.body
    para_idx_map = _para_idx_map(body)

    starts: dict[str, dict] = {}  # id → {name, para_idx}
    ends: dict[str, int] = {}     # id → end para_idx

    for bs in body.iter(f"{W}bookmarkStart"):
        bid = bs.get(f"{W}id")
        name = bs.get(f"{W}name") or ""
        if bid is None:
            continue
        starts[bid] = {
            "name": name,
            "start_para_idx": _ancestor_para_idx(bs, para_idx_map),
        }
    for be in body.iter(f"{W}bookmarkEnd"):
        bid = be.get(f"{W}id")
        if bid is None:
            continue
        ends[bid] = _ancestor_para_idx(be, para_idx_map)

    bookmarks: list[dict] = []
    orphan_starts: list[dict] = []
    orphan_ends: list[str] = []
    by_prefix: Counter = Counter()
    para_idx_for_each: dict[str, dict] = {}

    for bid, info in starts.items():
        name = info["name"]
        rec = {
            "id": bid,
            "name": name,
            "start_para_idx": info["start_para_idx"],
            "end_para_idx": ends.get(bid, -1),
            "prefix": _classify_name(name),
        }
        bookmarks.append(rec)
        by_prefix[rec["prefix"]] += 1
        para_idx_for_each[bid] = {
            "name": name,
            "start": info["start_para_idx"],
            "end": ends.get(bid, -1),
        }
        if bid not in ends:
            orphan_starts.append({"id": bid, "name": name,
                                  "start_para_idx": info["start_para_idx"]})
    for bid in ends:
        if bid not in starts:
            orphan_ends.append(bid)

    return {
        "bookmark_count": len(bookmarks),
        "bookmark_by_prefix": dict(by_prefix),
        "bookmarks": bookmarks,
        "orphan_starts": orphan_starts,
        "orphan_ends": orphan_ends,
        "para_idx_for_each": para_idx_for_each,
    }


# ---------------- pipeline adapter ----------------
def apply(doc, args=None) -> dict:
    """pipeline 调用 — 只读 audit, 不改 doc."""
    return audit(doc)


def main(argv: list[str] | None = None) -> int:
    ap = argparse.ArgumentParser(description="Audit bookmarks in a docx (read-only).")
    ap.add_argument("docx", type=Path)
    ap.add_argument("--report", type=Path, default=None,
                    help="Write full JSON report to this path")
    args = ap.parse_args(argv)

    if not args.docx.exists():
        print(f"[ERR] 文件不存在: {args.docx}", file=sys.stderr)
        return 1

    doc = Document(str(args.docx))
    report = audit(doc)
    report["docx"] = str(args.docx.resolve())

    summary = {
        "docx": report["docx"],
        "bookmark_count": report["bookmark_count"],
        "bookmark_by_prefix": report["bookmark_by_prefix"],
        "orphan_starts_count": len(report["orphan_starts"]),
        "orphan_ends_count": len(report["orphan_ends"]),
    }
    print(json.dumps(summary, ensure_ascii=False, indent=2))

    if args.report:
        args.report.parent.mkdir(parents=True, exist_ok=True)
        with args.report.open("w", encoding="utf-8") as fp:
            json.dump(report, fp, ensure_ascii=False, indent=2)
    return 0


if __name__ == "__main__":
    sys.exit(main())
