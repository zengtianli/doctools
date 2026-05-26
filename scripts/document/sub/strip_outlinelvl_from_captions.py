#!/usr/bin/env python3
# distilled from qual-supply/scripts/strip_outlinelvl_from_captions.py (2026-05-25 W1)
# -*- coding: utf-8 -*-
"""
strip_outlinelvl_from_captions.py
==================================

单功能描述
----------
从 docx 中匹配 `^(表|图)\\s*\\d+\\.\\d+-\\d+` 的段移除 ``<w:outlineLvl>`` 子元素,
让这些 caption 段不再出现在 Word "视图 > 导航窗格"的标题大纲里。

触发场景
--------
- docx 整合后表名/图名段被错误地写入了 outlineLvl=6,导致 Word 导航大纲污染。
- 仅处理"表 X.Y-Z" / "图 X.Y-Z" 三段式编号的 caption 段。
- 不动 H1-H4 章节标题段(它们的 outlineLvl 由 Heading 1-4 样式合法控制)。

CLI
---
    python3 scripts/strip_outlinelvl_from_captions.py <docx_path> \\
        [--dry-run] [--no-backup] [--report <json>]

启发规则
--------
- 命中: ``re.match(r'^(表|图)\\s*\\d+\\.\\d+-\\d+', paragraph.text.strip())``
- 移除: ``paragraph._p / w:pPr / w:outlineLvl`` 子元素(用 lxml ``pPr.remove(ol)``)
- 段已无 outlineLvl → skip (no_outlinelvl_skip++)

不许做
------
- 不改段文本(编号 ``表X.Y-Z`` 用户已锁定保留)
- 不改段 style.name (交给 apply_caption_styles.py)
- 不裸解 zip 改 XML (走 python-docx + lxml)
- 不动 H1-H4 / Normal 正文段 / 19 张图

约束
----
- 单文件实现, 仅依赖 python-docx
- 自动备份 .bak-N-YYYY-MM-DD.docx (除非 --no-backup)
- 写前 lsof 自检 Word/WPS 占用, 占用立即退出
"""

from __future__ import annotations

import argparse
import json
import re
import shutil
import subprocess
import sys
from datetime import date
from pathlib import Path

from docx import Document
from docx.oxml.ns import qn

CAPTION_PATTERN = re.compile(r'^(表|图)\s*\d+\.\d+-\d+')


def find_next_backup(docx_path: Path) -> Path:
    today = date.today().isoformat()
    stem = docx_path.stem
    parent = docx_path.parent
    n = 1
    while True:
        cand = parent / f"{stem}.bak-{n}-{today}.docx"
        if not cand.exists():
            return cand
        n += 1


def lsof_check(docx_path: Path) -> str | None:
    """返回非空字符串 = 被占用, 返回 None = 可写"""
    try:
        out = subprocess.run(
            ["lsof", str(docx_path)],
            capture_output=True, text=True, timeout=5,
        )
        # lsof 退出码 1 = 没人开;退出码 0 = 有人开
        if out.returncode == 0 and out.stdout.strip():
            return out.stdout
    except (FileNotFoundError, subprocess.TimeoutExpired):
        pass
    return None


def scan_and_strip(doc: Document, apply: bool) -> dict:
    """扫所有段, 命中 caption pattern 的从 pPr 移除 outlineLvl"""
    captions_processed = 0
    outlinelvl_removed = 0
    no_outlinelvl_skip = 0
    details = []

    for idx, p in enumerate(doc.paragraphs):
        text = p.text.strip()
        if not CAPTION_PATTERN.match(text):
            continue
        captions_processed += 1
        pPr = p._p.find(qn('w:pPr'))
        ol = None if pPr is None else pPr.find(qn('w:outlineLvl'))
        if ol is None:
            no_outlinelvl_skip += 1
            details.append({
                "idx": idx,
                "text": text[:60],
                "outlineLvl_before": None,
                "action": "skip",
            })
            continue
        ol_val = ol.get(qn('w:val'))
        details.append({
            "idx": idx,
            "text": text[:60],
            "outlineLvl_before": ol_val,
            "action": "remove",
        })
        if apply:
            pPr.remove(ol)
        outlinelvl_removed += 1

    return {
        "captions_processed": captions_processed,
        "outlinelvl_removed": outlinelvl_removed,
        "no_outlinelvl_skip": no_outlinelvl_skip,
        "details": details,
    }


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    parser.add_argument("docx", type=Path)
    parser.add_argument("--dry-run", action="store_true")
    parser.add_argument("--no-backup", action="store_true")
    parser.add_argument("--report", type=Path, default=None)
    args = parser.parse_args()

    if not args.docx.exists():
        print(f"[ERR] 文件不存在: {args.docx}", file=sys.stderr)
        return 2

    if not args.dry_run:
        occ = lsof_check(args.docx)
        if occ:
            print(f"[ERR] 文件被占用 (Word/WPS 在开?), 立即停止:\n{occ}", file=sys.stderr)
            return 3

    doc = Document(str(args.docx))
    result = scan_and_strip(doc, apply=not args.dry_run)

    report = {
        "docx": str(args.docx.resolve()),
        "dry_run": args.dry_run,
        "backup": None,
        "wrote": False,
        **result,
    }

    print(f"[INFO] 扫描 {args.docx.name}")
    print(f"  captions_processed   = {result['captions_processed']}")
    print(f"  outlinelvl_removed   = {result['outlinelvl_removed']}")
    print(f"  no_outlinelvl_skip   = {result['no_outlinelvl_skip']}")

    # 列前几条
    for d in result["details"][:5]:
        print(f"  idx={d['idx']:4d} | ol={d['outlineLvl_before']!s:4s} | {d['action']:6s} | {d['text']}")

    if args.dry_run:
        print("[DRY-RUN] 不写文件")
    elif result["outlinelvl_removed"] == 0:
        print("[INFO] 无需移除, 不写文件")
    else:
        if not args.no_backup:
            bak = find_next_backup(args.docx)
            shutil.copy2(args.docx, bak)
            report["backup"] = str(bak)
            print(f"[INFO] 备份 -> {bak.name}")
        doc.save(str(args.docx))
        report["wrote"] = True
        print(f"[OK] 已移除 {result['outlinelvl_removed']} 个 outlineLvl, 写回 {args.docx.name}")

    if args.report:
        args.report.parent.mkdir(parents=True, exist_ok=True)
        args.report.write_text(json.dumps(report, ensure_ascii=False, indent=2), encoding="utf-8")
        print(f"[INFO] report -> {args.report}")

    return 0


# ---------------- pipeline adapter ----------------
def apply(doc, args=None) -> dict:
    dry = bool(getattr(args, "dry_run", False)) if args else False
    return scan_and_strip(doc, apply=not dry)


if __name__ == "__main__":
    sys.exit(main())
