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
import re
import sys
from pathlib import Path

# ── surgical 收口：python-docx 存盘只重写点名的部件（炸开面 60→1）─────────────
import sys as _sys
from pathlib import Path as _Path
_sys.path.append(str(_Path(__file__).resolve().parents[3] / "lib"))
import docx_safe_save  # noqa: E402,F401  详见 lib/docx_safe_save.py

# sub/ 自身进 sys.path —— docx_cli 的 _dispatch 用 spec_from_file_location 加载,
# 不带脚本目录, 裸 import _cli_common 会 ImportError (append 不是 insert(0))
_sys.path.append(str(_Path(__file__).resolve().parent))
import _cli_common as _cc  # noqa: E402  家族 main() 样板 SSOT

from docx import Document
from docx.oxml.ns import qn

CAPTION_PATTERN = re.compile(r'^(表|图)\s*\d+\.\d+-\d+')


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
    _cc.add_write_flags(parser)
    args = parser.parse_args()

    if not args.docx.exists():
        print(f"[ERR] 文件不存在: {args.docx}", file=sys.stderr)
        return 2

    if not args.dry_run:
        occ = _cc.lsof_check(args.docx)
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
            bak = _cc.make_backup(args.docx)
            report["backup"] = str(bak)
            print(f"[INFO] 备份 -> {bak.name}")
        doc.save(str(args.docx))
        report["wrote"] = True
        print(f"[OK] 已移除 {result['outlinelvl_removed']} 个 outlineLvl, 写回 {args.docx.name}")

    _cc.write_report(report, args.report, announce="[INFO] report -> {path}")

    return 0


# ---------------- pipeline adapter ----------------
def apply(doc, args=None) -> dict:
    dry = bool(getattr(args, "dry_run", False)) if args else False
    return scan_and_strip(doc, apply=not dry)


if __name__ == "__main__":
    sys.exit(main())
