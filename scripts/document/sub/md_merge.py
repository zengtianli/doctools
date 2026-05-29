"""md_merge.py — group module: merge MD content into DOCX section (1 subcommand)

Subcommand (top-level):
  md-merge <md_file> <docx_file> <start_idx> <end_idx> [output_file]
                ← md_merge_impl.py (distilled from panan-rigid scripts/merge_md_to_docx.py)

  Safely replaces the content between two paragraph indices in a DOCX with
  paragraphs parsed from a Markdown file.  Tables and non-paragraph XML
  elements inside the replaced range are preserved and reinserted by anchor
  text matching.

  Note: registered as top-level `md-merge` (not under `md` group) because
  the `md` group is a legacy subcommand dispatching to md_tools.py.

Standalone CLI:
    python3 sub/md_merge_impl.py <md_file> <docx_file> <start_idx> <end_idx> [output]

Trigger scenarios:
  - "把 MD 内容合入 docx 第 N 章"
  - "用新写的 md 替换 docx 某章节"
  - 编写完 Markdown 草稿后需要安全回写到已有 Word 交付物
  - 配合 `section read --list` 先确认段落索引再合入
"""
from __future__ import annotations

import argparse

from ._dispatch import exec_script, _rest_argv


def _run(args) -> int:
    argv: list[str] = []
    if getattr(args, "md_file", None):
        argv.append(str(args.md_file))
    if getattr(args, "docx_file", None):
        argv.append(str(args.docx_file))
    if getattr(args, "start_idx", None) is not None:
        argv.append(str(args.start_idx))
    if getattr(args, "end_idx", None) is not None:
        argv.append(str(args.end_idx))
    if getattr(args, "output_file", None):
        argv.append(str(args.output_file))
    if getattr(args, "start_anchor", None):
        argv += ["--start-anchor", str(args.start_anchor)]
    if getattr(args, "end_anchor", None):
        argv += ["--end-anchor", str(args.end_anchor)]
    if getattr(args, "in_place", False):
        argv.append("--in-place")
    if getattr(args, "no_backup", False):
        argv.append("--no-backup")
    extra = _rest_argv(args)
    argv.extend(extra)
    return exec_script("md_merge_impl", argv)


def register(subparsers) -> None:
    """Register `md-merge` as a top-level subcommand."""
    p = subparsers.add_parser(
        "md-merge",
        help="replace a DOCX section with MD content (table-safe; --in-place + .bak; --start/end-anchor)",
    )
    p.add_argument("md_file", help="source Markdown file")
    p.add_argument("docx_file", help="target DOCX file")
    p.add_argument("start_idx", nargs="?", type=int,
                   help="start paragraph index (heading; preserved + title updated). 可用 --start-anchor 替代")
    p.add_argument("end_idx", nargs="?", type=int,
                   help="end paragraph index (exclusive; next section heading). 可用 --end-anchor 替代")
    p.add_argument("output_file", nargs="?",
                   help="output path (default: <docx>-merged.docx; --in-place 时忽略)")
    p.add_argument("--start-anchor", help="按标题文本定位 start_idx (省掉手查索引)")
    p.add_argument("--end-anchor", help="按标题文本定位 end_idx (下一节标题)")
    p.add_argument("--in-place", action="store_true", help="原地改 + 自动 .bak-时间戳 (Work §1.5)")
    p.add_argument("--no-backup", action="store_true", help="配合 --in-place 跳过备份")
    p.add_argument("rest", nargs=argparse.REMAINDER,
                   help="extra args forwarded to underlying script")
    p.set_defaults(func=_run)
