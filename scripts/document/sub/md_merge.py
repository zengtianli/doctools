"""md_merge.py — group module: merge MD content into DOCX section (1 subcommand)

Subcommand:
  md merge <md_file> <docx_file> <start_idx> <end_idx> [output_file]
                ← md_merge_impl.py (distilled from panan-rigid scripts/merge_md_to_docx.py)

  Safely replaces the content between two paragraph indices in a DOCX with
  paragraphs parsed from a Markdown file.  Tables and non-paragraph XML
  elements inside the replaced range are preserved and reinserted by anchor
  text matching.

Standalone CLI:
    python3 sub/md_merge_impl.py <md_file> <docx_file> <start_idx> <end_idx> [output]

Trigger scenarios:
  - "把 MD 内容合入 docx 第 N 章"
  - "用新写的 md 替换 docx 某章节"
  - 编写完 Markdown 草稿后需要安全回写到已有 Word 交付物
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
    extra = _rest_argv(args)
    argv.extend(extra)
    return exec_script("md_merge_impl", argv)


def register(subparsers) -> None:
    """Register `md merge` subcommand onto shared `md` group."""
    try:
        from ._dispatch import get_or_add_group, get_or_add_subparsers
    except ImportError:
        # Fallback: add directly as top-level
        p = subparsers.add_parser("md-merge", help="merge MD content into DOCX section")
        p.add_argument("md_file", help="source Markdown file")
        p.add_argument("docx_file", help="target DOCX file")
        p.add_argument("start_idx", type=int, help="start paragraph index (heading, kept)")
        p.add_argument("end_idx", type=int, help="end paragraph index (exclusive)")
        p.add_argument("output_file", nargs="?", help="output path (default: <docx>-merged.docx)")
        p.add_argument("rest", nargs=argparse.REMAINDER)
        p.set_defaults(func=_run)
        return

    grp = get_or_add_group(subparsers, "md",
                           "Markdown ops (format / merge / split / strip / to-docx / to-html / frontmatter / merge-into-docx)")
    sp = get_or_add_subparsers(grp, dest="md_target")
    existing = getattr(sp, "choices", {}) or {}
    if "merge-into-docx" not in existing:
        spp = sp.add_parser("merge-into-docx",
                            help="replace a DOCX section with MD content (table-safe)")
        spp.add_argument("md_file", help="source Markdown file")
        spp.add_argument("docx_file", help="target DOCX file")
        spp.add_argument("start_idx", type=int,
                         help="start paragraph index (heading paragraph, preserved)")
        spp.add_argument("end_idx", type=int,
                         help="end paragraph index (exclusive; next section heading)")
        spp.add_argument("output_file", nargs="?",
                         help="output path (default: <docx>-merged.docx)")
        spp.add_argument("rest", nargs=argparse.REMAINDER,
                         help="extra args forwarded to underlying script")
        spp.set_defaults(func=_run)
