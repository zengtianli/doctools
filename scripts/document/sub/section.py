"""section.py — group module: section read ops (1 subcommand)

Subcommand:
  section read <docx_file> <query>    ← section_read.py
  section read <docx_file> --list

  Read a named section from a DOCX by heading keyword (exact or fuzzy match).
  Use --list to enumerate all headings with paragraph indices — useful for
  determining start_idx / end_idx before running `md merge-into-docx`.

Distilled from panan-rigid-2026/scripts/read_section.py (B级通用, 2026-05-26).

Standalone CLI:
    python3 sub/section_read.py <docx_file> <query>
    python3 sub/section_read.py <docx_file> --list

Trigger scenarios:
  - "列出 docx 所有标题"
  - "读 docx 某章节内容"
  - 确认段落索引后配合 `md merge-into-docx` 安全回写
"""
from __future__ import annotations

import argparse

from ._dispatch import exec_script, _rest_argv


def _run(args) -> int:
    action = getattr(args, "section_action", None)
    if action != "read":
        print(f"[sub.section] unknown action: {action}; choices=['read']")
        return 2
    argv: list[str] = []
    if getattr(args, "docx_file", None):
        argv.append(str(args.docx_file))
    if getattr(args, "list_headings", False):
        argv.append("--list")
    elif getattr(args, "query", None):
        argv.append(str(args.query))
    extra = _rest_argv(args)
    argv.extend(extra)
    return exec_script("section_read", argv)


def register(subparsers) -> None:
    p = subparsers.add_parser(
        "section",
        help="section read/list ops on a DOCX",
    )
    sp = p.add_subparsers(dest="section_action", metavar="action")

    read_p = sp.add_parser("read", help="read a named section (or list all headings)")
    read_p.add_argument("docx_file", help="DOCX file path")
    read_p.add_argument(
        "query",
        nargs="?",
        help="heading keyword to match (fuzzy); omit with --list to enumerate",
    )
    read_p.add_argument(
        "--list", dest="list_headings", action="store_true",
        help="list all headings with paragraph indices",
    )
    read_p.add_argument("rest", nargs=argparse.REMAINDER)
    read_p.set_defaults(func=_run)
