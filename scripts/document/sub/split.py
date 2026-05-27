"""split.py — group module: split docx by structural unit (1 subcommand)

Subcommand:
  split by-h1 --docx <path> --out-dir <dir> [opts]   ← split_by_h1.py
        按 Heading 1 切分 docx → N 个独立 docx, 保留 styles/numbering/media。

Standalone CLI:
    python3 sub/split_by_h1.py --docx X --out-dir Y [--include-frontmatter] [--dry-run]

Trigger scenarios:
  - 大型 docx 报告按章节拆分(每章独立 review / 分发 / 并行编辑)
  - 模板分发(原始模板 → N 个空章节模板)
"""
from __future__ import annotations

import argparse

from ._dispatch import exec_script


def _run(args) -> int:
    action = getattr(args, "split_action", None)
    if action != "by-h1":
        print(f"[sub.split] unknown action: {action}; choices=['by-h1']")
        return 2
    argv: list[str] = []
    if getattr(args, "docx", None):
        argv.extend(["--docx", str(args.docx)])
    if getattr(args, "out_dir", None):
        argv.extend(["--out-dir", str(args.out_dir)])
    if getattr(args, "name_pattern", None):
        argv.extend(["--name-pattern", str(args.name_pattern)])
    if getattr(args, "include_frontmatter", False):
        argv.append("--include-frontmatter")
    if getattr(args, "dry_run", False):
        argv.append("--dry-run")
    if getattr(args, "allow_no_h1", False):
        argv.append("--allow-no-h1")
    extra = getattr(args, "rest", None) or []
    argv.extend(str(x) for x in extra)
    return exec_script("split_by_h1", argv)


def register(subparsers) -> None:
    p = subparsers.add_parser(
        "split",
        help="split docx by structural unit (by-h1)",
    )
    sp = p.add_subparsers(dest="split_action", metavar="action")

    by_h1 = sp.add_parser("by-h1", help="split docx by Heading 1 into N files")
    by_h1.add_argument("--docx", required=True, help="input docx path")
    by_h1.add_argument("--out-dir", required=True, help="output directory")
    by_h1.add_argument("--name-pattern", default=None,
                       help="filename pattern (default '{idx:02d}-{title}.docx')")
    by_h1.add_argument("--include-frontmatter", action="store_true",
                       help="emit content before first H1 as 00-frontmatter.docx")
    by_h1.add_argument("--dry-run", action="store_true",
                       help="print plan only, don't write files")
    by_h1.add_argument("--allow-no-h1", action="store_true",
                       help="suppress unhealthy-docx fail-fast (rare; default = "
                            "FAIL on 0 H1 and tell user to run /docx health first)")
    by_h1.add_argument("rest", nargs=argparse.REMAINDER)
    by_h1.set_defaults(func=_run)
