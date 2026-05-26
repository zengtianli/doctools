"""legacy.py — group module: deprecated catch-all (1 subcommand)

Subcommand:
  legacy fix-heading-disorder  ← fix_heading_disorder.py (DEPRECATED)
                                   6 类杂糅 (false_promotion / false_demotion /
                                   numbering_backward / numbering_skip /
                                   level_mismatch / duplicate_adjacent) 检测+修复.
                                   违反「一脚本一功能」(qual-supply CLAUDE.md 2026-05-25
                                   用户钦定), 已停止扩展.

新项目应使用拆分后的单功能脚本:
    docx outline normalize-arabic  ← normalize_outline_to_arabic.py
    docx outline promote-h1        ← promote_misclassified_h1.py
    docx outline demote-h2         ← demote_h2_with_h3_format.py
    docx headings renumber         ← renumber_headings.py
    docx blocks reorder            ← reorder_heading_blocks.py

Standalone CLI 仍可用 (老 docx 复用兼容):
    python3 sub/fix_heading_disorder.py <docx> [--dry-run] [--report x.json]
"""
from __future__ import annotations

import argparse

from ._dispatch import exec_script, _rest_argv


_TARGETS = {
    "fix-heading-disorder": "fix_heading_disorder",
}


def _run(args) -> int:
    target = getattr(args, "legacy_target", None)
    script = _TARGETS.get(target)
    if script is None:
        print(f"[sub.legacy] unknown target: {target}; choices={list(_TARGETS)}")
        return 2
    print(f"[sub.legacy] WARNING: '{target}' is DEPRECATED; see sub/legacy.py docstring "
          f"for the recommended single-purpose replacement scripts.",
          flush=True)
    return exec_script(script, _rest_argv(args))


def register(subparsers) -> None:
    p = subparsers.add_parser(
        "legacy",
        help="DEPRECATED catch-all scripts (fix-heading-disorder)",
    )
    sp = p.add_subparsers(dest="legacy_target", metavar="<target>", required=True)
    for t in _TARGETS:
        spp = sp.add_parser(t, help=f"[DEPRECATED] {t}", add_help=False)
        spp.add_argument("docx_path", nargs="?", help="target docx path")
        spp.add_argument("--dry-run", action="store_true")
        spp.add_argument("--no-backup", action="store_true")
        spp.add_argument("--report", help="write JSON report to this path")
        spp.add_argument("rest", nargs=argparse.REMAINDER, help="extra args forwarded")
        spp.set_defaults(func=_run)
