"""chapter.py — group module: chapter / heading text ops (2 subcommands)

Subcommands:
  chapter convert-arabic     ← convert_chapter_format.py
                                rewrite H1 chapter prefixes from CJK to arabic:
                                  "第三章 Y" / "三、Y" → "3 Y"
  chapter delete-empty-h1    ← delete_empty_h1.py
                                delete H1 paragraphs whose text is empty / whitespace-only
                                (合稿后整段冗余的空章节)
"""
from __future__ import annotations

import argparse

from ._dispatch import exec_script, _rest_argv, get_or_add_group, get_or_add_subparsers


_TARGETS = {
    "convert-arabic":   "convert_chapter_format",
    "delete-empty-h1":  "delete_empty_h1",
}


def _run(args) -> int:
    target = getattr(args, "chapter_target", None)
    script = _TARGETS.get(target)
    if script is None:
        print(f"[sub.chapter] unknown target: {target}; choices={list(_TARGETS)}")
        return 2
    return exec_script(script, _rest_argv(args))


def register(subparsers) -> None:
    # Shared `chapter` group with blocks.py (W3, `chapter delete`).
    p = get_or_add_group(subparsers, "chapter",
                         "chapter ops (convert-arabic / delete-empty-h1 / delete)")
    sp = get_or_add_subparsers(p, dest="chapter_target")
    existing = getattr(sp, "choices", {}) or {}
    for t in _TARGETS:
        if t in existing:
            continue
        spp = sp.add_parser(t, help=f"chapter {t}", add_help=False)
        spp.add_argument("docx_path", nargs="?", help="target docx path")
        spp.add_argument("--dry-run", action="store_true")
        spp.add_argument("--no-backup", action="store_true")
        spp.add_argument("--report", help="write JSON report to this path")
        spp.add_argument("rest", nargs=argparse.REMAINDER, help="extra args forwarded to underlying script")
        spp.set_defaults(func=_run)
