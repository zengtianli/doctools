"""audit.py — group module: audit-only docx checks (6 subcommands)

Subcommands:
  audit headings        ← audit_heading_numbers.py
  audit fields          ← audit_word_fields.py
  audit captions        ← audit_caption_outline.py
  audit images          ← audit_images.py
  audit table-pairing   ← audit_table_pairing.py
  audit bookmarks       ← audit_bookmarks.py

Each is read-only (no docx mutation). Forwards argv to the standalone
sub/<script>.py main(). Standalone CLI still works:
    python3 sub/audit_heading_numbers.py <docx> [--report x.json]
"""
from __future__ import annotations

import argparse

from ._dispatch import exec_script, _rest_argv


_TARGETS = {
    "headings":      "audit_heading_numbers",
    "fields":        "audit_word_fields",
    "captions":      "audit_caption_outline",
    "images":        "audit_images",
    "table-pairing": "audit_table_pairing",
    "bookmarks":     "audit_bookmarks",
}


def _run(args) -> int:
    target = getattr(args, "audit_target", None)
    script = _TARGETS.get(target)
    if script is None:
        print(f"[sub.audit] unknown target: {target}; choices={list(_TARGETS)}")
        return 2
    return exec_script(script, _rest_argv(args))


def register(subparsers) -> None:
    """Register `audit <target>` subcommands onto a parent subparsers object."""
    p = subparsers.add_parser(
        "audit",
        help="audit-only docx checks (headings/fields/captions/images/table-pairing/bookmarks)",
    )
    sp = p.add_subparsers(dest="audit_target", metavar="<target>", required=True)
    for t in _TARGETS:
        spp = sp.add_parser(t, help=f"audit {t} (read-only)", add_help=False)
        spp.add_argument("docx_path", nargs="?", help="target docx path")
        spp.add_argument("--dry-run", action="store_true")
        spp.add_argument("--no-backup", action="store_true")
        spp.add_argument("--report", help="write JSON report to this path")
        spp.add_argument("rest", nargs=argparse.REMAINDER, help="extra args forwarded to underlying script")
        spp.set_defaults(func=_run)
