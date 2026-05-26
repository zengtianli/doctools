"""captions.py — group module: caption ↔ table pairing ops (1 subcommand)

Subcommand:
  caption pair <decision.json>  ← pair_table_captions.py
                                  按 decision JSON 5 op 修「表名 ↔ tbl」配对:
                                    delete-caption / rename-caption /
                                    rename-orphan-tbl / pair-caption-to-tbl /
                                    renumber-all-tables
                                  decision 走 ~/Dev/tools/doctools/schemas/decision.schema.json v1.

Audit-first workflow:
    audit table-pairing → 人审 → decision.json → caption pair <decision.json>

Standalone CLI 仍可用:
    python3 sub/pair_table_captions.py <docx> --decision decision.json
"""
from __future__ import annotations

import argparse

from ._dispatch import exec_script, _rest_argv, get_or_add_group, get_or_add_subparsers


def _run(args) -> int:
    # caption_target unified across modules (caption.py/captions.py/styles.py)
    action = getattr(args, "caption_target", None) or getattr(args, "caption_action", None)
    if action != "pair":
        print(f"[sub.captions] unknown action: {action}; choices=['pair']")
        return 2
    argv = _rest_argv(args)
    decision = getattr(args, "decision", None)
    if decision:
        argv.extend(["--decision", str(decision)])
    return exec_script("pair_table_captions", argv)


def register(subparsers) -> None:
    """Register `caption pair` subcommand. Shares `caption` parent with caption.py/styles.py."""
    p = get_or_add_group(subparsers, "caption", "caption ops (number / pair / number-by-style)")
    sp = get_or_add_subparsers(p, dest="caption_target")
    if "pair" in (getattr(sp, "choices", {}) or {}):
        return
    spp = sp.add_parser("pair", help="apply decision.json to fix caption-table pairing",
                        add_help=False)
    spp.add_argument("docx_path", nargs="?", help="target docx path")
    spp.add_argument("--decision", help="decision JSON path (schemas/decision.schema.json v1)")
    spp.add_argument("--dry-run", action="store_true")
    spp.add_argument("--no-backup", action="store_true")
    spp.add_argument("--report", help="write JSON report to this path")
    spp.add_argument("rest", nargs=argparse.REMAINDER, help="extra args forwarded")
    spp.set_defaults(func=_run)
