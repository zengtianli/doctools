"""blocks.py — group module: paragraph-block structural ops (3 subcommands)

Subcommands:
  blocks reorder              ← reorder_heading_blocks.py
                                按段块兄弟排序 (heading + 下属正文段) + 删同号重复块.
                                不跨 H1, 不删 H1 重复.
  blocks relocate <plan.json> ← relocate_orphan_blocks.py
                                外部 plan JSON 驱动机械搬段 (孤儿块跨章移位).
                                plan 走 ~/Dev/tools/doctools/schemas/plan.schema.json v1.
  chapter delete <h1>         ← delete_chapter.py
                                按 H1 编号或 text 前缀删整章 (含表/段, 保留 sectPr).

Standalone CLI 仍可用:
    python3 sub/reorder_heading_blocks.py <docx>
    python3 sub/relocate_orphan_blocks.py <docx> --plan plan.json
    python3 sub/delete_chapter.py <docx> --h1 3
"""
from __future__ import annotations

import argparse

from ._dispatch import exec_script, _rest_argv, get_or_add_group, get_or_add_subparsers


# ─── blocks reorder / relocate ───────────────────────────────────────────
_BLOCK_TARGETS = {
    "reorder":  "reorder_heading_blocks",
    "relocate": "relocate_orphan_blocks",
}


def _run_blocks(args) -> int:
    target = getattr(args, "block_target", None)
    script = _BLOCK_TARGETS.get(target)
    if script is None:
        print(f"[sub.blocks] unknown target: {target}; choices={list(_BLOCK_TARGETS)}")
        return 2
    argv = _rest_argv(args)
    # relocate 需要 --plan 透传, 已在 rest 里
    plan = getattr(args, "plan", None)
    if plan:
        argv.extend(["--plan", str(plan)])
    return exec_script(script, argv)


# ─── chapter delete ──────────────────────────────────────────────────────
def _run_chapter(args) -> int:
    sub = getattr(args, "chapter_target", None) or getattr(args, "chapter_action", None)
    if sub != "delete":
        print(f"[sub.chapter] unknown action: {sub}; choices=['delete']")
        return 2
    argv: list[str] = []
    if getattr(args, "docx_path", None):
        argv.append(str(args.docx_path))
    # delete_chapter.py 接受 --h1 / --prefix
    if getattr(args, "h1", None):
        argv.extend(["--h1", str(args.h1)])
    if getattr(args, "prefix", None):
        argv.extend(["--prefix", str(args.prefix)])
    if getattr(args, "dry_run", False):
        argv.append("--dry-run")
    if getattr(args, "no_backup", False):
        argv.append("--no-backup")
    extra = getattr(args, "rest", None) or []
    argv.extend(str(x) for x in extra)
    return exec_script("delete_chapter", argv)


def register(subparsers) -> None:
    """Register `blocks` + `chapter delete` subcommands. Shares `chapter` group."""
    # blocks <reorder|relocate>
    pb = get_or_add_group(subparsers, "blocks", "paragraph-block structural ops (reorder/relocate)")
    spb = get_or_add_subparsers(pb, dest="block_target")
    existing_b = getattr(spb, "choices", {}) or {}
    for t in _BLOCK_TARGETS:
        if t in existing_b:
            continue
        spp = spb.add_parser(t, help=f"blocks {t}", add_help=False)
        spp.add_argument("docx_path", nargs="?", help="target docx path")
        spp.add_argument("--dry-run", action="store_true")
        spp.add_argument("--no-backup", action="store_true")
        spp.add_argument("--report", help="write JSON report to this path")
        if t == "relocate":
            spp.add_argument("--plan", help="plan JSON path (schemas/plan.schema.json v1)")
        spp.add_argument("rest", nargs=argparse.REMAINDER, help="extra args forwarded")
        spp.set_defaults(func=_run_blocks)

    # chapter delete — shared `chapter` group with chapter.py (W1)
    pc = get_or_add_group(subparsers, "chapter",
                          "chapter ops (convert-arabic / delete-empty-h1 / delete)")
    spc = get_or_add_subparsers(pc, dest="chapter_target")
    existing_c = getattr(spc, "choices", {}) or {}
    if "delete" not in existing_c:
        spcd = spc.add_parser("delete", help="delete an entire H1 chapter", add_help=False)
        spcd.add_argument("docx_path", nargs="?", help="target docx path")
        spcd.add_argument("--h1", help="H1 chapter number (e.g. 3)")
        spcd.add_argument("--prefix", help="H1 text prefix to match")
        spcd.add_argument("--dry-run", action="store_true")
        spcd.add_argument("--no-backup", action="store_true")
        spcd.add_argument("rest", nargs=argparse.REMAINDER, help="extra args forwarded")
        spcd.set_defaults(func=_run_chapter)
