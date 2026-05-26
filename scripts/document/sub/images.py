"""images.py — group module: image relink ops (1 subcommand)

Subcommand:
  image relink <source.docx> [--apply-patch <patch.json>] [--patch <out.json>]
                                ← relink_images_from_source.py
                                  从源 docx 提取 word/media + rels 重嵌入 target docx
                                  的 dangling drawing/pict. 启发配对 (caption / 段位置 /
                                  邻接 blip), 支持 detect-only 出 patch + 独立 apply
                                  (W3/W4 并发不冲突).
                                  patch 走 ~/Dev/tools/doctools/schemas/patch.schema.json v1.

Standalone CLI 仍可用:
    # 一步式 detect + apply
    python3 sub/relink_images_from_source.py <target.docx> --source <source.docx>
    # 两步: detect-only -> apply patch
    python3 sub/relink_images_from_source.py <target.docx> --source <source.docx> \
        --report patch.json --dry-run
    python3 sub/relink_images_from_source.py <target.docx> --apply-patch patch.json
"""
from __future__ import annotations

import argparse

from ._dispatch import exec_script


def _run(args) -> int:
    action = getattr(args, "image_action", None)
    if action != "relink":
        print(f"[sub.image] unknown action: {action}; choices=['relink']")
        return 2
    argv: list[str] = []
    if getattr(args, "target_docx", None):
        argv.append(str(args.target_docx))
    if getattr(args, "source", None):
        argv.extend(["--source", str(args.source)])
    if getattr(args, "apply_patch", None):
        argv.extend(["--apply-patch", str(args.apply_patch)])
    if getattr(args, "report", None):
        argv.extend(["--report", str(args.report)])
    if getattr(args, "dry_run", False):
        argv.append("--dry-run")
    if getattr(args, "no_backup", False):
        argv.append("--no-backup")
    extra = getattr(args, "rest", None) or []
    argv.extend(str(x) for x in extra)
    return exec_script("relink_images_from_source", argv)


def register(subparsers) -> None:
    p = subparsers.add_parser(
        "image",
        help="image relink ops (re-embed media from source docx)",
    )
    sp = p.add_subparsers(dest="image_action", metavar="<action>", required=True)
    spp = sp.add_parser("relink", help="relink/re-embed images from a source docx",
                        add_help=False)
    spp.add_argument("target_docx", nargs="?", help="target docx (with dangling drawings)")
    spp.add_argument("--source", help="source docx to lift media from")
    spp.add_argument("--apply-patch", help="apply existing patch JSON (skip detect)")
    spp.add_argument("--report", help="write patch JSON to this path (use with --dry-run for detect-only)")
    spp.add_argument("--dry-run", action="store_true")
    spp.add_argument("--no-backup", action="store_true")
    spp.add_argument("rest", nargs=argparse.REMAINDER, help="extra args forwarded")
    spp.set_defaults(func=_run)
