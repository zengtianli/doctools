"""images.py — group module: image ops (2 subcommands)

Subcommands:
  image relink <source.docx> [--apply-patch <patch.json>] [--patch <out.json>]
                                ← relink_images_from_source.py
                                  从源 docx 提取 word/media + rels 重嵌入 target docx
                                  的 dangling drawing/pict. 启发配对 (caption / 段位置 /
                                  邻接 blip), 支持 detect-only 出 patch + 独立 apply
                                  (W3/W4 并发不冲突).
                                  patch 走 ~/Dev/tools/doctools/schemas/patch.schema.json v1.

  image extract <docx> --out-dir <dir>
                                ← image_extract.py (W · 2026-05-28)
                                  抽 docx 内所有 <a:blip>/<v:imagedata> 引用的 word/media/
                                  二进制成文件, 命名 = 邻近 caption 文字 sanitize
                                  ("图x-y …"), fallback image-NN. 纯读不改 docx.

Standalone CLI 仍可用:
    # relink: 一步式 detect + apply
    python3 sub/relink_images_from_source.py <target.docx> --source <source.docx>
    # relink: 两步 detect-only -> apply patch
    python3 sub/relink_images_from_source.py <target.docx> --source <source.docx> \
        --report patch.json --dry-run
    python3 sub/relink_images_from_source.py <target.docx> --apply-patch patch.json
    # extract
    python3 sub/image_extract.py <docx> --out-dir <dir>
"""
from __future__ import annotations

import argparse

from ._dispatch import exec_script


def _run(args) -> int:
    action = getattr(args, "image_action", None)
    if action == "relink":
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
    if action == "extract":
        argv = []
        if getattr(args, "docx", None):
            argv.append(str(args.docx))
        if getattr(args, "out_dir", None):
            argv.extend(["--out-dir", str(args.out_dir)])
        if getattr(args, "quiet", False):
            argv.append("--quiet")
        extra = getattr(args, "rest", None) or []
        argv.extend(str(x) for x in extra)
        return exec_script("image_extract", argv)
    print(f"[sub.image] unknown action: {action}; choices=['relink', 'extract']")
    return 2


def register(subparsers) -> None:
    p = subparsers.add_parser(
        "image",
        help="image ops (relink media from source docx / extract images by caption)",
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

    spe = sp.add_parser("extract", help="extract images by neighboring caption text",
                        add_help=False)
    spe.add_argument("docx", nargs="?", help="source docx (read-only)")
    spe.add_argument("--out-dir", required=False, help="output directory")
    spe.add_argument("--quiet", action="store_true")
    spe.add_argument("rest", nargs=argparse.REMAINDER, help="extra args forwarded")
    spe.set_defaults(func=_run)
