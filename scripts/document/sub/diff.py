"""diff.py — group module: docx diff ops (2 subcommands)

Subcommands:
  diff seq <src.docx> <dst.docx> --out <report.md> [--noise-len N]
    ← seqdiff.py
    逐段 sequence diff: src vs dst，精确/高相似/中相似/新增 4 级分类
    输出 MD 报告含总体统计 + 跨章节迁移 + 雷同风险分组

  diff image <src.docx> <dst.docx> --out <report.md>
    ← image_dedup.py
    SHA256 图片去重: 识别 dst 中与 src 二进制相同的图片，定位章节

Standalone CLI 仍可用:
    python3 sub/seqdiff.py --src OLD.docx --dst NEW.docx --out report.md
    python3 sub/image_dedup.py --src OLD.docx --dst NEW.docx --out report.md

Distilled from ~/Work/shared/bids/.claude/skills/bid-diff-and-revise/scripts/ (2026-05-26)
"""
from __future__ import annotations

import argparse

from ._dispatch import exec_script


def _run_seq(args) -> int:
    argv: list[str] = []
    if getattr(args, "src", None):
        argv.extend(["--src", str(args.src)])
    if getattr(args, "dst", None):
        argv.extend(["--dst", str(args.dst)])
    if getattr(args, "out", None):
        argv.extend(["--out", str(args.out)])
    if getattr(args, "noise_len", None) is not None:
        argv.extend(["--noise-len", str(args.noise_len)])
    argv.extend(str(x) for x in (getattr(args, "rest", None) or []))
    return exec_script("seqdiff", argv)


def _run_image(args) -> int:
    argv: list[str] = []
    if getattr(args, "src", None):
        argv.extend(["--src", str(args.src)])
    if getattr(args, "dst", None):
        argv.extend(["--dst", str(args.dst)])
    if getattr(args, "out", None):
        argv.extend(["--out", str(args.out)])
    argv.extend(str(x) for x in (getattr(args, "rest", None) or []))
    return exec_script("image_dedup", argv)


def register(subparsers) -> None:
    p = subparsers.add_parser(
        "seqdiff",
        help="docx seqdiff ops: seq (逐段对照) / image (图片去重) — distilled from bid-diff-and-revise",
    )
    sp = p.add_subparsers(dest="diff_action", metavar="<action>", required=True)

    # seq
    sp_seq = sp.add_parser("seq", help="逐段 sequence diff: src vs dst → MD 报告", add_help=False)
    sp_seq.add_argument("--src", required=False, help="源 docx（旧版）")
    sp_seq.add_argument("--dst", required=False, help="目标 docx（新版）")
    sp_seq.add_argument("--out", required=False, help="输出 MD 路径")
    sp_seq.add_argument("--noise-len", type=int, default=8, dest="noise_len",
                        help="短于 N 字符视为噪声（默认 8）")
    sp_seq.add_argument("rest", nargs=argparse.REMAINDER)
    sp_seq.set_defaults(func=_run_seq)

    # image
    sp_img = sp.add_parser("image", help="SHA256 图片去重: src vs dst → 重复清单 MD", add_help=False)
    sp_img.add_argument("--src", required=False, help="源 docx")
    sp_img.add_argument("--dst", required=False, help="目标 docx")
    sp_img.add_argument("--out", required=False, help="输出 MD 路径")
    sp_img.add_argument("rest", nargs=argparse.REMAINDER)
    sp_img.set_defaults(func=_run_image)
