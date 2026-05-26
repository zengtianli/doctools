"""compare.py — group module: docx compare-vs-ref ops (1 subcommand)

Subcommand:
  compare ref --drafts-dir <dir> --ref <ref.docx> --out <report.md> [--glob *.md]
    ← compare_vs_ref.py
    对比多份改动 MD 里的"改为"段 vs 参考 docx，找雷同风险
    HIGH ratio≥0.85 → 必须重写；MID 0.6≤ratio<0.85 → 建议调整

Standalone CLI 仍可用:
    python3 sub/compare_vs_ref.py --drafts-dir 成果/md --ref 主标.docx --out report.md

Distilled from ~/Work/shared/bids/.claude/skills/bid-diff-and-revise/scripts/ (2026-05-26)
"""
from __future__ import annotations

import argparse

from ._dispatch import exec_script


def _run_ref(args) -> int:
    argv: list[str] = []
    if getattr(args, "drafts_dir", None):
        argv.extend(["--drafts-dir", str(args.drafts_dir)])
    if getattr(args, "ref", None):
        argv.extend(["--ref", str(args.ref)])
    if getattr(args, "out", None):
        argv.extend(["--out", str(args.out)])
    if getattr(args, "glob", None):
        argv.extend(["--glob", str(args.glob)])
    argv.extend(str(x) for x in (getattr(args, "rest", None) or []))
    return exec_script("compare_vs_ref", argv)


def register(subparsers) -> None:
    p = subparsers.add_parser(
        "compare-ref",
        help="compare-ref: 改动 MD 改为段 vs 参考 docx 雷同检查 — distilled from bid-diff-and-revise",
    )
    sp = p.add_subparsers(dest="compare_action", metavar="<action>", required=True)

    sp_ref = sp.add_parser("ref", help="改动草稿 MD 改为段 vs 参考 docx 雷同检查", add_help=False)
    sp_ref.add_argument("--drafts-dir", dest="drafts_dir", required=False,
                        help="改动草稿 MD 所在目录")
    sp_ref.add_argument("--ref", required=False, help="参考 docx（主标或基准）")
    sp_ref.add_argument("--out", required=False, help="输出 MD 路径")
    sp_ref.add_argument("--glob", default="*改动草稿.md", help="MD 文件 glob 模式")
    sp_ref.add_argument("rest", nargs=argparse.REMAINDER)
    sp_ref.set_defaults(func=_run_ref)
