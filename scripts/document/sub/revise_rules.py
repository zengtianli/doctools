"""revise_rules.py — group module: revision rules generation (1 subcommand)

Subcommand:
  revise-rules gen --drafts-dir <dir> --docx <target.docx> --out <rules.json>
    ← gen_rules.py
    解析改动草稿 MD → 生成合并 rules JSON 用于 docx_tools.py track-changes
    含 3 道守卫: 引号 swap / title-only 跳过 / 说明性括号剥除

    注: is_title() 含若干标书常见标题关键词 (子方案/总体技术/工作部署/技术路线等)，
    非技术耦合——任何长度<80 + 匹配结构标题的段落均命中；水利/通用报告均适用。
    如需禁用 title-skip，传 --no-title-skip。

Standalone CLI 仍可用:
    python3 sub/gen_rules.py --drafts-dir 成果/md --docx 目标.docx --out rules.json

Distilled from ~/Work/shared/bids/.claude/skills/bid-diff-and-revise/scripts/ (2026-05-26)
"""
from __future__ import annotations

import argparse

from ._dispatch import exec_script


def _run_gen(args) -> int:
    argv: list[str] = []
    if getattr(args, "drafts_dir", None):
        argv.extend(["--drafts-dir", str(args.drafts_dir)])
    if getattr(args, "docx", None):
        argv.extend(["--docx", str(args.docx)])
    if getattr(args, "out", None):
        argv.extend(["--out", str(args.out)])
    if getattr(args, "glob", None):
        argv.extend(["--glob", str(args.glob)])
    if getattr(args, "no_title_skip", False):
        argv.append("--no-title-skip")
    if getattr(args, "no_paren_strip", False):
        argv.append("--no-paren-strip")
    argv.extend(str(x) for x in (getattr(args, "rest", None) or []))
    return exec_script("gen_rules", argv)


def register(subparsers) -> None:
    p = subparsers.add_parser(
        "revise-rules",
        help="revision rules ops: gen (改动 MD → rules JSON for track-changes)",
    )
    sp = p.add_subparsers(dest="revise_rules_action", metavar="<action>", required=True)

    sp_gen = sp.add_parser("gen", help="解析改动草稿 MD → rules JSON", add_help=False)
    sp_gen.add_argument("--drafts-dir", dest="drafts_dir", required=False,
                        help="改动草稿 MD 目录")
    sp_gen.add_argument("--docx", required=False, help="目标 docx（find 存在性校验 + 引号自适配）")
    sp_gen.add_argument("--out", required=False, help="输出 rules JSON 路径")
    sp_gen.add_argument("--glob", default="*-改动草稿.md", help="MD glob（默认 *-改动草稿.md）")
    sp_gen.add_argument("--no-title-skip", action="store_true", dest="no_title_skip",
                        help="不跳过 title-only rule")
    sp_gen.add_argument("--no-paren-strip", action="store_true", dest="no_paren_strip",
                        help="不剥除说明性括号")
    sp_gen.add_argument("rest", nargs=argparse.REMAINDER)
    sp_gen.set_defaults(func=_run_gen)
