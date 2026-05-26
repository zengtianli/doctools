"""table — table structural ops (W4 distill · 2026-05-26)

Targets:
  delete-rows   — 删除 docx 指定表格的行范围，含安全校验（distilled from
                  ~/Work/shared/bids/.claude/skills/bid-diff-and-revise/scripts/delete_table_rows.py）
"""
from __future__ import annotations
import argparse
from pathlib import Path

from . import _dispatch

_GROUP = "table"
_HELP = "表格操作 (delete-rows / ...)"


def register(subparsers) -> None:
    """Register `table` group onto top-level subparsers."""
    grp = _dispatch.get_or_add_group(subparsers, _GROUP, _HELP)
    subs = _dispatch.get_or_add_subparsers(grp, dest=f"{_GROUP}_target",
                                            metavar="<target>")

    # ── delete-rows ──────────────────────────────────────────────────────────
    sp = subs.add_parser(
        "delete-rows",
        help="删除指定表格的行范围（安全校验 + 原地保存）",
        description=(
            "删除 docx 中指定表（0-based 索引）的行范围 FROM:TO（闭区间 0-based）。\n"
            "支持三段安全校验：保留行首列期望值、被删行关键字、删后末行期望值。\n"
            "原地修改 docx，保存后重读验证。\n\n"
            "用途：track-changes 不支持表行删除时，由脚本直接删除。"
        ),
    )
    sp.add_argument("docx", help="目标 docx（原地修改）")
    sp.add_argument("--table-index", type=int, required=True, help="表格索引（0-based）")
    sp.add_argument("--rows", required=True, help="行范围 FROM:TO（闭区间 0-based）")
    sp.add_argument("--expected-first-col", default="",
                    help="删除前保留行第一列期望值（逗号分隔），安全校验；留空跳过")
    sp.add_argument("--expected-residue", default="",
                    help="被删起始行应含此关键字，确保删对行")
    sp.add_argument("--expected-last", default="",
                    help="删后末行第二列期望值，校验删后结果")
    sp.set_defaults(_sub_target="delete-rows")
    grp.set_defaults(_group=_GROUP)


def handle(args: argparse.Namespace, rest: list[str]) -> int:
    target = getattr(args, f"{_GROUP}_target", None) or getattr(args, "_sub_target", None)
    if target == "delete-rows":
        argv = ["--docx", args.docx,
                "--table-index", str(args.table_index),
                "--rows", args.rows]
        if args.expected_first_col:
            argv += ["--expected-first-col", args.expected_first_col]
        if args.expected_residue:
            argv += ["--expected-residue", args.expected_residue]
        if args.expected_last:
            argv += ["--expected-last", args.expected_last]
        return _dispatch.exec_script("delete_table_rows", argv)
    print(f"[table] unknown target: {target}", flush=True)
    return 1
