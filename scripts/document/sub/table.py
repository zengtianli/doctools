"""table — table structural ops (W4 distill · 2026-05-26)

Targets:
  delete-rows   — 删除 docx 指定表格的行范围，含安全校验（distilled from
                  ~/Work/shared/bids/.claude/skills/bid-diff-and-revise/scripts/delete_table_rows.py）
  extract       — 抽 docx 内每张表为独立 docx，文件名 = 邻近 caption 段文字
                  (distilled from eco-flow/taizhou-天台 need · 2026-05-28)
"""
from __future__ import annotations
import argparse
from pathlib import Path

from . import _dispatch

_GROUP = "table"
_HELP = "表格操作 (delete-rows / extract / ...)"


def _run_extract(args) -> int:
    """func= handler for `table extract` (invoked by docx_cli.py's distilled path)."""
    docx = getattr(args, "docx_kw", None) or getattr(args, "docx_path_pos", None)
    if not docx:
        print("[table extract] missing docx (positional or --docx)", flush=True)
        return 2
    argv = ["--docx", str(docx),
            "--out-dir", str(args.out_dir),
            "--name-pattern", str(args.name_pattern)]
    if getattr(args, "dry_run", False):
        argv.append("--dry-run")
    return _dispatch.exec_script("extract_tables", argv)


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

    # ── extract ──────────────────────────────────────────────────────────────
    sp_ex = subs.add_parser(
        "extract",
        help="抽 docx 内每张表为独立 docx（文件名=邻近 caption）",
        description=(
            "遍历 docx body 内每张表 <w:tbl>，往前 1-2 段（或往后 1 段）扫"
            "「表…」开头的 caption 段，取文字作文件名 stem。\n"
            "无 caption → fallback `table-{idx:02d}`；重名加 -2/-3。\n\n"
            "输出策略：shutil.copy 源 docx → 删 body 内非目标元素（保留 sectPr）→ save，\n"
            "保留完整 styles / numbering / media / rels（与 split_by_h1 同套路）。"
        ),
    )
    sp_ex.add_argument("docx_path_pos", nargs="?", help="(positional) docx 路径，等价 --docx")
    sp_ex.add_argument("--docx", dest="docx_kw", help="目标 docx 路径")
    sp_ex.add_argument("--out-dir", required=True, help="输出目录（mkdir -p）")
    sp_ex.add_argument(
        "--name-pattern", default="{stem}.docx",
        help="文件名 pattern，默认 '{stem}.docx'（可用 {stem} {idx}）",
    )
    sp_ex.add_argument("--dry-run", action="store_true", help="只打印 plan，不写文件")
    sp_ex.set_defaults(_sub_target="extract", func=_run_extract)

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
    if target == "extract":
        docx = getattr(args, "docx_kw", None) or getattr(args, "docx_path_pos", None)
        if not docx:
            print("[table extract] missing docx (positional or --docx)", flush=True)
            return 2
        argv = ["--docx", str(docx),
                "--out-dir", str(args.out_dir),
                "--name-pattern", str(args.name_pattern)]
        if args.dry_run:
            argv.append("--dry-run")
        return _dispatch.exec_script("extract_tables", argv)
    print(f"[table] unknown target: {target}", flush=True)
    return 1
