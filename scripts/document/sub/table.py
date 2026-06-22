"""table — table structural ops (W4 distill · 2026-05-26)

Targets:
  delete-rows   — 删除 docx 指定表格的行范围，含安全校验（distilled from
                  ~/Work/shared/bids/.claude/skills/bid-diff-and-revise/scripts/delete_table_rows.py）
  extract       — 抽 docx 内每张表为独立 docx，文件名 = 邻近 caption 段文字
                  (distilled from eco-flow/taizhou-天台 need · 2026-05-28)
  borders       — 把所有表格统一为「满格实线」（表级 tblBorders 全 single +
                  默认清单元格级 tcBorders 的 nil 覆盖；含嵌套表 · 2026-06-01）
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


def _run_borders(args) -> int:
    """func= handler for `table borders`（distilled 路径直接调用）。

    --center 时附带跑 set_table_align（边框+整体居中常一起要）。
    """
    docx = getattr(args, "docx_kw", None) or getattr(args, "docx_path_pos", None)
    if not docx:
        print("[table borders] missing docx (positional or --docx)", flush=True)
        return 2
    argv = ["--docx", str(docx), "--val", str(args.val), "--sz", str(args.sz),
            "--color", str(args.color), "--space", str(args.space)]
    if getattr(args, "keep_cell_borders", False):
        argv.append("--keep-cell-borders")
    if getattr(args, "no_backup", False):
        argv.append("--no-backup")
    if getattr(args, "dry_run", False):
        argv.append("--dry-run")
    rc = _dispatch.exec_script("set_table_borders", argv)
    if rc == 0 and getattr(args, "center", False):
        rc = _run_center(args)
    return rc


def _run_center(args) -> int:
    """func= handler for `table center`（表格整体水平居中）。"""
    docx = getattr(args, "docx_kw", None) or getattr(args, "docx_path_pos", None)
    if not docx:
        print("[table center] missing docx (positional or --docx)", flush=True)
        return 2
    argv = ["--docx", str(docx)]
    if getattr(args, "cell_center", False):
        argv.append("--cell-center")
    if getattr(args, "no_backup", False):
        argv.append("--no-backup")
    if getattr(args, "dry_run", False):
        argv.append("--dry-run")
    return _dispatch.exec_script("set_table_align", argv)


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

    # ── borders ───────────────────────────────────────────────────────────────
    sp_bd = subs.add_parser(
        "borders",
        help="把所有表格统一为满格实线（表级全 single + 清单元格 nil 覆盖）",
        description=(
            "把 docx 内所有表格（含嵌套表）统一为「满格实线」边框。\n\n"
            "根因：表级 <w:tblBorders> 与单元格级 <w:tcBorders> 两级，后者优先级更高。\n"
            "某表「看着不是全实线」常因表级没设边框、且部分单元格 tcBorders 把内部边\n"
            "设成 val='nil'（无线）→ 内部竖/横线缺失。光设表级盖不住单元格 nil。\n\n"
            "本命令两手抓：① 每表设表级 tblBorders 6 边全 single；\n"
            "② 默认删每个单元格 tcBorders（让表级统一生效），或 --keep-cell-borders\n"
            "时只把非 single 边改写为实线。原地改 + 默认备份 .bak-时间戳。"
        ),
    )
    sp_bd.add_argument("docx_path_pos", nargs="?", help="(positional) docx 路径，等价 --docx")
    sp_bd.add_argument("--docx", dest="docx_kw", help="目标 docx（原地修改）")
    sp_bd.add_argument("--val", default="single", help="线型，默认 single 实线")
    sp_bd.add_argument("--sz", type=int, default=4, help="线宽 1/8pt（4=0.5pt），默认 4")
    sp_bd.add_argument("--color", default="auto", help="颜色，默认 auto（黑）")
    sp_bd.add_argument("--space", type=int, default=0, help="边距，默认 0")
    sp_bd.add_argument("--keep-cell-borders", action="store_true",
                       help="保留 tcBorders，仅改写非实线边（默认=删 tcBorders）")
    sp_bd.add_argument("--center", action="store_true",
                       help="顺带把所有表格整体水平居中（= 再跑一次 table center）")
    sp_bd.add_argument("--cell-center", action="store_true",
                       help="配合 --center：单元格内文字也水平+垂直居中")
    sp_bd.add_argument("--no-backup", action="store_true", help="不创建 .bak-时间戳")
    sp_bd.add_argument("--dry-run", action="store_true", help="只报告不写盘")
    sp_bd.set_defaults(_sub_target="borders", func=_run_borders)

    # ── center ─────────────────────────────────────────────────────────────────
    sp_ct = subs.add_parser(
        "center",
        help="把所有表格整体在页面水平居中（含嵌套表）",
        description=(
            "把 docx 内所有表格（含嵌套表）整体在页面水平居中——写表级\n"
            "<w:tblPr>/<w:jc w:val='center'/>（表作为块左右居中），与单元格内文字\n"
            "对齐是两回事。--cell-center 时同时把单元格内文字水平+垂直居中。\n"
            "原地改 + 默认备份 .bak-时间戳。"
        ),
    )
    sp_ct.add_argument("docx_path_pos", nargs="?", help="(positional) docx 路径，等价 --docx")
    sp_ct.add_argument("--docx", dest="docx_kw", help="目标 docx（原地修改）")
    sp_ct.add_argument("--cell-center", action="store_true",
                       help="同时把单元格内文字水平+垂直居中（默认只表格整体居中）")
    sp_ct.add_argument("--no-backup", action="store_true", help="不创建 .bak-时间戳")
    sp_ct.add_argument("--dry-run", action="store_true", help="只报告不写盘")
    sp_ct.set_defaults(_sub_target="center", func=_run_center)

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
    if target == "borders":
        return _run_borders(args)
    if target == "center":
        return _run_center(args)
    print(f"[table] unknown target: {target}", flush=True)
    return 1
