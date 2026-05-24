#!/usr/bin/env python3
"""data.py — 数据处理统一 CLI（Phase 1 colocate · GOAL script-consolidation）

合并 4 个 xlsx/csv 脚本到单一 entry，统一 argparse + 共享并行 API。

顶级子命令:
  convert      格式互转（CSV/TXT/XLSX/XLS 8 子-子命令；nested · cf_api 范本）
  xlsx-lower   Excel/Word 文本小写化
  xlsx-merge   多表合并（AI 智能匹配）
  xlsx-split   工作表分离为独立文件

convert 子-子命令（2 层 sub-sub）:
  csv-from-txt / csv-to-txt / csv-merge-txt /
  xlsx-from-csv / xlsx-from-txt / xlsx-from-xls /
  xlsx-to-csv / xlsx-to-txt / encode-duplicates

并行批量（统一走 ~/Dev/tools/dev/lib/parallel_contract.py）:
  --workers N --batch FILE.jsonl --phases ... --fanout-evidence PATH

引用总部: parallel_contract（lib/parallel_contract.py · distill 上提）
被合脚本: convert.py / xlsx_lowercase.py / xlsx_merge_tables.py / xlsx_splitsheets.py
范本: cf_api.py（2 层 sub-sub）· menus.py（add_parallel_args 双层声明）
"""
from __future__ import annotations

import argparse
import sys
from pathlib import Path

# --- 路径设置 ----------------------------------------------------
_SELF = Path(__file__).resolve()
_REPO_LIB = _SELF.parent.parent.parent / "lib"
_DEV_LIB = Path.home() / "Dev" / "tools" / "dev" / "lib"
_DATA_DIR = _SELF.parent

# 让 lib (display/file_ops/...) + 同目录旧脚本可被 import
for p in (_REPO_LIB, _DEV_LIB, _DATA_DIR):
    sp = str(p)
    if sp not in sys.path:
        sys.path.insert(0, sp)

# 总部并行合约（铁律 #5 不造轮子）
from parallel_contract import add_parallel_args  # noqa: E402

SCRIPT_VERSION = "1.0.0"


# --- dispatcher: 调用旧脚本 main(),保零破坏 ---------------------

def _dispatch(module_name: str, argv: list[str]) -> int:
    """import 旧脚本模块,改写 sys.argv 后调 main()。

    Why: 4 旧脚本 main 设计成读 sys.argv;包装最稳。旧脚本依然可独立调用
    (thin alias 留旧路径),本 CLI 仅做顶层 sub-router。
    """
    import importlib
    saved = sys.argv[:]
    try:
        sys.argv = [module_name] + argv
        mod = importlib.import_module(module_name)
        if hasattr(mod, "main"):
            rc = mod.main()
            return int(rc) if isinstance(rc, int) else 0
        return 0
    except SystemExit as e:
        return int(e.code) if isinstance(e.code, int) else 0
    finally:
        sys.argv = saved


# --- cmd handlers -------------------------------------------------

def cmd_convert(args: argparse.Namespace, rest: list[str]) -> int:
    """convert 转发到 convert.py main(),保持 8 sub-sub 完整。

    本层 argparse 只识别 sub-cmd 名,其余 flag 透传给 convert.py
    (它已有完整 --batch/--workers/--phases/--defer 实现)。
    """
    sub_argv: list[str] = []
    if args.subcommand:
        sub_argv.append(args.subcommand)
    # 顶层并行 flag → 透传
    if args.workers:
        sub_argv += ["--workers", str(args.workers)]
    if args.batch:
        sub_argv += ["--batch", args.batch]
    sub_argv += rest
    return _dispatch("convert", sub_argv)


def cmd_xlsx_lower(args: argparse.Namespace, rest: list[str]) -> int:
    """xlsx-lower 转发到 xlsx_lowercase.py (Word/Excel 文本小写化)。"""
    return _dispatch("xlsx_lowercase", rest)


def cmd_xlsx_merge(args: argparse.Namespace, rest: list[str]) -> int:
    """xlsx-merge 转发到 xlsx_merge_tables.py (AI 智能匹配多表合并)。"""
    return _dispatch("xlsx_merge_tables", rest)


def cmd_xlsx_split(args: argparse.Namespace, rest: list[str]) -> int:
    """xlsx-split 转发到 xlsx_splitsheets.py (工作表分离)。"""
    return _dispatch("xlsx_splitsheets", rest)


# --- parser 构建 --------------------------------------------------

# 8 个 convert 子-子命令名 (与 convert.py CONVERTERS 一致)
CONVERT_SUBCOMMANDS = (
    "csv-from-txt",
    "csv-to-txt",
    "csv-merge-txt",
    "xlsx-from-csv",
    "xlsx-from-txt",
    "xlsx-from-xls",
    "xlsx-to-csv",
    "xlsx-to-txt",
    "encode-duplicates",
)


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        prog="data",
        description="数据处理统一 CLI（4 顶级子命令 + convert 8 子-子）",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog=(
            "示例:\n"
            "  data convert csv-from-txt file.txt\n"
            "  data convert --batch tasks.jsonl --workers 8\n"
            "  data xlsx-lower file.xlsx\n"
            "  data xlsx-split book.xlsx -j 4\n"
            "  data xlsx-merge --master m.xlsx --master-key 名称 \\\n"
            "                   --aux a.xlsx --aux-key 工程 -o out.xlsx\n"
        ),
    )
    parser.add_argument("--version", action="version", version=f"data {SCRIPT_VERSION}")

    sub = parser.add_subparsers(dest="command", metavar="COMMAND")

    # convert (2 层 sub-sub · cf_api 范本)
    p_conv = sub.add_parser(
        "convert",
        help="格式互转 (CSV/TXT/XLSX/XLS 8 sub-sub)",
        description="数据格式转换（含批量并行）。详细 flag: data convert --help",
        add_help=False,  # 透传 --help 给 convert.py 显示其原 docstring
    )
    p_conv.add_argument(
        "subcommand",
        nargs="?",
        choices=CONVERT_SUBCOMMANDS,
        metavar="SUBCOMMAND",
        help=f"sub-sub 命令 ({'/'.join(CONVERT_SUBCOMMANDS)})",
    )
    # 顶层并行 flag (透传给 convert.py 已有的 --workers/--batch)
    add_parallel_args(p_conv, support_phases=False, support_fanout=False)
    p_conv.add_argument("-h", "--help", action="store_true", help="显示 convert 详细帮助")
    p_conv.set_defaults(func=cmd_convert)

    # xlsx-lower
    p_low = sub.add_parser(
        "xlsx-lower",
        help="Excel/Word 文本小写化 (.docx/.xlsx/.xlsm)",
        add_help=False,
    )
    p_low.set_defaults(func=cmd_xlsx_lower)

    # xlsx-merge
    p_merge = sub.add_parser(
        "xlsx-merge",
        help="多表合并 (AI 智能匹配)",
        add_help=False,
    )
    p_merge.set_defaults(func=cmd_xlsx_merge)

    # xlsx-split
    p_split = sub.add_parser(
        "xlsx-split",
        help="工作表分离为独立 .xlsx 文件",
        add_help=False,
    )
    p_split.set_defaults(func=cmd_xlsx_split)

    return parser


def main(argv: list[str] | None = None) -> int:
    argv = list(sys.argv[1:]) if argv is None else list(argv)
    parser = build_parser()
    # 用 parse_known_args 透传剩余 flag 给被合脚本(它们各自有 argparse)
    args, rest = parser.parse_known_args(argv)

    if not args.command:
        parser.print_help()
        return 0

    func = getattr(args, "func", None)
    if func is None:
        parser.print_help()
        return 1

    # convert 的 --help 透传
    if args.command == "convert" and getattr(args, "help", False):
        rest = ["--help"] + rest

    return func(args, rest)


if __name__ == "__main__":
    try:
        sys.exit(main())
    except KeyboardInterrupt:
        print("用户中断", file=sys.stderr)
        sys.exit(130)
