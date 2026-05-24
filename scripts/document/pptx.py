#!/usr/bin/env python3
"""
PPTX 文档工具集统一入口 v1.0.0

合并 3 个 PPTX 脚本到单一 argparse subparsers 多子命令工具:

子命令:
    font    - 字体统一为微软雅黑（嵌入自 pptx_tools.py）
    format  - 文本格式修复（引号/标点/单位，嵌入自 pptx_tools.py）
    table   - 表格样式设置（嵌入自 pptx_tools.py）
    all     - 一键标准化 format -> font -> table（嵌入自 pptx_tools.py）
    to-md   - PPTX 转 Markdown（取代 pptx_to_md.py）
    chart   - 数据驱动图表生成 JSON -> PNG（取代 chart.py）

用法:
    python3 pptx.py <subcommand> [args...]
    python3 pptx.py font presentation.pptx
    python3 pptx.py all *.pptx --workers 8
    python3 pptx.py to-md slides.pptx
    python3 pptx.py chart bar config.json -o out.png

向后兼容:
    旧脚本 pptx_tools.py / pptx_to_md.py / chart.py 仍保留，不删除。
    本入口只是统一调度，将 font/format/table/all 内部转发到 pptx_tools，
    to-md 转发到 pptx_to_md，chart 转发到 chart。

作者: tianli
日期: 2026-05-23
版本: 1.0.0
"""

import argparse
import importlib.util
import sys
from pathlib import Path

SCRIPT_VERSION = "1.0.0"
SCRIPT_DIR = Path(__file__).resolve().parent

# 把 lib 加入 sys.path（与 pptx_tools.py / pptx_to_md.py / chart.py 一致）
# 注意：**不把 SCRIPT_DIR 加入 sys.path** — 本文件名为 `pptx.py`，会和
# 第三方 `python-pptx` 包的 `pptx` 顶级模块冲突，导致 sibling 脚本里
# `from pptx import Presentation` 拿到我们这个文件而不是 python-pptx。
sys.path.insert(0, str(SCRIPT_DIR.parent.parent / "lib"))
sys.path.insert(0, str(Path.home() / "Dev" / "tools" / "dev" / "lib"))


def _purge_script_dir_from_path():
    """从 sys.path 删除「本脚本所在目录」+ sys.modules 删除可能误注的 `pptx`。

    原因：Python 启动时把 `__main__` 脚本的目录自动塞进 sys.path[0]。
    本脚本叫 `pptx.py`，会作为顶级模块 `pptx` 命中 sibling 脚本里
    `from pptx import Presentation`，把 python-pptx 包整个屏蔽。
    因此在加载 sibling 之前先清掉这个目录 + 强制让 `pptx` 重新解析为 python-pptx。
    """
    script_dir = str(SCRIPT_DIR)
    sys.path[:] = [p for p in sys.path if Path(p).resolve() != SCRIPT_DIR and p != script_dir and p != ""]
    # 把可能被误加载的 `pptx` 顶级模块踢出 sys.modules，强制下次 import 从 site-packages 走
    for name in list(sys.modules.keys()):
        if name == "pptx" or name.startswith("pptx."):
            mod = sys.modules[name]
            mod_file = getattr(mod, "__file__", "") or ""
            if str(SCRIPT_DIR) in mod_file:
                del sys.modules[name]


def _load_sibling(module_alias: str, filename: str):
    """用 importlib 按绝对路径加载同目录脚本，避免和 python-pptx 包名冲突。

    module_alias 必须避开 `pptx` 顶级名 — 用 `_pptx_tools_sibling` 这种别名。
    """
    _purge_script_dir_from_path()
    path = SCRIPT_DIR / filename
    spec = importlib.util.spec_from_file_location(module_alias, path)
    if spec is None or spec.loader is None:
        raise ImportError(f"无法加载 {path}")
    mod = importlib.util.module_from_spec(spec)
    sys.modules[module_alias] = mod
    spec.loader.exec_module(mod)
    return mod


def _run_module_main(mod, argv0: str, argv: list[str]) -> int:
    """用 argv 包装调用 mod.main()，捕获 SystemExit。"""
    saved = sys.argv
    try:
        sys.argv = [argv0, *argv]
        try:
            mod.main()
        except SystemExit as e:
            return int(e.code) if e.code is not None else 0
        return 0
    finally:
        sys.argv = saved


def _dispatch_pptx_tools(subcommand: str, argv: list[str]) -> int:
    """转发 font/format/table/all 到 pptx_tools.py"""
    mod = _load_sibling("_pptx_tools_sibling", "pptx_tools.py")
    return _run_module_main(mod, "pptx_tools", [subcommand, *argv])


def _dispatch_to_md(argv: list[str]) -> int:
    """转发 to-md 到 pptx_to_md.py"""
    mod = _load_sibling("_pptx_to_md_sibling", "pptx_to_md.py")
    return _run_module_main(mod, "pptx_to_md", argv)


def _dispatch_chart(argv: list[str]) -> int:
    """转发 chart 到 chart.py"""
    mod = _load_sibling("_chart_sibling", "chart.py")
    return _run_module_main(mod, "chart", argv)


SUBCOMMANDS = {
    "font": "字体统一为微软雅黑（pptx_tools font）",
    "format": "文本格式修复 引号/标点/单位（pptx_tools format）",
    "table": "表格样式设置 标题行/镶边行/首列（pptx_tools table）",
    "all": "一键标准化 format -> font -> table（pptx_tools all）",
    "to-md": "PPTX 转 Markdown（取代 pptx_to_md.py）",
    "chart": "数据驱动图表生成 JSON -> PNG（取代 chart.py）",
}


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        prog="pptx.py",
        description="PPTX 文档工具集统一入口 v%s（6 子命令：font/format/table/all/to-md/chart）"
        % SCRIPT_VERSION,
        epilog=(
            "子命令说明:\n"
            + "\n".join(f"  {k:<8}{v}" for k, v in SUBCOMMANDS.items())
            + "\n\n"
            "各子命令完整 --help: python3 pptx.py <subcommand> --help"
        ),
        formatter_class=argparse.RawDescriptionHelpFormatter,
        add_help=True,
    )
    parser.add_argument(
        "subcommand",
        choices=list(SUBCOMMANDS.keys()),
        help="子命令: " + " | ".join(SUBCOMMANDS.keys()),
    )
    parser.add_argument(
        "args",
        nargs=argparse.REMAINDER,
        help="转发给子命令的参数（用 --help 看具体用法）",
    )
    parser.add_argument(
        "--version", action="version", version=f"%(prog)s {SCRIPT_VERSION}"
    )
    return parser


def main() -> int:
    parser = build_parser()

    # 无参数时显示帮助
    if len(sys.argv) < 2:
        parser.print_help()
        return 1

    args = parser.parse_args()
    sub = args.subcommand
    sub_argv = args.args or []

    if sub in ("font", "format", "table", "all"):
        return _dispatch_pptx_tools(sub, sub_argv)
    if sub == "to-md":
        return _dispatch_to_md(sub_argv)
    if sub == "chart":
        return _dispatch_chart(sub_argv)

    parser.error(f"未知子命令: {sub}")
    return 2


if __name__ == "__main__":
    sys.exit(main())
