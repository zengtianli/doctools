"""fonts — 字体/颜色规整族（2026-06-21）

Targets:
  normalize  — 统一中西文字体 + 标题颜色压黑（杀 pandoc 主题蓝）。
               标题→黑体/黑色 · 正文→宋体 · 西文→Times New Roman。
               样式级 + run 级双保险。distilled from reclaim 节水年会征文需求。
"""
from __future__ import annotations
import argparse

from . import _dispatch

_GROUP = "fonts"
_HELP = "字体/颜色规整 (normalize)"


def _run_normalize(args) -> int:
    docx = getattr(args, "docx_kw", None) or getattr(args, "docx_path_pos", None)
    if not docx:
        print("[fonts normalize] missing docx (positional or --docx)", flush=True)
        return 2
    argv = ["--docx", str(docx),
            "--body-cjk", str(args.body_cjk),
            "--heading-cjk", str(args.heading_cjk),
            "--latin", str(args.latin),
            "--heading-color", str(args.heading_color)]
    if getattr(args, "no_backup", False):
        argv.append("--no-backup")
    if getattr(args, "dry_run", False):
        argv.append("--dry-run")
    return _dispatch.exec_script("normalize_fonts", argv)


def register(subparsers) -> None:
    grp = _dispatch.get_or_add_group(subparsers, _GROUP, _HELP)
    subs = _dispatch.get_or_add_subparsers(grp, dest=f"{_GROUP}_target",
                                            metavar="<target>")
    sp = subs.add_parser(
        "normalize",
        help="统一中西文字体 + 标题颜色压黑（杀 pandoc 主题蓝）",
        description=(
            "统一 docx 中西文字体并把标题颜色压成黑色。\n\n"
            "根因：pandoc 套 Word 内置 Heading 样式自带主题蓝（H1 无色继承蓝、\n"
            "H3/4/5=0F4761 青蓝），生成脚本只改字号没改色 → 标题蓝。本命令：\n"
            "  · 标题段（样式名含 heading/标题/title/subtitle/toc）：中文→黑体、\n"
            "    西文→Times New Roman、颜色→黑(000000)。\n"
            "  · 其余段（正文/图注/表格内）：中文→宋体、西文→Times New Roman；颜色不动。\n"
            "样式级 + run 级双保险（防 pandoc run 直接 rFonts 盖样式）。\n"
            "原地改 + 默认备份 .bak-时间戳。"
        ),
    )
    sp.add_argument("docx_path_pos", nargs="?", help="(positional) docx 路径，等价 --docx")
    sp.add_argument("--docx", dest="docx_kw", help="目标 docx（原地修改）")
    sp.add_argument("--body-cjk", default="宋体", help="正文中文字体，默认 宋体")
    sp.add_argument("--heading-cjk", default="黑体", help="标题中文字体，默认 黑体")
    sp.add_argument("--latin", default="Times New Roman", help="西文字体，默认 Times New Roman")
    sp.add_argument("--heading-color", default="000000", help="标题颜色 hex，默认 000000")
    sp.add_argument("--no-backup", action="store_true", help="不创建 .bak-时间戳")
    sp.add_argument("--dry-run", action="store_true", help="只报告不写盘")
    sp.set_defaults(_sub_target="normalize", func=_run_normalize)

    grp.set_defaults(_group=_GROUP)


def handle(args: argparse.Namespace, rest: list[str]) -> int:
    target = getattr(args, f"{_GROUP}_target", None) or getattr(args, "_sub_target", None)
    if target == "normalize":
        return _run_normalize(args)
    print(f"[fonts] unknown target: {target}", flush=True)
    return 1
