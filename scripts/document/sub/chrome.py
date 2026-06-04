"""chrome.py — group module: 院报告版面装帧 (逐章分节 + 逐章页眉页脚水印 + 宽表横向节).

Subcommand:
  chrome --raw <raw.docx> --template <范式.docx> [--out <out.docx>] [--county 天台县]
        把 raw 合并稿重建成院报告版面: 逐章分节·每章 running-title 页眉(标题+章名)·
        院名页码页脚·宽表(gridCol>11000twip)/宽图(extent cx>纵向可用宽)就地横向节·
        附表/附图整章横向。页眉脚部件取自 --template, 县名 swap 复用他县。
        纯 zip+lxml, 媒体/公式/OLE verbatim。默认输出 <raw stem>_chrome.docx。
  chrome --validate <out.docx> <范式.docx>
        diff 节结构(orient+章名) 对照范式, 验证复刻正确性。

Standalone CLI:
    python3 sub/docx_chrome.py --raw X --template Y [--out Z] [--county 景宁县]

Trigger: 县市报告交付前版面定型(eco-flow 院范式); 多县复用(--county swap)。
distilled from eco-flow/taizhou-天台 (2026-06-04).
"""
from __future__ import annotations

import argparse

from ._dispatch import exec_script


def _run(args) -> int:
    argv: list[str] = []
    if getattr(args, "validate", None):
        argv = ["--validate"] + [str(x) for x in args.validate]
        return exec_script("docx_chrome", argv)
    if getattr(args, "raw", None):       argv += ["--raw", str(args.raw)]
    if getattr(args, "template", None):  argv += ["--template", str(args.template)]
    if getattr(args, "out", None):       argv += ["--out", str(args.out)]
    if getattr(args, "county", None):    argv += ["--county", str(args.county)]
    extra = getattr(args, "rest", None) or []
    argv.extend(str(x) for x in extra)
    return exec_script("docx_chrome", argv)


def register(subparsers) -> None:
    p = subparsers.add_parser(
        "chrome",
        help="院报告版面装帧: 逐章分节+逐章页眉页脚水印+宽表横向节 (→ <raw>_chrome.docx)",
    )
    p.add_argument("--raw", help="输入 raw 合并 docx(正文已成型)")
    p.add_argument("--template", help="范式 docx(取页眉/页脚/水印部件)")
    p.add_argument("--out", help="输出(默认 <raw stem>_chrome.docx)")
    p.add_argument("--county", default="天台县", help="目标县名(替换范式县名)")
    p.add_argument("--validate", nargs=2, metavar=("OUT", "TEMPLATE"),
                   help="diff 节结构对照范式而非构建")
    p.add_argument("rest", nargs=argparse.REMAINDER)
    p.set_defaults(func=_run)
