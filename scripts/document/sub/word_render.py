#!/usr/bin/env python3
"""用 **Microsoft Word 自己** 把 docx 导成 PDF —— 版面核验的唯一可信渲染。

为什么必须是 Word 而不是 LibreOffice（2026-08-09 实证，代价是一整轮返工）：
断行规则不同。同一份 271 页论文，LO 数出 8 处「脚注号排行首」，Word 数出 40 处，
且具体位置几乎不重合。**拿 LO 的位置去改 Word 的版面 = 改了个替身**，会假绿。
LO 仍适合验内容完整性（字数/图数/部件），但凡涉及**断行、分页、行首行尾、
题注位置、页面填充**，一律用本模块。

两条硬纪律，都是踩出来的：
① **只关目标那一个文档，绝不 quit / pkill Word** —— 用户同时开着别的文档还要用
   （2026-08-09 用户原话「你关 word 的时候 就关对应的 docx 就好了…我其他的还要用」）。
② **产物必须过规模门** —— Word 偶尔会导出一份几 MB 的残废 PDF（实测 3.8MB / 13.9pt /
   页数不足），不设门就会拿它报绿。

CLI:
    python3 word_render.py <docx 绝对路径> <pdf 绝对路径> [--min-pages N]
"""
import argparse
import os
import subprocess
import sys
import time
from pathlib import Path

# 只关目标文档；System Events 先判进程在不在，Word 没开就什么都不做
_CLOSE_ONE = '''
tell application "System Events"
  if not (exists process "Microsoft Word") then return
end tell
tell application "Microsoft Word"
  repeat with d in documents
    try
      if (name of d) is "{base}" then close d saving no
    end try
  end repeat
end tell
'''

_EXPORT = '''
set src to POSIX file "{src}"
tell application "Microsoft Word"
  open src
  save as active document file name "{dst}" file format format PDF
  close active document saving no
end tell
'''


def _osa(script):
    # Word 收尾常报 -128（User canceled）而产物已出 —— 一律以产物为准，不看退出码
    subprocess.run(["osascript", "-e", script],
                   stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL,
                   stdin=subprocess.DEVNULL, check=False)


def page_count(pdf):
    try:
        import pdfplumber
        with pdfplumber.open(pdf) as p:
            return len(p.pages)
    except Exception:
        return 0


def render(src, dst, min_pages=1, timeout=180):
    """docx → pdf（Word 排版）。返回页数；不达标抛 SystemExit。"""
    src, dst = Path(src).resolve(), Path(dst).resolve()
    if not src.is_file():
        raise SystemExit(f"❌ 找不到 {src}")
    if dst.exists():
        dst.unlink()

    _osa(_CLOSE_ONE.format(base=src.name))
    _osa(_EXPORT.format(src=src, dst=dst))

    waited = 0
    while waited < timeout:
        if dst.exists() and dst.stat().st_size > 0:
            size = dst.stat().st_size
            time.sleep(2)
            if dst.stat().st_size == size:      # 大小稳定 = 写完了
                break
        time.sleep(2)
        waited += 2
    if not dst.exists() or dst.stat().st_size == 0:
        raise SystemExit(f"❌ Word 没有产出 {dst}（它可能弹了对话框，去看一眼）")

    n = page_count(dst)
    if n < min_pages:
        raise SystemExit(
            f"❌ 导出的 PDF 只有 {n} 页（要求 ≥ {min_pages}）—— 明显不是这份稿，"
            f"拒绝当有效渲染。实测出过 3.8MB/13.9pt 的残废产物。")
    print(f"✓ {dst}  ({dst.stat().st_size / 1048576:.0f}MB, {n} 页)")
    return n


def main():
    ap = argparse.ArgumentParser(description=__doc__,
                                 formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("docx")
    ap.add_argument("pdf")
    ap.add_argument("--min-pages", type=int, default=1,
                    help="规模门：少于这个页数就判渲染无效（长报告务必设，如 200）")
    a = ap.parse_args()
    if sys.platform != "darwin":
        raise SystemExit("❌ 本模块靠 macOS 的 osascript 驱动 Word")
    if not os.path.exists("/Applications/Microsoft Word.app"):
        raise SystemExit("❌ 本机没装 Microsoft Word")
    render(a.docx, a.pdf, a.min_pages)


if __name__ == "__main__":
    main()
