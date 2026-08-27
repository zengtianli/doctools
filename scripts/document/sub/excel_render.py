#!/usr/bin/env python3
"""用 **Microsoft Excel 自己** 把 xlsx 导成 PDF —— xlsx 版面核验的唯一可信渲染。

与同目录 `word_render.py` 同一条纪律的 xlsx 版（全局禁令：版面只认 Word / WPS / Excel，
第三方渲染器的断行分页与 Office 不同，拿它判版面 = 判了个替身）。

两条硬纪律，抄自 word_render 的实证：
① **只关目标那一个工作簿，绝不 quit Excel** —— 用户同时开着别的表还要用。
② **产物必须过页数门** —— 导出可能出一份 0 页/残废 PDF，不设门就会拿它报绿。

CLI:
    python3 excel_render.py <xlsx 绝对路径> <pdf 绝对路径> [--min-pages N]
"""
import argparse
import subprocess
import sys
from pathlib import Path

_CLOSE_ONE = '''
tell application "System Events"
  if not (exists process "Microsoft Excel") then return
end tell
tell application "Microsoft Excel"
  repeat with w in workbooks
    try
      if (name of w) is "{base}" then close w saving no
    end try
  end repeat
end tell
'''

_EXPORT = '''
tell application "Microsoft Excel"
  open POSIX file "{src}"
  delay 1
  set wb to active workbook
  save wb in (POSIX file "{dst}") as PDF file format
  close wb saving no
end tell
'''


def _osa(script: str) -> None:
    # Excel 收尾偶报 -128（User canceled）而产物已出 —— 一律以产物为准，不看退出码
    subprocess.run(["osascript", "-e", script],
                   stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL,
                   stdin=subprocess.DEVNULL, check=False)


def page_count(pdf: Path) -> int:
    r = subprocess.run(["pdfinfo", str(pdf)], capture_output=True, text=True)
    for line in r.stdout.splitlines():
        if line.startswith("Pages:"):
            return int(line.split()[1])
    return 0


def render(src: Path, dst: Path, min_pages: int = 1) -> int:
    src, dst = Path(src).resolve(), Path(dst).resolve()
    if not src.is_file():
        print(f"[excel_render] 源文件不存在: {src}", file=sys.stderr)
        return 2
    dst.parent.mkdir(parents=True, exist_ok=True)
    if dst.exists():
        dst.unlink()
    _osa(_CLOSE_ONE.format(base=src.name))      # 目标表若已开着，先只关它
    _osa(_EXPORT.format(src=src, dst=dst))
    _osa(_CLOSE_ONE.format(base=src.name))      # 收尾同样只关它
    if not dst.is_file():
        print("[excel_render] Excel 未产出 PDF", file=sys.stderr)
        return 1
    pages = page_count(dst)
    if pages < min_pages:
        print(f"[excel_render] 页数门未过：{pages} < {min_pages}", file=sys.stderr)
        return 1
    print(f"[excel_render] ✅ {dst}（{pages} 页）")
    return 0


def main(argv=None) -> int:
    p = argparse.ArgumentParser(prog="excel_render", description=__doc__,
                                formatter_class=argparse.RawDescriptionHelpFormatter)
    p.add_argument("xlsx")
    p.add_argument("pdf")
    p.add_argument("--min-pages", type=int, default=1)
    a = p.parse_args(argv)
    return render(Path(a.xlsx), Path(a.pdf), a.min_pages)


if __name__ == "__main__":
    sys.exit(main())
