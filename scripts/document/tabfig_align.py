#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""tabfig_align — 表/图题注号与所在章号对齐（单一职责,报告/标书通用）。

问题域: 章号位移(chapter_renumber)后,正文里的 `表 9.2-1　…` / `图 8-1` 类编号
前缀还是旧章号。本脚本只干一件事: **让每个 表/图 编号的章号段 = 它所在
章文件的章号**(从文件名 ch<N>-*.md 取),自愈式、幂等、与位移映射解耦——
不管中间改过几轮章号,跑一次就对。

不做的事(各归各的脚本): 章文件名/H1/PNG/FACTS 位移=chapter_renumber.py;
子标题派生=项目 number_headings.py;渲染=gen_bid_docx.py。

用法:
  python3 tabfig_align.py <chapters.yaml|章节目录> [--apply|--check]
    (默认)   干跑,列出待改项,exit 0
    --apply  写回文件
    --check  机检门: 有漂移 exit 2,干净 exit 0
"""
import re
import sys
from pathlib import Path

TOKEN_RE = re.compile(r"([表图])(\s*)(\d+(?:\.\d+)*)(-\d+)")
CH_RE = re.compile(r"^ch(\d+(?:\.\d+)*)-")


def chapter_files(arg: Path):
    if arg.is_dir():
        return sorted(arg.glob("ch*.md"))
    # chapters.yaml → 走总部 lib 的 targets(chapters_glob 相对 config 目录)
    sys.path.insert(0, str(Path(__file__).resolve().parents[2] / "lib"))
    from chapter_numbering import ChapterNumbering
    cn = ChapterNumbering(arg)
    return sorted(cn.root.glob(cn.targets()["chapters_glob"]))


def align_text(text: str, ch_num: str):
    """返回 (新文本, [(旧token, 新token), ...])。只改章号段,保留序号与空白。"""
    changes = []

    def sub(m):
        kind, sp, num, tail = m.groups()
        if num == ch_num:
            return m.group(0)
        new = f"{kind}{sp}{ch_num}{tail}"
        changes.append((m.group(0), new))
        return new

    return TOKEN_RE.sub(sub, text), changes


def main(argv):
    args = [a for a in argv if not a.startswith("--")]
    if not args:
        print(__doc__)
        return 1
    apply_ = "--apply" in argv
    check = "--check" in argv
    src = Path(args[0]).expanduser().resolve()

    total = 0
    for f in chapter_files(src):
        m = CH_RE.match(f.name)
        if not m:
            continue
        new_text, changes = align_text(f.read_text(encoding="utf-8"), m.group(1))
        if not changes:
            continue
        total += len(changes)
        print(f"{f.name}  ({len(changes)} 处)")
        for old, new in changes:
            print(f"  {old}  →  {new}")
        if apply_:
            f.write_text(new_text, encoding="utf-8")

    if total == 0:
        print("✅ 表/图编号与章号全部对齐,无需改动")
        return 0
    if apply_:
        print(f"✅ 已写回 {total} 处")
        return 0
    print(f"⚠ 共 {total} 处待对齐(干跑未写)。--apply 执行,--check 作机检门")
    return 2 if check else 0


if __name__ == "__main__":
    sys.exit(main(sys.argv[1:]))
