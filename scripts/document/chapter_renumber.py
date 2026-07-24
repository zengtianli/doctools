#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""章号统一位移引擎(通用·config 驱动) —— 报告/标书序号调整的总部 SSOT 工具。

改章号 = 只改项目 chapters.yaml 的 number_base 一个数,跑本引擎 --apply;源码/正文不用手改。
按「当前号→目标号」整数映射,位移所有引用点的前导章号:
  ① 章节 md: H1 首行章号 + 图题注号(默认 【图 X-Y】)
  ② 章节文件名: ch<号>-<slug>.md
  ③ 成图 PNG: 图<号>-<k>_<名>.png   (两段式改名防撞号)
  ④ FACTS.md facts-machine 块的章号键 "X"/"X.Y"
  ⑤ 目录大纲: 第N章 / chN / 列表与标题前导号 / N—M 区间
目标路径来自 chapters.yaml 的 renumber_targets(缺省=标书/报告约定,见 chapter_numbering.DEFAULT_TARGETS)。
幂等: 磁盘已达标 → no-op。位移后需项目侧重跑 number_headings(重派子号)+ 重渲。

chapters.yaml 最小 schema:
    number_base: 10                 # 首章显示章号
    sequence:                       # 逻辑章序(固定);含 subs 的章占一个整号,子节 .1~.n
      - {slug: 基础资料, title: 对本项目的基础资料掌握情况}
      - slug: 工作方案
        title: 工作方案
        subs: [{slug: 思路, title: 工作思路…}, …]
      - …
    header_body: |                   # (可选)含 subs 章的章头引语,勿写子节号
    wide_figure_keywords: [路线图, 甘特图]   # (可选)满幅宽图关键词
    renumber_targets: {…}           # (可选)覆盖默认位移路径

用法:
    python3 chapter_renumber.py [chapters.yaml]          # 干跑(默认找 ./技术标/chapters.yaml 或 ./chapters.yaml)
    python3 chapter_renumber.py 技术标/chapters.yaml --apply
"""
import os
import re
import subprocess
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[2] / "lib"))
from chapter_numbering import ChapterNumbering  # noqa: E402

FACTS_FENCE = re.compile(r"(```yaml[ \t]+facts-machine[ \t]*\n)(.*?)(\n```)", re.DOTALL)


def find_config(argv):
    for a in argv:
        if a.endswith(".yaml") or a.endswith(".yml"):
            return Path(a).resolve()
    for cand in (Path("技术标/chapters.yaml"), Path("chapters.yaml")):
        if cand.exists():
            return cand.resolve()
    sys.exit("错误: 未找到 chapters.yaml,请显式传入路径。")


def remap_num(numstr, imap):
    head, dot, tail = numstr.partition(".")
    if not head.isdigit() or int(head) not in imap:
        return numstr
    return str(imap[int(head)]) + (dot + tail if dot else "")


def xform_chapter_md(text, imap, caption_prefix):
    n = [0]

    def h1_sub(m):
        new = remap_num(m.group(2), imap)
        if new == m.group(2):
            return m.group(0)
        n[0] += 1
        return f"{m.group(1)}{new}{m.group(3)}"

    def cap_sub(m):
        new = remap_num(m.group(2), imap)
        if new == m.group(2):
            return m.group(0)
        n[0] += 1
        return f"{m.group(1)}{new}{m.group(3)}"

    lines = text.split("\n")
    for i, ln in enumerate(lines):
        if ln.startswith("# "):  # 只动 H1;H2+ 由 number_headings 从新 base 派生
            lines[i] = re.sub(r"^(#\s+)(\d+(?:\.\d+)*)(\s.*)$", h1_sub, ln)
    text = "\n".join(lines)
    text = re.sub(
        r"(" + re.escape(caption_prefix) + r"\s*)(\d+(?:\.\d+)*)(-\d+)", cap_sub, text
    )
    return text, n[0]


def xform_facts(text, imap):
    def block_sub(bm):
        body = re.sub(
            r'"(\d+(?:\.\d+)*)"(\s*):',
            lambda m: f'"{remap_num(m.group(1), imap)}"{m.group(2)}:',
            bm.group(2),
        )
        return bm.group(1) + body + bm.group(3)

    return FACTS_FENCE.sub(block_sub, text)


def xform_outline(text, imap):
    lines = text.split("\n")
    for i, ln in enumerate(lines):
        if ln.startswith(">"):  # 引语/说明行(语义)人工维护,引擎不碰
            continue
        s = ln
        s = re.sub(r"第(\d+)章", lambda m: f"第{remap_num(m.group(1), imap)}章", s)
        s = re.sub(r"ch(\d+(?:\.\d+)*)", lambda m: f"ch{remap_num(m.group(1), imap)}", s)
        s = re.sub(
            r"^(\s*(?:#{2,}\s+|-\s+))(\d+(?:\.\d+)*)(\s|　)",
            lambda m: f"{m.group(1)}{remap_num(m.group(2), imap)}{m.group(3)}",
            s,
        )
        s = re.sub(
            r"(\d+)—(\d+)",
            lambda m: f"{remap_num(m.group(1), imap)}—{remap_num(m.group(2), imap)}",
            s,
        )
        lines[i] = s
    return "\n".join(lines)


def new_leading(name_pat, name, imap):
    m = re.match(name_pat, name)
    if not m:
        return None
    new = remap_num(m.group(1), imap)
    if new == m.group(1):
        return None
    return new, m


def main():
    apply = "--apply" in sys.argv
    cfg_path = find_config(sys.argv[1:])
    cn = ChapterNumbering(cfg_path)
    root = cn.root
    t = cn.targets()
    ch_dir = (root / t["chapters_glob"]).parent
    imap = cn.integer_map(ch_dir)
    identity = all(k == v for k, v in imap.items())
    print(f"=== 章号位移引擎(通用) ({'APPLY' if apply else 'DRY-RUN'}) · {cfg_path} ===")
    print(f"number_base={cn.load()['number_base']} · 整数章号映射: {dict(sorted(imap.items()))}")
    if identity:
        print("磁盘已与 config 一致,无需位移 (no-op)。")
        return

    # ① + ② 章节 md 内容 + 文件重命名
    md_files = sorted(ch_dir.glob(Path(t["chapters_glob"]).name))
    print(f"\n[md] {len(md_files)} 个章节文件:")
    for f in md_files:
        text = f.read_text(encoding="utf-8")
        new_text, nchg = xform_chapter_md(text, imap, t["caption_prefix"])
        r = new_leading(r"^ch(\d+(?:\.\d+)*)-", f.name, imap)
        nn = (f"ch{r[0]}-" + f.name[r[1].end():]) if r else None
        print(f"  {f.name}  H1/题注×{nchg}" + (f"  → {nn}" if nn else ""))
        if apply:
            if new_text != text:
                f.write_text(new_text, encoding="utf-8")
            if nn and nn != f.name:
                subprocess.run(["git", "mv", f.name, nn], cwd=str(ch_dir), capture_output=True)
                if (ch_dir / f.name).exists():
                    os.rename(ch_dir / f.name, ch_dir / nn)

    # ③ 成图 PNG 两段式改名
    png_dir = (root / t["figure_png_glob"]).parent
    pngs = sorted(png_dir.glob(Path(t["figure_png_glob"]).name)) if png_dir.is_dir() else []
    renames = []
    for p in pngs:
        r = new_leading(r"^图(\d+(?:\.\d+)*)(-\d+_)", p.name, imap)
        if r:
            renames.append((p, f"图{r[0]}" + p.name[r[1].start(2):]))
    print(f"\n[png] {len(renames)}/{len(pngs)} 个成图改名:")
    for p, nn in renames:
        print(f"  {p.name}  → {nn}")
    if apply and renames:
        for p, _ in renames:
            os.rename(p, p.with_name("__tmp__" + p.name))
        for p, nn in renames:
            os.rename(p.with_name("__tmp__" + p.name), png_dir / nn)

    # ④ FACTS
    facts = root / t["facts_file"]
    if facts.exists():
        ftxt = facts.read_text(encoding="utf-8")
        fnew = xform_facts(ftxt, imap)
        print(f"\n[FACTS] facts-machine 键位移: {'有改动' if fnew != ftxt else '无'}")
        if apply and fnew != ftxt:
            facts.write_text(fnew, encoding="utf-8")

    # ⑤ 目录大纲
    for outline in sorted(root.glob(t["outline_glob"])):
        otxt = outline.read_text(encoding="utf-8")
        onew = xform_outline(otxt, imap)
        print(f"[大纲] {outline.name} 章号位移: {'有改动' if onew != otxt else '无'}")
        if apply and onew != otxt:
            outline.write_text(onew, encoding="utf-8")

    print(f"\n{'✅ 已执行' if apply else '（干跑,加 --apply 执行）'}")
    print("后续: 项目侧 number_headings.py --apply(重派子号) → 重渲 docx。")
    print("提示: 引擎不碰散文类 SSOT(CLAUDE.md/总纲/大纲引语行),需人工同步语义。")


if __name__ == "__main__":
    main()
