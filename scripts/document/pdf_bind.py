#!/usr/bin/env python3
"""pdf_bind.py —— 把一批 PDF 按「篇 / 章」结构装订成一本正式册子。

产出形制（每一段都可选）：

    封面            不编页码
    前置件          保留原件自带页码，不叠印
    目录            页码 = 1，自动回填各篇 / 各章起始页
    篇扉页          「一、水量评估」
    章扉页          「1 供水保障程度」
    正文            该章各份材料依次拼接
    ……
    尾件            正文之后的附加件（可带自己的扉页）

全篇连续页码从目录页起算，**旋转感知**：/Rotate 90/180/270 的页
（扫描件里很常见）页码照样印在视觉底边中央且正着读，不会跑到侧边或倒印。

为什么不是 `pdfunite` / `pdftk` 一条命令：那些只做 concat，不生成扉页、
不叠页码、不回填目录页码。这三件才是「装订成册」与「拼在一起」的区别。

依赖：pypdf、reportlab（CJK 走内置 STSong-Light CID 字体，无需外部字体文件）、
      Pillow（仅当输入含图片时）。

用法（库）：

    from pdf_bind import Chapter, Part, bind
    bind(out_path=...,
         cover=Path('封面.pdf'),
         front=[Path('省厅通知.pdf')],
         parts=[Part('一、水量评估', [Chapter('1 供水保障程度', [...])])],
         tail=[Chapter('自评估表', [Path('z2 自评估表.pdf')])])

用法（CLI，接 JSON 配方）：

    python3 pdf_bind.py --recipe recipe.json --out 汇编.pdf
"""

from __future__ import annotations

import argparse
import io
import json
import math
import re
import sys
from dataclasses import dataclass, field
from pathlib import Path
from typing import Iterable, Sequence

import pypdf
from reportlab.lib.pagesizes import A4
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.cidfonts import UnicodeCIDFont
from reportlab.pdfgen import canvas as rl_canvas

# ---------------------------------------------------------------------------
# 字体：reportlab 内置的 Adobe CJK CID 字体，装 reportlab 就有，不依赖系统字体
# ---------------------------------------------------------------------------
CJK_FONT = "STSong-Light"
_FONT_READY = False


def ensure_font() -> str:
    global _FONT_READY
    if not _FONT_READY:
        pdfmetrics.registerFont(UnicodeCIDFont(CJK_FONT))
        _FONT_READY = True
    return CJK_FONT


# ---------------------------------------------------------------------------
# 结构
# ---------------------------------------------------------------------------
@dataclass
class Chapter:
    """一章 = 一个扉页 + 若干份材料。title 为 None 时不出扉页。"""

    title: str | None
    files: list[Path] = field(default_factory=list)


@dataclass
class Part:
    """一篇 = 一个扉页 + 若干章。title 为 None 时不出扉页。"""

    title: str | None
    chapters: list[Chapter] = field(default_factory=list)


@dataclass
class _Entry:
    """目录里的一行。level 0 = 篇，1 = 章。"""

    level: int
    title: str
    page: int  # 连续页码（目录页 = 1）


# ---------------------------------------------------------------------------
# 排版参数（默认值对齐水源地汇编的既有实物）
# ---------------------------------------------------------------------------
@dataclass
class Style:
    page_size: tuple[float, float] = A4
    part_font_size: float = 24
    chapter_font_size: float = 22
    toc_title_size: float = 18
    toc_entry_size: float = 12
    toc_part_size: float = 13
    page_number_size: float = 10.5
    page_number_margin: float = 30  # 页码基线距视觉底边（pt）
    toc_top_margin: float = 90
    toc_bottom_margin: float = 70
    toc_side_margin: float = 70
    toc_line_gap: float = 21


# ---------------------------------------------------------------------------
# 单页生成
# ---------------------------------------------------------------------------
def _center_text_page(text: str, size: float, st: Style) -> bytes:
    """一张只有一行居中大字的扉页。"""
    ensure_font()
    buf = io.BytesIO()
    w, h = st.page_size
    c = rl_canvas.Canvas(buf, pagesize=st.page_size)
    c.setFont(CJK_FONT, size)
    c.drawCentredString(w / 2, h / 2, text)
    c.showPage()
    c.save()
    return buf.getvalue()


def _toc_pages(entries: Sequence[_Entry], st: Style, title: str = "目录") -> bytes:
    """目录页（可能多页）。带点导线，右对齐页码。"""
    ensure_font()
    buf = io.BytesIO()
    w, h = st.page_size
    c = rl_canvas.Canvas(buf, pagesize=st.page_size)

    left = st.toc_side_margin
    right = w - st.toc_side_margin
    y = h - st.toc_top_margin
    first = True

    for e in entries:
        if first:
            c.setFont(CJK_FONT, st.toc_title_size)
            c.drawCentredString(w / 2, y, title)
            y -= st.toc_line_gap * 2
            first = False
        if y < st.toc_bottom_margin:
            c.showPage()
            y = h - st.toc_top_margin

        size = st.toc_part_size if e.level == 0 else st.toc_entry_size
        indent = left + (0 if e.level == 0 else 18)
        c.setFont(CJK_FONT, size)
        c.drawString(indent, y, e.title)

        num = str(e.page)
        num_w = pdfmetrics.stringWidth(num, CJK_FONT, size)
        c.drawString(right - num_w, y, num)

        # 点导线
        title_w = pdfmetrics.stringWidth(e.title, CJK_FONT, size)
        dot_start = indent + title_w + 4
        dot_end = right - num_w - 4
        if dot_end > dot_start:
            dot_w = pdfmetrics.stringWidth(".", CJK_FONT, size)
            n = int((dot_end - dot_start) / dot_w)
            if n > 0:
                c.drawString(dot_start, y, "." * n)
        y -= st.toc_line_gap

    c.showPage()
    c.save()
    return buf.getvalue()


def _image_to_pdf(img_path: Path, st: Style) -> bytes:
    """图片 → 单页 PDF（等比缩放居中留白）。不转就会在汇编里整份丢失。"""
    from PIL import Image

    ensure_font()
    buf = io.BytesIO()
    pw, ph = st.page_size
    with Image.open(img_path) as im:
        iw, ih = im.size
    margin = 36
    scale = min((pw - 2 * margin) / iw, (ph - 2 * margin) / ih)
    dw, dh = iw * scale, ih * scale
    c = rl_canvas.Canvas(buf, pagesize=st.page_size)
    c.drawImage(str(img_path), (pw - dw) / 2, (ph - dh) / 2, width=dw, height=dh,
                preserveAspectRatio=True, anchor="c")
    c.showPage()
    c.save()
    return buf.getvalue()


# ---------------------------------------------------------------------------
# 页码叠印（旋转感知）
# ---------------------------------------------------------------------------
def _number_overlay_doc(specs: Sequence[tuple[float, float, float, float, int, int]],
                        st: Style) -> pypdf.PdfReader:
    """一次性生成所有页码叠印层。

    specs 每项 = (x0, y0, w, h, rotation, number)，其中 x0/y0 是 mediabox 原点
    （不一定是 0,0），rotation 取 0/90/180/270。

    旋转的处理：叠印层与被叠页在**同一个未旋转坐标系**里合并，显示时一起被
    /Rotate 转。所以要让页码显示在视觉底边且正着读，得反推它在未旋转空间
    里的位置和角度：

        rot   视觉底边 = 原始的哪条边      文字在 PDF 空间的转角
        0     下边 (y=y0)                  0
        90    右边 (x=x0+w)                90
        180   上边 (y=y0+h)                180
        270   左边 (x=x0)                  270
    """
    ensure_font()
    buf = io.BytesIO()
    c = rl_canvas.Canvas(buf)
    m = st.page_number_margin
    for x0, y0, w, h, rot, num in specs:
        c.setPageSize((w, h))
        c.setFont(CJK_FONT, st.page_number_size)
        text = str(num)
        c.saveState()
        # 注：drawCentredString 以当前坐标系原点为基准，故先平移再旋转
        if rot == 90:
            c.translate(x0 + w - m, y0 + h / 2)
            c.rotate(90)
        elif rot == 180:
            c.translate(x0 + w / 2, y0 + h - m)
            c.rotate(180)
        elif rot == 270:
            c.translate(x0 + m, y0 + h / 2)
            c.rotate(270)
        else:
            c.translate(x0 + w / 2, y0 + m)
        c.drawCentredString(0, 0, text)
        c.restoreState()
        c.showPage()
    c.save()
    buf.seek(0)
    return pypdf.PdfReader(buf)


# ---------------------------------------------------------------------------
# 输入
# ---------------------------------------------------------------------------
IMAGE_SUFFIXES = {".png", ".jpg", ".jpeg", ".tif", ".tiff", ".bmp", ".gif"}


def _reader_for(path: Path, st: Style) -> pypdf.PdfReader:
    if path.suffix.lower() in IMAGE_SUFFIXES:
        return pypdf.PdfReader(io.BytesIO(_image_to_pdf(path, st)))
    return pypdf.PdfReader(str(path))


def natural_key(p: Path):
    """按材料编号自然排序：1-2 在 1-10 前面，纯字典序会反过来。"""
    stem = p.name
    m = re.match(r"^\s*([0-9]+(?:[-.][0-9]+)*)", stem)
    if not m:
        return (1, [], stem)
    nums = [int(x) for x in re.split(r"[-.]", m.group(1))]
    return (0, nums, stem)


# ---------------------------------------------------------------------------
# 主流程
# ---------------------------------------------------------------------------
def bind(
    out_path: Path,
    cover: Path | None = None,
    front: Sequence[Path] = (),
    parts: Sequence[Part] = (),
    tail: Sequence[Chapter] = (),
    toc_title: str = "目录",
    with_toc: bool = True,
    style: Style | None = None,
    verbose: bool = True,
) -> dict:
    """装订并落盘，返回统计字典。"""
    st = style or Style()
    out_path = Path(out_path)

    def log(msg: str) -> None:
        if verbose:
            print(msg, flush=True)

    # --- 1. 先量出正文各段页数，才能算目录页码 -----------------------------
    # 段 = ('part-title'|'chapter-title'|'file', 载荷, 页数)
    body: list[tuple[str, object, int]] = []
    entries_plan: list[tuple[int, str, int]] = []  # (level, title, body 内的段序号)

    def push_title(level: int, text: str) -> None:
        entries_plan.append((level, text, len(body)))
        body.append(("title", (text, st.part_font_size if level == 0
                               else st.chapter_font_size), 1))

    page_counts: dict[Path, int] = {}

    def push_file(f: Path) -> None:
        r = _reader_for(f, st)
        n = len(r.pages)
        page_counts[f] = n
        body.append(("file", f, n))

    for part in parts:
        if part.title:
            push_title(0, part.title)
        for ch in part.chapters:
            if ch.title:
                push_title(1, ch.title)
            for f in ch.files:
                push_file(f)
    for ch in tail:
        if ch.title:
            push_title(0, ch.title)
        for f in ch.files:
            push_file(f)

    body_pages = sum(n for _, _, n in body)
    log(f"[bind] 正文 {body_pages} 页（扉页 {sum(1 for k, _, _ in body if k == 'title')} 张）")

    # --- 2. 目录页数：先按估算排一遍，稳定后定下 ---------------------------
    toc_page_count = 0
    if with_toc and entries_plan:
        usable = st.page_size[1] - st.toc_top_margin - st.toc_bottom_margin
        per_page = max(1, int(usable / st.toc_line_gap))
        # 首页被标题占去两行
        toc_page_count = max(1, math.ceil((len(entries_plan) + 2) / per_page))

    # --- 3. 回填页码：目录页 = 1 -------------------------------------------
    offsets: list[int] = []  # body 段序号 → 该段首页的连续页码
    cur = toc_page_count + 1
    for _, _, n in body:
        offsets.append(cur)
        cur += n
    total_numbered = toc_page_count + body_pages

    entries = [_Entry(level, title, offsets[idx]) for level, title, idx in entries_plan]

    # --- 4. 组装 ------------------------------------------------------------
    writer = pypdf.PdfWriter()
    numbered_specs: list[tuple[float, float, float, float, int, int]] = []
    numbered_start_index = 0  # writer 里从第几页开始编号

    if cover is not None:
        for p in _reader_for(Path(cover), st).pages:
            writer.add_page(p)
    for f in front:
        for p in _reader_for(Path(f), st).pages:
            writer.add_page(p)
    numbered_start_index = len(writer.pages)
    log(f"[bind] 封面+前置件 {numbered_start_index} 页（不编页码）")

    if toc_page_count:
        toc_reader = pypdf.PdfReader(io.BytesIO(_toc_pages(entries, st, toc_title)))
        got = len(toc_reader.pages)
        if got != toc_page_count:
            # 估算与实排不符 → 用实排页数重算一次页码，再渲染一遍（收敛快，最多两轮）
            log(f"[bind] 目录实排 {got} 页 ≠ 估算 {toc_page_count} 页，重算页码")
            toc_page_count = got
            cur = toc_page_count + 1
            offsets = []
            for _, _, n in body:
                offsets.append(cur)
                cur += n
            total_numbered = toc_page_count + body_pages
            entries = [_Entry(l, t, offsets[i]) for l, t, i in entries_plan]
            toc_reader = pypdf.PdfReader(io.BytesIO(_toc_pages(entries, st, toc_title)))
        for p in toc_reader.pages:
            writer.add_page(p)

    for kind, payload, _n in body:
        if kind == "title":
            text, size = payload  # type: ignore[misc]
            r = pypdf.PdfReader(io.BytesIO(_center_text_page(text, size, st)))
        else:
            r = _reader_for(payload, st)  # type: ignore[arg-type]
        for p in r.pages:
            writer.add_page(p)

    # --- 5. 叠页码 ----------------------------------------------------------
    for i in range(numbered_start_index, len(writer.pages)):
        page = writer.pages[i]
        mb = page.mediabox
        numbered_specs.append((
            float(mb.left), float(mb.bottom),
            float(mb.width), float(mb.height),
            page.rotation % 360,
            i - numbered_start_index + 1,
        ))
    log(f"[bind] 叠页码 {len(numbered_specs)} 页（1..{total_numbered}）")
    ov = _number_overlay_doc(numbered_specs, st)
    for k, i in enumerate(range(numbered_start_index, len(writer.pages))):
        writer.pages[i].merge_page(ov.pages[k])

    out_path.parent.mkdir(parents=True, exist_ok=True)
    with out_path.open("wb") as fh:
        writer.write(fh)

    stats = {
        "out": str(out_path),
        "total_pages": len(writer.pages),
        "unnumbered_pages": numbered_start_index,
        "numbered_pages": len(numbered_specs),
        "toc_pages": toc_page_count,
        "title_pages": sum(1 for k, _, _ in body if k == "title"),
        "source_files": len(page_counts),
        "source_pages": sum(page_counts.values()),
        "entries": [(e.level, e.title, e.page) for e in entries],
        "size_bytes": out_path.stat().st_size,
    }
    log(f"[bind] 完成 {out_path} · {stats['total_pages']} 页 · "
        f"{stats['size_bytes'] / 1048576:.0f} MB")
    return stats


# ---------------------------------------------------------------------------
# CLI
# ---------------------------------------------------------------------------
def _from_recipe(recipe: dict) -> dict:
    def P(x):
        return Path(x)

    parts = [Part(p.get("title"), [Chapter(c.get("title"), [P(f) for f in c.get("files", [])])
                                   for c in p.get("chapters", [])])
             for p in recipe.get("parts", [])]
    tail = [Chapter(c.get("title"), [P(f) for f in c.get("files", [])])
            for c in recipe.get("tail", [])]
    return dict(
        cover=P(recipe["cover"]) if recipe.get("cover") else None,
        front=[P(f) for f in recipe.get("front", [])],
        parts=parts,
        tail=tail,
        toc_title=recipe.get("toc_title", "目录"),
        with_toc=recipe.get("with_toc", True),
    )


def main(argv: list[str] | None = None) -> int:
    ap = argparse.ArgumentParser(description="把一批 PDF 按篇/章结构装订成册")
    ap.add_argument("--recipe", required=True, type=Path, help="JSON 配方")
    ap.add_argument("--out", required=True, type=Path, help="产物路径")
    ap.add_argument("--stats", type=Path, help="把统计写成 JSON")
    ap.add_argument("--quiet", action="store_true")
    args = ap.parse_args(argv)

    if not args.recipe.exists():
        print(f"⛔ 配方不存在：{args.recipe}", file=sys.stderr)
        return 2
    recipe = json.loads(args.recipe.read_text(encoding="utf-8"))
    kw = _from_recipe(recipe)
    if not kw["parts"] and not kw["tail"]:
        print("⛔ 配方里 parts / tail 都是空的，拒绝在空集上产出一本空册子", file=sys.stderr)
        return 2

    stats = bind(args.out, verbose=not args.quiet, **kw)
    if args.stats:
        args.stats.write_text(json.dumps(stats, ensure_ascii=False, indent=2), encoding="utf-8")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
