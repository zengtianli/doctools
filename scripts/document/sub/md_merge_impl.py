#!/usr/bin/env python3
"""md_merge_impl.py — 将 MD 内容安全替换到 DOCX 指定章节

Distilled from panan-rigid-2026/scripts/merge_md_to_docx.py (A级通用, 2026-05-26).
零业务硬编码; 纯 python-docx XML 操作.

用法:
    python3 md_merge_impl.py <md_file> <docx_file> <start_idx> <end_idx> [output_file]

参数:
    md_file     : 要并入的 MD 文件路径
    docx_file   : 目标 DOCX 文件路径
    start_idx   : 替换起始段落索引（Heading 段落，会保留并更新标题）
    end_idx     : 替换结束段落索引（不含，即下一个章节的 Heading）
    output_file : 可选，输出文件路径，默认在 docx 同目录加 -merged 后缀

安全保证:
    - 只删除 w:p 段落元素
    - 保留 w:tbl 表格和其他非段落 XML 元素
    - 表格按前导段落文本锚点回插

触发场景:
    - 把 MD 内容合入 docx 某章节（知道起止段落索引）
    - 用新写的 MD 草稿替换 Word 交付物已有章节内容
    - 配合 `section read-section --list` 先确认段落索引再合入
"""
from __future__ import annotations

import argparse
import os
import re
import sys
import shutil
from datetime import datetime
from pathlib import Path
from docx import Document
from docx.shared import Inches

# 复用 md_docx_template 的 md-table 解析/边框 helper, 不重写 (铁律 #5)
sys.path.insert(0, str(Path(__file__).resolve().parent.parent))
from md_docx_template import (  # noqa: E402
    parse_table_row,
    is_separator_row,
    set_table_border,
    clean_markdown_text,
)


def parse_md(filepath: str) -> list:
    """解析 MD 为 block 列表,支持 标题 / 段落 / markdown 表格→w:tbl。

    block 形态:
      ("heading", level:int, text)     —— ## .. ##### → level 2..5
      ("para", text)                   —— 普通段落(经 clean_markdown_text 去 **bold**/`code`/$math$ 语法噪音)
      ("table", headers:list, rows:list[list]) —— markdown 表格(管道符 + 分隔行)

    2026-05-29 (GOAL report-automation Phase 0-A): 修复原 parse_md 只认标题+段、
    markdown 表格静默写成字面管道符 Normal 段的 🔴 缺口。表格解析复用 md_docx_template
    的 parse_table_row/is_separator_row(铁律 #5 不重写)。
    """
    blocks: list = []
    with open(filepath, "r", encoding="utf-8") as f:
        lines = f.read().split("\n")
    i, n = 0, len(lines)
    while i < n:
        line = lines[i].rstrip()
        stripped = line.strip()
        if stripped == "":
            i += 1
            continue
        # 表格: 连续 | 开头行 (≥2 行: 表头 + 分隔 + 数据)
        if stripped.startswith("|"):
            tbl_lines = []
            while i < n and lines[i].strip().startswith("|"):
                tbl_lines.append(lines[i])
                i += 1
            if len(tbl_lines) >= 2:
                headers = [clean_markdown_text(c) for c in parse_table_row(tbl_lines[0])]
                rows = [
                    [clean_markdown_text(c) for c in parse_table_row(tl)]
                    for tl in tbl_lines[2:]
                    if not is_separator_row(tl)
                ]
                blocks.append(("table", headers, rows))
            else:
                for tl in tbl_lines:  # 退化: 不足 2 行不成表 → 普通段落
                    blocks.append(("para", clean_markdown_text(tl.strip())))
            continue
        # 图片 ![alt](path) — 路径相对 md 文件目录解析
        img_m = re.match(r"^!\[(.*?)\]\((.+?)\)\s*$", stripped)
        if img_m:
            alt, path = img_m.group(1), img_m.group(2).strip()
            if not os.path.isabs(path):
                path = os.path.normpath(
                    os.path.join(os.path.dirname(os.path.abspath(filepath)), path)
                )
            blocks.append(("image", path, alt))
            i += 1
            continue
        # 标题 ##..#####
        m = re.match(r"^(#{2,5})\s+(.+)$", line)
        if m:
            blocks.append(("heading", len(m.group(1)), m.group(2).strip()))
            i += 1
            continue
        # 普通段落
        blocks.append(("para", clean_markdown_text(stripped)))
        i += 1
    return blocks


def resolve_anchor(doc, anchor: str, *, from_idx: int = 0) -> int:
    """返回 from_idx 起首个 stripped 文本 startswith/contains anchor 的段落 idx。

    未匹配 → ValueError。用于 --start-anchor/--end-anchor 省掉手查索引那步。
    """
    norm = anchor.strip()
    for i in range(from_idx, len(doc.paragraphs)):
        t = (doc.paragraphs[i].text or "").strip()
        if t.startswith(norm) or norm in t:
            return i
    raise ValueError(f"anchor 未匹配: {anchor!r}")


def apply(
    md_file: str,
    docx_file: str,
    start_idx: int,
    end_idx: int,
    output_file: str | None = None,
    in_place: bool = False,
    no_backup: bool = False,
) -> str:
    """Merge MD into DOCX section and return output path.

    Args:
        md_file: Path to source Markdown file.
        docx_file: Path to target DOCX file.
        start_idx: Index of section heading paragraph (kept, title updated from MD).
        end_idx: Index of next section heading (exclusive; content up to here replaced).
        output_file: Optional output path. Defaults to <docx>-merged.docx.

    Returns:
        Absolute path of saved output file.
    """
    if in_place:
        if not no_backup:
            bak = f"{docx_file}.bak-{datetime.now().strftime('%Y%m%d-%H%M%S')}"
            shutil.copy2(docx_file, bak)
            print(f"已备份: {bak}")
        output_file = docx_file
    elif output_file is None:
        p = Path(docx_file)
        output_file = str(p.parent / (p.stem + "-merged" + p.suffix))

    if output_file != docx_file:
        shutil.copy2(docx_file, output_file)
    doc = Document(output_file)
    body = doc.element.body

    start_elem = doc.paragraphs[start_idx]._element
    end_elem = doc.paragraphs[end_idx]._element

    print(f"替换范围: 段落[{start_idx}]~[{end_idx - 1}]")
    print(f"  起始: {doc.paragraphs[start_idx].text[:60]}")
    print(f"  结束前: {doc.paragraphs[end_idx - 1].text[:60]}")
    print(f"  下一章: {doc.paragraphs[end_idx].text[:60]}")

    # Collect children between start and end (exclusive)
    children = []
    in_range = False
    for child in list(body):
        if child is start_elem:
            in_range = True
            continue  # keep start (section heading)
        if child is end_elem:
            break
        if in_range:
            children.append(child)

    # Classify: paragraphs vs non-paragraph elements (tables etc.)
    paras = []
    tables: list[tuple] = []  # (element, anchor_text)
    last_para_text = ""
    for child in children:
        tag = child.tag.split("}")[-1]
        if tag == "p":
            last_para_text = "".join(child.itertext()).strip()[:80]
            paras.append(child)
        else:
            tables.append((child, last_para_text))

    print(f"\n范围内: {len(paras)} 段落, {len(tables)} 个非段落元素")

    # Detach non-paragraph elements then delete paragraphs
    for elem, _ in tables:
        body.remove(elem)
    for p_elem in paras:
        body.remove(p_elem)

    blocks = parse_md(md_file)

    # If MD starts with a heading, update the section heading text
    if blocks and blocks[0][0] == "heading":
        title_text = blocks[0][2]
        p_start = doc.paragraphs[start_idx]
        for run in p_start.runs:
            run.text = ""
        if p_start.runs:
            p_start.runs[0].text = title_text
        blocks = blocks[1:]
        print(f"标题更新为: {title_text}")

    n_tbl = sum(1 for b in blocks if b[0] == "table")
    print(f"插入 {len(blocks)} 个 block (含 {n_tbl} 个表格)")

    # Insert blocks after the section heading (para/heading/table)
    ref = start_elem
    for blk in blocks:
        if blk[0] == "heading":
            new_p = doc.add_paragraph(blk[2], style=f"Heading {min(blk[1], 5)}")
            new_elem = new_p._element
        elif blk[0] == "table":
            headers, rows = blk[1], blk[2]
            ncols = len(headers)
            table = doc.add_table(rows=1 + len(rows), cols=ncols)
            set_table_border(table)
            for j, h in enumerate(headers):
                table.rows[0].cells[j].text = h
            for ri, rd in enumerate(rows):
                for j, ct in enumerate(rd):
                    if j < ncols:
                        table.rows[ri + 1].cells[j].text = ct
            new_elem = table._tbl  # add_table 追加在 body 末; addnext 移到 anchor 后
        elif blk[0] == "image":
            img_path, _alt = blk[1], blk[2]
            new_p = doc.add_paragraph(style="Normal")
            if os.path.exists(img_path):
                pic = new_p.add_run().add_picture(img_path)
                maxw = Inches(5.8)  # 超过正文宽则等比缩放
                if pic.width > maxw:
                    ratio = maxw / pic.width
                    pic.width = int(pic.width * ratio)
                    pic.height = int(pic.height * ratio)
            else:
                new_p.add_run(f"[图片缺失: {img_path}]")
                print(f"WARN: 图片不存在, 插占位文本: {img_path}")
            new_elem = new_p._element
        else:  # para
            new_p = doc.add_paragraph(blk[1], style="Normal")
            new_elem = new_p._element
        ref.addnext(new_elem)
        ref = new_elem

    # Reinsert non-paragraph elements by anchor text matching
    for tbl_elem, anchor in tables:
        inserted = False
        if anchor:
            for child in list(body):
                if child.tag.split("}")[-1] == "p":
                    p_text = "".join(child.itertext()).strip()[:80]
                    if anchor and anchor in p_text:
                        child.addnext(tbl_elem)
                        print(f'  非段落元素回插到 "{anchor[:40]}" 之后')
                        inserted = True
                        break
        if not inserted:
            ref.addnext(tbl_elem)
            ref = tbl_elem
            print(f'  非段落元素插到章末（锚点 "{anchor[:40]}" 未匹配）')

    doc.save(output_file)
    print(f"\n已保存到 {output_file}")

    # Verification pass
    doc2 = Document(output_file)
    print("\n--- 验证 ---")
    for i in range(
        max(0, start_idx - 1),
        min(start_idx + len(md_paras) + 5, len(doc2.paragraphs)),
    ):
        p = doc2.paragraphs[i]
        if "Heading" in p.style.name:
            print(f"  [{i}] ({p.style.name}) {p.text[:80]}")

    return output_file


# Alias for pipeline adapter compatibility
def apply_path(docx_path=None, args=None) -> dict:
    """pipeline adapter — delegates to apply(); requires positional args via sys.argv."""
    try:
        main()
        return {"status": "ok", "script": "md_merge_impl.py"}
    except SystemExit as e:
        return {"status": "sysexit", "code": e.code, "script": "md_merge_impl.py"}
    except Exception as e:
        return {"status": "error", "error": repr(e), "script": "md_merge_impl.py"}


def main() -> int:
    ap = argparse.ArgumentParser(
        prog="md-merge",
        description="用 MD 内容替换 DOCX 指定节 (表格 anchor 安全回插)。"
                    "支持位置索引 或 --start-anchor/--end-anchor 按标题定位；"
                    "--in-place 原地改 + 自动 .bak-时间戳 (Work §1.5 协议)。",
    )
    ap.add_argument("md_file", help="源 Markdown")
    ap.add_argument("docx_file", help="目标 DOCX")
    ap.add_argument("start_idx", nargs="?", type=int, help="起始段落 idx (节标题段)")
    ap.add_argument("end_idx", nargs="?", type=int, help="结束段落 idx (不含, 下一节标题)")
    ap.add_argument("output_file", nargs="?", help="输出路径 (默认 <docx>-merged.docx; --in-place 时忽略)")
    ap.add_argument("--start-anchor", help="按标题文本定位 start_idx (替代位置参数)")
    ap.add_argument("--end-anchor", help="按标题文本定位 end_idx (下一节标题)")
    ap.add_argument("--in-place", action="store_true", help="原地改 + 自动备份 .bak-时间戳")
    ap.add_argument("--no-backup", action="store_true", help="配合 --in-place 跳过备份")
    args = ap.parse_args()

    start_idx, end_idx = args.start_idx, args.end_idx
    if args.start_anchor or args.end_anchor:
        doc = Document(args.docx_file)
        if args.start_anchor:
            start_idx = resolve_anchor(doc, args.start_anchor)
        if args.end_anchor:
            end_idx = resolve_anchor(doc, args.end_anchor, from_idx=(start_idx or 0) + 1)
        print(f"锚点解析: start_idx={start_idx} end_idx={end_idx}")

    if start_idx is None or end_idx is None:
        print("[md-merge] 需位置参数 start_idx/end_idx, 或 --start-anchor/--end-anchor", file=sys.stderr)
        return 2

    apply(args.md_file, args.docx_file, start_idx, end_idx,
          args.output_file, in_place=args.in_place, no_backup=args.no_backup)
    return 0


if __name__ == "__main__":
    sys.exit(main())
