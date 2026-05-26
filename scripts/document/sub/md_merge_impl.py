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

import sys
import shutil
from pathlib import Path
from docx import Document


def parse_md(filepath: str) -> list[tuple[str, str]]:
    """解析 MD 为 [(style_name, text)] 列表。

    只处理 ##/###/####/##### 标题和普通段落（空行跳过）。
    Returns list of (Word style name, text) tuples.
    """
    paragraphs: list[tuple[str, str]] = []
    with open(filepath, "r", encoding="utf-8") as f:
        for line in f:
            line = line.rstrip("\n")
            if line.startswith("##### "):
                paragraphs.append(("Heading 5", line[6:]))
            elif line.startswith("#### "):
                paragraphs.append(("Heading 4", line[5:]))
            elif line.startswith("### "):
                paragraphs.append(("Heading 3", line[4:]))
            elif line.startswith("## "):
                paragraphs.append(("Heading 2", line[3:]))
            elif line.strip() == "":
                continue
            else:
                paragraphs.append(("Normal", line))
    return paragraphs


def apply(
    md_file: str,
    docx_file: str,
    start_idx: int,
    end_idx: int,
    output_file: str | None = None,
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
    if output_file is None:
        p = Path(docx_file)
        output_file = str(p.parent / (p.stem + "-merged" + p.suffix))

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

    md_paras = parse_md(md_file)

    # If MD starts with a heading, update the section heading text
    if md_paras and md_paras[0][0].startswith("Heading"):
        title_text = md_paras[0][1]
        p_start = doc.paragraphs[start_idx]
        for run in p_start.runs:
            run.text = ""
        if p_start.runs:
            p_start.runs[0].text = title_text
        md_paras = md_paras[1:]
        print(f"标题更新为: {title_text}")

    print(f"插入 {len(md_paras)} 段新内容")

    # Insert MD paragraphs after the section heading
    ref = start_elem
    for style, text in md_paras:
        new_p = doc.add_paragraph(text, style=style)
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
    if len(sys.argv) < 5:
        print(__doc__)
        sys.exit(1)

    md_file = sys.argv[1]
    docx_file = sys.argv[2]
    start_idx = int(sys.argv[3])
    end_idx = int(sys.argv[4])
    output_file = sys.argv[5] if len(sys.argv) > 5 else None

    apply(md_file, docx_file, start_idx, end_idx, output_file)
    return 0


if __name__ == "__main__":
    sys.exit(main())
