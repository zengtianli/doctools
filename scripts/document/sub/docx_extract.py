#!/usr/bin/env python3
"""docx extract — 从 .docx 提取纯文本（Markdown 格式），支持分章节输出。

docx_tools.py（1705 行单体）2026-07-31 P2 拆解产物之一：本模块是 extract 段的
实现层（纯 zipfile+lxml 只读）。docx_tools.py 仍是组合入口 + library re-export 面，
CLI 契约（子命令名/flag/默认值）以 docx_tools.py 为准、声明在本模块 add_extract_parser()
只写一遍。也可独立敲：

    python3 sub/docx_extract.py extract input.docx [--info|--json|--split-chapters] [-o OUT]
"""

import argparse
import json
import os
import re
import sys
import zipfile
from pathlib import Path

# ── surgical 收口：python-docx 存盘只重写点名的部件（炸开面 60→1）─────────────
import sys as _sys
from pathlib import Path as _Path
_sys.path.append(str(_Path(__file__).resolve().parents[3] / "lib"))
import docx_safe_save  # noqa: E402,F401  详见 lib/docx_safe_save.py

from lxml import etree  # noqa: E402

from docx_xml import qn  # noqa: E402


def extract_paragraphs(docx_path: str) -> list[dict]:
    """提取所有段落，返回 [{style, text, level}]"""
    paragraphs = []
    with zipfile.ZipFile(docx_path, "r") as zf:
        doc_xml = zf.read("word/document.xml")
        # 读取样式映射
        styles_map = {}
        if "word/styles.xml" in zf.namelist():
            styles_xml = zf.read("word/styles.xml")
            stree = etree.fromstring(styles_xml)
            for s in stree.iter(qn("w:style")):
                sid = s.get(qn("w:styleId"), "")
                name_elem = s.find(qn("w:name"))
                name = name_elem.get(qn("w:val"), sid) if name_elem is not None else sid
                styles_map[sid] = name

        tree = etree.fromstring(doc_xml)
        body = tree.find(qn("w:body"))
        if body is None:
            return paragraphs

        for para in body.iter(qn("w:p")):
            # 提取样式
            ppr = para.find(qn("w:pPr"))
            style_id = ""
            outline_level = -1
            if ppr is not None:
                ps = ppr.find(qn("w:pStyle"))
                if ps is not None:
                    style_id = ps.get(qn("w:val"), "")
                ol = ppr.find(qn("w:outlineLvl"))
                if ol is not None:
                    outline_level = int(ol.get(qn("w:val"), "-1"))

            style_name = styles_map.get(style_id, style_id)

            # 计算标题级别
            level = -1
            heading_match = re.match(r"[Hh]eading\s*(\d+)", style_name)
            if heading_match:
                level = int(heading_match.group(1))
            elif outline_level >= 0:
                level = outline_level + 1

            # 提取文本
            texts = []
            for t in para.iter(qn("w:t")):
                if t.text:
                    texts.append(t.text)
            text = "".join(texts).strip()

            if text:
                paragraphs.append(
                    {
                        "style": style_name,
                        "style_id": style_id,
                        "text": text,
                        "level": level,
                    }
                )

    return paragraphs


def paragraphs_to_markdown(paragraphs: list[dict]) -> str:
    """将段落列表转为 Markdown"""
    lines = []
    for p in paragraphs:
        style = p["style"].lower()
        text = p["text"]
        level = p["level"]

        if level >= 1:
            lines.append(f"\n{'#' * level} {text}\n")
        elif "表" in style or "图" in style or "caption" in style:
            lines.append(f"\n**[{text}]**\n")
        elif "题目" in style or "title" in style:
            lines.append(f"\n**{text}**\n")
        else:
            lines.append(f"\n{text}\n")

    return "\n".join(lines).strip() + "\n"


def split_by_chapters(paragraphs: list[dict]) -> list[dict]:
    """按一级标题拆分章节，返回 [{title, paragraphs, markdown}]"""
    chapters = []
    current = {"title": "前置内容", "paragraphs": []}

    for p in paragraphs:
        if p["level"] == 1:
            if current["paragraphs"]:
                current["markdown"] = paragraphs_to_markdown(current["paragraphs"])
                chapters.append(current)
            current = {"title": p["text"], "paragraphs": [p]}
        else:
            current["paragraphs"].append(p)

    if current["paragraphs"]:
        current["markdown"] = paragraphs_to_markdown(current["paragraphs"])
        chapters.append(current)

    return chapters


def document_info(paragraphs: list[dict]) -> str:
    """输出文档结构摘要"""
    lines = ["## 文档结构\n"]
    total_chars = sum(len(p["text"]) for p in paragraphs)
    lines.append(f"- 总段落数：{len(paragraphs)}")
    lines.append(f"- 总字符数：{total_chars}")
    lines.append("- 样式统计：")

    style_counts = {}
    for p in paragraphs:
        s = p["style"] or "(无样式)"
        style_counts[s] = style_counts.get(s, 0) + 1
    for s, c in sorted(style_counts.items(), key=lambda x: -x[1]):
        lines.append(f"  - {s}: {c}")

    lines.append("\n## 章节目录\n")
    for p in paragraphs:
        if p["level"] >= 1:
            indent = "  " * (p["level"] - 1)
            lines.append(f"{indent}- {p['text']}")

    return "\n".join(lines)


def cmd_extract(args):
    """extract 子命令入口"""
    paragraphs = extract_paragraphs(args.input)

    if args.info:
        print(document_info(paragraphs))
        return

    if args.json:
        output = json.dumps(paragraphs, ensure_ascii=False, indent=2)
        if args.output:
            with open(args.output, "w", encoding="utf-8") as f:
                f.write(output)
            print(f"已输出到 {args.output}")
        else:
            print(output)
        return

    if args.split_chapters:
        chapters = split_by_chapters(paragraphs)
        if args.output:
            os.makedirs(args.output, exist_ok=True)
            for i, ch in enumerate(chapters):
                safe_title = re.sub(r"[^\w\u4e00-\u9fff]+", "_", ch["title"])[:30]
                filename = f"{i:02d}_{safe_title}.md"
                filepath = os.path.join(args.output, filename)
                with open(filepath, "w", encoding="utf-8") as f:
                    f.write(ch["markdown"])
                print(f"  {filepath} ({len(ch['paragraphs'])} 段)")
        else:
            for ch in chapters:
                print(f"\n{'=' * 60}")
                print(f"## {ch['title']}")
                print(f"{'=' * 60}")
                print(ch["markdown"])
        return

    # 默认：全文 Markdown
    md = paragraphs_to_markdown(paragraphs)
    if args.output:
        with open(args.output, "w", encoding="utf-8") as f:
            f.write(md)
        print(f"已输出到 {args.output}（{len(paragraphs)} 段，{len(md)} 字符）")
    else:
        print(md)


def add_extract_parser(sub):
    """把 extract 子命令声明挂到组合入口的 subparsers（声明只写这一遍）。"""
    ep = sub.add_parser("extract", help="从 .docx 提取纯文本（Markdown 格式）")
    ep.add_argument("input", help="输入 .docx 文件")
    ep.add_argument("-o", "--output", help="输出文件/目录路径")
    ep.add_argument("--split-chapters", action="store_true", help="按一级标题拆分输出")
    ep.add_argument("--info", action="store_true", help="仅输出文档结构信息")
    ep.add_argument("--json", action="store_true", help="输出 JSON 格式（段落列表）")
    return ep


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(
        prog="docx-extract", description="docx extract — 文本提取（docx_tools 拆解实现模块）"
    )
    sub = parser.add_subparsers(dest="command", required=True)
    add_extract_parser(sub)
    args = parser.parse_args(argv)
    cmd_extract(args)
    return 0


if __name__ == "__main__":
    sys.exit(main())
