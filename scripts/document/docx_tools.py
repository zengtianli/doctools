#!/usr/bin/env python3
"""
Word 文档工具集 (docx_tools.py)

合并三个 docx 工具为统一入口：

子命令：
  extract        — 从 .docx 提取纯文本（Markdown 格式），支持分章节输出
  check          — 两层格式检查（ZIP 完整性 + 格式语义）
  track-changes  — 读取/写入修订标记和批注

用法：
  python3 docx_tools.py extract input.docx                     # 全文输出到 stdout
  python3 docx_tools.py extract input.docx -o output.md        # 全文输出到文件
  python3 docx_tools.py extract input.docx --split-chapters    # 按章节拆分输出
  python3 docx_tools.py extract input.docx --info              # 仅输出文档结构信息

  python3 docx_tools.py check snapshot input.docx              # 输出格式报告
  python3 docx_tools.py check snapshot input.docx -o snap.json # 存为 JSON 快照
  python3 docx_tools.py check compare  before.docx after.docx  # 对比两个文件

  python3 docx_tools.py track-changes read input.docx [--format md|json]
  python3 docx_tools.py track-changes review input.docx -o output.docx --rules rules.json
"""

import argparse
import copy
import hashlib
import json
import os
import re
import shutil
import sys
import tempfile
import zipfile
from collections import Counter
from datetime import UTC, datetime
from pathlib import Path

sys.path.insert(0, str(Path(__file__).parent.parent.parent / "lib"))

from lxml import etree

from docx_xml import R_NS, REL_COMMENTS, W, qn


# ══════════════════════════════════════════════════════════════════════
#  extract — 文本提取
# ══════════════════════════════════════════════════════════════════════


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


# ══════════════════════════════════════════════════════════════════════
#  check — 格式检查
# ══════════════════════════════════════════════════════════════════════

VML_NS = "urn:schemas-microsoft-com:vml"


# ── 半角 twips → cm 转换 ──
def twips_to_cm(val: str) -> float:
    return round(int(val) / 567, 2) if val else 0


def half_pt(val: str) -> float:
    """Word XML 字号是半磅，转为磅"""
    return int(val) / 2 if val else 0


# ── A 层：ZIP 条目级 hash ──


def zip_hashes(docx_path: str) -> dict[str, str]:
    """计算 docx 内每个文件的 MD5"""
    hashes = {}
    with zipfile.ZipFile(docx_path, "r") as zf:
        for info in zf.infolist():
            data = zf.read(info.filename)
            hashes[info.filename] = hashlib.md5(data).hexdigest()
    return hashes


# 审阅操作（添加 comments + 修改 document.xml）会连带改变以下文件
# 分为"必需"和"安全副作用"两类
EXPECTED_CHANGES = {"word/document.xml", "word/comments.xml"}
SAFE_SIDE_EFFECTS = {
    "[Content_Types].xml",  # 新增 comments 内容类型声明
    "word/_rels/document.xml.rels",  # 新增 comments 关系
    "docProps/core.xml",  # 修订号+1，修改时间更新
    "docProps/app.xml",  # 行数/段落数统计更新
    "word/settings.xml",  # rsid 修订标识更新
    "word/endnotes.xml",  # rsid 变化
    "word/footnotes.xml",  # rsid 变化
    "word/commentsExtended.xml",  # 批注扩展元数据（新增）
    "word/commentsIds.xml",  # 批注 ID 映射（新增）
}


def compare_zip_integrity(hashes_before: dict, hashes_after: dict) -> list[dict]:
    """对比两份 hash，返回差异列表"""
    diffs = []
    all_keys = set(hashes_before) | set(hashes_after)
    for key in sorted(all_keys):
        h1 = hashes_before.get(key)
        h2 = hashes_after.get(key)
        if h1 == h2:
            continue
        if key in EXPECTED_CHANGES:
            level = "expected"
        elif key in SAFE_SIDE_EFFECTS:
            level = "safe"
        else:
            level = "unexpected"

        if h1 is None:
            diffs.append({"file": key, "type": "added", "level": level})
        elif h2 is None:
            diffs.append({"file": key, "type": "removed", "level": level})
        else:
            diffs.append({"file": key, "type": "changed", "level": level})
    return diffs


# ── B 层：格式语义提取 ──


def _extract_hf_text(tree) -> str:
    """从页眉/页脚 XML 中提取文本，正确处理 PAGE 域和 Tab 分隔"""
    parts = []
    for para in tree.iter(f"{{{W}}}p"):
        run_texts = []
        in_field = False
        field_instr = ""
        for elem in para.iter():
            tag = elem.tag.split("}")[-1] if "}" in elem.tag else elem.tag
            if tag == "fldChar":
                ftype = elem.get(f"{{{W}}}fldCharType", "")
                if ftype == "begin":
                    in_field = True
                    field_instr = ""
                elif ftype == "end":
                    if "PAGE" in field_instr.upper():
                        run_texts.append("{页码}")
                    in_field = False
            elif tag == "instrText" and in_field:
                field_instr += elem.text or ""
            elif tag == "t" and not in_field:
                run_texts.append(elem.text or "")
            elif tag == "tab":
                run_texts.append(" | ")
        if run_texts:
            line = "".join(run_texts).strip()
            # 压缩连续空白为单个空格
            line = re.sub(r"\s{2,}", " ", line)
            if line:
                parts.append(line)
    return " / ".join(parts) if parts else ""


def extract_format_snapshot(docx_path: str) -> dict:
    """提取文档的完整格式快照"""
    snap = {
        "file": docx_path,
        "page_setup": [],
        "headers_footers": [],
        "watermark": None,
        "styles": [],
        "style_usage": {},
        "direct_overrides_count": 0,
        "images_count": 0,
        "zip_hashes": zip_hashes(docx_path),
    }

    REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships"

    with zipfile.ZipFile(docx_path, "r") as zf:
        names = zf.namelist()

        # ── 读取 relationship 映射：rId → 文件路径 ──
        rid_map = {}
        rels_path = "word/_rels/document.xml.rels"
        if rels_path in names:
            rels_xml = zf.read(rels_path)
            rels_tree = etree.fromstring(rels_xml)
            for rel in rels_tree.findall(f"{{{REL_NS}}}Relationship"):
                rid = rel.get("Id", "")
                target = rel.get("Target", "")
                # Target 是相对于 word/ 的路径
                if not target.startswith("/"):
                    target = "word/" + target
                rid_map[rid] = target

        # ── 预提取所有 header/footer 文件的文本 ──
        hf_text_cache = {}  # "word/header1.xml" → text
        hf_watermark = {}  # "word/header1.xml" → watermark text
        for name in names:
            if ("header" in name.lower() or "footer" in name.lower()) and name.endswith(".xml"):
                xml = zf.read(name)
                htree = etree.fromstring(xml)
                hf_text_cache[name] = _extract_hf_text(htree) or "(空)"
                # 检查水印
                content = xml.decode("utf-8", errors="ignore")
                if "v:shape" in content or "mso-position-horizontal:center" in content:
                    wm_matches = re.findall(r'string="([^"]+)"', content)
                    if wm_matches:
                        hf_watermark[name] = wm_matches[0]
                        snap["watermark"] = wm_matches[0]

        # ── 页面设置 + 页眉页脚（按分节符读取） ──
        doc_xml = zf.read("word/document.xml")
        tree = etree.fromstring(doc_xml)

        section_idx = 0
        for sectPr in tree.iter(f"{{{W}}}sectPr"):
            section_idx += 1  # noqa: SIM113
            sec = {"section": section_idx}

            pgSz = sectPr.find(f"{{{W}}}pgSz")
            if pgSz is not None:
                w = pgSz.get(f"{{{W}}}w", "")
                h = pgSz.get(f"{{{W}}}h", "")
                orient = pgSz.get(f"{{{W}}}orient", "portrait")
                sec["paper_w_cm"] = twips_to_cm(w)
                sec["paper_h_cm"] = twips_to_cm(h)
                sec["orientation"] = orient

            pgMar = sectPr.find(f"{{{W}}}pgMar")
            if pgMar is not None:
                sec["margin_top_cm"] = twips_to_cm(pgMar.get(f"{{{W}}}top", "0"))
                sec["margin_bottom_cm"] = twips_to_cm(pgMar.get(f"{{{W}}}bottom", "0"))
                sec["margin_left_cm"] = twips_to_cm(pgMar.get(f"{{{W}}}left", "0"))
                sec["margin_right_cm"] = twips_to_cm(pgMar.get(f"{{{W}}}right", "0"))
                sec["header_cm"] = twips_to_cm(pgMar.get(f"{{{W}}}header", "0"))
                sec["footer_cm"] = twips_to_cm(pgMar.get(f"{{{W}}}footer", "0"))

            # 从 sectPr 读取 headerReference / footerReference
            header_text = None
            footer_text = None
            for ref in sectPr.findall(f"{{{W}}}headerReference"):
                htype = ref.get(f"{{{W}}}type", "")
                if htype == "default":
                    rid = ref.get("{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id", "")
                    fname = rid_map.get(rid, "")
                    header_text = hf_text_cache.get(fname, "(空)")
            for ref in sectPr.findall(f"{{{W}}}footerReference"):
                ftype = ref.get(f"{{{W}}}type", "")
                if ftype == "default":
                    rid = ref.get("{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id", "")
                    fname = rid_map.get(rid, "")
                    footer_text = hf_text_cache.get(fname, "(空)")

            sec["header"] = header_text
            sec["footer"] = footer_text
            snap["page_setup"].append(sec)

        # 继承逻辑：没有显式引用的节继承前一节（Word "链接到上一节"）
        for i, sec in enumerate(snap["page_setup"]):
            if i > 0:
                prev = snap["page_setup"][i - 1]
                if sec["header"] is None:
                    sec["header"] = prev["header"]
                    sec["header_inherited"] = True
                if sec["footer"] is None:
                    sec["footer"] = prev["footer"]
                    sec["footer_inherited"] = True

        # 构建 headers_footers（兼容 compare 逻辑）
        for sec in snap["page_setup"]:
            hf_entry = {
                "section": sec["section"],
                "header": sec.get("header") or "(空)",
                "footer": sec.get("footer") or "(空)",
                "header_inherited": sec.get("header_inherited", False),
                "footer_inherited": sec.get("footer_inherited", False),
            }
            snap["headers_footers"].append(hf_entry)

        # ── 样式定义 ──
        if "word/styles.xml" in names:
            styles_xml = zf.read("word/styles.xml")
            stree = etree.fromstring(styles_xml)

            style_names_map = {}
            for s in stree.findall(f".//{{{W}}}style"):
                sid = s.get(f"{{{W}}}styleId", "")
                stype = s.get(f"{{{W}}}type", "")
                if stype != "paragraph":
                    continue

                name_elem = s.find(f"{{{W}}}name")
                name = name_elem.get(f"{{{W}}}val", sid) if name_elem is not None else sid
                style_names_map[sid] = name

                # 提取字体、字号、加粗
                info = {"id": sid, "name": name}

                rpr = s.find(f".//{{{W}}}rPr")
                if rpr is not None:
                    rf = rpr.find(f"{{{W}}}rFonts")
                    if rf is not None:
                        info["font_cn"] = rf.get(f"{{{W}}}eastAsia", "")
                        info["font_en"] = rf.get(f"{{{W}}}ascii", "")
                    sz = rpr.find(f"{{{W}}}sz")
                    if sz is not None:
                        info["size_pt"] = half_pt(sz.get(f"{{{W}}}val", ""))
                    if rpr.find(f"{{{W}}}b") is not None:
                        info["bold"] = True

                ppr = s.find(f".//{{{W}}}pPr")
                if ppr is not None:
                    jc = ppr.find(f"{{{W}}}jc")
                    if jc is not None:
                        info["align"] = jc.get(f"{{{W}}}val", "")
                    sp = ppr.find(f"{{{W}}}spacing")
                    if sp is not None:
                        line = sp.get(f"{{{W}}}line", "")
                        if line:
                            info["line_spacing"] = int(line)
                    ind = ppr.find(f"{{{W}}}ind")
                    if ind is not None:
                        fc = ind.get(f"{{{W}}}firstLineChars", "")
                        if fc:
                            info["indent_first_chars"] = int(fc)

                snap["styles"].append(info)

            # ── 样式使用统计 ──
            usage = Counter()
            direct_count = 0
            for para in tree.iter(f"{{{W}}}p"):
                ppr = para.find(f"{{{W}}}pPr")
                sid = ""
                if ppr is not None:
                    ps = ppr.find(f"{{{W}}}pStyle")
                    if ps is not None:
                        sid = ps.get(f"{{{W}}}val", "")
                sname = style_names_map.get(sid, sid or "Normal")
                usage[sname] += 1

                for run in para.iter(f"{{{W}}}r"):
                    parent = run.getparent()
                    if parent is not None and parent.tag in (f"{{{W}}}del", f"{{{W}}}ins"):
                        continue
                    rpr = run.find(f"{{{W}}}rPr")
                    if rpr is not None:
                        has_font = rpr.find(f"{{{W}}}rFonts") is not None
                        has_size = rpr.find(f"{{{W}}}sz") is not None
                        if has_font or has_size:
                            direct_count += 1

            snap["style_usage"] = dict(usage.most_common())
            snap["direct_overrides_count"] = direct_count

        # ── 图片数量 ──
        snap["images_count"] = len([n for n in names if n.startswith("word/media/")])

    return snap


# ── 格式报告生成 ──

ALIGN_MAP = {"both": "两端对齐", "center": "居中", "left": "左对齐", "right": "右对齐"}


def _format_style_row(st: dict) -> str:
    """格式化一行样式定义"""
    name = st["name"]
    font_cn = st.get("font_cn", "")
    font_en = st.get("font_en", "")
    size = f"{st['size_pt']}pt" if st.get("size_pt") else "-"
    bold = "**是**" if st.get("bold") else "-"
    align = ALIGN_MAP.get(st.get("align", ""), st.get("align", "-"))
    if st.get("line_spacing"):
        val = st["line_spacing"]
        line_sp = {240: "单倍", 300: "1.25倍", 360: "1.5倍", 480: "2倍"}.get(val, f"{val}twips")
    else:
        line_sp = "-"
    indent = f"{st['indent_first_chars']}字符" if st.get("indent_first_chars") else "-"
    return f"| {name} | {font_cn or '-'} | {font_en or '-'} | {size} | {bold} | {align} | {line_sp} | {indent} |"


def format_report(snap: dict) -> str:
    """生成人可读的格式报告"""
    lines = []
    lines.append("# 文档格式报告\n")
    lines.append(f"文件：`{snap['file']}`\n")

    # ── 1. 页面设置 ──
    lines.append("## 1. 页面设置\n")
    orientations = Counter(s.get("orientation", "portrait") for s in snap["page_setup"])
    lines.append(
        f"共 {len(snap['page_setup'])} 个分节符"
        f"（纵向 {orientations.get('portrait', 0)} 个"
        f"，横向 {orientations.get('landscape', 0)} 个）\n"
    )
    if snap["page_setup"]:
        s0 = snap["page_setup"][0]
        lines.append("| 项目 | 值 |")
        lines.append("|------|------|")
        lines.append(f"| 纸张 | {s0.get('paper_w_cm', 0)}cm × {s0.get('paper_h_cm', 0)}cm (A4) |")
        lines.append(f"| 上边距 | {s0.get('margin_top_cm', 0)}cm |")
        lines.append(f"| 下边距 | {s0.get('margin_bottom_cm', 0)}cm |")
        lines.append(f"| 左边距 | {s0.get('margin_left_cm', 0)}cm |")
        lines.append(f"| 右边距 | {s0.get('margin_right_cm', 0)}cm |")
        lines.append(f"| 页眉距 | {s0.get('header_cm', 0)}cm |")
        lines.append(f"| 页脚距 | {s0.get('footer_cm', 0)}cm |")
    lines.append("")

    # ── 2. 页眉页脚 & 水印 ──
    lines.append("## 2. 页眉页脚\n")
    if snap["watermark"]:
        lines.append(f"水印：**{snap['watermark']}**\n")
    else:
        lines.append("水印：无\n")

    if snap["headers_footers"]:
        lines.append("| 节 | 页眉 | 页脚 |")
        lines.append("|------|----------|----------|")
        for hf in snap["headers_footers"]:
            sec_num = hf.get("section", "?")
            h_text = hf.get("header") or "(无)"
            f_text = hf.get("footer") or "(无)"
            if hf.get("header_inherited") and h_text != "(空)":
                h_text = f"↑ {h_text}"
            if hf.get("footer_inherited") and f_text != "(空)":
                f_text = f"↑ {f_text}"
            lines.append(f"| {sec_num} | {h_text} | {f_text} |")
    else:
        lines.append("无页眉页脚。")
    lines.append("")

    # ── 3. 样式定义（只显示使用中的） ──
    _ = set()
    style_by_id = {}
    for st in snap["styles"]:
        style_by_id[st["id"]] = st
        # 也用 name 做匹配（style_usage 存的是 name）
        style_by_id[st["name"]] = st

    used_names = set(snap.get("style_usage", {}).keys())

    # 分为：使用中 vs 未使用
    used_styles = []
    unused_styles = []
    for st in snap["styles"]:
        if st["name"] in used_names or st["id"] in used_names:
            used_styles.append(st)
        else:
            unused_styles.append(st)

    # 按使用量排序
    usage = snap.get("style_usage", {})
    used_styles.sort(key=lambda s: -usage.get(s["name"], usage.get(s["id"], 0)))

    lines.append("## 3. 使用中的样式\n")
    lines.append(f"共 {len(used_styles)} 个样式正在使用，{len(unused_styles)} 个已定义但未使用。\n")
    lines.append("| 样式名 | 段落数 | 中文字体 | 西文字体 | 字号 | 加粗 | 对齐 | 行距 | 首行缩进 |")
    lines.append("|--------|--------|----------|----------|------|------|------|------|----------|")
    for st in used_styles:
        count = usage.get(st["name"], usage.get(st["id"], 0))
        name = st["name"]
        font_cn = st.get("font_cn", "")
        font_en = st.get("font_en", "")
        size = f"{st['size_pt']}pt" if st.get("size_pt") else "-"
        bold = "**是**" if st.get("bold") else "-"
        align = ALIGN_MAP.get(st.get("align", ""), st.get("align", "-"))
        if st.get("line_spacing"):
            val = st["line_spacing"]
            line_sp = {240: "单倍", 300: "1.25倍", 360: "1.5倍", 480: "2倍"}.get(val, f"{val}twips")
        else:
            line_sp = "-"
        indent = f"{st['indent_first_chars']}字符" if st.get("indent_first_chars") else "-"
        lines.append(
            f"| {name} | {count} | {font_cn or '-'} | {font_en or '-'} | {size} | {bold} | {align} | {line_sp} | {indent} |"
        )
    lines.append("")

    # ── 4. 未使用的样式（仅列名称） ──
    if unused_styles:
        lines.append("## 4. 未使用的样式\n")
        lines.append("<details>\n<summary>展开查看 " + str(len(unused_styles)) + " 个未使用样式</summary>\n")
        lines.append("| 样式名 | 中文字体 | 西文字体 | 字号 | 加粗 | 对齐 | 行距 | 首行缩进 |")
        lines.append("|--------|----------|----------|------|------|------|------|----------|")
        for st in sorted(unused_styles, key=lambda x: x["name"]):
            lines.append(_format_style_row(st))
        lines.append("\n</details>")
    lines.append("")

    # ── 5. 格式指纹 ──
    lines.append("## 5. 格式指纹\n")
    lines.append("| 项目 | 值 |")
    lines.append("|------|------|")
    lines.append(f"| 直接格式覆盖（run 级） | {snap['direct_overrides_count']} 个 |")
    lines.append(f"| 嵌入图片 | {snap['images_count']} 张 |")
    lines.append(f"| 段落样式定义总数 | {len(snap['styles'])} 个 |")
    total_paras = sum(snap.get("style_usage", {}).values())
    lines.append(f"| 段落总数 | {total_paras} 个 |")
    lines.append("")

    return "\n".join(lines)


def compare_report(snap1: dict, snap2: dict) -> str:
    """对比两份快照，生成差异报告"""
    lines = []
    lines.append("# 格式对比报告\n")
    lines.append(f"- 原始：`{snap1['file']}`")
    lines.append(f"- 修改后：`{snap2['file']}`\n")

    all_ok = True

    # ── A 层：ZIP 完整性 ──
    lines.append("## A 层：ZIP 文件完整性\n")
    diffs = compare_zip_integrity(snap1["zip_hashes"], snap2["zip_hashes"])
    if not diffs:
        lines.append("✅ 所有文件完全一致（无任何改动）\n")
    else:
        expected = [d for d in diffs if d["level"] == "expected"]
        safe = [d for d in diffs if d["level"] == "safe"]
        unexpected = [d for d in diffs if d["level"] == "unexpected"]

        if expected:
            lines.append("✅ 预期变化（审阅修订核心文件）：")
            for d in expected:
                lines.append(f"  - `{d['file']}` ({d['type']})")

        if safe:
            lines.append("\nℹ️  安全副作用（批注/修订的基础设施文件，不影响格式）：")
            for d in safe:
                lines.append(f"  - `{d['file']}` ({d['type']})")

        if unexpected:
            all_ok = False
            lines.append("\n❌ **非预期变化**（可能影响格式/页面/图片）：")
            for d in unexpected:
                lines.append(f"  - `{d['file']}` ({d['type']})")

        if not unexpected:
            lines.append(f"\n  共 {len(expected) + len(safe)} 个文件变化，全部在预期范围内。")
    lines.append("")

    # ── B 层：语义对比 ──
    lines.append("## B 层：格式语义对比\n")

    # 页面设置
    ps1 = snap1["page_setup"]
    ps2 = snap2["page_setup"]
    if json.dumps(ps1) == json.dumps(ps2):
        lines.append(f"✅ 页面设置：未变化（{len(ps1)} 个分节符）")
    else:
        all_ok = False
        lines.append("❌ 页面设置：有变化！")
        lines.append(f"  原始 {len(ps1)} 个分节符 → 修改后 {len(ps2)} 个")

    # 页眉页脚
    hf1 = [(h.get("section"), h.get("header"), h.get("footer")) for h in snap1["headers_footers"]]
    hf2 = [(h.get("section"), h.get("header"), h.get("footer")) for h in snap2["headers_footers"]]
    if hf1 == hf2:
        lines.append("✅ 页眉页脚：未变化" + (f"（{len(hf1)} 节）" if hf1 else "（无）"))
    else:
        all_ok = False
        lines.append("❌ 页眉页脚：有变化！")

    # 水印
    wm1 = snap1["watermark"]
    wm2 = snap2["watermark"]
    if wm1 == wm2:
        lines.append("✅ 水印：未变化" + (f'（"{wm1}"）' if wm1 else "（无）"))
    else:
        all_ok = False
        lines.append(f'❌ 水印："{wm1}" → "{wm2}"')

    # 样式数量
    s1 = len(snap1["styles"])
    s2 = len(snap2["styles"])
    if s1 == s2:
        lines.append(f"✅ 样式定义：未变化（{s1} 个段落样式）")
    else:
        all_ok = False
        lines.append(f"❌ 样式定义：{s1} → {s2} 个")

    # 样式使用
    u1 = snap1["style_usage"]
    u2 = snap2["style_usage"]
    usage_diff = {}
    for k in set(u1) | set(u2):
        v1 = u1.get(k, 0)
        v2 = u2.get(k, 0)
        if v1 != v2:
            usage_diff[k] = (v1, v2)
    if not usage_diff:
        lines.append("✅ 样式使用：段落分布完全一致")
    else:
        # 样式使用变化可能是正常的（比如删除了一段），标为 info
        lines.append(f"ℹ️  样式使用：{len(usage_diff)} 个样式的段落数有变化")
        for k, (v1, v2) in sorted(usage_diff.items()):
            lines.append(f"  - {k}: {v1} → {v2}")

    # 直接覆盖
    do1 = snap1["direct_overrides_count"]
    do2 = snap2["direct_overrides_count"]
    if do1 == do2:
        lines.append(f"✅ 直接格式覆盖：{do1} 个 run（未变化）")
    else:
        diff = do2 - do1
        sign = "+" if diff > 0 else ""
        lines.append(f"ℹ️  直接格式覆盖：{do1} → {do2}（{sign}{diff}），可能因修订标记导致")

    # 图片
    i1 = snap1["images_count"]
    i2 = snap2["images_count"]
    if i1 == i2:
        lines.append(f"✅ 嵌入图片：{i1} 张（未变化）")
    else:
        all_ok = False
        lines.append(f"❌ 嵌入图片：{i1} → {i2} 张")

    lines.append("")
    if all_ok:
        lines.append("## 结论：✅ 格式完整性通过\n")
        lines.append("只有文本内容被修改（修订标记），文档格式、页面设置、图片等均未变化。")
    else:
        lines.append("## 结论：❌ 发现非预期变化\n")
        lines.append("请检查上述标记为 ❌ 的项目。")

    return "\n".join(lines)


def cmd_check(args):
    """check 子命令入口"""
    if args.check_command == "snapshot":
        snap = extract_format_snapshot(args.input)
        report = format_report(snap)
        # 终端始终输出
        print(report)
        # 同时写文件
        if args.output:
            output = args.output
        elif args.md:
            output = os.path.splitext(args.input)[0] + "_格式报告.md"
        else:
            output = None
        if output:
            if output.endswith(".json"):
                snap_out = {k: v for k, v in snap.items() if k != "zip_hashes"}
                with open(output, "w", encoding="utf-8") as f:
                    json.dump(snap_out, f, ensure_ascii=False, indent=2)
            else:
                with open(output, "w", encoding="utf-8") as f:
                    f.write(report)
            print(f"\n已同时输出到 {output}")

    elif args.check_command == "compare":
        snap1 = extract_format_snapshot(args.before)
        snap2 = extract_format_snapshot(args.after)
        report = compare_report(snap1, snap2)
        # 终端始终输出
        print(report)
        # 同时写文件
        if args.output:
            output = args.output
        elif args.md:
            output = os.path.splitext(args.before)[0] + "_格式对比.md"
        else:
            output = None
        if output:
            with open(output, "w", encoding="utf-8") as f:
                f.write(report)
            print(f"\n已同时输出到 {output}")

    else:
        print("请指定子命令：snapshot 或 compare")
        sys.exit(1)


# ══════════════════════════════════════════════════════════════════════
#  track-changes — 修订标记
# ══════════════════════════════════════════════════════════════════════


def read_track_changes(docx_path: str, fmt: str = "md") -> str:
    """读取 .docx 中的修订和批注，返回 Markdown 或 JSON"""
    changes = []
    comments = []

    with zipfile.ZipFile(docx_path, "r") as zf:
        doc_xml = zf.read("word/document.xml")
        tree = etree.fromstring(doc_xml)

        for ins in tree.iter(qn("w:ins")):
            author = ins.get(qn("w:author"), "未知")
            date = ins.get(qn("w:date"), "")
            text = _tc_extract_text(ins)
            if text.strip():
                changes.append({"type": "insert", "author": author, "date": date, "text": text})

        for dl in tree.iter(qn("w:del")):
            author = dl.get(qn("w:author"), "未知")
            date = dl.get(qn("w:date"), "")
            text = _tc_extract_del_text(dl)
            if text.strip():
                changes.append({"type": "delete", "author": author, "date": date, "text": text})

        if "word/comments.xml" in zf.namelist():
            comments_xml = zf.read("word/comments.xml")
            ctree = etree.fromstring(comments_xml)
            for comment in ctree.iter(qn("w:comment")):
                comments.append(
                    {
                        "id": comment.get(qn("w:id"), ""),
                        "author": comment.get(qn("w:author"), "未知"),
                        "date": comment.get(qn("w:date"), ""),
                        "text": _tc_extract_text(comment),
                    }
                )

    if fmt == "json":
        return json.dumps({"changes": changes, "comments": comments}, ensure_ascii=False, indent=2)

    # Markdown 格式
    lines = []
    if changes:
        lines.append("## 修订记录\n")
        for i, c in enumerate(changes, 1):
            icon = "插入" if c["type"] == "insert" else "删除"
            date_str = c["date"][:10] if c["date"] else ""
            lines.append(f"{i}. **{icon}** | {c['author']} | {date_str}")
            lines.append(f"   > {c['text']}\n")
    else:
        lines.append("## 修订记录\n\n无修订。\n")

    if comments:
        lines.append("## 批注\n")
        for c in comments:
            date_str = c["date"][:10] if c["date"] else ""
            lines.append(f"- **[{c['id']}]** {c['author']} ({date_str}):")
            lines.append(f"  > {c['text']}\n")

    return "\n".join(lines)


def _tc_extract_text(node) -> str:
    """从 XML 节点中提取所有 <w:t> 文本"""
    return "".join(t.text for t in node.iter(qn("w:t")) if t.text)


def _tc_extract_del_text(node) -> str:
    """从删除标记中提取 <w:delText> 文本"""
    parts = [t.text for t in node.iter(qn("w:delText")) if t.text]
    if not parts:
        parts = [t.text for t in node.iter(qn("w:t")) if t.text]
    return "".join(parts)


class DocxReviewer:
    """对 .docx 文件应用替换规则，生成带修订标记的新文件。保持原文格式不变。"""

    def __init__(self, docx_path: str, author: str = "CC审阅"):
        self.docx_path = docx_path
        self.author = author
        self.date = datetime.now(UTC).strftime("%Y-%m-%dT%H:%M:%SZ")
        self.comment_id_counter = 0
        self.comments = []
        self.revision_id_counter = 100

        self.tmpdir = tempfile.mkdtemp(prefix="docx_review_")
        with zipfile.ZipFile(docx_path, "r") as zf:
            zf.extractall(self.tmpdir)

        doc_path = os.path.join(self.tmpdir, "word", "document.xml")
        self.doc_tree = etree.parse(doc_path)
        self.doc_root = self.doc_tree.getroot()
        self._init_comment_ids()

    def _init_comment_ids(self):
        comments_path = os.path.join(self.tmpdir, "word", "comments.xml")
        if os.path.exists(comments_path):
            ctree = etree.parse(comments_path)
            for c in ctree.getroot().iter(qn("w:comment")):
                cid = int(c.get(qn("w:id"), "0"))
                if cid >= self.comment_id_counter:
                    self.comment_id_counter = cid + 1

    def _next_rid(self) -> str:
        self.revision_id_counter += 1
        return str(self.revision_id_counter)

    def _next_comment_id(self) -> int:
        cid = self.comment_id_counter
        self.comment_id_counter += 1
        return cid

    def apply_rules(self, rules: list[dict]) -> int:
        """应用替换规则列表。返回成功替换的数量。"""
        count = 0
        for rule in rules:
            n = self._apply_one_rule(rule["find"], rule["replace"], rule.get("comment"))
            count += n
        return count

    def _apply_one_rule(self, find: str, replace: str, comment: str | None) -> int:
        body = self.doc_root.find(qn("w:body"))
        if body is None:
            return 0
        count = 0
        for para in body.iter(qn("w:p")):
            while True:
                result = self._find_in_paragraph(para, find)
                if result is None:
                    break
                self._replace_in_paragraph(para, result, find, replace, comment)
                count += 1
        return count

    def _find_in_paragraph(self, para, find_text: str):
        """在段落中跨 run 搜索文本（跳过已有修订标记内的 run）"""
        runs = list(para.iter(qn("w:r")))
        if not runs:
            return None

        active_runs = []
        for r in runs:
            parent = r.getparent()
            if parent is not None and parent.tag in (qn("w:del"), qn("w:ins")):
                continue
            active_runs.append(r)
        if not active_runs:
            return None

        run_texts = []
        for r in active_runs:
            t_elem = r.find(qn("w:t"))
            run_texts.append(t_elem.text if t_elem is not None and t_elem.text else "")

        full_text = "".join(run_texts)
        idx = full_text.find(find_text)
        if idx == -1:
            return None

        start_pos, end_pos = idx, idx + len(find_text)
        cumulative = 0
        start_run_idx = end_run_idx = None
        start_offset = end_offset = 0

        for i, text in enumerate(run_texts):
            run_start = cumulative
            run_end = cumulative + len(text)
            if start_run_idx is None and run_end > start_pos:
                start_run_idx = i
                start_offset = start_pos - run_start
            if run_end >= end_pos:
                end_run_idx = i
                end_offset = end_pos - run_start
                break
            cumulative = run_end

        if start_run_idx is None or end_run_idx is None:
            return None

        return {
            "runs": active_runs[start_run_idx : end_run_idx + 1],
            "start_offset": start_offset,
            "end_offset": end_offset,
        }

    def _replace_in_paragraph(self, para, match, find_text, replace_text, comment_text):
        """拆分 run，插入 del/ins 标记，保持原有格式"""
        runs = match["runs"]
        start_offset = match["start_offset"]
        end_offset = match["end_offset"]

        # 继承第一个 run 的格式
        rpr_template = runs[0].find(qn("w:rPr"))
        if rpr_template is not None:
            rpr_template = copy.deepcopy(rpr_template)

        # 前缀文本（第一个 run 中匹配之前的部分）
        first_t = runs[0].find(qn("w:t"))
        first_text = first_t.text if first_t is not None and first_t.text else ""
        prefix_text = first_text[:start_offset]

        # 后缀文本（最后一个 run 中匹配之后的部分）
        last_t = runs[-1].find(qn("w:t"))
        last_text = last_t.text if last_t is not None and last_t.text else ""
        suffix_text = last_text[end_offset:]

        # 记录插入位置，删除原 run
        parent = runs[0].getparent()
        insert_pos = list(parent).index(runs[0])
        for r in runs:
            r.getparent().remove(r)

        # 构建替换节点
        nodes = []

        if prefix_text:
            nodes.append(self._make_run(prefix_text, rpr_template))

        # 批注起始
        comment_id = None
        if comment_text:
            comment_id = self._next_comment_id()
            cs = etree.Element(qn("w:commentRangeStart"))
            cs.set(qn("w:id"), str(comment_id))
            nodes.append(cs)

        # <w:del>
        rid = self._next_rid()
        del_node = etree.Element(qn("w:del"))
        del_node.set(qn("w:id"), rid)
        del_node.set(qn("w:author"), self.author)
        del_node.set(qn("w:date"), self.date)
        del_node.append(self._make_del_run(find_text, rpr_template))
        nodes.append(del_node)

        # <w:ins>
        ins_node = etree.Element(qn("w:ins"))
        ins_node.set(qn("w:id"), self._next_rid())
        ins_node.set(qn("w:author"), self.author)
        ins_node.set(qn("w:date"), self.date)
        ins_node.append(self._make_run(replace_text, rpr_template))
        nodes.append(ins_node)

        # 批注结束 + 引用
        if comment_text and comment_id is not None:
            ce = etree.Element(qn("w:commentRangeEnd"))
            ce.set(qn("w:id"), str(comment_id))
            nodes.append(ce)

            ref_run = etree.Element(qn("w:r"))
            ref_rpr = etree.SubElement(ref_run, qn("w:rPr"))
            ref_style = etree.SubElement(ref_rpr, qn("w:rStyle"))
            ref_style.set(qn("w:val"), "CommentReference")
            ref_elem = etree.SubElement(ref_run, qn("w:commentReference"))
            ref_elem.set(qn("w:id"), str(comment_id))
            nodes.append(ref_run)

            self.comments.append(
                {
                    "id": comment_id,
                    "author": self.author,
                    "date": self.date,
                    "text": comment_text,
                }
            )

        if suffix_text:
            nodes.append(self._make_run(suffix_text, rpr_template))

        for i, node in enumerate(nodes):
            parent.insert(insert_pos + i, node)

    def _make_run(self, text: str, rpr=None) -> etree._Element:
        """创建 <w:r>，继承原有 rPr 格式"""
        run = etree.Element(qn("w:r"))
        if rpr is not None:
            run.append(copy.deepcopy(rpr))
        t = etree.SubElement(run, qn("w:t"))
        t.text = text
        t.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
        return run

    def _make_del_run(self, text: str, rpr=None) -> etree._Element:
        """创建 <w:del> 内部的 run（用 <w:delText>）"""
        run = etree.Element(qn("w:r"))
        if rpr is not None:
            run.append(copy.deepcopy(rpr))
        dt = etree.SubElement(run, qn("w:delText"))
        dt.text = text
        dt.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
        return run

    def _write_comments_xml(self):
        if not self.comments:
            return

        comments_path = os.path.join(self.tmpdir, "word", "comments.xml")
        if os.path.exists(comments_path):
            ctree = etree.parse(comments_path)
            croot = ctree.getroot()
        else:
            croot = etree.Element(qn("w:comments"), nsmap={"w": W, "r": R_NS})

        for c in self.comments:
            ce = etree.SubElement(croot, qn("w:comment"))
            ce.set(qn("w:id"), str(c["id"]))
            ce.set(qn("w:author"), c["author"])
            ce.set(qn("w:date"), c["date"])
            ce.set(qn("w:initials"), c["author"][:2])

            p = etree.SubElement(ce, qn("w:p"))
            etree.SubElement(p, qn("w:pPr"))
            r = etree.SubElement(p, qn("w:r"))
            rpr = etree.SubElement(r, qn("w:rPr"))
            rs = etree.SubElement(rpr, qn("w:rStyle"))
            rs.set(qn("w:val"), "CommentReference")
            etree.SubElement(r, qn("w:annotationRef"))

            r2 = etree.SubElement(p, qn("w:r"))
            t = etree.SubElement(r2, qn("w:t"))
            t.text = c["text"]
            t.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")

        with open(comments_path, "wb") as f:
            f.write(etree.tostring(croot, xml_declaration=True, encoding="UTF-8", standalone=True))
        self._ensure_content_type("comments")
        self._ensure_rels("comments")

    def _ensure_content_type(self, part: str):
        ct_path = os.path.join(self.tmpdir, "[Content_Types].xml")
        ct_tree = etree.parse(ct_path)
        ct_root = ct_tree.getroot()
        ct_ns = "http://schemas.openxmlformats.org/package/2006/content-types"

        if part == "comments":
            part_name = "/word/comments.xml"
            ct = "application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml"
            for o in ct_root.iter(f"{{{ct_ns}}}Override"):
                if o.get("PartName") == part_name:
                    return
            o = etree.SubElement(ct_root, f"{{{ct_ns}}}Override")
            o.set("PartName", part_name)
            o.set("ContentType", ct)
            with open(ct_path, "wb") as f:
                f.write(etree.tostring(ct_tree, xml_declaration=True, encoding="UTF-8", standalone=True))

    def _ensure_rels(self, part: str):
        rels_path = os.path.join(self.tmpdir, "word", "_rels", "document.xml.rels")
        rels_tree = etree.parse(rels_path)
        rels_root = rels_tree.getroot()
        rels_ns = "http://schemas.openxmlformats.org/package/2006/relationships"

        if part == "comments":
            for rel in rels_root.iter(f"{{{rels_ns}}}Relationship"):
                if rel.get("Type") == REL_COMMENTS:
                    return
            max_id = 0
            for rel in rels_root.iter(f"{{{rels_ns}}}Relationship"):
                m = re.search(r"(\d+)", rel.get("Id", "rId0"))
                if m:
                    max_id = max(max_id, int(m.group(1)))
            new_rel = etree.SubElement(rels_root, f"{{{rels_ns}}}Relationship")
            new_rel.set("Id", f"rId{max_id + 1}")
            new_rel.set("Type", REL_COMMENTS)
            new_rel.set("Target", "comments.xml")
            with open(rels_path, "wb") as f:
                f.write(etree.tostring(rels_tree, xml_declaration=True, encoding="UTF-8", standalone=True))

    def save(self, output_path: str):
        """保存修改后的 .docx"""
        doc_path = os.path.join(self.tmpdir, "word", "document.xml")
        with open(doc_path, "wb") as f:
            f.write(etree.tostring(self.doc_tree, xml_declaration=True, encoding="UTF-8", standalone=True))
        self._write_comments_xml()

        output_path = os.path.abspath(output_path)
        with zipfile.ZipFile(output_path, "w", zipfile.ZIP_DEFLATED) as zf:
            for root, _dirs, files in os.walk(self.tmpdir):
                for fn in files:
                    abs_path = os.path.join(root, fn)
                    zf.write(abs_path, os.path.relpath(abs_path, self.tmpdir))
        shutil.rmtree(self.tmpdir, ignore_errors=True)

    def cleanup(self):
        if os.path.exists(self.tmpdir):
            shutil.rmtree(self.tmpdir, ignore_errors=True)


def review_docx(input_path: str, output_path: str, rules: list[dict], author: str = "CC审阅") -> int:
    """便捷函数：对 .docx 应用替换规则，输出带修订标记的新文件。"""
    reviewer = DocxReviewer(input_path, author=author)
    try:
        count = reviewer.apply_rules(rules)
        reviewer.save(output_path)
        return count
    except Exception:
        reviewer.cleanup()
        raise


def cmd_track_changes(args):
    """track-changes 子命令入口"""
    if args.tc_command == "read":
        print(read_track_changes(args.input, args.format))
    elif args.tc_command == "review":
        with open(args.rules, encoding="utf-8") as f:
            rules = json.load(f)
        count = review_docx(args.input, args.output, rules, author=args.author)
        print(f"完成：{count} 处替换已写入 {args.output}")
    elif args.tc_command == "compare":
        print("compare 功能将在 v2 实现。")
        sys.exit(1)
    else:
        print("请指定子命令：read、review 或 compare")
        sys.exit(1)


# ══════════════════════════════════════════════════════════════════════
#  CLI 入口
# ══════════════════════════════════════════════════════════════════════


def main():
    parser = argparse.ArgumentParser(
        description="Word 文档工具集",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog=__doc__,
    )
    sub = parser.add_subparsers(dest="command", help="子命令")

    # ── extract ──
    ep = sub.add_parser("extract", help="从 .docx 提取纯文本（Markdown 格式）")
    ep.add_argument("input", help="输入 .docx 文件")
    ep.add_argument("-o", "--output", help="输出文件/目录路径")
    ep.add_argument("--split-chapters", action="store_true", help="按一级标题拆分输出")
    ep.add_argument("--info", action="store_true", help="仅输出文档结构信息")
    ep.add_argument("--json", action="store_true", help="输出 JSON 格式（段落列表）")

    # ── check ──
    ck = sub.add_parser("check", help="两层格式检查（ZIP 完整性 + 格式语义）")
    ck_sub = ck.add_subparsers(dest="check_command", help="check 子命令")

    sp = ck_sub.add_parser("snapshot", help="提取格式快照")
    sp.add_argument("input", help="输入 .docx 文件")
    sp.add_argument("-o", "--output", help="输出文件（.md 或 .json）")
    sp.add_argument("--md", action="store_true", help="自动输出为同名 .md 文件")

    cp = ck_sub.add_parser("compare", help="对比两个 .docx 文件的格式")
    cp.add_argument("before", help="原始 .docx 文件")
    cp.add_argument("after", help="修改后 .docx 文件")
    cp.add_argument("-o", "--output", help="输出报告文件")
    cp.add_argument("--md", action="store_true", help="自动输出为 .md 文件")

    # ── track-changes ──
    tc = sub.add_parser("track-changes", help="读取/写入修订标记和批注")
    tc_sub = tc.add_subparsers(dest="tc_command", help="track-changes 子命令")

    rp = tc_sub.add_parser("read", help="读取修订和批注")
    rp.add_argument("input", help="输入 .docx 文件")
    rp.add_argument("--format", "-f", choices=["md", "json"], default="md", help="输出格式 (默认: md)")

    wp = tc_sub.add_parser("review", help="写入修订标记")
    wp.add_argument("input", help="输入 .docx 文件")
    wp.add_argument("--output", "-o", required=True, help="输出 .docx 文件")
    wp.add_argument("--rules", "-r", required=True, help="替换规则 JSON 文件")
    wp.add_argument("--author", "-a", default="CC审阅", help="作者名")

    tcp = tc_sub.add_parser("compare", help="对比生成修订 (v2)")
    tcp.add_argument("original", help="原始 .docx")
    tcp.add_argument("modified", help="修改后 .docx")
    tcp.add_argument("--output", "-o", required=True, help="输出 .docx")

    args = parser.parse_args()

    if args.command == "extract":
        cmd_extract(args)
    elif args.command == "check":
        cmd_check(args)
    elif args.command == "track-changes":
        cmd_track_changes(args)
    else:
        parser.print_help()


if __name__ == "__main__":
    main()
