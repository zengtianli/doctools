#!/usr/bin/env python3
"""修复 DOCX 文档中的引用角标：将 [数字] 或 [数字-数字] 格式的引用改为上标。

用法:
    python3 fix_superscript_refs.py input.docx [-o output.docx] [--dry-run]

说明:
    扫描文档所有段落（含修订标记中的文本），将 [1] [3-4] [18-19] 等
    文献引用标记拆分为独立 run 并设为上标格式。

    如果引用已经是上标，则跳过。不处理参考文献列表中的编号（行首 [数字]）。
"""

import argparse
import copy
import re
import sys
from pathlib import Path

from docx import Document
from docx.oxml.ns import qn
from docx.oxml import OxmlElement


REF_PATTERN = re.compile(r"(\[[1-9][0-9]*(?:[,-][1-9][0-9]*)*\])")


def is_superscript(r_elem):
    """检查 run 是否已经是上标。"""
    rPr = r_elem.find(qn("w:rPr"))
    if rPr is None:
        return False
    vert = rPr.find(qn("w:vertAlign"))
    return vert is not None and vert.get(qn("w:val")) == "superscript"


def is_ref_list_line(text):
    """判断是否为参考文献列表行（以 [数字] 开头）。"""
    return bool(re.match(r"^\s*\[\d+\]", text.strip()))


def clone_rPr(source_rPr):
    if source_rPr is not None:
        return copy.deepcopy(source_rPr)
    return OxmlElement("w:rPr")


def split_run(r_elem):
    """将 run 中的 [数字] 拆成普通文本 + 上标引用。返回修改数量。"""
    if is_superscript(r_elem):
        return 0

    t_elem = r_elem.find(qn("w:t"))
    if t_elem is None or not t_elem.text:
        return 0

    text = t_elem.text
    if not REF_PATTERN.search(text):
        return 0

    rPr = r_elem.find(qn("w:rPr"))
    parts = REF_PATTERN.split(text)
    parts = [p for p in parts if p]  # 去空

    if len(parts) <= 1 and not REF_PATTERN.match(text):
        return 0

    parent = r_elem.getparent()
    insert_point = r_elem
    count = 0

    for part in parts:
        new_r = OxmlElement("w:r")
        new_rPr = clone_rPr(rPr)

        if REF_PATTERN.match(part):
            # 上标
            vert = new_rPr.find(qn("w:vertAlign"))
            if vert is not None:
                new_rPr.remove(vert)
            vert = OxmlElement("w:vertAlign")
            vert.set(qn("w:val"), "superscript")
            new_rPr.append(vert)
            count += 1

        new_r.append(new_rPr)
        new_t = OxmlElement("w:t")
        new_t.set(qn("xml:space"), "preserve")
        new_t.text = part
        new_r.append(new_t)

        insert_point.addnext(new_r)
        insert_point = new_r

    parent.remove(r_elem)
    return count


def full_text(para):
    """段落完整文本（含修订标记）。"""
    return "".join(
        n.text for n in para._element.iter()
        if n.tag == qn("w:t") and n.text
    )


def process_document(doc, dry_run=False):
    """处理整个文档，返回 (修改段落数, 修改引用数, 详情列表)。"""
    total_paras = 0
    total_refs = 0
    details = []

    for i, para in enumerate(doc.paragraphs):
        ft = full_text(para)

        # 跳过参考文献列表行
        if is_ref_list_line(ft):
            continue

        # 跳过没有引用的段落
        if not REF_PATTERN.search(ft):
            continue

        if dry_run:
            refs = REF_PATTERN.findall(ft)
            # 检查哪些还不是上标
            pending = []
            for r_elem in para._element.iter(qn("w:r")):
                if is_superscript(r_elem):
                    continue
                t = r_elem.find(qn("w:t"))
                if t is not None and t.text and REF_PATTERN.search(t.text):
                    pending.extend(REF_PATTERN.findall(t.text))
            # 也检查 w:ins 中的 run
            for ins in para._element.findall(qn("w:ins")):
                for r_elem in ins.findall(qn("w:r")):
                    if is_superscript(r_elem):
                        continue
                    t = r_elem.find(qn("w:t"))
                    if t is not None and t.text and REF_PATTERN.search(t.text):
                        pending.extend(REF_PATTERN.findall(t.text))
            if pending:
                total_paras += 1
                total_refs += len(pending)
                details.append((i, ft[:60], pending))
            continue

        para_count = 0

        # 处理正常 run
        runs_to_check = list(para._element.iter(qn("w:r")))
        for r_elem in runs_to_check:
            # 跳过已处理过的（parent 可能已变）
            if r_elem.getparent() is None:
                continue
            para_count += split_run(r_elem)

        if para_count > 0:
            total_paras += 1
            total_refs += para_count
            details.append((i, ft[:60], para_count))

    return total_paras, total_refs, details


def main():
    parser = argparse.ArgumentParser(
        description="修复 DOCX 文献引用角标（[数字] → 上标）"
    )
    parser.add_argument("input", help="输入 .docx 文件")
    parser.add_argument("-o", "--output", help="输出文件（默认覆盖输入）")
    parser.add_argument("--dry-run", action="store_true",
                        help="仅检查，不修改")
    args = parser.parse_args()

    inp = Path(args.input).expanduser()
    if not inp.exists():
        print(f"✗ 文件不存在: {inp}")
        sys.exit(1)

    doc = Document(str(inp))
    n_paras, n_refs, details = process_document(doc, dry_run=args.dry_run)

    if args.dry_run:
        print(f"扫描结果: {n_paras} 个段落, {n_refs} 处引用需要上标")
        for idx, text, refs in details:
            print(f"  段落{idx}: {text}... → {refs}")
        return

    if n_refs == 0:
        print("所有引用已是上标，无需修改。")
        return

    out = Path(args.output) if args.output else inp
    doc.save(str(out))

    print(f"✓ {n_paras} 个段落, {n_refs} 处引用已改为上标")
    for idx, text, count in details:
        print(f"  段落{idx}: {text}... ({count}处)")
    print(f"已保存: {out}")


if __name__ == "__main__":
    main()
