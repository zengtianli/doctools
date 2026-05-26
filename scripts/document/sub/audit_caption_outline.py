#!/usr/bin/env python3
# distilled from qual-supply/scripts/audit_caption_outline.py (2026-05-25 W1)
"""
audit_caption_outline.py — audit-only · 扫 docx 找 caption 段污染 outlineLvl + 错样式

单功能 audit-only:扫 body 段,识别 caption 段(`^(表|图)\s*\d+[.\-]\d+`,兼容 `1.2-1` 三段和
`1-1` 两段),统计 outlineLvl 污染 + style 错配。**不修改 docx**,只产 JSON 报告。

触发场景:
- W1 worker 给 caption 段套 Caption-family style + 删污染 outlineLvl 之前/之后做 before/after 对比
- 任何 docx 整合后图表标题段被错继承大纲层级时

CLI:
    python3 scripts/audit_caption_outline.py <docx> [--report <json>]

输出:控制台 summary + 可选 JSON 报告
"""
from __future__ import annotations

import argparse
import json
import re
import sys
from pathlib import Path
from collections import Counter

from docx import Document
from lxml import etree

W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
NSMAP = {"w": W_NS}

# 兼容 `图1.2-1` 三段、`表1-1` 两段、`表 1.2-1`、`图1.1` 两段(纯点)
CAPTION_RE = re.compile(r"^\s*(表|图)\s*\d+([\.\-]\d+){1,2}")

# Caption family styles — docx 里允许的 caption 类样式名(集团 ZDWP 命名 + 通用)
CAPTION_FAMILY_KEYWORDS = ("caption", "表名", "图名", "0图", "0表", "zdwp 表名", "zdwp 图名")


def is_caption_style(style_name: str) -> bool:
    if not style_name:
        return False
    low = style_name.lower().strip()
    for kw in CAPTION_FAMILY_KEYWORDS:
        if kw in low:
            return True
    return False


def get_outline_lvl(p_elem) -> str | None:
    """返回 paragraph 的 outlineLvl 值(字符串)或 None。"""
    pPr = p_elem.find(f"{{{W_NS}}}pPr")
    if pPr is None:
        return None
    ol = pPr.find(f"{{{W_NS}}}outlineLvl")
    if ol is None:
        return None
    return ol.get(f"{{{W_NS}}}val")


def get_style_name(p, doc) -> str:
    """返回 paragraph 的 style 名称(从 styles.xml 解析 styleId → name)。"""
    try:
        return p.style.name if p.style is not None else ""
    except Exception:
        return ""


def collect_available_caption_styles(doc) -> list[str]:
    """列 docx styles.xml 里所有 caption-family 样式名。"""
    result = []
    for s in doc.styles:
        try:
            if is_caption_style(s.name):
                result.append(s.name)
        except Exception:
            continue
    return sorted(set(result))


def _audit_from_doc(doc, docx_path_label: str = "") -> dict:
    body_paragraphs = doc.paragraphs

    total = len(body_paragraphs)
    captions_total = 0
    captions_with_outlinelvl: list[dict] = []
    captions_by_style: Counter = Counter()
    captions_clean_examples: list[dict] = []
    h_count: Counter = Counter()
    polluted_count = 0
    wrong_style_count = 0
    empty_caption = 0
    all_caption_records: list[dict] = []

    for idx, p in enumerate(body_paragraphs):
        style_name = get_style_name(p, doc)
        text = (p.text or "").strip()

        # H styles 统计
        if style_name and style_name.startswith("Heading "):
            h_count[style_name] += 1

        # caption 识别:文本匹配 caption 正则 OR style 已套 caption-family (但 text 是空的也算被识别为 caption)
        text_is_caption = bool(CAPTION_RE.match(text))
        style_is_caption = is_caption_style(style_name)

        if text_is_caption or style_is_caption:
            captions_total += 1
            outline_lvl = get_outline_lvl(p._p)
            captions_by_style[style_name or "(no-style)"] += 1

            record = {
                "idx": idx,
                "style": style_name,
                "outlineLvl": outline_lvl,
                "text": text[:80],
            }
            all_caption_records.append(record)

            # 空 caption(style 套了但文本空)
            if not text:
                empty_caption += 1

            # 污染:有 outlineLvl 且数值 <= 4(0-4 表示 H1-H5 级被错继承)
            if outline_lvl is not None:
                try:
                    lvl_int = int(outline_lvl)
                    if lvl_int <= 6:  # 0-6 都算污染(<=6 = 进 outline 大纲)
                        polluted_count += 1
                        if len(captions_with_outlinelvl) < 30:
                            captions_with_outlinelvl.append(record)
                except ValueError:
                    pass

            # style 错配:文本是 caption 形态但 style 不是 caption-family 也不是空(像 Normal/正文)
            if text_is_caption and not style_is_caption:
                wrong_style_count += 1

            # clean 样本:outline 清 + style 正确
            if outline_lvl is None and style_is_caption and len(captions_clean_examples) < 10:
                captions_clean_examples.append(record)

    h_styles_present = sorted(h_count.keys())
    available_caption_styles = collect_available_caption_styles(doc)

    return {
        "docx_path": docx_path_label,
        "total_paragraphs": total,
        "captions_total": captions_total,
        "captions_with_outlinelvl": captions_with_outlinelvl,
        "captions_by_style": dict(captions_by_style),
        "captions_clean_examples": captions_clean_examples,
        "h_styles_present": h_styles_present,
        "h_count": dict(h_count),
        "caption_styles_available": available_caption_styles,
        "all_caption_records": all_caption_records,  # 用于 sample 5 抽样
        "issues": {
            "polluted_outline_count": polluted_count,
            "wrong_style_count": wrong_style_count,
            "empty_caption": empty_caption,
        },
    }


def audit(docx_path: Path) -> dict:
    doc = Document(str(docx_path))
    return _audit_from_doc(doc, str(docx_path))


def apply(doc, args=None) -> dict:
    """pipeline read-only adapter"""
    label = str(getattr(args, "docx", "")) if args else ""
    return _audit_from_doc(doc, label)


def main():
    ap = argparse.ArgumentParser(description="audit-only · caption 段 outlineLvl + style 污染审计")
    ap.add_argument("docx", help="docx 路径")
    ap.add_argument("--report", help="JSON 报告输出路径(可选)")
    args = ap.parse_args()

    docx_path = Path(args.docx)
    if not docx_path.exists():
        print(f"ERROR: docx 不存在: {docx_path}", file=sys.stderr)
        sys.exit(1)

    report = audit(docx_path)

    # 控制台 summary
    print(f"=== audit_caption_outline · {docx_path.name} ===")
    print(f"total_paragraphs       : {report['total_paragraphs']}")
    print(f"captions_total         : {report['captions_total']}")
    print(f"polluted_outline_count : {report['issues']['polluted_outline_count']}")
    print(f"wrong_style_count      : {report['issues']['wrong_style_count']}")
    print(f"empty_caption          : {report['issues']['empty_caption']}")
    print(f"captions_by_style      : {report['captions_by_style']}")
    print(f"h_count                : {report['h_count']}")
    print(f"caption_styles_avail   : {report['caption_styles_available']}")

    if args.report:
        Path(args.report).write_text(json.dumps(report, ensure_ascii=False, indent=2))
        print(f"[report] {args.report}")


if __name__ == "__main__":
    main()
