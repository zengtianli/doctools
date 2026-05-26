#!/usr/bin/env python3
# distilled from qual-supply/scripts/audit_heading_numbers.py (2026-05-25 W1)
"""audit_heading_numbers.py — W2 read-only audit · heading 自动编号 freeze 前后状态评估

单功能扫:
1. 每个 H 段 (Heading 1/2/3/4) 的 idx / style / text(50) / 段级 numPr / 当前 prefix 形态
2. styles.xml 内 heading 1-4 的 numPr 状态 (exists/removed/numId/ilvl)
3. numbering.xml 内 heading 用到的 numId → abstractNumId → 每级 lvlText 模板

CLI: python3 scripts/audit_heading_numbers.py <docx> [--report <json>]

不改 docx。produce JSON to stdout (and --report path 同步落盘)。
"""
from __future__ import annotations

import argparse
import json
import re
import sys
import zipfile
from pathlib import Path

W_NS = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"

PREFIX_RE = re.compile(r"^\d+(?:\.\d+)*\s")


def _strip_ns(tag: str) -> str:
    return tag.split("}", 1)[-1] if "}" in tag else tag


def _read_xml(z: zipfile.ZipFile, name: str) -> str:
    with z.open(name) as f:
        return f.read().decode("utf-8")


def _para_text(p_xml: str) -> str:
    """concat all w:t in paragraph xml string"""
    texts = re.findall(r"<w:t[^>]*>([^<]*)</w:t>", p_xml)
    return "".join(texts)


def _para_style(p_xml: str) -> str | None:
    m = re.search(r'<w:pStyle w:val="([^"]+)"', p_xml)
    return m.group(1) if m else None


def _para_numpr(p_xml: str) -> tuple[bool, str | None, str | None]:
    """returns (has_numPr, numId, ilvl) at paragraph level only"""
    # only look in <w:pPr> .. </w:pPr>
    ppr_m = re.search(r"<w:pPr>(.*?)</w:pPr>", p_xml, re.DOTALL)
    if not ppr_m:
        return (False, None, None)
    ppr = ppr_m.group(1)
    if "<w:numPr>" not in ppr and "<w:numPr/>" not in ppr:
        return (False, None, None)
    numId_m = re.search(r'<w:numId w:val="(\d+)"', ppr)
    ilvl_m = re.search(r'<w:ilvl w:val="(\d+)"', ppr)
    return (True, numId_m.group(1) if numId_m else None, ilvl_m.group(1) if ilvl_m else None)


def _split_paragraphs(doc_xml: str) -> list[str]:
    # naive splitter on top-level w:p; good enough for read-only flat scan
    paras = []
    i = 0
    while True:
        start = doc_xml.find("<w:p ", i)
        s2 = doc_xml.find("<w:p>", i)
        if start < 0 and s2 < 0:
            break
        if start < 0 or (s2 >= 0 and s2 < start):
            start = s2
        # find matching </w:p> (no nested w:p in OOXML)
        end = doc_xml.find("</w:p>", start)
        if end < 0:
            break
        paras.append(doc_xml[start:end + 6])
        i = end + 6
    return paras


HEADING_STYLE_MAP = {
    "1": "Heading 1",
    "2": "Heading 2",
    "3": "Heading 3",
    "4": "Heading 4",
    "Heading1": "Heading 1",
    "Heading2": "Heading 2",
    "Heading3": "Heading 3",
    "Heading4": "Heading 4",
}


def _normalize_style(style_id: str | None) -> str | None:
    if style_id is None:
        return None
    return HEADING_STYLE_MAP.get(style_id, style_id)


def _scan_styles(styles_xml: str) -> dict:
    """find heading 1-4 style blocks; report numPr status"""
    out: dict[str, dict] = {}
    for sid in ["1", "2", "3", "4", "Heading1", "Heading2", "Heading3", "Heading4"]:
        pat = f'<w:style w:type="paragraph" w:styleId="{sid}"'
        idx = styles_xml.find(pat)
        if idx < 0:
            # fallback: any w:style w:styleId="{sid}"
            idx = styles_xml.find(f'w:styleId="{sid}"')
            if idx < 0:
                continue
            idx = styles_xml.rfind("<w:style ", 0, idx)
            if idx < 0:
                continue
        end = styles_xml.find("</w:style>", idx)
        block = styles_xml[idx:end + 10]
        name_m = re.search(r'<w:name w:val="([^"]+)"', block)
        name = name_m.group(1) if name_m else f"styleId={sid}"
        if not name.lower().startswith("heading"):
            continue
        norm = _normalize_style(name) or name
        if norm in out:
            continue
        has_numpr = "<w:numPr>" in block or "<w:numPr/>" in block
        numId_m = re.search(r'<w:numPr>.*?<w:numId w:val="(\d+)"', block, re.DOTALL)
        ilvl_m = re.search(r'<w:numPr>.*?<w:ilvl w:val="(\d+)"', block, re.DOTALL)
        if has_numpr:
            status = f"exists numId={numId_m.group(1) if numId_m else None} ilvl={ilvl_m.group(1) if ilvl_m else None}"
        else:
            status = "removed (no w:numPr)"
        out[norm.lower()] = {
            "styleId": sid,
            "name": name,
            "has_numpr": has_numpr,
            "numId": numId_m.group(1) if numId_m else None,
            "ilvl": ilvl_m.group(1) if ilvl_m else None,
            "status": status,
        }
    return out


def _scan_numbering(numbering_xml: str, num_ids: list[str]) -> dict:
    """for each numId, resolve abstractNumId and dump lvlText per ilvl 0..4"""
    out: dict = {}
    for num_id in num_ids:
        if num_id is None or num_id in out:
            continue
        m = re.search(rf'<w:num w:numId="{num_id}"[^>]*>(.*?)</w:num>', numbering_xml, re.DOTALL)
        if not m:
            out[num_id] = {"error": "numId not found"}
            continue
        abs_m = re.search(r'<w:abstractNumId w:val="(\d+)"', m.group(1))
        if not abs_m:
            out[num_id] = {"error": "abstractNumId not found"}
            continue
        abs_id = abs_m.group(1)
        start = numbering_xml.find(f'<w:abstractNum w:abstractNumId="{abs_id}"')
        end = numbering_xml.find("</w:abstractNum>", start)
        block = numbering_xml[start:end + 16] if start >= 0 else ""
        lvltexts: dict[str, dict] = {}
        for ilvl in range(5):
            lvl_start = block.find(f'<w:lvl w:ilvl="{ilvl}"')
            if lvl_start < 0:
                continue
            lvl_end = block.find("</w:lvl>", lvl_start)
            lvl_block = block[lvl_start:lvl_end]
            lt = re.search(r'<w:lvlText w:val="([^"]*)"', lvl_block)
            nfmt = re.search(r'<w:numFmt w:val="([^"]+)"', lvl_block)
            lvltexts[str(ilvl)] = {
                "numFmt": nfmt.group(1) if nfmt else None,
                "lvlText": lt.group(1) if lt else None,
            }
        out[num_id] = {"abstractNumId": abs_id, "lvltexts": lvltexts}
    return out


def audit(docx_path: Path) -> dict:
    with zipfile.ZipFile(docx_path) as z:
        doc_xml = _read_xml(z, "word/document.xml")
        styles_xml = _read_xml(z, "word/styles.xml")
        try:
            numbering_xml = _read_xml(z, "word/numbering.xml")
        except KeyError:
            numbering_xml = ""

    paras = _split_paragraphs(doc_xml)

    h_count_by_level: dict[str, int] = {}
    h_with_prefix = 0
    h_without_prefix = 0
    h_with_para_numpr = 0
    samples_no_prefix: list[list] = []
    samples_with_prefix: list[list] = []
    h_details: list[dict] = []

    for idx, p in enumerate(paras):
        style_raw = _para_style(p)
        style_norm = _normalize_style(style_raw)
        if style_norm not in {"Heading 1", "Heading 2", "Heading 3", "Heading 4"}:
            continue
        text = _para_text(p)
        text50 = text[:50]
        has_npr, p_numId, p_ilvl = _para_numpr(p)
        has_prefix = bool(PREFIX_RE.match(text))
        prefix_literal = ""
        if has_prefix:
            m = PREFIX_RE.match(text)
            if m:
                prefix_literal = m.group(0).strip()
        h_count_by_level[style_norm] = h_count_by_level.get(style_norm, 0) + 1
        if has_prefix:
            h_with_prefix += 1
            if len(samples_with_prefix) < 10:
                samples_with_prefix.append([idx, style_norm, prefix_literal, text50])
        else:
            h_without_prefix += 1
            if len(samples_no_prefix) < 10:
                samples_no_prefix.append([idx, style_norm, text50])
        if has_npr:
            h_with_para_numpr += 1
        h_details.append({
            "idx": idx,
            "style": style_norm,
            "text50": text50,
            "para_numpr": has_npr,
            "para_numId": p_numId,
            "para_ilvl": p_ilvl,
            "has_prefix": has_prefix,
            "prefix": prefix_literal,
        })

    styles_status = _scan_styles(styles_xml)

    # numIds used by heading styles
    heading_num_ids = sorted({
        v.get("numId") for v in styles_status.values() if v.get("numId")
    } | {"4"})  # always inspect numId=4

    numbering_status = _scan_numbering(numbering_xml, list(heading_num_ids)) if numbering_xml else {}

    # styles_numpr_status flat
    flat_styles_status = {}
    for k, v in styles_status.items():
        flat_styles_status[v["name"]] = v["status"]

    return {
        "docx": str(docx_path),
        "total_h_paragraphs": h_with_prefix + h_without_prefix,
        "h_count_by_level": h_count_by_level,
        "h_with_prefix": h_with_prefix,
        "h_without_prefix": h_without_prefix,
        "h_with_para_numpr": h_with_para_numpr,
        "styles_numpr_status": flat_styles_status,
        "numbering_lvltext": {
            num_id: v.get("lvltexts", {}) for num_id, v in numbering_status.items()
        },
        "samples_no_prefix": samples_no_prefix,
        "samples_with_prefix": samples_with_prefix,
        "h_details": h_details,
    }


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("docx", type=Path)
    ap.add_argument("--report", type=Path, default=None)
    args = ap.parse_args()
    if not args.docx.exists():
        print(f"ERR: {args.docx} not found", file=sys.stderr)
        sys.exit(2)
    result = audit(args.docx)
    j = json.dumps(result, ensure_ascii=False, indent=2)
    if args.report:
        args.report.write_text(j, encoding="utf-8")
    # summary to stdout
    summary = {k: v for k, v in result.items() if k != "h_details"}
    print(json.dumps(summary, ensure_ascii=False, indent=2))


# ---------------- pipeline adapter ----------------
def apply_path(docx_path, args=None) -> dict:
    return audit(Path(docx_path))


if __name__ == "__main__":
    main()
