#!/usr/bin/env python3
# distilled from qual-supply/scripts/audit_word_fields.py (2026-05-25 W1)
"""audit_word_fields.py — audit-only 扫 docx 内 Word 字段域(complex + simple)。

单功能:统计 <w:fldChar>/<w:instrText>/<w:fldSimple>,解析 complex field
(begin/separate/end 配对)出 instr + result plain text,给出字段类型分布、
嵌套深度、所在 paragraph idx 列表。

CLI:
    python3 scripts/audit_word_fields.py <docx> [--include-headers] [--report <json>]

不改文件,只产 JSON 报告 + stdout 摘要。
"""
from __future__ import annotations

import argparse
import json
import re
import sys
import zipfile
from pathlib import Path
from typing import Any

from lxml import etree

W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
W = f"{{{W_NS}}}"
NSMAP = {"w": W_NS}


def _localname(tag: str) -> str:
    return tag.split("}", 1)[-1] if "}" in tag else tag


def _iter_paragraph_idx(root: etree._Element) -> dict[etree._Element, int]:
    """给文档每个 <w:p> 标注一个 idx(物理顺序)。"""
    idx_map: dict[etree._Element, int] = {}
    for i, p in enumerate(root.iter(f"{W}p")):
        idx_map[p] = i
    return idx_map


def _ancestor_para_idx(elem: etree._Element, para_idx_map: dict[etree._Element, int]) -> int:
    cur = elem
    while cur is not None:
        if cur.tag == f"{W}p":
            return para_idx_map.get(cur, -1)
        cur = cur.getparent()
    return -1


def _classify_instr(instr: str) -> str:
    """从 instrText 文本里取字段类型(如 TOC / PAGEREF / SEQ / STYLEREF / REF / DATE / TIME / HYPERLINK)。
    Special:开头是 '=' 视作 '='(formula)。"""
    s = instr.strip()
    if not s:
        return "EMPTY"
    if s.startswith("="):
        return "="
    m = re.match(r"^([A-Z][A-Z0-9_]*)", s)
    if m:
        return m.group(1)
    return "OTHER"


def _parse_complex_fields(root: etree._Element, para_idx_map: dict[etree._Element, int]) -> list[dict]:
    """按文档物理顺序遍历 fldChar,依 begin/separate/end 配对组装字段。
    支持嵌套(begin 栈),返回每个字段记录。"""
    fields: list[dict] = []
    stack: list[dict] = []  # 栈,每个元素 = 一个未闭合 field 的临时状态

    for elem in root.iter():
        tag = _localname(elem.tag)
        if tag == "fldChar":
            ftype = elem.get(f"{W}fldCharType")
            if ftype == "begin":
                stack.append({
                    "instr_parts": [],
                    "result_parts": [],
                    "phase": "instr",  # instr -> result(after separate)
                    "para_idx": _ancestor_para_idx(elem, para_idx_map),
                    "depth": len(stack),  # 0 = top-level
                })
            elif ftype == "separate":
                if stack:
                    stack[-1]["phase"] = "result"
            elif ftype == "end":
                if stack:
                    rec = stack.pop()
                    instr = "".join(rec["instr_parts"])
                    result = "".join(rec["result_parts"])
                    fields.append({
                        "kind": "complex",
                        "type": _classify_instr(instr),
                        "instr": instr.strip(),
                        "result_sample": result.strip()[:200],
                        "para_idx": rec["para_idx"],
                        "depth": rec["depth"],
                    })
        elif tag == "instrText":
            if stack:
                stack[-1]["instr_parts"].append(elem.text or "")
        elif tag == "t":
            # plain text — 在 result phase 时,把外层 field 的 result 累计
            if stack and stack[-1]["phase"] == "result":
                # 只算最内层的 result(避免外层重复累计 inner field 文本)
                stack[-1]["result_parts"].append(elem.text or "")

    # 处理未闭合(异常)
    while stack:
        rec = stack.pop()
        instr = "".join(rec["instr_parts"])
        fields.append({
            "kind": "complex-unclosed",
            "type": _classify_instr(instr),
            "instr": instr.strip(),
            "result_sample": "",
            "para_idx": rec["para_idx"],
            "depth": rec["depth"],
        })
    return fields


def _parse_simple_fields(root: etree._Element, para_idx_map: dict[etree._Element, int]) -> list[dict]:
    fields: list[dict] = []
    for elem in root.iter(f"{W}fldSimple"):
        instr = elem.get(f"{W}instr", "") or ""
        # result = 内部所有 w:t 拼接
        result = "".join((t.text or "") for t in elem.iter(f"{W}t"))
        fields.append({
            "kind": "simple",
            "type": _classify_instr(instr),
            "instr": instr.strip(),
            "result_sample": result.strip()[:200],
            "para_idx": _ancestor_para_idx(elem, para_idx_map),
            "depth": 0,
        })
    return fields


def _count_artifacts(root: etree._Element) -> dict[str, int]:
    return {
        "fldChar_count": sum(1 for _ in root.iter(f"{W}fldChar")),
        "instrText_count": sum(1 for _ in root.iter(f"{W}instrText")),
        "fldSimple_count": sum(1 for _ in root.iter(f"{W}fldSimple")),
    }


def _scan_one_xml(xml_bytes: bytes, xml_name: str) -> dict[str, Any]:
    parser = etree.XMLParser(huge_tree=True, recover=False)
    root = etree.fromstring(xml_bytes, parser=parser)
    para_idx_map = _iter_paragraph_idx(root)
    counts = _count_artifacts(root)
    complex_fields = _parse_complex_fields(root, para_idx_map)
    simple_fields = _parse_simple_fields(root, para_idx_map)
    all_fields = complex_fields + simple_fields

    type_dist: dict[str, int] = {}
    for f in all_fields:
        type_dist[f["type"]] = type_dist.get(f["type"], 0) + 1

    depths = [f.get("depth", 0) for f in complex_fields]
    return {
        "xml_name": xml_name,
        **counts,
        "type_distribution": type_dist,
        "nested_depth_max": (max(depths) if depths else 0),
        "field_total": len(all_fields),
        "fields": all_fields,
    }


def audit(docx_path: Path, include_headers: bool) -> dict[str, Any]:
    if not docx_path.exists():
        raise FileNotFoundError(docx_path)
    targets: list[str] = ["word/document.xml"]
    files_scanned: list[str] = []
    aggregate = {
        "fldChar_count": 0,
        "instrText_count": 0,
        "fldSimple_count": 0,
        "type_distribution": {},
        "nested_depth_max": 0,
        "field_total": 0,
        "fields": [],
        "per_file": [],
    }
    with zipfile.ZipFile(docx_path, "r") as z:
        names = set(z.namelist())
        if include_headers:
            for n in names:
                if n.startswith("word/header") and n.endswith(".xml"):
                    targets.append(n)
                if n.startswith("word/footer") and n.endswith(".xml"):
                    targets.append(n)
        for name in targets:
            if name not in names:
                continue
            data = z.read(name)
            try:
                res = _scan_one_xml(data, name)
            except etree.XMLSyntaxError as e:
                aggregate["per_file"].append({"xml_name": name, "error": str(e)})
                continue
            files_scanned.append(name)
            aggregate["fldChar_count"] += res["fldChar_count"]
            aggregate["instrText_count"] += res["instrText_count"]
            aggregate["fldSimple_count"] += res["fldSimple_count"]
            aggregate["field_total"] += res["field_total"]
            aggregate["nested_depth_max"] = max(aggregate["nested_depth_max"], res["nested_depth_max"])
            for k, v in res["type_distribution"].items():
                aggregate["type_distribution"][k] = aggregate["type_distribution"].get(k, 0) + v
            aggregate["fields"].extend(res["fields"])
            aggregate["per_file"].append({
                "xml_name": name,
                "fldChar_count": res["fldChar_count"],
                "instrText_count": res["instrText_count"],
                "fldSimple_count": res["fldSimple_count"],
                "field_total": res["field_total"],
                "type_distribution": res["type_distribution"],
            })
    aggregate["files_scanned"] = files_scanned
    aggregate["docx"] = str(docx_path)
    return aggregate


def main(argv: list[str] | None = None) -> int:
    ap = argparse.ArgumentParser(description="Audit Word field artifacts in a docx (read-only).")
    ap.add_argument("docx", type=Path)
    ap.add_argument("--include-headers", action="store_true",
                    help="Also scan word/header*.xml and word/footer*.xml")
    ap.add_argument("--report", type=Path, default=None, help="Write full JSON report to this path")
    args = ap.parse_args(argv)

    report = audit(args.docx, include_headers=args.include_headers)

    # stdout 摘要 (短)
    summary = {
        "docx": report["docx"],
        "files_scanned": report["files_scanned"],
        "fldChar_count": report["fldChar_count"],
        "instrText_count": report["instrText_count"],
        "fldSimple_count": report["fldSimple_count"],
        "field_total": report["field_total"],
        "nested_depth_max": report["nested_depth_max"],
        "type_distribution": report["type_distribution"],
    }
    print(json.dumps(summary, ensure_ascii=False, indent=2))

    if args.report:
        args.report.parent.mkdir(parents=True, exist_ok=True)
        # 全报告(含 fields 详情)
        with args.report.open("w", encoding="utf-8") as fp:
            json.dump(report, fp, ensure_ascii=False, indent=2)
    return 0


# ---------------- pipeline adapter ----------------
def apply_path(docx_path, args=None) -> dict:
    include_headers = bool(getattr(args, "include_headers", False)) if args else False
    return audit(Path(docx_path), include_headers=include_headers)


if __name__ == "__main__":
    sys.exit(main())
