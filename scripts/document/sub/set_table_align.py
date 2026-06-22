#!/usr/bin/env python3
"""set_table_align — 把 docx 内所有表格整体在页面水平居中。

Why：
  Word 表格的「整体居中」由表级 <w:tblPr>/<w:jc w:val="center"/> 控制（表作为块
  在页面左右居中），与单元格内文字对齐（段落 jc / tcPr vAlign）是两回事。
  pandoc 生成的表默认 jc=left（或缺省）→ 表靠左，看着不居中。本命令统一写
  tblPr 的 jc=center。

  可选 --cell-center：同时把每个单元格内段落水平居中（段落 jc=center）+ 垂直居中
  （tcPr vAlign=center），用于表内容也要居中的场景。

遍历所有 <w:tbl>（含嵌套表）。原地修改，默认先备份 .bak-时间戳。

Usage:
  python3 set_table_align.py 报告.docx                  # 表格整体居中 + 自动备份
  python3 set_table_align.py 报告.docx --cell-center    # 同时单元格内文字居中
  python3 set_table_align.py 报告.docx --dry-run        # 只报告不写盘
  python3 set_table_align.py 报告.docx --no-backup
"""
from __future__ import annotations

import argparse
import shutil
import sys
from datetime import datetime
from pathlib import Path

from docx import Document
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

# tblPr 子元素顺序（schema 序）：jc 须排在这些之前
_TBLPR_AFTER = ("tblBorders", "shd", "tblLayout", "tblCellMar", "tblLook",
                "tblCaption", "tblDescription")
# pPr 子元素顺序：jc 须排在这些之前（足够覆盖常见情况）
_PPR_AFTER = ("rPr",)


def _insert_in_order(parent, child, after_tags):
    """把 child 插到 parent 中第一个属于 after_tags 的元素之前，否则追加到末尾。"""
    after_qn = {qn(f"w:{t}") for t in after_tags}
    for existing in parent:
        if existing.tag in after_qn:
            existing.addprevious(child)
            return
    parent.append(child)


def _set_table_jc(tbl_el, val: str = "center") -> None:
    """设/重置表级 tblPr/jc。tblPr 是 tbl 的必有首子元素。"""
    tblPr = tbl_el.find(qn("w:tblPr"))
    if tblPr is None:
        tblPr = OxmlElement("w:tblPr")
        tbl_el.insert(0, tblPr)
    jc = tblPr.find(qn("w:jc"))
    if jc is None:
        jc = OxmlElement("w:jc")
        _insert_in_order(tblPr, jc, _TBLPR_AFTER)
    jc.set(qn("w:val"), val)


def _center_cells(tbl_el) -> int:
    """单元格内段落水平居中 + 单元格垂直居中。返回处理的单元格数。"""
    n = 0
    for tc in tbl_el.iter(qn("w:tc")):
        # 垂直居中：tcPr/vAlign=center
        tcPr = tc.find(qn("w:tcPr"))
        if tcPr is None:
            tcPr = OxmlElement("w:tcPr")
            tc.insert(0, tcPr)
        v = tcPr.find(qn("w:vAlign"))
        if v is None:
            v = OxmlElement("w:vAlign")
            tcPr.append(v)
        v.set(qn("w:val"), "center")
        # 水平居中：每个直属段落 pPr/jc=center
        for p in tc.findall(qn("w:p")):
            pPr = p.find(qn("w:pPr"))
            if pPr is None:
                pPr = OxmlElement("w:pPr")
                p.insert(0, pPr)
            jc = pPr.find(qn("w:jc"))
            if jc is None:
                jc = OxmlElement("w:jc")
                _insert_in_order(pPr, jc, _PPR_AFTER)
            jc.set(qn("w:val"), "center")
        n += 1
    return n


def _all_tbl_elements(doc):
    """所有 <w:tbl>，含单元格内嵌套表。"""
    return list(doc.element.body.iter(qn("w:tbl")))


def process(docx_path: Path, cell_center: bool, dry_run: bool, backup: bool) -> dict:
    doc = Document(str(docx_path))
    tbls = _all_tbl_elements(doc)
    n_tbl = len(tbls)
    cells = 0
    for tbl_el in tbls:
        _set_table_jc(tbl_el, "center")
        if cell_center:
            cells += _center_cells(tbl_el)

    result = {
        "file": str(docx_path),
        "tables": n_tbl,
        "cell_center": cell_center,
        "cells": cells,
        "dry_run": dry_run,
    }
    if dry_run or n_tbl == 0:
        result["written"] = False
        return result

    if backup:
        bak = docx_path.with_name(
            docx_path.name + ".bak-" + datetime.now().strftime("%Y%m%d-%H%M%S"))
        shutil.copy2(docx_path, bak)
        result["backup"] = str(bak)

    doc.save(str(docx_path))
    result["written"] = True

    # 验证：重读，统计表级 jc=center 的表数
    doc2 = Document(str(docx_path))
    centered = 0
    for tbl_el in _all_tbl_elements(doc2):
        tblPr = tbl_el.find(qn("w:tblPr"))
        jc = tblPr.find(qn("w:jc")) if tblPr is not None else None
        if jc is not None and jc.get(qn("w:val")) == "center":
            centered += 1
    result["verify_centered"] = centered
    return result


def _fmt(r: dict) -> str:
    head = f"[table center] {Path(r['file']).name}"
    if r["tables"] == 0:
        return f"{head}: 无表格，跳过"
    tail = ""
    if r.get("written"):
        tail = f" → centered {r.get('verify_centered')}/{r['tables']}"
        if r["cell_center"]:
            tail += f" · cells={r['cells']}"
        if r.get("backup"):
            tail += f" · bak={Path(r['backup']).name}"
    elif r["dry_run"]:
        tail = " (dry-run)"
    return f"{head}: {r['tables']} 表 · 整体居中{tail}"


def main() -> int:
    ap = argparse.ArgumentParser(
        description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("docx_pos", nargs="?", help="(positional) docx 路径，等价 --docx")
    ap.add_argument("--docx", dest="docx_kw", help="目标 docx（原地修改）")
    ap.add_argument("--cell-center", action="store_true",
                    help="同时把单元格内文字水平+垂直居中（默认只表格整体居中）")
    ap.add_argument("--no-backup", action="store_true", help="不创建 .bak-时间戳 备份")
    ap.add_argument("--dry-run", action="store_true", help="只报告不写盘")
    args = ap.parse_args()

    docx = args.docx_kw or args.docx_pos
    if not docx:
        print("[table center] missing docx (positional or --docx)", file=sys.stderr)
        return 2
    docx_path = Path(docx)
    if not docx_path.exists():
        print(f"[table center] not found: {docx_path}", file=sys.stderr)
        return 2

    r = process(docx_path, cell_center=args.cell_center,
                dry_run=args.dry_run, backup=not args.no_backup)
    print(_fmt(r))
    return 0


# ---------------- pipeline adapter ----------------
def apply_path(docx_path, args=None) -> dict:
    """pipeline-compatible adapter（原地 mutator）。"""
    cell = bool(getattr(args, "cell_center", False)) if args else False
    dry = bool(getattr(args, "dry_run", False)) if args else False
    backup = not bool(getattr(args, "no_backup", False)) if args else True
    return process(Path(docx_path), cell, dry, backup)


if __name__ == "__main__":
    sys.exit(main())
