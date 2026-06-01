#!/usr/bin/env python3
"""set_table_borders — 把 docx 内所有表格统一为「满格实线」边框。

Why（踩坑根因）：
  Word docx 表格边框有两级——表级 <w:tblBorders> 与单元格级 <w:tcBorders>，
  后者优先级更高。实务里常见某张表「看着不是全实线」其实是：表级没有 tblBorders、
  且部分单元格 tcBorders 把内部边（left/top 等）设成 val="nil"（无线），
  于是内部竖线/横线缺失，渲染成断断续续。光设表级边框盖不住单元格级 nil。

本引擎的「全实线」保证 = 两手抓：
  1. 每张表设表级 tblBorders：top/left/bottom/right/insideH/insideV 全 = single。
  2. 默认清掉每个单元格的 tcBorders（让表级统一生效）；或 --keep-cell-borders
     时只把非 single 的边（nil/none/dashed/dotted/double…）改写为 single。

遍历所有 <w:tbl>（含嵌套表）。原地修改，默认先备份 .bak-时间戳。

Usage:
  python3 set_table_borders.py 报告.docx                    # 原地 + 自动备份
  python3 set_table_borders.py --docx 报告.docx --sz 4 --color auto
  python3 set_table_borders.py 报告.docx --keep-cell-borders # 保留 tcBorders 仅改写非实线
  python3 set_table_borders.py 报告.docx --dry-run           # 只报告不写盘
  python3 set_table_borders.py 报告.docx --no-backup
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

EDGES = ("top", "left", "bottom", "right", "insideH", "insideV")
# tblPr 子元素中必须排在 tblBorders 之后者（用于定位插入点）
_TBLPR_AFTER = ("shd", "tblLayout", "tblCellMar", "tblLook",
                "tblCaption", "tblDescription")
# tcPr 子元素中必须排在 tcBorders 之后者
_TCPR_AFTER = ("shd", "noWrap", "tcMar", "textDirection", "tcFitText",
               "vAlign", "hideMark")


def _mk_border(tag: str, val: str, sz: int, color: str, space: int):
    el = OxmlElement(f"w:{tag}")
    el.set(qn("w:val"), val)
    el.set(qn("w:sz"), str(sz))
    el.set(qn("w:space"), str(space))
    el.set(qn("w:color"), color)
    return el


def _insert_in_order(parent, child, after_tags):
    """把 child 插到 parent 中第一个属于 after_tags 的元素之前，否则追加到末尾。"""
    after_qn = {qn(f"w:{t}") for t in after_tags}
    for existing in parent:
        if existing.tag in after_qn:
            existing.addprevious(child)
            return
    parent.append(child)


def _set_table_borders(tbl_el, val: str, sz: int, color: str, space: int) -> None:
    """设/重置表级 tblBorders 为指定 6 边。tblPr 是 tbl 的必有首子元素。"""
    tblPr = tbl_el.find(qn("w:tblPr"))
    if tblPr is None:  # 理论上不会发生（schema 必有）
        tblPr = OxmlElement("w:tblPr")
        tbl_el.insert(0, tblPr)
    old = tblPr.find(qn("w:tblBorders"))
    if old is not None:
        tblPr.remove(old)
    borders = OxmlElement("w:tblBorders")
    for edge in EDGES:
        borders.append(_mk_border(edge, val, sz, color, space))
    _insert_in_order(tblPr, borders, _TBLPR_AFTER)


def _strip_cell_borders(tbl_el) -> int:
    """删除每个单元格的 tcBorders（让表级统一生效）。返回删除个数。"""
    n = 0
    for tcPr in tbl_el.iter(qn("w:tcPr")):
        tcB = tcPr.find(qn("w:tcBorders"))
        if tcB is not None:
            tcPr.remove(tcB)
            n += 1
    return n


def _solidify_cell_borders(tbl_el, val: str, sz: int, color: str, space: int) -> int:
    """保留 tcBorders，但把任何非 single 的边（nil/none/dashed…）改写为实线。
    返回改写的边数。"""
    n = 0
    for tcB in tbl_el.iter(qn("w:tcBorders")):
        for edge in EDGES:
            el = tcB.find(qn(f"w:{edge}"))
            if el is None:
                continue
            if el.get(qn("w:val")) != val:
                el.set(qn("w:val"), val)
                el.set(qn("w:sz"), str(sz))
                el.set(qn("w:space"), str(space))
                el.set(qn("w:color"), color)
                n += 1
    return n


def _all_tbl_elements(doc):
    """所有 <w:tbl>，含单元格内嵌套表。"""
    return doc.element.body.iter(qn("w:tbl"))


def process(docx_path: Path, val: str, sz: int, color: str, space: int,
            keep_cell: bool, dry_run: bool, backup: bool) -> dict:
    doc = Document(str(docx_path))
    tbls = list(_all_tbl_elements(doc))
    n_tbl = len(tbls)
    cell_changed = 0
    for tbl_el in tbls:
        _set_table_borders(tbl_el, val, sz, color, space)
        if keep_cell:
            cell_changed += _solidify_cell_borders(tbl_el, val, sz, color, space)
        else:
            cell_changed += _strip_cell_borders(tbl_el)

    result = {
        "file": str(docx_path),
        "tables": n_tbl,
        "mode": "keep-cell" if keep_cell else "strip-cell",
        "cell_changed": cell_changed,
        "border": f"{val}/sz{sz}/{color}",
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

    # 验证：重读，统计表级 6 边齐全 + 残留 nil 边
    doc2 = Document(str(docx_path))
    full = 0
    residual_nil = 0
    for tbl_el in _all_tbl_elements(doc2):
        tblB = tbl_el.find(qn("w:tblPr"))
        tblB = tblB.find(qn("w:tblBorders")) if tblB is not None else None
        if tblB is not None and all(
            (e := tblB.find(qn(f"w:{edge}"))) is not None
            and e.get(qn("w:val")) == val for edge in EDGES
        ):
            full += 1
        for tcB in tbl_el.iter(qn("w:tcBorders")):
            for edge in EDGES:
                e = tcB.find(qn(f"w:{edge}"))
                if e is not None and e.get(qn("w:val")) in ("nil", "none"):
                    residual_nil += 1
    result["verify_full_grid"] = full
    result["verify_residual_nil"] = residual_nil
    return result


def _fmt(r: dict) -> str:
    head = f"[table borders] {Path(r['file']).name}"
    if r["tables"] == 0:
        return f"{head}: 无表格，跳过"
    tail = ""
    if r.get("written"):
        tail = (f" → full-grid {r.get('verify_full_grid')}/{r['tables']}"
                f" · residual-nil {r.get('verify_residual_nil')}")
        if r.get("backup"):
            tail += f" · bak={Path(r['backup']).name}"
    elif r["dry_run"]:
        tail = " (dry-run)"
    return (f"{head}: {r['tables']} 表 · {r['mode']} · cell_changed={r['cell_changed']}"
            f" · {r['border']}{tail}")


def main() -> int:
    ap = argparse.ArgumentParser(
        description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("docx_pos", nargs="?", help="(positional) docx 路径，等价 --docx")
    ap.add_argument("--docx", dest="docx_kw", help="目标 docx（原地修改）")
    ap.add_argument("--val", default="single",
                    help="边框线型 (single/double/dashed/...)，默认 single 实线")
    ap.add_argument("--sz", type=int, default=4,
                    help="线宽，单位 1/8 pt（4=0.5pt），默认 4")
    ap.add_argument("--color", default="auto", help="边框颜色，默认 auto（=黑）")
    ap.add_argument("--space", type=int, default=0, help="边距，默认 0")
    ap.add_argument("--keep-cell-borders", action="store_true",
                    help="保留单元格 tcBorders，仅把非实线边改写为实线（默认=直接删 tcBorders）")
    ap.add_argument("--no-backup", action="store_true", help="不创建 .bak-时间戳 备份")
    ap.add_argument("--dry-run", action="store_true", help="只报告不写盘")
    args = ap.parse_args()

    docx = args.docx_kw or args.docx_pos
    if not docx:
        print("[table borders] missing docx (positional or --docx)", file=sys.stderr)
        return 2
    docx_path = Path(docx)
    if not docx_path.exists():
        print(f"[table borders] not found: {docx_path}", file=sys.stderr)
        return 2

    r = process(docx_path, args.val, args.sz, args.color, args.space,
                keep_cell=args.keep_cell_borders, dry_run=args.dry_run,
                backup=not args.no_backup)
    print(_fmt(r))
    return 0


# ---------------- pipeline adapter ----------------
def apply_path(docx_path, args=None) -> dict:
    """pipeline-compatible adapter（原地 mutator）。"""
    val = getattr(args, "val", "single") if args else "single"
    sz = int(getattr(args, "sz", 4)) if args else 4
    color = getattr(args, "color", "auto") if args else "auto"
    space = int(getattr(args, "space", 0)) if args else 0
    keep = bool(getattr(args, "keep_cell_borders", False)) if args else False
    dry = bool(getattr(args, "dry_run", False)) if args else False
    backup = not bool(getattr(args, "no_backup", False)) if args else True
    return process(Path(docx_path), val, sz, color, space, keep, dry, backup)


if __name__ == "__main__":
    sys.exit(main())
