#!/usr/bin/env python3
"""line_spacing — 正文段落固定行距规整（surgical · OLE/公式安全）。

Why（踩坑根因 · 2026-06-25 天台 0624 vs 0625 diff）：
  pandoc/半拉子排版导出的 docx，正文段常**没设固定行距**（pPr 无 w:spacing
  line/lineRule）→ Word 用默认单倍行距，与院范式「固定值 28 磅(line=560
  lineRule=exact)」不一致，191 段行距不齐。

本脚本只补「缺固定行距的正文段」：
  · 判定正文段 = 有 CJK 文本、pStyle 不含 heading/title/toc/caption/图/表。
  · 已有 explicit line spacing 的段 **跳过**（幂等）。
  · 目标 (line,lineRule)：--ref 给参照件则取其正文段众数；否则默认 560/exact。

surgical：只重写 word/document.xml，其余 zip 项（媒体/embeddings/OLE/OMML 公式）
**逐字节 verbatim**，CRC 全等 → 含公式章安全（不走 python-docx）。

Usage:
  python3 line_spacing.py report.docx --check                 # 只读机检, exit2=有缺
  python3 line_spacing.py report.docx --check --ref 正确.docx  # 对参照件机检
  python3 line_spacing.py report.docx --fix                   # 修(默认 560/exact)+.bak
  python3 line_spacing.py report.docx --fix --ref 正确.docx    # 修(行距值取自参照)
"""
from __future__ import annotations

import argparse
import shutil
import sys
import zipfile
from collections import Counter
from datetime import datetime
from pathlib import Path

from lxml import etree

W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"


def q(tag: str) -> str:
    return W + tag


_SKIP_STYLE_KEYS = ("heading", "title", "subtitle", "toc", "caption",
                    "标题", "题注", "图", "表", "目录")


def _para_text(p) -> str:
    return "".join(t.text or "" for t in p.iter(q("t")))


def _style(p) -> str:
    pPr = p.find(q("pPr"))
    if pPr is None:
        return ""
    st = pPr.find(q("pStyle"))
    return (st.get(q("val")) if st is not None else "") or ""


def _is_body(p) -> bool:
    """正文段：有 CJK 文本、样式不在 skip 列表。"""
    txt = _para_text(p).strip()
    if not txt:
        return False
    if not any("一" <= c <= "鿿" for c in txt):
        return False
    st = _style(p).lower()
    return not any(k in st for k in _SKIP_STYLE_KEYS)


def _explicit_line(p):
    """返回 (line, lineRule) 若该段已显式设固定/精确行距，否则 None。"""
    pPr = p.find(q("pPr"))
    if pPr is None:
        return None
    sp = pPr.find(q("spacing"))
    if sp is None:
        return None
    line = sp.get(q("line"))
    rule = sp.get(q("lineRule"))
    if line and rule in ("exact", "atLeast", "auto"):
        return (line, rule)
    if line:
        return (line, rule or "auto")
    return None


def _body_root(docx_path: Path):
    with zipfile.ZipFile(docx_path) as z:
        xml = z.read("word/document.xml")
    return etree.fromstring(xml)


def _ref_target(ref_path: Path):
    """参照件正文段固定行距的众数 (line, lineRule)。"""
    root = _body_root(ref_path)
    body = root.find(q("body"))
    c = Counter()
    for p in body.iter(q("p")):
        if not _is_body(p):
            continue
        ex = _explicit_line(p)
        if ex and ex[1] == "exact":
            c[ex] += 1
    if not c:
        return None
    return c.most_common(1)[0][0]


def _ref_spacing_map(ref_path: Path):
    """参照件 文本→显式exact行距 映射（只收 ref 中显式设了 exact 的正文段）。"""
    root = _body_root(ref_path)
    body = root.find(q("body"))
    m = {}
    for p in body.iter(q("p")):
        if not _is_body(p):
            continue
        ex = _explicit_line(p)
        if ex and ex[1] == "exact":
            m.setdefault(_para_text(p).strip(), ex)
    return m


def _scan_vs_ref(root, refmap):
    """ref-matched：返回 (可比段数, 缺失段[(text20,目标值)])。
    可比段 = 文本在 ref 且 ref 该段有显式 exact 行距。缺失 = 目标段没设。"""
    body = root.find(q("body"))
    comparable, missing = 0, []
    for p in body.iter(q("p")):
        if not _is_body(p):
            continue
        tx = _para_text(p).strip()
        tgt = refmap.get(tx)
        if tgt is None:
            continue
        comparable += 1
        if _explicit_line(p) is None:
            missing.append((tx[:20], tgt))
    return comparable, missing


def cmd_check(docx_path: Path, ref: Path | None) -> int:
    if ref is None:
        print("行距机检需 --ref 参照件（行距差异在直接格式层，须逐段对照 golden）",
              file=sys.stderr)
        return 1
    refmap = _ref_spacing_map(ref)
    root = _body_root(docx_path)
    comparable, missing = _scan_vs_ref(root, refmap)
    print(f"[行距机检 · 对照 {ref.name}] {docx_path.name}")
    print(f"  参照件设了固定行距的正文段: {len(refmap)}")
    print(f"  本件可比段（文本匹配）    : {comparable}")
    print(f"  缺固定行距（参照有本件无）: {len(missing)}")
    if missing:
        for s, t in missing[:8]:
            print(f"    · 「{s}」 应={t[0]}/{t[1]}")
        if len(missing) > 8:
            print(f"    …共 {len(missing)} 段")
        print("✗ 有正文段缺固定行距（对照参照）")
        return 2
    print("✓ 对照参照，正文段固定行距齐")
    return 0


def cmd_fix(docx_path: Path, ref: Path | None, no_backup: bool) -> int:
    if ref is None:
        print("行距修复需 --ref 参照件", file=sys.stderr)
        return 1
    refmap = _ref_spacing_map(ref)
    root = _body_root(docx_path)
    body = root.find(q("body"))
    fixed = 0
    for p in body.iter(q("p")):
        if not _is_body(p):
            continue
        tx = _para_text(p).strip()
        tgt = refmap.get(tx)
        if tgt is None or _explicit_line(p) is not None:
            continue
        line, rule = tgt
        pPr = p.find(q("pPr"))
        if pPr is None:
            pPr = etree.SubElement(p, q("pPr"))
            p.insert(0, pPr)
        sp = pPr.find(q("spacing"))
        if sp is None:
            sp = etree.SubElement(pPr, q("spacing"))
        sp.set(q("line"), line)
        sp.set(q("lineRule"), rule)
        fixed += 1

    if fixed == 0:
        print(f"[行距修复] {docx_path.name}: 无需修改（已全部设固定行距）")
        return 0

    if not no_backup:
        bak = docx_path.with_suffix(
            docx_path.suffix + f".bak-{datetime.now():%Y%m%d-%H%M%S}")
        shutil.copy2(docx_path, bak)
        print(f"  备份 → {bak.name}")

    new_doc = etree.tostring(root, xml_declaration=True,
                             encoding="UTF-8", standalone=True)
    # surgical 重打包：只替换 document.xml，其余项 verbatim
    tmp = docx_path.with_suffix(docx_path.suffix + ".tmp")
    with zipfile.ZipFile(docx_path) as zin, \
         zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as zout:
        for item in zin.infolist():
            data = zin.read(item.filename)
            if item.filename == "word/document.xml":
                data = new_doc
            zout.writestr(item, data)
    tmp.replace(docx_path)
    print(f"[行距修复] {docx_path.name}: 对照参照补固定行距于 {fixed} 段")
    return 0


def main(argv=None):
    ap = argparse.ArgumentParser(description="正文段固定行距规整（surgical）")
    ap.add_argument("docx", type=Path)
    ap.add_argument("--check", action="store_true", help="只读机检, exit2=有缺")
    ap.add_argument("--fix", action="store_true", help="补固定行距")
    ap.add_argument("--ref", type=Path, help="参照件（行距值取其众数）")
    ap.add_argument("--no-backup", action="store_true")
    a = ap.parse_args(argv)
    if not a.docx.exists():
        print(f"找不到: {a.docx}", file=sys.stderr)
        return 1
    if a.check or not a.fix:
        return cmd_check(a.docx, a.ref)
    return cmd_fix(a.docx, a.ref, a.no_backup)


if __name__ == "__main__":
    sys.exit(main())
