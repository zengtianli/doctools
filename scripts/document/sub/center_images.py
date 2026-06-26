#!/usr/bin/env python3
"""center_images — 把每个含图片的段落显式居中 + 零缩进（surgical · OLE 安全 · 可视化）。

Why（2026-06-25 天台 0624 排版踩坑）：
  靠 pStyle(ZDWP1)样式继承居中 → Word 与 LibreOffice 渲染不一致，**Word 里图就左对齐**。
  renumber-fig --fix-center 只补了无样式的图段(4/20)，套了 ZDWP1 的 16 段没显式居中。
  根治 = 每个含 <w:drawing>/<w:pict> 的段**显式打 jc=center + ind 清零**，不赌样式。

单功能脚本（你自己可跑）：
  python3 center_images.py x.docx --check          # 机器检查：报 N 张图段, 几张已显式居中
  python3 center_images.py x.docx --apply          # 修：全部显式居中+零缩进 (+.bak)
  python3 center_images.py x.docx --apply --no-backup
  python3 center_images.py x.docx --shot [--out-dir DIR]  # 可视化：把含图页渲染成 PNG 供眼检

surgical：只重写 word/document.xml，其余 zip 项 verbatim。
"""
from __future__ import annotations

import argparse
import shutil
import subprocess
import sys
import zipfile
from datetime import datetime
from pathlib import Path

from lxml import etree

W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
SOFFICE = "/Applications/LibreOffice.app/Contents/MacOS/soffice"


def q(t):
    return W + t


def _has_image(p) -> bool:
    return p.find(".//" + q("drawing")) is not None or p.find(".//" + q("pict")) is not None


def _is_centered(p) -> bool:
    pPr = p.find(q("pPr"))
    if pPr is None:
        return False
    jc = pPr.find(q("jc"))
    return jc is not None and jc.get(q("val")) == "center"


# CT_PPr 子元素 schema 顺序（相关子集）—— jc/ind 必须按此插入，否则 Word 拒读
_PPR_ORDER = [
    "pStyle", "keepNext", "keepLines", "pageBreakBefore", "framePr", "widowControl",
    "numPr", "suppressLineNumbers", "pBdr", "shd", "tabs", "suppressAutoHyphens",
    "kinsoku", "wordWrap", "overflowPunct", "topLinePunct", "autoSpaceDE",
    "autoSpaceDN", "bidi", "adjustRightInd", "snapToGrid", "spacing", "ind",
    "contextualSpacing", "mirrorIndents", "suppressOverlap", "jc", "textDirection",
    "textAlignment", "textboxTightWrap", "outlineLvl", "divId", "cnfStyle", "rPr",
    "sectPr", "pPrChange",
]
_ORDER_IDX = {q(t): i for i, t in enumerate(_PPR_ORDER)}


def _ensure_in_order(pPr, tag):
    """取/建 pPr 下 tag 子元素，按 schema 顺序就位。返回该元素。"""
    el = pPr.find(q(tag))
    if el is not None:
        return el
    el = etree.Element(q(tag))
    my = _ORDER_IDX.get(q(tag), 999)
    pos = len(pPr)
    for i, ch in enumerate(pPr):
        if _ORDER_IDX.get(ch.tag, 999) > my:
            pos = i
            break
    pPr.insert(pos, el)
    return el


def _root(docx):
    with zipfile.ZipFile(docx) as z:
        return etree.fromstring(z.read("word/document.xml"))


def _scan(docx):
    root = _root(docx)
    imgs = [p for p in root.iter(q("p")) if _has_image(p)]
    centered = [p for p in imgs if _is_centered(p)]
    return root, imgs, centered


def cmd_check(docx) -> int:
    _, imgs, centered = _scan(docx)
    print(f"[图片居中机检] {docx.name}")
    print(f"  含图片段总数      : {len(imgs)}")
    print(f"  已显式居中(jc=center): {len(centered)}")
    print(f"  未显式居中(赌样式) : {len(imgs) - len(centered)}")
    if len(centered) < len(imgs):
        print("✗ 有图片段未显式居中（Word 里可能左对齐）")
        return 2
    print("✓ 所有图片段已显式居中")
    return 0


def cmd_apply(docx, no_backup) -> int:
    root, imgs, _ = _scan(docx)
    fixed = 0
    for p in imgs:
        pPr = p.find(q("pPr"))
        if pPr is None:
            pPr = etree.Element(q("pPr"))
            p.insert(0, pPr)
        # jc=center（按 schema 顺序就位）
        jc = _ensure_in_order(pPr, "jc")
        jc.set(q("val"), "center")
        # ind 清零（去首行缩进/左缩进，按 schema 顺序就位）
        ind = _ensure_in_order(pPr, "ind")
        for k in ("firstLine", "firstLineChars", "left", "leftChars", "hanging"):
            if ind.get(q(k)) is not None:
                ind.set(q(k), "0")
        ind.set(q("firstLine"), "0")
        fixed += 1

    if fixed == 0:
        print(f"[图片居中] {docx.name}: 无图片段")
        return 0
    if not no_backup:
        bak = docx.with_suffix(docx.suffix + f".bak-{datetime.now():%Y%m%d-%H%M%S}")
        shutil.copy2(docx, bak)
        print(f"  备份 → {bak.name}")
    new_doc = etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True)
    tmp = docx.with_suffix(docx.suffix + ".tmp")
    with zipfile.ZipFile(docx) as zin, \
         zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as zout:
        for item in zin.infolist():
            data = new_doc if item.filename == "word/document.xml" else zin.read(item.filename)
            zout.writestr(item, data)
    tmp.replace(docx)
    print(f"[图片居中] {docx.name}: {fixed} 个图片段显式居中+零缩进")
    return 0


def _img_pages(pdf):
    """返回含图片的页码（按页面绘图对象多寡近似——这里用文本'图X'题注定位）。"""
    try:
        from pypdf import PdfReader
    except Exception:
        from PyPDF2 import PdfReader
    import re
    rd = PdfReader(str(pdf))
    pages = []
    for i, pg in enumerate(rd.pages, 1):
        t = pg.extract_text() or ""
        if re.search(r"图\s?\d+[\.．]\d+-\d+", t) or re.search(r"附图\s?\d+", t):
            pages.append(i)
    return pages


def cmd_shot(docx, out_dir) -> int:
    """可视化：docx → PDF → 含图页 PNG。"""
    out_dir = out_dir or docx.parent / (docx.stem + "_图检")
    out_dir.mkdir(parents=True, exist_ok=True)
    subprocess.run([SOFFICE, "--headless", "--convert-to", "pdf",
                    "--outdir", str(out_dir), str(docx)],
                   capture_output=True, timeout=180)
    pdf = out_dir / (docx.stem + ".pdf")
    if not pdf.exists():
        print("✗ LibreOffice 转 PDF 失败", file=sys.stderr)
        return 1
    pages = _img_pages(pdf)
    print(f"[图片可视化] {docx.name}: 含图页 {len(pages)} 页 → {out_dir}")
    for pg in pages:
        subprocess.run(["/opt/homebrew/bin/pdftoppm", "-png", "-f", str(pg), "-l", str(pg),
                        "-r", "80", str(pdf), str(out_dir / f"图检-p{pg:03d}")],
                       capture_output=True)
    print(f"  PNG 已出 {len(pages)} 张于 {out_dir}（打开眼检图是否居中）")
    return 0


def main(argv=None):
    ap = argparse.ArgumentParser(description="图片段显式居中+零缩进（单功能·可视化）")
    ap.add_argument("docx", type=Path)
    ap.add_argument("--check", action="store_true")
    ap.add_argument("--apply", action="store_true")
    ap.add_argument("--shot", action="store_true", help="渲染含图页为 PNG 供眼检")
    ap.add_argument("--out-dir", type=Path, default=None)
    ap.add_argument("--no-backup", action="store_true")
    a = ap.parse_args(argv)
    if not a.docx.exists():
        print(f"找不到: {a.docx}", file=sys.stderr)
        return 1
    if a.shot:
        return cmd_shot(a.docx, a.out_dir)
    if a.apply:
        return cmd_apply(a.docx, a.no_backup)
    return cmd_check(a.docx)


if __name__ == "__main__":
    sys.exit(main())
