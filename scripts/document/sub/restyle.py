#!/usr/bin/env python3
"""restyle — 按同源 golden 把段落**完整格式**精确移植回「格式被扒光」的稿子（surgical · OLE 安全）。

Why（2026-06-25 天台 0624 vs 0625）：
  用户把 0624 的院格式全扒光做成扁平稿（pStyle 全空 + 直接格式也清掉，封面标题塌成
  正文小字左对齐），要求「排版成 0625 那样」。golden 0625 与 0624 内容 99% 同源、
  且保留完整院格式。
  ⚠ 只搬 pStyle 不够：院封面/声明/目录的大字居中靠**直接格式(pPr 对齐/间距 + run rPr
  字号/加粗)**不靠样式，pStyle-only 补不回封面 → 整本扒光稿的封面/特殊页全错（实测教训）。
  → 对**文本能匹配**的段落，整段克隆 golden 的 pPr + runs（内容一样，克隆即精确还原格式，
  含封面大字居中加粗）。装帧引擎 chrome 靠 pStyle 找章边界，restyle 是它的前置必需步。

策略（序列对齐 · full-clone 优先 · pStyle 兜底 · 保守不破媒体）：
  · 目标/golden 段序列按归一文本(strip + 压缩内部连续空白) difflib 对齐（内容同序，
    重复文本/表格数值/空段按位置正确配对，避免「唯一文本字典」漏掉 1000+ 多义段）。
  · equal 块里逐一配对 (目标段 ↔ golden 段)：两边都**不含媒体/特殊引用**(drawing/pict/
    object/OLE/公式/脚注/批注/超链接) → **整段克隆**golden 的 pPr+runs（最精确，含封面直接格式）。
  · 含媒体段（图片/公式）→ 不整段克隆（媒体 r:id 必须留目标件），退化为**只搬 pStyle**。
  · replace/insert/delete 块（内容真差异，如法人名/编写名单）→ 跳过，不动。

surgical：只重写 word/document.xml，其余 zip 项（媒体/embeddings/OLE/OMML 公式）verbatim。

Usage:
  python3 restyle.py target.docx --ref golden.docx --check          # 只读：报克隆/兜底/无源
  python3 restyle.py target.docx --ref golden.docx --apply          # 移植格式 + .bak
  python3 restyle.py target.docx --ref golden.docx --apply --no-backup
"""
from __future__ import annotations

import argparse
import re
import shutil
import sys
import zipfile
from copy import deepcopy
from datetime import datetime
from difflib import SequenceMatcher
from pathlib import Path

from lxml import etree

W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"


def q(tag: str) -> str:
    return W + tag


# 含这些子元素的段 = 带媒体/特殊引用，不整段克隆（r:id/embeddings 必须留目标件）
_OMML = "{http://schemas.openxmlformats.org/officeDocument/2006/math}"
_SPECIAL = {
    q("drawing"), q("pict"), q("object"),
    q("footnoteReference"), q("endnoteReference"),
    q("commentReference"), q("hyperlink"),
    _OMML + "oMath", _OMML + "oMathPara",
}


_WS = re.compile(r"\s+")


def _norm(s: str) -> str:
    return _WS.sub(" ", s).strip()


def _para_text(p) -> str:
    return "".join(t.text or "" for t in p.iter(q("t")))


def _pstyle(p):
    pPr = p.find(q("pPr"))
    if pPr is None:
        return None
    st = pPr.find(q("pStyle"))
    return st.get(q("val")) if st is not None else None


def _has_special(p) -> bool:
    for el in p.iter():
        if el.tag in _SPECIAL:
            return True
    return False


def _body_root(path: Path):
    with zipfile.ZipFile(path) as z:
        return etree.fromstring(z.read("word/document.xml"))


def _paras(root):
    """body 直接 + 嵌套的所有 <w:p>，连同归一文本。"""
    body = root.find(q("body"))
    ps = list(body.iter(q("p")))
    return ps, [_norm(_para_text(p)) for p in ps]


def _align(target: Path, ref: Path):
    """序列对齐目标段 vs golden 段。归类：clone[(p,gp)] / restyle[(p,style)] / 计数。"""
    troot = _body_root(target)
    groot = _body_root(ref)
    tps, ttx = _paras(troot)
    gps, gtx = _paras(groot)
    sm = SequenceMatcher(None, ttx, gtx, autojunk=False)
    clone, restyle = [], []
    kept = diff = 0
    for tag, i1, i2, j1, j2 in sm.get_opcodes():
        if tag != "equal":
            diff += (i2 - i1)            # 内容真差异(法人名/编写名单)，不动
            continue
        for k in range(i2 - i1):
            p, gp = tps[i1 + k], gps[j1 + k]
            if not _has_special(p) and not _has_special(gp):
                clone.append((p, gp))     # 整段克隆（最精确，含封面直接格式）
            else:
                # 含媒体/公式 → 只搬 pStyle（媒体 r:id 必须留目标件）
                if _pstyle(p) is None and _pstyle(gp):
                    restyle.append((p, _pstyle(gp)))
                else:
                    kept += 1
    return troot, clone, restyle, kept, diff


def cmd_check(target: Path, ref: Path) -> int:
    _, clone, restyle, kept, diff = _align(target, ref)
    print(f"[restyle 机检 · 对照 {ref.name}] {target.name}")
    print(f"  整段克隆格式（最精确）    : {len(clone)}")
    print(f"  仅搬 pStyle（含媒体段兜底）: {len(restyle)}")
    print(f"  跳过（已有样式/媒体无源）  : {kept}")
    print(f"  内容真差异段（不动）      : {diff}")
    if clone or restyle:
        print("✗ 有段落格式待移植（格式被扒）")
        return 2
    print("✓ 无待移植段")
    return 0


def cmd_apply(target: Path, ref: Path, no_backup: bool) -> int:
    root, clone, restyle, kept, diff = _align(target, ref)
    if not clone and not restyle:
        print(f"[restyle] {target.name}: 无需修改（kept={kept} 差异={diff}）")
        return 0

    # ① 整段克隆：保留目标 <w:p> 元素本身，把子节点换成 golden 段的深拷贝。
    #    ⚠ 剥掉 pPr 内的 sectPr —— 节断引用 golden 的页眉页脚 rId，克隆进目标件=悬空
    #    rId 致文档损坏；节结构是 chrome 的活，restyle 只搬段落/字符格式。
    for p, gp in clone:
        for ch in list(p):
            p.remove(ch)
        for ch in gp:
            c = deepcopy(ch)
            if c.tag == q("pPr"):
                for sect in c.findall(q("sectPr")):
                    c.remove(sect)
            p.append(c)

    # ② 含媒体段兜底：只补 pStyle
    for p, st in restyle:
        pPr = p.find(q("pPr"))
        if pPr is None:
            pPr = etree.Element(q("pPr"))
            p.insert(0, pPr)
        ps = etree.Element(q("pStyle"))
        ps.set(q("val"), st)
        pPr.insert(0, ps)

    if not no_backup:
        bak = target.with_suffix(target.suffix + f".bak-{datetime.now():%Y%m%d-%H%M%S}")
        shutil.copy2(target, bak)
        print(f"  备份 → {bak.name}")

    new_doc = etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True)
    tmp = target.with_suffix(target.suffix + ".tmp")
    with zipfile.ZipFile(target) as zin, \
         zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as zout:
        for item in zin.infolist():
            data = zin.read(item.filename)
            if item.filename == "word/document.xml":
                data = new_doc
            zout.writestr(item, data)
    tmp.replace(target)
    print(f"[restyle] {target.name}: 整段克隆格式 {len(clone)} 段 + 媒体段兜底 pStyle {len(restyle)} 段"
          f"（跳过 kept={kept} / 内容真差异 {diff} 段不动）")
    return 0


def main(argv=None):
    ap = argparse.ArgumentParser(description="按同源 golden 移植段落完整格式（surgical）")
    ap.add_argument("docx", type=Path)
    ap.add_argument("--ref", type=Path, required=True, help="同源 golden（格式源）")
    ap.add_argument("--check", action="store_true", help="只读机检, exit2=有待移植")
    ap.add_argument("--apply", action="store_true", help="移植格式")
    ap.add_argument("--no-backup", action="store_true")
    a = ap.parse_args(argv)
    if not a.docx.exists():
        print(f"找不到: {a.docx}", file=sys.stderr)
        return 1
    if not a.ref.exists():
        print(f"找不到参照: {a.ref}", file=sys.stderr)
        return 1
    if a.apply:
        return cmd_apply(a.docx, a.ref, a.no_backup)
    return cmd_check(a.docx, a.ref)


if __name__ == "__main__":
    sys.exit(main())
