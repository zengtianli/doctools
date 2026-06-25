#!/usr/bin/env python3
"""port_sections — 从同源 golden 精确移植节结构(分节+页眉页脚水印+横向岛+封面titlePg)。

Why（2026-06-25 天台 0624 vs 0625）：
  chrome 按自己的章约定(前言/第N章 pStyle=1)启发式重建装帧，复刻不出 golden 那份
  手工排的 19 节(数字章 pStyle=2 + 统一标题页眉 + evenAndOdd + 封面无页眉)。
  既然有内容同源 golden，直接把它的真实节结构按文本锚搬过来 = 1:1 复刻，最贴合。
  前置步：restyle 已把 body/封面段落格式克隆对（本引擎只管「节/页眉/页脚」层）。

机制（纯 zip + lxml · 媒体/OLE/公式 verbatim · 页眉自包含水印无需搬媒体）：
  · 复制 golden 所有 header*/footer* 部件(加前缀防撞名) + 各自 .rels + Content_Types
    override + document.xml.rels 关系，rId 重映射到目标件新号。
  · 收集 golden 顶层 body 的节断(段内 pPr/sectPr) + 末尾 body sectPr，深拷贝并按
    重映射改写 headerReference/footerReference 的 r:id。
  · difflib 对齐目标 vs golden 顶层 body 段(归一文本)；金段节断落点 → 映射到目标对应段
    (不在 equal 块=落在内容真差异段如编写名单 → 退化锚到最近的前一个对齐段，仍在同一
     前置组内，视觉无碍)。把改写后的 sectPr 注入目标段 pPr；末节 sectPr 设为 golden 末节。
  · 注入前先清掉目标件原有 sectPr(段内+末尾)，避免双节断。

Usage:
  python3 port_sections.py target.docx --ref golden.docx [--out OUT] [--no-backup]
  python3 port_sections.py target.docx --ref golden.docx --check     # 只读：报金节数/可锚数
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
R = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}"
RELNS = "http://schemas.openxmlformats.org/package/2006/relationships"
CTNS = "http://schemas.openxmlformats.org/package/2006/content-types"
HDR_CT = "application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"
FTR_CT = "application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml"
HDR_TYPE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/header"
FTR_TYPE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer"

_WS = re.compile(r"\s+")


def q(t):
    return W + t


def _norm(s):
    return _WS.sub(" ", s).strip()


def _ptext(p):
    return "".join(t.text or "" for t in p.iter(q("t")))


def _zget(z, name):
    return z.read(name)


def _root(z, name):
    return etree.fromstring(_zget(z, name))


def _top_paras(root):
    """body 的顶层 <w:p> 直接子节点（节断只可能在这一层）。"""
    body = root.find(q("body"))
    return [el for el in body if el.tag == q("p")]


def _golden_breaks(groot):
    """golden 顶层 body 节断：返回 [(top_para_index | None, sectPr_elem)]。
    None index = 末尾 body sectPr。顺序即文档顺序。"""
    body = groot.find(q("body"))
    tops = _top_paras(groot)
    idx = {id(p): i for i, p in enumerate(tops)}
    breaks = []
    for el in body:
        if el.tag == q("p"):
            pPr = el.find(q("pPr"))
            if pPr is not None:
                sect = pPr.find(q("sectPr"))
                if sect is not None:
                    breaks.append((idx[id(el)], sect))
        elif el.tag == q("sectPr"):
            breaks.append((None, el))
    return breaks


def _hf_rels(z):
    """golden document.xml.rels 里 header/footer 关系: [(rId, kind, target)]。"""
    rels = _zget(z, "word/_rels/document.xml.rels").decode()
    out = []
    for m in re.finditer(r'Id="(rId\d+)"[^>]*Type="[^"]*/(header|footer)"[^>]*Target="([^"]+)"', rels):
        out.append((m.group(1), m.group(2), m.group(3)))
    # 也兼容 Type 在 Target 之后的顺序
    for m in re.finditer(r'Id="(rId\d+)"[^>]*Target="([^"]+)"[^>]*Type="[^"]*/(header|footer)"', rels):
        out.append((m.group(1), m.group(3), m.group(2)))
    seen = {}
    for rid, kind, tgt in out:
        seen[rid] = (kind, tgt)
    return [(rid, k, t) for rid, (k, t) in seen.items()]


def _target_max_rid(z):
    rels = _zget(z, "word/_rels/document.xml.rels").decode()
    return max([int(m.group(1)) for m in re.finditer(r'Id="rId(\d+)"', rels)] or [0])


def main(argv=None):
    ap = argparse.ArgumentParser(description="从同源 golden 移植节结构")
    ap.add_argument("docx", type=Path)
    ap.add_argument("--ref", type=Path, required=True)
    ap.add_argument("--out", type=Path, default=None)
    ap.add_argument("--check", action="store_true")
    ap.add_argument("--no-backup", action="store_true")
    a = ap.parse_args(argv)
    if not a.docx.exists() or not a.ref.exists():
        print("找不到 docx 或 ref", file=sys.stderr)
        return 1

    zt = zipfile.ZipFile(a.docx)
    zg = zipfile.ZipFile(a.ref)
    troot = _root(zt, "word/document.xml")
    groot = _root(zg, "word/document.xml")

    breaks = _golden_breaks(groot)
    # 对齐顶层 body 段
    t_tops = _top_paras(troot)
    g_tops = _top_paras(groot)
    t_tx = [_norm(_ptext(p)) for p in t_tops]
    g_tx = [_norm(_ptext(p)) for p in g_tops]
    sm = SequenceMatcher(None, g_tx, t_tx, autojunk=False)
    ratio = sm.ratio()              # 同源度（0~1）
    # 非同源参照（别县范式文本对不上）→ 不动文档，避免乱锚节断把版面搞坏
    if ratio < 0.5:
        print(f"[port_sections] 非同源参照（同源度 {ratio:.2f} < 0.5），跳过节移植不动文档。"
              f"节结构移植仅适用内容同源 golden。", file=sys.stderr)
        return 0
    g2t = {}                       # golden top idx -> target top idx (equal 块)
    equal_g = []
    for tag, i1, i2, j1, j2 in sm.get_opcodes():
        if tag == "equal":
            for k in range(i2 - i1):
                g2t[i1 + k] = j1 + k
                equal_g.append(i1 + k)
    equal_g.sort()

    def resolve(gi):
        """golden top idx → target top idx（精确 or 最近前一个对齐段）。"""
        if gi in g2t:
            return g2t[gi]
        # 最近 ≤ gi 的对齐金段
        lo = [e for e in equal_g if e <= gi]
        if lo:
            return g2t[lo[-1]]
        hi = [e for e in equal_g if e > gi]
        if hi:
            return g2t[hi[0]]
        return None

    anchored = sum(1 for gi, _ in breaks if gi is not None and resolve(gi) is not None)
    if a.check:
        print(f"[port_sections 机检] {a.docx.name} ← {a.ref.name}")
        print(f"  golden 节数            : {len(breaks)}")
        print(f"  含段节断(非末节)       : {sum(1 for gi,_ in breaks if gi is not None)}")
        print(f"  可锚定到目标段         : {anchored}")
        print(f"  golden header/footer 部件: {len(_hf_rels(zg))}")
        return 0 if anchored else 2

    # ── 复制 golden header/footer 部件，重映射 rId ──
    next_rid = _target_max_rid(zt) + 1
    rid_map = {}            # golden rId -> target new rId
    new_parts = {}          # 新部件名 -> bytes
    new_partrels = {}       # 新 .rels 名 -> bytes
    ct_adds = []            # [(partname, content-type)]
    rel_adds = []           # [(new_rid, type, target)]
    g_names = set(zg.namelist())
    for rid, kind, tgt in _hf_rels(zg):
        src = "word/" + tgt.lstrip("/")
        if src not in g_names:
            continue
        n = re.search(r'(header|footer)(\d+)\.xml', tgt)
        newname = f"{kind}G{n.group(2)}.xml" if n else f"{kind}G{next_rid}.xml"
        new_rid = f"rId{next_rid}"; next_rid += 1
        rid_map[rid] = new_rid
        new_parts["word/" + newname] = _zget(zg, src)
        relsrc = f"word/_rels/{Path(tgt).name}.rels"
        if relsrc in g_names:
            new_partrels[f"word/_rels/{newname}.rels"] = _zget(zg, relsrc)
        ct_adds.append(("/word/" + newname, HDR_CT if kind == "header" else FTR_CT))
        rel_adds.append((new_rid, HDR_TYPE if kind == "header" else FTR_TYPE, newname))

    def rewrite_refs(sect):
        s = deepcopy(sect)
        for ref in s.findall(q("headerReference")) + s.findall(q("footerReference")):
            old = ref.get(R + "id")
            if old in rid_map:
                ref.set(R + "id", rid_map[old])
            else:
                ref.getparent().remove(ref)   # 引用了没搬的 rId → 去掉，避免悬空
        return s

    # ── 清目标原有 sectPr，注入 golden 节断 ──
    tbody = troot.find(q("body"))
    for p in t_tops:
        pPr = p.find(q("pPr"))
        if pPr is not None:
            for s in pPr.findall(q("sectPr")):
                pPr.remove(s)
    for s in tbody.findall(q("sectPr")):
        tbody.remove(s)

    placed = 0
    final_sect = None
    for gi, sect in breaks:
        new_sect = rewrite_refs(sect)
        if gi is None:
            final_sect = new_sect
            continue
        ti = resolve(gi)
        if ti is None:
            continue
        tp = t_tops[ti]
        pPr = tp.find(q("pPr"))
        if pPr is None:
            pPr = etree.SubElement(tp, q("pPr"))
            tp.insert(0, pPr)
        # sectPr 须为 pPr 末子元素
        pPr.append(new_sect)
        placed += 1
    if final_sect is not None:
        tbody.append(final_sect)

    # ── 改写 [Content_Types].xml + document.xml.rels ──
    ct = _root(zt, "[Content_Types].xml")
    existing_ct = {o.get("PartName") for o in ct.findall(f"{{{CTNS}}}Override")}
    for pn, cty in ct_adds:
        if pn not in existing_ct:
            o = etree.SubElement(ct, f"{{{CTNS}}}Override")
            o.set("PartName", pn); o.set("ContentType", cty)
    drels = _root(zt, "word/_rels/document.xml.rels")
    for new_rid, typ, tgt in rel_adds:
        r = etree.SubElement(drels, f"{{{RELNS}}}Relationship")
        r.set("Id", new_rid); r.set("Type", typ); r.set("Target", tgt)

    # ── 重打包 ──
    out = a.out or (a.docx if False else a.docx.with_name(a.docx.stem + "_sect.docx"))
    if not a.no_backup and a.out is None:
        pass
    new_doc = etree.tostring(troot, xml_declaration=True, encoding="UTF-8", standalone=True)
    new_ct = etree.tostring(ct, xml_declaration=True, encoding="UTF-8", standalone=True)
    new_dr = etree.tostring(drels, xml_declaration=True, encoding="UTF-8", standalone=True)
    repl = {"word/document.xml": new_doc, "[Content_Types].xml": new_ct,
            "word/_rels/document.xml.rels": new_dr}
    tmp = out.with_suffix(out.suffix + ".tmp")
    with zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as zout:
        for item in zt.infolist():
            zout.writestr(item, repl.get(item.filename, _zget(zt, item.filename)))
        for pn, data in {**new_parts, **new_partrels}.items():
            zout.writestr(pn, data)
    tmp.replace(out)
    print(f"[port_sections] {out.name}: 移植 golden {len(breaks)} 节"
          f"（注入段节断 {placed} + 末节{'1' if final_sect is not None else '0'}）"
          f"，复制 header/footer {len(new_parts)} 部件")
    return 0


if __name__ == "__main__":
    sys.exit(main())
