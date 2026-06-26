#!/usr/bin/env python3
"""sync_toc — 让目录(TOC)与同源 golden 完全一致（surgical · OLE 安全 · 可视化）。

Why（2026-06-25 天台 0624 排版踩坑·Word 文档比较「目录还是不一样」）：
  扒光格式的 0624 经 restyle 套回了正文，但 **目录字段块** 被 restyle 当「special
  (含 hyperlink/PAGEREF 字段)」跳过 → 目录条目保留了扒光后的直接格式（无点导引 tab、
  无加粗、无 Times New Roman 28、缩进错），且字段缓存条目仍引用旧书签锚点 `_Toc231…`
  （正文里已被 restyle 换成 golden 的 `_Toc233…` → 19/19 锚点全悬空）。
  更深一层：扒光稿的 styles.xml 里 styleId 21/31/5/TOC1 与 golden 语义冲突
  （golden 21=toc2 段落样式；mine 21=正文首行缩进**字符**样式）→ 目录条目 pStyle=21
  指向字符样式被 Word 忽略 = 退化成正文。

根治（两道·都 surgical 改 xml + verbatim 重打包）：
  ① 样式表对账：mine 文中被引用、但与 golden 同 ID 不同义(名/类型)的 styleId
     → 用 golden 的定义覆盖(并补齐 golden 定义依赖的 basedOn/link/next)。
     （安全前提：这些 ID 在 mine 里只被 golden-克隆内容引用，无其它 rStyle/交叉引用。）
  ② 目录块移植：把 golden 的「目录标题段 + TOC 字段块」整段克隆进 mine，替换原块。
     带回 golden 的直接 tab(点导引)/rPr(加粗/TNR28)/缩进 + golden 锚点(_Toc233,
     在 mine 正文 19/19 可解析)。

单功能脚本（你自己可跑）：
  python3 sync_toc.py mine.docx --ref golden.docx --check          # 机器检查：报每条不一致
  python3 sync_toc.py mine.docx --ref golden.docx --apply          # 修(+.bak)
  python3 sync_toc.py mine.docx --ref golden.docx --apply --no-backup
  python3 sync_toc.py mine.docx --shot [--out-dir DIR]            # 可视化：渲染目录页 PNG 眼检

surgical：只重写 word/document.xml + word/styles.xml，其余 zip 项 verbatim。禁 python-docx。
"""
from __future__ import annotations

import argparse
import re
import shutil
import subprocess
import sys
import zipfile
from copy import deepcopy
from datetime import datetime
from pathlib import Path

from lxml import etree

W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
SOFFICE = "/Applications/LibreOffice.app/Contents/MacOS/soffice"


def q(t):
    return W + t


def _text(p):
    return "".join(t.text or "" for t in p.iter(q("t")))


def _instr(p):
    return "".join(t.text or "" for t in p.iter(q("instrText")))


def _pstyle(p):
    pPr = p.find(q("pPr"))
    if pPr is None:
        return None
    ps = pPr.find(q("pStyle"))
    return ps.get(q("val")) if ps is not None else None


def _load(docx):
    with zipfile.ZipFile(docx) as z:
        droot = etree.fromstring(z.read("word/document.xml"))
        sroot = etree.fromstring(z.read("word/styles.xml"))
    return droot, sroot


# ── 样式表索引 ────────────────────────────────────────────────────────────
def _style_index(sroot):
    """styleId -> (name, type, element)。"""
    idx = {}
    for s in sroot.iter(q("style")):
        sid = s.get(q("styleId"))
        nm = s.find(q("name"))
        idx[sid] = (nm.get(q("val")) if nm is not None else "", s.get(q("type")), s)
    return idx


def _referenced_styleids(droot):
    ids = set()
    for ps in droot.iter(q("pStyle")):
        ids.add(ps.get(q("val")))
    for rs in droot.iter(q("rStyle")):
        ids.add(rs.get(q("val")))
    return ids


# ── 目录区域定位（整个 TOC <w:sdt> 内容控件，含标题+字段+条目）──────────────
def _toc_sdt(droot):
    """返回包裹主目录(TOC \\o)的 <w:sdt> 元素；找不到返回 None。

    现代 Word 把目录包在 sdt/sdtContent 内容控件里（含「目录」标题段 + 字段 + 条目）。
    """
    for p in droot.iter(q("p")):
        if "TOC \\o" in _instr(p):
            e = p
            while e is not None and e.tag != q("sdt"):
                e = e.getparent()
            return e
    return None


def _toc_entries(droot):
    """目录条目段列表（sdt 内、含 PAGEREF 的段）：[(style, text)]。"""
    sdt = _toc_sdt(droot)
    if sdt is None:
        return []
    return [(_pstyle(el), _text(el).strip())
            for el in sdt.iter(q("p")) if "PAGEREF" in _instr(el)]


def _toc_anchors(droot):
    anch = []
    for it in droot.iter(q("instrText")):
        anch += re.findall(r"PAGEREF\s+(_Toc\d+)", it.text or "")
    return anch


def _body_bookmarks(droot):
    return {b.get(q("name")) for b in droot.iter(q("bookmarkStart")) if (b.get(q("name")) or "").startswith("_Toc")}


def _has_leader_tab(droot):
    """目录块内是否存在点导引 tab（golden 的视觉特征）。"""
    sdt = _toc_sdt(droot)
    if sdt is None:
        return False
    for tab in sdt.iter(q("tab")):
        if tab.get(q("leader")) == "dot":
            return True
    return False


# ── --check ────────────────────────────────────────────────────────────────
def cmd_check(mine, golden):
    md, ms = _load(mine)
    gd, gs = _load(golden)
    g_idx, m_idx = _style_index(gs), _style_index(ms)

    me, ge = _toc_entries(md), _toc_entries(gd)
    print(f"[目录机检] {mine.name}  vs  golden {golden.name}")
    ok = True

    # 1) 条目数 + 文本
    print(f"  条目数: mine {len(me)} / golden {len(ge)}")
    if [t for _, t in me] != [t for _, t in ge]:
        print("  ✗ 条目文本不一致")
        ok = False
        n = max(len(me), len(ge))
        for i in range(n):
            mt = me[i][1] if i < len(me) else "∅"
            gt = ge[i][1] if i < len(ge) else "∅"
            if mt != gt:
                print(f"      [{i}] mine={mt!r}  golden={gt!r}")
    else:
        print("  ✓ 条目文本逐条一致")

    # 2) 每条样式「名」对齐（Word 比较按样式名判异同）
    def sname(idx, sid):
        return idx.get(sid, ("∅", "?", None))[0]
    mismatch = []
    for i in range(min(len(me), len(ge))):
        mn = sname(m_idx, me[i][0])
        gn = sname(g_idx, ge[i][0])
        if mn != gn:
            mismatch.append((i, me[i][0], mn, ge[i][0], gn))
    if mismatch:
        print(f"  ✗ {len(mismatch)} 条样式名不一致（Word 比较会标格式变更）:")
        for i, msid, mn, gsid, gn in mismatch[:6]:
            print(f"      [{i}] mine pStyle={msid}({mn})  golden pStyle={gsid}({gn})")
        ok = False
    else:
        print("  ✓ 每条样式名与 golden 一致")

    # 3) 点导引 tab
    ml, gl = _has_leader_tab(md), _has_leader_tab(gd)
    print(f"  点导引 tab(……页码): mine {'有' if ml else '无'} / golden {'有' if gl else '无'}")
    if gl and not ml:
        print("  ✗ mine 目录缺点导引 tab")
        ok = False

    # 4) 锚点解析
    ma, mb = _toc_anchors(md), _body_bookmarks(md)
    resolved = sum(1 for a in ma if a in mb)
    print(f"  锚点解析: {resolved}/{len(ma)}")
    if resolved < len(ma):
        print(f"  ✗ {len(ma) - resolved} 个 PAGEREF 锚点悬空（F9 更新会丢页码）")
        ok = False

    # 5) 样式表语义冲突（mine 文中引用、与 golden 同 ID 不同义）
    refd = _referenced_styleids(md)
    coll = []
    for sid in sorted(refd):
        if sid in g_idx and sid in m_idx and g_idx[sid][:2] != m_idx[sid][:2]:
            coll.append((sid, g_idx[sid][0], g_idx[sid][1], m_idx[sid][0], m_idx[sid][1]))
    if coll:
        print(f"  ✗ 样式表 {len(coll)} 处同 ID 不同义（目录/标题会退化渲染）:")
        for sid, gn, gt, mn, mt in coll:
            print(f"      id={sid:5} golden={gn}({gt})  mine={mn}({mt})")
        ok = False
    else:
        print("  ✓ 被引用 styleId 与 golden 语义一致")

    print("✓ 目录与 golden 一致" if ok else "✗ 目录与 golden 不一致（--apply 修复）")
    return 0 if ok else 2


# ── 样式对账 ────────────────────────────────────────────────────────────────
def _reconcile_styles(ms_root, gs_root, droot):
    """把 mine 文中引用、与 golden 同 ID 不同义的 styleId 用 golden 定义覆盖，
    并补齐 golden 定义依赖的 basedOn/link/next（mine 缺则从 golden 导入；仍缺则剥引用）。
    返回 (overwritten:list, imported:list)。"""
    g_idx = _style_index(gs_root)
    m_idx = _style_index(ms_root)
    refd = _referenced_styleids(droot)
    styles_parent = ms_root  # <w:styles> 根，<w:style> 是其直接 child

    overwritten, imported = [], []

    def m_has(sid):
        return sid in _style_index(ms_root)

    def import_style(sid):
        """从 golden 深拷一份 style 定义进 mine（若 mine 已有则跳过）。"""
        if m_has(sid) or sid not in g_idx:
            return
        styles_parent.append(deepcopy(g_idx[sid][2]))
        imported.append(sid)

    # 先确定要覆盖的 ID
    targets = [sid for sid in refd
               if sid in g_idx and sid in m_idx and g_idx[sid][:2] != m_idx[sid][:2]]

    for sid in targets:
        # 删 mine 旧定义，插 golden 定义深拷
        old = m_idx[sid][2]
        gnew = deepcopy(g_idx[sid][2])
        old.getparent().replace(old, gnew)
        overwritten.append(sid)

    # 补依赖：覆盖后，逐个新定义的 basedOn/link/next 指向必须在 mine 存在
    for sid in targets:
        gnew = _style_index(ms_root)[sid][2]
        for tag in ("basedOn", "link", "next"):
            e = gnew.find(q(tag))
            if e is None:
                continue
            dep = e.get(q("val"))
            if not m_has(dep):
                import_style(dep)          # 从 golden 拉
                if not m_has(dep):          # golden 也没有 → 剥掉悬空引用
                    gnew.remove(e)
    return overwritten, imported


# ── 目录块移植 ──────────────────────────────────────────────────────────────
def _port_toc_block(md_root, gd_root):
    """用 golden 的整个目录 <w:sdt> 替换 mine 的。返回 (mine段数, golden段数)。"""
    msdt = _toc_sdt(md_root)
    gsdt = _toc_sdt(gd_root)
    if msdt is None or gsdt is None:
        raise RuntimeError("未定位到目录 sdt")
    m_paras = len(msdt.findall(".//" + q("p")))
    g_paras = len(gsdt.findall(".//" + q("p")))
    msdt.getparent().replace(msdt, deepcopy(gsdt))
    return m_paras, g_paras


# ── --apply ──────────────────────────────────────────────────────────────────
def cmd_apply(mine, golden, no_backup):
    md, ms = _load(mine)
    gd, gs = _load(golden)

    overw, imp = _reconcile_styles(ms, gs, md)
    removed, inserted = _port_toc_block(md, gd)

    if not no_backup:
        bak = mine.with_suffix(mine.suffix + f".bak-{datetime.now():%Y%m%d-%H%M%S}")
        shutil.copy2(mine, bak)
        print(f"  备份 → {bak.name}")

    new_doc = etree.tostring(md, xml_declaration=True, encoding="UTF-8", standalone=True)
    new_sty = etree.tostring(ms, xml_declaration=True, encoding="UTF-8", standalone=True)
    tmp = mine.with_suffix(mine.suffix + ".tmp")
    with zipfile.ZipFile(mine) as zin, \
         zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as zout:
        for item in zin.infolist():
            if item.filename == "word/document.xml":
                data = new_doc
            elif item.filename == "word/styles.xml":
                data = new_sty
            else:
                data = zin.read(item.filename)
            zout.writestr(item, data)
    tmp.replace(mine)

    print(f"[目录同步] {mine.name}")
    print(f"  样式对账: 覆盖 {len(overw)} 处 {overw}" + (f" + 导入依赖 {imp}" if imp else ""))
    print(f"  目录块移植: 删 mine {removed} 段 → 插 golden {inserted} 段")
    # 复验
    md2, ms2 = _load(mine)
    ma, mb = _toc_anchors(md2), _body_bookmarks(md2)
    print(f"  复验锚点解析: {sum(1 for a in ma if a in mb)}/{len(ma)}  · 点导引 tab: {'有' if _has_leader_tab(md2) else '无'}")
    return 0


# ── --shot ───────────────────────────────────────────────────────────────────
def _toc_pages(pdf):
    try:
        from pypdf import PdfReader
    except Exception:
        from PyPDF2 import PdfReader
    rd = PdfReader(str(pdf))
    pages = []
    for i, pg in enumerate(rd.pages, 1):
        t = pg.extract_text() or ""
        if "目" in t and "录" in t and ("前  言" in t or "编制目的" in t):
            pages.append(i)
        if re.search(r"编制目的与依据\s*\d", t) and i not in pages:
            pages.append(i)
    return pages[:4] or [3, 4]


def cmd_shot(docx, out_dir):
    out_dir = out_dir or docx.parent / (docx.stem + "_目录检")
    out_dir.mkdir(parents=True, exist_ok=True)
    subprocess.run([SOFFICE, "--headless", "--convert-to", "pdf",
                    "--outdir", str(out_dir), str(docx)], capture_output=True, timeout=240)
    pdf = out_dir / (docx.stem + ".pdf")
    if not pdf.exists():
        print("✗ LibreOffice 转 PDF 失败", file=sys.stderr)
        return 1
    pages = _toc_pages(pdf)
    print(f"[目录可视化] {docx.name}: 目录页 {pages} → {out_dir}")
    for pg in pages:
        subprocess.run(["/opt/homebrew/bin/pdftoppm", "-png", "-f", str(pg), "-l", str(pg),
                        "-r", "100", str(pdf), str(out_dir / f"目录检-p{pg:03d}")],
                       capture_output=True)
    print(f"  PNG 已出于 {out_dir}（打开眼检点导引/页码/缩进）")
    return 0


def main(argv=None):
    ap = argparse.ArgumentParser(description="目录与 golden 同步（单功能·可视化）")
    ap.add_argument("docx", type=Path)
    ap.add_argument("--ref", type=Path, help="同源 golden 参照件")
    ap.add_argument("--check", action="store_true")
    ap.add_argument("--apply", action="store_true")
    ap.add_argument("--shot", action="store_true")
    ap.add_argument("--out-dir", type=Path, default=None)
    ap.add_argument("--no-backup", action="store_true")
    a = ap.parse_args(argv)
    if not a.docx.exists():
        print(f"找不到: {a.docx}", file=sys.stderr)
        return 1
    if a.shot:
        return cmd_shot(a.docx, a.out_dir)
    if not a.ref or not a.ref.exists():
        print("--check/--apply 需 --ref golden 参照件", file=sys.stderr)
        return 1
    if a.apply:
        return cmd_apply(a.docx, a.ref, a.no_backup)
    return cmd_check(a.docx, a.ref)


if __name__ == "__main__":
    sys.exit(main())
