#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""typeset_ops.py — 排版链五件套合一（2026-07-31 家族折叠；OQ1 用户批准五合一：
五件在 Work/CLAUDE.md 是同一条外部契约 --check/--apply/--shot，typeset_pipeline
五个常量收拢一处）

子命令 ↔ 原脚本（函数体逐字搬移；模块级 main/apply_path/cmd_* 改名 _<sub> 后缀；
q/_norm/_para_text/_pstyle/_body_root 各段逐字相同的重复定义合并/保留后影同义，
port_sections 的 2 参 _root 改名 _zroot 避让 center_images 的 1 参 _root）：

    restyle        ← restyle.py（从同源 golden 整段克隆段落/字符格式）
    sync-toc       ← sync_toc.py（目录块+目录样式对账移植）
    port-sections  ← port_sections.py（节结构移植：分节/横竖/页眉脚水印）
    center-images  ← center_images.py（图片段显式居中+零缩进）
    line-spacing   ← line_spacing.py（对照 golden 补正文固定行距）

各子命令 CLI 与原独立脚本逐字一致：python3 sub/typeset_ops.py <sub> <docx> …。
退役原件在 ~/.Trash/consolidation-20260731/typeset_ops/（含 MANIFEST.md）。
"""
from __future__ import annotations

import argparse  # noqa: E402
import re  # noqa: E402
import shutil  # noqa: E402
import subprocess  # noqa: E402
import sys  # noqa: E402
import zipfile  # noqa: E402
from copy import deepcopy  # noqa: E402
from datetime import datetime  # noqa: E402
from difflib import SequenceMatcher  # noqa: E402
from pathlib import Path  # noqa: E402

from lxml import etree  # noqa: E402

# 仓根 lib 进 sys.path —— 部件完整性断言 + soffice SSOT（append 不是 insert(0)，
# 防顶掉 sub/ 同名模块）
import sys as _sys
from pathlib import Path as _Path
_sys.path.append(str(_Path(__file__).resolve().parents[3] / "lib"))
from soffice import find_soffice, require_soffice  # noqa: E402  doctools SSOT: soffice 路径解析
from docx_parts import DEFAULT_ALLOW_CHANGED, assert_parts_intact  # noqa: E402  部件完整性断言

W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"


def q(tag: str) -> str:
    return W + tag


SOFFICE = find_soffice() or "/Applications/LibreOffice.app/Contents/MacOS/soffice"


# ══════════ restyle ← restyle.py ══════════

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


def cmd_check_restyle(target: Path, ref: Path) -> int:
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


def _restyle(target: Path, ref: Path | None, *, no_backup: bool, dry: bool = False) -> dict:
    """格式移植的唯一实现，CLI 与 spec 引擎共用。"""
    if ref is None:
        return {"changed": 0, "skipped": "缺 --ref golden，没有格式源可移植"}
    root, clone, restyle, kept, diff = _align(target, ref)
    if not clone and not restyle:
        return {"changed": 0, "kept": kept, "content_diff": diff}

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

    if dry:
        return {"changed": 0, "would_change": len(clone) + len(restyle),
                "kept": kept, "content_diff": diff}
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
    # 部件完整性：replace 之前断言（此刻 target 未动 = 天然基线，断言炸则源件无损）。
    # CLI cmd_apply 与 pipeline apply_path 都汇到本函数，一处覆盖双入口；
    # 唯一改动部件 word/document.xml 在 DEFAULT_ALLOW_CHANGED，零白名单。
    try:
        assert_parts_intact(target, tmp, verbose=False)
    except Exception:
        tmp.unlink(missing_ok=True)
        raise
    tmp.replace(target)
    return {"changed": len(clone) + len(restyle), "cloned": len(clone),
            "pstyle_only": len(restyle), "kept": kept, "content_diff": diff}


def cmd_apply_restyle(target: Path, ref: Path, no_backup: bool) -> int:
    r = _restyle(target, ref, no_backup=no_backup)
    if r["changed"] == 0:
        print(f"[restyle] {target.name}: 无需修改（kept={r.get('kept')} 差异={r.get('content_diff')}）")
        return 0
    print(f"[restyle] {target.name}: 整段克隆格式 {r['cloned']} 段 + 媒体段兜底 pStyle "
          f"{r['pstyle_only']} 段（跳过 kept={r['kept']} / 内容真差异 {r['content_diff']} 段不动）")
    return 0


def apply_path_restyle(docx_path, args=None) -> dict:
    """pipeline: 从同源 golden 移植段落/字符格式。spec 里给 `restyle: {ref: <golden.docx>}`。"""
    ref = getattr(args, "restyle_ref", None) if args else None
    return _restyle(Path(docx_path), Path(ref) if ref else None, no_backup=True,
                    dry=bool(getattr(args, "dry_run", False)) if args else False)


def main_restyle(argv=None):
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
        return cmd_apply_restyle(a.docx, a.ref, a.no_backup)
    return cmd_check_restyle(a.docx, a.ref)


# ══════════ sync-toc ← sync_toc.py ══════════

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
def cmd_check_sync_toc(mine, golden):
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
def _sync(mine, golden, *, no_backup, dry=False) -> dict:
    """目录同步的唯一实现，CLI 与 spec 引擎共用。"""
    if golden is None:
        return {"changed": 0, "skipped": "缺 --ref golden，没有目录块可移植"}
    md, ms = _load(mine)
    gd, gs = _load(golden)

    overw, imp = _reconcile_styles(ms, gs, md)
    removed, inserted = _port_toc_block(md, gd)

    if dry:
        return {"changed": 0, "would_change": len(overw) + removed + inserted,
                "styles_overwritten": len(overw), "toc_removed": removed,
                "toc_inserted": inserted}
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
    # 部件完整性断言(replace 前:mine 仍是未动源件=天然基线;炸则 tmp 被清、源件无损)。
    # styles.xml 不在 DEFAULT 白名单,显式报备。CLI cmd_apply 与 pipeline apply_path
    # (恒 no_backup=True)都汇到本函数,一处覆盖双入口,不依赖 .bak 存在。
    try:
        assert_parts_intact(mine, tmp,
                            allow_changed=set(DEFAULT_ALLOW_CHANGED) | {"word/styles.xml"},
                            verbose=False)
    except Exception:
        tmp.unlink(missing_ok=True)
        raise
    tmp.replace(mine)

    # 复验：移植完的目录锚点是不是真能解析到正文书签。解析不到 = 目录点不动，
    # 而 Word 里看上去一切正常 —— 所以这个数字必须进返回值，不能只 print。
    md2, ms2 = _load(mine)
    ma, mb = _toc_anchors(md2), _body_bookmarks(md2)
    return {"changed": len(overw) + removed + inserted,
            "styles_overwritten": len(overw), "styles_imported": imp,
            "toc_removed": removed, "toc_inserted": inserted,
            "anchors_resolved": sum(1 for a in ma if a in mb), "anchors_total": len(ma),
            "leader_tab": _has_leader_tab(md2)}


def cmd_apply_sync_toc(mine, golden, no_backup):
    r = _sync(mine, golden, no_backup=no_backup)
    if "skipped" in r:
        print(f"[目录同步] {mine.name}: {r['skipped']}", file=sys.stderr)
        return 1
    print(f"[目录同步] {mine.name}")
    print(f"  样式对账: 覆盖 {r['styles_overwritten']} 处"
          + (f" + 导入依赖 {r['styles_imported']}" if r["styles_imported"] else ""))
    print(f"  目录块移植: 删 mine {r['toc_removed']} 段 → 插 golden {r['toc_inserted']} 段")
    print(f"  复验锚点解析: {r['anchors_resolved']}/{r['anchors_total']}"
          f"  · 点导引 tab: {'有' if r['leader_tab'] else '无'}")
    return 0


def apply_path_sync_toc(docx_path, args=None) -> dict:
    """pipeline: 从同源 golden 移植目录块 + 对账目录样式。spec 里给 `sync_toc: {ref: <golden.docx>}`。"""
    ref = getattr(args, "sync_toc_ref", None) if args else None
    return _sync(Path(docx_path), Path(ref) if ref else None, no_backup=True,
                 dry=bool(getattr(args, "dry_run", False)) if args else False)


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


def cmd_shot_sync_toc(docx, out_dir):
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


def main_sync_toc(argv=None):
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
        return cmd_shot_sync_toc(a.docx, a.out_dir)
    if not a.ref or not a.ref.exists():
        print("--check/--apply 需 --ref golden 参照件", file=sys.stderr)
        return 1
    if a.apply:
        return cmd_apply_sync_toc(a.docx, a.ref, a.no_backup)
    return cmd_check_sync_toc(a.docx, a.ref)


# ══════════ port-sections ← port_sections.py ══════════

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


def _zroot(z, name):
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


def main_port_sections(argv=None):
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
    troot = _zroot(zt, "word/document.xml")
    groot = _zroot(zg, "word/document.xml")

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
    ct = _zroot(zt, "[Content_Types].xml")
    existing_ct = {o.get("PartName") for o in ct.findall(f"{{{CTNS}}}Override")}
    for pn, cty in ct_adds:
        if pn not in existing_ct:
            o = etree.SubElement(ct, f"{{{CTNS}}}Override")
            o.set("PartName", pn); o.set("ContentType", cty)
    drels = _zroot(zt, "word/_rels/document.xml.rels")
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
    # 部件完整性断言放 replace 之前：基线 = a.docx（产物部件全部继承自它，此刻未动），
    # 断言炸则 tmp 不落位。allow_added = 从 golden 有意搬来的 header/footer 及其 .rels，
    # 白名单与收集代码同源（new_parts/new_partrels 就是 writestr 进去的那份）；
    # repl 改写的 3 部件全在 DEFAULT_ALLOW_CHANGED。
    assert_parts_intact(a.docx, tmp,
                        allow_added=set(new_parts) | set(new_partrels),
                        verbose=False)
    tmp.replace(out)
    print(f"[port_sections] {out.name}: 移植 golden {len(breaks)} 节"
          f"（注入段节断 {placed} + 末节{'1' if final_sect is not None else '0'}）"
          f"，复制 header/footer {len(new_parts)} 部件")
    return 0


# ══════════ center-images ← center_images.py ══════════

def _has_image(p) -> bool:
    return p.find(".//" + q("drawing")) is not None or p.find(".//" + q("pict")) is not None


def _body_start_idx(paras) -> int:
    """正文起始段索引 = 目录(TOC)字段之后。封面/批准/落款的院 logo 是装帧不是正文图，
    不该被强制居中（否则打乱落款版式）——只居中正文图。无 TOC 则返 0（保持旧行为）。"""
    last = -1
    for idx, p in enumerate(paras):
        instr = "".join(n.text or "" for n in p.iter(q("instrText")))
        if "TOC" in instr or "PAGEREF _Toc" in instr:
            last = idx
    return last + 1


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
    paras = list(root.iter(q("p")))
    body0 = _body_start_idx(paras)
    imgs = [p for i, p in enumerate(paras) if i >= body0 and _has_image(p)]
    centered = [p for p in imgs if _is_centered(p)]
    return root, imgs, centered


def cmd_check_center_images(docx) -> int:
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


def _center(docx, *, no_backup: bool, dry: bool = False) -> dict:
    """居中动作的唯一实现。cmd_apply（CLI）与 apply_path（spec 引擎）都走这里 ——
    两条路各写一遍逻辑，就等于验证的时候测了个替身。返回 pipeline 约定的计数 dict。"""
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
        return {"changed": 0, "images": len(imgs), "note": "无图片段"}
    if dry:
        return {"changed": 0, "images": len(imgs), "would_change": fixed}
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
    # 部件完整性断言（fail-closed）：此刻 docx 仍是未动源件 = 天然基线；
    # 断言炸则不 replace，源件毫发无损。只改 document.xml，默认白名单即可。
    assert_parts_intact(docx, tmp, verbose=False)
    tmp.replace(docx)
    return {"changed": fixed, "images": len(imgs)}


def cmd_apply_center_images(docx, no_backup) -> int:
    r = _center(docx, no_backup=no_backup)
    if r["changed"] == 0:
        print(f"[图片居中] {docx.name}: 无图片段")
        return 0
    print(f"[图片居中] {docx.name}: {r['changed']} 个图片段显式居中+零缩进")
    return 0


def apply_path_center_images(docx_path, args=None) -> dict:
    """pipeline: 图片段显式居中 + 缩进清零（zip 级，走 doc-post 之后的 path 段）。
    备份由引擎统一做，这里恒 no_backup。"""
    return _center(Path(docx_path), no_backup=True,
                   dry=bool(getattr(args, "dry_run", False)) if args else False)


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


def cmd_shot_center_images(docx, out_dir) -> int:
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


def main_center_images(argv=None):
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
        return cmd_shot_center_images(a.docx, a.out_dir)
    if a.apply:
        return cmd_apply_center_images(a.docx, a.no_backup)
    return cmd_check_center_images(a.docx)


# ══════════ line-spacing ← line_spacing.py ══════════

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


def cmd_check_line_spacing(docx_path: Path, ref: Path | None) -> int:
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


def _fix(docx_path: Path, ref: Path | None, *, no_backup: bool, dry: bool = False) -> dict:
    """行距修复的唯一实现，CLI 与 spec 引擎共用。ref 缺失时**不静默 noop** ——
    这个动作没有参照件就无事可做，返回 skipped 让调用方能在报告里看见。"""
    if ref is None:
        return {"changed": 0, "skipped": "缺 --ref 参照件，行距值无从取"}
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
        return {"changed": 0, "note": "已全部设固定行距"}
    if dry:
        return {"changed": 0, "would_change": fixed}

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
    # 部件完整性断言放 replace 之前：docx_path 此刻仍是未动源件（天然基线，
    # pipeline 恒 no_backup 也不漏），断言炸则 tmp 不落位、源件毫发无损。
    # 这才是 docstring「其余 zip 项逐字节 verbatim, CRC 全等」的机器兜底。
    assert_parts_intact(docx_path, tmp, verbose=False)
    tmp.replace(docx_path)
    return {"changed": fixed}


def cmd_fix(docx_path: Path, ref: Path | None, no_backup: bool) -> int:
    if ref is None:
        print("行距修复需 --ref 参照件", file=sys.stderr)
        return 1
    r = _fix(docx_path, ref, no_backup=no_backup)
    if r["changed"] == 0:
        print(f"[行距修复] {docx_path.name}: 无需修改（已全部设固定行距）")
        return 0
    print(f"[行距修复] {docx_path.name}: 对照参照补固定行距于 {r['changed']} 段")
    return 0


def apply_path_line_spacing(docx_path, args=None) -> dict:
    """pipeline: 按参照件众数补固定行距。spec 里给 `line_spacing: {ref: <golden.docx>}`。"""
    ref = getattr(args, "line_spacing_ref", None) if args else None
    return _fix(Path(docx_path), Path(ref) if ref else None, no_backup=True,
                dry=bool(getattr(args, "dry_run", False)) if args else False)


def main_line_spacing(argv=None):
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
        return cmd_check_line_spacing(a.docx, a.ref)
    return cmd_fix(a.docx, a.ref, a.no_backup)


# ──────────────────────────── 家族入口（子命令分发）────────────────────────────

SUBCOMMANDS = {
    "restyle": main_restyle,
    "sync-toc": main_sync_toc,
    "port-sections": main_port_sections,
    "center-images": main_center_images,
    "line-spacing": main_line_spacing,
}


def main(argv: list[str] | None = None) -> int:
    args = list(sys.argv[1:] if argv is None else argv)
    if not args or args[0] in ("-h", "--help"):
        print("usage: typeset_ops.py {" + ",".join(SUBCOMMANDS) + "} <args…>\n"
              "每个子命令的参数与原独立脚本逐字一致：typeset_ops.py <sub> --help 查看。")
        return 0 if args else 2
    sub, rest = args[0], args[1:]
    fn = SUBCOMMANDS.get(sub)
    if fn is None:
        print(f"[typeset_ops] unknown subcommand: {sub!r}; choices={list(SUBCOMMANDS)}",
              file=sys.stderr)
        return 2
    saved = sys.argv[:]
    sys.argv = [sys.argv[0]] + rest
    try:
        rc = fn()
        return int(rc) if isinstance(rc, int) else 0
    finally:
        sys.argv = saved


if __name__ == "__main__":
    sys.exit(main())
