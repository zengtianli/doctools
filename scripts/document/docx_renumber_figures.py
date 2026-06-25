#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""docx_renumber_figures.py — 按文档出现顺序重排 Figure 编号 + 同步全部正文引用。

用途：论文/报告里图被增删/挪位后，图题号与正文 "Figure N" 引用全乱。本脚本扫
所有图题（caption），按它们在文档里**实际出现的物理顺序**重编号为 1..N，并把
captions + 所有正文引用（含范围 "Figures N–M"、列举 "Figures N, M and K"）一并改对。

为什么不是简单 str.replace：
  1. **跨 run 分裂**：Word 常把 "Figure " 和 "23" 拆进相邻 w:r/w:t，朴素正则
     按单个 w:t 改会漏掉数字在独立节点的 caption。本脚本按"段落级 concat 文本 +
     字符偏移定位 → 写回覆盖该偏移所在 w:t"，跨 run 也能改。
  2. **轮转/置换防碰撞**：重排常是置换（如 28→20, 20→21…27→28）。逐个 token
     读旧值、原子写新值，绝不串改（朴素全局 replace 会把 23→25 又被 25→27 二次改）。
  3. **排除 w:del**：track-changes 删除态(w:del/delText)里的旧文本不能算数、不改。
  4. ⚠ **python-docx 陷阱**：`Paragraph.text` 静默漏掉 `w:ins`（修订插入）里的
     run 文本 → 用它扫图题会漏掉"修订态插入的图/引用"。本脚本直接走 lxml 遍历
     w:t（含 w:ins、排除 w:del），不踩这个坑。

CLI:
  docx_renumber_figures.py <docx> [-o OUT] [--dry-run] [--prefix Figure] [--inplace]
  --dry-run : 只打印 现号→新号 映射 + 受影响引用，不写
  -o OUT    : 输出路径（默认 <name>.renumbered.docx）
  --inplace : 覆盖原文件（自动留 <name>.bak）
  --prefix  : 图前缀，默认 "Figure"（也匹配 "Fig."/"Fig"）

退出码：0 成功且重编号后 captions 连续 1..N；2 检测到重复图号（引用无法安全remap）。
"""
import argparse, re, shutil, sys, zipfile
from lxml import etree

W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"


def _under_del(t):
    return any(a.tag in (f"{W}del", f"{W}delText") for a in t.iterancestors())


def _visible_t_nodes(p):
    """段落内可见 w:t（含 w:ins，排除 w:del）按文档顺序。"""
    return [t for t in p.iter(f"{W}t") if not _under_del(t)]


def _apply_edits(nodes, edits):
    """把 (start,end,repl) 编辑（基于 nodes 拼接串的字符偏移，互不重叠）落到各 w:t。
    repl 仅在编辑**起点**所在节点写入；跨节点的尾部节点只做删除，避免 repl 重复。"""
    spans, pos = [], 0
    for n in nodes:
        tx = n.text or ""
        spans.append((pos, pos + len(tx)))
        pos += len(tx)
    edits = sorted(edits)
    for n, (ns, ne) in zip(nodes, spans):
        tx = n.text or ""
        local = []
        for s, e, repl in edits:
            if e <= ns or s >= ne:
                continue
            local.append((max(s, ns) - ns, min(e, ne) - ns, repl if s >= ns else ""))
        if local:
            local.sort()
            out, last = [], 0
            for cs, ce, repl in local:
                out.append(tx[last:cs]); out.append(repl); last = ce
            out.append(tx[last:])
            n.text = "".join(out)


def renumber(docx_path, prefix="Figure", dry_run=False):
    """返回 (root, remap, caption_order, ok)。dry_run 不修改 root（仍可读 remap）。"""
    zin = zipfile.ZipFile(docx_path)
    root = etree.fromstring(zin.read("word/document.xml"))
    paras = list(root.iter(f"{W}p"))

    def ptext(p):
        return "".join(n.text or "" for n in _visible_t_nodes(p))

    cap_re = re.compile(rf'^\s*(?:{prefix}|Fig\.?)\s*(\d+)\b', re.I)
    caption_order = []  # 现号，按物理顺序
    for p in paras:
        m = cap_re.match(ptext(p).strip())
        if m:
            caption_order.append(int(m.group(1)))

    # 重复图号 → 引用无法安全 remap
    if len(caption_order) != len(set(caption_order)):
        dup = [x for x in caption_order if caption_order.count(x) > 1]
        return root, {}, caption_order, False, sorted(set(dup))

    remap = {old: i + 1 for i, old in enumerate(caption_order)}
    if dry_run:
        return root, remap, caption_order, True, []

    # 引用匹配：Figure(s)/Fig. + 数字 + 可选 范围/列举（–,-,—,，and）
    cit = re.compile(rf'(?:{prefix}s?|Figs?\.?)\s*\d+(?:\s*(?:[–\-—,]|and)\s*\d+)*', re.I)
    num = re.compile(r'\d+')
    for p in paras:
        nodes = _visible_t_nodes(p)
        full = "".join(n.text or "" for n in nodes)
        if not re.search(rf'{prefix}|Fig', full, re.I):
            continue
        edits = []
        for m in cit.finditer(full):
            for nm in num.finditer(m.group(0)):
                v = int(nm.group(0))
                if v in remap and remap[v] != v:
                    edits.append((m.start() + nm.start(), m.start() + nm.end(), str(remap[v])))
        if edits:
            _apply_edits(nodes, edits)
    return root, remap, caption_order, True, []


def _write(src_docx, root, out_path):
    new_xml = etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True)
    zin = zipfile.ZipFile(src_docx)
    with zipfile.ZipFile(out_path, "w", zipfile.ZIP_DEFLATED) as zout:
        for it in zin.infolist():
            data = new_xml if it.filename == "word/document.xml" else zin.read(it.filename)
            zout.writestr(it, data)


def _verify(out_path, prefix):
    """重读输出确认 captions 连续 1..N。"""
    root = etree.fromstring(zipfile.ZipFile(out_path).read("word/document.xml"))
    cap_re = re.compile(rf'^\s*(?:{prefix}|Fig\.?)\s*(\d+)\b', re.I)
    nums = []
    for p in root.iter(f"{W}p"):
        txt = "".join(n.text or "" for n in _visible_t_nodes(p)).strip()
        m = cap_re.match(txt)
        if m:
            nums.append(int(m.group(1)))
    return nums, nums == list(range(1, len(nums) + 1))


def _para_style(p):
    pPr = p.find(f"{W}pPr")
    if pPr is None:
        return None
    ps = pPr.find(f"{W}pStyle")
    return ps.get(f"{W}val") if ps is not None else None


def _has_drawing(p):
    return p.find(f".//{W}drawing") is not None or p.find(f".//{W}pict") is not None


def _ptext(p):
    return "".join(n.text or "" for n in _visible_t_nodes(p))


def _center_style_ids(docx_path):
    """从 styles.xml 取「有效 jc=center」的段落样式 id 集合（含 basedOn 继承链）。"""
    try:
        sroot = etree.fromstring(zipfile.ZipFile(docx_path).read("word/styles.xml"))
    except Exception:
        return set()
    jc, based = {}, {}
    for st in sroot.iter(f"{W}style"):
        if st.get(f"{W}type") != "paragraph":
            continue
        sid = st.get(f"{W}styleId")
        if not sid:
            continue
        ppr = st.find(f"{W}pPr")
        j = ppr.find(f"{W}jc") if ppr is not None else None
        if j is not None:
            jc[sid] = j.get(f"{W}val")
        b = st.find(f"{W}basedOn")
        if b is not None:
            based[sid] = b.get(f"{W}val")

    def eff(sid, seen=None):
        seen = seen or set()
        if sid is None or sid in seen:
            return None
        seen.add(sid)
        return jc[sid] if sid in jc else eff(based.get(sid), seen)

    return {sid for sid in set(list(jc) + list(based)) if eff(sid) == "center"}


def _para_centered(p, center_ids):
    """段落有效对齐是否 center：显式 jc 优先，否则看 pStyle 是否继承 center。"""
    pPr = p.find(f"{W}pPr")
    jc = pPr.find(f"{W}jc") if pPr is not None else None
    if jc is not None:
        return jc.get(f"{W}val") == "center"
    return _para_style(p) in center_ids


def _collect_captions(paras, kind):
    """返回 (numbered, unnumbered, cap_style_ids)。
    numbered  : [(idx, sec, old_n)]  行首匹配 图X.Y-N 的题注
    unnumbered: [idx]                紧跟图片、属 caption 样式、无号的题注段（待补号）
    caption 样式 = numbered 题注所用 pStyle 的并集（据此识别同款无号题注，排除封面 logo 的非题注后段）。
    """
    cap_re = re.compile(rf'^\s*{kind}\s*(\d+(?:\.\d+)?)\s*[-－—–]\s*(\d+)')
    numbered, cap_styles = [], set()
    for idx, p in enumerate(paras):
        s = _ptext(p).strip()
        if 0 < len(s) < 80:
            m = cap_re.match(s)
            if m:
                numbered.append((idx, m.group(1), int(m.group(2))))
                st = _para_style(p)
                if st:
                    cap_styles.add(st)
    # 附图/附表 = 独立扁平编号体系（附图1、附图2…），不属 图X.Y-N 范畴 → 排除，否则误判为无号
    appendix_re = re.compile(r'^\s*附[图表]')
    unnumbered = []
    for idx, p in enumerate(paras):
        if idx == 0 or not _has_drawing(paras[idx - 1]):
            continue
        s = _ptext(p).strip()
        if not s or len(s) >= 80 or cap_re.match(s) or appendix_re.match(s):
            continue
        if cap_styles and _para_style(p) in cap_styles:
            unnumbered.append(idx)
    return numbered, unnumbered, cap_styles


def _prepend_caption_number(p, text):
    """在 caption 段首插入一个 run「图X.Y-N 」，rPr 克隆自原首 run（字体一致）。"""
    import copy
    first_r = p.find(f"{W}r")
    new_r = etree.Element(f"{W}r")
    if first_r is not None:
        rpr = first_r.find(f"{W}rPr")
        if rpr is not None:
            new_r.append(copy.deepcopy(rpr))
        first_r.addprevious(new_r)
    else:
        p.append(new_r)
    t = etree.SubElement(new_r, f"{W}t")
    t.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
    t.text = text


def check_cn_section(docx_path, kind="图", check_center=True):
    """只读机检：返回 issues dict（全空=干净）。供 gate / report figs 调。
    {unnumbered:[题注文本], gaps:{sec:[缺号]}, duplicates:{sec:[重号]}, uncentered:[题注/段号]}
    """
    zin = zipfile.ZipFile(docx_path)
    root = etree.fromstring(zin.read("word/document.xml"))
    paras = list(root.iter(f"{W}p"))
    numbered, unnumbered, cap_styles = _collect_captions(paras, kind)

    by_sec = {}
    for _, sec, n in numbered:
        by_sec.setdefault(sec, []).append(n)
    gaps, dups = {}, {}
    for sec, ns in by_sec.items():
        seen, d = set(), []
        for n in ns:
            if n in seen:
                d.append(n)
            seen.add(n)
        if d:
            dups[sec] = sorted(set(d))
        miss = sorted(set(range(1, max(ns) + 1)) - set(ns))
        if miss:
            gaps[sec] = miss

    uncentered = []
    if check_center and cap_styles:   # 仅对「有题注的报告」查居中，零题注文档跳过
        center_ids = _center_style_ids(docx_path)
        for idx, p in enumerate(paras):
            if not _has_drawing(p) or _para_centered(p, center_ids):
                continue
            cap = _ptext(paras[idx + 1]).strip()[:30] if idx + 1 < len(paras) else ""
            uncentered.append(cap or f"段{idx}")

    return {
        "unnumbered": [_ptext(paras[i]).strip()[:30] for i in unnumbered],
        "gaps": gaps,
        "duplicates": dups,
        "uncentered": uncentered,
    }


def center_figure_paragraphs(root, center_ids):
    """把所有含图段落的有效对齐改为 center（已居中的跳过）。返回改动数。"""
    n = 0
    for p in root.iter(f"{W}p"):
        if not _has_drawing(p) or _para_centered(p, center_ids):
            continue
        pPr = p.find(f"{W}pPr")
        if pPr is None:
            pPr = etree.Element(f"{W}pPr")
            p.insert(0, pPr)
        jc = pPr.find(f"{W}jc")
        if jc is None:
            jc = etree.SubElement(pPr, f"{W}jc")
        jc.set(f"{W}val", "center")
        n += 1
    return n


def renumber_cn_section(docx_path, kind="图", dry_run=False, supplement=True, fix_center=False):
    """中文章节式 图{X.Y}-{N} / 表{X.Y}-{N} 按**节内物理顺序**重排 + 补号 + 同步正文引用。

    编号是「节内」的（同 X.Y 前缀各自从 1 递增）：断号(图2.1-2 后直接 2.1-4)、重复号
    (两个图2.2-1) 按物理顺序顺排修正。**supplement=True 时**，紧跟图片、与已编号题注同
    样式、却无号的题注段（如「狮子口水库溢洪道出口段」）会被**补号**（节号据物理上最近的
    已编号题注推断，前优先后兜底；定位不到则记 warning 不补）。重复号有正文引用 → warning
    不动引用、caption 仍按物理位置改对。fix_center=True 同时把含图段落居中。

    kind: '图' 或 '表'。返回 (root, plan, caps, ok, warnings)。
    plan: [(para_idx, typ, sec, old_n, new_n)]；typ='num' 改号 / 'new' 补号。dry_run 不改 root。
    """
    zin = zipfile.ZipFile(docx_path)
    root = etree.fromstring(zin.read("word/document.xml"))
    paras = list(root.iter(f"{W}p"))
    numbered, unnumbered, _ = _collect_captions(paras, kind)

    num_idx_sec = [(idx, sec) for idx, sec, _ in numbered]

    def infer_sec(idx):
        prev = [s for i, s in num_idx_sec if i < idx]
        if prev:
            return prev[-1]
        nxt = [s for i, s in num_idx_sec if i > idx]
        return nxt[0] if nxt else None

    items = sorted([(idx, "num", sec, old_n) for idx, sec, old_n in numbered]
                   + [(idx, "new", None, None) for idx in (unnumbered if supplement else [])])

    sec_ctr, plan, skipped = {}, [], []
    for idx, typ, sec, old_n in items:
        if typ == "new":
            sec = infer_sec(idx)
            if sec is None:
                skipped.append(idx)
                continue
        sec_ctr[sec] = sec_ctr.get(sec, 0) + 1
        plan.append((idx, typ, sec, old_n, sec_ctr[sec]))

    # (sec,old_n)->new_n 映射（仅已编号题注，供正文引用 remap）；同 key 多 new_n = 重复号冲突
    remap, conflict = {}, set()
    for idx, typ, sec, old_n, new_n in plan:
        if typ != "num":
            continue
        k = (sec, old_n)
        if k in remap and remap[k] != new_n:
            conflict.add(k)
        remap.setdefault(k, new_n)

    caps_compat = list(numbered)  # 兼容旧返回签名（仅已编号题注）

    if dry_run:
        return root, plan, caps_compat, True, sorted(conflict)

    para_by_idx = dict(enumerate(paras))
    n_after_re = re.compile(rf'(\s*{kind}\s*\d+(?:\.\d+)?\s*[-－—–]\s*)(\d+)')

    # 1) caption 改写
    for idx, typ, sec, old_n, new_n in plan:
        p = para_by_idx[idx]
        if typ == "num":
            if old_n == new_n:
                continue
            nodes = _visible_t_nodes(p)
            full = "".join(n.text or "" for n in nodes)
            m2 = n_after_re.match(full)
            if m2:
                _apply_edits(nodes, [(m2.start(2), m2.end(2), str(new_n))])
        else:  # 补号
            _prepend_caption_number(p, f"{kind}{sec}-{new_n} ")

    # 2) 正文引用改写（排除**所有** caption 段：已编号 + 补号）；重复号有引用 → warning 不动
    #    ⚠ 必须含补号段——否则刚 prepend 的「图X.Y-N」会被本循环当正文引用再 remap 一次（曾踩）。
    cap_ids = {row[0] for row in plan}
    ref_re = re.compile(rf'{kind}\s*(\d+(?:\.\d+)?)\s*[-－—–]\s*(\d+)')
    warnings = []
    for idx, p in enumerate(paras):
        if idx in cap_ids:
            continue
        nodes = _visible_t_nodes(p)
        full = "".join(n.text or "" for n in nodes)
        if kind not in full:
            continue
        edits = []
        for m in ref_re.finditer(full):
            k = (m.group(1), int(m.group(2)))
            if k in conflict:
                warnings.append(f"{kind}{k[0]}-{k[1]}（重复号，引用需人工确认）")
                continue
            if k in remap and remap[k] != k[1]:
                edits.append((m.start(2), m.end(2), str(remap[k])))
        if edits:
            _apply_edits(nodes, edits)

    if fix_center:
        n_c = center_figure_paragraphs(root, _center_style_ids(docx_path))
        if n_c:
            warnings.append(f"已居中 {n_c} 个含图段落")
    if skipped:
        warnings.append(f"{len(skipped)} 个无号题注无法定位章节（无相邻已编号题注），未补号")
    return root, plan, caps_compat, True, warnings


def _verify_cn(out_path, kind):
    """重读输出确认每个 sec 内 n 连续 1..k。返回 ({sec:[n...]}, all_ok)。"""
    root = etree.fromstring(zipfile.ZipFile(out_path).read("word/document.xml"))
    cap_re = re.compile(rf'^\s*{kind}\s*(\d+(?:\.\d+)?)\s*[-－—–]\s*(\d+)')
    by_sec = {}
    for p in root.iter(f"{W}p"):
        s = "".join(n.text or "" for n in _visible_t_nodes(p)).strip()
        if len(s) >= 80:
            continue
        m = cap_re.match(s)
        if m:
            by_sec.setdefault(m.group(1), []).append(int(m.group(2)))
    ok = all(v == list(range(1, len(v) + 1)) for v in by_sec.values())
    return by_sec, ok


def main():
    ap = argparse.ArgumentParser(description="按出现顺序重排 docx 图号 + 同步正文引用")
    ap.add_argument("docx")
    ap.add_argument("-o", "--output")
    ap.add_argument("--dry-run", action="store_true")
    ap.add_argument("--inplace", action="store_true")
    ap.add_argument("--prefix", default="Figure")
    ap.add_argument("--cn-section", action="store_true",
                    help="中文章节式 图X.Y-N / 表X.Y-N 节内重排（修断号+重复号）")
    ap.add_argument("--kind", default="图", choices=["图", "表"],
                    help="--cn-section 模式下的题注类型，默认 图")
    ap.add_argument("--check", action="store_true",
                    help="只读机检（配 --cn-section）：报无序号题注/断号/重复号/未居中，有问题 exit 2")
    ap.add_argument("--no-supplement", action="store_true",
                    help="--cn-section 重排时不给无号题注补号（默认补号）")
    ap.add_argument("--fix-center", action="store_true",
                    help="--cn-section 重排时顺带把含图段落居中")
    a = ap.parse_args()

    # 只读机检模式（gate / report figs 调用入口）
    if a.check:
        kind = a.kind if a.cn_section else "图"
        iss = check_cn_section(a.docx, kind)
        bad = any(iss[k] for k in ("unnumbered", "gaps", "duplicates", "uncentered"))
        print(f"[{kind}题机检] {a.docx}")
        print(f"  无序号题注: {iss['unnumbered'] or '无'}")
        print(f"  节内断号  : {iss['gaps'] or '无'}")
        print(f"  重复号    : {iss['duplicates'] or '无'}")
        print(f"  未居中图片: {iss['uncentered'] or '无'}")
        print("✗ 发现问题，需重排/补号/居中" if bad else "✓ 图序号与居中均合规")
        sys.exit(2 if bad else 0)

    if a.cn_section:
        root, plan, caps, ok, warns = renumber_cn_section(
            a.docx, a.kind, dry_run=a.dry_run,
            supplement=not a.no_supplement, fix_center=a.fix_center)
        changes = [(f"{a.kind}{s}-{o or '—'}", f"{a.kind}{s}-{n}")
                   for _, typ, s, o, n in plan if typ == "new" or o != n]
        print(f"{a.kind}题数: {len(caps)} 已编号 + {sum(1 for _,t,*_ in plan if t=='new')} 补号（节内分组）")
        print(f"变动 (现→新): {changes or '无（已连续）'}")
        if warns:
            print(f"⚠ 提示: {warns}")
        if a.dry_run:
            print("[dry-run] 未写文件")
            return
        out = a.docx if a.inplace else (a.output or re.sub(r'\.docx$', '.renumbered.docx', a.docx))
        if a.inplace:
            shutil.copy2(a.docx, a.docx + ".bak")
            _write(a.docx + ".bak", root, out)   # 从 .bak 读、写回原文件，避免读写同路径截断
        else:
            _write(a.docx, root, out)
        by_sec, seq = _verify_cn(out, a.kind)
        print(f"已写: {out}")
        print(f"验证: 各节 {a.kind}号 = {by_sec}")
        print("✓ 每节连续 1..k" if seq else "✗ 重编号后仍不连续，请检查")
        sys.exit(0 if seq else 2)

    root, remap, order, ok, dup = renumber(a.docx, a.prefix, dry_run=a.dry_run)
    if not ok:
        print(f"✗ 检测到重复图号 {dup}，正文引用无法安全 remap。先消重再跑。", file=sys.stderr)
        sys.exit(2)

    changes = {o: n for o, n in remap.items() if o != n}
    print(f"图题数: {len(order)}  现号顺序: {order}")
    print(f"变动 (现→新): {changes or '无（已连续）'}")

    if a.dry_run:
        print("[dry-run] 未写文件")
        return

    out = a.docx if a.inplace else (a.output or re.sub(r'\.docx$', '.renumbered.docx', a.docx))
    if a.inplace:
        shutil.copy2(a.docx, a.docx + ".bak")
        _write(a.docx + ".bak", root, out)   # 从 .bak 读、写回原文件，避免读写同路径截断
    else:
        _write(a.docx, root, out)
    nums, seq = _verify(out, a.prefix)
    print(f"已写: {out}")
    print(f"验证: captions = {nums}")
    print("✓ 连续 1..N" if seq else "✗ 重编号后仍不连续，请检查")
    sys.exit(0 if seq else 2)


if __name__ == "__main__":
    main()
