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


def renumber_cn_section(docx_path, kind="图", dry_run=False):
    """中文章节式 图{X.Y}-{N} / 表{X.Y}-{N} 按**节内物理顺序**重排 + 同步正文引用。

    与全局 renumber 的区别：编号是「节内」的（同 X.Y 前缀各自从 1 递增），断号
    (图2.1-2 后直接 2.1-4) 与重复号 (两个图2.2-1) 都按物理顺序顺排修正。重复号
    在本模式下**不报错退出**（重复正是要修的）——但若该重复号有正文引用，无法判定
    引用指向哪个实例 → 记 warning，引用不动，caption 仍按物理位置改对。

    kind: '图' 或 '表'。返回 (root, plan, caps, ok, warnings)。
    plan: [(para_idx, sec, old_n, new_n)] 物理顺序。dry_run 不改 root。
    """
    zin = zipfile.ZipFile(docx_path)
    root = etree.fromstring(zin.read("word/document.xml"))
    paras = list(root.iter(f"{W}p"))

    def ptext(p):
        return "".join(n.text or "" for n in _visible_t_nodes(p))

    # caption 行首：图2.1-4 / 表2.1-1（sec=2.1 可为 X 或 X.Y，n 为节内序），<80 字
    cap_re = re.compile(rf'^\s*{kind}\s*(\d+(?:\.\d+)?)\s*[-－—–]\s*(\d+)')
    caps = []  # (para_idx, sec, old_n) 物理顺序
    for idx, p in enumerate(paras):
        s = ptext(p).strip()
        if len(s) >= 80:
            continue
        m = cap_re.match(s)
        if m:
            caps.append((idx, m.group(1), int(m.group(2))))

    sec_ctr, plan = {}, []
    for idx, sec, old_n in caps:
        sec_ctr[sec] = sec_ctr.get(sec, 0) + 1
        plan.append((idx, sec, old_n, sec_ctr[sec]))

    # (sec,old_n)->new_n 映射；同 key 多个 new_n = 重复号冲突
    remap, conflict = {}, set()
    for idx, sec, old_n, new_n in plan:
        k = (sec, old_n)
        if k in remap and remap[k] != new_n:
            conflict.add(k)
        remap.setdefault(k, new_n)

    if dry_run:
        return root, plan, caps, True, sorted(conflict)

    para_by_idx = dict(enumerate(paras))
    n_after_re = re.compile(rf'(\s*{kind}\s*\d+(?:\.\d+)?\s*[-－—–]\s*)(\d+)')

    # 1) caption 改写：按物理位置直接改 n 段（不查 remap，避免重复号歧义）
    for idx, sec, old_n, new_n in plan:
        if old_n == new_n:
            continue
        nodes = _visible_t_nodes(para_by_idx[idx])
        full = "".join(n.text or "" for n in nodes)
        m2 = n_after_re.match(full)
        if not m2:
            continue
        _apply_edits(nodes, [(m2.start(2), m2.end(2), str(new_n))])

    # 2) 正文引用改写（排除 caption 段自身）；重复号有引用 → warning 不动
    cap_ids = {idx for idx, _, _ in caps}
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
    return root, plan, caps, True, warnings


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
    a = ap.parse_args()

    if a.cn_section:
        root, plan, caps, ok, warns = renumber_cn_section(a.docx, a.kind, dry_run=a.dry_run)
        changes = [(f"{a.kind}{s}-{o}", f"{a.kind}{s}-{n}") for _, s, o, n in plan if o != n]
        print(f"{a.kind}题数: {len(caps)}（节内分组）")
        print(f"变动 (现→新): {changes or '无（已连续）'}")
        if warns:
            print(f"⚠ 重复号引用需人工确认: {warns}")
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
