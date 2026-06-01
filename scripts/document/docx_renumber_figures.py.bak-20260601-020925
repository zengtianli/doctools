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


def main():
    ap = argparse.ArgumentParser(description="按出现顺序重排 docx 图号 + 同步正文引用")
    ap.add_argument("docx")
    ap.add_argument("-o", "--output")
    ap.add_argument("--dry-run", action="store_true")
    ap.add_argument("--inplace", action="store_true")
    ap.add_argument("--prefix", default="Figure")
    a = ap.parse_args()

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
    _write(a.docx, root, out)
    nums, seq = _verify(out, a.prefix)
    print(f"已写: {out}")
    print(f"验证: captions = {nums}")
    print("✓ 连续 1..N" if seq else "✗ 重编号后仍不连续，请检查")
    sys.exit(0 if seq else 2)


if __name__ == "__main__":
    main()
