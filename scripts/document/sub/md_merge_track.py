#!/usr/bin/env python3
"""md_merge_track.py — 把 MD 段落以 Word 修订标记(track changes)插入到 DOCX 锚点之前。

GOAL report-automation Phase 0-B (2026-05-29): 从 reclaim 项目级孤本
`reclaim/.claude/skills/merge-tracked/scripts/merge.py` 上提总部 (铁律 #5 子公司不另造轮子)。
RevisionWriter / parse_md / paragraph_to_runs / find_paragraph 为验证过的 delicate 资产,
verbatim 抬入零回归;新增 register()(接 docx_cli `md-merge-track`)+ --in-place(.bak, Work §1.5)。

与 `md-merge`(md_merge_impl.py · 替换整节 · plain)的区别:
  - md-merge       = 替换 [start,end) 整节, 无 track changes, 支持 md-table
  - md-merge-track = 在锚点段**之前**插入新段(全 w:ins) + 可选 renumber 锚点内子串(w:del+w:ins)

能力:
  - 锚点前插入新段(全 w:ins 修订标记)
  - renumber 锚点段内子串(w:del + w:ins)
  - 剥除 MD 里 [n] 引用标记
  - 识别 **bold** run
MD 首行(**标题**)用锚点 pStyle 成标题; 后续段为正文(默认无 pStyle)。

CLI:
  docx_cli.py md-merge-track --src S.docx --out O.docx --md C.md --anchor "1.2.2.1 评估指标" \
      [--renumber "1.2.2.1=>1.2.2.2"] [--style-anchor ...] [--body-same-style] \
      [--author ...] [--date ...]
  原地: --in-place(忽略 --out, 改原文件 + .bak-时间戳)
"""
from __future__ import annotations

import argparse
import os
import re
import shutil
import sys
import tempfile
import zipfile
from datetime import datetime

from lxml import etree

W_NS = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
W = f'{{{W_NS}}}'
XMLNS_SPACE = '{http://www.w3.org/XML/1998/namespace}space'


def parse_md(md_path):
    with open(md_path, 'r', encoding='utf-8') as f:
        text = f.read()
    text = text.split('\n---\n')[0].strip()
    lines = text.split('\n')
    title = re.sub(r'\*\*', '', lines[0]).strip()
    body = '\n'.join(lines[1:]).strip()
    paras = [p.strip() for p in re.split(r'\n\s*\n', body) if p.strip()]
    paras = [re.sub(r'\[\d+\]', '', p) for p in paras]
    return title, paras


def paragraph_to_runs(p):
    runs, pos = [], 0
    for m in re.finditer(r'\*\*(.+?)\*\*', p):
        if m.start() > pos:
            runs.append((p[pos:m.start()], False))
        runs.append((m.group(1), True))
        pos = m.end()
    if pos < len(p):
        runs.append((p[pos:], False))
    return runs


class RevisionWriter:
    def __init__(self, author, date):
        self.author = author
        self.date = date
        self.counter = 1000

    def nid(self):
        self.counter += 1
        return str(self.counter)

    def make_inserted_paragraph(self, runs, style=None):
        p = etree.Element(f'{W}p')
        pPr = etree.SubElement(p, f'{W}pPr')
        if style:
            pStyle = etree.SubElement(pPr, f'{W}pStyle')
            pStyle.set(f'{W}val', style)
        rPr = etree.SubElement(pPr, f'{W}rPr')
        ins_mark = etree.SubElement(rPr, f'{W}ins')
        ins_mark.set(f'{W}id', self.nid())
        ins_mark.set(f'{W}author', self.author)
        ins_mark.set(f'{W}date', self.date)

        ins_wrap = etree.SubElement(p, f'{W}ins')
        ins_wrap.set(f'{W}id', self.nid())
        ins_wrap.set(f'{W}author', self.author)
        ins_wrap.set(f'{W}date', self.date)

        for text, bold in runs:
            r = etree.SubElement(ins_wrap, f'{W}r')
            if bold:
                rrPr = etree.SubElement(r, f'{W}rPr')
                etree.SubElement(rrPr, f'{W}b')
            t = etree.SubElement(r, f'{W}t')
            t.set(XMLNS_SPACE, 'preserve')
            t.text = text
        return p

    def replace_text_in_paragraph(self, target_p, old, new):
        """Track-change rename: find old in a w:t and split into before / del / ins / after."""
        for t_el in target_p.findall(f'.//{W}t'):
            if t_el.text and old in t_el.text:
                orig = t_el.text
                idx = orig.find(old)
                before, after = orig[:idx], orig[idx + len(old):]
                r = t_el.getparent()
                rPr = r.find(f'{W}rPr')
                rpar = r.getparent()
                ridx = list(rpar).index(r)

                def clone_rpr():
                    return etree.fromstring(etree.tostring(rPr)) if rPr is not None else None

                def mk_run(text):
                    nr = etree.Element(f'{W}r')
                    c = clone_rpr()
                    if c is not None:
                        nr.append(c)
                    nt = etree.SubElement(nr, f'{W}t')
                    nt.set(XMLNS_SPACE, 'preserve')
                    nt.text = text
                    return nr

                new_els = []
                if before:
                    new_els.append(mk_run(before))

                dw = etree.Element(f'{W}del')
                dw.set(f'{W}id', self.nid())
                dw.set(f'{W}author', self.author)
                dw.set(f'{W}date', self.date)
                dr = etree.SubElement(dw, f'{W}r')
                c = clone_rpr()
                if c is not None:
                    dr.append(c)
                dt = etree.SubElement(dr, f'{W}delText')
                dt.set(XMLNS_SPACE, 'preserve')
                dt.text = old
                new_els.append(dw)

                iw = etree.Element(f'{W}ins')
                iw.set(f'{W}id', self.nid())
                iw.set(f'{W}author', self.author)
                iw.set(f'{W}date', self.date)
                ir = etree.SubElement(iw, f'{W}r')
                c = clone_rpr()
                if c is not None:
                    ir.append(c)
                it = etree.SubElement(ir, f'{W}t')
                it.set(XMLNS_SPACE, 'preserve')
                it.text = new
                new_els.append(iw)

                if after:
                    new_els.append(mk_run(after))

                rpar.remove(r)
                for j, el in enumerate(new_els):
                    rpar.insert(ridx + j, el)
                return True
        return False


def find_paragraph(body, anchor_prefix):
    for p in body.iter(f'{W}p'):
        text = ''.join(t.text or '' for t in p.findall(f'.//{W}t'))
        if text.strip().startswith(anchor_prefix):
            return p
    return None


def apply_track(src, out, md, anchor, *, style_anchor=None, body_same_style=False,
                renumber=None, author='Tianli Zeng', date='2026-04-23T00:00:00Z',
                in_place=False, no_backup=False) -> str:
    """核心: MD → 锚点前插入(track changes) + 可选 renumber。返回输出路径。"""
    if in_place:
        if not no_backup:
            bak = f"{src}.bak-{datetime.now().strftime('%Y%m%d-%H%M%S')}"
            shutil.copy2(src, bak)
            print(f'已备份: {bak}')
        out = src

    title, paragraphs = parse_md(md)
    paragraph_runs = [paragraph_to_runs(p) for p in paragraphs]
    print(f'Title : {title}')
    print(f'Paras : {len(paragraphs)}')

    if out != src:
        shutil.copy2(src, out)
    tmp = tempfile.mkdtemp(prefix='md_merge_track_')
    try:
        with zipfile.ZipFile(out, 'r') as z:
            z.extractall(tmp)

        doc_xml = os.path.join(tmp, 'word', 'document.xml')
        tree = etree.parse(doc_xml, etree.XMLParser(remove_blank_text=False))
        root = tree.getroot()
        body = root.find(f'{W}body')

        target = find_paragraph(body, anchor)
        if target is None:
            sys.exit(f'ERROR: anchor not found: {anchor}')
        anchor_text = ''.join(t.text or '' for t in target.findall(f'.//{W}t'))
        print(f'Anchor: "{anchor_text[:60]}"')

        style_para = target
        if style_anchor:
            sp = find_paragraph(body, style_anchor)
            if sp is None:
                sys.exit(f'ERROR: style-anchor not found: {style_anchor}')
            style_para = sp
        tpPr = style_para.find(f'{W}pPr')
        style = None
        if tpPr is not None:
            ps = tpPr.find(f'{W}pStyle')
            if ps is not None:
                style = ps.get(f'{W}val')
        print(f'Style : {style}')

        rw = RevisionWriter(author, date)
        body_style = style if body_same_style else None
        new_ps = [rw.make_inserted_paragraph([(title, False)], style=style)]
        for runs in paragraph_runs:
            new_ps.append(rw.make_inserted_paragraph(runs, style=body_style))

        for np in new_ps:
            target.addprevious(np)
        print(f'Inserted {len(new_ps)} paragraphs before anchor')

        if renumber:
            old, new = renumber.split('=>')
            if rw.replace_text_in_paragraph(target, old, new):
                print(f'Renumber: {old} -> {new} done')
            else:
                print(f'WARN: renumber source "{old}" not found in anchor')

        tree.write(doc_xml, xml_declaration=True, encoding='UTF-8', standalone=True)

        os.remove(out)
        with zipfile.ZipFile(out, 'w', zipfile.ZIP_DEFLATED) as z:
            for rd, _ds, fs in os.walk(tmp):
                for f in fs:
                    full = os.path.join(rd, f)
                    arc = os.path.relpath(full, tmp)
                    z.write(full, arc)
    finally:
        shutil.rmtree(tmp, ignore_errors=True)

    print(f'OK: {out}')
    return out


# ─── docx_cli 接入 ───────────────────────────────────────────────────────

def _run(args) -> int:
    apply_track(
        args.src, args.out, args.md, args.anchor,
        style_anchor=args.style_anchor, body_same_style=args.body_same_style,
        renumber=args.renumber, author=args.author, date=args.date,
        in_place=args.in_place, no_backup=args.no_backup,
    )
    return 0


def register(subparsers) -> None:
    """Register `md-merge-track` as a top-level subcommand."""
    p = subparsers.add_parser(
        "md-merge-track",
        help="MD 段以 track-changes 插入到 DOCX 锚点前(+可选 renumber); --in-place 原地+.bak",
    )
    p.add_argument("--src", required=True, help="源 DOCX")
    p.add_argument("--out", default=None, help="输出 DOCX(--in-place 时忽略)")
    p.add_argument("--md", required=True, help="要插入的 MD")
    p.add_argument("--anchor", required=True, help="锚点段文本(前缀匹配),新段插其前")
    p.add_argument("--style-anchor", default=None, help="取 pStyle 的备用锚点(默认 = --anchor)")
    p.add_argument("--body-same-style", action="store_true", help="所有插入段套锚点 pStyle(非仅标题)")
    p.add_argument("--renumber", default=None, help="锚点段内子串改号 OLD=>NEW(w:del+w:ins)")
    p.add_argument("--author", default="Tianli Zeng")
    p.add_argument("--date", default="2026-04-23T00:00:00Z")
    p.add_argument("--in-place", action="store_true", help="原地改源文件 + 自动 .bak-时间戳(Work §1.5)")
    p.add_argument("--no-backup", action="store_true", help="配合 --in-place 跳过备份")
    p.set_defaults(func=_run)


def main() -> int:
    ap = argparse.ArgumentParser(prog="md-merge-track", description=__doc__,
                                 formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument('--src', required=True)
    ap.add_argument('--out', default=None)
    ap.add_argument('--md', required=True)
    ap.add_argument('--anchor', required=True)
    ap.add_argument('--style-anchor', default=None)
    ap.add_argument('--body-same-style', action='store_true')
    ap.add_argument('--renumber', default=None)
    ap.add_argument('--author', default='Tianli Zeng')
    ap.add_argument('--date', default='2026-04-23T00:00:00Z')
    ap.add_argument('--in-place', action='store_true')
    ap.add_argument('--no-backup', action='store_true')
    ap.add_argument('--tmp', default=None, help=argparse.SUPPRESS)  # 向后兼容(忽略;内部用 mkdtemp)
    args = ap.parse_args()
    if not args.in_place and not args.out:
        ap.error('需 --out 或 --in-place')
    return _run(args)


if __name__ == '__main__':
    sys.exit(main())
