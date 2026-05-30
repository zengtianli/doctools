#!/usr/bin/env python3
"""
pptx_align.py — PPTX 样式对齐工具集（audit / layout / tablestyle / titlecolor / render）

distill 自 qual-supply 2026-05-30 会话:把"成果2 PPT 格式对齐成果1"反复手搓的 5 类
OOXML 操作沉淀为总部单功能子命令。与 pptx_tools.py(font/format/table 开关)互补 —
本模块专做 pptx_tools 缺的:结构侦察、layout 引用切换、表格 styleId 整体替换、
标题 run 改色、soffice 渲染验证。

依赖: 仅 stdlib(zipfile/re/subprocess) — 刻意不依赖 dockit/python-pptx,系统 python3 即可跑,
       避免 pptx_tools 那种 venv 依赖陷阱(用错解释器 → 误判工具坏 → 手搓重造轮子)。

统一 CLI:
    python3 pptx_align.py audit       <pptx>                  # 只读:列每页 layout/表styleId/标题色
    python3 pptx_align.py layout      <pptx> --map "3:13,4:14" # 改 slide→layout 引用(逗号分隔 slideN:layoutN)
    python3 pptx_align.py tablestyle  <pptx> --style "{GUID}"  # 全部表统一为某 tableStyleId
    python3 pptx_align.py titlecolor  <pptx> --color bg1 [--pattern '^[（(][一二三四五六七八九十]']  # 标题run补颜色
    python3 pptx_align.py render      <pptx> [--pages 1,3,10] [--dpi 110] [--outdir DIR]  # soffice→PNG 验证

所有写操作默认 backup(.bak-YYYY... 由调用方负责或加 --backup);本模块写前自检 lsof 占用。
audit 是只读。改前先 audit + render 看现状(铁律: 样式对齐先建黄金基准再逐页 diff)。
"""
import argparse
import os
import re
import subprocess
import sys
import zipfile
from pathlib import Path

SLIDE_RE = re.compile(r'ppt/slides/slide(\d+)\.xml$')


def _slide_nums(z):
    nums = []
    for n in z.namelist():
        m = SLIDE_RE.match(n)
        if m:
            nums.append(int(m.group(1)))
    return sorted(nums)


def _lsof_guard(path):
    """Office 占用自检 — 占用则退出(调用方/用户先关或 kill)。"""
    try:
        r = subprocess.run(["lsof", str(path)], capture_output=True, text=True)
        if r.stdout.strip():
            print(f"⚠️  文件被占用(关闭 PowerPoint 后重试 或 backup+kill):\n{r.stdout}", file=sys.stderr)
            return False
    except FileNotFoundError:
        pass  # 无 lsof(非 mac)跳过
    return True


# ── audit:只读侦察 ────────────────────────────────────────────────
def cmd_audit(path, args):
    z = zipfile.ZipFile(path)
    # tableStyles 名称表
    name_of = {}
    if 'ppt/tableStyles.xml' in z.namelist():
        ts = z.read('ppt/tableStyles.xml').decode('utf-8', 'ignore')
        name_of = dict(re.findall(r'<a:tblStyle styleId="(\{[^}]*\})"\s+styleName="([^"]*)"', ts))

    print(f"# PPTX 结构侦察: {Path(path).name}")
    print(f"slides={len(_slide_nums(z))}  "
          f"layouts={len([n for n in z.namelist() if 'slideLayouts/slideLayout' in n and n.endswith('.xml')])}  "
          f"masters={len([n for n in z.namelist() if 'slideMasters/slideMaster' in n and n.endswith('.xml')])}  "
          f"themes={len([n for n in z.namelist() if n.startswith('ppt/theme/theme') and n.endswith('.xml')])}")
    print()
    print(f"{'slide':>6} | {'layout':>7} | {'表(styleId名)':<22} | {'标题色':<10} | 首文本")
    print("-" * 90)
    for num in _slide_nums(z):
        x = z.read(f'ppt/slides/slide{num}.xml').decode('utf-8', 'ignore')
        # layout
        relp = f'ppt/slides/_rels/slide{num}.xml.rels'
        lay = '?'
        if relp in z.namelist():
            m = re.search(r'slideLayout(\d+)\.xml', z.read(relp).decode())
            lay = m.group(1) if m else '?'
        # 表 styleId
        sids = re.findall(r'<a:tableStyleId>([^<]*)</a:tableStyleId>', x)
        tbl = ','.join(name_of.get(s, s[-8:]) for s in sids) if sids else '-'
        # 标题色(第一个 (一)(二) 或 中文数字标题 run 的 solidFill)
        tcolor = '-'
        for sp in re.findall(r'<p:sp>.*?</p:sp>', x, re.S):
            t = ''.join(re.findall(r'<a:t>([^<]*)</a:t>', sp)).strip()
            if re.match(r'^[（(][一二三四五六七八九十][）)]|^[一二三四五六七八九十]、', t):
                r = re.search(r'<a:r>.*?<a:t>', sp, re.S)
                rt = r.group(0) if r else ''
                fm = re.search(r'<a:solidFill>\s*<a:(?:schemeClr val="(\w+)"|srgbClr val="([0-9A-Fa-f]{6})")', rt)
                tcolor = (fm.group(1) or fm.group(2)) if fm else '黑/缺'
                break
        txt = ''.join(re.findall(r'<a:t>([^<]*)</a:t>', x))[:16]
        print(f"{num:>6} | {lay:>7} | {tbl:<22} | {tcolor:<10} | {txt}")
    return 0


# ── layout:改 slide→layout 引用 ────────────────────────────────
def cmd_layout(path, args):
    if not _lsof_guard(path):
        return 1
    mapping = {}
    for pair in args.map.split(','):
        s, l = pair.split(':')
        mapping[int(s)] = int(l)
    with zipfile.ZipFile(path) as z:
        items = {n: z.read(n) for n in z.namelist()}
        infos = {n: z.getinfo(n) for n in z.namelist()}
    changed = []
    for snum, lnum in mapping.items():
        rel = f'ppt/slides/_rels/slide{snum}.xml.rels'
        if rel not in items:
            print(f"  slide{snum} 无 rels,跳过", file=sys.stderr)
            continue
        txt = items[rel].decode()
        new = re.sub(r'slideLayout\d+\.xml', f'slideLayout{lnum}.xml', txt, count=1)
        if new != txt:
            items[rel] = new.encode()
            changed.append((snum, lnum))
    if not args.dry_run and changed:
        _rewrite(path, items, infos)
    print(f"{'[dry-run] ' if args.dry_run else ''}改 layout: " +
          ', '.join(f'slide{s}→layout{l}' for s, l in changed))
    return 0


# ── tablestyle:全部表统一 styleId ──────────────────────────────
def cmd_tablestyle(path, args):
    if not _lsof_guard(path):
        return 1
    sid = args.style if args.style.startswith('{') else '{' + args.style + '}'
    with zipfile.ZipFile(path) as z:
        items = {n: z.read(n) for n in z.namelist()}
        infos = {n: z.getinfo(n) for n in z.namelist()}
        # 校验 styleId 在 tableStyles.xml 里有定义
        if 'ppt/tableStyles.xml' in items:
            if sid not in items['ppt/tableStyles.xml'].decode('utf-8', 'ignore'):
                print(f"⚠️  styleId {sid} 不在 tableStyles.xml 中,改了会丢样式", file=sys.stderr)
                return 1
    cnt = 0
    for n in list(items):
        if SLIDE_RE.match(n):
            x = items[n].decode('utf-8', 'ignore')
            c = len(re.findall(r'<a:tableStyleId>', x))
            if c:
                x2 = re.sub(r'<a:tableStyleId>[^<]*</a:tableStyleId>',
                            f'<a:tableStyleId>{sid}</a:tableStyleId>', x)
                if x2 != x:
                    items[n] = x2.encode()
                    cnt += c
    if not args.dry_run and cnt:
        _rewrite(path, items, infos)
    print(f"{'[dry-run] ' if args.dry_run else ''}统一 {cnt} 张表 → {sid}")
    return 0


# ── titlecolor:给标题 run 补颜色 ───────────────────────────────
def cmd_titlecolor(path, args):
    if not _lsof_guard(path):
        return 1
    color = args.color
    fill = (f'<a:solidFill><a:schemeClr val="{color}"/></a:solidFill>'
            if not re.fullmatch(r'[0-9A-Fa-f]{6}', color)
            else f'<a:solidFill><a:srgbClr val="{color}"/></a:solidFill>')
    pat = re.compile(args.pattern)
    with zipfile.ZipFile(path) as z:
        items = {n: z.read(n) for n in z.namelist()}
        infos = {n: z.getinfo(n) for n in z.namelist()}
    pages = []
    for n in list(items):
        if not SLIDE_RE.match(n):
            continue
        num = int(SLIDE_RE.match(n).group(1))
        x = items[n].decode('utf-8', 'ignore')

        def fix_sp(m):
            sp = m.group(0)
            t = ''.join(re.findall(r'<a:t>([^<]*)</a:t>', sp)).strip()
            if not pat.match(t):
                return sp

            def addfill(rm):
                rpr = rm.group(0)
                if '<a:solidFill>' in rpr:
                    return rpr
                if rpr.endswith('/>'):
                    return rpr[:-2] + '>' + fill + '</a:rPr>'
                if '<a:latin' in rpr:
                    return rpr.replace('<a:latin', fill + '<a:latin', 1)
                return rpr.replace('</a:rPr>', fill + '</a:rPr>', 1)
            return re.sub(r'<a:rPr[^>]*(?:/>|>.*?</a:rPr>)', addfill, sp, count=1, flags=re.S)
        x2 = re.sub(r'<p:sp>.*?</p:sp>', fix_sp, x, flags=re.S)
        if x2 != x:
            items[n] = x2.encode()
            pages.append(num)
    if not args.dry_run and pages:
        _rewrite(path, items, infos)
    print(f"{'[dry-run] ' if args.dry_run else ''}标题补色 {color}: slide{pages}")
    return 0


# ── render:soffice → PNG 验证 ──────────────────────────────────
def cmd_render(path, args):
    soffice = next((p for p in ['/opt/homebrew/bin/soffice',
                                '/Applications/LibreOffice.app/Contents/MacOS/soffice',
                                'soffice', 'libreoffice'] if _which(p)), None)
    if not soffice:
        print("⚠️  未找到 LibreOffice(soffice)", file=sys.stderr)
        return 1
    outdir = Path(args.outdir or '/tmp/pptx-render')
    outdir.mkdir(parents=True, exist_ok=True)
    pdf = outdir / (Path(path).stem + '.pdf')
    subprocess.run([soffice, '--headless', '--convert-to', 'pdf',
                    '--outdir', str(outdir), str(path)],
                   capture_output=True, timeout=180)
    if not pdf.exists():
        print("⚠️  soffice 转 PDF 失败", file=sys.stderr)
        return 1
    pages = args.pages.split(',') if args.pages else None
    if pages:
        for p in pages:
            subprocess.run(['pdftoppm', '-png', '-r', str(args.dpi),
                            '-f', p, '-l', p, str(pdf), str(outdir / f'p{p}')],
                           capture_output=True)
    else:
        subprocess.run(['pdftoppm', '-png', '-r', str(args.dpi), str(pdf), str(outdir / 'page')],
                       capture_output=True)
    pngs = sorted(outdir.glob('*.png'))
    print(f"渲染 {len(pngs)} 张 → {outdir}/")
    for p in pngs:
        print(f"  {p}")
    return 0


def _which(cmd):
    if '/' in cmd:
        return Path(cmd).exists()
    return subprocess.run(['which', cmd], capture_output=True).returncode == 0


def _rewrite(path, items, infos):
    tmp = str(path) + '.tmp'
    with zipfile.ZipFile(tmp, 'w', zipfile.ZIP_DEFLATED) as zo:
        for n, d in items.items():
            zo.writestr(infos[n], d)
    os.replace(tmp, path)


def main():
    ap = argparse.ArgumentParser(prog='pptx_align', description='PPTX 样式对齐工具集')
    sub = ap.add_subparsers(dest='cmd', required=True)

    a = sub.add_parser('audit', help='只读:列每页 layout/表styleId/标题色')
    a.add_argument('pptx')

    l = sub.add_parser('layout', help='改 slide→layout 引用')
    l.add_argument('pptx')
    l.add_argument('--map', required=True, help='逗号分隔 slideN:layoutN,如 "3:13,4:14"')
    l.add_argument('--dry-run', action='store_true')

    t = sub.add_parser('tablestyle', help='全部表统一为某 tableStyleId')
    t.add_argument('pptx')
    t.add_argument('--style', required=True, help='tableStyleId GUID(带或不带花括号)')
    t.add_argument('--dry-run', action='store_true')

    c = sub.add_parser('titlecolor', help='标题 run 补颜色(schemeClr 名或 srgb 6位)')
    c.add_argument('pptx')
    c.add_argument('--color', required=True, help='bg1/tx1/accent1.. 或 FFFFFF')
    c.add_argument('--pattern', default=r'^[（(][一二三四五六七八九十][）)]',
                   help='标题文本匹配正则(默认 (一)(二)类)')
    c.add_argument('--dry-run', action='store_true')

    r = sub.add_parser('render', help='soffice→PNG 渲染验证')
    r.add_argument('pptx')
    r.add_argument('--pages', help='逗号分隔页号(演示顺序),如 1,3,10;省略=全部')
    r.add_argument('--dpi', type=int, default=110)
    r.add_argument('--outdir')

    args = ap.parse_args()
    path = Path(args.pptx)
    if not path.exists():
        print(f"文件不存在: {path}", file=sys.stderr)
        return 1
    return {'audit': cmd_audit, 'layout': cmd_layout, 'tablestyle': cmd_tablestyle,
            'titlecolor': cmd_titlecolor, 'render': cmd_render}[args.cmd](path, args)


if __name__ == '__main__':
    sys.exit(main())
