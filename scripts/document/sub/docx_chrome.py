#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""docx_chrome.py — 自动复刻院报告版面装帧「逐章分节 + 逐章页眉页脚水印 + 宽表/宽图横向节」.

引擎(distilled from eco-flow/taizhou-天台 chrome_full.py, 2026-06-04). 输入 raw 合并 docx
(正文已成型) → 重建节结构:
  · 前言~第5章: 纵向基底; 宽表(gridCol 顶层总宽>11000twip)或宽图(extent cx>纵向可用宽)
                就地包成横向节(表名段在上随表横向, 图名段在下随图横向, 尾随空段并入)
  · 附表/附图 : 整章横向
  · 每章一个 running-title 页眉(标题+章名 前言/第N章/附表/附图) + 院名页码页脚; 横向节用横向页眉
页眉/页脚部件取自 --template(已审定范式 docx), 县名正则 swap 可复用他县.
纯 zip + lxml surgery: 媒体/embeddings/OMML公式/OLE verbatim, 不剥公式.

CLI(独立可跑, 也经 docx_cli.py `chrome` 子命令转发):
  python3 sub/docx_chrome.py --raw <raw.docx> --template <范式.docx> [--out <out.docx>] [--county 天台县]
  python3 sub/docx_chrome.py --validate <out.docx> <范式.docx>   # diff 节结构(orient+章名)
默认输出 = <raw stem>_chrome.docx (同目录).
"""
from __future__ import annotations
import argparse, zipfile, sys, re
from pathlib import Path
from lxml import etree

W   = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
R   = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
CT  = 'http://schemas.openxmlformats.org/package/2006/content-types'
PR  = 'http://schemas.openxmlformats.org/package/2006/relationships'
def w(t): return f'{{{W}}}{t}'
def rr(t): return f'{{{R}}}{t}'

AVAIL_PORTRAIT = 8787              # 11906 - 1418(左) - 1701(右)
LAND_MIN_TWIP  = 11000            # 宽表→横向阈值(纵向窄表~9000, 横向宽表~12600+; gap 取中)
PORT_EMU       = AVAIL_PORTRAIT*635   # 纵向可用宽(EMU), 判宽图
DRAW_EXT = '{http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing}extent'
HDR_CT = 'application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml'
FTR_CT = 'application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml'

PG_PORT = {'w':'11906','h':'16838'}
PG_LAND = {'w':'16838','h':'11906','orient':'landscape'}
MAR_PORT = dict(top='1701',right='1701',bottom='1418',left='1418',header='1304',footer='1134',gutter='0')
MAR_LAND = dict(top='1701',right='1418',bottom='1418',left='1701',header='1304',footer='1134',gutter='0')

# ── 章节识别 ───────────────────────────────────────────────
def chapter_label(p):
    """段是章界则返回 label(前言/第N章/附表/附图), 否则 None."""
    ps = p.find(f'{w("pPr")}/{w("pStyle")}')
    if ps is None: return None
    sv = ps.get(w('val'))
    txt = ''.join(p.itertext()).strip()
    if sv == '1':
        if txt.startswith('前'): return '前言'
        m = re.match(r'^([1-5])', txt)
        if m: return f'第{m.group(1)}章'
    if sv == 'ZDWP5':
        if txt.startswith('附表'): return '附表'
        if txt.startswith('附图'): return '附图'
    return None

def grid_width(tbl):
    g = tbl.find(w('tblGrid'))          # 仅本表顶层 grid, 不数嵌套表(否则暴增误判)
    if g is None: return 0
    return sum(int(c.get(w('w'))) for c in g.findall(w('gridCol')) if c.get(w('w')))

def is_wide_block(el):
    t = etree.QName(el).localname
    if t == 'tbl': return grid_width(el) > LAND_MIN_TWIP
    if t == 'p':
        for ext in el.iter(DRAW_EXT):
            if int(ext.get('cx') or 0) > PORT_EMU: return True
    return False

# ── 模板部件抽取 ───────────────────────────────────────────
def load_template_parts(tpl_path, county):
    """从范式 docx 抽 每章(label,orient)->header_xml + 默认 footer_xml(port/land)."""
    z = zipfile.ZipFile(tpl_path)
    rels = {m.group(1): m.group(2) for m in
            re.finditer(r'Id="(rId\d+)"[^>]*Target="([^"]+)"',
                        z.read('word/_rels/document.xml.rels').decode())}
    doc = etree.fromstring(z.read('word/document.xml'))
    hdr = {}      # (label, orient) -> header bytes
    ftr = {}      # orient -> footer bytes (院名页码, 通用)
    for s in doc.findall(f'.//{w("sectPr")}'):
        pg = s.find(w('pgSz'))
        orient = 'land' if (pg is not None and pg.get(w('orient'))=='landscape') else 'port'
        hr = s.find(f'{w("headerReference")}[@{w("type")}="default"]')
        fr = s.find(f'{w("footerReference")}[@{w("type")}="default"]')
        if hr is not None:
            hx = z.read('word/'+rels[hr.get(rr("id"))])
            txt = ''.join(re.findall(r'<w:t[^>]*>(.*?)</w:t>', hx.decode('utf-8','ignore')))
            for lab in ('前言','第1章','第2章','第3章','第4章','第5章','附表','附图'):
                if lab in txt: hdr.setdefault((lab, orient), hx)
        if fr is not None and orient not in ftr:
            fx = z.read('word/'+rels[fr.get(rr("id"))])
            if 'PAGE' in fx.decode('utf-8','ignore'): ftr[orient] = fx
    def swap(b):
        s = b.decode('utf-8')
        return re.sub(r'[一-龥]{2,3}县小型水库生态流量核定与保障实施方案',
                      f'{county}小型水库生态流量核定与保障实施方案', s).encode('utf-8')
    hdr = {k: swap(v) for k,v in hdr.items()}
    if 'land' not in ftr and 'port' in ftr: ftr['land'] = ftr['port']
    return hdr, ftr

# ── 构造 sectPr ────────────────────────────────────────────
def make_sectpr(hdr_rid, ftr_rid, orient):
    pg = PG_LAND if orient=='land' else PG_PORT
    mar = MAR_LAND if orient=='land' else MAR_PORT
    s = etree.SubElement(etree.Element('root'), w('sectPr'))
    hRef = etree.SubElement(s, w('headerReference')); hRef.set(w('type'),'default'); hRef.set(rr('id'),hdr_rid)
    fRef = etree.SubElement(s, w('footerReference')); fRef.set(w('type'),'default'); fRef.set(rr('id'),ftr_rid)
    pgsz = etree.SubElement(s, w('pgSz'))
    for k,v in pg.items(): pgsz.set(w(k),v)
    pgmar = etree.SubElement(s, w('pgMar'))
    for k,v in mar.items(): pgmar.set(w(k),v)
    etree.SubElement(s, w('cols')).set(w('space'),'425')
    etree.SubElement(s, w('docGrid')).set(w('linePitch'),'312')
    return s

def sectpr_para(sectpr):
    p = etree.Element(w('p')); ppr = etree.SubElement(p, w('pPr')); ppr.append(sectpr); return p

# ── 主流程 ─────────────────────────────────────────────────
def build(raw, tpl, out, county):
    hdr, ftr = load_template_parts(tpl, county)
    zin = zipfile.ZipFile(raw)
    doc = etree.fromstring(zin.read('word/document.xml'))
    body = doc.find(w('body'))
    kids = list(body)
    first_ch = next((i for i,el in enumerate(kids)
                     if etree.QName(el).localname=='p' and chapter_label(el)), None)
    if first_ch is None: sys.exit('未找到章节标题(前言/第N章 pStyle=1, 附表/附图 pStyle=ZDWP5)')

    rels_xml = etree.fromstring(zin.read('word/_rels/document.xml.rels'))
    used_rids = [int(re.search(r'\d+',e.get('Id')).group()) for e in rels_xml]
    next_rid = max(used_rids)+1
    new_parts = {}; new_rels = []; cache = {}; pidx = [100]
    def reg_hdr(label, orient):
        key=(label,orient)
        if key in cache: return cache[key]
        data = hdr.get(key) or hdr.get((label,'port')) or hdr.get((label,'land'))
        if data is None: sys.exit(f'范式缺页眉部件: {key}')
        name=f'word/header{pidx[0]}.xml'; pidx[0]+=1
        nonlocal next_rid; rid=f'rId{next_rid}'; next_rid+=1
        new_parts[name]=(data,HDR_CT); new_rels.append((rid,name.split('/')[1],
            'http://schemas.openxmlformats.org/officeDocument/2006/relationships/header'))
        cache[key]=rid; return rid
    def reg_ftr(orient):
        key=('ftr',orient)
        if key in cache: return cache[key]
        data=ftr.get(orient) or ftr.get('port')
        if data is None: sys.exit('范式缺页脚部件(含 PAGE 域)')
        name=f'word/footer{pidx[0]}.xml'; pidx[0]+=1
        nonlocal next_rid; rid=f'rId{next_rid}'; next_rid+=1
        new_parts[name]=(data,FTR_CT); new_rels.append((rid,name.split('/')[1],
            'http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer'))
        cache[key]=rid; return rid

    body_final = kids[-1] if etree.QName(kids[-1]).localname == 'sectPr' else None
    end = len(kids) - (1 if body_final is not None else 0)
    body_blocks = kids[first_ch:end]
    for el in body_blocks:
        if etree.QName(el).localname=='p':
            sp=el.find(f'{w("pPr")}/{w("sectPr")}')
            if sp is not None: sp.getparent().remove(sp)
    body_final = body.find(w('sectPr'))

    N=len(body_blocks)
    def is_empty_p(el):
        return etree.QName(el).localname=='p' and not ''.join(el.itertext()).strip() and chapter_label(el) is None
    def is_caption_p(el):
        if etree.QName(el).localname!='p': return False
        t=''.join(el.itertext()).strip(); return t.startswith('表') or t.startswith('图')
    chap=[None]*N; cur=None
    for k,el in enumerate(body_blocks):
        lab = chapter_label(el) if etree.QName(el).localname=='p' else None
        if lab: cur=lab
        chap[k]=cur
    orient=['port']*N
    for k in range(N):
        if chap[k] in ('附表','附图'): orient[k]='land'
    for k in range(N):
        if is_wide_block(body_blocks[k]):
            orient[k]='land'
            if k-1>=0 and is_caption_p(body_blocks[k-1]) and chap[k-1]==chap[k]:
                orient[k-1]='land'              # 表名在上
            j=k+1                                # 图名在下 + 尾随空段
            while j<N and chap[j]==chap[k] and (is_empty_p(body_blocks[j]) or is_caption_p(body_blocks[j])):
                orient[j]='land'; j+=1
    sections=[]
    for k,el in enumerate(body_blocks):
        newchap = etree.QName(el).localname=='p' and chapter_label(el) is not None
        if (not sections or sections[-1]['chapter']!=chap[k]
                or sections[-1]['orient']!=orient[k] or newchap):
            sections.append(dict(chapter=chap[k],orient=orient[k],blocks=[]))
        sections[-1]['blocks'].append(el)
    sections=[s for s in sections if s['blocks']]

    new_body_children=[]; last_sectpr=None
    for si,sec in enumerate(sections):
        new_body_children.extend(sec['blocks'])
        hrid=reg_hdr(sec['chapter'],sec['orient']); frid=reg_ftr(sec['orient'])
        sp=make_sectpr(hrid,frid,sec['orient'])
        if si==len(sections)-1: last_sectpr=sp
        else: new_body_children.append(sectpr_para(sp))
    for el in body_blocks: body.remove(el)
    if body_final is not None: body.remove(body_final)
    anchor_idx=first_ch
    for el in new_body_children:
        body.insert(anchor_idx,el); anchor_idx+=1
    body.append(last_sectpr)

    for rid,target,reltype in new_rels:
        e=etree.SubElement(rels_xml,f'{{{PR}}}Relationship')
        e.set('Id',rid); e.set('Type',reltype); e.set('Target',target)
    ct=etree.fromstring(zin.read('[Content_Types].xml'))
    for name,(data,cttype) in new_parts.items():
        ov=etree.SubElement(ct,f'{{{CT}}}Override'); ov.set('PartName','/'+name); ov.set('ContentType',cttype)

    with zipfile.ZipFile(out,'w',zipfile.ZIP_DEFLATED) as zout:
        for it in zin.infolist():
            if it.filename=='word/document.xml':
                data=etree.tostring(doc,xml_declaration=True,encoding='UTF-8',standalone=True)
            elif it.filename=='word/_rels/document.xml.rels':
                data=etree.tostring(rels_xml,xml_declaration=True,encoding='UTF-8',standalone=True)
            elif it.filename=='[Content_Types].xml':
                data=etree.tostring(ct,xml_declaration=True,encoding='UTF-8',standalone=True)
            else:
                data=zin.read(it.filename)
            zi=zipfile.ZipInfo(it.filename,date_time=it.date_time)
            zi.compress_type=it.compress_type; zi.external_attr=it.external_attr
            zout.writestr(zi,data)
        for name,(data,cttype) in new_parts.items():
            zout.writestr(name,data)
    from collections import Counter
    print(f'OK  -> {out}')
    print(f'    节数={len(sections)}  新部件={len(new_parts)}(header+footer)')
    for k,v in Counter((s['chapter'],s['orient']) for s in sections).items():
        print(f'      {k[0]:5}/{k[1]:4} ×{v}')

def validate(out, tpl):
    def struct(p):
        z=zipfile.ZipFile(p)
        rels={m.group(1):m.group(2) for m in re.finditer(r'Id="(rId\d+)"[^>]*Target="([^"]+)"',z.read('word/_rels/document.xml.rels').decode())}
        doc=etree.fromstring(z.read('word/document.xml')); res=[]
        for s in doc.findall(f'.//{w("sectPr")}'):
            pg=s.find(w('pgSz')); o='land' if (pg is not None and pg.get(w('orient'))=='landscape') else 'port'
            hr=s.find(f'{w("headerReference")}[@{w("type")}="default"]'); txt=''
            if hr is not None:
                hx=z.read('word/'+rels[hr.get(rr("id"))]).decode('utf-8','ignore')
                txt=''.join(t for t in re.findall(r'<w:t[^>]*>(.*?)</w:t>',hx) if '<' not in t).strip()[-6:]
            res.append((o,txt))
        return res
    a,b=struct(out),struct(tpl)
    print(f'生成 {len(a)} 节 vs 范式 {len(b)} 节')
    for i in range(max(len(a),len(b))):
        xa=a[i] if i<len(a) else None; xb=b[i] if i<len(b) else None
        print(f'  {"✓" if xa==xb else "✗"} s{i:2} 生成={xa}  范式={xb}')

def main() -> int:
    ap=argparse.ArgumentParser(prog='docx chrome', description='院报告版面装帧: 逐章分节+逐章页眉页脚水印+宽表横向节')
    ap.add_argument('--raw', help='输入 raw 合并 docx(正文已成型)')
    ap.add_argument('--template', help='范式 docx(取页眉/页脚/水印部件)')
    ap.add_argument('--out', help='输出(默认 <raw stem>_chrome.docx)')
    ap.add_argument('--county', default='天台县', help='目标县名(替换范式县名, 默认 天台县)')
    ap.add_argument('--validate', nargs=2, metavar=('OUT','TEMPLATE'), help='diff 节结构而非构建')
    a=ap.parse_args()
    if a.validate:
        validate(*a.validate); return 0
    if not a.raw or not a.template:
        ap.error('--raw 和 --template 必填(或用 --validate)')
    out = a.out or str(Path(a.raw).with_name(Path(a.raw).stem + '_chrome.docx'))
    build(a.raw, a.template, out, a.county)
    return 0

if __name__=='__main__':
    sys.exit(main())
