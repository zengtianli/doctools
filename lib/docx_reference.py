"""Apply observed reference roles to a DOCX copy, retaining its package parts.

The caller supplies paragraph roles and explicit, hash-bound rewrite ranges.
This module owns OOXML formatting and repacking; it does not infer business
structure or silently copy reference text, numbering, headers, or attachments.
"""
from __future__ import annotations

import copy
import hashlib
from pathlib import Path
import re
import shutil
import zipfile

from lxml import etree
from docx_surgical import parse_document, para_text, serialize, surgical_rewrite, qn
from docx_parts import assert_parts_intact

P_TAGS = ('keepNext', 'keepLines', 'pageBreakBefore', 'widowControl', 'pBdr',
          'shd', 'tabs', 'spacing', 'ind', 'contextualSpacing', 'jc', 'outlineLvl')
R_TAGS = ('rFonts', 'b', 'bCs', 'i', 'iCs', 'color', 'spacing', 'sz', 'szCs',
          'u', 'vertAlign', 'lang', 'highlight', 'shd')
R_ORDER = ('rStyle', 'rFonts', 'b', 'bCs', 'i', 'iCs', 'caps', 'smallCaps',
           'strike', 'dstrike', 'outline', 'shadow', 'emboss', 'imprint',
           'noProof', 'snapToGrid', 'vanish', 'webHidden', 'color', 'spacing',
           'w', 'kern', 'position', 'sz', 'szCs', 'highlight', 'u', 'effect',
           'bdr', 'shd', 'fitText', 'vertAlign', 'rtl', 'cs', 'em', 'lang')
P_ORDER = ('pStyle', 'keepNext', 'keepLines', 'pageBreakBefore', 'framePr',
           'widowControl', 'numPr', 'suppressLineNumbers', 'pBdr', 'shd',
           'tabs', 'suppressAutoHyphens', 'kinsoku', 'wordWrap', 'overflowPunct',
           'topLinePunct', 'autoSpaceDE', 'autoSpaceDN', 'bidi', 'adjustRightInd',
           'snapToGrid', 'spacing', 'ind', 'contextualSpacing', 'mirrorIndents',
           'suppressOverlap', 'jc', 'textDirection', 'textAlignment',
           'textboxTightWrap', 'outlineLvl', 'divId', 'cnfStyle', 'rPr', 'sectPr')


def digest(path):
    return hashlib.sha256(Path(path).read_bytes()).hexdigest()


def logical_blocks(root):
    """Stemma/source-reader identity includes empty top-level paragraphs."""
    return [node for node in root.find(qn('w:body'))
            if node.tag in {qn('w:tbl'), qn('w:p')}]


def _properties(node, name):
    return node.find(qn('w:' + name)) if node is not None else None


def _merge(*nodes):
    result = {}
    for node in nodes:
        if node is not None:
            for child in node:
                name = etree.QName(child).localname
                previous = result.get(name)
                if previous is not None and name in {'rFonts', 'lang', 'spacing', 'ind'}:
                    merged = copy.deepcopy(previous)
                    merged.attrib.update(child.attrib)
                    if name == 'rFonts':
                        for field in ('ascii','hAnsi','eastAsia','cs'):
                            if child.get(qn('w:' + field)) is not None:
                                merged.attrib.pop(qn('w:' + field + 'Theme'), None)
                            elif child.get(qn('w:' + field + 'Theme')) is not None:
                                merged.attrib.pop(qn('w:' + field), None)
                    result[name] = merged
                else:
                    result[name] = copy.deepcopy(child)
    return result


def shell(reference, block, *, colors='reference', indent=None):
    """Resolve defaults, style inheritance, paragraph and first visible run."""
    root = parse_document(Path(reference))
    blocks = logical_blocks(root)
    if type(block) is not int or not 0 <= block < len(blocks):
        raise ValueError('范式段落索引无效')
    paragraph = blocks[block]
    if paragraph.tag != qn('w:p'):
        raise ValueError('范式角色须指向正文段落')
    with zipfile.ZipFile(reference) as archive:
        styles = etree.fromstring(archive.read('word/styles.xml'))
    catalog = {n.get(qn('w:styleId')): n for n in styles.findall(qn('w:style'))}
    default_style = next((n.get(qn('w:styleId')) for n in catalog.values()
                          if n.get(qn('w:default')) == '1' and
                          n.get(qn('w:type')) == 'paragraph'), None)
    defaults = styles.find(qn('w:docDefaults'))
    dp = defaults.find(qn('w:pPrDefault')) if defaults is not None else None
    dr = defaults.find(qn('w:rPrDefault')) if defaults is not None else None
    ppr = _properties(paragraph, 'pPr')
    ps = ppr.find(qn('w:pStyle')) if ppr is not None else None
    sid = ps.get(qn('w:val')) if ps is not None else default_style
    chain, seen = [], set()
    while sid:
        if sid in seen or sid not in catalog:
            raise ValueError('范式样式继承循环或缺失：' + sid)
        seen.add(sid)
        item = catalog[sid]; chain.insert(0, item)
        parent = item.find(qn('w:basedOn'))
        sid = parent.get(qn('w:val')) if parent is not None else None
    p = _merge(_properties(dp, 'pPr'), *[_properties(n, 'pPr') for n in chain], ppr)
    run = next((n for n in paragraph.iter(qn('w:r')) if para_text(n).strip()), None)
    cs = _properties(run, 'rPr')
    rs = cs.find(qn('w:rStyle')) if cs is not None else None
    chars, seen = [], set()
    sid = rs.get(qn('w:val')) if rs is not None else None
    while sid:
        if sid in seen or sid not in catalog:
            raise ValueError('范式字符样式继承循环或缺失：' + sid)
        seen.add(sid); item = catalog[sid]; chars.insert(0, item)
        parent = item.find(qn('w:basedOn'))
        sid = parent.get(qn('w:val')) if parent is not None else None
    r = _merge(_properties(dr, 'rPr'), *[_properties(n, 'rPr') for n in chain],
               _properties(ppr, 'rPr'), *[_properties(n, 'rPr') for n in chars], cs)
    # Explicit off prevents the destination style's bold/italic/underline bleed.
    for tag in ('b', 'bCs', 'i', 'iCs'):
        if tag not in r:
            r[tag] = etree.Element(qn('w:' + tag), {qn('w:val'): '0'})
    if 'u' not in r:
        r['u'] = etree.Element(qn('w:u'), {qn('w:val'): 'none'})
    if colors == 'black':
        r['color'] = etree.Element(qn('w:color'), {qn('w:val'): '000000'})
    elif colors != 'reference':
        raise ValueError('颜色策略须为 reference 或 black')
    if indent is not None:
        p['ind'] = etree.Element(qn('w:ind'), {qn('w:firstLineChars'): str(round(indent * 100))})
    return {'p': {k: v for k, v in p.items() if k in P_TAGS},
            'r': {k: v for k, v in r.items() if k in R_TAGS}}


def _ordered(node, order):
    nodes = list(node)
    ranks = {name: i for i, name in enumerate(order)}
    for child in nodes: node.remove(child)
    for child in sorted(nodes, key=lambda n: ranks.get(etree.QName(n).localname, len(order))):
        node.append(child)


def _apply_paragraph(paragraph, style):
    ppr = paragraph.find(qn('w:pPr'))
    if ppr is None:
        ppr = etree.Element(qn('w:pPr')); paragraph.insert(0, ppr)
    for tag in P_TAGS:
        for n in ppr.findall(qn('w:' + tag)): ppr.remove(n)
    for n in style['p'].values(): ppr.append(copy.deepcopy(n))
    _ordered(ppr, P_ORDER)
    for run in paragraph.iter(qn('w:r')):
        rpr = run.find(qn('w:rPr'))
        if rpr is None:
            rpr = etree.Element(qn('w:rPr')); run.insert(0, rpr)
        # Preserve scientific superscripts/subscripts and source emphasis.
        special = {tag: copy.deepcopy(rpr.find(qn('w:' + tag))) for tag in ('vertAlign',)}
        bold = rpr.find(qn('w:b'))
        if bold is not None: special['b'] = copy.deepcopy(bold)
        for tag in R_TAGS:
            for n in rpr.findall(qn('w:' + tag)): rpr.remove(n)
        for n in style['r'].values(): rpr.append(copy.deepcopy(n))
        for tag, n in special.items():
            if n is not None:
                for old in rpr.findall(qn('w:' + tag)): rpr.remove(old)
                rpr.append(n)
        _ordered(rpr, R_ORDER)


def _simple(paragraph):
    """Rewriting has a narrower contract than formatting; don't flatten objects."""
    for n in paragraph:
        if n.tag == qn('w:pPr'):
            if n.find(qn('w:sectPr')) is not None or n.find(qn('w:numPr')) is not None:
                return False
        elif n.tag != qn('w:r') or any(c.tag not in {qn('w:rPr'), qn('w:t')} for c in n):
            return False
    return True


def _paragraph_like(node, text):
    new = copy.deepcopy(node)
    for child in list(new):
        if child.tag != qn('w:pPr'): new.remove(child)
    run = etree.SubElement(new, qn('w:r'))
    t = etree.SubElement(run, qn('w:t'))
    t.set('{http://www.w3.org/XML/1998/namespace}space', 'preserve'); t.text = text
    return new


def apply(source, reference, output, *, assignments, roles, replacements=(),
          colors='reference', body_indent=None, page_layout=True, tables=True):
    source, reference, output = map(lambda p: Path(p).expanduser().resolve(), (source, reference, output))
    if output.exists() or output in {source, reference}:
        raise ValueError('候选输出必须是尚不存在的新文件')
    source_hash, ref_hash = digest(source), digest(reference)
    root = parse_document(source); original = logical_blocks(root)
    with zipfile.ZipFile(source) as z:
        for name in z.namelist():
            if name.startswith('word/') and name.endswith('.xml'):
                part = etree.fromstring(z.read(name))
                if any(etree.QName(n).localname in {'ins','del','moveFrom','moveTo'} for n in part.iter()):
                    raise ValueError('源稿有未接受修订，先确定内容；不自动接受修订')
    shells = {role: shell(reference, block, colors=colors,
                         indent=body_indent if role in {'body', 'heading2', 'task_heading'} else None)
              for role, block in roles.items()}
    for item in assignments:
        index, role = item['block'], item['role']
        if type(index) is not int or not 0 <= index < len(original): raise ValueError('目标段落索引无效')
        if original[index].tag != qn('w:p'): raise ValueError('正文角色不能指向表格')
        if role not in shells: raise ValueError('缺少范式角色：' + role)
    occupied = set()
    for change in replacements:
        start, end = change['start'], change['end']
        if type(start) is not int or type(end) is not int or not 0 <= start <= end < len(original):
            raise ValueError('改写区间无效')
        interval = set(range(start, end + 1))
        if occupied & interval: raise ValueError('改写区间重叠')
        occupied |= interval
        if not change.get('paragraphs'): raise ValueError('改写不能静默删除内容')
        for n in original[start:end + 1]:
            if n.tag != qn('w:p') or not _simple(n):
                raise ValueError('改写区间含表格、批注、书签、域、媒体或自动编号；保留该区间并明确处理')
        for p in change['paragraphs']:
            if not p.get('text', '').strip() or p.get('role', 'body') not in shells:
                raise ValueError('改写段落文字或角色无效')
    for item in assignments: _apply_paragraph(original[item['block']], shells[item['role']])
    replacements = sorted(replacements, key=lambda c: c['start'], reverse=True)
    rewritten = {}
    for change in replacements:
        start, end = change['start'], change['end']; first = original[start]
        parent, position = first.getparent(), first.getparent().index(first)
        created = []
        for p in change['paragraphs']:
            new = _paragraph_like(first, p['text']); _apply_paragraph(new, shells[p.get('role', 'body')])
            parent.insert(position + len(created), new); created.append(new)
        for old in original[start:end + 1]: parent.remove(old)
        for index in range(start, end + 1): rewritten[index] = created
    # Border and cell typography only. Target widths, merges, tables and text stay.
    if tables and 'body' in shells:
        ref_root = parse_document(reference)
        ref_table = next(ref_root.iter(qn('w:tbl')), None)
        if ref_table is not None:
            ref_pr = ref_table.find(qn('w:tblPr'))
            ref_border = ref_pr.find(qn('w:tblBorders')) if ref_pr is not None else None
            cell_run = next((r for r in ref_table.iter(qn('w:rPr')) if r.find(qn('w:sz')) is not None), None)
            typography = copy.deepcopy(shells['body'])
            if cell_run is not None:
                typography['r'].update({etree.QName(n).localname: copy.deepcopy(n) for n in cell_run if etree.QName(n).localname in {'rFonts','sz','szCs'}})
            typography['p'] = {}
            for table in root.iter(qn('w:tbl')):
                pr = table.find(qn('w:tblPr'))
                if pr is None: pr = etree.Element(qn('w:tblPr')); table.insert(0, pr)
                if ref_border is not None:
                    for old in pr.findall(qn('w:tblBorders')): pr.remove(old)
                    pr.append(copy.deepcopy(ref_border))
                for paragraph in table.iter(qn('w:p')):
                    # Keep the target's cell alignment, spacing and emphasis.
                    ppr = copy.deepcopy(paragraph.find(qn('w:pPr')))
                    _apply_paragraph(paragraph, typography)
                    current = paragraph.find(qn('w:pPr'))
                    paragraph.remove(current)
                    if ppr is not None: paragraph.insert(0, ppr)
                if ref_border is not None:
                    for cell in table.iter(qn('w:tc')):
                        cp = cell.find(qn('w:tcPr'))
                        if cp is not None:
                            for old in cp.findall(qn('w:tcBorders')): cp.remove(old)
    if page_layout:
        ref_section = next(parse_document(reference).iter(qn('w:sectPr')), None)
        if ref_section is not None:
            for section in root.iter(qn('w:sectPr')):
                for tag in ('pgSz', 'pgMar'):
                    sample = ref_section.find(qn('w:' + tag))
                    if sample is not None:
                        old = section.find(qn('w:' + tag))
                        if old is not None: section.replace(old, copy.deepcopy(sample))
                        else: section.append(copy.deepcopy(sample))
    final = logical_blocks(root)
    mapping = {str(i): [final.index(n) for n in rewritten.get(i, [node])]
               for i, node in enumerate(original)}
    output.parent.mkdir(parents=True, exist_ok=True)
    # Create exclusively; surgical_rewrite operates only on this new copy.
    with output.open('xb') as stream: stream.write(source.read_bytes())
    try:
        surgical_rewrite(output, serialize(root), backup=False)
        assert_parts_intact(source, output, allow_changed={'word/document.xml'},
                            allow_media_loss=False, verbose=False)
        if digest(source) != source_hash or digest(reference) != ref_hash:
            raise ValueError('执行期间输入发生变化，候选不能作为当前结果')
    except Exception:
        output.unlink(missing_ok=True)
        raise
    return {'ok': True, 'source_sha256': source_hash, 'reference_sha256': ref_hash,
            'output': str(output), 'block_mapping': mapping,
            'rewritten_ranges': len(replacements), 'changed_parts': ['word/document.xml']}
