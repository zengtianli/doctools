#!/usr/bin/env python3
"""docx_revise — 修订注入引擎：Word 修订标记(w:ins/w:del) + 批注，意见=数据(ops)。

为什么有这个文件：2026-07-31 qual-supply 一晚现编了 5 个注入脚本(4,030 行)，
同一套 XML 构件(取段文本/采样 rPr/造修订段/挂批注/注册 comments.xml)手搓了 5 遍，
专家意见原文全部硬编码在 py 常量里。内容与引擎必须解耦：引擎在这里，
意见写 ops.yaml(见 config/spec-examples/revise-ops-example.yaml)，项目里只留数据。

做法 = surgical：只重写 word/document.xml(+comments.xml 及其注册)，
其余部件逐字节 verbatim；禁 python-docx(会剥 OLE/原生图表——R2 曾因此丢 11 个 chart 整轮重做)。
存盘后 assert_parts_intact 强制断言，白名单外任何部件变动 = 抛异常。

四种 op(action 字段)：
  comment       只在锚段挂批注
  insert_after  锚段后插 N 个新段(整段包 w:ins，段落标记也标插入 → Word 里一键整段接受/拒绝)
  replace       锚段内 old→new，w:del+w:ins 修订对
  replace_para  整段换掉(默认 tracked：旧 runs 包 w:del；tracked:false 用于换掉上一轮
                自己插的〔待补〕占位段——删自己的 w:ins 等价于拒绝它，无需再留痕)

锚点纪律(两条都是真事故换来的，别放松)：
  1. find 文本在检索范围内必须**唯一命中**，0 或 ≥2 一律抛错并列出所有命中位置。
     '在区域供用水现状的基础上' 在 §4.2.2 与 §5.1 逐字重复，首命中策略曾静默插错章 ——
     多命中时用 within(通常写节标题文本)收窄检索起点，而不是赌第一个是对的。
  2. 锚点检索跳过目录段(pStyle 以 TOC 开头)。目录项与正文标题文本几乎相同，
     纯文本判据曾把检索起点钉在目录里。
"""
from __future__ import annotations

import copy
import re
import shutil
import zipfile
from pathlib import Path

from lxml import etree

from docx_parts import assert_parts_intact

WNS = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
XML_SPACE = '{http://www.w3.org/XML/1998/namespace}space'


class ReviseError(SystemExit):
    """锚点/参数问题一律带上下文抛出——引擎宁可停也不静默插错位置。"""


def q(tag: str) -> str:
    return '{%s}%s' % (WNS, tag)


def mk(tag: str, **attrs) -> etree._Element:
    e = etree.Element(q(tag))
    for k, v in attrs.items():
        e.set(q(k), str(v))
    return e


def ptext(p) -> str:
    return ''.join(t.text or '' for t in p.iter(q('t')))


def _pstyle(p) -> str:
    ppr = p.find(q('pPr'))
    if ppr is None:
        return ''
    ps = ppr.find(q('pStyle'))
    return (ps.get(q('val')) or '') if ps is not None else ''


def _is_toc(p) -> bool:
    return _pstyle(p).upper().startswith('TOC')


def _is_p(k) -> bool:
    return etree.QName(k).localname == 'p'


# ── 锚点解析 ──────────────────────────────────────────────────────────────
def resolve_anchor(kids: list, op: dict) -> int:
    """返回锚段在 body 顶层的 index。唯一命中之外的一切情形都抛错。"""
    needle = op['find']
    start = 0
    within = op.get('within')
    if within:
        # within 通常写节标题('5.1'/'5.1 结论')：先按标题式匹配(段首缀+短段)，
        # 唯一命中才算；否则退回子串匹配。纯子串会被 '2015年2月'/'4.3.5.1' 这类误命中。
        head_hits = [i for i, k in enumerate(kids)
                     if _is_p(k) and not _is_toc(k)
                     and ptext(k).strip().startswith(within)
                     and len(ptext(k).strip()) <= 30]
        w_hits = head_hits if len(head_hits) == 1 else [
            i for i, k in enumerate(kids)
            if _is_p(k) and not _is_toc(k) and within in ptext(k)]
        if len(w_hits) != 1:
            raise ReviseError(
                f"✗ [{op.get('id', '?')}] within 锚点 {within!r} 命中 {len(w_hits)} 处"
                f"(要求唯一)：{[(i, ptext(kids[i])[:30]) for i in w_hits[:5]]}\n"
                f"   (标题式命中 {len(head_hits)} 处；写更完整的节标题文本可消歧)")
        start = w_hits[0] + 1
    hits = [i for i in range(start, len(kids))
            if _is_p(kids[i]) and not _is_toc(kids[i]) and needle in ptext(kids[i])]
    occ = op.get('occurrence')
    if occ:
        if not 1 <= occ <= len(hits):
            raise ReviseError(f"✗ [{op.get('id', '?')}] occurrence={occ} 超出命中数 {len(hits)}")
        return hits[occ - 1]
    if len(hits) != 1:
        detail = [(i, ptext(kids[i])[:36]) for i in hits[:6]]
        raise ReviseError(
            f"✗ [{op.get('id', '?')}] 锚点 {needle!r} 命中 {len(hits)} 处(要求唯一)：{detail}\n"
            f"   多命中 → 加 within: <节标题文本> 收窄；0 命中 → 源件已变，重新定位")
    return hits[0]


# ── 修订构件 ──────────────────────────────────────────────────────────────
def _sample_rpr(p):
    """取锚段第一个带 rPr 的 run(含 w:ins 内的)作新 run 格式模板——新段字体字号跟随邻段。"""
    for holder in [p] + p.findall(q('ins')):
        for r in holder.iter(q('r')):
            rpr = r.find(q('rPr'))
            if rpr is not None:
                return rpr
    return None


def _mk_run(text: str, rpr, bold: bool = False):
    r = mk('r')
    nr = copy.deepcopy(rpr) if rpr is not None else (mk('rPr') if bold else None)
    if nr is not None:
        if bold:
            for tag in ('bCs', 'b'):
                if nr.find(q(tag)) is None:
                    nr.insert(0, mk(tag))
        r.append(nr)
    t = mk('t')
    t.set(XML_SPACE, 'preserve')
    t.text = text
    r.append(t)
    return r


def build_ins_para(model_p, text: str, ids, meta: dict):
    """整段插入：pPr 照抄邻段，段落标记与全部 runs 都在修订标记里。**xx** 记加粗。"""
    p = mk('p')
    ppr = copy.deepcopy(model_p.find(q('pPr'))) if model_p.find(q('pPr')) is not None else mk('pPr')
    rpr_mark = ppr.find(q('rPr'))
    if rpr_mark is None:
        rpr_mark = mk('rPr')
        ppr.append(rpr_mark)
    rpr_mark.insert(0, mk('ins', id=next(ids), author=meta['author'], date=meta['date']))
    p.append(ppr)

    model_rpr = _sample_rpr(model_p)
    ins = mk('ins', id=next(ids), author=meta['author'], date=meta['date'])
    for seg in re.split(r'(\*\*[^*]+\*\*)', text):
        if not seg:
            continue
        bold = seg.startswith('**') and seg.endswith('**')
        ins.append(_mk_run(seg[2:-2] if bold else seg, model_rpr, bold=bold))
    p.append(ins)
    return p


def tracked_replace(p, op: dict, ids, meta: dict) -> None:
    """段内 old→new 修订对。old 须整段唯一(或 last:true 取末次)，且不能跨 run。"""
    old, new = op['old'], op['new']
    full = ptext(p)
    n = full.count(old)
    if n == 0 or (n > 1 and not op.get('last')):
        raise ReviseError(
            f"✗ [{op.get('id', '?')}] replace：{old!r} 在锚段出现 {n} 次"
            f"(要求恰 1 次，或加 last: true 取最后一次)")
    runs = [r for r in p.iter(q('r')) if r.find(q('t')) is not None]
    seq = runs if not op.get('last') else list(reversed(runs))
    for r in seq:
        t = r.find(q('t'))
        if old not in (t.text or ''):
            continue
        parent, pos = r.getparent(), list(r.getparent()).index(r)
        cut = t.text.rindex(old) if op.get('last') else t.text.index(old)
        head, tail = t.text[:cut], t.text[cut + len(old):]
        rpr = r.find(q('rPr'))
        if head:
            t.text = head
            pos += 1
        else:
            parent.remove(r)
        d = mk('del', id=next(ids), author=meta['author'], date=meta['date'])
        dr = mk('r')
        if rpr is not None:
            dr.append(copy.deepcopy(rpr))
        dt = mk('delText')
        dt.set(XML_SPACE, 'preserve')
        dt.text = old
        dr.append(dt)
        d.append(dr)
        ins = mk('ins', id=next(ids), author=meta['author'], date=meta['date'])
        ins.append(_mk_run(new, rpr))
        parent.insert(pos, ins)
        parent.insert(pos, d)
        if tail:
            tr = copy.deepcopy(r if head else dr)
            tt = tr.find(q('t')) if tr.find(q('t')) is not None else tr.find(q('delText'))
            tt.tag = q('t')
            tt.text = tail
            parent.insert(pos + 2, tr)
        return
    raise ReviseError(
        f"✗ [{op.get('id', '?')}] replace：{old!r} 跨 run 分布，本引擎 v1 不切跨 run 匹配——"
        f"改用更长的唯一锚文本，或先人工合并该段 runs")


_DEL_ANCESTORS = (q('del'), q('moveFrom'))


def _already_deleted(r) -> bool:
    """run 已在删除态里（w:del / w:moveFrom）—— 再包一层 w:del 会套娃并二次改 tag。"""
    return any(a.tag in _DEL_ANCESTORS for a in r.iterancestors())


def tracked_delete_runs(p, ids, meta: dict, include_nontext: bool = False) -> None:
    """整段旧内容标删除：w:t→w:delText 包进 w:del，段落标记标 w:del(replace_para tracked 用)。

    include_nontext=False（默认，= 2026-07-31 以来 replace_para 的历史行为，一个字节不变）：
      只标带 w:t 的 run；图片/公式/域这类没有 w:t 的 run 原样留下。
    include_nontext=True（`track compare` 整段删除用）：无 w:t 的 run 也包进 w:del。
      整段删除时不标它们 = 接受修订后「段落没了图还在」，rc 照样 0，属静默的错。
      w:del 里放 drawing 是 OOXML 的标准表示（CT_RunTrackChange 收 EG_ContentRunContent）。
    """
    for r in list(p.iter(q('r'))):
        if _already_deleted(r):
            continue
        t = r.find(q('t'))
        if t is None and not include_nontext:
            continue
        parent, pos = r.getparent(), list(r.getparent()).index(r)
        d = mk('del', id=next(ids), author=meta['author'], date=meta['date'])
        parent.remove(r)
        if t is not None:
            t.tag = q('delText')
        d.append(r)
        parent.insert(pos, d)
    ppr = p.find(q('pPr'))
    if ppr is None:
        ppr = mk('pPr')
        p.insert(0, ppr)
    rpr = ppr.find(q('rPr'))
    if rpr is None:
        rpr = mk('rPr')
        ppr.append(rpr)
    if rpr.find(q('del')) is None:   # 幂等：源件段落标记本就是删除态时不再叠一个
        rpr.insert(0, mk('del', id=next(ids), author=meta['author'], date=meta['date']))


def attach_comment(p, cid: int) -> None:
    start = mk('commentRangeStart', id=cid)
    p.insert(1 if p.find(q('pPr')) is not None else 0, start)
    p.append(mk('commentRangeEnd', id=cid))
    ref = mk('r')
    rpr = mk('rPr')
    st = mk('rStyle')
    st.set(q('val'), 'CommentReference')
    rpr.append(st)
    ref.append(rpr)
    ref.append(mk('commentReference', id=cid))
    p.append(ref)


# ── comments part：有则续，无则建 ─────────────────────────────────────────
class Comments:
    def __init__(self, z: zipfile.ZipFile, meta: dict):
        self.meta = meta
        if 'word/comments.xml' in z.namelist():
            self.root = etree.fromstring(z.read('word/comments.xml'))
            self.existed = True
        else:
            self.root = etree.Element(q('comments'), nsmap={'w': WNS})
            self.existed = False
        used = [int(c.get(q('id'))) for c in self.root.findall(q('comment'))
                if (c.get(q('id')) or '').isdigit()]
        self.next_id = max(used, default=0) + 1

    def add(self, text) -> int:
        cid = self.next_id
        self.next_id += 1
        c = mk('comment', id=cid, author=self.meta['author'], date=self.meta['date'],
               initials=self.meta.get('initials', ''))
        for para in ([text] if isinstance(text, str) else text):
            p = mk('p')
            r = mk('r')
            r.append(mk('annotationRef'))
            p.append(r)
            r2 = mk('r')
            t = mk('t')
            t.set(XML_SPACE, 'preserve')
            t.text = para
            r2.append(t)
            p.append(r2)
            c.append(p)
        self.root.append(c)
        return cid

    def tobytes(self) -> bytes:
        return etree.tostring(self.root, xml_declaration=True, encoding='UTF-8',
                              standalone=True)


def register_comments(ct: bytes, rels: bytes) -> tuple[bytes, bytes]:
    """[Content_Types].xml / document.xml.rels 注册 comments part(幂等)。"""
    cts = ct.decode('utf-8')
    if 'comments+xml' not in cts:
        ov = ('<Override PartName="/word/comments.xml" ContentType="application/'
              'vnd.openxmlformats-officedocument.wordprocessingml.comments+xml"/>')
        cts = cts.replace('</Types>', ov + '</Types>')
    rl = rels.decode('utf-8')
    if 'comments.xml' not in rl:
        nid = max((int(x) for x in re.findall(r'Id="rId(\d+)"', rl)), default=0) + 1
        rel = (f'<Relationship Id="rId{nid}" Type="http://schemas.openxmlformats.org/'
               'officeDocument/2006/relationships/comments" Target="comments.xml"/>')
        rl = rl.replace('</Relationships>', rel + '</Relationships>')
    return cts.encode('utf-8'), rl.encode('utf-8')


# ── 落位自检 ──────────────────────────────────────────────────────────────
def _heading_index(kids, token: str):
    for i, k in enumerate(kids):
        if _is_p(k) and not _is_toc(k):
            s = ptext(k).strip()
            if s.startswith(token) and len(s) <= 30:
                return i
    return None


def assert_between(kids, op: dict, pos: int) -> str:
    lo_t, hi_t = op['assert_between']
    lo, hi = _heading_index(kids, lo_t), _heading_index(kids, hi_t)
    if lo is None or hi is None or not (lo < pos < hi):
        raise ReviseError(
            f"✗ [{op.get('id', '?')}] 落位自检失败：落在 body[{pos}]，"
            f"应在 §{lo_t}(body[{lo}]) 与 §{hi_t}(body[{hi}]) 之间")
    return f'✓ 落位 [{op.get("id", "?")}] body[{pos}] ∈ §{lo_t}..§{hi_t}'


# ── 主流程 ────────────────────────────────────────────────────────────────
def apply_ops(src: Path, out: Path, cfg: dict, dry: bool = False) -> dict:
    src, out = Path(src), Path(out)
    if not src.exists():
        raise ReviseError(f'✗ 找不到源件：{src}')
    meta = {'author': cfg['author'], 'date': cfg['date'],
            'initials': cfg.get('initials', '')}
    z = zipfile.ZipFile(src)
    root = etree.fromstring(z.read('word/document.xml'))
    body = root.find(q('body'))

    used = [int(v) for e in root.iter() for v in [e.get(q('id'))]
            if v and v.isdigit()]
    ids = iter(range(max(used, default=0) + 1000, 10 ** 8))
    comments = Comments(z, meta)

    stats = {'ins_para': 0, 'replace': 0, 'replace_para': 0, 'comments': 0}
    lines: list[str] = []
    checks: list[tuple[dict, str]] = []   # (op, 首段文本) —— 全部 op 落完后统一落位自检

    for op in cfg['ops']:
        kids = list(body)
        i = resolve_anchor(kids, op)
        anchor_p = kids[i]
        action = op['action']
        lines.append(f"[{op.get('id', '?')}] {action} @ body[{i}] {ptext(anchor_p)[:32]!r}")

        if action == 'comment':
            attach_comment(anchor_p, comments.add(op['comment']))
            stats['comments'] += 1
            continue

        if action == 'replace':
            tracked_replace(anchor_p, op, ids, meta)
            stats['replace'] += 1
            if op.get('comment'):
                attach_comment(anchor_p, comments.add(op['comment']))
                stats['comments'] += 1
            continue

        if action in ('insert_after', 'replace_para'):
            # format_like：插入位置锚点(如表题)与格式模板段(如正文段)可以不是同一段——
            # 拿表题当格式模板会把新段套成题注样式(居中/黑体)，看着就错
            # 不继承 within：format_like 常常就是 within 锚本身(检索从它之后开始必然扑空)；
            # 要收窄时给单独的 format_within
            model_p = anchor_p
            if op.get('format_like'):
                model_p = kids[resolve_anchor(
                    kids, {'id': f"{op.get('id', '?')}/format_like",
                           'find': op['format_like'],
                           'within': op.get('format_within')})]
            new_ps = [build_ins_para(model_p, t, ids, meta) for t in op['paras']]
            at = i + 1
            if action == 'insert_after' and op.get('after_table'):
                # 锚段是表题时，越过紧邻的 w:tbl 再插——口径说明要跟在数据表后
                j = i + 1
                while j < len(kids) and etree.QName(kids[j]).localname != 'tbl':
                    j += 1
                at = j + 1 if j < len(kids) else i + 1
            for k, np_ in enumerate(new_ps):
                body.insert(at + k, np_)
            if action == 'replace_para':
                if op.get('tracked', True):
                    tracked_delete_runs(anchor_p, ids, meta)
                else:
                    body.remove(anchor_p)
                stats['replace_para'] += 1
            stats['ins_para'] += len(new_ps)
            if op.get('comment'):
                where = new_ps[-1] if op.get('comment_at') == 'last' else new_ps[0]
                attach_comment(where, comments.add(op['comment']))
                stats['comments'] += 1
            if op.get('assert_between'):
                checks.append((op, ptext(new_ps[0])[:16]))
            continue

        raise ReviseError(f"✗ [{op.get('id', '?')}] 未知 action: {action!r}")

    kids = list(body)
    for op, head in checks:
        pos = next((i for i, k in enumerate(kids) if _is_p(k) and head in ptext(k)), -1)
        lines.append(assert_between(kids, op, pos))

    if dry:
        z.close()
        return {**stats, 'dry_run': True, 'log': lines}

    override = {
        'word/document.xml': etree.tostring(root, xml_declaration=True,
                                            encoding='UTF-8', standalone=True),
        'word/comments.xml': comments.tobytes(),
        '[Content_Types].xml': None,
        'word/_rels/document.xml.rels': None,
    }
    ct, rels = register_comments(z.read('[Content_Types].xml'),
                                 z.read('word/_rels/document.xml.rels'))
    override['[Content_Types].xml'] = ct
    override['word/_rels/document.xml.rels'] = rels

    if out.exists():
        bak = out.with_suffix(f'.bak-{out.stat().st_mtime_ns}.docx')
        shutil.move(out, bak)
        lines.append(f'旧件已备份 → {bak.name}')
    out.parent.mkdir(parents=True, exist_ok=True)
    with zipfile.ZipFile(out, 'w', zipfile.ZIP_DEFLATED) as zo:
        wrote_comments = False
        for item in z.infolist():
            data = override.get(item.filename)
            zo.writestr(item, data if data is not None else z.read(item.filename))
            if item.filename == 'word/comments.xml':
                wrote_comments = True
        if not wrote_comments:
            zo.writestr('word/comments.xml', override['word/comments.xml'])
    z.close()

    # 白名单外任何部件变动 → 抛(丢 chart 事故的结构性防复发)；
    # 源件本无批注时 comments.xml 是本次有意新增，报备而非放宽
    added = frozenset() if comments.existed else frozenset({'word/comments.xml'})
    diff = assert_parts_intact(src, out, allow_added=added, verbose=False)
    lines.append(diff.summary())
    return {**stats, 'out': str(out), 'log': lines}
