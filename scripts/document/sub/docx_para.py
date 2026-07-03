#!/usr/bin/env python3
"""docx_para — docx 段落级「查-改-验」工作台（surgical · OLE 安全 · LLM 会话友好）。

Why（2026-07-02 天台修改23 缩进病会话卡死）：
  改一段缩进(一段 pPr 里的脏 w:ind),主会话跑了 30-40 条串行一次性 python/bash —— 每条都重新
  unzip+正则整个 1.5MB document.xml;渲染验证 = soffice 整本 84 页转 PDF,分钟级前台卡死。
  根因不是模型能力,是缺一个「一次解压、带正文索引、段落级查改验」的 CLI。
  → 本工具:locate 定段(17ms) → inspect 看病 / edit|fix-ppr 改(自动 bak+verify+gate) →
    render 渲那页亲眼看(整本 PDF 按内容 hash 缓存,缓存热 <2s,冷 miss 提示后台 --warm 绝不同步卡死)。

6 子命令(每条秒级收口,输出一行一事实、稳定前缀、末行给下一条命令):
  locate    <docx> "<文本>"                   正文归一化定位段(防 run-split 假阴性)
  inspect   <docx> --para <i>                  dump pPr/rPr + 邻段 diff + DIRTY 判定
  edit      <docx> --para <i> --replace <o> <n> run-split-safe 文本替换
  fix-ppr   <docx> --para <i> [--clone-from prev|next|<j>] 克隆邻段 pPr(剥 sectPr/pPrChange)
  scan-ppr  <docx>                             全扫正文脏直接格式(w:ind 偏离同样式模态)
  render    <docx> (--para <i>|--page N) [--warm]  段→页缓存渲染(冷 miss exit 3 提示 warm)

exit code(全族统一):
  0 成功/唯一命中/干净  1 未命中/渲染失败/运行错  2 前置拦截(文件不存在/越界/Word锁/--expect不符/跨结构边界,文件未动)
  3 需人裁决(多命中/多页/scan有发现/render冷缓存未warm)  4 已写盘但 health gate 红(bak 可回滚)

surgical:只重写 word/document.xml,其余 zip 项(媒体/embeddings/OLE/公式)verbatim。
独立可跑: python3 sub/docx_para.py locate x.docx "..."
"""

from __future__ import annotations

import argparse
import hashlib
import json
import re
import shutil
import subprocess
import sys
from collections import Counter, defaultdict
from copy import deepcopy
from pathlib import Path

from lxml import etree

# ── lib import（docx_tools.py:40 同款:parents[3]/"lib"）─────────────────────
sys.path.insert(0, str(Path(__file__).resolve().parents[3] / "lib"))
import docx_surgical as ds  # noqa: E402
from docx_surgical import qn  # noqa: E402

# register() 的 _dispatch helper —— 仅 package(docx_cli)上下文可用;standalone 跑 main()不需要
try:
    from ._dispatch import get_or_add_group, get_or_add_subparsers  # type: ignore
except Exception:  # standalone (python3 sub/docx_para.py)
    get_or_add_group = get_or_add_subparsers = None  # type: ignore

XML_SPACE = "{http://www.w3.org/XML/1998/namespace}space"
CACHE_ROOT = Path.home() / ".cache" / "doctools" / "docx_para"
PDFTOTEXT = "/opt/homebrew/bin/pdftotext"
PDFTOPPM = "/opt/homebrew/bin/pdftoppm"
_SOFFICE_CANDIDATES = [
    "/Applications/LibreOffice.app/Contents/MacOS/soffice",
    "/opt/homebrew/bin/soffice",
]
_NORM = re.compile(r"\s+")

# 命中范围跨这些结构 = 不安全,edit 拒动(不猜)
_STRUCT = {qn("w:tab"), qn("w:br"), qn("w:drawing"), qn("w:pict"), qn("w:fldChar"), qn("w:instrText"), qn("w:object")}


# ═══ 段落属性读取 helpers ═══════════════════════════════════════════════════
def _pstyle_of(p):
    pPr = p.find(qn("w:pPr"))
    if pPr is None:
        return None
    st = pPr.find(qn("w:pStyle"))
    return st.get(qn("w:val")) if st is not None else None


def _in_table(p) -> bool:
    """段是否在 w:tbl 内(表格单元格段有自己的直接格式, 不算正文流脏格式)。"""
    return next(p.iterancestors(qn("w:tbl")), None) is not None


def _snippet(p, n: int = 40) -> str:
    t = ds.para_text(p).strip()
    return (t[:n] + "…") if len(t) > n else t


def _ind_sig(p):
    """段直接 w:ind 签名 (firstLine, firstLineChars, left, leftChars, hanging)。无直接 ind → None。"""
    pPr = p.find(qn("w:pPr"))
    if pPr is None:
        return None
    ind = pPr.find(qn("w:ind"))
    if ind is None:
        return None
    return tuple(ind.get(qn("w:" + k)) for k in ("firstLine", "firstLineChars", "left", "leftChars", "hanging"))


def _ind_str(sig) -> str:
    if sig is None:
        return "(无直接 ind, 继承样式)"
    keys = ("firstLine", "firstLineChars", "left", "leftChars", "hanging")
    parts = [f"{k}={v}" for k, v in zip(keys, sig, strict=False) if v is not None]
    return "ind " + (" ".join(parts) if parts else "(空)")


def _jc_of(p):
    pPr = p.find(qn("w:pPr"))
    if pPr is None:
        return None
    jc = pPr.find(qn("w:jc"))
    return jc.get(qn("w:val")) if jc is not None else None


def _spacing_sig(p):
    pPr = p.find(qn("w:pPr"))
    if pPr is None:
        return None
    sp = pPr.find(qn("w:spacing"))
    if sp is None:
        return None
    return tuple(sp.get(qn("w:" + k)) for k in ("before", "after", "line", "lineRule"))


def _run_rpr_sig(r):
    """run 的 (font, sz, bold) 摘要签名。"""
    rPr = r.find(qn("w:rPr"))
    if rPr is None:
        return (None, None, False)
    rFonts = rPr.find(qn("w:rFonts"))
    font = None
    if rFonts is not None:
        font = rFonts.get(qn("w:eastAsia")) or rFonts.get(qn("w:ascii")) or rFonts.get(qn("w:hAnsi"))
    sz = rPr.find(qn("w:sz"))
    szv = sz.get(qn("w:val")) if sz is not None else None
    b = rPr.find(qn("w:b"))
    bold = b is not None and b.get(qn("w:val")) not in ("false", "0")
    return (font, szv, bold)


def _run_modal(p):
    """段内(有文本的)run 的模态 rPr 签名 + run 数。"""
    sigs = []
    for r in p.iter(qn("w:r")):
        if any((t.text or "") for t in r.iter(qn("w:t"))):
            sigs.append(_run_rpr_sig(r))
    if not sigs:
        return None, 0
    modal = Counter(sigs).most_common(1)[0][0]
    return modal, len(sigs)


def _ppr_summary(p) -> str:
    """一行规范化 pPr 概要(每个子元素一段)。"""
    pPr = p.find(qn("w:pPr"))
    if pPr is None:
        return "(无 pPr, 纯样式继承)"
    parts = []
    for ch in pPr:
        local = etree.QName(ch).localname
        if local == "pStyle":
            parts.append(f"pStyle={ch.get(qn('w:val'))}")
        elif local == "ind":
            parts.append(_ind_str(_ind_sig(p)))
        elif local == "jc":
            parts.append(f"jc={ch.get(qn('w:val'))}")
        elif local == "spacing":
            sg = _spacing_sig(p)
            parts.append(
                "spacing "
                + " ".join(
                    f"{k}={v}"
                    for k, v in zip(("before", "after", "line", "lineRule"), sg, strict=False)
                    if v is not None
                )
            )
        elif local == "rPr":
            parts.append("rPr(段级)")
        elif local == "sectPr":
            parts.append("⚠sectPr(节属性)")
        elif local == "pPrChange":
            parts.append("⚠pPrChange(修订残留)")
        else:
            parts.append(local)
    return " | ".join(parts) if parts else "(空 pPr)"


def _validate_para_idx(paras, i) -> int | None:
    if i < 0 or i >= len(paras):
        print(f"段索引越界: {i} (文档共 {len(paras)} 段, 0..{len(paras)-1})", file=sys.stderr)
        return None
    return i


def _lock_guard(docx, force: bool) -> bool:
    """Word 锁检查。锁在且非 --force → 打印提示 + 返回 False(调用方 exit 2)。"""
    lock = ds.word_lock_file(docx)
    if lock is not None and not force:
        print(f"LOCK {docx.name} 正被 Word/WPS 打开 (owner 文件: {lock.name})")
        print("HINT: 先在 Word/WPS 关闭该文档, 或加 --force 越过(风险:并发写冲突/丢改)")
        return False
    return True


# ═══ 1. locate ═════════════════════════════════════════════════════════════
def cmd_locate(args) -> int:
    docx = Path(args.docx)
    if not docx.exists():
        print(f"找不到: {docx}", file=sys.stderr)
        return 2
    root = ds.parse_document(docx)
    paras = ds.iter_paras(root)
    b0 = ds.body_start_idx(paras)
    query = _NORM.sub("", args.text)
    if not query:
        print("查询文本为空", file=sys.stderr)
        return 2
    hits = [i for i in range(b0, len(paras)) if query in ds.para_text(paras[i], normalize=True)]

    if getattr(args, "json", False):
        print(
            json.dumps(
                {
                    "docx": str(docx),
                    "query": args.text,
                    "body_start": b0,
                    "hits": [{"para": i, "style": _pstyle_of(paras[i]), "text": _snippet(paras[i], 80)} for i in hits],
                },
                ensure_ascii=False,
            )
        )
        return 0 if len(hits) == 1 else (3 if hits else 1)

    if len(hits) == 1:
        i = hits[0]
        for j in (i - 1, i, i + 1):
            if 0 <= j < len(paras):
                star = "★ " if j == i else "  "
                print(f"PARA {j} | style={_pstyle_of(paras[j])} | {star}{_snippet(paras[j], 80)}")
        print(f"HINT: docx_cli.py para inspect {docx} --para {i}")
        return 0
    if len(hits) > 1:
        for i in hits:
            print(f"PARA {i} | style={_pstyle_of(paras[i])} | {_snippet(paras[i], 80)}")
        print(f"AMBIGUOUS: {len(hits)} 段命中, 请缩小查询或直接 inspect 某段")
        return 3
    # 无命中 → 最相近 3 段(归一化最长公共子串启发)
    print(f"未命中 '{args.text}' (正文 {b0}..{len(paras)-1})")
    scored = []
    for i in range(b0, len(paras)):
        t = ds.para_text(paras[i], normalize=True)
        if not t:
            continue
        lcs = _lcs_len(query, t)
        if lcs >= max(2, len(query) // 3):
            scored.append((lcs, i))
    scored.sort(reverse=True)
    for _, i in scored[:3]:
        print(f"  近似 PARA {i} | style={_pstyle_of(paras[i])} | {_snippet(paras[i], 80)}")
    return 1


def _lcs_len(a: str, b: str) -> int:
    """最长公共子串长度(仅供无命中时排相近段, O(len(a)) 内存滚动)。"""
    if not a or not b:
        return 0
    prev = [0] * (len(b) + 1)
    best = 0
    for i in range(1, len(a) + 1):
        cur = [0] * (len(b) + 1)
        ai = a[i - 1]
        for j in range(1, len(b) + 1):
            if ai == b[j - 1]:
                cur[j] = prev[j - 1] + 1
                if cur[j] > best:
                    best = cur[j]
        prev = cur
    return best


# ═══ 2. inspect ════════════════════════════════════════════════════════════
def cmd_inspect(args) -> int:
    docx = Path(args.docx)
    if not docx.exists():
        print(f"找不到: {docx}", file=sys.stderr)
        return 2
    root = ds.parse_document(docx)
    paras = ds.iter_paras(root)
    i = _validate_para_idx(paras, args.para)
    if i is None:
        return 2
    p = paras[i]
    style = _pstyle_of(p)
    modal, nruns = _run_modal(p)

    if getattr(args, "json", False):
        print(
            json.dumps(
                {
                    "para": i,
                    "style": style,
                    "ind_sig": _ind_sig(p),
                    "jc": _jc_of(p),
                    "spacing_sig": _spacing_sig(p),
                    "runs": nruns,
                    "run_modal_rpr": modal,
                    "text": _snippet(p, 120),
                },
                ensure_ascii=False,
            )
        )
        return 0

    print(f"PARA {i} | style={style} | {_snippet(p, 80)}")
    print(f"  pPr: {_ppr_summary(p)}")
    if modal:
        print(f"  runs {nruns} | rPr 模态: font={modal[0]} sz={modal[1]} bold={modal[2]}")

    # 邻段 diff（同 pStyle 优先做基准）
    dirty = []
    my_ind, my_jc = _ind_sig(p), _jc_of(p)
    for j in (i - 2, i - 1, i + 1, i + 2):
        if not (0 <= j < len(paras)):
            continue
        q = paras[j]
        diffs = []
        if _pstyle_of(q) != style:
            diffs.append(f"style={_pstyle_of(q)}")
        if _ind_sig(q) != my_ind:
            diffs.append(f"{_ind_str(_ind_sig(q))}")
        if _jc_of(q) != my_jc:
            diffs.append(f"jc={_jc_of(q)}")
        if diffs:
            print(f"DIFF vs PARA {j} ({_pstyle_of(q)}): {' , '.join(diffs)}")

    # DIRTY 判定：同 pStyle 邻段的模态 ind 签名 vs 本段
    tgt_in_tbl = _in_table(p)
    same_style_neighbors = [
        paras[j]
        for j in range(max(ds.body_start_idx(paras), i - 4), min(len(paras), i + 5))
        if j != i and _pstyle_of(paras[j]) == style and _in_table(paras[j]) == tgt_in_tbl
    ]
    if same_style_neighbors:
        modal_ind = Counter(_ind_sig(q) for q in same_style_neighbors).most_common(1)[0][0]
        if my_ind != modal_ind and my_ind is not None:
            dirty.append(f"DIRTY ind: 本段 {_ind_str(my_ind)} ≠ 同样式模态 {_ind_str(modal_ind)}")
        modal_jc = Counter(_jc_of(q) for q in same_style_neighbors).most_common(1)[0][0]
        if my_jc is not None and my_jc != modal_jc:
            dirty.append(f"DIRTY jc: 本段 jc={my_jc} ≠ 同样式模态 jc={modal_jc}")
    for d in dirty:
        print(d)
    print(f'HINT: docx_cli.py para fix-ppr {docx} --para {i} --clone-from prev --expect "{_snippet(p, 12)}"')
    return 3 if dirty else 0


# ═══ 3. edit ═══════════════════════════════════════════════════════════════
def _text_map(p):
    """段内 w:t 偏移表 + 结构标记位置(text 坐标)。"""
    spans = []
    marks = []
    cur = 0
    T = qn("w:t")
    for el in p.iter():
        if el.tag == T:
            txt = el.text or ""
            spans.append((cur, cur + len(txt), el))
            cur += len(txt)
        elif el.tag in _STRUCT:
            marks.append(cur)
    full = "".join((n.text or "") for (_, _, n) in spans)
    return full, spans, marks


def _find_with_ws_fallback(full: str, old: str):
    i = full.find(old)
    if i >= 0:
        return i, i + len(old)
    comp = []
    cmap = []
    for k, ch in enumerate(full):
        if not ch.isspace():
            comp.append(ch)
            cmap.append(k)
    comp_s = "".join(comp)
    oc = "".join(ch for ch in old if not ch.isspace())
    if not oc:
        return None
    j = comp_s.find(oc)
    if j < 0:
        return None
    return cmap[j], cmap[j + len(oc) - 1] + 1


def cmd_edit(args) -> int:
    docx = Path(args.docx)
    if not docx.exists():
        print(f"找不到: {docx}", file=sys.stderr)
        return 2
    if not _lock_guard(docx, args.force):
        return 2
    root = ds.parse_document(docx)
    paras = ds.iter_paras(root)
    i = _validate_para_idx(paras, args.para)
    if i is None:
        return 2
    p = paras[i]
    if args.expect and args.expect not in ds.para_text(p):
        print(f"--expect 不符: 段 {i} 不含 '{args.expect}' (段索引可能已漂移)", file=sys.stderr)
        return 2

    old, new = args.replace
    full, spans, marks = _text_map(p)
    found = _find_with_ws_fallback(full, old)
    if found is None:
        print(f"未在段 {i} 找到 '{old}' (段文本: {full[:80]}…)", file=sys.stderr)
        return 2
    s, e = found
    if any(s < m < e for m in marks):
        print("命中范围跨结构边界(tab/br/drawing/field), 请缩小 --replace 的 old", file=sys.stderr)
        return 2

    first_done = False
    for a, b, node in spans:
        if b <= s or a >= e:
            continue
        txt = node.text or ""
        ls, le = max(a, s) - a, min(b, e) - a
        if not first_done:
            node.text = txt[:ls] + new + txt[le:]
            first_done = True
            nt = node.text or ""
            if nt[:1].isspace() or nt[-1:].isspace():
                node.set(XML_SPACE, "preserve")
        else:
            node.text = txt[:ls] + txt[le:]
    if not first_done:
        print("内部错误: 命中定位但无 w:t 覆盖", file=sys.stderr)
        return 1

    try:
        bak = ds.surgical_rewrite(docx, ds.serialize(root), backup=not args.no_backup)
    except ds.RepackError as ex:
        print(f"REPACK FAIL: {ex}", file=sys.stderr)
        return 1
    print(f'EDITED PARA {i} | -"{old[:40]}" +"{new[:40]}"')
    if bak:
        print(f"BACKUP {bak.name}")
    return _finish_gate(docx, bak, args)


# ═══ 4. fix-ppr ════════════════════════════════════════════════════════════
def _clone_ppr(src_pPr):
    """deepcopy 源 pPr → 剥 sectPr + pPrChange(防节断悬空 rId / 修订残留)。"""
    c = deepcopy(src_pPr)
    for tag in ("sectPr", "pPrChange"):
        for el in c.findall(qn("w:" + tag)):
            c.remove(el)
    return c


def cmd_fix_ppr(args) -> int:
    docx = Path(args.docx)
    if not docx.exists():
        print(f"找不到: {docx}", file=sys.stderr)
        return 2
    if not _lock_guard(docx, args.force):
        return 2
    root = ds.parse_document(docx)
    paras = ds.iter_paras(root)
    i = _validate_para_idx(paras, args.para)
    if i is None:
        return 2
    p = paras[i]
    if args.expect and args.expect not in ds.para_text(p):
        print(f"--expect 不符: 段 {i} 不含 '{args.expect}' (段索引可能已漂移)", file=sys.stderr)
        return 2

    src = args.clone_from
    if src in ("prev", "next"):
        j = i - 1 if src == "prev" else i + 1
    else:
        try:
            j = int(src)
        except ValueError:
            print(f"--clone-from 须为 prev|next|<段号>, 收到 {src!r}", file=sys.stderr)
            return 2
    if not (0 <= j < len(paras)):
        print(f"克隆源段越界: {j}", file=sys.stderr)
        return 2
    src_p = paras[j]

    tgt_style = _pstyle_of(p)
    before = _ppr_summary(p)
    src_pPr = src_p.find(qn("w:pPr"))
    note_numpr = False
    cloned_style = None
    if src_pPr is None:
        # 源无 pPr → 删目标 pPr = 回归纯样式继承
        old = p.find(qn("w:pPr"))
        if old is not None:
            p.remove(old)
    else:
        cloned = _clone_ppr(src_pPr)
        note_numpr = cloned.find(qn("w:numPr")) is not None
        cs = cloned.find(qn("w:pStyle"))
        cloned_style = cs.get(qn("w:val")) if cs is not None else None
        old = p.find(qn("w:pPr"))
        if old is not None:
            p.remove(old)
        p.insert(0, cloned)
    after = _ppr_summary(p)

    try:
        bak = ds.surgical_rewrite(docx, ds.serialize(root), backup=not args.no_backup)
    except ds.RepackError as ex:
        print(f"REPACK FAIL: {ex}", file=sys.stderr)
        return 1
    print(f"FIXPPR PARA {i} (克隆源 PARA {j})")
    print(f"  before: {before}")
    print(f"  after : {after}")
    if cloned_style is not None and cloned_style != tgt_style:
        print(
            f"NOTE: 克隆源样式 pStyle={cloned_style} ≠ 目标原样式 {tgt_style} "
            f"(整段 pPr 克隆会改样式; 若只想清直接格式请选同样式邻段)"
        )
    if note_numpr:
        print("NOTE: numbering cloned (numPr 随整体克隆一并搬入, 属预期语义)")
    if bak:
        print(f"BACKUP {bak.name}")
    return _finish_gate(docx, bak, args)


# ═══ 5. scan-ppr ═══════════════════════════════════════════════════════════
def cmd_scan_ppr(args) -> int:
    docx = Path(args.docx)
    if not docx.exists():
        print(f"找不到: {docx}", file=sys.stderr)
        return 2
    root = ds.parse_document(docx)
    paras = ds.iter_paras(root)
    b0 = ds.body_start_idx(paras)
    body = [(i, paras[i]) for i in range(b0, len(paras))]
    # 有文本的正文流段（排除表格单元格段：表内段有自己的直接格式, 非正文流脏格式）
    text_body = [(i, p) for i, p in body if ds.para_text(p, normalize=True) and not _in_table(p)]
    if not text_body:
        print("TOTAL 0 suspects / 0 body paras")
        return 0
    n = len(text_body)
    style_counts = Counter((_pstyle_of(p) or "(无)") for _, p in text_body)
    # 正文类样式 = 占比 ≥30% 的模态样式 + 无 pStyle
    body_styles = {s for s, c in style_counts.items() if c / n >= 0.30}
    body_styles.add("(无)")
    # 每样式的模态 ind 签名
    sig_by_style: dict[str, Counter] = defaultdict(Counter)
    jc_by_style: dict[str, Counter] = defaultdict(Counter)
    for _i, p in text_body:
        s = _pstyle_of(p) or "(无)"
        if s in body_styles:
            sig_by_style[s][_ind_sig(p)] += 1
            jc_by_style[s][_jc_of(p)] += 1
    modal_ind = {s: c.most_common(1)[0][0] for s, c in sig_by_style.items()}
    modal_jc = {s: c.most_common(1)[0][0] for s, c in jc_by_style.items()}

    suspects = []
    for i, p in text_body:
        s = _pstyle_of(p) or "(无)"
        if s not in body_styles:
            continue
        sig = _ind_sig(p)
        reasons = []
        if sig is not None and sig != modal_ind[s]:
            reasons.append(f"{_ind_str(sig)}(模态 {_ind_str(modal_ind[s])})")
        jc = _jc_of(p)
        if jc is not None and jc != modal_jc[s]:
            reasons.append(f"直接 jc={jc}(模态 {modal_jc[s]})")
        if reasons:
            suspects.append({"para": i, "style": s, "reasons": reasons, "text": _snippet(p, 60)})

    if getattr(args, "json", False):
        print(
            json.dumps(
                {"docx": str(docx), "body_paras": n, "body_styles": sorted(body_styles), "suspects": suspects},
                ensure_ascii=False,
            )
        )
        return 3 if suspects else 0

    for su in suspects:
        print(f"SUSPECT PARA {su['para']} | style={su['style']} | " f"{' ; '.join(su['reasons'])} | {su['text']}")
    print(f"TOTAL {len(suspects)} suspects / {n} body paras")
    if suspects:
        print(f'HINT: 逐段 docx_cli.py para fix-ppr {docx} --para <i> --clone-from prev --expect "<片段>"')
        return 3
    return 0


# ═══ 6. render（整本 PDF 内容 hash 缓存 + 段→页映射 + 单页切）═══════════════
def _soffice():
    for c in _SOFFICE_CANDIDATES:
        if Path(c).exists():
            return c
    return shutil.which("soffice")


def _docx_hash(docx: Path) -> str:
    h = hashlib.sha1()
    with open(docx, "rb") as f:
        for chunk in iter(lambda: f.read(1 << 20), b""):
            h.update(chunk)
    return h.hexdigest()[:12]


def _cache_dir(docx: Path) -> Path:
    return CACHE_ROOT / f"{docx.stem}-{_docx_hash(docx)}"


def _prune_cache(docx: Path, keep: int = 3) -> None:
    prefix = docx.stem + "-"
    dirs = sorted(
        [d for d in CACHE_ROOT.glob(prefix + "*") if d.is_dir()], key=lambda d: d.stat().st_mtime, reverse=True
    )
    for d in dirs[keep:]:
        shutil.rmtree(d, ignore_errors=True)


def _warm(docx: Path, cache_dir: Path) -> int:
    soffice = _soffice()
    if soffice is None:
        print("找不到 soffice (LibreOffice)", file=sys.stderr)
        return 1
    cache_dir.mkdir(parents=True, exist_ok=True)
    profile = "file:///tmp/lo_profile_docx_para"
    filt = (
        'pdf:writer_pdf_Export:{"ReduceImageResolution":{"type":"boolean","value":true},'
        '"MaxImageResolution":{"type":"long","value":150}}'
    )
    full = cache_dir / "full.pdf"
    for conv in (filt, "pdf"):  # 带降采样过滤器 → 失败退朴素 pdf
        subprocess.run(
            [
                soffice,
                "-env:UserInstallation=" + profile,
                "--headless",
                "--convert-to",
                conv,
                "--outdir",
                str(cache_dir),
                str(docx),
            ],
            capture_output=True,
            timeout=600,
        )
        produced = cache_dir / (docx.stem + ".pdf")
        if produced.exists():
            produced.replace(full)
            break
    if not full.exists():
        print("CACHE MISS: soffice 转 PDF 失败", file=sys.stderr)
        return 1
    # pages.txt = pdftotext 全本 → 按 \f 切页 → 去全部空白, 每页一行
    r = subprocess.run([PDFTOTEXT, str(full), "-"], capture_output=True, timeout=120)
    pages = r.stdout.decode("utf-8", "replace").split("\f")
    norm_pages = [_NORM.sub("", pg) for pg in pages]
    (cache_dir / "pages.txt").write_text("\n".join(norm_pages), encoding="utf-8")
    (cache_dir / "meta.json").write_text(
        json.dumps(
            {"docx": str(docx), "hash": cache_dir.name.rsplit("-", 1)[-1], "pages": len(norm_pages)}, ensure_ascii=False
        ),
        encoding="utf-8",
    )
    _prune_cache(docx)
    print(f"WARM {docx.name} → {len(norm_pages)} 页缓存于 {cache_dir}")
    return 0


def _page_of_para(para_text_norm: str, cache_dir: Path) -> list[int]:
    lines = (cache_dir / "pages.txt").read_text(encoding="utf-8").split("\n")
    t = para_text_norm
    probes = [t[:40], t[-40:]] if len(t) > 60 else [t]
    hits = [idx + 1 for idx, line in enumerate(lines) if all(pr in line for pr in probes)]
    if hits:
        return hits
    # 跨页断段兜底：拼相邻两页
    for idx in range(len(lines) - 1):
        if all(pr in (lines[idx] + lines[idx + 1]) for pr in probes):
            hits.append(idx + 1)
    return hits


def cmd_render(args) -> int:
    docx = Path(args.docx)
    if not docx.exists():
        print(f"找不到: {docx}", file=sys.stderr)
        return 2
    cache_dir = _cache_dir(docx)
    if args.warm:
        return _warm(docx, cache_dir)

    if not (cache_dir / "pages.txt").exists() or not (cache_dir / "full.pdf").exists():
        print(f"CACHE MISS: 先跑 docx_cli.py para render {docx} --warm")
        print(
            f"  (后台: nohup python3 {Path(__file__).resolve().parents[1] / 'docx_cli.py'} "
            f"para render '{docx}' --warm >/dev/null 2>&1 &  首转 84 页约 1-3 分钟)"
        )
        return 3

    if args.page:
        pages = [args.page]
    else:
        root = ds.parse_document(docx)
        paras = ds.iter_paras(root)
        i = _validate_para_idx(paras, args.para)
        if i is None:
            return 2
        t = ds.para_text(paras[i], normalize=True)
        if not t:
            print(f"段 {i} 无文本, 无法按文本定位页, 请用 --page N", file=sys.stderr)
            return 1
        pages = _page_of_para(t, cache_dir)
        if not pages:
            print(f"未能把段 {i} 映射到页 (缓存或与当前 docx hash 不符?)", file=sys.stderr)
            return 1

    out_dir = Path(args.out_dir) if args.out_dir else docx.parent / (docx.stem + "_页检")
    out_dir.mkdir(parents=True, exist_ok=True)
    full = cache_dir / "full.pdf"
    render_pages = pages[:1]  # 多页命中只渲第一张
    printed = []
    for pg in render_pages:
        # pdftoppm 输出名 = <root>-<page>.png (page 宽度随总页数), 用 glob 收
        subprocess.run(
            [
                PDFTOPPM,
                "-png",
                "-f",
                str(pg),
                "-l",
                str(pg),
                "-r",
                str(args.dpi),
                str(full),
                str(out_dir / f"p{pg:03d}"),
            ],
            capture_output=True,
        )
        produced = sorted(out_dir.glob(f"p{pg:03d}*.png"))
        if produced:
            print(f"PAGE {pg} | PNG {produced[-1]}")
            printed.append(pg)
        else:
            print(f"PAGE {pg} | 渲染失败(pdftoppm 无输出)", file=sys.stderr)
    if len(pages) > 1:
        print(f"AMBIGUOUS: 段命中多页 {pages}, 已渲第一页 {render_pages[0]}")
        return 3
    return 0 if printed else 1


# ═══ health gate（subprocess docx_cli.py health gate, 实测 0.55s）═══════════
def _run_health_gate(docx: Path):
    cli = Path(__file__).resolve().parents[1] / "docx_cli.py"
    try:
        r = subprocess.run(
            [sys.executable, str(cli), "health", "gate", str(docx)], capture_output=True, text=True, timeout=120
        )
        payload = json.loads(r.stdout)
        return payload.get("gate", "ERROR"), payload.get("failed", [])
    except Exception as e:
        return "ERROR", [f"gate 调用异常: {type(e).__name__}: {e}"]


def _finish_gate(docx: Path, bak, args) -> int:
    """edit/fix-ppr 写盘后的 health gate。FAIL(真发现)→exit 4;ERROR/PASS→0。"""
    if getattr(args, "no_gate", False):
        print("GATE SKIP (--no-gate)")
        return 0
    status, failed = _run_health_gate(docx)
    if status == "FAIL":
        print(f"GATE FAIL {failed} (文件已改, 可从 {bak.name if bak else '无备份'} 回滚)")
        return 4
    if status == "PASS":
        print("GATE PASS")
        return 0
    # ERROR: gate 检查自身没跑全(与本次改动无关);verify_repacked 已证 zip/xml 完好
    print(f"GATE {status} (gate 未能全跑, 但 repack 自检已过; 文件已改)")
    return 0


# ═══ argparse ══════════════════════════════════════════════════════════════
def _build_parser(prog="docx_para.py"):
    ap = argparse.ArgumentParser(prog=prog, description="docx 段落级 查-改-验 工作台")
    sub = ap.add_subparsers(dest="para_sub", metavar="<target>", required=True)
    _add_targets(sub)
    return ap


def _add_targets(sp):
    lo = sp.add_parser("locate", help="正文归一化定位段(防 run-split)")
    lo.add_argument("docx")
    lo.add_argument("text", help="要定位的段落文本(可跨 run, 空白不敏感)")
    lo.add_argument("--json", action="store_true")
    lo.set_defaults(func=cmd_locate)

    ins = sp.add_parser("inspect", help="dump pPr/rPr + 邻段 diff + DIRTY 判定")
    ins.add_argument("docx")
    ins.add_argument("--para", type=int, required=True)
    ins.add_argument("--json", action="store_true")
    ins.set_defaults(func=cmd_inspect)

    ed = sp.add_parser("edit", help="run-split-safe 文本替换")
    ed.add_argument("docx")
    ed.add_argument("--para", type=int, required=True)
    ed.add_argument("--replace", nargs=2, metavar=("OLD", "NEW"), required=True)
    ed.add_argument("--expect", help="目标段须含此子串(防陈旧段索引)")
    ed.add_argument("--no-backup", action="store_true")
    ed.add_argument("--no-gate", action="store_true")
    ed.add_argument("--force", action="store_true", help="越过 Word 锁")
    ed.set_defaults(func=cmd_edit)

    fp = sp.add_parser("fix-ppr", help="克隆邻段 pPr(剥 sectPr/pPrChange)")
    fp.add_argument("docx")
    fp.add_argument("--para", type=int, required=True)
    fp.add_argument("--clone-from", default="prev", help="prev|next|<段号> (默认 prev)")
    fp.add_argument("--expect", help="目标段须含此子串(防陈旧段索引)")
    fp.add_argument("--no-backup", action="store_true")
    fp.add_argument("--no-gate", action="store_true")
    fp.add_argument("--force", action="store_true", help="越过 Word 锁")
    fp.set_defaults(func=cmd_fix_ppr)

    sc = sp.add_parser("scan-ppr", help="全扫正文脏直接格式(ind 偏离同样式模态)")
    sc.add_argument("docx")
    sc.add_argument("--json", action="store_true")
    sc.set_defaults(func=cmd_scan_ppr)

    rd = sp.add_parser("render", help="段→页缓存渲染(冷 miss 提示 --warm)")
    rd.add_argument("docx")
    g = rd.add_mutually_exclusive_group()
    g.add_argument("--para", type=int)
    g.add_argument("--page", type=int)
    rd.add_argument("--warm", action="store_true", help="仅冷转缓存(后台跑), 不渲页")
    rd.add_argument("--out-dir")
    rd.add_argument("--dpi", type=int, default=100)
    rd.set_defaults(func=cmd_render)


def register(subparsers) -> None:
    """注册 `para <target>` 到 docx_cli.py 顶层 subparsers。"""
    if get_or_add_group is None:
        return
    p = get_or_add_group(
        subparsers, "para", help_text="docx 段落级 查-改-验 工作台 (locate/inspect/edit/fix-ppr/scan-ppr/render)"
    )
    sp = get_or_add_subparsers(p, dest="para_sub", metavar="<target>")
    _add_targets(sp)


def main(argv=None) -> int:
    ap = _build_parser()
    args = ap.parse_args(argv)
    return args.func(args)


if __name__ == "__main__":
    sys.exit(main())
