#!/usr/bin/env python3
"""全文逐段 sequence diff: src.docx vs dst.docx

为 dst 每一段在 src 里找匹配（精确/高相似/中相似/无）。
段落覆盖：正文 + 所有表格 cell 段落。

Usage:
  python3 seqdiff.py --src OLD.docx --dst NEW.docx --out 对照.md
"""
import argparse
import re
import sys
from difflib import SequenceMatcher
from pathlib import Path

from docx import Document

EXACT = "原样照搬"
HIGH = "小改"   # >= 0.85
MID = "改写"    # 0.6-0.85
NEW = "新增"


def normalize(text: str) -> str:
    t = text.strip()
    t = re.sub(r"\s+", "", t)
    pairs = [
        ("\uff0c", ","), ("\u3002", "."), ("\uff1b", ";"), ("\uff1a", ":"),
        ("\uff08", "("), ("\uff09", ")"), ("\uff01", "!"), ("\uff1f", "?"),
        ("\u201c", '"'), ("\u201d", '"'), ("\u2018", "'"), ("\u2019", "'"),
        ("\u3001", ","),
    ]
    for a, b in pairs:
        t = t.replace(a, b)
    return t


def heading_level(p) -> int:
    name = (p.style.name or "").strip()
    m = re.match(r"^Heading\s+(\d+)$", name, re.IGNORECASE)
    if m:
        return int(m.group(1))
    m = re.match(r"^标题\s*(\d+)$", name)
    if m:
        return int(m.group(1))
    return 0


def iter_block_items(parent):
    from docx.document import Document as _Doc
    from docx.oxml.table import CT_Tbl
    from docx.oxml.text.paragraph import CT_P
    from docx.table import _Cell, Table
    from docx.text.paragraph import Paragraph

    if isinstance(parent, _Doc):
        parent_elm = parent.element.body
    elif isinstance(parent, _Cell):
        parent_elm = parent._tc
    else:
        parent_elm = parent

    for child in parent_elm.iterchildren():
        if isinstance(child, CT_P):
            yield Paragraph(child, parent)
        elif isinstance(child, CT_Tbl):
            yield Table(child, parent)


def extract_paragraphs(doc_path: Path):
    doc = Document(str(doc_path))
    seq = 0
    heading_stack = {}
    out = []

    def emit(text, is_table):
        nonlocal seq
        text = text.strip()
        if not text:
            return
        chain = " > ".join(heading_stack[k] for k in sorted(heading_stack)) or "(前言)"
        out.append((seq, chain, is_table, text, normalize(text)))
        seq += 1

    def walk_block(block, in_table=False):
        if hasattr(block, "rows"):
            for row in block.rows:
                for cell in row.cells:
                    for sub in iter_block_items(cell):
                        walk_block(sub, in_table=True)
        else:
            lvl = heading_level(block)
            text = block.text.strip()
            if lvl > 0 and text and not in_table:
                heading_stack[lvl] = text
                for k in list(heading_stack):
                    if k > lvl:
                        del heading_stack[k]
                emit(text, False)
            else:
                emit(text, in_table)

    for blk in iter_block_items(doc):
        walk_block(blk)
    return out


def best_match(target_norm, src_index, src_norms_list):
    if not target_norm:
        return (NEW, None, None)
    if target_norm in src_index:
        return (EXACT, 1.0, src_index[target_norm])
    target_len = len(target_norm)
    if target_len < 4:
        return (NEW, None, None)
    best_r = 0.0
    best_seq = None
    sm = SequenceMatcher(autojunk=False)
    sm.set_seq2(target_norm)
    for src_norm, src_seq in src_norms_list:
        if abs(len(src_norm) - target_len) > target_len * 0.6:
            continue
        sm.set_seq1(src_norm)
        r = sm.ratio()
        if r > best_r:
            best_r = r
            best_seq = src_seq
            if r >= 0.99:
                break
    if best_r >= 0.85:
        return (HIGH, best_r, best_seq)
    if best_r >= 0.6:
        return (MID, best_r, best_seq)
    return (NEW, best_r if best_r > 0 else None, best_seq)


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--src", required=True, help="源 docx（旧版）")
    ap.add_argument("--dst", required=True, help="目标 docx（新版，查重对象）")
    ap.add_argument("--out", required=True, help="输出 MD 路径")
    ap.add_argument("--noise-len", type=int, default=8, help="短于 N 字符的段视为表格字段噪声（默认 8）")
    args = ap.parse_args()

    src_paras = extract_paragraphs(Path(args.src))
    dst_paras = extract_paragraphs(Path(args.dst))
    print(f"源 {len(src_paras)} 段 / 目标 {len(dst_paras)} 段", file=sys.stderr)

    src_index, src_chain, src_text = {}, {}, {}
    src_norms_list = []
    for s, ch, _, raw, norm in src_paras:
        if norm not in src_index:
            src_index[norm] = s
        src_chain[s] = ch
        src_text[s] = raw
        src_norms_list.append((norm, s))

    rows = []
    counts = {EXACT: 0, HIGH: 0, MID: 0, NEW: 0}
    cross_chapter = []
    for seq, chain, is_table, raw, norm in dst_paras:
        status, ratio, src_seq = best_match(norm, src_index, src_norms_list)
        counts[status] += 1
        rows.append((seq, chain, is_table, raw, status, ratio, src_seq))
        if status in (EXACT, HIGH) and src_seq is not None:
            if src_chain.get(src_seq) != chain:
                cross_chapter.append((seq, chain, raw, status, src_seq, src_chain.get(src_seq, "?")))

    L = []
    L.append(f"# 逐段对照 — {Path(args.src).stem} vs {Path(args.dst).stem}\n")
    L.append("> **判定**：归一化精确=原样照搬；ratio≥0.85=小改；0.6≤ratio<0.85=改写；其余=新增\n")
    L.append("")
    L.append("## 总体统计\n")
    L.append("| 类别 | 段数 | 占比 |")
    L.append("|------|------|------|")
    total = len(dst_paras)
    for k in [EXACT, HIGH, MID, NEW]:
        L.append(f"| {k} | {counts[k]} | {counts[k]*100/total:.1f}% |")
    L.append(f"| **总段数** | **{total}** | 100% |")
    L.append(f"\n**雷同风险段** = 原样照搬 + 小改 = {counts[EXACT] + counts[HIGH]} 段（占 {(counts[EXACT]+counts[HIGH])*100/total:.1f}%）\n")

    if cross_chapter:
        L.append("\n## ⚠️ 跨章节迁移\n")
        L.append("| 新版段 | 新章节 | → 源章节 | 状态 | 原文摘要 |")
        L.append("|--------|--------|---------|------|----------|")
        for seq, chain, raw, status, src_seq, src_ch in cross_chapter[:50]:
            snippet = raw[:40].replace("|", "\\|")
            L.append(f"| {seq} | {chain[:30]} | {src_ch[:30]} | {status} | {snippet}… |")
        if len(cross_chapter) > 50:
            L.append(f"\n（另有 {len(cross_chapter)-50} 条略）")
        L.append("")

    def is_noise(raw):
        n = normalize(raw)
        return len(n) < args.noise_len or bool(re.match(r"^[\d\W]+$", n))

    L.append("\n## 雷同风险段落（按新版章节分组，已过滤短字段）\n")
    grouped, noise_count = {}, 0
    for seq, chain, is_table, raw, status, ratio, src_seq in rows:
        if status in (EXACT, HIGH):
            if is_noise(raw):
                noise_count += 1
                continue
            grouped.setdefault(chain, []).append((seq, raw, status, ratio, src_seq))
    L.append(f"> 噪声过滤 {noise_count} 段；实质雷同 {sum(len(v) for v in grouped.values())} 段\n")

    if not grouped:
        L.append("✅ 无雷同段落\n")
    else:
        for chain in sorted(grouped):
            items = grouped[chain]
            L.append(f"### 「{chain}」 — {len(items)} 段")
            for seq, raw, status, ratio, src_seq in items:
                src_ch = src_chain.get(src_seq, "?")
                same = "✓同章节" if src_ch == chain else f"✗→「{src_ch}」"
                rs = f"{ratio:.2f}" if ratio is not None else "-"
                L.append(f"- **#{seq}** [{status} ratio={rs} {same}]")
                L.append(f"  > {raw}")
                if status == HIGH and src_seq is not None:
                    L.append(f"  > **源原文(#{src_seq})**: {src_text.get(src_seq,'')}")
            L.append("")

    L.append("\n## 附录：完整明细\n")
    L.append("| # | 章节链 | T? | 状态 | ratio | 源# | 原文 |")
    L.append("|---|--------|----|------|-------|-----|------|")
    for seq, chain, is_table, raw, status, ratio, src_seq in rows:
        snippet = raw[:50].replace("|", "\\|").replace("\n", " ")
        rs = f"{ratio:.2f}" if ratio is not None else ""
        srcs = str(src_seq) if src_seq is not None else ""
        tbl = "T" if is_table else ""
        L.append(f"| {seq} | {chain[:25]} | {tbl} | {status} | {rs} | {srcs} | {snippet} |")

    Path(args.out).write_text("\n".join(L), encoding="utf-8")
    print(f"OK -> {args.out}", file=sys.stderr)
    print(f"照搬 {counts[EXACT]} | 小改 {counts[HIGH]} | 改写 {counts[MID]} | 新增 {counts[NEW]}", file=sys.stderr)


# ---------------- pipeline adapter ----------------
def apply_path(docx_path, args=None) -> dict:
    """pipeline-compatible adapter (跨文件 analyzer).

    docx_path = dst 新版 (查重对象). args 透传:
      - src (必需): 源 docx 旧版
      - noise_len: 噪声过滤短段阈值 (默认 8)
      - out / out_dir: 输出 MD 路径
    """
    from pathlib import Path as _P
    src_path = getattr(args, "src", None) if args else None
    if not src_path:
        return {"skipped": "no --src; seqdiff needs old-version docx"}
    src_path = _P(src_path)
    dst_path = _P(docx_path)
    noise_len = int(getattr(args, "noise_len", 8) or 8)
    out_path = getattr(args, "out", None)
    out_dir = getattr(args, "out_dir", None)
    if out_path:
        out = _P(out_path)
    elif out_dir:
        out = _P(out_dir) / f"seqdiff-{src_path.stem}-vs-{dst_path.stem}.md"
    else:
        out = dst_path.parent / "reports" / f"seqdiff-{src_path.stem}-vs-{dst_path.stem}.md"
    out.parent.mkdir(parents=True, exist_ok=True)

    src_paras = extract_paragraphs(src_path)
    dst_paras = extract_paragraphs(dst_path)
    src_index, src_chain, src_text = {}, {}, {}
    src_norms_list = []
    for s, ch, _, raw, norm in src_paras:
        if norm not in src_index:
            src_index[norm] = s
        src_chain[s] = ch
        src_text[s] = raw
        src_norms_list.append((norm, s))

    rows = []
    counts = {EXACT: 0, HIGH: 0, MID: 0, NEW: 0}
    for seq, chain, is_table, raw, norm in dst_paras:
        status, ratio, src_seq = best_match(norm, src_index, src_norms_list)
        counts[status] += 1
        rows.append((seq, chain, is_table, raw, status, ratio, src_seq))

    L = [f"# 逐段对照 — {src_path.stem} vs {dst_path.stem}\n"]
    total = len(dst_paras)
    L.append("## 总体统计\n")
    L.append("| 类别 | 段数 | 占比 |")
    L.append("|------|------|------|")
    for k in [EXACT, HIGH, MID, NEW]:
        pct = counts[k] * 100 / total if total else 0
        L.append(f"| {k} | {counts[k]} | {pct:.1f}% |")
    L.append(f"| **总段数** | **{total}** | 100% |")
    out.write_text("\n".join(L), encoding="utf-8")

    return {
        "src": str(src_path),
        "dst": str(dst_path),
        "src_paras": len(src_paras),
        "dst_paras": total,
        "exact": counts[EXACT],
        "high": counts[HIGH],
        "mid": counts[MID],
        "new": counts[NEW],
        "out": str(out),
    }


if __name__ == "__main__":
    main()
