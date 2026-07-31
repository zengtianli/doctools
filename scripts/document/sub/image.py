#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""image.py — image 家族三合一（2026-07-31 家族折叠）

子命令 ↔ 原脚本（函数体逐字搬移；模块级 main/apply_path 改名
main_<sub>/apply_path_<sub>，其余公有名一律保留）：

    dedup    ← image_dedup.py（media_hashes 公有名保留：slim.py 靠它）
    extract  ← image_extract.py（extract_images 公有名保留：pipeline_lib builtin step 靠它）
    relink   ← relink_images_from_source.py

各子命令 CLI 与原独立脚本逐字一致：python3 sub/image.py <sub> …。
退役原件在 ~/.Trash/consolidation-20260731/image/（含 MANIFEST.md）。
"""
from __future__ import annotations

import argparse
import hashlib
import json
import os
import re
import shutil
import sys
import zipfile
from copy import deepcopy
from datetime import date
from pathlib import Path
from xml.etree import ElementTree as ET

from lxml import etree

# 仓根 lib 进 sys.path —— 部件完整性断言（append 不是 insert(0)，防顶掉同名模块）
import sys as _sys
from pathlib import Path as _Path
_sys.path.append(str(_Path(__file__).resolve().parents[3] / 'lib'))
from docx_parts import assert_parts_intact  # noqa: E402

from docx import Document  # noqa: E402


# ══════════ dedup ← image_dedup.py ══════════

def media_hashes(docx_path: Path) -> dict:
    out = {}
    with zipfile.ZipFile(docx_path) as z:
        for name in z.namelist():
            if name.startswith("word/media/") and not name.endswith("/"):
                fn = name.split("/")[-1]
                if not fn:
                    continue
                data = z.read(name)
                h = hashlib.sha256(data).hexdigest()
                out[fn] = (h, len(data))
    return out


def rels_map(docx_path: Path) -> dict:
    out = {}
    with zipfile.ZipFile(docx_path) as z:
        rels_xml = z.read("word/_rels/document.xml.rels")
    root = ET.fromstring(rels_xml)
    for rel in root:
        if rel.tag.endswith("Relationship"):
            target = rel.get("Target", "")
            if "media/" in target:
                rid = rel.get("Id")
                out[rid] = target.split("/")[-1]
    return out


def image_positions(docx_path: Path) -> list:
    doc = Document(str(docx_path))
    positions = []
    heading_stack = {}
    for i, p in enumerate(doc.paragraphs):
        style = (p.style.name or "").strip()
        m = re.match(r"^(?:Heading|标题)\s*(\d+)$", style, re.IGNORECASE)
        if m and p.text.strip():
            lvl = int(m.group(1))
            heading_stack[lvl] = p.text.strip()
            for k in list(heading_stack):
                if k > lvl:
                    del heading_stack[k]
        xml = p._p.xml
        for embed in re.findall(r'r:embed="([^"]+)"', xml):
            chain = " > ".join(heading_stack[k] for k in sorted(heading_stack))
            positions.append((i, embed, chain or "（前言）"))
    return positions


def main_dedup():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--src", required=True, help="源 docx")
    ap.add_argument("--dst", required=True, help="目标 docx")
    ap.add_argument("--out", required=True, help="输出 MD 路径")
    args = ap.parse_args()

    SRC, DST = Path(args.src), Path(args.dst)
    src_hashes = media_hashes(SRC)
    dst_hashes = media_hashes(DST)
    print(f"源 {SRC.name}: {len(src_hashes)} 张", file=sys.stderr)
    print(f"目标 {DST.name}: {len(dst_hashes)} 张", file=sys.stderr)

    src_by_hash = {}
    for fn, (h, sz) in src_hashes.items():
        src_by_hash.setdefault(h, []).append((fn, sz))

    duplicates = {}
    for fn, (h, sz) in dst_hashes.items():
        if h in src_by_hash:
            duplicates[fn] = (h, sz, [n for n, _ in src_by_hash[h]])

    dst_rels = rels_map(DST)
    dst_pos = image_positions(DST)
    media_to_chains = {}
    for _, rid, chain in dst_pos:
        media = dst_rels.get(rid)
        if media:
            media_to_chains.setdefault(media, []).append(chain)

    lines = [f"# 图片重复清单 — {SRC.stem} → {DST.stem}\n",
             f"> 源 {len(src_hashes)} 张 / 目标 {len(dst_hashes)} 张 / **重复 {len(duplicates)} 张**\n",
             ""]
    if not duplicates:
        lines.append("✅ 无重复图片。\n")
    else:
        lines.append("## 二进制完全相同的图片\n")
        lines.append("| 目标文件 | 源文件 | KB | 出现章节 |")
        lines.append("|---------|--------|-----|---------|")
        for fn in sorted(duplicates, key=lambda x: int(re.sub(r"\D", "", x) or 0)):
            h, sz, src_fns = duplicates[fn]
            chains = media_to_chains.get(fn, ["（未在文档中引用）"])
            chains_str = "<br>".join(chains)
            lines.append(f"| {fn} | {', '.join(src_fns)} | {sz//1024} | {chains_str} |")

    Path(args.out).write_text("\n".join(lines), encoding="utf-8")
    print(f"OK -> {args.out}", file=sys.stderr)
    print(f"重复 {len(duplicates)} 张", file=sys.stderr)


# ---------------- pipeline adapter ----------------
def apply_path_dedup(docx_path, args=None) -> dict:
    """pipeline-compatible adapter (跨文件 analyzer).

    docx_path = dst (新版 / 待查重). args 透传:
      - src (必需): 源 docx
      - out / out_dir: 输出 MD 路径
    """
    from pathlib import Path as _P
    src_path = getattr(args, "src", None) if args else None
    if not src_path:
        return {"skipped": "no --src; image_dedup needs source docx"}
    SRC = _P(src_path)
    DST = _P(docx_path)
    out_path = getattr(args, "out", None)
    out_dir = getattr(args, "out_dir", None)
    if out_path:
        out = _P(out_path)
    elif out_dir:
        out = _P(out_dir) / f"image-dedup-{SRC.stem}-vs-{DST.stem}.md"
    else:
        out = DST.parent / "reports" / f"image-dedup-{SRC.stem}-vs-{DST.stem}.md"
    out.parent.mkdir(parents=True, exist_ok=True)

    src_hashes = media_hashes(SRC)
    dst_hashes = media_hashes(DST)
    src_by_hash = {}
    for fn, (h, sz) in src_hashes.items():
        src_by_hash.setdefault(h, []).append((fn, sz))
    duplicates = {}
    for fn, (h, sz) in dst_hashes.items():
        if h in src_by_hash:
            duplicates[fn] = (h, sz, [n for n, _ in src_by_hash[h]])
    dst_rels = rels_map(DST)
    dst_pos = image_positions(DST)
    media_to_chains = {}
    for _, rid, chain in dst_pos:
        media = dst_rels.get(rid)
        if media:
            media_to_chains.setdefault(media, []).append(chain)
    lines = [f"# 图片重复清单 — {SRC.stem} → {DST.stem}\n",
             f"> 源 {len(src_hashes)} / 目标 {len(dst_hashes)} / **重复 {len(duplicates)}**\n"]
    if not duplicates:
        lines.append("✅ 无重复图片。\n")
    else:
        lines.append("| 目标文件 | 源文件 | KB | 章节 |")
        lines.append("|---------|--------|-----|------|")
        for fn in sorted(duplicates, key=lambda x: int(re.sub(r"\D", "", x) or 0)):
            h, sz, src_fns = duplicates[fn]
            chains = media_to_chains.get(fn, ["（未引用）"])
            lines.append(f"| {fn} | {', '.join(src_fns)} | {sz//1024} | {'<br>'.join(chains)} |")
    out.write_text("\n".join(lines), encoding="utf-8")
    return {
        "src": str(SRC),
        "dst": str(DST),
        "src_images": len(src_hashes),
        "dst_images": len(dst_hashes),
        "duplicates": len(duplicates),
        "out": str(out),
    }


# ══════════ extract ← image_extract.py ══════════

W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
A_NS = "http://schemas.openxmlformats.org/drawingml/2006/main"
R_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
V_NS = "urn:schemas-microsoft-com:vml"
WP_NS = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
PIC_NS = "http://schemas.openxmlformats.org/drawingml/2006/picture"

NS_EXTRACT = {
    "w": W_NS,
    "a": A_NS,
    "r": R_NS,
    "v": V_NS,
    "wp": WP_NS,
    "pic": PIC_NS,
}

ILLEGAL_FS = re.compile(r'[\\/:*?"<>|\r\n\t]')
WS_RUN = re.compile(r"\s+")
CAPTION_PREFIX = re.compile(r"^图")


def _para_text(p: etree._Element) -> str:
    """Concatenate visible text in a <w:p>."""
    parts = []
    for t in p.iter(f"{{{W_NS}}}t"):
        if t.text:
            parts.append(t.text)
    return "".join(parts)


def _sanitize_stem(text: str, max_len: int = 100) -> str:
    s = ILLEGAL_FS.sub("_", text)
    s = WS_RUN.sub(" ", s).strip()
    s = s.rstrip(".")  # windows trailing dot
    if len(s) > max_len:
        s = s[:max_len].rstrip()
    return s


def _load_rels(z: zipfile.ZipFile) -> dict[str, str]:
    """Return rId -> target (e.g. media/image3.png) for word/_rels/document.xml.rels."""
    rels_path = "word/_rels/document.xml.rels"
    if rels_path not in z.namelist():
        return {}
    data = z.read(rels_path)
    root = etree.fromstring(data)
    out: dict[str, str] = {}
    for rel in root.findall("{http://schemas.openxmlformats.org/package/2006/relationships}Relationship"):
        rid = rel.get("Id")
        target = rel.get("Target")
        if rid and target:
            out[rid] = target
    return out


def _iter_image_anchors(body: etree._Element):
    """Yield (paragraph_element, rid) in document order for each embedded image.

    Walks <w:p> in body; for each paragraph, finds nested <a:blip r:embed=rId>
    and <v:imagedata r:id=rId>. Yields paragraph element + rid per image.
    """
    for p in body.iter(f"{{{W_NS}}}p"):
        # drawing-based images (modern)
        for blip in p.iter(f"{{{A_NS}}}blip"):
            rid = blip.get(f"{{{R_NS}}}embed") or blip.get(f"{{{R_NS}}}link")
            if rid:
                yield p, rid
        # legacy VML <v:imagedata>
        for vid in p.iter(f"{{{V_NS}}}imagedata"):
            rid = vid.get(f"{{{R_NS}}}id") or vid.get(f"{{{R_NS}}}href")
            if rid:
                yield p, rid


def _find_caption(body_paras: list[etree._Element], idx: int, lookahead: int = 2) -> str | None:
    """Look at body_paras[idx+1 .. idx+lookahead] for a paragraph starting with 图 and < 80 chars."""
    n = len(body_paras)
    for k in range(1, lookahead + 1):
        j = idx + k
        if j >= n:
            break
        txt = _para_text(body_paras[j]).strip()
        if not txt:
            continue
        if CAPTION_PREFIX.match(txt) and len(txt) < 80:
            return txt
    return None


def extract_images(docx_path: Path, out_dir: Path, quiet: bool = False) -> int:
    if not docx_path.exists():
        print(f"[image_extract] docx not found: {docx_path}", file=sys.stderr)
        return 2
    out_dir.mkdir(parents=True, exist_ok=True)
    with zipfile.ZipFile(docx_path, "r") as z:
        if "word/document.xml" not in z.namelist():
            print(f"[image_extract] word/document.xml missing in {docx_path}", file=sys.stderr)
            return 2
        rels = _load_rels(z)
        doc_xml = etree.fromstring(z.read("word/document.xml"))
        body = doc_xml.find(f"{{{W_NS}}}body")
        if body is None:
            print(f"[image_extract] no <w:body> in {docx_path}", file=sys.stderr)
            return 2

        body_paras = list(body.iter(f"{{{W_NS}}}p"))
        para_index = {id(p): i for i, p in enumerate(body_paras)}

        # Build image entries in physical order, dedupe by (paragraph_id, rid) tuple
        seen_keys: set[tuple[int, str]] = set()
        entries: list[tuple[etree._Element, str]] = []
        for p, rid in _iter_image_anchors(body):
            key = (id(p), rid)
            if key in seen_keys:
                continue
            seen_keys.add(key)
            entries.append((p, rid))

        if not entries:
            if not quiet:
                print(f"[image_extract] no images in {docx_path}")
            return 0

        used_names: set[str] = set()
        written: list[tuple[str, str]] = []
        for idx, (p, rid) in enumerate(entries, start=1):
            target = rels.get(rid)
            if not target:
                if not quiet:
                    print(f"[image_extract] WARN rId {rid} not in rels, skip", file=sys.stderr)
                continue
            # Resolve relative path from word/_rels/document.xml.rels
            # Most targets look like 'media/image3.png' -> word/media/image3.png
            if target.startswith("/"):
                zip_path = target.lstrip("/")
            else:
                zip_path = "word/" + target
            # normalize ../ if any
            zip_path = str(Path(zip_path)).replace("\\", "/")
            # collapse any leading word/.. patterns
            while "/../" in zip_path:
                head, _, tail = zip_path.partition("/../")
                head_parts = head.split("/")
                if head_parts:
                    head_parts.pop()
                zip_path = "/".join(head_parts + [tail])
            if zip_path not in z.namelist():
                if not quiet:
                    print(f"[image_extract] WARN media path missing: {zip_path}", file=sys.stderr)
                continue
            ext = Path(zip_path).suffix.lower() or ".bin"

            p_idx = para_index.get(id(p))
            caption = _find_caption(body_paras, p_idx) if p_idx is not None else None
            if caption:
                stem = _sanitize_stem(caption)
                if not stem:
                    stem = f"image-{idx:02d}"
            else:
                stem = f"image-{idx:02d}"

            # Dedupe filename within out_dir
            base = stem
            suffix_n = 1
            final_name = f"{base}{ext}"
            while final_name in used_names or (out_dir / final_name).exists():
                suffix_n += 1
                final_name = f"{base}-{suffix_n}{ext}"
            used_names.add(final_name)

            data = z.read(zip_path)
            (out_dir / final_name).write_bytes(data)
            written.append((final_name, zip_path))

        if not quiet:
            print(f"[image_extract] wrote {len(written)} image(s) to {out_dir}")
            for name, src in written:
                print(f"  {name}  <- {src}")
    return 0


def main_extract() -> int:
    ap = argparse.ArgumentParser(description=__doc__.splitlines()[0])
    ap.add_argument("docx", type=Path, help="source docx (read-only)")
    ap.add_argument("--out-dir", type=Path, required=True, help="output directory")
    ap.add_argument("--quiet", action="store_true")
    args = ap.parse_args()
    return extract_images(args.docx, args.out_dir, quiet=args.quiet)


# ══════════ relink ← relink_images_from_source.py ══════════

W_NS = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
A_NS = 'http://schemas.openxmlformats.org/drawingml/2006/main'
R_NS = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
WP_NS = 'http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing'
PIC_NS = 'http://schemas.openxmlformats.org/drawingml/2006/picture'
V_NS = 'urn:schemas-microsoft-com:vml'
MC_NS = 'http://schemas.openxmlformats.org/markup-compatibility/2006'
WPS_NS = 'http://schemas.microsoft.com/office/word/2010/wordprocessingShape'
CHART_NS = 'http://schemas.openxmlformats.org/drawingml/2006/chart'
PKG_REL_NS = 'http://schemas.openxmlformats.org/package/2006/relationships'
CT_NS = 'http://schemas.openxmlformats.org/package/2006/content-types'

NS = {
    'w': W_NS, 'a': A_NS, 'r': R_NS, 'wp': WP_NS, 'pic': PIC_NS,
    'v': V_NS, 'mc': MC_NS, 'wps': WPS_NS, 'c': CHART_NS,
}
IMAGE_REL_TYPE = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/image'

# ----- helpers -----
def qn(prefix, tag):
    return f'{{{NS[prefix]}}}{tag}'


FIG_PREFIX_RE = re.compile(r'^图\s*\d+\s*[-－]\s*\d+\s*')


def normalize_caption(text: str) -> str:
    """Strip 图 X-Y prefix and ideographic whitespace."""
    t = text.strip()
    t = FIG_PREFIX_RE.sub('', t)
    t = re.sub(r'^表\s*\d+\s*[-－]\s*\d+\s*', '', t)
    t = t.replace('　', '').replace(' ', '').strip()
    return t


def has_fig_prefix(text: str) -> bool:
    return bool(FIG_PREFIX_RE.match(text.strip()))


def canonical_caption(after: str, before: str) -> tuple[str, str]:
    """Pick canonical caption for an image: prefer side starting with '图X-Y'.
       Returns (canonical_text, source: 'after'|'before'|'none').
    """
    if has_fig_prefix(after):
        return after, 'after'
    if has_fig_prefix(before):
        return before, 'before'
    # 附图/无编号: pick the non-empty closer side (caller already truncated)
    if after:
        return after, 'after'
    if before:
        return before, 'before'
    return '', 'none'


def list_paragraphs(zf: zipfile.ZipFile):
    """Return list of (idx, w:p element) from word/document.xml."""
    xml = zf.read('word/document.xml')
    root = etree.fromstring(xml)
    body = root.find(qn('w', 'body'))
    paras = body.findall(qn('w', 'p'))
    return root, paras


def para_text(p) -> str:
    return ''.join(p.itertext()).strip()


def find_source_images(src_path: Path):
    """Scan source docx, return list of dicts:
       {idx, blip_rids:[...], caption_text, caption_idx, target_paths:[media/imageN.ext]}

       Strategy for canonical caption:
       - First collect all "图X-Y" prefix captions in document order.
       - Collect all blip-bearing paragraphs in document order (each blip = 1 slot).
       - Pair Nth fig-prefix caption ↔ Nth blip slot (deterministic, matches reading order).
       - For images without a matched prefix caption, fall back to nearest AFTER/BEFORE.
    """
    with zipfile.ZipFile(src_path) as z:
        root, paras = list_paragraphs(z)
        # rid -> target map
        rels_root = etree.fromstring(z.read('word/_rels/document.xml.rels'))
        rid_target = {}
        for r in rels_root:
            if r.get('Type') == IMAGE_REL_TYPE:
                rid_target[r.get('Id')] = r.get('Target')
        # media bytes
        media = {}
        for name in z.namelist():
            if name.startswith('word/media/'):
                media[name[len('word/'):]] = z.read(name)

    # Collect all 图X-Y prefix captions in document order
    fig_captions = []  # list of (para_idx, text)
    for idx, p in enumerate(paras):
        t = para_text(p)
        if has_fig_prefix(t):
            fig_captions.append((idx, t))

    # Collect all blip-bearing slots in document order
    blip_slots = []  # list of (src_idx, sub_i, rid)
    blip_para_indices = set()
    for idx, p in enumerate(paras):
        blips = p.findall('.//' + qn('a', 'blip'))
        embeds = [b.get(qn('r', 'embed')) for b in blips if b.get(qn('r', 'embed'))]
        for i, rid in enumerate(embeds):
            blip_slots.append((idx, i, rid))
        if embeds:
            blip_para_indices.add(idx)

    # For each caption, find nearest blip-bearing paragraph (within ±6 paras), prefer
    # the one IMMEDIATELY before (caption-after convention) then immediately after.
    # Each para can own multiple captions if it has multiple sub-blips, but for now we
    # map caption -> para_idx, then per-para multiple captions distribute by sub_i.
    para_to_captions = {}  # para_idx -> list of caption texts in document order
    for cap_idx, cap_text in fig_captions:
        # search nearest blip para
        best = None
        best_dist = 99
        for bp in blip_para_indices:
            # prefer caption immediately AFTER blip (bp < cap_idx with small distance)
            if bp < cap_idx:
                dist = cap_idx - bp
            else:
                dist = (bp - cap_idx) * 2 + 1  # penalize BEFORE-blip case
            if dist < best_dist:
                best = bp
                best_dist = dist
        if best is not None and best_dist <= 6:
            para_to_captions.setdefault(best, []).append(cap_text)

    # Build slot -> canonical caption: sub_i-th caption of the para
    slot_to_canonical = {}
    for src_idx, sub_i, rid in blip_slots:
        caps = para_to_captions.get(src_idx, [])
        if sub_i < len(caps):
            slot_to_canonical[(src_idx, sub_i)] = caps[sub_i]
        elif caps:
            # multiple blips share single caption — assign first to all
            slot_to_canonical[(src_idx, sub_i)] = caps[0]
        else:
            slot_to_canonical[(src_idx, sub_i)] = ''

    # Build per-paragraph output (1 entry per blip-bearing paragraph, listing all blips inside)
    out = []
    para_blips = {}
    for src_idx, sub_i, rid in blip_slots:
        para_blips.setdefault(src_idx, []).append((sub_i, rid))

    for idx in sorted(para_blips):
        p = paras[idx]
        sub_rids = para_blips[idx]
        embeds = [rid for _, rid in sub_rids]
        # caption AFTER
        cap_text = ''
        cap_idx = None
        for j in range(idx + 1, min(idx + 6, len(paras))):
            t = para_text(paras[j])
            if t:
                cap_text = t
                cap_idx = j
                break
        cap_before = ''
        for j in range(idx - 1, max(idx - 6, -1), -1):
            t = para_text(paras[j])
            if t:
                cap_before = t
                break
        # canonical per sub-blip
        sub_canonicals = []
        for sub_i, rid in sub_rids:
            canon = slot_to_canonical.get((idx, sub_i), '')
            if not canon:
                canon = cap_text or cap_before
            sub_canonicals.append(canon)
        out.append({
            'src_idx': idx,
            'blip_rids': embeds,
            'targets': [rid_target.get(rid) for rid in embeds],
            'caption_after': cap_text,
            'caption_after_idx': cap_idx,
            'caption_before': cap_before,
            'caption_after_norm': normalize_caption(cap_text),
            'caption_before_norm': normalize_caption(cap_before),
            'sub_canonicals': sub_canonicals,
            'sub_canonicals_norm': [normalize_caption(c) for c in sub_canonicals],
        })
    return out, media


def find_target_placeholders(tgt_path: Path):
    """Scan target docx, return list of dicts for paragraphs containing
       placeholder drawings/picts (wsp shapes WITHOUT a:blip):
       {tgt_idx, caption_after, caption_after_norm, shape_count}
    """
    with zipfile.ZipFile(tgt_path) as z:
        root, paras = list_paragraphs(z)
    out = []
    for idx, p in enumerate(paras):
        drawings = p.findall('.//' + qn('w', 'drawing'))
        picts = p.findall('.//' + qn('w', 'pict'))
        if not drawings and not picts:
            continue
        # skip chart-only (rId15/rId16 valid)
        blips = p.findall('.//' + qn('a', 'blip'))
        if blips:
            continue  # already has real image
        charts = p.findall('.//' + qn('c', 'chart'))
        # If only chart drawings (no wsp), skip
        wsps = p.findall('.//' + qn('wps', 'wsp'))
        v_rects = p.findall('.//' + qn('v', 'rect'))
        v_imgdata = p.findall('.//' + qn('v', 'imagedata'))
        if v_imgdata:
            continue  # already has v:imagedata
        if not (wsps or v_rects):
            continue  # nothing to relink (probably chart-only)
        # Count "shape slots". A single mc:AlternateContent block (Choice wsp + Fallback v:rect) = 1 shape.
        ac_blocks = p.findall('.//' + qn('mc', 'AlternateContent'))
        if ac_blocks:
            shape_count = len(ac_blocks)
        else:
            shape_count = max(len(wsps), len(v_rects), len(drawings))
        # caption AFTER
        cap_text = ''
        cap_idx = None
        for j in range(idx + 1, min(idx + 6, len(paras))):
            t = para_text(paras[j])
            if t:
                cap_text = t
                cap_idx = j
                break
        cap_before = ''
        for j in range(idx - 1, max(idx - 6, -1), -1):
            t = para_text(paras[j])
            if t:
                cap_before = t
                break
        out.append({
            'tgt_idx': idx,
            'shape_count': shape_count,
            'caption_after': cap_text,
            'caption_after_idx': cap_idx,
            'caption_before': cap_before,
            'caption_after_norm': normalize_caption(cap_text),
            'caption_before_norm': normalize_caption(cap_before),
        })
    return out


def _pair_score(tn: str, sn: str) -> int:
    if not tn or not sn:
        return 0
    if tn == sn:
        return 100
    if len(tn) >= 6 and tn in sn:
        return 80
    if len(sn) >= 6 and sn in tn:
        return 75
    set_t = set(tn)
    set_s = set(sn)
    if set_t and set_s:
        overlap = len(set_t & set_s) / max(len(set_t), len(set_s))
        if overlap >= 0.7:
            return 50
        if overlap >= 0.5:
            return 35
    return 0


def score_match(tgt: dict, src: dict) -> tuple[int, str]:
    """Heuristic match score 0-100 + reason. Source side picks canonical caption
       (the one with 图X-Y prefix). Target side compares against both BEFORE/AFTER
       (target captions are stripped, so we don't know which side).
    """
    tn_a = tgt['caption_after_norm']
    tn_b = tgt.get('caption_before_norm', '')
    src_canonical = src.get('caption_canonical_norm', '') or src['caption_after_norm']
    src_side = src.get('caption_canonical_side', 'after')

    # Score target's both sides against source's canonical.
    candidates = []
    if tn_a:
        candidates.append((_pair_score(tn_a, src_canonical), f'tgt-after~src-{src_side}', tn_a, src_canonical))
    if tn_b:
        candidates.append((_pair_score(tn_b, src_canonical), f'tgt-before~src-{src_side}', tn_b, src_canonical))
    if not candidates:
        return 0, 'empty target captions'
    candidates.sort(key=lambda x: x[0], reverse=True)
    score, mode, tn, sn = candidates[0]
    if score == 0:
        return 0, f'no-match tgt({tn_a[:20]!r}|{tn_b[:20]!r}) vs src({src_canonical[:25]!r})'
    return score, f'{mode}: {tn[:35]!r} ~ {sn[:35]!r}'


def build_patch(target_path: Path, source_path: Path) -> dict:
    """Build relink patch JSON."""
    src_images, src_media = find_source_images(source_path)
    tgt_placeholders = find_target_placeholders(target_path)

    # Flatten: each src image is one "blip slot"
    # src_images entries may have multiple blip_rids (e.g. 2 in one para)
    src_slots = []  # list of {src_idx, sub_i, rid, target_media, caption_after_norm, caption_after}
    for s in src_images:
        for i, (rid, mtarget) in enumerate(zip(s['blip_rids'], s['targets'])):
            canon = s['sub_canonicals'][i] if i < len(s['sub_canonicals']) else ''
            canon_norm = s['sub_canonicals_norm'][i] if i < len(s['sub_canonicals_norm']) else ''
            src_slots.append({
                'src_idx': s['src_idx'],
                'sub_i': i,
                'rid': rid,
                'target_media': mtarget,
                'caption_after': s['caption_after'],
                'caption_after_norm': s['caption_after_norm'],
                'caption_before_norm': s.get('caption_before_norm', ''),
                'caption_canonical': canon,
                'caption_canonical_norm': canon_norm,
                'caption_canonical_side': 'paired',
                'caption_after_idx': s['caption_after_idx'],
            })

    # Flatten target: each shape slot in placeholder = 1 slot.
    # For paragraphs with shape_count>1 (e.g. idx=116 has 2), we need 2 slots, but they
    # share the same caption_after. We will round-robin assign source slots ordered by
    # position to them.
    tgt_slots = []
    for t in tgt_placeholders:
        for k in range(t['shape_count']):
            tgt_slots.append({
                'tgt_idx': t['tgt_idx'],
                'sub_i': k,
                'shape_count': t['shape_count'],
                'caption_after': t['caption_after'],
                'caption_after_norm': t['caption_after_norm'],
                'caption_before_norm': t.get('caption_before_norm', ''),
                'caption_after_idx': t['caption_after_idx'],
            })

    # Match: pair each target slot to best source slot (by caption).
    # Strategy: for each target slot, find best-scoring source slot among UNUSED ones.
    # Sort target slots in document order; iterate and greedily pick best.
    used_src = set()
    matches = []
    unmatched = []
    for ti, t in enumerate(tgt_slots):
        best = (-1, None, '')
        for si, s in enumerate(src_slots):
            if si in used_src:
                continue
            score, reason = score_match(t, s)
            # tiebreaker: prefer source slot near same document order (sub_i)
            if s['sub_i'] == t['sub_i']:
                score += 1
            if score > best[0]:
                best = (score, si, reason)
        if best[0] >= 50 and best[1] is not None:
            si = best[1]
            used_src.add(si)
            s = src_slots[si]
            matches.append({
                'tgt_idx': t['tgt_idx'],
                'tgt_sub_i': t['sub_i'],
                'tgt_shape_count': t['shape_count'],
                'tgt_caption': t['caption_after'],
                'src_idx': s['src_idx'],
                'src_sub_i': s['sub_i'],
                'src_rid': s['rid'],
                'src_target_media': s['target_media'],
                'src_caption': s['caption_after'],
                'score': best[0],
                'reason': best[2],
            })
        else:
            unmatched.append({
                'tgt_idx': t['tgt_idx'],
                'tgt_sub_i': t['sub_i'],
                'tgt_caption': t['caption_after'],
                'best_score': best[0],
                'best_reason': best[2],
            })

    # images_to_copy = unique target_media used in matches
    used_media = {}
    for m in matches:
        tm = m['src_target_media']
        if tm and tm not in used_media:
            with zipfile.ZipFile(source_path) as z:
                size = z.getinfo('word/' + tm).file_size
            used_media[tm] = size
    images_to_copy = [{'source_path': 'word/' + k, 'target_name': k, 'size': v}
                      for k, v in used_media.items()]

    # rels_to_add: assign new rIds in target, one per unique media target
    # Read existing target rels max id
    with zipfile.ZipFile(target_path) as z:
        rels_xml = z.read('word/_rels/document.xml.rels')
    rels_root = etree.fromstring(rels_xml)
    existing_ids = [r.get('Id') for r in rels_root]
    existing_targets = {r.get('Target'): r.get('Id') for r in rels_root
                        if r.get('Type') == IMAGE_REL_TYPE}
    next_num = max([int(re.sub(r'\D', '', i) or '0') for i in existing_ids] + [0]) + 1
    media_to_newrid = {}
    rels_to_add = []
    for tm in used_media:
        if tm in existing_targets:
            media_to_newrid[tm] = existing_targets[tm]
        else:
            new_rid = f'rId{next_num}'
            next_num += 1
            media_to_newrid[tm] = new_rid
            rels_to_add.append({'rid': new_rid, 'type': 'image', 'target': tm})

    # drawing_rid_remap: per match, attach new_rid
    for m in matches:
        m['new_rid'] = media_to_newrid[m['src_target_media']]
        m['image_name'] = os.path.basename(m['src_target_media'])

    # extensions needed
    extensions = set()
    for tm in used_media:
        ext = os.path.splitext(tm)[1].lstrip('.').lower()
        if ext:
            extensions.add(ext)

    return {
        'target_docx': str(target_path),
        'source_docx': str(source_path),
        'summary': {
            'src_image_slots': len(src_slots),
            'tgt_placeholder_slots': len(tgt_slots),
            'matched': len(matches),
            'unmatched': len(unmatched),
            'images_to_copy': len(images_to_copy),
            'rels_to_add': len(rels_to_add),
        },
        'images_to_copy': images_to_copy,
        'rels_to_add': rels_to_add,
        'drawing_rid_remap': matches,
        'unmatched': unmatched,
        'extensions': sorted(extensions),
    }


# ---- apply ----
INLINE_DRAWING_XML = '''<w:drawing xmlns:w="{w}" xmlns:wp="{wp}" xmlns:a="{a}" xmlns:r="{r}" xmlns:pic="{pic}">
<wp:inline distT="0" distB="0" distL="0" distR="0">
<wp:extent cx="{cx}" cy="{cy}"/>
<wp:effectExtent l="0" t="0" r="0" b="0"/>
<wp:docPr id="{docpr_id}" name="{name}"/>
<wp:cNvGraphicFramePr><a:graphicFrameLocks noChangeAspect="1"/></wp:cNvGraphicFramePr>
<a:graphic><a:graphicData uri="{pic}">
<pic:pic><pic:nvPicPr><pic:cNvPr id="{docpr_id}" name="{name}"/><pic:cNvPicPr/></pic:nvPicPr>
<pic:blipFill><a:blip r:embed="{rid}"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill>
<pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="{cx}" cy="{cy}"/></a:xfrm>
<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr></pic:pic>
</a:graphicData></a:graphic></wp:inline></w:drawing>'''


def build_image_drawing(rid: str, name: str, cx: int, cy: int, doc_id: int):
    xml = INLINE_DRAWING_XML.format(
        w=W_NS, wp=WP_NS, a=A_NS, r=R_NS, pic=PIC_NS,
        cx=cx, cy=cy, rid=rid, name=name, docpr_id=doc_id,
    )
    return etree.fromstring(xml)


def get_shape_extent(shape_block) -> tuple[int, int]:
    """Try to get cx/cy from original placeholder shape; fallback 5400000 x 3000000."""
    ext = shape_block.find('.//' + qn('wp', 'extent'))
    if ext is not None:
        try:
            return int(ext.get('cx')), int(ext.get('cy'))
        except (TypeError, ValueError):
            pass
    return 5400000, 3000000


def apply_patch(target_path: Path, patch: dict, backup: bool = True) -> dict:
    """Apply patch to target docx (in-place by default with backup)."""
    if backup:
        n = 1
        while True:
            bak = target_path.with_name(
                f'{target_path.stem}.bak-{n}-{date.today().isoformat()}{target_path.suffix}'
            )
            if not bak.exists():
                break
            n += 1
        shutil.copy2(target_path, bak)
    else:
        bak = None

    # Load source media bytes once
    src_path = Path(patch['source_docx'])
    src_media_bytes = {}
    needed = {c['target_name'] for c in patch['images_to_copy']}
    with zipfile.ZipFile(src_path) as z:
        for tm in needed:
            src_media_bytes[tm] = z.read('word/' + tm)

    # Load target zip into memory, manipulate
    with zipfile.ZipFile(target_path, 'r') as z:
        contents = {name: z.read(name) for name in z.namelist()}

    # 1. Patch [Content_Types].xml
    ct_xml = contents['[Content_Types].xml']
    ct_root = etree.fromstring(ct_xml)
    existing_defaults = {d.get('Extension'): d.get('ContentType') for d in ct_root
                         if etree.QName(d).localname == 'Default'}
    ct_added = []
    ext_ct = {'png': 'image/png', 'jpeg': 'image/jpeg', 'jpg': 'image/jpeg',
              'gif': 'image/gif', 'bmp': 'image/bmp'}
    for ext in patch.get('extensions', []):
        if ext not in existing_defaults and ext in ext_ct:
            new_def = etree.SubElement(
                ct_root, etree.QName(CT_NS, 'Default'),
                Extension=ext, ContentType=ext_ct[ext]
            )
            ct_added.append(ext)
    contents['[Content_Types].xml'] = etree.tostring(
        ct_root, xml_declaration=True, encoding='UTF-8', standalone=True
    )

    # 2. Patch word/_rels/document.xml.rels
    rels_xml = contents['word/_rels/document.xml.rels']
    rels_root = etree.fromstring(rels_xml)
    existing_ids = {r.get('Id') for r in rels_root}
    rels_added = []
    for rel in patch.get('rels_to_add', []):
        if rel['rid'] in existing_ids:
            continue
        new_rel = etree.SubElement(
            rels_root, etree.QName(PKG_REL_NS, 'Relationship'),
            Id=rel['rid'], Type=IMAGE_REL_TYPE, Target=rel['target']
        )
        rels_added.append(rel['rid'])
    contents['word/_rels/document.xml.rels'] = etree.tostring(
        rels_root, xml_declaration=True, encoding='UTF-8', standalone=True
    )

    # 3. Patch word/document.xml — replace placeholder shapes
    doc_xml = contents['word/document.xml']
    doc_root = etree.fromstring(doc_xml)
    body = doc_root.find(qn('w', 'body'))
    paras = body.findall(qn('w', 'p'))

    # Group matches by tgt_idx
    by_para = {}
    for m in patch['drawing_rid_remap']:
        by_para.setdefault(m['tgt_idx'], []).append(m)

    replaced = 0
    doc_id_counter = 90000
    for idx, mlist in by_para.items():
        if idx >= len(paras):
            continue
        p = paras[idx]
        # Find all mc:AlternateContent blocks (each = 1 shape slot) in document order
        ac_blocks = p.findall('.//' + qn('mc', 'AlternateContent'))
        # If no AC blocks, look for bare w:drawing (containing wsp) or w:pict (containing v:rect)
        if not ac_blocks:
            # find drawings without blip
            shape_blocks = []
            for d in p.findall('.//' + qn('w', 'drawing')):
                if d.findall('.//' + qn('a', 'blip')):
                    continue
                if d.findall('.//' + qn('wps', 'wsp')):
                    shape_blocks.append(d)
            for pt in p.findall('.//' + qn('w', 'pict')):
                if pt.findall('.//' + qn('v', 'imagedata')):
                    continue
                if pt.findall('.//' + qn('v', 'rect')):
                    shape_blocks.append(pt)
            target_blocks = shape_blocks
        else:
            target_blocks = ac_blocks

        # Sort matches by tgt_sub_i
        mlist.sort(key=lambda x: x.get('tgt_sub_i', 0))
        for k, m in enumerate(mlist):
            if k >= len(target_blocks):
                break
            blk = target_blocks[k]
            cx, cy = get_shape_extent(blk)
            new_drawing = build_image_drawing(
                rid=m['new_rid'],
                name=m['image_name'],
                cx=cx, cy=cy,
                doc_id=doc_id_counter,
            )
            doc_id_counter += 1
            parent = blk.getparent()
            parent.replace(blk, new_drawing)
            replaced += 1

    contents['word/document.xml'] = etree.tostring(
        doc_root, xml_declaration=True, encoding='UTF-8', standalone=True
    )

    # 4. Add media files
    media_added = []
    for c in patch['images_to_copy']:
        key = 'word/' + c['target_name']
        contents[key] = src_media_bytes[c['target_name']]
        media_added.append(c['target_name'])

    # 5. Write zip
    tmp_path = target_path.with_suffix('.tmp.docx')
    with zipfile.ZipFile(tmp_path, 'w', zipfile.ZIP_DEFLATED) as z:
        for name, data in contents.items():
            z.writestr(name, data)
    # 部件完整性：replace 之前断言（此刻 target 未动 = 天然基线，断言炸则源件无损）。
    # 整包读进 dict 再重建是 137→35 事故同型的高危形态；有意新增 = 从 source 拷来的
    # media，改动的 document.xml / rels / [Content_Types].xml 在 DEFAULT_ALLOW_CHANGED。
    try:
        assert_parts_intact(
            target_path, tmp_path,
            allow_added={'word/' + c['target_name'] for c in patch['images_to_copy']},
            verbose=False)
    except Exception:
        tmp_path.unlink(missing_ok=True)
        raise
    os.replace(tmp_path, target_path)

    return {
        'backup': str(bak) if bak else None,
        'replaced_drawings': replaced,
        'media_added': media_added,
        'rels_added': rels_added,
        'content_type_extensions_added': ct_added,
    }


# ---- CLI ----
def main_relink():
    ap = argparse.ArgumentParser(description=__doc__.splitlines()[0])
    ap.add_argument('target_docx', type=Path)
    ap.add_argument('--source', type=Path, help='Source docx with real images')
    ap.add_argument('--apply-patch', type=Path, help='Skip detect, apply existing patch JSON')
    ap.add_argument('--dry-run', action='store_true', help='Detect only, no write')
    ap.add_argument('--no-backup', action='store_true', help='Skip .bak file')
    ap.add_argument('--report', type=Path, help='Write patch JSON here')
    args = ap.parse_args()

    if not args.target_docx.exists():
        print(f'ERROR: target {args.target_docx} not found', file=sys.stderr)
        sys.exit(2)

    if args.apply_patch:
        patch = json.loads(args.apply_patch.read_text(encoding='utf-8'))
        # doctools v1 schema 校验 (best-effort, 不阻断本地 ad-hoc patch)
        try:
            from lib.schemas import validate as _validate_schema
            _err = _validate_schema(patch, "patch")
            if _err:
                print(f'[warn] patch schema (v1) 校验未通过: {_err}', file=sys.stderr)
        except Exception:
            pass
    else:
        if not args.source or not args.source.exists():
            print(f'ERROR: --source required and must exist', file=sys.stderr)
            sys.exit(2)
        patch = build_patch(args.target_docx, args.source)

    if args.report:
        args.report.write_text(json.dumps(patch, ensure_ascii=False, indent=2),
                                encoding='utf-8')
        print(f'patch written: {args.report}')

    print(json.dumps(patch.get('summary', {}), ensure_ascii=False, indent=2))

    if args.dry_run:
        print('DRY RUN — no write.')
        return

    if not args.apply_patch and not patch.get('drawing_rid_remap'):
        print('No matches found — nothing to apply.')
        return

    result = apply_patch(args.target_docx, patch, backup=not args.no_backup)
    print(json.dumps(result, ensure_ascii=False, indent=2))


# ---------------- pipeline adapter ----------------
def apply_path_relink(docx_path, args=None) -> dict:
    """pipeline: 仅当 args.relink_source 提供时执行"""
    source = getattr(args, "relink_source", None) if args else None
    if not source:
        return {"changed": 0, "skipped": "no relink_source in args"}
    dry = bool(getattr(args, "dry_run", False)) if args else False
    target = Path(docx_path)
    patch = build_patch(target, Path(source))
    if dry or not patch.get("drawing_rid_remap"):
        return {"changed": 0, "patch_summary": patch.get("summary", {})}
    result = apply_patch(target, patch, backup=False)
    return {
        "changed": len(patch.get("drawing_rid_remap", {})),
        "result": result,
    }


# ──────────────────────────── 家族入口（子命令分发）────────────────────────────

SUBCOMMANDS = {
    "dedup": main_dedup,
    "extract": main_extract,
    "relink": main_relink,
}


def main(argv: list[str] | None = None) -> int:
    args = list(sys.argv[1:] if argv is None else argv)
    if not args or args[0] in ("-h", "--help"):
        print("usage: image.py {" + ",".join(SUBCOMMANDS) + "} <args…>\n"
              "每个子命令的参数与原独立脚本逐字一致：image.py <sub> --help 查看。")
        return 0 if args else 2
    sub, rest = args[0], args[1:]
    fn = SUBCOMMANDS.get(sub)
    if fn is None:
        print(f"[image] unknown subcommand: {sub!r}; choices={list(SUBCOMMANDS)}",
              file=sys.stderr)
        return 2
    saved = sys.argv[:]
    sys.argv = [sys.argv[0]] + rest
    try:
        rc = fn()
        return int(rc) if isinstance(rc, int) else 0
    finally:
        sys.argv = saved


if __name__ == "__main__":
    sys.exit(main())
