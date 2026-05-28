#!/usr/bin/env python3
"""image_extract.py — extract images from a docx, name files by neighboring caption.

Goal:
  Pull every embedded image out of a docx into a directory, naming each file by
  the nearest following "图x-y …" caption paragraph (if present). Fallback to
  `image-NN.<ext>` (NN = physical order) when no caption is detectable.

Why distilled (W · 2026-05-28):
  - 报告校审常需要逐张抽图, 旧办法 = 手动 unzip + 在 Word 里数 caption + rename, 慢且易错.
  - 既有 sub/image_dedup.py 处理"按 sha 去重", sub/audit_images.py 处理"审计",
    都不抽文件. 加一个 extract action 走 docx_cli image extract <docx> --out-dir <dir>.

CLI:
  python3 sub/image_extract.py <docx> --out-dir <dir> [--quiet]

Naming:
  - Caption probe: 取图节点所在 <w:p> 之后最近 1-2 段 <w:p>, 文本以 "图" 开头且 < 80 字符
    -> sanitize 后用作 file stem (剥除非法字符 \\/:*?"<>| -> _, 多空格压缩, trim, 截 100 字符)
  - Fallback: image-{idx:02d} (idx 按 docx 内首次出现顺序, 从 1 起)
  - 重名加 -2, -3 后缀
  - 扩展名沿用 word/media/imageN.<ext> 真扩展名 (.png/.jpeg/.wmf/.emf/...)

Not in scope:
  - 改 docx (纯读)
  - 提取 oleObject 嵌入文件 (只跟 drawing/pict <a:blip r:embed>)
  - caption 之外的命名启发 (alt text / 文件元数据)
"""
from __future__ import annotations

import argparse
import re
import sys
import zipfile
from pathlib import Path

from lxml import etree

W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
A_NS = "http://schemas.openxmlformats.org/drawingml/2006/main"
R_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
V_NS = "urn:schemas-microsoft-com:vml"
WP_NS = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
PIC_NS = "http://schemas.openxmlformats.org/drawingml/2006/picture"

NS = {
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


def main() -> int:
    ap = argparse.ArgumentParser(description=__doc__.splitlines()[0])
    ap.add_argument("docx", type=Path, help="source docx (read-only)")
    ap.add_argument("--out-dir", type=Path, required=True, help="output directory")
    ap.add_argument("--quiet", action="store_true")
    args = ap.parse_args()
    return extract_images(args.docx, args.out_dir, quiet=args.quiet)


if __name__ == "__main__":
    raise SystemExit(main())
