#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
slim.py
=======

单功能描述
----------
**docx 瘦身 / one-shot slim** — 综合现有 strip + image-dedup 能力,
统一入口,两 mode:

* ``safe`` (默认): ensemble 串跑 strip_revisions / strip_bookmarks /
  strip_empty_captions / image_dedup (intra-doc) / strip_orphan_media,
  **保所有引用内容**, audit-styleset 不受影响, 典型 30-50% 瘦身.
* ``aggressive``: 套用 ``sub/extract_tables.py`` 的"最小 docx 骨架"
  构造范式 — 重建 docx 只保 ``[Content_Types].xml`` / ``_rels/.rels`` /
  ``word/document.xml`` / ``word/styles.xml`` / (按需) ``numbering.xml``
  / ``theme/theme1.xml`` + document.xml.rels 实际引用的 media. 砍掉
  header / footer / comments / footnotes / endnotes / customXml /
  fontTable / settings / webSettings / docProps (除 core.xml). 不可逆.

CLI
---
    python3 sub/slim.py <docx> [-o OUT] [--mode {safe,aggressive}]

如 ``-o`` 缺省, safe 模式覆写原文件 (保留 .bak), aggressive 模式拒绝原
位覆写 (避免误伤), 必须显式 ``-o``.

依赖复用 (不要 reimplement)
---------------------------
* ``sub.strip_revisions``    — body 内 ins/del/comment* 接受 + 清 trackChanges/comments.xml
* ``sub.strip_bookmarks``    — _Toc / _Ref / _Hlk / _GoBack 自动 bookmark
* ``sub.strip_empty_captions`` — caption 样式 strict-empty 段
* ``sub.image_dedup``        — ``media_hashes`` 复用做 intra-doc 同源图去重
* ``sub.strip_orphan_media`` — 删 word/media/* 未被任何 rId 引用的孤儿
"""
from __future__ import annotations

import argparse
import hashlib
import re
import shutil
import subprocess
import sys
import zipfile
from datetime import date
from pathlib import Path

try:
    from lxml import etree
    from docx import Document
except ImportError as e:
    print(f"[ERR] 缺依赖 (lxml / python-docx): {e}", file=sys.stderr)
    sys.exit(2)


# 复用现有 sub/*.py 函数 (silent on import, no top-level side effect)
from . import strip_revisions
from . import strip_bookmarks
from . import strip_empty_captions
from . import strip_orphan_media
from . import image_dedup  # media_hashes 函数复用


NS_REL = "http://schemas.openxmlformats.org/package/2006/relationships"
REL = f"{{{NS_REL}}}"
NS_W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
NS_R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"

_AGGRESSIVE_DROP_PATTERNS = (
    re.compile(r"^word/header\d*\.xml$"),
    re.compile(r"^word/footer\d*\.xml$"),
    re.compile(r"^word/_rels/header\d*\.xml\.rels$"),
    re.compile(r"^word/_rels/footer\d*\.xml\.rels$"),
    re.compile(r"^word/comments\.xml$"),
    re.compile(r"^word/_rels/comments\.xml\.rels$"),
    re.compile(r"^word/commentsExtended\.xml$"),
    re.compile(r"^word/commentsIds\.xml$"),
    re.compile(r"^word/footnotes\.xml$"),
    re.compile(r"^word/_rels/footnotes\.xml\.rels$"),
    re.compile(r"^word/endnotes\.xml$"),
    re.compile(r"^word/_rels/endnotes\.xml\.rels$"),
    re.compile(r"^word/fontTable\.xml$"),
    re.compile(r"^word/webSettings\.xml$"),
    re.compile(r"^word/settings\.xml$"),
    re.compile(r"^word/people\.xml$"),
    re.compile(r"^word/glossary/.*$"),
    re.compile(r"^customXml/.*$"),
    re.compile(r"^docProps/app\.xml$"),
    re.compile(r"^docProps/custom\.xml$"),
    re.compile(r"^docProps/thumbnail\..*$"),
)

# Aggressive keeps these (whitelist). All else under word/ except referenced media is dropped.
_AGGRESSIVE_KEEP_EXACT = {
    "[Content_Types].xml",
    "_rels/.rels",
    "word/document.xml",
    "word/_rels/document.xml.rels",
    "word/styles.xml",
    "word/numbering.xml",
    "word/theme/theme1.xml",
    "docProps/core.xml",
}


def _lsof_check(p: Path) -> str | None:
    try:
        r = subprocess.run(
            ["lsof", str(p)], capture_output=True, text=True, timeout=5
        )
        if r.returncode == 0 and r.stdout.strip():
            return r.stdout
    except (FileNotFoundError, subprocess.TimeoutExpired):
        pass
    return None


def _find_next_backup(docx_path: Path) -> Path:
    today = date.today().isoformat()
    n = 1
    while True:
        cand = docx_path.with_name(
            f"{docx_path.stem}.bak-{n}-{today}{docx_path.suffix}"
        )
        if not cand.exists():
            return cand
        n += 1


# ---------------- safe mode helpers ----------------

def _intra_doc_image_dedup(docx_path: Path) -> dict:
    """**单 docx 内**图片去重:
    使用 ``image_dedup.media_hashes`` 复用计算 SHA256 → 同 hash 多图保 1 张 +
    重写 ``word/_rels/document.xml.rels`` (+ header/footer rels) 让所有原引用 rId
    指向保留下来的那张. 删冗余 media 文件.

    Returns: {"duplicate_groups": N, "media_removed": K, "bytes_freed": B}
    """
    # 1. 算 SHA256 (复用 image_dedup)
    hash_map = image_dedup.media_hashes(docx_path)  # {filename: (sha256, size)}
    if not hash_map:
        return {"duplicate_groups": 0, "media_removed": 0, "bytes_freed": 0}

    # 2. 按 hash 聚 group
    by_hash: dict[str, list[str]] = {}
    for fn, (h, sz) in hash_map.items():
        by_hash.setdefault(h, []).append(fn)
    dup_groups = [(h, names) for h, names in by_hash.items() if len(names) > 1]
    if not dup_groups:
        return {"duplicate_groups": 0, "media_removed": 0, "bytes_freed": 0}

    # 3. 建 fname → keeper 映射 (字典序首个保留)
    fname_to_keeper: dict[str, str] = {}
    media_to_remove: set[str] = set()
    bytes_freed = 0
    for _h, names in dup_groups:
        names_sorted = sorted(names)
        keeper = names_sorted[0]
        for n in names_sorted[1:]:
            fname_to_keeper[n] = keeper
            media_to_remove.add(f"word/media/{n}")
            bytes_freed += hash_map[n][1]

    if not media_to_remove:
        return {"duplicate_groups": 0, "media_removed": 0, "bytes_freed": 0}

    # 4. 重写 zip: 删 dup media + rewrite 所有 rels Target 指向 keeper
    tmp = docx_path.with_suffix(docx_path.suffix + ".dedup.tmp")
    rels_re = re.compile(r"^word/_rels/.+\.xml\.rels$")
    parser = etree.XMLParser(remove_blank_text=False, recover=True)

    with zipfile.ZipFile(str(docx_path), "r") as zin:
        with zipfile.ZipFile(str(tmp), "w", zipfile.ZIP_DEFLATED) as zout:
            for it in zin.infolist():
                # 删冗余 media
                if it.filename in media_to_remove:
                    continue
                data = zin.read(it.filename)
                # 改 rels (Target 指向被合并的 fname → 改指 keeper)
                if rels_re.match(it.filename):
                    try:
                        root = etree.fromstring(data, parser=parser)
                    except etree.XMLSyntaxError:
                        root = None
                    if root is not None:
                        changed = False
                        for rel in root.findall(f"{REL}Relationship"):
                            tgt = (rel.get("Target") or "").replace("\\", "/")
                            tgt_norm = tgt.lstrip("/")
                            if "media/" in tgt_norm:
                                idx = tgt_norm.find("media/")
                                fname = tgt_norm[idx + len("media/"):]
                                if fname in fname_to_keeper:
                                    new_fname = fname_to_keeper[fname]
                                    prefix = tgt[: len(tgt) - len(fname)]
                                    rel.set("Target", prefix + new_fname)
                                    changed = True
                        if changed:
                            data = etree.tostring(
                                root, xml_declaration=True,
                                encoding="UTF-8", standalone=True,
                            )
                zout.writestr(it, data)

    shutil.move(str(tmp), str(docx_path))
    return {
        "duplicate_groups": len(dup_groups),
        "media_removed": len(media_to_remove),
        "bytes_freed": bytes_freed,
    }


def _call_strip_module(mod, docx_path: Path) -> dict:
    """Call mod.main() with argv = [docx, --no-backup] on working file.

    Returns {"ok": bool, "rc": int, "module": name}.
    """
    saved = sys.argv[:]
    sys.argv = [mod.__name__.split(".")[-1], str(docx_path), "--no-backup"]
    try:
        rc = mod.main() if hasattr(mod, "main") else 1
        rc = int(rc) if isinstance(rc, int) else 0
    except SystemExit as se:
        rc = int(se.code) if isinstance(se.code, int) else (0 if se.code is None else 1)
    except Exception as e:
        return {"ok": False, "rc": -1, "module": mod.__name__,
                "error": f"{type(e).__name__}: {e}"}
    finally:
        sys.argv = saved
    return {"ok": rc == 0, "rc": rc, "module": mod.__name__.split(".")[-1]}


def run_safe(src_docx: Path, out_path: Path | None) -> dict:
    """Safe ensemble: strip-revisions + strip-bookmarks + strip-empty-captions +
    intra-doc image-dedup + strip-orphan-media. Preserves all referenced content.

    If out_path is None, rewrites src in place (with .bak); else writes to out.
    Returns report dict.
    """
    src_docx = Path(src_docx).expanduser().resolve()
    if not src_docx.exists():
        raise FileNotFoundError(f"input docx not found: {src_docx}")

    size_before = src_docx.stat().st_size

    # Decide working file
    inplace = out_path is None
    if inplace:
        occ = _lsof_check(src_docx)
        if occ:
            raise RuntimeError(f"文件被占用 (Word/WPS): {occ}")
        bak = _find_next_backup(src_docx)
        shutil.copy2(src_docx, bak)
        work = src_docx
        backup = bak
    else:
        out_path = Path(out_path).expanduser().resolve()
        out_path.parent.mkdir(parents=True, exist_ok=True)
        shutil.copy2(src_docx, out_path)
        work = out_path
        backup = None

    steps_report = []

    # 1. revisions
    steps_report.append(("strip-revisions", _call_strip_module(strip_revisions, work)))
    # 2. bookmarks
    steps_report.append(("strip-bookmarks", _call_strip_module(strip_bookmarks, work)))
    # 3. empty-captions
    steps_report.append(("strip-empty-captions", _call_strip_module(strip_empty_captions, work)))
    # 4. intra-doc image-dedup (reuse media_hashes)
    try:
        dedup_rep = _intra_doc_image_dedup(work)
        steps_report.append(("image-dedup", {"ok": True, **dedup_rep}))
    except Exception as e:
        steps_report.append(("image-dedup", {"ok": False, "error": f"{type(e).__name__}: {e}"}))
    # 5. orphan-media (dedup 后留下的 rels-orphan + 原本就 orphan 一锅烩)
    steps_report.append(("strip-orphan-media", _call_strip_module(strip_orphan_media, work)))

    size_after = work.stat().st_size
    return {
        "mode": "safe",
        "input": str(src_docx),
        "output": str(work),
        "backup": str(backup) if backup else None,
        "size_before": size_before,
        "size_after": size_after,
        "delta_bytes": size_before - size_after,
        "delta_pct": round((size_before - size_after) / size_before * 100, 2) if size_before else 0,
        "steps": steps_report,
    }


# ---------------- aggressive mode ----------------

def _collect_referenced_media_from_doc(z: zipfile.ZipFile) -> tuple[set[str], set[str]]:
    """Aggressive mode media filter: keep only media whose rId is **actually used**
    inside ``word/document.xml`` body (r:embed / r:link / r:id attrs).

    Returns (kept_media_zip_paths, used_rids_in_body).
    """
    parser = etree.XMLParser(remove_blank_text=False, recover=True)
    # 1. Scan document.xml for actually used rIds
    used_rids: set[str] = set()
    try:
        doc_bytes = z.read("word/document.xml")
        doc_root = etree.fromstring(doc_bytes, parser=parser)
    except (KeyError, etree.XMLSyntaxError):
        doc_root = None
    if doc_root is not None:
        # Only count rIds used by image/blip references (r:embed/r:link).
        # r:id is also used by headerReference/footerReference/footnoteReference
        # etc which aggressive drops anyway — counting those would falsely keep
        # their media. Aggressive also strips header/footer/footnote refs from
        # document.xml before this point so this is conservative.
        attrs = (f"{{{NS_R}}}embed", f"{{{NS_R}}}link")
        for elem in doc_root.iter():
            for a in attrs:
                v = elem.get(a)
                if v:
                    used_rids.add(v)

    # 2. Map rIds → media targets via document.xml.rels
    referenced: set[str] = set()
    try:
        rels_data = z.read("word/_rels/document.xml.rels")
        rels_root = etree.fromstring(rels_data, parser=parser)
    except (KeyError, etree.XMLSyntaxError):
        rels_root = None
    if rels_root is not None:
        for rel in rels_root.findall(f"{REL}Relationship"):
            rid = rel.get("Id")
            if rid not in used_rids:
                continue
            tgt = (rel.get("Target") or "").replace("\\", "/").lstrip("/")
            if tgt.startswith("media/"):
                referenced.add("word/" + tgt)
            elif "/media/" in tgt:
                idx = tgt.find("media/")
                referenced.add("word/" + tgt[idx:])
    return referenced, used_rids


def _strip_doc_rels_for_aggressive(
    rels_xml: bytes, kept_media: set[str], used_rids: set[str]
) -> bytes:
    """Filter word/_rels/document.xml.rels: keep ONLY (styles/numbering/theme) +
    rels whose Id is actually used in document.xml body AND points to kept media.
    Drop hyperlinks/footnotes/endnotes/comments/header/footer/customXml/embeddings
    etc.
    """
    parser = etree.XMLParser(remove_blank_text=False, recover=True)
    try:
        root = etree.fromstring(rels_xml, parser=parser)
    except etree.XMLSyntaxError:
        return rels_xml
    if root is None:
        return rels_xml

    structural_types = {
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles",
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering",
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme",
    }
    body_used_types = {
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image",
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/diagramData",
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart",
    }
    image_type = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image"

    for rel in list(root.findall(f"{REL}Relationship")):
        rtype = rel.get("Type") or ""
        rid = rel.get("Id")
        if rtype in structural_types:
            continue  # always keep styles/numbering/theme
        if rtype in body_used_types:
            if rid not in used_rids:
                root.remove(rel)
                continue
            if rtype == image_type:
                tgt = (rel.get("Target") or "").replace("\\", "/").lstrip("/")
                if tgt.startswith("media/"):
                    zp = "word/" + tgt
                elif "/media/" in tgt:
                    idx = tgt.find("media/")
                    zp = "word/" + tgt[idx:]
                else:
                    root.remove(rel)
                    continue
                if zp not in kept_media:
                    root.remove(rel)
            continue
        # default: drop everything else
        root.remove(rel)
    return etree.tostring(
        root, xml_declaration=True, encoding="UTF-8", standalone=True
    )


def _strip_document_xml_for_aggressive(doc_xml: bytes) -> bytes:
    """Strip references to dropped parts inside word/document.xml:
    - <w:sectPr> children <w:headerReference> / <w:footerReference> (header/footer dropped)
    - footnote/endnote references inside runs

    Returns rewritten bytes; if no change, returns input.
    """
    parser = etree.XMLParser(remove_blank_text=False, recover=True)
    try:
        root = etree.fromstring(doc_xml, parser=parser)
    except etree.XMLSyntaxError:
        return doc_xml
    if root is None:
        return doc_xml
    W = f"{{{NS_W}}}"
    changed = False
    # Remove header/footer references
    for tag in (f"{W}headerReference", f"{W}footerReference"):
        for el in list(root.iter(tag)):
            p = el.getparent()
            if p is not None:
                p.remove(el)
                changed = True
    # Remove footnote/endnote references
    for tag in (f"{W}footnoteReference", f"{W}endnoteReference"):
        for el in list(root.iter(tag)):
            p = el.getparent()
            if p is not None:
                p.remove(el)
                changed = True
    if not changed:
        return doc_xml
    return b'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' + etree.tostring(
        root, xml_declaration=False, encoding="utf-8"
    )


def _build_aggressive_content_types(has_numbering: bool, has_theme: bool, kept_media: set[str]) -> bytes:
    """Build minimal [Content_Types].xml: Default for rels/xml/media extensions +
    Override for document/styles/numbering/theme/core."""
    # Collect media extensions
    exts = set()
    for mp in kept_media:
        ext = mp.rsplit(".", 1)[-1].lower() if "." in mp else ""
        if ext:
            exts.add(ext)
    # Common content-type map
    ct_map = {
        "png": "image/png",
        "jpg": "image/jpeg",
        "jpeg": "image/jpeg",
        "gif": "image/gif",
        "bmp": "image/bmp",
        "tiff": "image/tiff",
        "tif": "image/tiff",
        "wmf": "image/x-wmf",
        "emf": "image/x-emf",
        "svg": "image/svg+xml",
        "webp": "image/webp",
        "ico": "image/x-icon",
    }
    parts = [
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
        '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">',
        '  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>',
        '  <Default Extension="xml" ContentType="application/xml"/>',
    ]
    for ext in sorted(exts):
        ct = ct_map.get(ext, f"image/{ext}")
        parts.append(f'  <Default Extension="{ext}" ContentType="{ct}"/>')
    parts.append('  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>')
    parts.append('  <Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>')
    if has_numbering:
        parts.append('  <Override PartName="/word/numbering.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml"/>')
    if has_theme:
        parts.append('  <Override PartName="/word/theme/theme1.xml" ContentType="application/vnd.openxmlformats-officedocument.theme+xml"/>')
    parts.append('  <Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>')
    parts.append('</Types>')
    return ("\n".join(parts) + "\n").encode("utf-8")


def _build_aggressive_root_rels() -> bytes:
    return (
        b'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
        b'<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">\n'
        b'  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>\n'
        b'  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>\n'
        b'</Relationships>\n'
    )


def run_aggressive(src_docx: Path, out_path: Path) -> dict:
    """Aggressive minimal-skeleton rebuild.

    Refuses to overwrite source in-place (must pass explicit -o).
    """
    src_docx = Path(src_docx).expanduser().resolve()
    out_path = Path(out_path).expanduser().resolve()
    if not src_docx.exists():
        raise FileNotFoundError(f"input docx not found: {src_docx}")
    if src_docx == out_path:
        raise ValueError("aggressive 模式拒绝原位覆写, 请指定 -o 到新路径")

    # WARN to stderr
    print(
        "WARN: aggressive 模式将丢失 header/footer/comments/footnotes/endnotes/customXml, "
        "不可逆。如需保留这些用 --mode safe",
        file=sys.stderr,
    )

    out_path.parent.mkdir(parents=True, exist_ok=True)
    size_before = src_docx.stat().st_size

    dropped_parts: list[str] = []
    with zipfile.ZipFile(str(src_docx), "r") as zin:
        names = set(zin.namelist())
        # Identify media actually used inside word/document.xml body
        referenced_media, used_rids = _collect_referenced_media_from_doc(zin)
        # Restrict to media actually present
        kept_media = {m for m in referenced_media if m in names}

        # Load source parts we keep verbatim
        styles_bytes = zin.read("word/styles.xml") if "word/styles.xml" in names else None
        has_numbering = "word/numbering.xml" in names
        numbering_bytes = zin.read("word/numbering.xml") if has_numbering else None
        has_theme = "word/theme/theme1.xml" in names
        theme_bytes = zin.read("word/theme/theme1.xml") if has_theme else None
        core_bytes = zin.read("docProps/core.xml") if "docProps/core.xml" in names else None
        doc_xml_bytes = zin.read("word/document.xml")
        doc_rels_bytes = zin.read("word/_rels/document.xml.rels") if "word/_rels/document.xml.rels" in names else b""

        # Strip body refs to dropped parts
        doc_xml_bytes_stripped = _strip_document_xml_for_aggressive(doc_xml_bytes)
        # Strip rels: keep only styles/numbering/theme + body-used image rels
        doc_rels_stripped = _strip_doc_rels_for_aggressive(doc_rels_bytes, kept_media, used_rids)

        # Build content types
        content_types = _build_aggressive_content_types(has_numbering, has_theme, kept_media)

        # Track dropped parts (for report)
        for n in sorted(names):
            if n in _AGGRESSIVE_KEEP_EXACT:
                continue
            if n in kept_media:
                continue
            if n.startswith("word/media/"):
                # media not referenced by document.xml.rels → drop
                dropped_parts.append(n)
                continue
            dropped_parts.append(n)

        # Write new zip
        tmp = out_path.with_suffix(out_path.suffix + ".aggressive.tmp")
        with zipfile.ZipFile(str(tmp), "w", zipfile.ZIP_DEFLATED) as zout:
            zout.writestr("[Content_Types].xml", content_types)
            zout.writestr("_rels/.rels", _build_aggressive_root_rels())
            zout.writestr("word/document.xml", doc_xml_bytes_stripped)
            zout.writestr("word/_rels/document.xml.rels", doc_rels_stripped)
            if styles_bytes is not None:
                zout.writestr("word/styles.xml", styles_bytes)
            else:
                # degrade safely
                zout.writestr(
                    "word/styles.xml",
                    b'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
                    b'<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"/>',
                )
            if has_numbering:
                zout.writestr("word/numbering.xml", numbering_bytes)
            if has_theme:
                zout.writestr("word/theme/theme1.xml", theme_bytes)
            if core_bytes is not None:
                zout.writestr("docProps/core.xml", core_bytes)
            else:
                zout.writestr(
                    "docProps/core.xml",
                    b'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
                    b'<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties"/>',
                )
            # Copy kept media verbatim
            for mp in sorted(kept_media):
                zout.writestr(mp, zin.read(mp))

    shutil.move(str(tmp), str(out_path))
    size_after = out_path.stat().st_size

    return {
        "mode": "aggressive",
        "input": str(src_docx),
        "output": str(out_path),
        "size_before": size_before,
        "size_after": size_after,
        "delta_bytes": size_before - size_after,
        "delta_pct": round((size_before - size_after) / size_before * 100, 2) if size_before else 0,
        "kept_media": len(kept_media),
        "dropped_parts_count": len(dropped_parts),
        "dropped_parts_sample": dropped_parts[:20],
    }


# ---------------- CLI ----------------

def main(argv: list[str] | None = None) -> int:
    ap = argparse.ArgumentParser(
        description=__doc__,
        formatter_class=argparse.RawDescriptionHelpFormatter,
    )
    ap.add_argument("docx", type=Path, help="source docx path")
    ap.add_argument("-o", "--output", type=Path, default=None,
                    help="output path (safe: 默认覆写原文件保留 .bak; aggressive: 必填)")
    ap.add_argument("--mode", choices=["safe", "aggressive"], default="safe",
                    help="safe = ensemble strip (默认); aggressive = 最小骨架重建 (不可逆)")
    args = ap.parse_args(argv)

    if args.mode == "safe":
        rep = run_safe(args.docx, args.output)
    else:
        if args.output is None:
            print("[ERR] aggressive 模式必须显式 -o OUT (拒绝原位覆写)", file=sys.stderr)
            return 2
        rep = run_aggressive(args.docx, args.output)

    print(f"[slim/{rep['mode']}] {rep['input']}")
    print(f"  size: {rep['size_before']} -> {rep['size_after']} "
          f"(delta {rep['delta_bytes']} bytes / {rep['delta_pct']}%)")
    if rep["mode"] == "safe":
        for step_name, sr in rep["steps"]:
            print(f"  - {step_name}: ok={sr.get('ok', '?')}")
    else:
        print(f"  kept_media={rep['kept_media']} dropped_parts={rep['dropped_parts_count']}")
        if rep["dropped_parts_sample"]:
            print(f"  dropped_sample: {rep['dropped_parts_sample'][:5]}...")
    print(f"  out: {rep['output']}")
    return 0


# ---------------- group register (called by sub/__init__.py register_all) ----------------

def register(subparsers) -> None:
    """Register top-level 'slim' command group on docx_cli main subparsers."""
    p = subparsers.add_parser(
        "slim",
        help="docx slim - safe (ensemble strip, 30-50%% reduction) / aggressive (minimal skeleton, 80-95%%, irreversible)",
    )
    p.add_argument("docx", nargs="?", type=Path, help="source docx path")
    p.add_argument("-o", "--output", type=Path, default=None,
                   help="output path (safe: default in-place + .bak; aggressive: required)")
    p.add_argument("--mode", choices=["safe", "aggressive"], default="safe",
                   help="safe = ensemble strip (default); aggressive = minimal skeleton (irreversible)")
    p.set_defaults(func=_dispatch_run)


def _dispatch_run(args) -> int:
    if args.docx is None:
        print(__doc__)
        print("\nUsage: docx_cli slim <docx> [-o OUT] [--mode {safe,aggressive}]")
        return 0
    argv = [str(args.docx)]
    if args.output:
        argv.extend(["-o", str(args.output)])
    argv.extend(["--mode", args.mode])
    return main(argv)


if __name__ == "__main__":
    sys.exit(main())
