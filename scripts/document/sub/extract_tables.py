#!/usr/bin/env python3
# distilled from eco-flow/taizhou-天台 table extract need (2026-05-28)
r"""extract_tables.py — 从 docx 抽出每张表为独立 docx。

文件名 = 邻近 caption 段文字（如 "表2.2-1-水库基本情况.docx"）。
保留原 docx 全套 styles / numbering / sectPr / page setup / rels / media
(整套 zip parts 复制, 只改 document.xml 的 body), 与 split_by_h1.py 同套路。

CLI:
    python3 scripts/document/sub/extract_tables.py \
        --docx <path> --out-dir <dir> \
        [--name-pattern '{stem}.docx'] [--dry-run]

算法:
    1. 用 lxml 遍历 word/document.xml body, 找所有 <w:tbl> 元素位置
    2. 每张表往前扫最近 1-2 段 <w:p>, 取以 "表" 开头 < 80 字符的 → caption
       若前面没有, 往后扫 1 段; 再无 → fallback `table-{idx:02d}`
    3. caption sanitize → 文件名 stem
    4. shutil.copy 源 docx → 打开 copy, body 内只留目标 <w:tbl> + 它的 caption 段
       + 末尾 <w:sectPr>, 删其余 → save
    5. 重名加 `-2`, `-3` ...
"""
from __future__ import annotations

import argparse
import re
import shutil
import sys
from pathlib import Path
from typing import Optional

try:
    from docx import Document
    from docx.oxml.ns import qn
except ImportError:
    print("ERROR: python-docx 未安装 (pip install python-docx)", file=sys.stderr)
    sys.exit(2)


_ILLEGAL_FILENAME_RE = re.compile(r'[/\\:*?"<>|\r\n\t]')
_MULTI_WS_RE = re.compile(r"\s+")

# Caption heuristic: starts with 表 / Table; length < 80 chars
_CAPTION_PREFIX_RE = re.compile(r"^\s*(?:表|Table\b)", re.IGNORECASE)
_CAPTION_MAX_LEN = 80


def sanitize_filename(name: str, max_len: int = 100) -> str:
    """Replace illegal filename chars with _, compress whitespace, strip, truncate."""
    if not name:
        return "untitled"
    n = _ILLEGAL_FILENAME_RE.sub("_", name)
    n = _MULTI_WS_RE.sub(" ", n).strip()
    if not n:
        return "untitled"
    if len(n) > max_len:
        n = n[:max_len].rstrip()
    return n


def _paragraph_text(p_elem) -> str:
    """Extract concatenated text from <w:t> nodes under a <w:p>."""
    return "".join(t.text or "" for t in p_elem.iter(qn("w:t")))


def _is_caption_like(text: str) -> bool:
    """True if text looks like a table caption (starts with 表/Table, short)."""
    if not text:
        return False
    t = text.strip()
    if not t or len(t) >= _CAPTION_MAX_LEN:
        return False
    return bool(_CAPTION_PREFIX_RE.match(t))


def _find_caption(children: list, tbl_idx: int) -> Optional[str]:
    """Scan back up to 2 paragraphs, else forward 1, for a caption-like <w:p>.

    Returns the caption text (raw, un-sanitized) or None.
    """
    # Backward up to 2 non-empty paragraphs
    seen_paras = 0
    for i in range(tbl_idx - 1, -1, -1):
        elem = children[i]
        if elem.tag != qn("w:p"):
            continue
        text = _paragraph_text(elem).strip()
        if not text:
            continue
        seen_paras += 1
        if _is_caption_like(text):
            return text
        if seen_paras >= 2:
            break

    # Forward 1 paragraph (some templates put caption below the table)
    for i in range(tbl_idx + 1, len(children)):
        elem = children[i]
        if elem.tag == qn("w:tbl"):
            break  # hit next table; stop
        if elem.tag != qn("w:p"):
            continue
        text = _paragraph_text(elem).strip()
        if not text:
            continue
        if _is_caption_like(text):
            return text
        break  # first non-empty para wasn't a caption; give up

    return None


def plan_extracts(docx_path: Path, doc=None) -> tuple[list[dict], int]:
    """Inspect docx → return (extracts, tbl_count).

    extracts: list[{idx, caption, stem, tbl_idx_in_body, caption_idx_in_body}]
        caption_idx_in_body == -1 if no caption found (fallback to table-XX)
    """
    if doc is None:
        doc = Document(str(docx_path))
    body = doc.element.body
    children = list(body)

    # Locate sectPr index (final node, exclude from scan)
    sect_idx = len(children)
    for i in range(len(children) - 1, -1, -1):
        if children[i].tag == qn("w:sectPr"):
            sect_idx = i
            break

    extracts: list[dict] = []
    idx_counter = 0
    used_stems: dict[str, int] = {}

    for i in range(sect_idx):
        elem = children[i]
        if elem.tag != qn("w:tbl"):
            continue
        caption = _find_caption(children, i)
        if caption:
            stem_raw = caption
        else:
            stem_raw = f"table-{idx_counter:02d}"
        stem = sanitize_filename(stem_raw)
        # Disambiguate duplicates
        base_stem = stem
        n = used_stems.get(base_stem, 0)
        if n > 0:
            stem = f"{base_stem}-{n + 1}"
        used_stems[base_stem] = n + 1

        # Determine caption_idx (the actual element index of the caption paragraph)
        caption_idx = -1
        if caption:
            # Walk back to find the matching caption paragraph
            seen = 0
            for j in range(i - 1, -1, -1):
                e2 = children[j]
                if e2.tag != qn("w:p"):
                    continue
                t2 = _paragraph_text(e2).strip()
                if not t2:
                    continue
                seen += 1
                if t2 == caption.strip():
                    caption_idx = j
                    break
                if seen >= 2:
                    break
            if caption_idx == -1:
                # try forward
                for j in range(i + 1, len(children)):
                    e2 = children[j]
                    if e2.tag == qn("w:tbl"):
                        break
                    if e2.tag != qn("w:p"):
                        continue
                    t2 = _paragraph_text(e2).strip()
                    if not t2:
                        continue
                    if t2 == caption.strip():
                        caption_idx = j
                    break

        extracts.append({
            "idx": idx_counter,
            "caption": caption,
            "stem": stem,
            "tbl_idx_in_body": i,
            "caption_idx_in_body": caption_idx,
        })
        idx_counter += 1

    return extracts, idx_counter


def write_extract(
    src_docx: Path,
    dst_docx: Path,
    tbl_idx: int,
    caption_idx: int,
) -> tuple[int, int]:
    """Copy src→dst, prune body to keep only target <w:tbl> (+ optional caption) + sectPr.

    Returns (tables_in_extracted, file_bytes).
    """
    dst_docx.parent.mkdir(parents=True, exist_ok=True)
    shutil.copy(str(src_docx), str(dst_docx))
    doc = Document(str(dst_docx))
    body = doc.element.body
    children = list(body)

    # Identify sectPr (last)
    sect_elem = None
    for i in range(len(children) - 1, -1, -1):
        if children[i].tag == qn("w:sectPr"):
            sect_elem = children[i]
            break

    keep_indices = {tbl_idx}
    if caption_idx >= 0:
        keep_indices.add(caption_idx)

    for i, elem in enumerate(children):
        if elem is sect_elem:
            continue
        if i not in keep_indices:
            body.remove(elem)
    doc.save(str(dst_docx))

    doc2 = Document(str(dst_docx))
    tbl_count = len(doc2.tables)
    file_bytes = dst_docx.stat().st_size
    return tbl_count, file_bytes


def run_extract(
    src_docx: Path,
    out_dir: Path,
    dry_run: bool = False,
    name_pattern: str = "{stem}.docx",
    doc=None,
) -> dict:
    """Execute table-extract; reuses provided `doc` for planning if given."""
    src = Path(src_docx).expanduser().resolve()
    if not src.exists():
        return {"error": f"input docx not found: {src}", "exit_code": 2}
    out_dir = Path(out_dir).expanduser().resolve()

    extracts, tbl_count = plan_extracts(src, doc=doc)

    if tbl_count == 0:
        return {"tbl_count": 0, "extracts_emitted": 0, "note": "no tables in docx"}

    if dry_run:
        plan = []
        for e in extracts:
            fname = name_pattern.format(stem=e["stem"], idx=e["idx"])
            plan.append({
                "idx": e["idx"], "fname": fname,
                "caption": e["caption"], "tbl_idx": e["tbl_idx_in_body"],
            })
        return {"tbl_count": tbl_count, "extracts_planned": plan, "dry_run": True}

    out_dir.mkdir(parents=True, exist_ok=True)
    emitted = []
    failed = []
    for e in extracts:
        fname = name_pattern.format(stem=e["stem"], idx=e["idx"])
        dst = out_dir / fname
        try:
            tc, nbytes = write_extract(
                src, dst, e["tbl_idx_in_body"], e["caption_idx_in_body"],
            )
            emitted.append({
                "idx": e["idx"], "fname": fname,
                "tables": tc, "bytes": nbytes,
                "caption": e["caption"],
            })
        except Exception as exc:
            failed.append({
                "idx": e["idx"], "fname": fname,
                "error": f"{type(exc).__name__}: {exc}",
            })
    return {
        "tbl_count": tbl_count,
        "extracts_emitted": len(emitted),
        "extracts_failed": len(failed),
        "out_dir": str(out_dir),
        "emitted": emitted,
        "failed": failed,
    }


def main() -> int:
    ap = argparse.ArgumentParser(
        description="Extract each table in a DOCX as an independent DOCX file "
                    "(filename = neighboring caption text, or table-XX fallback).",
    )
    ap.add_argument("--docx", required=True, help="input docx path")
    ap.add_argument("--out-dir", required=True, help="output directory (mkdir -p)")
    ap.add_argument(
        "--name-pattern",
        default="{stem}.docx",
        help="output filename pattern, default '{stem}.docx' "
             "(available: {stem}, {idx})",
    )
    ap.add_argument(
        "--dry-run", action="store_true",
        help="print plan only, don't write files",
    )
    args = ap.parse_args()

    src = Path(args.docx).expanduser().resolve()
    if not src.exists():
        print(f"ERROR: input docx not found: {src}", file=sys.stderr)
        return 2
    out_dir = Path(args.out_dir).expanduser().resolve()

    extracts, tbl_count = plan_extracts(src)

    if tbl_count == 0:
        print(f"[extract-tables] 0 tables in docx — nothing to extract")
        return 0

    if args.dry_run:
        print(f"# DRY RUN — would write {len(extracts)} files to {out_dir}/")
        for e in extracts:
            fname = args.name_pattern.format(stem=e["stem"], idx=e["idx"])
            cap = (e["caption"] or "[no caption]")[:60]
            print(f"  [{e['idx']:>2}] tbl@body[{e['tbl_idx_in_body']}]  → {fname}  (caption={cap!r})")
        return 0

    out_dir.mkdir(parents=True, exist_ok=True)
    n_ok = 0
    for e in extracts:
        fname = args.name_pattern.format(stem=e["stem"], idx=e["idx"])
        dst = out_dir / fname
        try:
            tc, nbytes = write_extract(
                src, dst, e["tbl_idx_in_body"], e["caption_idx_in_body"],
            )
            print(f"  [{e['idx']:>2}] {fname}  · {tc} tables · {nbytes:,} bytes")
            n_ok += 1
        except Exception as exc:
            print(f"  [{e['idx']:>2}] FAILED {fname}: {type(exc).__name__}: {exc}",
                  file=sys.stderr)
    print(f"OK: {n_ok} files written to {out_dir}")
    return 0 if n_ok == len(extracts) else 1


if __name__ == "__main__":
    sys.exit(main())
