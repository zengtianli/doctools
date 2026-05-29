"""pdf_pipeline_lib.py — PDF batch pipeline engine + step registry (2026-05-28)

Same design as docx pipeline (`sub/pipeline_lib.py`),swap underlay to
pdfplumber + pypdf + Poppler CLI (pdfimages / pdftotext / pdfinfo).

Two step kinds
--------------
1. **pdf-based step** — `fn(pdf, args, out_dir) -> dict`
   - `pdf` is a `pdfplumber.open(path)` object already opened by caller
   - Suits: same-PDF multi-pass (text / table extract)
   - Each PDF opened once,N steps reuse,no re-parse

2. **path-based step** — `fn(pdf_path: Path, args, out_dir) -> dict`
   - Manages file IO itself (typical: subprocess wrap pdfimages)
   - No pdfplumber object required → avoid unnecessary parse

Step registry
-------------
`_BUILTIN_STEPS: dict[str, tuple[str, Callable]]`, kind ∈ {"pdf", "path"}.
Add a step = write `_<verb>(...)` + one registry line.

Concurrency model
-----------------
- Single PDF + N steps → `run_pipeline_single`,one pdfplumber open reused
- N PDFs + N steps → `run_pipeline_parallel`,ProcessPoolExecutor across PDFs,
  each worker still reuses open within its PDF

Per-step out_dir resolution
---------------------------
Per step override via `--<step-name>-out-dir`,default → `<pdf-parent>/<verb>/`.
"""

from __future__ import annotations

import argparse
import csv
import os
import re
import subprocess
import tempfile
import time
from concurrent.futures import ProcessPoolExecutor, as_completed
from dataclasses import dataclass
from pathlib import Path
from typing import Any, Callable, Optional

try:
    import pdfplumber  # type: ignore
except ImportError:
    pdfplumber = None  # type: ignore


# ─── page-range parsing ────────────────────────────────────────────────

def parse_page_spec(spec: Optional[str], total_pages: int) -> list[int]:
    """Parse "1-5,8,10-12" → sorted 0-indexed page list.

    spec=None → all pages.
    Validates against total_pages; clamps to [1, total_pages]; raises on invalid syntax.
    """
    if not spec:
        return list(range(total_pages))
    out: set[int] = set()
    for tok in spec.split(","):
        tok = tok.strip()
        if not tok:
            continue
        m = re.match(r"^(\d+)(?:-(\d+))?$", tok)
        if not m:
            raise ValueError(f"invalid page spec token: {tok!r}")
        a = int(m.group(1))
        b = int(m.group(2)) if m.group(2) else a
        if a < 1 or b < 1 or a > b:
            raise ValueError(f"invalid page range: {tok!r}")
        for p in range(a, b + 1):
            if 1 <= p <= total_pages:
                out.add(p - 1)
    return sorted(out)


def parse_ranges_spec(spec: str) -> list[tuple[int, int]]:
    """Parse "1-10,11-20,21-30" → [(1,10),(11,20),(21,30)] (1-indexed inclusive)."""
    out: list[tuple[int, int]] = []
    for tok in spec.split(","):
        tok = tok.strip()
        if not tok:
            continue
        m = re.match(r"^(\d+)-(\d+)$", tok)
        if not m:
            raise ValueError(f"invalid range token: {tok!r} (need A-B)")
        a, b = int(m.group(1)), int(m.group(2))
        if a < 1 or b < a:
            raise ValueError(f"invalid range: {tok!r}")
        out.append((a, b))
    if not out:
        raise ValueError("no ranges parsed")
    return out


def safe_filename(s: str, maxlen: int = 120) -> str:
    """Sanitize bookmark titles for filesystem use."""
    s = re.sub(r"[/\\\x00-\x1f]+", "_", s).strip()
    s = re.sub(r"\s+", "_", s)
    return (s[:maxlen] or "untitled")


# ─── caption detection helpers ────────────────────────────────────────
#
# Detect lines like "图 3-7 台州市干旱指数年内分配过程图" / "表 1-1 ...".
# Anchored at start of line to avoid in-body refs like "如图 1-1 所示".
# Separator allowed: "-", "．", "."; gap between code and title allowed
# to be one or more whitespace (incl. full-width).
_IMG_CAPTION_RE = re.compile(
    r"^\s*图\s*(\d+)[\-．.](\d+)\s+(.{2,60}?)\s*$"
)
_TBL_CAPTION_RE = re.compile(
    r"^\s*表\s*(\d+)[\-．.](\d+)\s+(.{2,60}?)\s*$"
)

# Filesystem-illegal chars in caption titles → replace with "_".
_FN_BAD_CHARS_RE = re.compile(r'[/\\:*?"<>|\x00-\x1f]+')
# Whitespace (including full-width space U+3000) → "_".
_FN_WS_RE = re.compile(r"[\s　]+")


def _sanitize_caption_filename(s: str, maxlen: int = 80) -> str:
    """Sanitize a caption string into a filesystem-safe stem."""
    s = _FN_BAD_CHARS_RE.sub("_", s).strip()
    s = _FN_WS_RE.sub("_", s)
    s = s.strip("._")
    if len(s) > maxlen:
        s = s[:maxlen].rstrip("._")
    return s or "untitled"


def _build_caption_stem(m: re.Match) -> str:
    """From a caption regex match → stem like '图3-7_台州市干旱指数年内分配过程图'."""
    kind_char = m.string.lstrip()[0]  # "图" or "表"
    major, minor, title = m.group(1), m.group(2), m.group(3)
    title_clean = _sanitize_caption_filename(title)
    return f"{kind_char}{major}-{minor}_{title_clean}"


def _find_captions_on_page(page, regex: re.Pattern) -> list[tuple[str, float]]:
    """Return [(stem, y_top), ...] for caption-matching lines on the page,
    sorted by y_top (reading order).

    Uses pdfplumber.Page.extract_text_lines() when available (gives both
    line text and top y-coord). Falls back to extract_text() split-by-line
    with synthetic y = line index when not.
    """
    out: list[tuple[str, float]] = []
    lines_data: list[tuple[str, float]] = []
    try:
        for ln in page.extract_text_lines() or []:
            text = ln.get("text") or ""
            top = float(ln.get("top", 0.0))
            lines_data.append((text, top))
    except Exception:
        try:
            txt = page.extract_text() or ""
        except Exception:
            txt = ""
        for idx, line in enumerate(txt.split("\n")):
            lines_data.append((line, float(idx)))

    for text, top in lines_data:
        m = regex.match(text)
        if not m:
            continue
        stem = _build_caption_stem(m)
        out.append((stem, top))
    out.sort(key=lambda x: x[1])
    return out


def _unique_path(out_dir: Path, stem: str, ext: str) -> Path:
    """Return out_dir/<stem><ext> avoiding collision via -2/-3 suffix.

    `ext` includes the leading dot (".jpg", ".csv", etc).
    """
    base = out_dir / f"{stem}{ext}"
    if not base.exists():
        return base
    i = 2
    while True:
        cand = out_dir / f"{stem}-{i}{ext}"
        if not cand.exists():
            return cand
        i += 1


# Threshold (bytes) below which a pdfimages output is treated as noise
# (soft-mask, single-color filler). Empirically 2-2.5KB files are noise,
# real figures start at ~25KB. 3KB cleanly separates the two.
MIN_IMAGE_BYTES = 3072


# ─── built-in step implementations ────────────────────────────────────

def _text_extract(pdf, args, out_dir: Path) -> dict:
    """pdf-based: pdfplumber extract_text() per page to .txt files.

    args.text_extract_pages — page spec (e.g. "1-5"), None=all
    args.text_extract_single — bool, concat all into full.txt
    """
    if pdf is None:
        return {"step": "text-extract",
                "error": "pdfplumber not available or pdf is None"}
    t0 = time.perf_counter()
    out_dir.mkdir(parents=True, exist_ok=True)
    spec = (getattr(args, "text_extract_pages", None)
            or getattr(args, "pages", None))
    single = bool(getattr(args, "text_extract_single", False)
                  or getattr(args, "single", False))
    total = len(pdf.pages)
    page_idx = parse_page_spec(spec, total)
    pad = max(3, len(str(total)))
    pages_written: list[str] = []
    chunks: list[str] = []
    for i in page_idx:
        try:
            txt = pdf.pages[i].extract_text() or ""
        except Exception as e:
            txt = f"[ERROR extract_text page {i+1}: {type(e).__name__}: {e}]"
        if single:
            chunks.append(f"\n\n===== page {i+1} =====\n\n{txt}")
        else:
            fn = out_dir / f"page-{str(i+1).zfill(pad)}.txt"
            fn.write_text(txt, encoding="utf-8")
            pages_written.append(fn.name)
    if single:
        full = out_dir / "full.txt"
        full.write_text("".join(chunks), encoding="utf-8")
        pages_written = ["full.txt"]
    return {
        "step": "text-extract",
        "pages_requested": len(page_idx),
        "files_written": len(pages_written),
        "out_dir": str(out_dir),
        "elapsed_s": round(time.perf_counter() - t0, 3),
        "mode": "single" if single else "per-page",
    }


def _table_extract(pdf, args, out_dir: Path) -> dict:
    """pdf-based: pdfplumber.extract_tables() per page → CSV files.

    Naming: prefer matched table-caption stems on the page
    (e.g. "表1-1_台州市主要河流特征表.csv"); fall back to
    "page-{NNN}-table-{M}.csv" when caption count < table count.

    Pairing strategy: per page, sort captions by y_top (reading order),
    and pair with tables in the order returned by pdfplumber
    (pdfplumber.extract_tables() already returns tables in page-y order).

    args.table_extract_pages — page spec; None=all
    """
    if pdf is None:
        return {"step": "table-extract",
                "error": "pdfplumber not available or pdf is None"}
    t0 = time.perf_counter()
    out_dir.mkdir(parents=True, exist_ok=True)
    spec = (getattr(args, "table_extract_pages", None)
            or getattr(args, "pages", None))
    total = len(pdf.pages)
    page_idx = parse_page_spec(spec, total)
    pad = max(3, len(str(total)))
    per_page: dict[str, int] = {}
    total_tables = 0
    caption_named = 0
    fallback_named = 0
    files: list[str] = []
    for i in page_idx:
        page = pdf.pages[i]
        try:
            tables = page.extract_tables() or []
        except Exception as e:
            per_page[str(i + 1)] = -1
            files.append(f"[ERROR p{i+1}: {type(e).__name__}: {e}]")
            continue
        per_page[str(i + 1)] = len(tables)
        if not tables:
            continue
        try:
            captions = _find_captions_on_page(page, _TBL_CAPTION_RE)
        except Exception:
            captions = []
        for j, tbl in enumerate(tables):
            if j < len(captions):
                stem = captions[j][0]
                fn = _unique_path(out_dir, stem, ".csv")
                caption_named += 1
            else:
                stem = f"page-{str(i+1).zfill(pad)}-table-{j+1}"
                fn = _unique_path(out_dir, stem, ".csv")
                fallback_named += 1
            with fn.open("w", newline="", encoding="utf-8") as f:
                w = csv.writer(f)
                for row in tbl:
                    w.writerow(["" if c is None else c for c in row])
            files.append(fn.name)
            total_tables += 1
    return {
        "step": "table-extract",
        "pages_scanned": len(page_idx),
        "total_tables": total_tables,
        "caption_named": caption_named,
        "fallback_named": fallback_named,
        "per_page_table_counts": per_page,
        "out_dir": str(out_dir),
        "elapsed_s": round(time.perf_counter() - t0, 3),
    }


def _image_extract(pdf_path: Path, args, out_dir: Path) -> dict:
    """path-based: per-page pdfimages -all + caption-based renaming.

    Algorithm (per page):
      a. open page in pdfplumber, regex-match "图 X-Y title" captions →
         list of (stem, y_top), sorted by y_top.
      b. subprocess pdfimages -f N -l N to a tempdir.
      c. drop files < MIN_IMAGE_BYTES (soft-mask / pure-color noise).
      d. zip captions ↔ remaining files in extraction order;
         caption stem → final name. Extras → fallback "page-NNN-img-MM.<ext>".

    args.image_extract_pages — page spec; None=all
    """
    t0 = time.perf_counter()
    out_dir.mkdir(parents=True, exist_ok=True)

    if pdfplumber is None:
        return {
            "step": "image-extract",
            "error": "pdfplumber not installed; required for caption-named image extraction",
        }

    # Determine page list. Need total pages → open pdfplumber once.
    spec = (getattr(args, "image_extract_pages", None)
            or getattr(args, "pages", None))

    pdfimages_bin = "/opt/homebrew/bin/pdfimages"
    rep_files: list[str] = []
    caption_named = 0
    fallback_named = 0
    noise_dropped = 0
    rc_nonzero_pages: list[int] = []
    stderr_first: str = ""

    try:
        pdf_obj = pdfplumber.open(str(pdf_path))
    except Exception as e:
        return {
            "step": "image-extract",
            "error": f"pdfplumber.open failed: {type(e).__name__}: {e}",
        }

    try:
        total = len(pdf_obj.pages)
        page_idx = parse_page_spec(spec, total)
        pad = max(3, len(str(total)))

        for i in page_idx:
            page_num = i + 1  # pdfimages is 1-based
            try:
                page = pdf_obj.pages[i]
                captions = _find_captions_on_page(page, _IMG_CAPTION_RE)
            except Exception:
                captions = []

            with tempfile.TemporaryDirectory(prefix="pdfimg_") as td:
                tdp = Path(td)
                cmd = [
                    pdfimages_bin, "-all",
                    "-f", str(page_num), "-l", str(page_num),
                    str(pdf_path), str(tdp / "img"),
                ]
                cp = subprocess.run(cmd, capture_output=True, text=True)
                if cp.returncode != 0:
                    rc_nonzero_pages.append(page_num)
                    if not stderr_first:
                        stderr_first = (cp.stderr or "").strip()[:300]

                # Collect produced files in extraction order
                # (pdfimages numbers img-NNN sequentially in page-y order;
                # main images, masks and smasks interleave by object).
                produced = sorted(
                    [p for p in tdp.iterdir() if p.is_file()],
                    key=lambda p: p.name,
                )

                # Filter noise by size threshold.
                kept: list[Path] = []
                for p in produced:
                    try:
                        if p.stat().st_size < MIN_IMAGE_BYTES:
                            noise_dropped += 1
                            continue
                    except OSError:
                        continue
                    kept.append(p)

                # Pair captions (already sorted by y) with kept files
                # (sorted by pdfimages order = page-y order). Extras → fallback.
                for k, srcp in enumerate(kept):
                    ext = srcp.suffix.lower() or ".bin"
                    if k < len(captions):
                        stem = captions[k][0]
                        dstp = _unique_path(out_dir, stem, ext)
                        caption_named += 1
                    else:
                        stem = f"page-{str(page_num).zfill(pad)}-img-{k:02d}"
                        dstp = _unique_path(out_dir, stem, ext)
                        fallback_named += 1
                    try:
                        srcp.replace(dstp)
                    except OSError:
                        # cross-fs fallback
                        dstp.write_bytes(srcp.read_bytes())
                        try:
                            srcp.unlink()
                        except OSError:
                            pass
                    rep_files.append(dstp.name)
    finally:
        try:
            pdf_obj.close()
        except Exception:
            pass

    rep = {
        "step": "image-extract",
        "pages_scanned": len(page_idx),
        "images_written": len(rep_files),
        "caption_named": caption_named,
        "fallback_named": fallback_named,
        "noise_dropped": noise_dropped,
        "min_image_bytes": MIN_IMAGE_BYTES,
        "out_dir": str(out_dir),
        "elapsed_s": round(time.perf_counter() - t0, 3),
    }
    if rc_nonzero_pages:
        rep["pdfimages_rc_nonzero_pages"] = rc_nonzero_pages
        if stderr_first:
            rep["pdfimages_stderr_first"] = stderr_first
    return rep


# Registry: name → (kind, callable)
# kind="pdf"  → fn(pdf: pdfplumber.PDF, args, out_dir: Path) -> dict
# kind="path" → fn(pdf_path: Path,      args, out_dir: Path) -> dict
_BUILTIN_STEPS: dict[str, tuple[str, Callable[..., dict]]] = {
    "image-extract": ("path", _image_extract),
    "text-extract":  ("pdf",  _text_extract),
    "table-extract": ("pdf",  _table_extract),
}


def list_builtin_steps() -> list[str]:
    return sorted(_BUILTIN_STEPS.keys())


def is_builtin_step(name: str) -> bool:
    return name in _BUILTIN_STEPS


@dataclass
class LoadedStep:
    name: str
    kind: str  # "pdf" or "path"
    fn: Callable[..., dict]


def load_step(name: str) -> LoadedStep:
    if name not in _BUILTIN_STEPS:
        raise KeyError(
            f"unknown step: {name!r}; known: {', '.join(list_builtin_steps())}"
        )
    kind, fn = _BUILTIN_STEPS[name]
    return LoadedStep(name=name, kind=kind, fn=fn)


# ─── per-step out_dir resolution ──────────────────────────────────────

def _resolve_step_out_dir(step_name: str, args, pdf_path: Path) -> Path:
    """Resolve out_dir for a step in priority order:
    1) explicit --<step-key>-out-dir (e.g. --text-extract-out-dir DIR)
    2) global --out-dir DIR → DIR/<step-verb>/
    3) <pdf-parent>/<step-verb>/
    """
    attr = step_name.replace("-", "_") + "_out_dir"
    explicit = getattr(args, attr, None)
    if explicit:
        return Path(str(explicit))
    global_root = getattr(args, "out_dir", None)
    if global_root:
        return Path(str(global_root)) / step_name
    return pdf_path.parent / step_name


# ─── single-pdf pipeline (multi-step, one pdfplumber.open) ────────────

def run_pipeline_single(
    pdf_path: Path | str,
    step_names: list[str],
    args: argparse.Namespace | None = None,
) -> dict:
    """Run N steps against 1 PDF; open with pdfplumber at most once."""
    pdf_path = Path(pdf_path).resolve()
    if not pdf_path.is_file():
        raise FileNotFoundError(pdf_path)

    if args is None:
        args = argparse.Namespace()
    if not hasattr(args, "pdf"):
        args.pdf = pdf_path

    t0 = time.perf_counter()
    report: dict[str, Any] = {
        "pdf": str(pdf_path),
        "steps": {},
        "timing": {},
    }

    loaded = [load_step(n) for n in step_names]
    needs_pdf = any(s.kind == "pdf" for s in loaded)

    pdf_obj = None
    if needs_pdf:
        if pdfplumber is None:
            raise RuntimeError(
                "pdfplumber not installed; required by steps: "
                + ",".join(s.name for s in loaded if s.kind == "pdf")
            )
        t_open = time.perf_counter()
        try:
            pdf_obj = pdfplumber.open(str(pdf_path))
        except Exception as e:
            report["error"] = f"pdfplumber.open failed: {type(e).__name__}: {e}"
            report["timing"]["total"] = round(time.perf_counter() - t0, 3)
            return report
        report["timing"]["open"] = round(time.perf_counter() - t_open, 3)

    try:
        for s in loaded:
            out_dir = _resolve_step_out_dir(s.name, args, pdf_path)
            t_s = time.perf_counter()
            try:
                if s.kind == "pdf":
                    rep = s.fn(pdf_obj, args, out_dir)
                else:
                    rep = s.fn(pdf_path, args, out_dir)
            except Exception as exc:
                rep = {"step": s.name,
                       "error": f"{type(exc).__name__}: {exc}"}
            report["steps"][s.name] = rep
            report["timing"][f"step:{s.name}"] = round(
                time.perf_counter() - t_s, 3
            )
    finally:
        if pdf_obj is not None:
            try:
                pdf_obj.close()
            except Exception:
                pass

    report["timing"]["total"] = round(time.perf_counter() - t0, 3)
    return report


# ─── parallel across PDFs ─────────────────────────────────────────────

def _worker_payload_to_args(payload: dict) -> argparse.Namespace:
    """Reconstruct args namespace inside child process."""
    ns = argparse.Namespace()
    for k, v in payload.get("args_dict", {}).items():
        setattr(ns, k, v)
    ns.pdf = Path(payload["pdf"])
    return ns


def _worker(payload: dict) -> dict:
    args = _worker_payload_to_args(payload)
    return run_pipeline_single(
        pdf_path=payload["pdf"],
        step_names=payload["steps"],
        args=args,
    )


def run_pipeline_parallel(
    pdf_list: list[Path | str],
    step_names: list[str],
    max_workers: int | None = None,
    args_dict: dict | None = None,
) -> dict[str, dict]:
    """Process-pool across PDFs; each worker re-runs run_pipeline_single."""
    if not pdf_list:
        return {}
    if max_workers is None:
        max_workers = min(len(pdf_list), os.cpu_count() or 4)
    args_dict = args_dict or {}
    payloads = [
        {"pdf": str(p), "steps": list(step_names), "args_dict": args_dict}
        for p in pdf_list
    ]
    results: dict[str, dict] = {}
    if max_workers <= 1 or len(payloads) == 1:
        for p in payloads:
            try:
                results[p["pdf"]] = _worker(p)
            except Exception as exc:
                results[p["pdf"]] = {
                    "pdf": p["pdf"],
                    "error": f"{type(exc).__name__}: {exc}",
                }
        return results
    with ProcessPoolExecutor(max_workers=max_workers) as ex:
        fut_map = {ex.submit(_worker, p): p["pdf"] for p in payloads}
        for fut in as_completed(fut_map):
            pdf = fut_map[fut]
            try:
                results[pdf] = fut.result()
            except Exception as exc:
                results[pdf] = {
                    "pdf": pdf,
                    "error": f"{type(exc).__name__}: {exc}",
                }
    return results


# ─── timing pretty-printer ─────────────────────────────────────────────

def format_timing_table(results: dict[str, dict], wall: float) -> str:
    lines = []
    lines.append("=" * 78)
    lines.append("PDF PIPELINE TIMING REPORT")
    lines.append("=" * 78)
    lines.append(f"{'pdf':46s}  {'open':>6s}  {'total':>8s}")
    lines.append("-" * 78)
    step_set: list[str] = []
    seen: set[str] = set()
    sum_total = 0.0
    for pdf, r in results.items():
        name = Path(pdf).name[:46]
        if not isinstance(r, dict) or "timing" not in r:
            err = r.get("error", str(r)) if isinstance(r, dict) else str(r)
            lines.append(f"{name:46s}  ERROR: {err[:60]}")
            continue
        t = r["timing"]
        open_t = t.get("open", 0.0)
        tot = t.get("total", 0.0)
        sum_total += tot
        lines.append(f"{name:46s}  {open_t:6.3f}  {tot:8.3f}")
        for k in r.get("steps", {}):
            if k not in seen:
                step_set.append(k)
                seen.add(k)
    lines.append("-" * 78)
    if step_set:
        lines.append("Per-step (sum across all PDFs):")
        for step in step_set:
            tot = 0.0
            cnt = 0
            for r in results.values():
                if isinstance(r, dict) and "timing" in r:
                    v = r["timing"].get(f"step:{step}")
                    if v is not None:
                        tot += v
                        cnt += 1
            lines.append(f"  {step:46s}  {tot:7.3f}s  ({cnt} pdfs)")
    lines.append("-" * 78)
    ratio = (sum_total / wall) if wall > 0 else 0
    lines.append(
        f"Wall clock: {wall:.3f}s  |  serial-sum: {sum_total:.3f}s  "
        f"|  ratio = {ratio:.2f}x  |  N={len(results)} PDFs"
    )
    lines.append("=" * 78)
    return "\n".join(lines)


__all__ = [
    "parse_page_spec",
    "parse_ranges_spec",
    "safe_filename",
    "load_step",
    "list_builtin_steps",
    "is_builtin_step",
    "run_pipeline_single",
    "run_pipeline_parallel",
    "format_timing_table",
    "LoadedStep",
]
