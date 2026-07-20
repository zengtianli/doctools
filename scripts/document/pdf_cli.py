#!/usr/bin/env python3
"""pdf_cli.py — PDF 处理统一 CLI (2026-05-28 · 配 pdf_pipeline_lib.py)

8 + 1 子命令:
  read                    pdfplumber 提文本 / --list 显示 outline + pageinfo
  extract image           pdfimages -all
  extract text            pdfplumber 每页 txt (或 --single 合一)
  extract table           pdfplumber 每页 CSV
  split  by-bookmark      pypdf outline 切分
  split  by-page-range    pypdf 范围切分 (--ranges "1-10,11-20")
  merge                   pypdf PdfWriter 拼接
  decrypt                 qpdf 直解 (--password) / 无密走 cc-home pdf-decrypt skill
  pipeline run            <glob> --steps <names> [--parallel ...]

底座依赖:
  pdfplumber 0.11.9 / pypdf 6.12.2
  /opt/homebrew/bin/{pdfimages,pdftotext,pdfinfo,qpdf,ocrmypdf,tesseract}

设计同源 `docx_cli.py` 双层 dispatch,但所有子命令本仓 native (不转发旧脚本)。
pipeline 引擎在 `pdf_pipeline_lib.py`。
"""

from __future__ import annotations

import argparse
import glob as _glob
import json
import os
import subprocess
import sys
import time
from pathlib import Path
from typing import Optional

# ─── parallel_contract for --workers/--max-workers cohesion ────────────
_LIB = Path.home() / "Dev" / "tools" / "dev" / "lib"
if str(_LIB) not in sys.path:
    sys.path.insert(0, str(_LIB))
try:
    from parallel_contract import add_parallel_args  # type: ignore
except ImportError:
    def add_parallel_args(p):  # type: ignore  # noqa
        p.add_argument("--workers", type=int, default=None)

# pdf_pipeline_lib (sibling)
_HERE = Path(__file__).resolve().parent
if str(_HERE) not in sys.path:
    sys.path.insert(0, str(_HERE))
try:
    import pdf_pipeline_lib as ppl  # type: ignore
except ImportError as e:
    print(f"[pdf_cli.py] FATAL: cannot import pdf_pipeline_lib: {e}",
          file=sys.stderr)
    sys.exit(2)

# Lazy imports of pdfplumber / pypdf — heavy deps; only some subcommands need
try:
    import pdfplumber  # type: ignore
except ImportError:
    pdfplumber = None  # type: ignore

try:
    import pypdf  # type: ignore
    from pypdf import PdfReader, PdfWriter  # type: ignore
except ImportError:
    pypdf = None  # type: ignore
    PdfReader = PdfWriter = None  # type: ignore


# Canonical binary paths
_PDFINFO = "/opt/homebrew/bin/pdfinfo"
_PDFIMAGES = "/opt/homebrew/bin/pdfimages"
_PDFTOTEXT = "/opt/homebrew/bin/pdftotext"
_QPDF = "/opt/homebrew/bin/qpdf"
_DECRYPT_SKILL = (
    Path.home() / "Dev" / "tools" / "cc-home"
    / "skills" / "pdf-decrypt" / "scripts" / "decrypt.py"
)
_PYTHON3 = "/opt/homebrew/bin/python3"


# ═══════════════════════════════════════════════════════════════════════
# 1. read
# ═══════════════════════════════════════════════════════════════════════

def _cmd_read(args: argparse.Namespace) -> int:
    pdf_path = Path(args.pdf).resolve()
    if not pdf_path.is_file():
        print(f"ERROR: file not found: {pdf_path}", file=sys.stderr)
        return 2

    if args.list:
        # pdfinfo for page count / metadata
        cp = subprocess.run(
            [_PDFINFO, str(pdf_path)], capture_output=True, text=True
        )
        if cp.returncode == 0:
            print(cp.stdout.rstrip())
        else:
            print(f"[pdfinfo rc={cp.returncode}] {cp.stderr.strip()}",
                  file=sys.stderr)

        # pypdf outline tree
        if PdfReader is None:
            print("\n[outline] pypdf not installed", file=sys.stderr)
            return 0 if cp.returncode == 0 else cp.returncode
        try:
            reader = PdfReader(str(pdf_path))
            outline = reader.outline
        except Exception as e:
            print(f"\n[outline] error: {type(e).__name__}: {e}",
                  file=sys.stderr)
            return 0 if cp.returncode == 0 else cp.returncode

        print("\n── outline ──")
        if not outline:
            print("(no outline / bookmarks)")
            return 0 if cp.returncode == 0 else cp.returncode

        def _walk(node, depth=0):
            if isinstance(node, list):
                for item in node:
                    _walk(item, depth)
                return
            title = getattr(node, "title", str(node))
            try:
                page_num = reader.get_destination_page_number(node) + 1
            except Exception:
                page_num = "?"
            print(f"{'  ' * depth}- [{page_num}] {title}")

        _walk(outline)
        return 0

    # Default: extract text from page(s) to stdout
    if pdfplumber is None:
        print("ERROR: pdfplumber not installed", file=sys.stderr)
        return 2
    try:
        pdf = pdfplumber.open(str(pdf_path))
    except Exception as e:
        print(f"ERROR: pdfplumber.open failed: {type(e).__name__}: {e}",
              file=sys.stderr)
        return 1
    try:
        total = len(pdf.pages)
        if args.page is not None:
            spec = str(args.page)
        elif args.pages:
            spec = args.pages
        else:
            spec = None
        try:
            idx = ppl.parse_page_spec(spec, total)
        except ValueError as e:
            print(f"ERROR: {e}", file=sys.stderr)
            return 2
        for i in idx:
            try:
                txt = pdf.pages[i].extract_text() or ""
            except Exception as e:
                txt = f"[ERROR page {i+1}: {type(e).__name__}: {e}]"
            print(f"===== page {i+1} =====")
            print(txt)
        return 0
    finally:
        try:
            pdf.close()
        except Exception:
            pass


# ═══════════════════════════════════════════════════════════════════════
# 2. extract image / 3. extract text / 4. extract table
# ═══════════════════════════════════════════════════════════════════════

def _cmd_extract_image(args: argparse.Namespace) -> int:
    pdf_path = Path(args.pdf).resolve()
    if not pdf_path.is_file():
        print(f"ERROR: file not found: {pdf_path}", file=sys.stderr)
        return 2
    out_dir = Path(args.out_dir).resolve()
    out_dir.mkdir(parents=True, exist_ok=True)
    t0 = time.perf_counter()
    cmd = [_PDFIMAGES, "-all", str(pdf_path), str(out_dir / "img")]
    cp = subprocess.run(cmd, capture_output=True, text=True)
    elapsed = time.perf_counter() - t0
    n_files = sum(1 for p in out_dir.iterdir() if p.is_file())
    print(f"[extract image] rc={cp.returncode} files={n_files} "
          f"elapsed={elapsed:.3f}s out={out_dir}")
    if cp.returncode != 0:
        print(cp.stderr, file=sys.stderr)
    return cp.returncode


def _cmd_extract_text(args: argparse.Namespace) -> int:
    pdf_path = Path(args.pdf).resolve()
    if not pdf_path.is_file():
        print(f"ERROR: file not found: {pdf_path}", file=sys.stderr)
        return 2
    if pdfplumber is None:
        print("ERROR: pdfplumber not installed", file=sys.stderr)
        return 2
    out_dir = Path(args.out_dir).resolve()
    # Wrap into pipeline-style namespace for code reuse
    ns = argparse.Namespace(
        pdf=pdf_path,
        text_extract_pages=args.pages,
        text_extract_single=args.single,
        text_extract_out_dir=out_dir,
    )
    rep = ppl.run_pipeline_single(pdf_path, ["text-extract"], args=ns)
    sub = rep.get("steps", {}).get("text-extract", {})
    if "error" in sub:
        print(f"[extract text] ERROR: {sub['error']}", file=sys.stderr)
        return 1
    print(f"[extract text] mode={sub.get('mode')} "
          f"pages={sub.get('pages_requested')} "
          f"files={sub.get('files_written')} "
          f"elapsed={sub.get('elapsed_s')}s out={sub.get('out_dir')}")
    return 0


def _cmd_extract_table(args: argparse.Namespace) -> int:
    pdf_path = Path(args.pdf).resolve()
    if not pdf_path.is_file():
        print(f"ERROR: file not found: {pdf_path}", file=sys.stderr)
        return 2
    if pdfplumber is None:
        print("ERROR: pdfplumber not installed", file=sys.stderr)
        return 2
    out_dir = Path(args.out_dir).resolve()
    ns = argparse.Namespace(
        pdf=pdf_path,
        table_extract_pages=args.pages,
        table_extract_out_dir=out_dir,
    )
    rep = ppl.run_pipeline_single(pdf_path, ["table-extract"], args=ns)
    sub = rep.get("steps", {}).get("table-extract", {})
    if "error" in sub:
        print(f"[extract table] ERROR: {sub['error']}", file=sys.stderr)
        return 1
    print(f"[extract table] tables={sub.get('total_tables')} "
          f"pages_scanned={sub.get('pages_scanned')} "
          f"elapsed={sub.get('elapsed_s')}s out={sub.get('out_dir')}")
    counts = sub.get("per_page_table_counts", {})
    nonzero = {k: v for k, v in counts.items() if v}
    if nonzero:
        print(f"  per-page (nonzero): {nonzero}")
    return 0


# ═══════════════════════════════════════════════════════════════════════
# 5. split by-bookmark / 6. split by-page-range
# ═══════════════════════════════════════════════════════════════════════

def _flatten_top_outline(reader) -> list[tuple[str, int]]:
    """Return top-level [(title, page_num_0idx), ...] from pypdf outline."""
    outline = reader.outline
    out: list[tuple[str, int]] = []
    if not outline:
        return out
    for item in outline:
        if isinstance(item, list):
            # nested children of previous top item → skip (top-only)
            continue
        title = getattr(item, "title", None) or "untitled"
        try:
            pno = reader.get_destination_page_number(item)
        except Exception:
            continue
        out.append((title, pno))
    return out


def _cmd_split_by_bookmark(args: argparse.Namespace) -> int:
    pdf_path = Path(args.pdf).resolve()
    if not pdf_path.is_file():
        print(f"ERROR: file not found: {pdf_path}", file=sys.stderr)
        return 2
    if PdfReader is None or PdfWriter is None:
        print("ERROR: pypdf not installed", file=sys.stderr)
        return 2
    out_dir = Path(args.out_dir).resolve()
    out_dir.mkdir(parents=True, exist_ok=True)

    reader = PdfReader(str(pdf_path))
    items = _flatten_top_outline(reader)
    if not items:
        print("no outline found", file=sys.stderr)
        return 3

    total_pages = len(reader.pages)
    # Build (title, start, end) — end = next start - 1, last = total - 1
    segments: list[tuple[str, int, int]] = []
    for i, (title, start) in enumerate(items):
        end = (items[i + 1][1] - 1) if i + 1 < len(items) else (total_pages - 1)
        if end < start:
            end = start
        segments.append((title, start, end))

    n_written = 0
    pad = max(2, len(str(len(segments))))
    for idx, (title, start, end) in enumerate(segments, 1):
        writer = PdfWriter()
        for p in range(start, end + 1):
            writer.add_page(reader.pages[p])
        safe = ppl.safe_filename(title)
        out_file = out_dir / f"{str(idx).zfill(pad)}-{safe}.pdf"
        with out_file.open("wb") as f:
            writer.write(f)
        n_written += 1
        print(f"  [{idx}] pages {start+1}-{end+1} → {out_file.name}")
    print(f"[split by-bookmark] {n_written} parts → {out_dir}")
    return 0


def _cmd_split_by_page_range(args: argparse.Namespace) -> int:
    pdf_path = Path(args.pdf).resolve()
    if not pdf_path.is_file():
        print(f"ERROR: file not found: {pdf_path}", file=sys.stderr)
        return 2
    if PdfReader is None or PdfWriter is None:
        print("ERROR: pypdf not installed", file=sys.stderr)
        return 2
    out_dir = Path(args.out_dir).resolve()
    out_dir.mkdir(parents=True, exist_ok=True)
    try:
        ranges = ppl.parse_ranges_spec(args.ranges)
    except ValueError as e:
        print(f"ERROR: {e}", file=sys.stderr)
        return 2

    reader = PdfReader(str(pdf_path))
    total = len(reader.pages)
    pad = max(2, len(str(len(ranges))))
    for idx, (a, b) in enumerate(ranges, 1):
        if a > total:
            print(f"  [{idx}] skip {a}-{b} (file has {total} pages)",
                  file=sys.stderr)
            continue
        b_eff = min(b, total)
        writer = PdfWriter()
        for p in range(a - 1, b_eff):
            writer.add_page(reader.pages[p])
        out_file = out_dir / f"{str(idx).zfill(pad)}-pages-{a}-{b_eff}.pdf"
        with out_file.open("wb") as f:
            writer.write(f)
        print(f"  [{idx}] pages {a}-{b_eff} → {out_file.name}")
    print(f"[split by-page-range] {len(ranges)} ranges → {out_dir}")
    return 0


# ═══════════════════════════════════════════════════════════════════════
# 7. merge
# ═══════════════════════════════════════════════════════════════════════

def _cmd_merge(args: argparse.Namespace) -> int:
    if PdfWriter is None:
        print("ERROR: pypdf not installed", file=sys.stderr)
        return 2
    inputs = [Path(p).resolve() for p in args.pdfs]
    missing = [p for p in inputs if not p.is_file()]
    if missing:
        print(f"ERROR: files not found: {missing}", file=sys.stderr)
        return 2
    out = Path(args.out).resolve()
    out.parent.mkdir(parents=True, exist_ok=True)
    writer = PdfWriter()
    total_pages = 0
    for p in inputs:
        try:
            reader = PdfReader(str(p))
            for page in reader.pages:
                writer.add_page(page)
                total_pages += 1
            print(f"  + {p.name} ({len(reader.pages)} pages)")
        except Exception as e:
            print(f"ERROR: {p}: {type(e).__name__}: {e}", file=sys.stderr)
            return 1
    with out.open("wb") as f:
        writer.write(f)
    print(f"[merge] {len(inputs)} files, {total_pages} pages → {out}")
    return 0


# ═══════════════════════════════════════════════════════════════════════
# 8. decrypt
# ═══════════════════════════════════════════════════════════════════════

def _cmd_convert_to_docx(args: argparse.Namespace) -> int:
    """PDF → 可编辑 Word。引擎在 pdf_to_docx.py（结构提取路线，见该文件头注）。"""
    from pdf_to_docx import convert as _pdf2docx

    r = _pdf2docx(
        Path(args.pdf),
        Path(args.out) if args.out else None,
        with_images=args.images,
        with_tables=not args.no_tables,
        norm_width=not args.keep_fullwidth,
    )
    if not r["ok"]:
        print(f"ERROR: {r['error']}", file=sys.stderr)
        return 1
    bits = [f"{r['pages']}p", f"{r['paragraphs']} paras"]
    if r["tables"]:
        bits.append(f"{r['tables']} tables")
    if r["images"]:
        bits.append(f"{r['images']} images")
    print(f"✓ {r['output']}  ({', '.join(bits)})")
    return 0


def _cmd_decrypt(args: argparse.Namespace) -> int:
    pdf_path = Path(args.pdf).resolve()
    if not pdf_path.is_file():
        print(f"ERROR: file not found: {pdf_path}", file=sys.stderr)
        return 2
    if args.password:
        out = Path(args.out).resolve() if args.out else (
            pdf_path.with_name(pdf_path.stem + ".decrypted.pdf")
        )
        out.parent.mkdir(parents=True, exist_ok=True)
        cmd = [_QPDF, f"--password={args.password}", "--decrypt",
               str(pdf_path), str(out)]
        cp = subprocess.run(cmd, capture_output=True, text=True)
        if cp.returncode == 0:
            print(f"[decrypt] OK → {out}")
        else:
            print(f"[decrypt] qpdf rc={cp.returncode}: {cp.stderr.strip()}",
                  file=sys.stderr)
        return cp.returncode
    # No password: invoke the cc-home pdf-decrypt skill auto-guesser
    if not _DECRYPT_SKILL.is_file():
        print(f"ERROR: decrypt skill not found at {_DECRYPT_SKILL}",
              file=sys.stderr)
        return 2
    cmd = [_PYTHON3, str(_DECRYPT_SKILL), str(pdf_path)]
    if args.out:
        cmd.extend(["--out", str(Path(args.out).resolve())])
    cp = subprocess.run(cmd)
    return cp.returncode


# ═══════════════════════════════════════════════════════════════════════
# 9. pipeline run
# ═══════════════════════════════════════════════════════════════════════

def _parse_steps(s: str) -> list[str]:
    out = [t.strip() for t in s.split(",") if t.strip()]
    if not out:
        raise argparse.ArgumentTypeError("--steps cannot be empty")
    return out


def _resolve_glob(token: str) -> list[Path]:
    """Expand a glob token to existing PDF paths.

    Quoted glob from CLI arrives literal; we expand here.
    Also handles literal file paths.
    """
    p = Path(token)
    if p.is_file():
        return [p.resolve()]
    matches = _glob.glob(token, recursive=True)
    return [Path(m).resolve() for m in matches if Path(m).is_file()]


def _cmd_pipeline_run(args: argparse.Namespace) -> int:
    # Expand globs across all positional tokens
    pdfs: list[Path] = []
    for tok in args.pdfs:
        found = _resolve_glob(tok)
        if not found:
            print(f"[pipeline] WARN: no matches for {tok!r}", file=sys.stderr)
        pdfs.extend(found)
    # De-dup preserving order
    seen: set[str] = set()
    pdfs_dedup: list[Path] = []
    for p in pdfs:
        sp = str(p)
        if sp not in seen:
            seen.add(sp)
            pdfs_dedup.append(p)
    pdfs = pdfs_dedup
    if not pdfs:
        print("[pipeline] no PDFs matched after glob expansion",
              file=sys.stderr)
        return 2

    steps = args.steps
    unknown = [s for s in steps if not ppl.is_builtin_step(s)]
    if unknown:
        print(f"[pipeline] unknown steps: {unknown}; "
              f"known: {ppl.list_builtin_steps()}", file=sys.stderr)
        return 2

    print(f"[pipeline] {len(pdfs)} PDFs × {len(steps)} steps "
          f"({'parallel' if args.parallel else 'serial'} mode)")
    print(f"[pipeline] steps: {' → '.join(steps)}")
    for p in pdfs:
        print(f"  - {p.name}")

    # Build args_dict for child workers (only JSON-serializable scalars)
    args_dict = {
        "out_dir": str(args.out_dir) if args.out_dir else None,
        "image_extract_out_dir":
            str(args.image_extract_out_dir) if args.image_extract_out_dir
            else None,
        "text_extract_out_dir":
            str(args.text_extract_out_dir) if args.text_extract_out_dir
            else None,
        "text_extract_pages": args.text_extract_pages,
        "text_extract_single": args.text_extract_single,
        "table_extract_out_dir":
            str(args.table_extract_out_dir) if args.table_extract_out_dir
            else None,
        "table_extract_pages": args.table_extract_pages,
        "pages": args.pages,
        "single": args.single,
    }

    t0 = time.perf_counter()
    if not args.parallel or len(pdfs) == 1:
        results: dict[str, dict] = {}
        ns = argparse.Namespace(**args_dict)
        for p in pdfs:
            ns.pdf = p
            try:
                results[str(p)] = ppl.run_pipeline_single(p, steps, args=ns)
            except Exception as exc:
                results[str(p)] = {
                    "pdf": str(p),
                    "error": f"{type(exc).__name__}: {exc}",
                }
    else:
        results = ppl.run_pipeline_parallel(
            pdfs, steps,
            max_workers=args.max_workers,
            args_dict=args_dict,
        )
    wall = time.perf_counter() - t0

    print()
    print(ppl.format_timing_table(results, wall))

    # Per-PDF report JSON
    if args.report_dir:
        rd = Path(args.report_dir).resolve()
        rd.mkdir(parents=True, exist_ok=True)
        for pdf, rep in results.items():
            stem = Path(pdf).stem
            out = rd / f"pipeline-{stem}.json"
            out.write_text(
                json.dumps(rep, ensure_ascii=False, indent=2, default=str),
                encoding="utf-8",
            )
        print(f"[pipeline] reports → {rd}")

    any_err = any(
        isinstance(r, dict) and ("error" in r) for r in results.values()
    )
    # Also check step-level errors
    if not any_err:
        for r in results.values():
            if isinstance(r, dict):
                for sub in r.get("steps", {}).values():
                    if isinstance(sub, dict) and "error" in sub:
                        any_err = True
                        break
    return 1 if any_err else 0


# ═══════════════════════════════════════════════════════════════════════
# Argparse wiring
# ═══════════════════════════════════════════════════════════════════════

def _build_parser() -> argparse.ArgumentParser:
    p = argparse.ArgumentParser(
        prog="pdf_cli",
        description=(
            "PDF unified CLI (2026-05-28)\n"
            "Subcommands: read / extract image|text|table / "
            "split by-bookmark|by-page-range / merge / decrypt / pipeline run"
        ),
        formatter_class=argparse.RawDescriptionHelpFormatter,
    )
    add_parallel_args(p)
    sub = p.add_subparsers(dest="command", metavar="<subcommand>")

    # read
    rp = sub.add_parser("read", help="extract page text / list outline")
    rp.add_argument("pdf")
    g = rp.add_mutually_exclusive_group()
    g.add_argument("--page", type=int, help="single page number (1-indexed)")
    g.add_argument("--pages", help="page spec, e.g. '1-5,8'")
    g.add_argument("--list", action="store_true",
                   help="pdfinfo + outline tree")
    rp.set_defaults(func=_cmd_read)

    # extract <kind>
    ep = sub.add_parser("extract", help="extract image / text / table")
    ep_sub = ep.add_subparsers(dest="extract_kind", metavar="<kind>")
    ep.set_defaults(func=lambda a: (ep.print_help() or 0))

    # extract image
    eip = ep_sub.add_parser("image", help="pdfimages -all → out-dir/img-*")
    eip.add_argument("pdf")
    eip.add_argument("--out-dir", required=True)
    eip.set_defaults(func=_cmd_extract_image)

    # extract text
    etp = ep_sub.add_parser("text", help="pdfplumber per-page txt")
    etp.add_argument("pdf")
    etp.add_argument("--out-dir", required=True)
    etp.add_argument("--single", action="store_true",
                     help="concat all into full.txt")
    etp.add_argument("--pages", help="page spec, e.g. '1-5'")
    etp.set_defaults(func=_cmd_extract_text)

    # extract table
    ettp = ep_sub.add_parser("table", help="pdfplumber per-page CSV tables")
    ettp.add_argument("pdf")
    ettp.add_argument("--out-dir", required=True)
    ettp.add_argument("--pages", help="page spec, e.g. '1-5'")
    ettp.set_defaults(func=_cmd_extract_table)

    # split <kind>
    sp = sub.add_parser("split", help="split PDF by bookmark / page-range")
    sp_sub = sp.add_subparsers(dest="split_kind", metavar="<kind>")
    sp.set_defaults(func=lambda a: (sp.print_help() or 0))

    # split by-bookmark
    sbb = sp_sub.add_parser("by-bookmark",
                            help="split by top-level outline entries")
    sbb.add_argument("pdf")
    sbb.add_argument("--out-dir", required=True)
    sbb.set_defaults(func=_cmd_split_by_bookmark)

    # split by-page-range
    sbpr = sp_sub.add_parser("by-page-range",
                             help='split by --ranges "1-10,11-20"')
    sbpr.add_argument("pdf")
    sbpr.add_argument("--ranges", required=True,
                      help='comma-separated A-B ranges (1-indexed)')
    sbpr.add_argument("--out-dir", required=True)
    sbpr.set_defaults(func=_cmd_split_by_page_range)

    # merge
    mp = sub.add_parser("merge", help="concatenate multiple PDFs")
    mp.add_argument("pdfs", nargs="+", help="input PDFs in order")
    mp.add_argument("--out", required=True, help="output combined PDF")
    mp.set_defaults(func=_cmd_merge)

    # decrypt
    dp = sub.add_parser("decrypt", help="decrypt PDF via qpdf / auto-guess")
    dp.add_argument("pdf")
    dp.add_argument("--password", default=None,
                    help="explicit password (qpdf direct)")
    dp.add_argument("--out", default=None, help="output path")
    dp.set_defaults(func=_cmd_decrypt)

    # convert group
    cp = sub.add_parser("convert", help="convert PDF to other formats")
    cp_sub = cp.add_subparsers(dest="convert_cmd", metavar="<cmd>")
    cp.set_defaults(func=lambda a: (cp.print_help() or 0))

    ctd = cp_sub.add_parser("to-docx", help="PDF → editable Word (structure extraction)")
    ctd.add_argument("pdf")
    ctd.add_argument("-o", "--out", default=None, help="output .docx (default: alongside)")
    ctd.add_argument("--images", action="store_true", help="append embedded images")
    ctd.add_argument("--no-tables", action="store_true", help="skip table detection")
    ctd.add_argument("--keep-fullwidth", action="store_true",
                     help="keep full-width latin/digits (default: convert to half-width)")
    ctd.set_defaults(func=_cmd_convert_to_docx)

    # pipeline group
    pp = sub.add_parser("pipeline", help="multi-step + multi-PDF pipeline")
    pp_sub = pp.add_subparsers(dest="pipeline_cmd", metavar="<cmd>")
    pp.set_defaults(func=lambda a: (pp.print_help() or 0))

    pprun = pp_sub.add_parser(
        "run",
        help="run --steps across PDFs",
        description=(
            "Examples:\n"
            "  pdf_cli.py pipeline run report.pdf "
            "--steps text-extract,table-extract --report-dir reports/\n"
            "  pdf_cli.py pipeline run 'data/*.pdf' "
            "--steps image-extract --parallel --max-workers 4"
        ),
        formatter_class=argparse.RawDescriptionHelpFormatter,
    )
    pprun.add_argument("pdfs", nargs="+",
                       help="PDF paths or globs (quote globs to defer "
                            "shell expansion)")
    pprun.add_argument("--steps", type=_parse_steps, required=True,
                       help="comma-separated step names: "
                            + ",".join(ppl.list_builtin_steps()))
    pprun.add_argument("--parallel", action="store_true",
                       help="cross-PDF process pool")
    pprun.add_argument("--max-workers", type=int, default=None,
                       help="parallel worker count "
                            "(default: min(N_pdf, cpu_count))")
    pprun.add_argument("--report-dir", default=None,
                       help="dir to write pipeline-<stem>.json per PDF")
    pprun.add_argument("--out-dir", default=None,
                       help="global out-dir root; "
                            "step uses <root>/<step-name>/")
    # per-step out_dir overrides
    pprun.add_argument("--image-extract-out-dir", default=None)
    pprun.add_argument("--text-extract-out-dir", default=None)
    pprun.add_argument("--table-extract-out-dir", default=None)
    # step-shared page selection
    pprun.add_argument("--pages", default=None,
                       help="page spec applied to all pdf-based steps "
                            "(text/table)")
    pprun.add_argument("--single", action="store_true",
                       help="text-extract: single full.txt")
    pprun.add_argument("--text-extract-pages", default=None)
    pprun.add_argument("--text-extract-single", action="store_true")
    pprun.add_argument("--table-extract-pages", default=None)
    pprun.set_defaults(func=_cmd_pipeline_run)

    return p


def main(argv: Optional[list[str]] = None) -> int:
    raw = list(argv) if argv is not None else sys.argv[1:]
    if not raw or raw[0] in ("-h", "--help"):
        _build_parser().print_help()
        return 0
    parser = _build_parser()
    try:
        args = parser.parse_args(raw)
    except SystemExit as se:
        return int(se.code) if isinstance(se.code, int) else 2

    func = getattr(args, "func", None)
    if func is None:
        parser.print_help()
        return 0
    try:
        rc = func(args)
        return int(rc) if isinstance(rc, int) else (0 if rc is None else 1)
    except SystemExit as se:
        return int(se.code) if isinstance(se.code, int) else 0
    except Exception as e:
        print(f"[pdf_cli.py] error: {type(e).__name__}: {e}", file=sys.stderr)
        return 1


if __name__ == "__main__":
    sys.exit(main())
