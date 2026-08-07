#!/usr/bin/env python3
"""
Word 文档工具集 (docx_tools.py)

合并三个 docx 工具为统一入口：

子命令：
  extract        — 从 .docx 提取纯文本（Markdown 格式），支持分章节输出
  check          — 两层格式检查（ZIP 完整性 + 格式语义）
  track-changes  — 读取/写入修订标记和批注

用法：
  python3 docx_tools.py extract input.docx                     # 全文输出到 stdout
  python3 docx_tools.py extract input.docx -o output.md        # 全文输出到文件
  python3 docx_tools.py extract input.docx --split-chapters    # 按章节拆分输出
  python3 docx_tools.py extract input.docx --info              # 仅输出文档结构信息

  python3 docx_tools.py check snapshot input.docx              # 输出格式报告
  python3 docx_tools.py check snapshot input.docx -o snap.json # 存为 JSON 快照
  python3 docx_tools.py check compare  before.docx after.docx  # 对比两个文件

  python3 docx_tools.py track-changes read input.docx [--format md|json]
  python3 docx_tools.py track-changes review input.docx -o output.docx --rules rules.json

2026-07-31 P2 拆薄：三段实现移至 sub/docx_extract.py / sub/docx_check.py /
sub/docx_track.py（各自带 main() 可独立敲，argparse 声明在各自 add_*_parser() 只写一遍）。
本文件保留：batch 并行层 + 组合 argparse 入口 + library re-export（外部会话实证有人
`from docx_tools import extract_paragraphs` 当库用，名字面必须原样保留）。
"""

import argparse
import json
import os
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).parent.parent.parent / "lib"))

# ── surgical 收口：python-docx 存盘只重写点名的部件（炸开面 60→1）─────────────
import sys as _sys
from pathlib import Path as _Path
_sys.path.append(str(_Path(__file__).resolve().parents[2] / "lib"))
import docx_safe_save  # noqa: E402,F401  详见 lib/docx_safe_save.py

# ── 实现模块直载（不 `import sub`：包 __init__ 会连带加载 12 个业务模块）─────────
import importlib.util as _ilu

_SUB_DIR = Path(__file__).resolve().parent / "sub"


def _load_impl(name: str):
    """spec 直载 sub/<name>.py（同 docx_cli._exec_script 范式），注册进 sys.modules 防重复加载。"""
    mod = _sys.modules.get(name)
    if mod is not None:
        return mod
    spec = _ilu.spec_from_file_location(name, _SUB_DIR / f"{name}.py")
    mod = _ilu.module_from_spec(spec)
    _sys.modules[name] = mod
    spec.loader.exec_module(mod)
    return mod


_extract_impl = _load_impl("docx_extract")
_check_impl = _load_impl("docx_check")
_track_impl = _load_impl("docx_track")

# ── library re-export：旧单体的模块级名字面，外部 import 契约，只增不减 ──────────
# extract 段
extract_paragraphs = _extract_impl.extract_paragraphs
paragraphs_to_markdown = _extract_impl.paragraphs_to_markdown
split_by_chapters = _extract_impl.split_by_chapters
document_info = _extract_impl.document_info
cmd_extract = _extract_impl.cmd_extract
# check 段
VML_NS = _check_impl.VML_NS
twips_to_cm = _check_impl.twips_to_cm
half_pt = _check_impl.half_pt
zip_hashes = _check_impl.zip_hashes
EXPECTED_CHANGES = _check_impl.EXPECTED_CHANGES
SAFE_SIDE_EFFECTS = _check_impl.SAFE_SIDE_EFFECTS
compare_zip_integrity = _check_impl.compare_zip_integrity
_extract_hf_text = _check_impl._extract_hf_text
extract_format_snapshot = _check_impl.extract_format_snapshot
ALIGN_MAP = _check_impl.ALIGN_MAP
_format_style_row = _check_impl._format_style_row
format_report = _check_impl.format_report
compare_report = _check_impl.compare_report
cmd_check = _check_impl.cmd_check
# track-changes 段
read_track_changes = _track_impl.read_track_changes
_tc_extract_text = _track_impl._tc_extract_text
_tc_extract_del_text = _track_impl._tc_extract_del_text
DocxReviewer = _track_impl.DocxReviewer
review_docx = _track_impl.review_docx
compare_docx = _track_impl.compare_docx      # 2026-08-03：compare 落地（原 v2 空桩）
DocxAccepter = _track_impl.DocxAccepter      # 2026-08-07：accept 落地（中间稿→成稿）
accept_docx = _track_impl.accept_docx
cmd_track_changes = _track_impl.cmd_track_changes


# ══════════════════════════════════════════════════════════════════════
#  Batch / 并行 API
# ══════════════════════════════════════════════════════════════════════
#
# JSONL 行格式：
#   {"file": "/path/a.docx", "subcommand": "extract", "options": {"json": true}}
#   {"file": "/path/b.docx", "subcommand": "check", "options": {"check_command": "snapshot", "md": true}}
#   {"file": "/path/c.docx", "subcommand": "track-changes", "options": {"tc_command": "read", "format": "json"}}
#   {"file": "/path/d.docx", "subcommand": "snapshot", "options": {"md": true}}                  # alias: check snapshot
#   {"file": "/path/d.docx", "subcommand": "compare", "options": {"after": "/path/e.docx"}}      # alias: check compare
#
# 阶段（phase）：
#   extract / snapshot / track-changes-read  → IO 重，可高并发
#   check-compare / track-changes-review     → 较重，但仍 IO bound
#   --defer PHASE 跳过指定阶段；--phases 仅运行指定阶段（逗号分隔）。


class _BatchArgs:
    """轻量 Namespace，把 JSONL row 的 options 字典套进 cmd_* 期望的 args 形态。"""

    def __init__(self, **kwargs):
        for k, v in kwargs.items():
            setattr(self, k, v)

    def __getattr__(self, name):
        return None


_PHASE_MAP = {
    "extract": "extract",
    "check-snapshot": "snapshot",
    "snapshot": "snapshot",
    "check-compare": "compare",
    "compare": "compare",
    "track-changes-read": "tc-read",
    "track-changes-review": "tc-review",
}


def _row_phase(subcommand: str, options: dict) -> str:
    if subcommand == "extract":
        return "extract"
    if subcommand == "check":
        return options.get("check_command", "snapshot")
    if subcommand == "snapshot":
        return "snapshot"
    if subcommand == "compare":
        return "compare"
    if subcommand == "track-changes":
        tc = options.get("tc_command", "read")
        return f"tc-{tc}"
    return subcommand


def _run_one(row: dict) -> dict:
    """执行单条 batch row，返回 {file, subcommand, ok, error, stdout_lines}."""
    import io
    from contextlib import redirect_stdout

    file_path = row.get("file") or row.get("input")
    subcommand = row["subcommand"]
    options = dict(row.get("options") or {})

    buf = io.StringIO()
    try:
        with redirect_stdout(buf):
            if subcommand == "extract":
                args = _BatchArgs(
                    input=file_path,
                    output=options.get("output"),
                    split_chapters=options.get("split_chapters", False),
                    info=options.get("info", False),
                    json=options.get("json", False),
                )
                cmd_extract(args)
            elif subcommand in ("snapshot", "compare"):
                # 顶层 alias → check.<subcmd>
                args = _BatchArgs(
                    check_command=subcommand,
                    input=file_path,
                    before=options.get("before", file_path),
                    after=options.get("after"),
                    output=options.get("output"),
                    md=options.get("md", False),
                )
                cmd_check(args)
            elif subcommand == "check":
                args = _BatchArgs(
                    check_command=options.get("check_command", "snapshot"),
                    input=file_path,
                    before=options.get("before", file_path),
                    after=options.get("after"),
                    output=options.get("output"),
                    md=options.get("md", False),
                )
                cmd_check(args)
            elif subcommand == "track-changes":
                args = _BatchArgs(
                    tc_command=options.get("tc_command", "read"),
                    input=file_path,
                    format=options.get("format", "md"),
                    output=options.get("output"),
                    rules=options.get("rules"),
                    author=options.get("author", "CC审阅"),
                    # compare 用（2026-08-03）：不传这两个，compare 分支只能报「缺参数」
                    original=options.get("original", file_path),
                    modified=options.get("modified"),
                )
                cmd_track_changes(args)
            else:
                raise ValueError(f"未知 subcommand: {subcommand}")
        return {
            "file": file_path,
            "subcommand": subcommand,
            "phase": _row_phase(subcommand, options),
            "ok": True,
            "stdout": buf.getvalue(),
        }
    except SystemExit as e:
        return {
            "file": file_path,
            "subcommand": subcommand,
            "phase": _row_phase(subcommand, options),
            "ok": (e.code in (0, None)),
            "error": None if e.code in (0, None) else f"SystemExit({e.code})",
            "stdout": buf.getvalue(),
        }
    except Exception as e:
        return {
            "file": file_path,
            "subcommand": subcommand,
            "phase": _row_phase(subcommand, options),
            "ok": False,
            "error": f"{type(e).__name__}: {e}",
            "stdout": buf.getvalue(),
        }


def run_batch(jsonl_path: str, workers: int | None = None,
              defer: list[str] | None = None, phases: list[str] | None = None) -> dict:
    """并行执行 JSONL batch。

    ThreadPoolExecutor（python-docx / zipfile / lxml 均为 IO 重 + Python C 扩展，
    GIL 期间释放，线程并行有效）。
    """
    from concurrent.futures import ThreadPoolExecutor, as_completed

    if workers is None or workers <= 0:
        workers = min(os.cpu_count() or 4, 8)

    defer_set = set(defer or [])
    phases_set = set(phases or [])

    rows = []
    skipped = []
    with open(jsonl_path, encoding="utf-8") as f:
        for ln, line in enumerate(f, 1):
            line = line.strip()
            if not line or line.startswith("#"):
                continue
            try:
                row = json.loads(line)
            except json.JSONDecodeError as e:
                skipped.append({"line": ln, "error": f"JSON 解析失败: {e}"})
                continue
            ph = _row_phase(row.get("subcommand", ""), row.get("options") or {})
            if defer_set and ph in defer_set:
                skipped.append({"line": ln, "phase": ph, "reason": "defer"})
                continue
            if phases_set and ph not in phases_set:
                skipped.append({"line": ln, "phase": ph, "reason": "not-in-phases"})
                continue
            rows.append(row)

    results = []
    with ThreadPoolExecutor(max_workers=workers) as ex:
        futs = {ex.submit(_run_one, r): r for r in rows}
        for fut in as_completed(futs):
            results.append(fut.result())

    summary = {
        "total": len(rows),
        "ok": sum(1 for r in results if r["ok"]),
        "failed": sum(1 for r in results if not r["ok"]),
        "skipped": len(skipped),
        "workers": workers,
        "results": results,
        "skipped_detail": skipped,
    }
    return summary


# ══════════════════════════════════════════════════════════════════════
#  CLI 入口
# ══════════════════════════════════════════════════════════════════════


def main():
    parser = argparse.ArgumentParser(
        description="Word 文档工具集",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog=__doc__,
    )

    # ── 批处理 / 并行 API（顶层 flag，优先于 subcommand）──
    parser.add_argument(
        "--batch",
        metavar="FILE",
        help="JSONL 批处理文件，每行 {file, subcommand, options} 并行调度",
    )
    parser.add_argument(
        "--workers",
        type=int,
        default=0,
        help="并发 worker 数（默认 min(cpu, 8)；仅 --batch 模式生效）",
    )
    parser.add_argument(
        "--defer",
        action="append",
        default=[],
        metavar="PHASE",
        help="跳过指定阶段（extract/snapshot/compare/tc-read/tc-review），可重复",
    )
    parser.add_argument(
        "--phases",
        default="",
        help="仅运行指定阶段（逗号分隔，e.g. 'extract,snapshot'）",
    )
    parser.add_argument(
        "--batch-json",
        action="store_true",
        help="--batch 完成后输出 JSON 汇总到 stdout（默认人类可读）",
    )

    sub = parser.add_subparsers(dest="command", help="子命令")

    # 三段的 argparse 声明在各实现模块里只写一遍（契约面禁漂移）
    _extract_impl.add_extract_parser(sub)
    _check_impl.add_check_parser(sub)
    _track_impl.add_track_parser(sub)

    args = parser.parse_args()

    # ── 批处理优先 ──
    if args.batch:
        phases = [p.strip() for p in args.phases.split(",") if p.strip()] if args.phases else None
        summary = run_batch(
            args.batch,
            workers=args.workers or None,
            defer=args.defer or None,
            phases=phases,
        )
        if args.batch_json:
            print(json.dumps(summary, ensure_ascii=False, indent=2))
        else:
            print(
                f"[batch] total={summary['total']} ok={summary['ok']} "
                f"failed={summary['failed']} skipped={summary['skipped']} "
                f"workers={summary['workers']}"
            )
            for r in summary["results"]:
                tag = "OK" if r["ok"] else "FAIL"
                err = f" :: {r.get('error')}" if not r["ok"] else ""
                print(f"  [{tag}] {r['phase']:<10} {r['file']}{err}")
        sys.exit(0 if summary["failed"] == 0 else 1)

    if args.command == "extract":
        cmd_extract(args)
    elif args.command == "check":
        cmd_check(args)
    elif args.command == "track-changes":
        cmd_track_changes(args)
    else:
        parser.print_help()


if __name__ == "__main__":
    main()
