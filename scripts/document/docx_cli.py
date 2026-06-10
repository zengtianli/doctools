#!/usr/bin/env python3
"""docx_cli.py — 文档处理统一 CLI (Phase 1 colocate · script-consolidation GOAL)

合并 12 个旧 docx/md 处理脚本为单一 multi-subcommand 入口。dispatcher 模式：
每个子命令构造 argv 并调用旧脚本 `main()`，不复刻代码。

子命令 (15+):
  extract / check / snapshot / compare / track   ← docx_tools.py
  bullet                                          ← bullet_to_paragraph.py
  image-caption                                   ← docx_apply_image_caption.py
  template                                        ← docx_apply_template.py
  renumber-fig                                    ← docx_renumber_figures.py
  text-fmt                                        ← docx_text_formatter.py
  fix-ref                                         ← fix_superscript_refs.py
  md-to-docx                                      ← md_docx_template.py
  quality-check                                   ← report_quality_check.py
  review                                          ← review_deep.py
  scan-sensitive                                  ← scan_sensitive_words.py
  md ...                                          ← md_tools.py (sub-group: format/merge/split/strip/to-docx/to-html/frontmatter)

并行契约：消费 `parallel_contract.add_parallel_args` (--workers / --batch / --phases / --defer / --fanout-evidence)。
单文件交付走旧脚本；多文件 --batch 走 `parallel_contract.run_batch`。

旧脚本不删（thin alias 责由别 worker 生成）。

Why dispatcher 模式：
  - 12 脚本共 7133 行，复刻 = 揉巨大 boilerplate 违反铁律 #5
  - 旧脚本 `main()` 已稳定 + 含完整 argparse + 业务逻辑
  - dispatcher 只做 argv 转发 + 收尾，本文件 ~400 行
"""
from __future__ import annotations

import argparse
import importlib
import importlib.util
import json
import os
import sys
from contextlib import contextmanager
from pathlib import Path
from typing import Any, Callable, Optional

# ─── parallel_contract 兜底导入 ──────────────────────────────────────────
_LIB = Path.home() / "Dev" / "tools" / "dev" / "lib"
if str(_LIB) not in sys.path:
    sys.path.insert(0, str(_LIB))
try:
    from parallel_contract import add_parallel_args, run_batch, parse_batch_jsonl  # type: ignore
except ImportError as e:  # pragma: no cover
    print(f"[docx_cli.py] FATAL: cannot import parallel_contract from {_LIB}: {e}", file=sys.stderr)
    sys.exit(2)

# ─── 旧脚本同目录定位 ────────────────────────────────────────────────────
_HERE = Path(__file__).resolve().parent


# ALL_DOCX_CMDS — for batch-all 编排（外部消费者读此常量）
ALL_DOCX_CMDS = [
    "extract", "check", "snapshot", "compare", "track",
    "bullet", "image-caption", "template",
    "renumber-fig", "text-fmt", "fix-ref",
    "md-to-docx", "quality-check", "review", "scan-sensitive",
    "md",
]


# ─── 旧脚本加载（spec_from_file_location 路径直载） ─────────────────────
_LOADED_MODS: dict[str, Any] = {}


def _load_script_module(filename: str) -> Any:
    """按文件路径载入旧脚本 module；以别名注册防遮蔽 python-docx。"""
    if filename in _LOADED_MODS:
        return _LOADED_MODS[filename]
    path = _HERE / filename
    if not path.exists():
        raise FileNotFoundError(f"script not found: {path}")
    alias = f"_docx_dispatch__{filename.replace('.', '_')}"
    spec = importlib.util.spec_from_file_location(alias, str(path))
    if spec is None or spec.loader is None:
        raise ImportError(f"cannot spec {path}")
    mod = importlib.util.module_from_spec(spec)
    # 注册到 sys.modules 让 spec.loader 能解析旧脚本内 from X import Y（无）
    sys.modules[alias] = mod
    try:
        spec.loader.exec_module(mod)
    except Exception:
        sys.modules.pop(alias, None)
        raise
    _LOADED_MODS[filename] = mod
    return mod


def _exec_script(filename_stem: str, argv: list[str]) -> int:
    """执行 <filename_stem>.py 的 main()；支持无 main 的脚本（exec module 本身）。

    sys.argv 临时替换；SystemExit/异常兜底。
    """
    filename = filename_stem if filename_stem.endswith(".py") else f"{filename_stem}.py"
    saved_argv = sys.argv[:]
    saved_cwd = os.getcwd()
    sys.argv = [filename] + list(argv)
    try:
        try:
            mod = _load_script_module(filename)
        except SystemExit as se:
            # 加载时旧脚本顶层就退出（罕见；视为 1）
            return int(se.code) if isinstance(se.code, int) else 1
        except Exception as e:
            print(f"[docx_cli.py] load error {filename}: {type(e).__name__}: {e}", file=sys.stderr)
            return 2
        try:
            if hasattr(mod, "main"):
                rc = mod.main()
                if rc is None:
                    rc = 0
                return int(rc) if isinstance(rc, int) else 0
            # 无 main → 已在 exec_module 阶段执行了顶层逻辑（含 __main__ 块？否，
            # 因为 spec 模式 __name__ != "__main__"，所以无 main 脚本需 re-exec）
            # 对 docx_text_formatter.py 等：用 runpy 形式重跑（__name__ = '__main__'）
            import runpy
            runpy.run_path(str(_HERE / filename), run_name="__main__")
            return 0
        except SystemExit as se:
            return int(se.code) if isinstance(se.code, int) else (0 if se.code is None else 1)
        except Exception as e:
            print(f"[docx_cli.py] exec error in {filename}: {type(e).__name__}: {e}", file=sys.stderr)
            return 1
    finally:
        sys.argv = saved_argv
        try:
            os.chdir(saved_cwd)
        except Exception:
            pass


def _exec_script_file(filename: str, argv: list[str]) -> int:
    """兼容别名（旧 cmd_text_fmt 等）。"""
    return _exec_script(filename, argv)


# ─── dispatcher 子命令 ─────────────────────────────────────────────────
# 形式 1: 单段 → 旧脚本里也叫这个 subcommand
def cmd_extract(args: argparse.Namespace, rest: list[str]) -> int:
    return _exec_script("docx_tools", ["extract"] + rest)

def cmd_check(args: argparse.Namespace, rest: list[str]) -> int:
    return _exec_script("docx_tools", ["check"] + rest)

def cmd_snapshot(args: argparse.Namespace, rest: list[str]) -> int:
    return _exec_script("docx_tools", ["check", "snapshot"] + rest)

def cmd_compare(args: argparse.Namespace, rest: list[str]) -> int:
    return _exec_script("docx_tools", ["check", "compare"] + rest)

def cmd_track(args: argparse.Namespace, rest: list[str]) -> int:
    return _exec_script("docx_tools", ["track-changes"] + rest)

def cmd_bullet(args: argparse.Namespace, rest: list[str]) -> int:
    return _exec_script("bullet_to_paragraph", rest)

def cmd_image_caption(args: argparse.Namespace, rest: list[str]) -> int:
    return _exec_script("docx_apply_image_caption", rest)

def cmd_template(args: argparse.Namespace, rest: list[str]) -> int:
    return _exec_script("docx_apply_template", rest)

def cmd_renumber_fig(args: argparse.Namespace, rest: list[str]) -> int:
    return _exec_script("docx_renumber_figures", rest)

def cmd_text_fmt(args: argparse.Namespace, rest: list[str]) -> int:
    # docx_text_formatter.py 无 def main() → spec 执行
    return _exec_script_file("docx_text_formatter.py", rest)

def cmd_fix_ref(args: argparse.Namespace, rest: list[str]) -> int:
    return _exec_script("fix_superscript_refs", rest)

def cmd_md_to_docx(args: argparse.Namespace, rest: list[str]) -> int:
    return _exec_script("md_docx_template", rest)

def cmd_quality_check(args: argparse.Namespace, rest: list[str]) -> int:
    return _exec_script("report_quality_check", rest)

def cmd_review(args: argparse.Namespace, rest: list[str]) -> int:
    return _exec_script("review_deep", rest)

def cmd_scan_sensitive(args: argparse.Namespace, rest: list[str]) -> int:
    return _exec_script("scan_sensitive_words", rest)

def cmd_md(args: argparse.Namespace, rest: list[str]) -> int:
    """md 子组：直接转发 md_tools.py 的 subcommand (format/merge/split/...)"""
    return _exec_script("md_tools", rest)


CMD_TABLE: dict[str, Callable[[argparse.Namespace, list[str]], int]] = {
    "extract": cmd_extract,
    "read": cmd_extract,      # alias: skill 叫 read → cli extract
    "check": cmd_check,
    "snapshot": cmd_snapshot,
    "compare": cmd_compare,
    "diff": cmd_compare,      # alias: skill 叫 diff → cli compare
    "track": cmd_track,
    "bullet": cmd_bullet,
    "image-caption": cmd_image_caption,
    "template": cmd_template,
    "renumber-fig": cmd_renumber_fig,
    "text-fmt": cmd_text_fmt,
    "fix-ref": cmd_fix_ref,
    "md-to-docx": cmd_md_to_docx,
    "quality-check": cmd_quality_check,
    "review": cmd_review,
    "scan-sensitive": cmd_scan_sensitive,
    "md": cmd_md,
}


# ─── batch 模式 ────────────────────────────────────────────────────────
def _handle_batch(args: argparse.Namespace, sub_cmd: str, base_rest: list[str]) -> int:
    """--batch FILE.jsonl 形式：JSONL 每行 dict 含 'file' 或 'argv' 字段。

    单条 task schema (任选)：
      {"file": "/path/x.docx", "extra": ["-o", "out.md"]}
      {"argv": ["x.docx", "-o", "out.md"]}

    通过 run_batch 并发执行，evidence_path 走 fanout-evidence。
    """
    tasks = parse_batch_jsonl(args.batch)
    if not tasks:
        print("[docx_cli.py] batch jsonl empty", file=sys.stderr)
        return 0

    sub_fn = CMD_TABLE.get(sub_cmd)
    if sub_fn is None:
        print(f"[docx_cli.py] unknown subcommand for batch: {sub_cmd}", file=sys.stderr)
        return 2

    def handler(task: dict) -> dict:
        if "argv" in task:
            argv = list(task["argv"])
        elif "file" in task:
            argv = [str(task["file"])] + [str(x) for x in task.get("extra", [])]
        else:
            return {"ok": False, "error": f"task lacks 'file'/'argv' (ln {task.get('_ln_no')})"}
        argv = base_rest + argv
        rc = sub_fn(args, argv)
        return {"ok": rc == 0, "rc": rc, "argv": argv}

    rc, results = run_batch(
        tasks,
        handler,
        workers=args.workers,
        evidence_path=getattr(args, "fanout_evidence", None),
        progress=True,
    )
    ok = sum(1 for r in results if r.get("ok"))
    print(f"[docx_cli.py] batch done: {ok}/{len(results)} ok (rc={rc})", file=sys.stderr)
    return rc


# ─── 顶层 CLI ──────────────────────────────────────────────────────────
def _build_parser() -> argparse.ArgumentParser:
    p = argparse.ArgumentParser(
        prog="docx_cli",
        description=(
            "doctools 文档处理统一 CLI (36+ subcommands · 2026-05-26 W4 table delete-rows)\n"
            "Legacy (16 旧族): extract / check / snapshot / compare / track / bullet /\n"
            "  image-caption / template / renumber-fig / text-fmt / fix-ref / md-to-docx /\n"
            "  quality-check / review / scan-sensitive / md\n"
            "Distilled (15 新族 · sub/*.py): audit / freeze / strip / header-footer /\n"
            "  chapter / renumber / caption / blocks / outline / style / image / legacy /\n"
            "  seqdiff / compare-ref / revise-rules\n"
            "Pipeline (新): pipeline run <docx>... --steps <step,...> [--parallel] [--step-dir]"
        ),
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog=(
            "示例:\n"
            "  docx_cli.py extract input.docx -o out.md\n"
            "  docx_cli.py check snapshot a.docx\n"
            "  docx_cli.py md format -i x.md\n"
            "  docx_cli.py audit headings X.docx --report /tmp/h.json\n"
            "  docx_cli.py freeze headings X.docx\n"
            "  docx_cli.py style body X.docx --profile zdwp\n"
            "  docx_cli.py renumber h4-figures X.docx --profile eco-flow\n"
            "  docx_cli.py --batch tasks.jsonl --workers 8 extract\n"
            "\n详见 GOAL: ~/Dev/tools/cc-home/goals/script-consolidation/GOAL.md\n"
            "  hq_capabilities.yaml doctools.sub_capabilities (子命令清单)"
        ),
    )
    add_parallel_args(p)
    sub = p.add_subparsers(dest="command", metavar="<subcommand>")
    # Legacy 16 旧族 — REMAINDER 透传到旧脚本
    # alias map: skill 名 → cli 规范名（skill 叫 read/diff，cli 叫 extract/compare）
    _CMD_ALIASES: dict[str, list[str]] = {
        "extract": ["read"],
        "compare": ["diff"],
    }
    for name in ALL_DOCX_CMDS:
        aliases = _CMD_ALIASES.get(name, [])
        sp = sub.add_parser(
            name,
            aliases=aliases,
            help=f"→ 转发到旧脚本 ({name})" + (f" [alias: {','.join(aliases)}]" if aliases else ""),
            add_help=False,  # 让旧脚本自己处理 -h（透传 rest）
        )
        sp.add_argument("rest", nargs=argparse.REMAINDER, help="透传到旧脚本")
    # Distilled 11 新族 — sub/*.py 各自 register()
    _register_distilled_subcommands(sub)
    return p


def _register_distilled_subcommands(sub) -> None:
    """Load sub/ package (sibling dir) and register all distilled group modules.

    sub/ is `scripts/document/sub/__init__.py` — siblings of this file.
    We add its parent (scripts/document/) to sys.path then `import sub`.
    """
    here_parent = str(_HERE)
    inserted = False
    if here_parent not in sys.path:
        sys.path.insert(0, here_parent)
        inserted = True
    try:
        import importlib
        sub_pkg = importlib.import_module("sub")
        sub_pkg.register_all(sub)
    except Exception as e:  # pragma: no cover
        print(f"[docx_cli.py] WARN: failed to register sub/* modules: "
              f"{type(e).__name__}: {e}", file=sys.stderr)
    finally:
        if inserted:
            try:
                sys.path.remove(here_parent)
            except ValueError:
                pass


def main(argv: Optional[list[str]] = None) -> int:
    raw = list(argv) if argv is not None else sys.argv[1:]

    # 顶层 --help / -h (无 subcommand 时显示)
    if not raw or raw[0] in ("-h", "--help"):
        _build_parser().print_help()
        return 0

    # 手动分割：第一个非顶层-flag 即 subcommand，其余全透传
    TOP_FLAGS_WITH_VAL = {"--workers", "--batch", "--phases", "--defer", "--fanout-evidence"}
    parser = _build_parser()
    top_argv: list[str] = []
    sub_cmd: Optional[str] = None
    rest: list[str] = []
    i = 0
    # Distilled top-level groups (12+) — argparse handles their internals
    DISTILLED_GROUPS = {
        "audit", "audit-styleset", "styleset", "freeze", "strip", "header-footer", "chapter",
        "renumber", "caption", "blocks", "outline", "style", "image", "legacy",
        "pipeline", "health", "health-split",
        "section",    # section read/list (distilled from panan-rigid, 2026-05-26)
        "md-merge",   # merge MD content into DOCX section (distilled from panan-rigid, 2026-05-26)
        "md-merge-track",  # MD→track-changes 锚点前插 (上提 reclaim merge-tracked, GOAL 0-B 2026-05-29)
        "table",      # table structural ops: delete-rows (W4 distill, 2026-05-26)
        "split",      # split docx by-h1 (eco-flow/taizhou-天台 distill, W1 2026-05-26)
        "combine",    # combine N docx → 1 (docxcompose; inverse of split by-h1, 2026-06-07)
        "chapters-sync",  # 成品 docx 反向回写成品章节目录 (merge 的逆操作; govern 2026-06-08)
        "slim",       # docx-slim: safe ensemble + aggressive minimal skeleton (W 2026-05-28)
        "chrome",     # 院报告版面装帧: 逐章分节+逐章页眉页脚水印+宽表横向节 (eco-flow distill, 2026-06-04)
    }
    while i < len(raw):
        tok = raw[i]
        if sub_cmd is None:
            if tok in TOP_FLAGS_WITH_VAL:
                # 顶层 flag + 值
                top_argv.append(tok)
                if i + 1 < len(raw):
                    top_argv.append(raw[i + 1])
                    i += 2
                    continue
                i += 1
                continue
            if tok in CMD_TABLE or tok in DISTILLED_GROUPS:
                sub_cmd = tok
                i += 1
                continue
            # 未知顶层 token → 让 argparse 报错
            top_argv.append(tok)
            i += 1
        else:
            rest.append(tok)
            i += 1

    # 顶层 argparse 仅解析 top_argv（不含 subcommand 的 rest）
    # 为避免 subparsers required 报错，单独构造一个无 sub 的顶层 parser
    top_p = argparse.ArgumentParser(add_help=False)
    add_parallel_args(top_p)
    try:
        args, _unknown = top_p.parse_known_args(top_argv)
    except SystemExit:
        return 2

    if sub_cmd is None:
        parser.print_help()
        return 0

    # Legacy (CMD_TABLE) — fast-path REMAINDER dispatch
    if sub_cmd in CMD_TABLE:
        sub_fn = CMD_TABLE[sub_cmd]
        args.command = sub_cmd
        if getattr(args, "batch", None):
            return _handle_batch(args, sub_cmd, rest)
        return sub_fn(args, rest)

    # Distilled (sub/*.py) — full argparse path
    if sub_cmd in DISTILLED_GROUPS:
        try:
            full = parser.parse_args([sub_cmd] + rest)
        except SystemExit as se:
            return int(se.code) if isinstance(se.code, int) else 2
        func = getattr(full, "func", None)
        if func is None:
            print(f"[docx_cli.py] no handler for {sub_cmd} (incomplete subcommand?)",
                  file=sys.stderr)
            return 2
        try:
            rc = func(full)
            return int(rc) if isinstance(rc, int) else (0 if rc is None else 1)
        except SystemExit as se:
            return int(se.code) if isinstance(se.code, int) else 0

    print(f"[docx_cli.py] unknown subcommand: {sub_cmd}", file=sys.stderr)
    return 2


if __name__ == "__main__":
    sys.exit(main())
