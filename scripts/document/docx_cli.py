#!/usr/bin/env python3
"""docx_cli.py — 文档处理统一 CLI (Phase 1 colocate · script-consolidation GOAL)

合并 12 个旧 docx/md 处理脚本为单一 multi-subcommand 入口。dispatcher 模式：
每个子命令构造 argv 并调用旧脚本 `main()`，不复刻代码。

子命令:
  extract / check / snapshot / compare / track   ← docx_tools.py
  image-caption                                   ← sub/docx_apply_image_caption.py
  template                                        ← docx_fmt.py template（原 docx_apply_template.py）
  format                                          ← docx_fmt.py clone（原 docx_format_clone.py）
  renumber-fig                                    ← renum.py figures（原 docx_renumber_figures.py）
  text-fmt                                        ← docx_fmt.py text（原 docx_text_formatter.py）
  fix-ref                                         ← fix_superscript_refs.py
  md-to-docx                                      ← md_tools.py md2docx（原 md_docx_template.py）
  scan-sensitive                                  ← scan_sensitive_words.py
  md ...                                          ← md_tools.py (sub-group: format/merge/split/strip/to-docx/to-html/frontmatter)

（bullet / quality-check / review 2026-07-30 随 bullet_to_paragraph.py /
  report_quality_check.py / review_deep.py 退役 —— 全域零消费，
  见 ~/.Trash/doctools-orphans-20260730/MANIFEST.md）

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
# 别的机器上没有 ~/Dev —— 而缺 parallel_contract 在这里是 **fail-closed exit 2**，
# 不是降级，所以装出来的 wheel 会一条动词都跑不了。wheel 把这个总部 SSOT 镜像到了
# `<根>/lib/`（见 pyproject 的 force-include），故同源目录也纳入搜索。
# **append 不是 insert(0)**：`<根>/lib` 里有 styles.py / schemas.py / progress.py 这些
# 与他处同名的模块，把它顶到 sys.path 首位会改变全进程的解析优先级。这里只是要多一个
# 兜底来源，不是要提权。本机行为因此完全不变：`<repo>/lib` 没有 parallel_contract.py，
# 解析照旧落到上面的 _LIB。（形状对齐 scripts/data/data.py 既有先例。）
_BUNDLED_LIB = Path(__file__).resolve().parents[2] / "lib"
if str(_BUNDLED_LIB) not in sys.path:
    sys.path.append(str(_BUNDLED_LIB))
try:
    from parallel_contract import add_parallel_args, run_batch, parse_batch_jsonl  # type: ignore
except ImportError as e:  # pragma: no cover
    print(f"[docx_cli.py] FATAL: cannot import parallel_contract from {_LIB} / {_BUNDLED_LIB}: {e}", file=sys.stderr)
    sys.exit(2)

# ─── 旧脚本同目录定位 ────────────────────────────────────────────────────
_HERE = Path(__file__).resolve().parent


# ─── 版本号（SSOT = src/doctools/__init__.py） ──────────────────────────
def _pkg_version() -> str:
    """读 `src/doctools/__init__.py` 的 `__version__`。

    那一个字面量同时喂三处：`pyproject.toml`（hatchling `[tool.hatch.version]`
    动态读取）、装出来的 `doctools` 命令、和这里。**别在任何地方抄第二份。**

    走 spec_from_file_location 而不是 `import doctools`，是因为本脚本的主要跑法
    是 `python3 <abs>/docx_cli.py`（系统 python，没装过包），而把仓根塞进
    sys.path 会让 `tools/` `lib/` `config/` 这些平铺目录变成 namespace package
    去遮蔽同名模块 —— 为读一个字符串不值当。

    fail-closed：读不到就 ValueError，不打「unknown」糊弄过去。
    """
    init = _HERE.parents[1] / "src" / "doctools" / "__init__.py"
    if init.is_file():
        spec = importlib.util.spec_from_file_location("_doctools_version_probe", init)
        if spec is not None and spec.loader is not None:
            mod = importlib.util.module_from_spec(spec)
            spec.loader.exec_module(mod)
            v = getattr(mod, "__version__", None)
            if v:
                return str(v)
    # 先读文件、后读包元数据的顺序是有意的：editable 装法下 `uv sync` 不会因为
    # 版本号变了就重建 dist-info（实测改完 __version__ 再 sync，
    # `importlib.metadata.version("doctools")` 还停在旧值）。读源文件的这条路
    # 永远说真话，元数据只当源码树不在旁边时的兜底。
    try:  # 装成 wheel 后源码树可能不在旁边，退回包元数据（同一个 SSOT 派生的）
        from importlib.metadata import version as _md_version
        return str(_md_version("doctools"))
    except Exception as e:  # noqa: BLE001
        raise ValueError(f"版本号 SSOT 读不到（{init}）: {e}") from e


# ALL_DOCX_CMDS — for batch-all 编排（外部消费者读此常量）
ALL_DOCX_CMDS = [
    "extract", "check", "snapshot", "compare", "track",
    "image-caption", "template", "format",
    "renumber-fig", "text-fmt", "fix-ref",
    "md-to-docx", "scan-sensitive",
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
            # （2026-07-31 后本表脚本全都有 main()；runpy 分支留作无 main 脚本的兜底）
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

def cmd_image_caption(args: argparse.Namespace, rest: list[str]) -> int:
    return _exec_script("sub/docx_apply_image_caption", rest)

def cmd_template(args: argparse.Namespace, rest: list[str]) -> int:
    return _exec_script("docx_fmt", ["template"] + rest)

def cmd_format(args: argparse.Namespace, rest: list[str]) -> int:
    # 提取/复刻版式: format extract <ref> / format apply <content> --ref <ref>
    return _exec_script("docx_fmt", ["clone"] + rest)

def cmd_renumber_fig(args: argparse.Namespace, rest: list[str]) -> int:
    return _exec_script("renum", ["figures"] + rest)

def cmd_text_fmt(args: argparse.Namespace, rest: list[str]) -> int:
    # 2026-07-31 折进 docx_fmt.py text（有 def main() 了，不再走 runpy 分支）
    return _exec_script("docx_fmt", ["text"] + rest)

def cmd_fix_ref(args: argparse.Namespace, rest: list[str]) -> int:
    return _exec_script("sub/fix_superscript_refs", rest)

def cmd_md_to_docx(args: argparse.Namespace, rest: list[str]) -> int:
    return _exec_script("md_tools", ["md2docx"] + rest)

def cmd_scan_sensitive(args: argparse.Namespace, rest: list[str]) -> int:
    return _exec_script("sub/scan_sensitive_words", rest)

def cmd_md(args: argparse.Namespace, rest: list[str]) -> int:
    """md 子组：直接转发 md_tools.py 的 subcommand (format/merge/split/...)"""
    return _exec_script("md_tools", rest)


# ─── 职能轴（只读） ─────────────────────────────────────────────────────
_FN_AXIS: Any = None


def _function_axis() -> Any:
    """载入职能轴 SSOT（`sub/_function_axis.py`）。

    按文件路径直载、不走 `import sub` —— 只要一张纯数据表，不必把 12 个实现模块
    全拉起来。载不进来就让异常往上抛：`verbs` 的 `--fn` 取值域是从这张表派生的，
    静默兜底会让它变成一个「选项还在、取值域空了」的假绿。
    """
    global _FN_AXIS
    if _FN_AXIS is None:
        path = _HERE / "sub" / "_function_axis.py"
        spec = importlib.util.spec_from_file_location("_docx_cli_function_axis", str(path))
        if spec is None or spec.loader is None:
            raise ImportError(f"cannot spec {path}")
        mod = importlib.util.module_from_spec(spec)
        sys.modules["_docx_cli_function_axis"] = mod
        spec.loader.exec_module(mod)
        _FN_AXIS = mod
    return _FN_AXIS


def cmd_verbs(args: argparse.Namespace) -> int:
    """列出子命令及其职能（只读，不碰任何文件）。

    数据源是 `sub/_function_axis.py` 那张表，跟 CLI 的一致性由
    `tools/check_function_axis.py` 机检（表里缺一条或多一条都判红）。
    """
    axis = _function_axis()
    fn = getattr(args, "fn", None)
    rows = axis.rows(fn)
    if getattr(args, "json", False):
        print(json.dumps(
            {"total": len(rows),
             "fn_definitions": axis.FN_DEFINITIONS,
             "verbs": [{"path": axis.path_of(top, action), "top": top,
                        "action": action, "fn": f, "note": note}
                       for top, action, f, note in rows]},
            ensure_ascii=False, indent=1))
        return 0
    tags = [fn] if fn else list(axis.FN_TAGS)
    for tag in tags:
        group = [r for r in rows if r[2] == tag]
        print(f"\n## {tag} ({len(group)})  —— {axis.FN_DEFINITIONS[tag]}")
        for top, action, _f, note in group:
            tail = f"    # {note}" if note else ""
            print(f"  {axis.path_of(top, action)}{tail}")
    print(f"\n共 {len(rows)} 条"
          f"（别名 read=extract · diff=compare · styleset=audit-styleset 与被别名者"
          f"共用同一条职能，不单独列）")
    return 0


CMD_TABLE: dict[str, Callable[[argparse.Namespace, list[str]], int]] = {
    "extract": cmd_extract,
    "read": cmd_extract,      # alias: skill 叫 read → cli extract
    "check": cmd_check,
    "snapshot": cmd_snapshot,
    "compare": cmd_compare,
    "diff": cmd_compare,      # alias: skill 叫 diff → cli compare
    "track": cmd_track,
    "image-caption": cmd_image_caption,
    "template": cmd_template,
    "format": cmd_format,
    "renumber-fig": cmd_renumber_fig,
    "text-fmt": cmd_text_fmt,
    "fix-ref": cmd_fix_ref,
    "md-to-docx": cmd_md_to_docx,
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
            "doctools 文档处理统一 CLI (45 subcommands · 2026-07-30 计数核对)\n"
            "Legacy (13 旧族): extract / check / snapshot / compare / track /\n"
            "  image-caption / template / renumber-fig / text-fmt / fix-ref / md-to-docx /\n"
            "  scan-sensitive / md\n"
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
            "  docx_cli.py --version           # 版本号（装包后等价：doctools --version）\n"
            "\n详见 (script-consolidation GOAL 已随 goals/ 注册表 2026-07-10 退役):\n"
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
    # verbs — 只读：列出子命令及其职能（数据源 sub/_function_axis.py）
    # **载入失败不许拖垮整个 CLI**：这张表只服务 verbs 一条子命令，让它的异常上抛
    # 等于「表里一个拼写错误 → 全部 125 条子命令一起挂，还是个裸 traceback」。
    # 降级策略：--fn 的取值域退回不限，真正的报错留到 cmd_verbs 被调用时抛。
    try:
        fn_choices = list(_function_axis().FN_TAGS)
    except Exception as e:                      # noqa: BLE001 —— 什么错都不该拖垮 CLI
        print(f"\u26a0 \u804c\u80fd\u8f74\u8868\u8f7d\u5165\u5931\u8d25"
              f"\uff08\u53ea\u5f71\u54cd `docx_cli verbs`\uff09\uff1a{e}", file=sys.stderr)
        fn_choices = None
    vp = sub.add_parser(
        "verbs",
        help="列出子命令及其职能 (format/content/review/inspect/convert/dispatch)",
    )
    vp.add_argument("--fn", choices=fn_choices, help="只列该职能的子命令")
    vp.add_argument("--json", action="store_true", help="JSON 输出")
    vp.set_defaults(func=cmd_verbs)
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

    # 顶层 --version / -V
    # **故意不注册成 argparse 参数**：tools/cli_surface.py 的等价闸门是逐 action
    # 比对整棵子命令树，多一个 root action 就会改指纹；而「指纹逐字节不变」是这
    # 一轮的硬前提。只认第 0 位（那一位不是 subcommand 名就没有别的合法含义），
    # 所以不会和任何子命令自己的 --version 撞车。
    if raw and raw[0] in ("-V", "--version"):
        try:
            print(f"doctools {_pkg_version()}")
        except ValueError as e:
            print(f"[docx_cli.py] {e}", file=sys.stderr)
            return 2
        return 0

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
    # 顶层组名**从 parser 派生**，不手维护（2026-08-02 立）。
    #
    # 这里原来是一份手抄的 29 个名字的 set。它和 `_build_parser()` 里真正注册的
    # 49 个顶层名是两份需要人肉同步的东西，于是漏了 4 个：
    # `fix`(7 条) / `seqdiff`(2) / `compare-ref`(1) / `revise-rules`(1)。
    # 后果不是报错，是**静默无操作**：这 11 条动词（占 93 条的 12%）走到下面
    # 「未知顶层 token」分支 → top_p 是 add_help=False + parse_known_args 不报错 →
    # sub_cmd 仍是 None → 打印根 help、**rc=0、不写盘、不报错**。
    # `docx_cli.py fix clear-direct-format X.docx --inplace` 敲下去看着像跑了，
    # 其实连备份都没产生。而 check_function_axis(93 条) 与 cli_forward_probe(67 条)
    # 当时 rc 全 0 —— 闸门绿着，1/8 的动词是死的。
    #
    # 派生之后这类漏同步在结构上不可能再发生。机检：tools/check_verbs_reachable.py。
    DISTILLED_GROUPS = {
        name
        for a in parser._actions
        if isinstance(a, argparse._SubParsersAction)
        for name in a.choices
    }
    if not DISTILLED_GROUPS:            # fail-closed：一个都没派生出来 = 判据坏了
        print("[docx_cli.py] 内部错误：顶层子命令表为空", file=sys.stderr)
        return 2
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
        # 裸敲 `docx_cli.py` = 求助，打 help 返 0。
        # 但**给了个不认识的子命令**必须是错误 —— 原来这两种情况共用一个出口，
        # 于是 `docx_cli.py bogusxyz` 也是「打印 help + rc=0」，
        # 脚本里 `docx_cli.py <拼错的动词> && echo 成功` 会一路报成功（2026-08-02 实测）。
        unknown = [t for t in top_argv if not t.startswith("-")]
        if unknown:
            print(f"[docx_cli.py] 未知子命令: {unknown[0]!r}", file=sys.stderr)
            print(f"  可用: {', '.join(sorted(set(CMD_TABLE) | DISTILLED_GROUPS))}",
                  file=sys.stderr)
            return 2
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
