"""_dispatch.py — shared helper to exec a sub/*.py script's main() with argv.

Used by group modules (audit.py, freeze.py, strip.py, etc.) to forward
subcommand args to the standalone script's main() without re-implementing
business logic. Mirrors docx_cli.py's `_exec_script` pattern.
"""
from __future__ import annotations

import importlib.util
import os
import runpy
import sys
from pathlib import Path
from typing import Any

_HERE = Path(__file__).resolve().parent
_LOADED: dict[str, Any] = {}


def _load(filename: str) -> Any:
    if filename in _LOADED:
        return _LOADED[filename]
    path = _HERE / filename
    if not path.exists():
        raise FileNotFoundError(f"script not found: {path}")
    alias = f"_doctools_sub__{filename.replace('.', '_')}"
    spec = importlib.util.spec_from_file_location(alias, str(path))
    if spec is None or spec.loader is None:
        raise ImportError(f"cannot spec {path}")
    mod = importlib.util.module_from_spec(spec)
    sys.modules[alias] = mod
    try:
        spec.loader.exec_module(mod)
    except Exception:
        sys.modules.pop(alias, None)
        raise
    _LOADED[filename] = mod
    return mod


def exec_script(filename_stem: str, argv: list[str]) -> int:
    """Execute sub/<filename_stem>.py main() with given argv. Returns int rc."""
    filename = filename_stem if filename_stem.endswith(".py") else f"{filename_stem}.py"
    saved_argv = sys.argv[:]
    saved_cwd = os.getcwd()
    sys.argv = [filename] + list(argv)
    try:
        try:
            mod = _load(filename)
        except SystemExit as se:
            return int(se.code) if isinstance(se.code, int) else 1
        except Exception as e:
            print(f"[sub._dispatch] load error {filename}: {type(e).__name__}: {e}", file=sys.stderr)
            return 2
        try:
            if hasattr(mod, "main"):
                rc = mod.main()
                if rc is None:
                    rc = 0
                return int(rc) if isinstance(rc, int) else 0
            runpy.run_path(str(_HERE / filename), run_name="__main__")
            return 0
        except SystemExit as se:
            return int(se.code) if isinstance(se.code, int) else (0 if se.code is None else 1)
        except Exception as e:
            print(f"[sub._dispatch] exec error in {filename}: {type(e).__name__}: {e}", file=sys.stderr)
            return 1
    finally:
        sys.argv = saved_argv
        try:
            os.chdir(saved_cwd)
        except Exception:
            pass


def get_or_add_group(subparsers, name: str, help_text: str = ""):
    """Return existing top-level group parser if already registered, else add new.

    Why: multiple distill batches (W1/W2/W3) each register `caption` / `chapter` /
    `renumber` as their own group -> argparse raises ArgumentError on duplicate.
    This helper lets each module's register() share the same group parent so
    targets from different sources coexist (e.g. `caption number` from W1,
    `caption pair` from W3, `caption number-by-style` from W2).
    """
    # argparse subparsers actions expose `choices` (name -> parser map)
    existing = getattr(subparsers, "choices", None) or getattr(subparsers, "_name_parser_map", {})
    if name in existing:
        return existing[name]
    return subparsers.add_parser(name, help=help_text)


def get_or_add_subparsers(group_parser, dest: str, metavar: str = "<target>", required: bool = True):
    """Return existing sub-subparsers action on a group parser, else add new.

    Why: when get_or_add_group returns an existing parser, calling
    `parser.add_subparsers()` again raises ArgumentError. We introspect
    `_actions` to detect and reuse.
    """
    for act in group_parser._actions:
        if act.__class__.__name__ == "_SubParsersAction":
            return act
    return group_parser.add_subparsers(dest=dest, metavar=metavar, required=required)


def _rest_argv(args) -> list[str]:
    """Extract argparse Namespace -> argv list for forwarding to standalone script."""
    argv: list[str] = []
    if getattr(args, "docx_path", None):
        argv.append(str(args.docx_path))
    if getattr(args, "dry_run", False):
        argv.append("--dry-run")
    if getattr(args, "no_backup", False):
        argv.append("--no-backup")
    if getattr(args, "report", None):
        argv.extend(["--report", str(args.report)])
    # any extra raw rest
    extra = getattr(args, "rest", None) or []
    argv.extend(str(x) for x in extra)
    return argv
