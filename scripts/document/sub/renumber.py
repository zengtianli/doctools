"""renumber.py — group module: renumber headings / table-figure captions (1 subcommand)

Subcommands:
  renumber headings   ← renumber_headings.py
                          re-number H1/H2/H3 + table captions (表X-Y) based on physical
                          paragraph order. Run-level prefix replacement preserves
                          bold/font-size. Title style untouched.
"""
from __future__ import annotations

import argparse

from ._dispatch import exec_script, _rest_argv


_TARGETS = {
    "headings": "renumber_headings",
}


def _run(args) -> int:
    target = getattr(args, "renumber_target", None)
    script = _TARGETS.get(target)
    if script is None:
        print(f"[sub.renumber] unknown target: {target}; choices={list(_TARGETS)}")
        return 2
    return exec_script(script, _rest_argv(args))


def register(subparsers) -> None:
    p = subparsers.add_parser(
        "renumber",
        help="renumber headings + caption numbers by physical order",
    )
    sp = p.add_subparsers(dest="renumber_target", metavar="<target>", required=True)
    for t in _TARGETS:
        spp = sp.add_parser(t, help=f"renumber {t}", add_help=False)
        spp.add_argument("docx_path", nargs="?", help="target docx path")
        spp.add_argument("--dry-run", action="store_true")
        spp.add_argument("--no-backup", action="store_true")
        spp.add_argument("--report", help="write JSON report to this path")
        spp.add_argument("rest", nargs=argparse.REMAINDER, help="extra args forwarded to underlying script")
        spp.set_defaults(func=_run)
