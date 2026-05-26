"""freeze.py — group module: freeze Word automation (2 subcommands)

Subcommands:
  freeze headings   ← freeze_heading_numbers.py
                       write H1-H4 auto-numbering as static text prefix +
                       remove style-level numPr (合稿断链)
  freeze fields     ← freeze_all_fields.py
                       freeze TOC / PAGEREF / SEQ / formula fields to static result text
                       (合稿防重算)
"""
from __future__ import annotations

import argparse

from ._dispatch import exec_script, _rest_argv


_TARGETS = {
    "headings": "freeze_heading_numbers",
    "fields":   "freeze_all_fields",
}


def _run(args) -> int:
    target = getattr(args, "freeze_target", None)
    script = _TARGETS.get(target)
    if script is None:
        print(f"[sub.freeze] unknown target: {target}; choices={list(_TARGETS)}")
        return 2
    return exec_script(script, _rest_argv(args))


def register(subparsers) -> None:
    p = subparsers.add_parser(
        "freeze",
        help="freeze auto-computed Word elements (headings numbering / fields)",
    )
    sp = p.add_subparsers(dest="freeze_target", metavar="<target>", required=True)
    for t in _TARGETS:
        spp = sp.add_parser(t, help=f"freeze {t}", add_help=False)
        spp.add_argument("docx_path", nargs="?", help="target docx path")
        spp.add_argument("--dry-run", action="store_true")
        spp.add_argument("--no-backup", action="store_true")
        spp.add_argument("--report", help="write JSON report to this path")
        spp.add_argument("rest", nargs=argparse.REMAINDER, help="extra args (e.g. --types TOC,PAGEREF)")
        spp.set_defaults(func=_run)
