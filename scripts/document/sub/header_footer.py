"""header_footer.py — group module: docx header/footer ops (1 subcommand)

Subcommands:
  header-footer add   ← add_header_footer.py
                          add standard hydrology-院 header (report title, right-aligned)
                          + footer (院 name + PAGE field, centered) to all sections.

Required extras: --header <text> --footer-prefix <院名> [--page-number]
"""
from __future__ import annotations

import argparse

from ._dispatch import exec_script, _rest_argv


_TARGETS = {
    "add": "add_header_footer",
}


def _run(args) -> int:
    target = getattr(args, "hf_target", None)
    script = _TARGETS.get(target)
    if script is None:
        print(f"[sub.header_footer] unknown target: {target}; choices={list(_TARGETS)}")
        return 2
    return exec_script(script, _rest_argv(args))


def register(subparsers) -> None:
    p = subparsers.add_parser(
        "header-footer",
        help="add/manage docx header & footer (院标准格式)",
    )
    sp = p.add_subparsers(dest="hf_target", metavar="<action>", required=True)
    for t in _TARGETS:
        spp = sp.add_parser(t, help=f"header-footer {t}", add_help=False)
        spp.add_argument("docx_path", nargs="?", help="target docx path")
        spp.add_argument("--dry-run", action="store_true")
        spp.add_argument("--no-backup", action="store_true")
        spp.add_argument("--report", help="write JSON report to this path")
        spp.add_argument("rest", nargs=argparse.REMAINDER,
                         help="forwarded args (--header / --footer-prefix / --page-number / --font-size / --gap-spaces / --header-align / --footer-align)")
        spp.set_defaults(func=_run)
