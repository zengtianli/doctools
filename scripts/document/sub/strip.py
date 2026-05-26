"""strip.py — group module: strip stale/polluting docx elements (5 subcommands)

Subcommands:
  strip outlinelvl         ← strip_outlinelvl_from_captions.py
                              remove paragraph-level <w:outlineLvl> from caption paragraphs
                              (caption 段从 Word 导航大纲消失)
  strip style-outlinelvl   ← strip_style_outlinelvl.py
                              remove style-level <w:outlineLvl> from caption styles in
                              word/styles.xml (heading 真标题保护)
  strip bookmarks          ← strip_bookmarks.py
  strip revisions          ← strip_revisions.py
                              clean tracked changes / comments
  strip doc-protection     ← strip_doc_protection.py
                              remove <w:documentProtection> from settings.xml
"""
from __future__ import annotations

import argparse

from ._dispatch import exec_script, _rest_argv


_TARGETS = {
    "outlinelvl":         "strip_outlinelvl_from_captions",
    "style-outlinelvl":   "strip_style_outlinelvl",
    "bookmarks":          "strip_bookmarks",
    "revisions":          "strip_revisions",
    "doc-protection":     "strip_doc_protection",
}


def _run(args) -> int:
    target = getattr(args, "strip_target", None)
    script = _TARGETS.get(target)
    if script is None:
        print(f"[sub.strip] unknown target: {target}; choices={list(_TARGETS)}")
        return 2
    return exec_script(script, _rest_argv(args))


def register(subparsers) -> None:
    p = subparsers.add_parser(
        "strip",
        help="strip stale/polluting docx elements (outlinelvl / bookmarks / revisions / doc-protection)",
    )
    sp = p.add_subparsers(dest="strip_target", metavar="<target>", required=True)
    for t in _TARGETS:
        spp = sp.add_parser(t, help=f"strip {t}", add_help=False)
        spp.add_argument("docx_path", nargs="?", help="target docx path")
        spp.add_argument("--dry-run", action="store_true")
        spp.add_argument("--no-backup", action="store_true")
        spp.add_argument("--report", help="write JSON report to this path")
        spp.add_argument("rest", nargs=argparse.REMAINDER, help="extra args forwarded to underlying script")
        spp.set_defaults(func=_run)
