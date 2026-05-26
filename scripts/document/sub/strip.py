"""strip.py — group module: strip stale/polluting docx elements (7 subcommands)

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
  strip orphan-media       ← strip_orphan_media.py  (2026-05-26 added)
                              remove word/media/* not referenced by any rId
  strip empty-captions     ← strip_empty_captions.py  (2026-05-26 added)
                              remove caption-style paragraphs that are strict-empty
                              (no text, no inline drawing, no field)
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
    "orphan-media":       "strip_orphan_media",
    "empty-captions":     "strip_empty_captions",
}

# targets that additionally support -o/--output (write to new path, leave original)
_TARGETS_WITH_OUTPUT = {"orphan-media", "empty-captions"}


def _rest_argv_with_output(args) -> list[str]:
    """Extend _rest_argv to also forward -o/--output if present."""
    argv = _rest_argv(args)
    out = getattr(args, "output", None)
    if out:
        argv.extend(["-o", str(out)])
    return argv


def _run(args) -> int:
    target = getattr(args, "strip_target", None)
    script = _TARGETS.get(target)
    if script is None:
        print(f"[sub.strip] unknown target: {target}; choices={list(_TARGETS)}")
        return 2
    argv_fn = _rest_argv_with_output if target in _TARGETS_WITH_OUTPUT else _rest_argv
    return exec_script(script, argv_fn(args))


def register(subparsers) -> None:
    p = subparsers.add_parser(
        "strip",
        help="strip stale/polluting docx elements (outlinelvl / bookmarks / revisions / doc-protection / orphan-media / empty-captions)",
    )
    sp = p.add_subparsers(dest="strip_target", metavar="<target>", required=True)
    for t in _TARGETS:
        spp = sp.add_parser(t, help=f"strip {t}", add_help=False)
        spp.add_argument("docx_path", nargs="?", help="target docx path")
        spp.add_argument("--dry-run", action="store_true")
        spp.add_argument("--no-backup", action="store_true")
        spp.add_argument("--report", help="write JSON report to this path")
        if t in _TARGETS_WITH_OUTPUT:
            spp.add_argument("-o", "--output", default=None,
                             help="write to new path (do not modify original, no bak)")
        spp.add_argument("rest", nargs=argparse.REMAINDER, help="extra args forwarded to underlying script")
        spp.set_defaults(func=_run)
