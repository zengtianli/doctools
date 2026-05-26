"""caption.py — group module: caption (table/figure) operations (1 subcommand)

Subcommands:
  caption number   ← number_captions.py
                       heuristic prepend "表 X-Y" / "图 X-Y" caption numbers to caption
                       paragraphs based on H1 chapter context. Text-based heuristic
                       (CJK numerals / arabic short paragraphs as H1 anchors).

Sister script (pStyle-aware variant) `number_captions_by_style.py` lives in qual-supply/
and is not part of this distill batch (W2/W3 candidates).
"""
from __future__ import annotations

import argparse

from ._dispatch import exec_script, _rest_argv, get_or_add_group, get_or_add_subparsers


_TARGETS = {
    "number": "number_captions",
}


def _run(args) -> int:
    target = getattr(args, "caption_target", None)
    script = _TARGETS.get(target)
    if script is None:
        print(f"[sub.caption] unknown target: {target}; choices={list(_TARGETS)}")
        return 2
    return exec_script(script, _rest_argv(args))


def register(subparsers) -> None:
    # Conflict-tolerant: shared `caption` group with captions.py (W3, pair) +
    # styles.py (W2, number-by-style). Use shared `caption_target` dest.
    p = get_or_add_group(subparsers, "caption", "caption ops (number / pair / number-by-style)")
    sp = get_or_add_subparsers(p, dest="caption_target")
    for t in _TARGETS:
        if t in (getattr(sp, "choices", {}) or {}):
            continue  # already registered by another module
        spp = sp.add_parser(t, help=f"caption {t}", add_help=False)
        spp.add_argument("docx_path", nargs="?", help="target docx path")
        spp.add_argument("--dry-run", action="store_true")
        spp.add_argument("--no-backup", action="store_true")
        spp.add_argument("--report", help="write JSON report to this path")
        spp.add_argument("rest", nargs=argparse.REMAINDER, help="extra args forwarded to underlying script")
        spp.set_defaults(func=_run)
