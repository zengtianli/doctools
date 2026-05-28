"""sub — distilled docx-processing subcommand modules (W1+W2+W3 · 2026-05-25)

Distilled from `~/Work/projects/zdwp/projects/qual-supply/scripts/` per
qual-supply CLAUDE.md « Word/docx 操作 SOP » + 一脚本一功能 ironclad rule.

Group modules (each exposes `register(subparsers)` for docx_cli.py dispatcher):

  audit         — audit-only docx checks (6 targets):
                    headings / fields / captions / images / table-pairing / bookmarks
  freeze        — freeze auto-numbering & fields (2 targets):
                    headings / fields
  strip         — strip stale/polluting elements (5 targets):
                    outlinelvl / style-outlinelvl / bookmarks / revisions / doc-protection
  header_footer — header/footer ops (1 target): add
  chapter       — chapter / H1 text ops (shared group; 3 targets):
                    convert-arabic / delete-empty-h1 (chapter.py) + delete (blocks.py)
  renumber      — renumber headings + h4-figures (shared group; 2 targets):
                    headings (renumber.py) + h4-figures (styles.py)
  caption       — caption ops (shared group; 3 targets):
                    number (caption.py) + pair (captions.py) + number-by-style (styles.py)
  blocks        — paragraph-block structural ops (2 targets): reorder / relocate
  outline       — outline level normalization (3 targets):
                    promote-h1 / demote-h2 / normalize-arabic
  style         — style application (profile-driven, 3 targets):
                    body / table / caption
  image         — image ops (1 target): relink
  section       — section read/list ops (1 target): read (distilled from panan-rigid, 2026-05-26)
  md_merge      — merge MD into DOCX section (1 target): md merge-into-docx (distilled from panan-rigid, 2026-05-26)
  table         — table structural ops (1 target): delete-rows (distilled from bid-diff-and-revise, 2026-05-26)
  legacy        — deprecated/spike (1 target): fix-heading-disorder

Total: 14 group modules -> 33+ subcommands (group dedupe handled by
_dispatch.get_or_add_group / get_or_add_subparsers).

Each underlying script also remains independently runnable:
    python3 sub/<script>.py <docx> [--dry-run] [--no-backup] [--report x.json]
"""

from . import (
    audit,
    audit_styleset,
    blocks,
    caption,
    captions,
    chapter,
    compare,
    diff,
    fix_styleset,
    freeze,
    header_footer,
    health,
    health_split,
    images,
    legacy,
    md_merge,
    outline,
    pipeline,
    renumber,
    revise_rules,
    section,
    slim,
    split,
    strip,
    styles,
    table,
)

__all__ = [
    "audit",
    "audit_styleset",
    "blocks",
    "caption",
    "captions",
    "chapter",
    "compare",
    "diff",
    "fix_styleset",
    "freeze",
    "header_footer",
    "health",
    "health_split",
    "images",
    "legacy",
    "md_merge",
    "outline",
    "pipeline",
    "renumber",
    "revise_rules",
    "section",
    "slim",
    "split",
    "strip",
    "styles",
    "table",
]


def register_all(subparsers) -> None:
    """Convenience: register every group's subcommands onto a parent subparsers.

    Registration order matters for shared groups (caption / chapter / renumber):
    first-registrant defines the group parser; later modules add targets via
    `get_or_add_group` / `get_or_add_subparsers` helpers in _dispatch.py.

    Usage in docx_cli.py:
        from sub import register_all
        register_all(top_subparsers)
    """
    # Order: unique-group first, then shared-group contributors
    for mod in (
        audit, audit_styleset, freeze, strip, header_footer, outline, blocks, images, legacy,  # unique
        diff, compare, revise_rules,                                             # distilled from bid-diff-and-revise
        health,                                                                  # health diagnose/fix/full
        pipeline,                                                                # pipeline driver
        section,                                                                 # section read/list (distilled from panan-rigid)
        md_merge,                                                                # md merge-into-docx (distilled from panan-rigid)
        table,                                                                   # table structural ops (delete-rows, W4 2026-05-26)
        split,                                                                   # split docx by-h1 (distilled from eco-flow/taizhou-天台, W1 2026-05-26)
        fix_styleset,                                                            # style-set fix family + shape_contract gate (W13 2026-05-26)
        health_split,                                                            # one-shot health + split thin wrapper (distilled from 业务模板 SOP, 2026-05-28)
        slim,                                                                    # docx-slim: safe ensemble + aggressive minimal skeleton (W docx-slim 2026-05-28)
        chapter, renumber, caption, captions, styles,                           # shared
    ):
        mod.register(subparsers)
