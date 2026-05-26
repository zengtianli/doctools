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
  legacy        — deprecated/spike (1 target): fix-heading-disorder

Total: 11 group modules -> 30+ subcommands (group dedupe handled by
_dispatch.get_or_add_group / get_or_add_subparsers).

Each underlying script also remains independently runnable:
    python3 sub/<script>.py <docx> [--dry-run] [--no-backup] [--report x.json]
"""

from . import (
    audit,
    blocks,
    caption,
    captions,
    chapter,
    freeze,
    header_footer,
    images,
    legacy,
    outline,
    renumber,
    strip,
    styles,
)

__all__ = [
    "audit",
    "blocks",
    "caption",
    "captions",
    "chapter",
    "freeze",
    "header_footer",
    "images",
    "legacy",
    "outline",
    "renumber",
    "strip",
    "styles",
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
        audit, freeze, strip, header_footer, outline, blocks, images, legacy,  # unique
        chapter, renumber, caption, captions, styles,                           # shared
    ):
        mod.register(subparsers)
