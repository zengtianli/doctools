"""sub — distilled docx-processing subcommand modules (W1 · 2026-05-25)

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
  chapter       — chapter / H1 text ops (2 targets):
                    convert-arabic / delete-empty-h1
  renumber      — renumber headings (1 target): headings
  caption       — caption ops (1 target): number

Total: 7 group modules → 18 subcommands → 18 distilled standalone scripts.

Each underlying script also remains independently runnable:
    python3 sub/<script>.py <docx> [--dry-run] [--no-backup] [--report x.json]
"""

from . import (
    audit,
    caption,
    chapter,
    freeze,
    header_footer,
    renumber,
    strip,
)

__all__ = [
    "audit",
    "caption",
    "chapter",
    "freeze",
    "header_footer",
    "renumber",
    "strip",
]


def register_all(subparsers) -> None:
    """Convenience: register every group's subcommands onto a parent subparsers.

    Usage in docx_cli.py:
        from sub import register_all
        register_all(top_subparsers)
    """
    for mod in (audit, freeze, strip, header_footer, chapter, renumber, caption):
        mod.register(subparsers)
