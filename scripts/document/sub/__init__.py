"""sub — distilled docx-processing subcommand modules (W1+W2+W3 · 2026-05-25)

Distilled from 浙东引水 qual-supply 项目的 Word/docx 操作 SOP（原 zdwp monorepo,
2026-05-26 已拆为独立 repo `~/Work/projects/qual-supply/`）+ 一脚本一功能 ironclad rule.

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

from . import _groups
from . import (
    audit_styleset,
    blocks,
    captions,
    chapters_sync,
    chrome,
    combine,
    compare,
    diff,
    docx_para,
    fix_styleset,
    fonts,
    health,
    health_split,
    images,
    md_merge,
    md_merge_track,
    outline,
    pipeline,
    revise_rules,
    section,
    slim,
    split,
    styles,
    table,
)

__all__ = [
    "audit",
    "audit_styleset",
    "blocks",
    "caption",
    "captions",
    "chapters_sync",
    "chrome",
    "chapter",
    "combine",
    "compare",
    "diff",
    "docx_para",
    "fix_styleset",
    "fonts",
    "freeze",
    "header_footer",
    "health",
    "health_split",
    "images",
    "legacy",
    "md_merge",
    "md_merge_track",
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
    # 顺序 = 顶层 --help 里组的显示顺序，别随手调。
    # `_G("x")` = 由 _groups.py 的声明表注册的组（2026-07-30 起，取代 8 个只做转发的
    # group 模块）；`mod.register` = 还有自己逻辑、没被收编的模块。
    def _G(name):
        return lambda sp: _groups.register_group(sp, name)

    for reg in (
        _G("audit"), audit_styleset.register, _G("freeze"), _G("strip"),
        _G("header-footer"), outline.register, blocks.register, images.register,
        _G("legacy"),  # unique
        docx_para.register,                                                               # 段落级 查-改-验 工作台 (locate/inspect/edit/fix-ppr/scan-ppr/render, 2026-07-03)
        combine.register,                                                                 # combine N docx → 1 (docxcompose; inverse of split by-h1, 2026-06-07)
        chapters_sync.register,                                                           # 成品 docx 反向回写成品章节目录 (merge 的逆操作; govern 2026-06-08)
        chrome.register,                                                                  # 院报告版面装帧(逐章分节+横向节, distilled eco-flow 2026-06-04)
        diff.register, compare.register, revise_rules.register,                                             # distilled from bid-diff-and-revise
        health.register,                                                                  # health diagnose/fix/full
        pipeline.register,                                                                # pipeline driver
        section.register,                                                                 # section read/list (distilled from panan-rigid)
        md_merge.register,                                                                # md merge-into-docx (distilled from panan-rigid)
        md_merge_track.register,                                                          # md→track-changes 锚点前插 (上提 reclaim merge-tracked, 0-B 2026-05-29)
        table.register,                                                                   # table structural ops (delete-rows / borders / center, W4 2026-05-26)
        fonts.register,                                                                    # 字体/标题色规整 normalize (distilled reclaim 节水年会征文, 2026-06-21)
        split.register,                                                                   # split docx by-h1 (distilled from eco-flow/taizhou-天台, W1 2026-05-26)
        fix_styleset.register,                                                            # style-set fix family + shape_contract gate (W13 2026-05-26)
        health_split.register,                                                            # one-shot health + split thin wrapper (distilled from 业务模板 SOP, 2026-05-28)
        slim.register,                                                                    # docx-slim: safe ensemble + aggressive minimal skeleton (W docx-slim 2026-05-28)
        _G("chapter"), _G("renumber"), _G("caption"), captions.register, styles.register,   # shared
    ):
        reg(subparsers)
