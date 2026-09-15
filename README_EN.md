# doctools

[中文](README.md) | **English**

A document-processing and data-conversion toolkit that maintains both command-line engines and the [DocKit desktop client](mac/README.md). Both entry points reuse the same processing capabilities. Desktop builds and JSON contracts are documented in its directory; CLI entry points are listed below.

- Current inventory (script counts / how docx files are changed / gates): [`handoffs/_archive/2026-08-04-docx-scripts-inventory.md`](handoffs/_archive/2026-08-04-docx-scripts-inventory.md)
- Refactoring roadmap and closure ledger: [`handoffs/_archive/2026-08-04-docx-refactor-roadmap.md`](handoffs/_archive/2026-08-04-docx-refactor-roadmap.md)
- Development conventions (consolidation / element-level traversal / subcommand table): [`CLAUDE.md`](CLAUDE.md)

**103 scripts**: entry points 17 · `sub/` 44 · `data/` 5 · `lib/` 16 · `tools/` 8 · `src/doctools/` 2 · tests 11.

## Entry Layer (scripts/document/)

| Script | Function |
|------|------|
| `docx_cli.py` | **Main docx entry**: 46 unique families / 126 subcommands, dispatched to `sub/` |
| `word_table_format.command` | Removes shading, centers, and applies solid borders to Mac Word tables; applies the existing ZDWP图名 style to figure/table titles; [instructions](docs/word-table-format.md) |
| `docx_tools.py` | Combined extract / check / track entry (parallel batch + library re-export) |
| `typeset_apply.py` | spec(yaml)-driven typesetting engine (29 actions in a fixed order) |
| `typeset_pipeline.py` | End-to-end typesetting driver (each step: snapshot → self-check → retain/rollback) |
| `docx_revise.py` | Revision injection: comments represented as ops.yaml data → `w:ins/w:del` + comments |
| `bid_gate.py` | Final tender-document gate family (run / scan / sweep / identity / print / deref) |
| `docx_fmt.py` | docx layout/font/text normalization family (template / clone / fonts / text) |
| `renum.py` | Numbering/caption shifting and reordering family (chapter / tabfig / figures) |
| `md_tools.py` | Markdown toolkit with 8 subcommands (including style-replicating `md2docx` conversion) |
| `pptx_cli.py` | PPTX family with 13 subcommands (dual-interpreter contract: system python3 / venv each handle half) |
| `pdf_cli.py` | PDF family with 14 subcommands (read / convert / pipeline / chart extraction…) |
| `md_to_audiobook.py` | md → audiobook (concurrent edge-tts chapters; dependencies declared with PEP-723) |
| `chart.py` | Data-driven chart generation (bar / gantt / flow / insert, JSON → PNG) |
| `doc_dispatch.py` | Unified dispatcher routed by file extension (commands express only verbs; formats are identified at runtime) |
| `doc_gui_backend.py` | Wraps `doc_dispatch` in a JSON envelope for the SwiftUI app |
| `docx_write_gate.py` | In-place writeback concurrency gate (compares md5/mtime baselines before writing) |
| `bid_residue_lib.py` | SSOT for tender-residue detection (imported by `bid_gate`, not invoked directly) |

Supporting files: `config/` (styles_registry / spec-examples / schema), `schemas/`.

## Data-Conversion Scripts (data/)

| Script | Function |
|------|------|
| `data.py` | Unified data-processing CLI |
| `convert.py` | Unified format-conversion tool (consolidating conversions among 8 formats) |
| `xlsx_lowercase.py` | Lowercases text in Office documents |
| `xlsx_merge_tables.py` | Merges Excel tables (AI-assisted matching) |
| `xlsx_splitsheets.py` | Splits Excel worksheets into separate files |

## Shared Libraries (lib/)

`docx_surgical` (zipfile+lxml surgical engine) · `docx_safe_save` (centralized python-docx saving) ·
`docx_parts` (part-integrity assertions) · `docx_xml` (element-level traversal covering comments/footnotes/endnotes/headers/footers) ·
`docx_revise` (revision-injection engine) · `styles` · `chapter_numbering` · `text_fixes` · `schemas` ·
`soffice` · `llm_client` · `clipboard` · `progress` · `common.sh`

## Gates (Run All After Changes)

```bash
python3 tools/check_docx_collar.py    # 收口 23/23 · 部件断言 16/16，缺一判红（判据走 ast，只认真调用）
python3 tools/cli_surface.py          # CLI 接口指纹（126 子命令）
python3 tools/cli_forward_probe.py    # 67 条内嵌预期 argv 比对（真正转发出去的是什么）
python3 -m pytest scripts/document/tests scripts/document/sub/tests -q
python3 -m pytest tests/smoke -q      # 逐条真敲 91 条动词（rc + 源件 md5 双向断言）
python3 tools/check_smoke_coverage.py # 冒烟表 ↔ CLI 对账：真跑 91 / skip 2 / 共 93
python3 tools/check_function_axis.py  # 职能轴表 ↔ CLI 对账，缺一条或多一条都判红
python3 tools/check_external_refs.py   # 全生态引用存在性（已挂 pre-commit --changed-only）
python3 tools/check_verbs_reachable.py # 每个注册过的顶层子命令敲下去必须真进得去
python3 tools/script_graph.py --open  # 103 脚本 · 370 引用 · 93 动词 · 0 孤儿（三视图：图谱/清单/动词）
```

None of the first four gates **actually executes** any subcommand. Correct interface fingerprints, forwarded argv,
and function labels can still coexist with every real invocation failing. `tests/smoke` fills that gap:
it is the only layer that actually runs commands and checks whether the source file should change afterward.

## Installation

Python dependencies use the `~/Dev` uv workspace (this repository is a member; do not create an independent `.venv`):

```bash
cd ~/Dev && uv sync --all-packages
```

To change dependencies, edit this repository's `pyproject.toml`, then rerun the command above.

The same sync installs this repository as an editable package and creates the `doctools` command:

```bash
~/Dev/.venv/bin/doctools --version        # doctools 0.1.0
~/Dev/.venv/bin/doctools verbs --fn format
```

This **coexists with absolute-path invocation rather than replacing it**: `doctools <sub> …` and
`python3 ~/Dev/tools/doctools/scripts/document/docx_cli.py <sub> …` enter the same
`main()` (measured stdout for the same command is byte-for-byte identical), so there is no divergence
between the two entry points. The 130 absolute-path references under ~/Work need no changes.

**The only version SSOT is `__version__` in `src/doctools/__init__.py`**:
`pyproject.toml` declares `dynamic = ["version"]`, letting hatchling read that file;
`--version` reads it too. **Do not add a `version =` line to `pyproject.toml`**—that would create
two separate copies again. After changing the version, update installed metadata with
`uv pip install -e tools/doctools --no-deps --reinstall` (`uv sync` does not rebuild dist-info
merely because the version number changed).

## docx_cli Subcommand Families

**Main entry**: `python3 ~/Dev/tools/doctools/scripts/document/docx_cli.py <subcommand>`
(equivalent after package installation: `doctools <subcommand>`)

All declarations live in the `GROUPS` table in `scripts/document/sub/_groups.py`: **adding a subcommand means adding a data row, not a file**.

**126 commands = 49 top-level names (including 3 aliases: `read`=extract · `diff`=compare · `styleset`=audit-styleset,
leaving 46 unique families) + 77 actions within 25 families.**

| Family | Actions |
|---|---|
| `audit` (6) | bookmarks / captions / fields / headings / images / table-pairing — read-only checks |
| `audit-styleset` (6) | style-coherence / role-coverage / body-style-concentration / style-pool-cleanliness / style-pane-filter / restore |
| `fix` (7) | clear-direct-format / role-fill / style-create / style-pane-filter / style-pool-cleanup / style-rebrand / style-rename |
| `strip` (7) | bookmarks / doc-protection / empty-captions / orphan-media / outlinelvl / revisions / style-outlinelvl |
| `para` (6) | locate / inspect / edit / fix-ppr / scan-ppr / render — paragraph-level inspection/editing/validation workbench |
| `health` (4) | diagnose / fix / full / gate |
| `table` (4) | borders / center / delete-rows / extract |
| `caption` (3) | number / number-by-style / pair |
| `chapter` (3) | convert-arabic / delete / delete-empty-h1 |
| `outline` (3) | promote-h1 / demote-h2 / normalize-arabic |
| `style` (3) | body / table / caption — applies the correct group-named style family |
| `blocks` (2) | reorder / relocate |
| `freeze` (2) | headings / fields — freezes automatic numbering/fields before merging drafts |
| `image` (2) | relink / extract |
| `renumber` (2) | headings / h4-figures |
| `seqdiff` (2) | seq / image |
| `split` (2) | by-h1 / body-replace |
| `compare-ref` · `fonts` · `header-footer` · `legacy` · `pipeline` · `revise-rules` · `section` (1 each) | ref / normalize / add / fix-heading-disorder(DEPRECATED) / run / gen / read |

24 leaf commands (without sub-actions): `chapters-sync` `check` `chrome` `combine` `compare(diff)` `extract(read)`
`fix-ref` `format` `health-split` `image-caption` `md` `md-merge` `md-merge-track` `md-to-docx`
`renumber-fig` `scan-sensitive` `slim` `snapshot` `template` `text-fmt` `track` `verbs`

This table is derived from `python3 tools/cli_surface.py` output (checked on 2026-08-01); **that command is the SSOT, not this table**.

**SSOT index**:

- Subcommand capability list: `~/Dev/tools/dev/lib/tools/report/hq_capabilities.yaml` → `doctools.sub_capabilities`
- Style-family profiles: `config/styles_registry.yaml` (zdwp / eco-flow / generic)
- JSON schemas: `schemas/{plan,decision,patch}.schema.json`
- spec examples: `config/spec-examples/` (bid.yaml / report-generic.yaml)

**Invocation examples**:

```bash
D=~/Dev/tools/doctools/scripts/document

# audit 类（read-only）
python3 $D/docx_cli.py audit headings X.docx --report /tmp/h.json

# freeze（合稿前）+ renumber + style（--profile 选样式族）
python3 $D/docx_cli.py freeze headings X.docx
python3 $D/docx_cli.py renumber headings X.docx
python3 $D/docx_cli.py style body X.docx --profile zdwp

# spec 驱动整篇排版
python3 $D/typeset_apply.py X.docx --spec config/spec-examples/report-generic.yaml

# 各 sub/ 实现仍可独立敲
python3 $D/sub/audit.py headings X.docx --report /tmp/h.json
```

## 2026-05-25 Consolidation of the qual-supply docx Script Family

26+ qual-supply project scripts were distilled into this package so that water-sector projects such as eco-flow / shoreline / reclaim can call them through the CLI without rebuilding the same capabilities.
Details: `handoffs/_archive/2026-07-02-2026-05-25-qual-supply-distill.md`.

Water-sector project integration (eco-flow example): write project styles yaml or add a profile to `config/styles_registry.yaml` →
`docx_cli.py style body <docx> --profile eco-flow` applies the correct style family.
