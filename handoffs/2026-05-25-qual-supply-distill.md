# 2026-05-25 qual-supply docx 脚本族 distill 落地

> 4-worker fan-out (W1+W2+W3 distill + W4 dispatcher 整合) 把 qual-supply 项目 26+ 脚本 distill 上提到总部 `~/Dev/tools/doctools/`,让 eco-flow / shoreline / reclaim 等水利项目用 CLI 直接调用,不重复造轮子。

## 总览

- **落地模块**: 11 group module (sub/*.py) + 18 distilled standalone scripts(W1 cp)+ 2 merged monolith(W2 styles.py / outline.py)+ 4 group dispatcher(W3 blocks/captions/images/legacy)
- **总子命令**: 32 distilled + 16 legacy = **48 顶层入口** via `docx_cli.py`
- **样式 SSOT**: `config/styles_registry.yaml` 3 profile (zdwp / eco-flow / generic)
- **JSON schemas**: `schemas/{plan,decision,patch}.schema.json` (v1)
- **qual-supply 源**: 未动(只 cp 不 mv,验收前独立 CLI 仍可跑)

## 总入口

```bash
python3 ~/Dev/tools/doctools/scripts/document/docx_cli.py <subcommand> [--profile <name>] ...
```

## Distilled 11 子族 32 子命令清单

| 子族 | 子命令 | 落地 module | 源脚本 |
|---|---|---|---|
| `audit` (6 read-only) | headings / fields / captions / images / table-pairing / bookmarks | sub/audit.py | audit_heading_numbers / audit_word_fields / audit_caption_outline / audit_images / audit_table_pairing / audit_bookmarks |
| `freeze` (2) | headings / fields | sub/freeze.py | freeze_heading_numbers / freeze_all_fields |
| `strip` (5) | outlinelvl / style-outlinelvl / bookmarks / revisions / doc-protection | sub/strip.py | strip_outlinelvl_from_captions / strip_style_outlinelvl / strip_bookmarks / strip_revisions / strip_doc_protection |
| `renumber` (2 · 共享) | headings / h4-figures | sub/renumber.py + sub/styles.py | renumber_headings / renumber_h4_figures |
| `style` (3 · profile-driven) | body / table / caption | sub/styles.py | apply_body_styles / apply_table_styles / apply_caption_styles |
| `outline` (3) | promote-h1 / demote-h2 / normalize-arabic | sub/outline.py | promote_misclassified_h1 / demote_h2_with_h3_format / normalize_outline_to_arabic |
| `caption` (3 · 共享) | number / number-by-style / pair | sub/caption.py + sub/captions.py + sub/styles.py | number_captions / number_captions_by_style / pair_table_captions |
| `blocks` (2) | reorder / relocate | sub/blocks.py | reorder_heading_blocks / relocate_orphan_blocks |
| `chapter` (3 · 共享) | convert-arabic / delete-empty-h1 / delete | sub/chapter.py + sub/blocks.py | convert_chapter_format / delete_empty_h1 / delete_chapter |
| `header-footer` (1) | add | sub/header_footer.py | add_header_footer |
| `image` (1) | relink | sub/images.py | relink_images_from_source |
| `legacy` (1 · DEPRECATED) | fix-heading-disorder | sub/legacy.py | fix_heading_disorder |

## 共享 group 处理

W1+W2+W3 各自 distill 出独立 module,部分 group 名(`caption` / `chapter` / `renumber`)有多 module 贡献。`sub/_dispatch.py` 新增两个 helper 让 module 共享 group parent:

- `get_or_add_group(subparsers, name, help_text)` — 返回已有 group parser 或新建
- `get_or_add_subparsers(group_parser, dest)` — 复用已有 sub-subparsers action

子族 register() 函数检查 `existing = subparsers.choices` 防止重复添加同名 target。`sub/__init__.py register_all()` 按顺序调用 12 module 注册(unique-group 先注册,shared-group 贡献者后注册)。

## 样式 profile SSOT

`~/Dev/tools/doctools/config/styles_registry.yaml`:
- `zdwp` — ZDWP 集团数字 styleId (默认,兼容 qual-supply)
- `eco-flow` — 待填(模板已就位)
- `generic` — 兜底 (无 BODY_TARGET_STYLE_ID 等)

水利项目接入示例(eco-flow):

```bash
# 1. 扩 styles_registry.yaml 加 eco-flow profile (填 H1_STYLES / TARGET_STYLE_ID 等字段)
# 2. 调 docx_cli.py 子命令 + --profile eco-flow
python3 ~/Dev/tools/doctools/scripts/document/docx_cli.py style body X.docx --profile eco-flow
python3 ~/Dev/tools/doctools/scripts/document/docx_cli.py renumber h4-figures X.docx --profile eco-flow
```

## JSON schemas

3 schema 由 W3 立,位置 `~/Dev/tools/doctools/schemas/`:

| schema | version | consumer | 子命令 |
|---|---|---|---|
| `plan.schema.json` | 1 | relocate_orphan_blocks.py | `blocks relocate --plan plan.json` |
| `decision.schema.json` | 1 | pair_table_captions.py | `caption pair --decision decision.json` |
| `patch.schema.json` | 1 | relink_images_from_source.py | `image relink --apply-patch patch.json` |

Loader: `~/Dev/tools/doctools/lib/schemas.py`,exports `load_schema(name)` / `validate(data, name)` / `load_and_validate(json_path, name)`。

## SSOT 索引

- 子命令完整清单 + 描述: `~/Dev/tools/dev/lib/tools/report/hq_capabilities.yaml` doctools.sub_capabilities
- README: `~/Dev/tools/doctools/README.md` (头部 "2026-05-25 qual-supply distill" 节)
- qual-supply 项目 CLAUDE.md 引用: `~/Work/projects/zdwp/projects/qual-supply/CLAUDE.md` (引用总部节 + Word/docx 操作 SOP 节)

## 调用示例

```bash
# audit 套餐(read-only)
python3 ~/Dev/tools/doctools/scripts/document/docx_cli.py audit headings X.docx --report /tmp/h.json
python3 ~/Dev/tools/doctools/scripts/document/docx_cli.py audit captions X.docx
python3 ~/Dev/tools/doctools/scripts/document/docx_cli.py audit table-pairing X.docx

# freeze(合稿前)
python3 ~/Dev/tools/doctools/scripts/document/docx_cli.py freeze headings X.docx
python3 ~/Dev/tools/doctools/scripts/document/docx_cli.py freeze fields X.docx

# style(profile-driven)
python3 ~/Dev/tools/doctools/scripts/document/docx_cli.py style body X.docx --profile zdwp
python3 ~/Dev/tools/doctools/scripts/document/docx_cli.py style table X.docx --profile zdwp
python3 ~/Dev/tools/doctools/scripts/document/docx_cli.py style caption X.docx --profile zdwp

# outline 规范化
python3 ~/Dev/tools/doctools/scripts/document/docx_cli.py outline normalize-arabic X.docx

# renumber
python3 ~/Dev/tools/doctools/scripts/document/docx_cli.py renumber headings X.docx
python3 ~/Dev/tools/doctools/scripts/document/docx_cli.py renumber h4-figures X.docx

# blocks / chapter 结构
python3 ~/Dev/tools/doctools/scripts/document/docx_cli.py blocks reorder X.docx
python3 ~/Dev/tools/doctools/scripts/document/docx_cli.py chapter delete X.docx --h1 3
python3 ~/Dev/tools/doctools/scripts/document/docx_cli.py chapter delete-empty-h1 X.docx

# JSON-driven (decision/patch/plan)
python3 ~/Dev/tools/doctools/scripts/document/docx_cli.py caption pair X.docx --decision decision.json
python3 ~/Dev/tools/doctools/scripts/document/docx_cli.py blocks relocate X.docx --plan plan.json
python3 ~/Dev/tools/doctools/scripts/document/docx_cli.py image relink X.docx --apply-patch patch.json

# 旧族保留(REMAINDER 直转)
python3 ~/Dev/tools/doctools/scripts/document/docx_cli.py extract X.docx -o out.md
python3 ~/Dev/tools/doctools/scripts/document/docx_cli.py md format -i x.md
python3 ~/Dev/tools/doctools/scripts/document/docx_cli.py --batch tasks.jsonl --workers 8 extract

# 每脚本仍可独立 CLI 跑(standalone 兼容,验收前 qual-supply 源未动)
python3 ~/Dev/tools/doctools/scripts/document/sub/audit_heading_numbers.py X.docx
python3 ~/Work/projects/zdwp/projects/qual-supply/scripts/freeze_heading_numbers.py X.docx
```

## 状态

- [x] W1 distill — 18 脚本 + 7 group module + JSON `/tmp/distill-W1-subcommands.json`
- [x] W2 distill — 8 脚本 → 2 monolith (styles.py / outline.py) + styles_registry.yaml SSOT + lib/styles.py loader + JSON `/tmp/distill-W2-subcommands.json`
- [x] W3 distill — 6 脚本 + 4 group module (blocks/captions/images/legacy) + 3 JSON schemas + lib/schemas.py + JSON `/tmp/distill-W3-subcommands.json`
- [x] W4 整合 — sub/__init__.py register_all + docx_cli.py 路由 + hq_capabilities.yaml use_via + qual-supply CLAUDE.md + doctools README

## qual-supply 源状态

`~/Work/projects/zdwp/projects/qual-supply/scripts/` 26 脚本**未动**,仅被 cp 到 sub/。2026-06 验收期间仍可走原项目 CLI 跑;验收后用户决定是否移除 qual-supply 副本,改全走总部 CLI。
