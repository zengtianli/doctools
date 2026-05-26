# doctools

文档处理与数据转换工具集，从 [scripts](https://github.com/tianlizeng/scripts) 仓库拆分。macOS 环境，主要通过 Raycast 调用。

## 文档处理脚本 (document/)

| 脚本 | 功能 |
|------|------|
| `bullet_to_paragraph.py` | 要点转公文段落/表格（AI） |
| `chart.py` | 数据驱动图表生成（JSON -> PNG） |
| `docx_apply_image_caption.py` | 应用图片和图名样式 |
| `docx_apply_template.py` | Word 文档样式套模板 + 清理 |
| `docx_text_formatter.py` | 文本格式自动修复（DOCX） |
| `docx_tools.py` | Word 文档工具集 |
| `md_docx_template.py` | Markdown 转 Docx（样式复刻） |
| `md_tools.py` | Markdown 工具集 |
| `pptx_to_md.py` | PPTX 转 Markdown |
| `pptx_tools.py` | PPTX 文档标准化工具集 |
| `report_quality_check.py` | 报告/标书质量检查 + 自动修复 |
| `scan_sensitive_words.py` | AI 敏感词扫描器 |

辅助文件：`heading_styles.xml`（标题样式定义）、`styles_config.json`（样式配置）、`docx_to_md.sh`（转换脚本）

## 数据转换脚本 (data/)

| 脚本 | 功能 |
|------|------|
| `convert.py` | 数据格式转换统一工具（合并 8 种格式互转） |
| `xlsx_lowercase.py` | Office 文档文本小写化 |
| `xlsx_merge_tables.py` | Excel 多表合并（AI 智能匹配） |
| `xlsx_splitsheets.py` | Excel 工作表拆分为多个文件 |

## 公共库 (lib/)

`display` / `file_ops` / `finder` / `progress` / `docx_xml` / `clipboard` / `env` / `usage_log` / `llm_client`

## 安装

```bash
pip3 install -r requirements.txt
```

## 2026-05-25 qual-supply docx 脚本族 distill 落地

26+ qual-supply 项目脚本 distill 上提到本 package, 让 eco-flow / shoreline / reclaim 等水利项目用 CLI 直接调用,不重复造轮子。

**总入口**: `python3 ~/Dev/tools/doctools/scripts/document/docx_cli.py <subcommand>`

子命令 11 子族 30+ 命令:

- `audit` (6): headings / fields / captions / images / table-pairing / bookmarks  —— read-only 检查
- `freeze` (2): headings / fields  —— 合稿前冻结自动编号/字段域
- `strip` (5): outlinelvl / style-outlinelvl / bookmarks / revisions / doc-protection
- `renumber` (2): headings / h4-figures
- `style` (3): body / table / caption  —— 套对集团命名样式族
- `outline` (3): promote-h1 / demote-h2 / normalize-arabic
- `caption` (3): number / number-by-style / pair
- `blocks` (2): reorder / relocate
- `chapter` (3): delete / delete-empty-h1 / convert-arabic
- `header-footer` (1): add  —— 水利院标准页眉页脚
- `image` (1): relink  —— 从源 docx 提媒体重嵌
- 旧族保留: extract / check / snapshot / compare / track / bullet / image-caption / template / renumber-fig / text-fmt / fix-ref / md-to-docx / quality-check / review / scan-sensitive / md (子组)

**SSOT 索引**:

- 子命令完整能力清单: `~/Dev/tools/dev/lib/tools/report/hq_capabilities.yaml` doctools.sub_capabilities
- 样式族 profile SSOT: `~/Dev/tools/doctools/config/styles_registry.yaml` (zdwp / eco-flow / generic)
- JSON schemas: `~/Dev/tools/doctools/schemas/{plan,decision,patch}.schema.json`
- distill 详情: `~/Dev/tools/doctools/handoffs/2026-05-25-qual-supply-distill.md`

**调用示例**:

```bash
# audit 类(read-only, 默认 dry-run)
python3 ~/Dev/tools/doctools/scripts/document/docx_cli.py audit headings X.docx --report /tmp/h.json
python3 ~/Dev/tools/doctools/scripts/document/docx_cli.py audit captions X.docx

# freeze 类(合稿前)
python3 ~/Dev/tools/doctools/scripts/document/docx_cli.py freeze headings X.docx
python3 ~/Dev/tools/doctools/scripts/document/docx_cli.py freeze fields X.docx

# renumber + style(--profile 选样式族)
python3 ~/Dev/tools/doctools/scripts/document/docx_cli.py renumber headings X.docx --profile eco-flow
python3 ~/Dev/tools/doctools/scripts/document/docx_cli.py style body X.docx --profile zdwp

# 各 distilled 脚本仍可独立 CLI 跑
python3 ~/Dev/tools/doctools/scripts/document/sub/audit_heading_numbers.py X.docx --report /tmp/h.json
```

**水利项目接入**(eco-flow 范例):

1. 写 `~/Work/projects/zdwp/projects/eco-flow/scripts/eco-flow-styles.yaml` 或扩 `config/styles_registry.yaml` 加 `eco-flow` profile
2. 调用 `docx_cli.py style body <docx> --profile eco-flow` 即套对样式族
3. qual-supply 源脚本(`~/Work/projects/zdwp/projects/qual-supply/scripts/`)保留为 spike 备份(2026-06 验收用),已 cp 到总部不动源

