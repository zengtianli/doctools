# doctools

文档处理与数据转换工具集，从 [scripts](https://github.com/tianlizeng/scripts) 仓库拆分。macOS 环境，**入口是命令行**
（`raycast/` 整子树 2026-07-27 已归档，不再是调用方式）。

- 现状盘点（有几个脚本 / docx 怎么被改 / 闸门）：[`handoffs/docx-scripts-inventory.md`](handoffs/docx-scripts-inventory.md)
- 改造路线图与销项账：[`handoffs/docx-refactor-roadmap.md`](handoffs/docx-refactor-roadmap.md)
- 开发约定（收口 / 元素级遍历 / 子命令表）：[`CLAUDE.md`](CLAUDE.md)

**99 个脚本**：入口 17 · `sub/` 44 · `data/` 5 · `lib/` 16 · `tools/` 7 · 测试 10。

## 入口层 (scripts/document/)

| 脚本 | 功能 |
|------|------|
| `docx_cli.py` | **docx 总入口**：46 个唯一族 / 126 条子命令，dispatch 到 `sub/` |
| `docx_tools.py` | extract / check / track 组合入口（batch 并行 + library re-export） |
| `typeset_apply.py` | spec(yaml) 驱动的排版引擎（29 actions，固定顺序） |
| `typeset_pipeline.py` | 排版一条龙 driver（每步 snapshot → 自检 → 保留/回滚） |
| `docx_revise.py` | 修订注入：意见 = ops.yaml 数据 → `w:ins/w:del` + 批注 |
| `bid_gate.py` | 标书终稿门检族（run / scan / sweep / identity / print / deref） |
| `docx_fmt.py` | docx 版式/字体/文本规范化族（template / clone / fonts / text） |
| `renum.py` | 编号/题注位移与重排族（chapter / tabfig / figures） |
| `md_tools.py` | Markdown 工具集 8 子命令（含 `md2docx` 样式复刻转换） |
| `pptx_cli.py` | PPTX 族 13 子命令（双解释器契约：系统 python3 / venv 各管一半） |
| `pdf_cli.py` | PDF 族 14 子命令（read / convert / pipeline / 图表抽取…） |
| `md_to_audiobook.py` | md → 有声书（edge-tts 章节并发；PEP-723 自带依赖） |
| `chart.py` | 数据驱动图表生成（bar / gantt / flow / insert，JSON → PNG） |
| `doc_dispatch.py` | 按文件后缀路由的统一调度器（命令只表达动词，格式运行时认） |
| `doc_gui_backend.py` | 给 `doc_dispatch` 套 JSON 信封，供 SwiftUI app 调用 |
| `docx_write_gate.py` | 原地写回并发门（写回前 md5/mtime 基线比对） |
| `bid_residue_lib.py` | 标书残留检测逻辑 SSOT（被 `bid_gate` import，不单敲） |

辅助文件：`config/`（styles_registry / spec-examples / schema）、`schemas/`。

## 数据转换脚本 (data/)

| 脚本 | 功能 |
|------|------|
| `data.py` | 数据处理统一 CLI |
| `convert.py` | 数据格式转换统一工具（合并 8 种格式互转） |
| `xlsx_lowercase.py` | Office 文档文本小写化 |
| `xlsx_merge_tables.py` | Excel 多表合并（AI 智能匹配） |
| `xlsx_splitsheets.py` | Excel 工作表拆分为多个文件 |

## 公共库 (lib/)

`docx_surgical`（zipfile+lxml 手术引擎）· `docx_safe_save`（python-docx 存盘收口）·
`docx_parts`（部件完整性断言）· `docx_xml`（元素级遍历，覆盖批注/脚注/尾注/页眉页脚）·
`docx_revise`（修订注入引擎）· `styles` · `chapter_numbering` · `text_fixes` · `schemas` ·
`soffice` · `llm_client` · `clipboard` · `progress` · `common.sh`

## 闸门（改完全跑）

```bash
python3 tools/check_docx_collar.py    # 收口 23/23 · 部件断言 16/16，缺一判红（判据走 ast，只认真调用）
python3 tools/cli_surface.py          # CLI 接口指纹（126 子命令）
python3 tools/cli_forward_probe.py    # 67 条内嵌预期 argv 比对（真正转发出去的是什么）
python3 -m pytest scripts/document/tests scripts/document/sub/tests -q
python3 tools/check_function_axis.py  # 职能轴表 ↔ CLI 对账，缺一条或多一条都判红
python3 tools/check_external_refs.py   # 全生态引用存在性（已挂 pre-commit --changed-only）
python3 tools/script_graph.py --open  # 99 脚本 · 365 引用 · 93 动词 · 0 孤儿（三视图：图谱/清单/动词）
```

## 安装

Python 依赖走 `~/Dev` uv workspace（本 repo 是 member，不建独立 `.venv`）：

```bash
cd ~/Dev && uv sync --all-packages
```

改 deps → 改本 repo `pyproject.toml` 后重跑上面那行。

同一行 sync 会把本 repo 按 editable 装成包，落一个 `doctools` 命令：

```bash
~/Dev/.venv/bin/doctools --version        # doctools 0.1.0
~/Dev/.venv/bin/doctools verbs --fn format
```

它和绝对路径调用**是并存关系，不是替代** —— `doctools <sub> …` 与
`python3 ~/Dev/tools/doctools/scripts/document/docx_cli.py <sub> …` 进的是同一个
`main()`（实测同一条命令两边 stdout 逐字节相同），所以不会有「两个入口行为不一样」
这种事。~/Work 那 130 处绝对路径不用动。

**版本号 SSOT = `doctools/__init__.py` 的 `__version__`，只此一份**：
`pyproject.toml` 声明 `dynamic = ["version"]` 由 hatchling 从该文件读走，
`--version` 也读它。**别在 `pyproject.toml` 里补一行 `version =`** —— 那就又变回
两处各写一份了。改完版本号要让已装的元数据跟上，得
`uv pip install -e tools/doctools --no-deps --reinstall`（`uv sync` 不会因为版本号
变了就重建 dist-info）。

## docx_cli 子命令族

**总入口**：`python3 ~/Dev/tools/doctools/scripts/document/docx_cli.py <subcommand>`
（装包后等价：`doctools <subcommand>`）

声明全在 `scripts/document/sub/_groups.py` 的 `GROUPS` 表 —— **加子命令 = 加一行数据，不是加一个文件**。

**126 条 = 49 个顶层名（含 3 个 alias：`read`=extract · `diff`=compare · `styleset`=audit-styleset，
故唯一族 46）+ 25 个族里的 77 个动作。**

| 族 | 动作 |
|---|---|
| `audit` (6) | bookmarks / captions / fields / headings / images / table-pairing —— read-only 检查 |
| `audit-styleset` (6) | style-coherence / role-coverage / body-style-concentration / style-pool-cleanliness / style-pane-filter / restore |
| `fix` (7) | clear-direct-format / role-fill / style-create / style-pane-filter / style-pool-cleanup / style-rebrand / style-rename |
| `strip` (7) | bookmarks / doc-protection / empty-captions / orphan-media / outlinelvl / revisions / style-outlinelvl |
| `para` (6) | locate / inspect / edit / fix-ppr / scan-ppr / render —— 段落级查-改-验工作台 |
| `health` (4) | diagnose / fix / full / gate |
| `table` (4) | borders / center / delete-rows / extract |
| `caption` (3) | number / number-by-style / pair |
| `chapter` (3) | convert-arabic / delete / delete-empty-h1 |
| `outline` (3) | promote-h1 / demote-h2 / normalize-arabic |
| `style` (3) | body / table / caption —— 套对集团命名样式族 |
| `blocks` (2) | reorder / relocate |
| `freeze` (2) | headings / fields —— 合稿前冻结自动编号/字段域 |
| `image` (2) | relink / extract |
| `renumber` (2) | headings / h4-figures |
| `seqdiff` (2) | seq / image |
| `split` (2) | by-h1 / body-replace |
| `compare-ref` · `fonts` · `header-footer` · `legacy` · `pipeline` · `revise-rules` · `section` (各 1) | ref / normalize / add / fix-heading-disorder(DEPRECATED) / run / gen / read |

24 个叶命令（无子动作）：`chapters-sync` `check` `chrome` `combine` `compare(diff)` `extract(read)`
`fix-ref` `format` `health-split` `image-caption` `md` `md-merge` `md-merge-track` `md-to-docx`
`renumber-fig` `scan-sensitive` `slim` `snapshot` `template` `text-fmt` `track` `verbs`

本表由 `python3 tools/cli_surface.py` 的输出派生（2026-08-01 核对）；**SSOT 是那条命令，不是本表**。

**SSOT 索引**：

- 子命令能力清单：`~/Dev/tools/dev/lib/tools/report/hq_capabilities.yaml` → `doctools.sub_capabilities`
- 样式族 profile：`config/styles_registry.yaml`（zdwp / eco-flow / generic）
- JSON schemas：`schemas/{plan,decision,patch}.schema.json`
- spec 示例：`config/spec-examples/`（bid.yaml / report-generic.yaml）

**调用示例**：

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

## 2026-05-25 qual-supply docx 脚本族 distill 落地

26+ qual-supply 项目脚本 distill 上提到本 package，让 eco-flow / shoreline / reclaim 等水利项目用 CLI 直接调用，不重复造轮子。
distill 详情：`handoffs/_archive/2026-07-02-2026-05-25-qual-supply-distill.md`。

水利项目接入（eco-flow 范例）：写项目 styles yaml 或扩 `config/styles_registry.yaml` 加 profile →
`docx_cli.py style body <docx> --profile eco-flow` 即套对样式族。
