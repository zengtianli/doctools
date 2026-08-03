# doctools

文档处理与数据转换工具集，从 scripts 仓库拆分。**入口是 `/docx` skill 与命令行**（`raycast/` 整子树 2026-07-27 已归档：`commands/` 早已是空的，9 个 `doc_*.sh` 在自己的 `_archive/` 里）。

> **Python venv**：共享于 `~/Dev/.venv`（uv workspace member · 见 `~/Dev/CLAUDE.md` § uv workspace）。本 repo 不建独立 `.venv`。改 deps → 改 `pyproject.toml` + `cd ~/Dev && uv sync`。

## `doctools` 命令：装出来的入口与绝对路径入口**并存**（2026-08-02 立）

`cd ~/Dev && uv sync --all-packages` 会把本 repo 按 editable 装成包，落一个
`~/Dev/.venv/bin/doctools`。它**不替代**任何东西：

| 敲什么 | 进哪 |
|---|---|
| `doctools <sub> …` | `src/doctools/cli.py` → 同一个 `docx_cli.main()` |
| `python3 <abs>/scripts/document/docx_cli.py <sub> …` | 同一个 `docx_cli.main()` |

~/Work 有 130 处绝对路径在消费后者，**一处都不用改**。`src/doctools/cli.py` 里没有任何
解析或分发逻辑（只做 spec_from_file_location + 调 main），所以两条入口在结构上不可能
行为漂移 —— 实测同一条 `audit headings` 两边 stdout 逐字节相同。

**版本号 SSOT = `src/doctools/__init__.py` 的 `__version__`，全仓只此一份。**
`pyproject.toml` 用 `dynamic = ["version"]` + `[tool.hatch.version] path=` 从该文件读走，
`docx_cli.py --version` 也读该文件。**禁止在 `pyproject.toml` 里补 `version = "…"`**。
实证：把它改成 `9.9.9` 后，包元数据 / `doctools --version` / 老路径 `--version` 三处
一起变；改成非 PEP440 的 `9.9.9-probe` 会让 hatchling 构建直接失败（说明它真在读这个文件）。
改完版本号要让已装元数据跟上，跑
`uv pip install -e tools/doctools --no-deps --reinstall`（`uv sync` 不重建 dist-info）。

⚠ **`--version` 故意不是 argparse 参数**，写在 `main()` 里手动拦第 0 位。因为
`tools/cli_surface.py` 是逐 action 比对整棵子命令树的等价闸门，多注册一个 root action
就会改指纹。实证：临时改成 `p.add_argument("--version", action="version", …)` 之后
surface diff 立刻转红并点出 `_VersionAction` 那一节。要加就得先决定是否更新基线，
别顺手塞。（`--help` 里靠 epilog 一行告知，epilog 不进指纹。）

## 目录结构

```
src/doctools/         # 可安装包壳：__version__ SSOT + console_script 入口（无业务逻辑）
                      # 用 src-layout 是必需的：包放仓根时 editable 安装会把**整个仓根**
                      # 写进共享 venv 的 .pth，凭空多出 lib/scripts/tools/config 等顶层名
scripts/
├── document/ (17)    # 文档处理入口（docx_cli/typeset/revise/bid_gate/renum/docx_fmt/md_tools/pptx_cli/pdf_cli/chart…）
│   └── sub/ (44)     # docx_cli 子命令实现（_groups 声明表 + _function_axis 职能表 +
│                     #   业务模块 + _cli_common 样板；
│                     #   2026-07-31 家族折叠：strip/audit/freeze/image/table/caption/chapter/
│                     #   renumber/blocks/split/typeset_ops/biddiff 12 族 42 旧件并成 12 个子命令族文件；
│                     #   同日入口层折叠：chapter_renumber/tabfig_align/docx_renumber_figures →
│                     #   renum.py，docx_apply_template/format_clone/font_normalize/text_formatter →
│                     #   docx_fmt.py，docx_apply_image_caption 平移进 sub/；
│                     #   同日再折：pptx 4→1(pptx_cli) · pdf 3→1(pdf_cli) · md_docx_template 并入 md_tools）
└── data/ (5)         # 数据转换（xlsx_ + convert）

lib/                  # 公共模块
├── docx_xml.py       # DOCX XML 元素级遍历（全文改写必走，见下节）
├── docx_safe_save.py # surgical 收口（python-docx 存盘炸开面 60→1）
├── docx_parts.py     # 部件完整性断言（zipfile 路线必挂）
├── docx_surgical.py  # zipfile+lxml 手术引擎
├── docx_revise.py    # 修订注入引擎（w:ins/w:del+批注）
├── llm_client.py     # AI 调用（claude -p 封装）
├── cn_number.py      # 中文数字→int 唯一实现（纯 stdlib，见下节）
├── caption_re.py     # 题注/图表编号判据唯一实现（纯 stdlib，见下节）
├── styles.py · chapter_numbering.py · text_fixes.py · schemas.py · soffice.py
├── clipboard.py · progress.py
└── common.sh         # Shell 公共函数
```

## 独立入口脚本（不走 `docx_cli`，直接敲）

这几个不在 `docx_cli` 的命令表里、也没有 skill 指向，但**是活的**。登记在这里有两个作用：
让人找得到；让 `tools/script_graph.py` 把它们算作「有文档点名」而不是孤儿。

| 脚本 | 干什么 | 怎么跑 |
|---|---|---|
| `scripts/document/bid_gate.py` | 标书终稿门检族（2026-07-31 原 bid_final/bid_residue_scan/bid_finalize_sweep/bid_identity_gate/bid_print_ready/bid_deref 6 件合并；检测逻辑 SSOT 仍在 `bid_residue_lib.py`）。子命令：`run`=四门 driver（原 bid_final）· `scan`/`sweep`/`identity`/`print`=单门 · `deref`=交叉引用去耦合（原 bid_deref） | `python3 scripts/document/bid_gate.py run <docx> --mode main\|pei [--rules Y] [--apply]`；单门 `bid_gate.py scan\|sweep\|identity\|print <docx>`；去耦合 `bid_gate.py deref <docx> --check` |
| `scripts/document/md_to_audiobook.py` | md → 有声书（edge-tts，章节并发） | `uv run scripts/document/md_to_audiobook.py <md>`（PEP-723 自带依赖） |
| `scripts/document/docx_revise.py` | **修订注入：意见=ops.yaml 数据，禁在项目里现编注入脚本**（w:ins/w:del+批注；锚点唯一命中 fail-closed；引擎 `lib/docx_revise.py`） | `python3 scripts/document/docx_revise.py <ops.yaml> [--dry-run]`（写法 `config/spec-examples/revise-ops-example.yaml`） |
| `scripts/document/renum.py` | 编号/题注位移与重排族（2026-07-31 原 chapter_renumber/tabfig_align/docx_renumber_figures 3 件合并）。子命令：`chapter`=md 侧章号位移引擎（config 驱动，/renumber skill 指向）· `tabfig`=md 侧 表/图 题注号对齐（--check 机检门 exit 2）· `figures`=docx 图号重排+引用同步（docx_cli `renumber-fig` 即它）。**三个子命令统一 exit 3 = 枚举为空**，见下节 | `python3 scripts/document/renum.py chapter <chapters.yaml> [--apply]`；`renum.py tabfig <yaml\|目录> [--apply\|--check]`；`renum.py figures <docx> [--cn-section --kind 图\|表] [--dry-run\|--inplace]` |
| `scripts/document/docx_fmt.py` | docx 版式/字体/文本规范化族（2026-07-31 原 docx_apply_template/docx_format_clone/docx_font_normalize/docx_text_formatter 4 件合并）。子命令：`template`=套模板+样式清理（docx_cli `template`）· `clone`=版式提取/外壳克隆复刻（docx_cli `format`）· `fonts`=去等线（docx-font-guard hook 指向）· `text`=引号/标点/单位规范化（docx_cli `text-fmt`） | `python3 scripts/document/docx_fmt.py template <docx> [-t 模板]`；`docx_fmt.py clone extract\|apply …`；`docx_fmt.py fonts <docx...> --check\|--apply`；`docx_fmt.py text [flags] <docx...>` |

**加新的独立入口脚本 → 必须在这张表里加一行**，否则它对全仓不可见：`script_graph` 会把它
判成孤儿，下次清理就把它清了。

## 子命令：加子命令 = 加一行数据，不是加一个文件（2026-07-30 立）

`docx_cli.py` 的 126 个 parser 节点（2026-07-30 退役 bullet/quality-check/review、2026-08-01 加只读的 `verbs` 后）。
**三个数都对，报的时候必须说明口径**：`cli_surface` 数的是节点总数 126（含 audit / strip 这类组父节点）；
叶子路径 101；按 `id(parser)` 去重后 93 条可跑动词（其余 8 条是别名，见下节职能轴）。
声明全在 **`scripts/document/sub/_groups.py` 的 `GROUPS` 表**。
原来是 20 个只做「argparse 声明 + `exec_script` 转发」的 group 模块（2265 行），已折成一张表。

`Opt` 里同时写「怎么声明给 argparse」和「怎么转发给实现脚本」，**这两件事必须挨着** ——
原来它们分居 `register()` 与 `_run()` 两处，于是「加了选项忘了转发」是这类壳最常见的 bug：
用户传了参数、脚本收不到、静默按默认值跑。

改这张表**必须三道闸门都跑**：

```bash
python3 tools/cli_surface.py > /tmp/a.json        # 改之前
python3 tools/cli_surface.py > /tmp/b.json        # 改之后 → diff 必须空
python3 tools/cli_forward_probe.py                # 每条子命令真正转发出去的 argv
python3 tools/check_function_axis.py              # 每条子命令的职能标签（缺一条即红）
```

**只跑第一个不够**：它证明接口没变，证明不了参数传过去还是原样 —— 「命令能敲、
跑出来的东西不对」这一类改动它完全看不见。`cli_forward_probe.py` 拦掉 `exec_script`
把 argv 录下来，那才是折叠真正动到的东西。

（`cli_surface` 加上「子命令组的 dest/required/metavar」之后，当场抓到 `header-footer`
的 metavar 被从 `<action>` 写成了 `<target>` —— 这种差异不报错、只在 `--help` 里显示不同。）

自带业务逻辑的模块（`outline` / `styles` / `health` / `slim` / `fix_styleset` /
`docx_para` / `combine` / `md_merge_track` / `audit_styleset` / `chapters_sync` /
`health_split` / `pipeline`）**不在表里**，各自 `register()`。

**`docx_tools.py` 已拆薄（2026-07-31 P2）**：extract / check / track-changes 三段实现在
`sub/docx_{extract,check,track}.py`（各自带 main() 可独立敲，argparse 声明在各自
`add_*_parser()` 只写一遍），`docx_tools.py` 收薄为组合入口（batch 并行层 + 组合 argparse
+ library re-export——外部有人 `from docx_tools import extract_paragraphs` 当库用，名字面
只增不减）。这三段**也不进 `_groups` 表**：docx_cli 侧 extract/read/check/snapshot/compare/
diff/track 走 CMD_TABLE fast-path（带 aliases 与前置 token 注入），表模型装不下；docx_cli
转发仍指 docx_tools，surface/probe 指纹按构造不变。改 CLI 契约（子命令名/flag/默认值）前
先看外部调用面：cc-home `commands/docx.md`、zdys SKILL.md、`~/Work/CLAUDE.md` L161（track
review 的 include_ins+strict 默认全开是机器强制条款，禁改）。

### `track compare` 落地（2026-08-03 · 有回归门）

它此前是**假装成功的空桩**：`print("compare 功能将在 v2 实现。")`，parser 里却好好地
声明着三个参数。前四道闸门（surface / probe / function_axis / smoke）全绿 —— smoke 那条
只跑 `track read`，因为 `track` 是 CMD_TABLE fast-path，**它的子命令根本不进指纹**
（`cli_surface` 里 `track` 节点只有一个 `nargs="..."` 的 `rest`）。所以这类「二级子命令
是空桩」在本仓的闸门体系里是**结构性盲区**，只能靠真敲。

```bash
python3 scripts/document/docx_cli.py track compare 原.docx 改.docx -o 修订.docx [-a 作者]
```

| | |
|---|---|
| 粒度 | **段落级**：改过的段 = 整段 `w:del` + 整段 `w:ins`，不做 run 级字符 diff |
| 范围 | body 顶层 `w:p`（difflib 对齐，键 = pStyle + 可见文本，`w:del` 内的字不参与） |
| 做法 | surgical：以**原稿**的 zip 为底只重写 `word/document.xml`，其余部件逐字节 verbatim + `assert_parts_intact` |
| 回归门 | `python3 -m pytest scripts/document/tests/test_track_compare.py`（13 条） |

**退出码不是只有 0**（照抄 `expect_rc=0` 会误判）：`0` 已产出修订件 · `1` 段落级无差异
（**不产出文件** —— 产一个没有修订标记的副本就是另一种假装成功）· `2` 输入不合法 /
一边 0 段落（空集不报绿）· `3` 已产出，但有本引擎标不了的**范围外差异**。

范围外 = 表格内 / 页眉页脚 / 脚注尾注 / 改稿新增段里引用改稿 rels 的图（那种图照搬进
原稿包 = Word 开门就报损坏，故摘掉）。这些一律打到 stderr 并把 rc 顶成 3，**不静默吞**。

删除侧复用总部引擎 `lib/docx_revise.tracked_delete_runs`，为它加了
`include_nontext=True`（默认值不变，历史调用方一个字节不受影响）：整段删除时**图片 run
也要包进 `w:del`**，否则接受修订后「段落没了图还在」，rc 照样 0。

## 职能轴：每条子命令碰的是「什么」（2026-08-01 立 · 有闸门）

目录轴（`sub/` 里文件怎么摆）回答「实现在哪」；**职能轴**回答「这条命令碰的是什么」。
两条轴正交，所以职能不落成目录，落成一张表 + 一道机检：

| 干什么 | 敲什么 |
|---|---|
| 表本体（一条子命令一行） | `scripts/document/sub/_function_axis.py` 的 `_ROWS` |
| 对账闸门（fail-closed） | `python3 tools/check_function_axis.py` |
| 只看某一职能 | `python3 tools/check_function_axis.py --fn format`（`--md` 出可贴文档的表） |
| 用户侧只读入口 | `python3 scripts/document/docx_cli.py verbs [--fn format] [--json]` |

职能只有 6 个值，**别自创第七个**：`format`（只碰版式：样式/字体/段落属性/页眉页脚/
分节/编号题注号/大纲层级/表格边框/装帧/样式集）· `content`（碰正文写了什么）·
`review`（w:ins/w:del/批注/版本对照）· `inspect`（只读不写盘）· `convert`（跨格式）·
`dispatch`（调度编排）。当前分布 format 41 · content 15 · review 6 · inspect 28 ·
convert 2 · dispatch 1 = 93 条（另 8 条别名共用同一 parser，不单独占表项：
`read`=extract · `diff`=compare · `styleset *`=audit-styleset *）。

**加子命令 = `_ROWS` 里也加一行**，否则闸门转红并点名。反过来，表里留了 CLI 已经
没有的条目同样判红 —— 两个方向都堵死，这张表才不会变成「还在、只是不准了」。
别名不占表项：闸门按 `id(parser)` 归组自动认领，别名清单是从真 parser 派生的，不手抄。

## 冒烟轴：每条子命令**真敲一遍**（2026-08-02 立 · 有闸门）

前面那几道闸门（`cli_surface` 接口指纹 · `cli_forward_probe` 转发 argv ·
`check_function_axis` 职能标签）**没有一道真的执行过任何一条子命令**。三道全绿，
仍然可能每一条敲下去都是坏的。立表当天实跑 93 条就撞见一例：
`chapter delete --prefix` 声明了、也忠实转发了，可 `chapter.py` 只认
`--h1/--h1-text` —— 这个选项从写下那天起不可能生效，前三道闸门全绿。

| 干什么 | 敲什么 |
|---|---|
| 表本体（一条动词一行：argv / 预期 rc / 改不改源件 / 备注 / skip 理由） | `tests/smoke/_verb_specs.py` 的 `_ROWS` |
| 真跑（一条动词一个用例，各自独享新鲜 fixture 副本） | `python3 -m pytest tests/smoke -q` |
| 覆盖闸门（fail-closed，两个方向都堵） | `python3 tools/check_smoke_coverage.py` |
| 看表（含 skip 理由 / 可贴文档） | `check_smoke_coverage.py --list` · `--md` |
| fixture 造件器 **+ 自带富度门** | `python3 tests/smoke/_fixture.py <目录>`（造完当场 `selfcheck`，10 类对象计数低于下界即 **exit 1**） |

### fixture 的富度是机器判据，不是注释（2026-08-03 立）

`mutates` 抓的是「写盘动词其实没写」，但它抓不到**上一层**的空跑：动词写没写盘，
取决于 fixture 里有没有它认得的对象。2026-08-03 实测出的四处零对象面：

| 症状 | 后果（rc 全是 0，看输出也像干了活） |
|---|---|
| 6 个题注段全套 `Normal` | `renumber h4-figures` 打 `图=0 表=0`；`caption` 族与 `style caption` 一个对象都扫不到 |
| 题注短横只有 ASCII 一种 | `caption_re` 声明认 5 种，只测 1 种；「退回只认 ASCII」这类回归（2026-08-02 `strip.py` 真的发生过）冒烟看不见 |
| 题注全是两段式 `图1-1` | `SECTIONED_CAPTION`（`style caption` / `strip outlinelvl` 的写盘范围，只认三段式）恒为空集 |
| 中文题注再多也喂不到 `renumber-fig` 默认模式 | 中英两条编号线在 `caption_re` 里零共享，默认模式只认 `Figure N` |
| 草稿写成 `## 意见 N` | `revise-rules gen` 的 `SECTION_RE` 只认 `## 改动 N` → `parse_md` 直接 `return []` → 写出一个 `[]` 还 rc=0 |

所以 `_fixture.py` 自带 `selfcheck()`：题注样式 / 五种短横 / 三段式 / `outlineLvl`
病灶 / 内嵌图 / 英文图题 / 草稿节数逐条数，低于下界即 exit 1。
**下界留了余量**（加东西不会碰它，砍掉一整类对象才会）。

⚠ 两处判据看着像洁癖，其实都是被实现逼出来的，别顺手"统一"掉：

- **图题用 `Image Caption`、表题用 `表题`，两个名字不能合并成一个。**
  `styles_registry.yaml` 把 `Caption` 同时列进 FIG 与 TABLE 两族，而
  `styles._build_h4fig_plan` 是 `elif` 链、先判图后判表 —— 表题若也叫 `Caption`
  会被图分支先认领、再被 `startswith("图")` 否掉，`plan_tbl` 恒为 0。
  另外 `caption_re.is_fig_caption_style("Caption")` 是 **False**，用它会让
  `shape_contract.caption_figure_count` 恒为 0。
- **英文图题号故意是乱的（`Figure 3` 在 `Figure 1` 前面）。** 顺序对了 remap 就是
  恒等映射，`renumber-fig --inplace` 跑完源件一个字节不变，`mutates` 断言等于没写。

当前 **真跑 91 条 / skip 2 条 / 共 93 条**（skip：`scan-sensitive` 要调 LLM、
`para render --warm` 要起 soffice 并写 `~/.cache/`；两条都实测能跑，理由写在表里且
`-rs` 会打印出来）。

**核心断言是 `mutates` 那一列，它是双向的**：`mutates=False` 跑完源 docx 的 md5
必须没变（抓「只读动词偷偷写盘」），`mutates=True` 必须变了（抓「写盘动词其实没写」，
即静默空跑 —— rc 照样 0，看输出也像干了活）。注意它问的是**源件**变没变：
`template` / `audit images` / `table extract` 都写盘，但写的是 `-o`/`--report`/`--out-dir`，
源件必须一字节不动。

⚠ **`expect_rc` 不都是 0，且与 fixture 内容绑定**：`audit-styleset` 三条 rc=severity(fail=1)、
`health diagnose|full|gate` rc=健康度/gate 结果(0/1/2)、`para scan-ppr` rc=`3 if suspects else 0`、
`para edit` 与 `para fix-ppr` rc=4。换 fixture 就要重新实测这一列，**不能照抄邻居**。
另外 `health diagnose|gate` 的 rc=2 与 docx_cli 的「参数错误 rc=2」撞码，runner 不能把 2 一律当失败。

#### 2026-08-03 换 fixture 之后整列重测的结果（10 行变了，逐条真敲测出来的）

用一次性驱动（与 runner 同一造件路径、同一 copytree、同一 subprocess 参数）跑完 91 条，
只有这 10 条与旧声明不符 —— **全部属于「fixture 变强 → 动词现在有对象可处理」，
没有一条是 fixture 揭出来的崩溃**：

| 变的是什么 | 哪几条 | 实测证据 |
|---|---|---|
| `mutates` **False → True**（7 条） | `style caption` · `renumber h4-figures` · `renumber-fig` · `caption number` · `caption number-by-style` · `outline normalize-arabic` · `strip outlinelvl` | 分别打出 `tables=1 figures=1` / `图=4 表=3` / `remap {3:1,1:2}` / `chapters_detected=[4]` / `chapters=[1,2,3,4]` / `change_count=1` / `processed=2 removed=2`，旧 fixture 上全是 0 |
| `expect_rc` **0 → 2**（1 条） | `health gate` | FAIL，`failed=['caption-table-pairing','caption-count-consistency']` |
| `expect_rc` **0 → 4**（2 条） | `para edit` · `para fix-ppr` | 编辑本身成功（stdout 有 `EDITED PARA 30 …` / `FIXPPR PARA 12 …`），4 来自写盘后 `_finish_gate()` 复跑 health gate 拿到 FAIL |

**这三条 rc 编码的是同一个 fixture 事实**，会一起随 fixture 翻。`health gate` 的 FAIL 是
**真判定不是回归**：旧 fixture 的 PASS 是假绿 —— 那两个 check 都有「零 style 题注 → SKIP」
守卫，题注全 `Normal` 时压根没跑。现在报的是「图: 编号集 4 ≠ caption 段 7」「orphan_captions=1」，
而差额那几段正是 `caption number` 唯一的作业对象。**要把 rc 变回 0 就得删掉那些无编号题注 =
把 `caption number` 重新变成空跑**，所以选择保留对象、如实记 FAIL。

`para` 那两条**没有**加 `--no-gate` 把 4 抹平成 0（该 flag 存在、实测 `GATE SKIP (--no-gate)`
+ rc=0 + 源件仍变）：加了就再没有任何一条 smoke 覆盖 `_finish_gate` 的 FAIL 分支，
而那是这两条写盘动词唯一的安全网。

⚠ **`renumber-fig` 的 `renum.EMPTY_RC=3` 分支现在没有 smoke 覆盖**：加强后的 fixture
连 `--cn-section --kind 表 --check` 都非空（实测 rc=2「重复号 `{'1': [1]}`」）。
要覆盖 3 必须另造一份零题注 fixture —— **不能给 `renumber-fig` 加第二行**，
`_verb_specs._build()` 会判重复动词。

⚠ **`mutates` 是布尔的，看不见「部分退化」**：2026-08-03 把 `caption_re.DASH_CHARS`
注回只认 ASCII（本仓真发生过的那类回归），`strip outlinelvl` 的 `processed` 从 2 掉到 1，
**md5 照样变、smoke 全绿 92 passed**；红的是 `test_caption_re.py`（15 failed）。
所以那道 pytest 门不能省，冒烟替不了它。

**加子命令 = 这里也加一行**（连同 `_function_axis` 那行，一共三处），否则闸门转红并点名；
表里留了 CLI 已经没有的动词同样判红。`_verb_specs.py` 用 tuple-of-tuples 而非 dict 字面量，
理由与 `_groups` / `_function_axis` 相同：dict 里写重复键不报错、后一条静默覆盖前一条，
而那正是「覆盖率看起来是满的、实际少跑一条」的来源。

### 装出去的 wheel 必须自包含（2026-08-02 立 · 有门）

`scripts/` `lib/` **不在 `src/doctools/` 里**，而 `~/Work` 有 130 处绝对路径钉着它们的
现有位置，所以不能搬。改用 pyproject 的 `force-include` 在**构建时**把
`scripts/ lib/ config/ schemas/` 四个根镜像进 `doctools/_bundled/`，
**工作树一个字节都不动**。`src/doctools/cli.py` 按顺序试两个实现根：
工作树（editable 装）→ `_bundled/`（wheel 装）。

镜像时**相对深度与工作树完全一致**，所以全仓 40+ 处 `parents[2]/"lib"` /
`parents[3]/"lib"` 的层数算术在包内原样成立，一处都不用改。

| | |
|---|---|
| 分发能力对账门（构建可移植 + 中立 HOME 下实测可用动词数 == 声明值） | `python3 tools/check_wheel_selfcontained.py` |

**wheel 有条件可分发**（2026-08-03 实测 **49/49 全起得来**；同日中途 48/49、08-02 是 47/49、08-02 上午还是 0/49）。

⚠ 这个 49 是**按工作树快照**测的，不是按 HEAD。本门 `git archive` 的 ref 写死是 `HEAD`，
所以 image-caption / text-fmt / pyproject 三处改动 commit 之前直接跑它必红（rc=2，
`hq-devlib was not found in the package registry` —— HEAD 的 pyproject 还是裸 specifier）。
**那是时序不是回归。** 攒证据的正确姿势是把工作树做成悬空提交再让本门按那个 ref 取源：
`GIT_INDEX_FILE=<临时> git read-tree HEAD && git add -A && git write-tree` → `git commit-tree`，
全程不碰分支 / 真 index / 工作树。

⚠ **「有条件」三个字不能省，但条件已经从「摆两个 wheel」换成「有私库读权限」**
（2026-08-03 改）。原来 `hq-devlib>=0.1` 是普通 specifier，而它没发到任何公开
index，于是**只拿 doctools 的 wheel 会直接解析失败**（`Because hq-devlib was not
found in the package registry …`）—— 不是「装上后几条动词不能用」，是 pip 一步都
进不去；CI 的 `uv sync` 死在同一处。现在改成 **PEP 508 git 直接引用**
（用户拍板：走 git URL，不发 PyPI）：

```toml
"hq-devlib @ git+https://github.com/zengtianli/devtools.git",
```

三件必须一起记住的事：

| | |
|---|---|
| **不带 `#subdirectory=`** | `zengtianli/devtools` 这个仓的**根就是** `~/Dev/tools/dev`，pyproject 在仓根。写 `#subdirectory=tools/dev` 会 clone 完当场 `has no subdirectory tools/dev`（实测） |
| **`[tool.hatch.metadata] allow-direct-references = true` 必须跟着加** | 否则 hatchling 把直接引用判成硬错误，连本机 `uv sync` 都过不去（`cannot be a direct reference unless …`，实测）。它默认关着是因为**带直接引用的包 PyPI 拒收** —— 也就是说这条路和「发 index」互斥，哪天要发得两处一起改回去 |
| **`zengtianli/devtools` 是 PRIVATE** | 所以这条 URL 只在**有该仓读权限**的环境里成立：本机（git 常规凭证）✓ · CI（复用已有 secret `HQ_DEVTOOLS_TOKEN` 配 `url.insteadOf`，见 `.github/workflows/gates.yml`）✓ · 没凭证的第三方 ✗。要让任意第三方装得动，只有把 devtools 转 public 这一条路 |

**本机行为一个字节没变**：workspace 根 `~/Dev/pyproject.toml` 的
`[tool.uv.sources] hq-devlib = { workspace = true }` 覆盖掉这条 URL。实测
`cd ~/Dev && uv sync --all-packages` rc=0、`~/Dev/uv.lock` 改前改后**逐字节相同**
（仍是 `source = { virtual = "tools/dev" }`），且 `uv sync --all-packages --offline`
照样 rc=0 —— 根本不走网络。
另：`text-fmt` 的 `--help` 原来 rc=2（打自己那份「未知参数」清单），2026-08-03 已修 ——
`docx_fmt.py::text_main()` 开头拦 `-h/--help` 打 `TEXT_USAGE` 并 `return 0`。拦截**必须在
`--scope` 取值与未知 flag 判定之前**，否则 `--scope --help` 会被 `next(_it)` 先吃掉；
`-h` 也一并认了 —— 旧判据只看 `startswith("--")`，`-h` 会被当文件名吞进 `get_input_files()`。
至此 49 条动词的 `--help` 全部 rc=0，`DECLARED_WORKING = 49`。

**`image-caption` 2026-08-03 已修**（原 rc=1，与 text-fmt 并列在这里）。根因不是分发：
它**没有 argparse**，main() 把 argv 直接丢给 `get_input_files()`，而后者见「一个位置参数
都没有」就**回落去读 Finder 当前选中项**并按写模式处理 —— 于是 `--help` 会去动用户此刻
选中的文件。这正是本文「破坏性动作必须自己占一个动词」一节判过死刑的那条
（「`--help` 弹 Finder + 往选中文件写盘，曾写进 ~/Work 在跑的项目」，判据 =
**未知 flag 一律 `sys.exit(2)`，禁 fallthrough**），当时的漏网点。

修法与**位置**都要照抄：闸门在 `sub/docx_apply_image_caption.py` **模块顶部、
`from docx import ...` 之前**，不是在 `main()` 里。因为 docx_cli 的 `_exec_script` 走
`spec_from_file_location` + `exec_module` 再调 `main()`，只在 main() 里拦，重家伙在
模块加载阶段就已经 import 完了。顶层那道用 `_invoked_as_cli()` 把自己关掉，否则
`typeset_apply` 的 `load_step()`（同样是 spec 载入，只为拿 `apply()`，此刻 `sys.argv`
是 typeset 自己的、带 `--dry-run`）一进这步就会 exit(2)。main() 里还留了同一个纯函数
作第二道，兜 `sys.modules` 已缓存、顶层不再执行的第二次转发。

⚠ 连带行为变更：`image-caption <docx> --dry-run` 从「**静默忽略该 flag、照常写盘**」
变成 rc=2。本脚本从来没实现过 dry-run，旧行为是假装接受。

本仓运行时依赖总部
`~/Dev/tools/dev/lib/` 的 6 个**平铺模块**（finder / file_ops / display /
parallel_contract / usage_log / env，共 1156 行）。它们原来不是包、只能靠 sys.path
注入导入，所以「声明成 dependency」这条路当时走不通；现在总部仓把这 6 个打成了
分发名 **`hq-devlib`**（`lib/*.py` 零改动、44 处 sys.path 注入原样继续工作），
本仓 `pyproject.toml` 声明 `hq-devlib @ git+https://github.com/zengtianli/devtools.git`
（2026-08-03 从 `hq-devlib>=0.1` 改，理由见上一节），workspace 根的
`[tool.uv.sources]` 把它覆盖回 member `tools/dev`。
**本机行为一个字节没变**：`hq-devlib` 在 workspace 里是
virtual，`uv sync` 不会把它装进共享的 `~/Dev/.venv`（装了也只是多一条排在
sys.path 注入之后的来源）。

⚠ 在本机验「别的机器能不能用」时**必须中和 `$HOME`**：`docx_cli.py` 有一句
`Path.home()/"Dev"/"tools"/"dev"/"lib"` 兜底导入，不中和的话 0/49 的 wheel 也能
测成 49/49 —— 2026-08-02 上午就是这么报出一次假绿的。这门自己中和（`HOME` 指空
目录 + 清 `PYTHONPATH`），别绕过它手验。

⚠ 判据 B **带依赖装**（`pip install <whl>`，不是 `--no-deps`）。`--no-deps` 装法下
这门永远只能测出 0，而且卡点会随修复一路搬家（parallel_contract → yaml → lxml），
把「依赖声明对不对」这个真问题挡在门外。

⚠ 门里原来还有一段「把兄弟目录 `../dev` 构建成 wheel 丢进 `--find-links` 目录当
index 替身」，**2026-08-03 已删**：hq-devlib 改成 git 直接引用之后，直接引用的优先级
高于任何 index / find-links，那个目录**永远不会被看一眼**（实测：给了 find-links，
装出来的仍是 `+ hq-devlib==0.1.0 (from git+https://github.com/zengtianli/devtools.git@…)`）。
留着它只会让这门看起来在测「index 上取得到」，而它实际测的是「git URL clone 得动」。
代价是本门现在**要能上网 + 要有 devtools 私库读权限**，两者缺一即 `SystemExit(2)`
（fail-closed，不退化成假绿；反向验证：把 URL 换成不存在的仓，本门实测 rc=2 并打出
`Repository not found`）。装包那一步走真实环境，跑动词时 `HOME`/`PYTHONPATH` 照样中立。
| 只看实测能力不判定（改 `DECLARED_WORKING` 之前先跑它） | `python3 tools/check_wheel_selfcontained.py --scan` |

（`--tier struct` / `--tier full` 是 2026-08-02 写进本文却从未存在过的 flag，
已按实际 `--help` 更正。）

**「两份副本会不会漂」的答案是「仓库里根本没有第二份」**：`_bundled/` 只存在于构建
产物中，是构建那一刻从工作树 checkout 的逐字节快照。门里那条等式
（包内文件集 == `git ls-files scripts lib config schemas`，且逐文件 sha256 相同）
就是这句话的机器判据，**两个方向都查**（多一个/少一个都判红）。

⚠ `cli_surface` / `cli_forward_probe` **看不见包内副本那条分支** —— 它们在工作树里跑，
命中的永远是第一条。wheel-only 分支唯一的测法就是上面这道门真去装一遍。
2026-08-02 基线实测：加 force-include 之前 wheel 只有 **7 个文件**，
clean venv 装完 `doctools --version` 直接 `FATAL: 找不到实现入口` rc=2。

⚠ `parallel_contract`（总部 SSOT，`~/Dev/tools/dev/lib/`）**不在本仓**，而
`docx_cli.py` 缺它是 **fail-closed exit 2 不是降级** —— 不解决就一条动词都跑不了。
它现在有**三条来源，按这个顺序**：

| # | 来源 | 谁在用 |
|---|---|---|
| 1 | `~/Dev/tools/dev/lib`（`insert(0)`，**目录不存在就不塞**） | 本机所有直接敲绝对路径的调用 |
| 2 | `<根>/lib`（**append 不是 insert(0)**） | 留给包内镜像；本机这里**没有** parallel_contract.py |
| 3 | 装出来的 `hq-devlib` 包（site-packages 顶层） | wheel 装到别的机器上时 |

2 用 append 的理由：`lib/` 下有 styles.py / schemas.py / progress.py 等与他处同名的
模块，顶到 sys.path 首位会改变全进程解析优先级（形状对齐 `scripts/data/data.py`
既有先例）。3 排在最后，所以**本机行为一个字节没变** —— 解析照旧落到 1。
`pdf_cli.py` 同款三来源，但它的 except 分支是**静默降级**（丢 `--batch/--phases/
--defer/--fanout-evidence`，只留 `--workers`），比 exit 2 更难查，别以为它无所谓。

## 全仓脚本关系图 = 一页三视图（2026-08-02 扩）

```bash
python3 tools/script_graph.py --open     # 99 脚本 · 365 引用 · 93 动词 · 0 孤儿
```

| 视图 | 回答什么 | 数据从哪来 |
|---|---|---|
| **图谱** | 谁调谁（力导向，可拖可缩，点节点看双链） | ast：import + 字面量调用 |
| **清单** | 99 个脚本各是什么（层/引擎/格式/行数/入←→出/首行 docstring/接住哪些动词） | 同上 + 首行 docstring |
| **动词** | 93 条 CLI 动词各落到哪个脚本、属哪个职能 | `_groups.Target.impl` + parser 树 `func.__module__` + `CMD_TABLE` 源码 + `_function_axis` |

**「实现脚本」那一列不是手抄的**，也不是猜的：三条来源都读 SSOT 本体，
`Target.chain` 是 `(dest, 同组 target 名)`（**不是脚本名**，照字面收会把 `table borders`
报成实现在 `center.py`），扁平组的 Target 挂在 `g.flat` 不在 `g.targets`（漏掉这支会把
`chrome` / `md-merge` 报成实现在 `_groups.py`）—— 这两个坑本页第一版都踩过。
交叉核验：与 `cli_forward_probe` 实录的转发脚本比对，**65 条重合项 0 不一致**。

⚠ 清单里的**格式轴那一列是信号计数的启发式**（`FORMAT_SIG`），与边和动词映射不同级，
别拿它当事实用；页面顶部也这么标了。

孤儿判据是三条证据全无：代码里没人引用 + 没有文档点名 + 不是顶层入口。
**「代码里没人 import」单独不算死** —— 一多半脚本是文档告诉我去敲的。

## 开发约定

### 引用路径
- `scripts/xxx/` 下的 Python 脚本：`sys.path.insert(0, str(Path(__file__).parent.parent.parent / "lib"))`
- Shell 引用库：`source "$(dirname "$0")/../../lib/common.sh"`
- LLM 调用：`from llm_client import chat`

### 用 python-docx 存盘的脚本必须挂 surgical 收口（2026-07-30 立 · 有守卫）

一行：

```python
import sys as _sys
from pathlib import Path as _Path
_sys.path.append(str(_Path(__file__).resolve().parents[N] / "lib"))   # N = 到仓根的层数
import docx_safe_save  # noqa: E402,F401
```

**append 不是 `insert(0)`** —— `lib/` 和 `scripts/document/sub/` 都有 `styles.py`，插 0 位会顶掉脚本自己那份。

为什么：一份 301 部件 / 75 个嵌入公式的真报告，python-docx **什么都不改**地打开再存回，
**60 个部件字节变了、只有 1 个语义真变了**（59 个纯属被重新序列化）。收口把语义未变的部件
按原字节还原 → 炸开面 60→1，实测「无改动时输出与原件逐字节相同」。
`Document()` 新建文档自动不介入，所以从零造文件的脚本照样加，不碍事。

| 什么时候用什么 | |
|---|---|
| 守卫（枚举全仓，缺收口判红，fail-closed） | `python3 tools/check_docx_collar.py` |
| 量某条命令的炸开面 | `python3 tools/blast_radius.py run <docx> -- <命令，{docx} 占位>` |
| 迁移前后对拍（老 vs 新，判语义等价 + 炸开面收窄） | `python3 tools/blast_radius.py diff <docx> --old '…' --new '…'` |
| 逃生 | `DOCX_GRAFT_OFF=1`（退回裸存盘）· `DOCX_GRAFT_QUIET=1`（不打那行 stderr） |

`blast_radius.py diff` 的判据全部是「老 vs 新」的相对比较，**不是「新 vs 理想值」** ——
后者会连着造出三种假红（详见该文件里「判据的那根轴」）。

### `renum.py` 的枚举为空 = exit 3，不是 exit 0（2026-08-03 立）

`renum.py` 三个子命令原来在**一个对象都没枚举到**时全部打通过语并 exit 0。根因是
连续性判据在空集上**恒真**，于是自检对着「什么都没量到」发合规证：

| 路径 | 空集上原本打的 | 恒真的那个式子 |
|---|---|---|
| `figures`（默认英文线） | `✓ 连续 1..N` | `[] == list(range(1, 1))` |
| `figures --cn-section` | `✓ 每节连续 1..k` | `all(…)` over `{}` |
| `figures --cn-section --check` | `✓ 图序号与居中均合规` | `any(4 个空桶)` = False |
| `tabfig` | `✅ 表/图编号与章号全部对齐` | 循环体一次都没进 |
| `chapter` | `磁盘已与 config 一致 (no-op)` | `all(())`（`sequence: []`） |

现在这五处一律 **exit 3 + `✗ 未发现任何… —— 枚举为空`**（`renum.EMPTY_RC`）。

**3 而不是 2**：本文件的 2 已经被三种「真发现了问题」占着（`--check` 报断号/重号 ·
检测到重复图号 · 重编号后仍不连续）。空集是**没能做出任何判定**，与之正交；混进 2
会让上游把「这份文档没有图题」误读成「图号有问题」，实测两处会真坏 ——
`doc_dispatch.do_renum` 的 `图` 那轮返非 0 就 `break`，`--kind 表` 整轮不跑；
`typeset_pipeline` 步骤③ 见 `rc!=0` 即回滚，把 `--fix-center` 已经写进盘的居中一起撤掉。
两处都已放行 3（后者用 per-step `keep_rcs`，对账卡上显示 `○ 空集(rc=3)` 而非 `✅ 完成`）。
选 3 也对齐本仓既有先例：`para scan-ppr` 用 3 表示「非错误、但需要人看」。

⚠ **写盘与空集正交**：`--fix-center` 时先写盘再报空集（居中是真活），其余情况在写盘前
就退出 —— 产一份逐字节复制的 `.renumbered.docx` 只会让下游以为「重排过了」。

⚠ 空集判据必须问**枚举端**不能问**计划端**：`renumber_cn_section` 为此多返一个
`n_found`。`len(plan)` 在 `--no-supplement` / 无号题注定位不到节号时同样是 0，
「枚举到了但没排进计划」与「一个题注都没有」在 plan 上长得一模一样。

⚠ `tests/smoke/_verb_specs.py:137` 的 `renumber-fig` 那行 `expect_rc=0` 是**照着这个
空集绿写下来的**（备注还写着「本 fixture 图题数=0」）—— 冒烟表把一处 fail-open
固化成了预期。该行需改成 3；本轮没动它，`pytest tests/smoke` 会红这一条。

### 题注/图表编号判据必走 `lib/caption_re.py`（2026-08-01 立 · 有回归门）

**别再手写 `图\s*\d+[-–—]\d+` 这类正则。** 合并前全仓 12 个文件 ~30 条 pattern、
20 个互相竞争的判据点，光「短横」就有 **8 种互不相同的字符子集**，没有一处是全集
（`bid_gate` 有 U+2011 无全角 `－`，`renum` 恰好反过来，两个门互为盲区）。实测后果：
`renum figures` 把已编号的 `图1‑2` 当无号题注 prepend 出 `图1-3 图1‑2 …`，而自检用
**同一条瞎正则**读回，对着写坏的文档打印「✓ 每节连续」。

```python
from caption_re import parse, finditer, pattern, RENUM_CN_CAPTION
n = parse(text, RENUM_CN_CAPTION.for_kind("图"))   # n.section / n.seq / n.raw / n.appendix
pattern(PREFIX_STRIP_FIG).sub("", text)            # 也可直接拿 compiled 用
```

**不是一条正则打天下** —— 差异有一半是有意的（`bid_gate` 的右界断言防吃正文内联引用、
`bid_residue` 的非锚定是为扫交叉引用、`styles._CAPTION_PATTERN` 强制三段是写盘范围、
`table`/`image` 只看首字是为给无编号题注命名）。所以形状是「一套字符类 + 一个构造器 +
若干具名 spec」：**加调用点 = 加一个 spec 声明，不是再写一条正则**。

| 什么时候用什么 | |
|---|---|
| 回归门（55 条，含 9 种变异实证能抓） | `python3 -m pytest scripts/document/tests/test_caption_re.py` |
| 想「顺手统一」样式名/关键词三套判据 | **先别** —— 那是给写盘动词扩范围，模块 docstring 写了为什么不合并 |

⚠ `cli_surface` / `cli_forward_probe` **看不见这层**（那两个只管 argv），改判据必须跑
上面那个 pytest；改完还要用真 fixture 走一遍 CLI（`renum figures` / `health diagnose` /
`caption pair` 是三条已知会现形的链）。

**2026-08-02 补一处漏网**：`sub/strip.py` 的 `CAPTION_PATTERN` 与迁移前的
`styles._CAPTION_PATTERN` 逐字节相同却没跟着搬，于是归并那轮**自己造出一处新分歧** ——
strip 只认 ASCII `-`，styles 侧已认 5 种。现共用 `STRIP_OUTLINELVL_CAPTION`
（= `SECTIONED_CAPTION` 同一对象）。实测 `strip outlinelvl` 在 6 种编号的 fixture 上
从 processed 2/7 变 7/7（U+2011 / U+2013 / U+2014 / U+FF0D / 全角句点 5 条原本漏网、
`w:outlineLvl` 继续污染 Word 导航窗格）。教训：**「逐字节相同」正是最容易被跳过的那种，
归并收尾必须按 grep 结果逐条销号，不能靠「看起来都搬完了」**。

### 全文改写 docx 必走 `lib/docx_xml.py` 的元素级遍历（2026-07-26 立）

python-docx 的 `Document.paragraphs` 只认 `w:body/./w:p`、`Paragraph.runs` 只认 `./w:r`，
**静默漏掉审阅修订（`w:ins`/`w:del`/`w:moveFrom`/`w:moveTo`）、`w:hyperlink`、文本框、
嵌套表里的 run**；批注 `comments.xml` / 脚注 / 尾注更是整个 part 碰不到。所以凡是
「把全文某类字符改一遍」的引擎（规范化、替换、术语统一…）都用：

```python
from docx_xml import iter_text_roots, iter_paragraphs, para_own_runs, run_text, set_run_text
for label, root, flush in iter_text_roots(doc):      # 正文/批注/脚注/尾注/页眉页脚
    for p in iter_paragraphs(root):
        for r in para_own_runs(p):
            set_run_text(r, fix(run_text(r)))
    if flush: flush()                                 # 脚注/尾注是通用 Part,必须回写 blob
```

删除态（`w:del`/`w:moveFrom`）的文本载体是 `w:delText`，写成 `w:t` = 把删掉的字变回正文；
`set_run_text` / `text_tag_for` 已挡住这条。回归门：
`scripts/document/tests/test_docx_text_formatter_scopes.py`（老实现在此测试下 9 红）。

### 中文数字转 int 只有 `lib/cn_number.py` 一份（2026-08-01 立 · 有测试门）

合并前全仓 **5 处** 各写各的（chapter / outline / blocks / caption / styles），分三档能力，
同一个输入给三种答案：`十六` 在 caption/styles 侧返 None、`一百零五` 在 blocks 侧返 None。
后果不是学术问题 —— `caption number` 的章计数器解析失败就**不换章**，第 16 章往后的表图
继续按上一章编（表15-7、表15-8…）；`styles` 那侧上层写的是 `_parse_chapter_from_text(t) or (chapter + 1)`，
**静默拿「上一章+1」顶上**。三个入口都挂在 `typeset_apply.py` 步骤表里，/typeset 一条龙每次都在跑。

```python
from cn_number import chinese_to_arabic   # 严格版：解析不了抛 ValueError
from cn_number import cn_to_int           # 宽松版：解析不了返 None
```

**两个 API 必须并存，别合成一个** —— chapter/outline 三处靠 `except ValueError` 控流
（收 None 会拿着 None 往下算、写出「None、标题」且不报错），blocks/caption/styles 三处靠
None 分支（抛异常会直接崩）。`cn_to_int` 就是 `chinese_to_arabic` 外包一层 try，语义不会再分叉。

能力上界 = 十/百/千 + 「十X」省略一 + 〇/两 + 阿拉伯直通；**「万」不支持是有意的**（旧的三档
没一档支持，加它要改累加器结构）。**本模块不管正则** —— 各调用点的章标题字符类没统一
（caption/styles 侧不含 `百`/`零`，所以「一百零五、」压根匹配不到），那是另一根轴。

回归门：`scripts/document/tests/test_cn_number.py`（含「sub/ 下不许再有本地副本」那条）。

### lsof 占用检查 + `.bak-N-日期` 备份路径只有 `_cli_common` 一份（2026-08-02 立 · 有对拍门）

收敛前全仓 **10 份 `lsof_check`，四种语义，没有任何两份逐字相同**：A 裸 `lsof <path>`（5 份）·
B `lsof -- <path>` 且要求行数>1（3 份）· C 返 bool、不看 returncode、`except Exception` 全吞（1 份）·
D 无 timeout（1 份）。用户拍板**取最健壮的 B**，现在真身在 `sub/_cli_common.py`：

```python
import _cli_common as _cc          # sub/ 自身 append 进 sys.path，别 insert(0)
occ = _cc.lsof_check(path)         # str|None
bak = _cc.find_next_backup(path)   # 只算路径；_cc.make_backup 才 copy2
```

**`--` 不是洁癖**：实测 `lsof -hold.docx` 被解析成 `-h`（帮助）→ **rc=0、stdout 空** →
老的 A/C/D 三派一致判「空闲」，对着 Word 正开着的文件放行写盘。returncode 检查救不了它，
只有 `--` 能。所以别顺手把它删掉当「多余参数」。

⚠ `pipeline_lib.lsof_check` / `make_backup_path` 现在只是**委派壳**，别再去那里找真身 ——
它们留着是因为 `pipeline_lib.__all__` 列着、`typeset_apply.py:98` 按名 import、外仓还有
`import *` 的 shim；**名字面只增不减**。

| 什么时候用什么 | |
|---|---|
| 对拍门（10 处同解 + `--` 生效 + 6 处备份委派 + 内联循环已拆） | `python3 handoffs/_loc_plan_harness/lsof_backup_ab.py` |
| **不归本机制管** | `.bak-<时间戳>` 是**另一套命名约定**（`table` / `normalize_fonts` / `md_merge_impl` / `docx_fmt` / `bid_gate` / `lib/docx_surgical.make_backup`），扫进来就是行为变更 |

`cli_surface` / `cli_forward_probe` **看不见这层**（那两个只管 argv），改判据必须跑上面那个对拍器。

### per-op 选项走后端声明,别写进 Swift（2026-07-26 立）

GUI 的勾选框由 `doc_gui_backend.OPS[<op>]["options"]` + `["option_groups"]` 声明，
Swift 只按 `type` 泛化渲染（`bool`→checkbox），**Sources/ 里不许出现任何 option id 字面量**。
加一个旋钮 = 只改 Python，不重编 `.app`。

```
gui-ops  → {"id":"rule.units","group":"rule","type":"bool","default":true,"title":"…","note":"…"}
gui-run  --op clean --opt rule.units=0 --opt scope.comments=0 --files a.docx
```

契约铁律：未知 key / 非法值一律走信封 `{"ok": false, "error": …}` 且 **exit 0**
（argparse 的 exit 2 + 空 stdout 会让 Swift 只看到「后端崩了」）。
引擎侧配置一律走 `FormatConfig` 冻结 dataclass 显式穿参，**禁再加模块级可变全局**
（旧的 5 个全局只在 `__main__` 赋值，被 import 时永远是默认值、还会跨调用串味）。

⚠ 域过滤与引号规则**不正交**：引号左右方向靠段落级奇偶计数器，被跳过的 run 仍要
推进 counter 但不写回，否则它后面的引号整体反相。回归门在
`tests/test_docx_text_formatter_safety.py::test_skipped_scope_does_not_flip_quote_direction`。

### 破坏性动作必须自己占一个动词（2026-07-26 立）

**不可逆操作不许搭在非破坏性动词的默认行为里**，也不许只靠 flag 才关得掉。2026-07-26 实测的四例：

| 动词叫 | 默认还偷偷干了 | 修法 |
|---|---|---|
| `clean` 规范化 | 删光全部 `headerReference`/`footerReference`（连 typeset 套的院模板 14/13 个也没了） | 拆出独立动词 `stripchrome`，标 `danger` |
| `--help` | 弹 Finder + 往选中文件写盘（曾写进 `~/Work` 在跑的项目） | 未知 flag 一律 `sys.exit(2)`，禁 fallthrough |
| 「统一单位」 | 裸 `str.replace`：小时候→h候 · 毫米波→mm波 · Item2→Item² | 加词边界（中文单位须紧跟数字，ASCII 上标须两侧非字母数字） |
| GUI 未勾的选项 | 当作**显式关闭** | `_clean_flags` 改为「没传 = 保持默认」，只有显式 `=0` 才关 |

判据：**动词名字承诺什么，默认就只做什么**；超出的部分要么另立动词，要么 opt-in flag。

### Raycast 脚本（2026-07-27 已整体归档）

`raycast/` 整子树已进 `~/.Trash/dead-scripts-20260727/`：`commands/` 早就是空的，
9 个 `doc_*.sh` 在它自己的 `_archive/` 里，只剩一个没人 source 的 `lib/run_python.sh`。
现在的入口是直接命令行（`/docx` skill 已退役）。

## Claude CLI 依赖脚本

| 脚本 | 功能 | 模型 |
|------|------|------|
| `document/md_tools.py frontmatter` | 批量生成 MD frontmatter | haiku |
| `document/sub/scan_sensitive_words.py` | 标书敏感词检测 | haiku |

`llm_client.py` 接口：`chat(system, message, model="haiku")` -> `claude -p --model <model>`
