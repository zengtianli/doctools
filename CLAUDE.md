# doctools

文档处理与数据转换工具集，从 scripts 仓库拆分。**入口是 `/docx` skill 与命令行**（`raycast/` 整子树 2026-07-27 已归档：`commands/` 早已是空的，9 个 `doc_*.sh` 在自己的 `_archive/` 里）。

> **Python venv**：共享于 `~/Dev/.venv`（uv workspace member · 见 `~/Dev/CLAUDE.md` § uv workspace）。本 repo 不建独立 `.venv`。改 deps → 改 `pyproject.toml` + `cd ~/Dev && uv sync`。

## 目录结构

```
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
| `scripts/document/renum.py` | 编号/题注位移与重排族（2026-07-31 原 chapter_renumber/tabfig_align/docx_renumber_figures 3 件合并）。子命令：`chapter`=md 侧章号位移引擎（config 驱动，/renumber skill 指向）· `tabfig`=md 侧 表/图 题注号对齐（--check 机检门 exit 2）· `figures`=docx 图号重排+引用同步（docx_cli `renumber-fig` 即它） | `python3 scripts/document/renum.py chapter <chapters.yaml> [--apply]`；`renum.py tabfig <yaml\|目录> [--apply\|--check]`；`renum.py figures <docx> [--cn-section --kind 图\|表] [--dry-run\|--inplace]` |
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

## 全仓脚本关系图

```bash
python3 tools/script_graph.py --open     # 92 个脚本 · 谁调谁 · 双链可点
```

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
