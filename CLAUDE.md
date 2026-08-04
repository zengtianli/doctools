# doctools

文档处理与数据转换工具集。**入口是命令行**（`/docx` skill 与 `raycast/` 已退役，史见 `docs/history.md`）。

> **Python venv**：共享于 `~/Dev/.venv`（uv workspace member · 见 `~/Dev/CLAUDE.md` § uv workspace）。本 repo 不建独立 `.venv`。改 deps → 改 `pyproject.toml` + `cd ~/Dev && uv sync`。

> **长叙事外迁 `docs/`（2026-08-04）**：本文只留硬约束、闸门命令与指针。合并史/实证/踩坑
> 全文在 `docs/{wheel-dist,smoke-fixture,predicates-ssot,track-compare,history}.md`，
> 改对应部位前先读对应页。已闭环轮次的 handoff 在 `handoffs/_archive/`。

## 闸门总表（改完对号跑，全部 fail-closed）

| 闸门 | 管什么 |
|---|---|
| `python3 tools/cli_surface.py` | 接口指纹（改前改后 diff 必须空；节点 126 · 叶子 101 · 去重动词 93，报数必须说明口径） |
| `python3 tools/cli_forward_probe.py` | 每条子命令真正转发出去的 argv（内嵌预期比对） |
| `python3 tools/check_function_axis.py` | 职能轴对账（`--fn X` 过滤 · `--md` 出表） |
| `python3 tools/check_verbs_reachable.py` | 每个顶层子命令敲下去真进得去 |
| `python3 tools/check_smoke_coverage.py` | 冒烟表 ↔ CLI 子命令树双向对账（`--list` / `--md`） |
| `python3 tools/check_docx_collar.py` | python-docx 存盘必挂 surgical 收口 |
| `python3 tools/check_external_refs.py` | 仓外（~/Work 等）指进来的路径全部存在 |
| `python3 tools/script_graph.py` | 全仓关系图三视图 + 孤儿；动词轴解不出 = exit 2 |
| `python3 tools/check_wheel_selfcontained.py` | 分发门 A/B/C（详 `docs/wheel-dist.md`） |
| `python3 tools/lsof_backup_ab.py` | lsof/备份收敛对拍（10 处同解 + `--` 生效 + 6 处委派） |
| `python3 tools/allowed_deltas_ab.py` | fix_styleset 的 allowed_deltas 等价对拍（shape_contract 唯一覆盖） |
| `python3 tools/family_ab.py <仓根> <out.json>` | 家族折叠行为快照（144 例） |
| `python3 tools/blast_radius.py run\|diff …` | 量/对拍某条命令的 docx 炸开面 |
| `python3 -m pytest -q` | 全部单元/回归/冒烟 |

`cli_surface` / `cli_forward_probe` **只管 argv**：判据层（caption_re/cn_number/lsof）、
CMD_TABLE fast-path 的二级子命令、wheel 包内分支它们都看不见 —— 对应各自的 pytest 门
与分发门不能省。

## 入口与版本

`cd ~/Dev && uv sync --all-packages` 把本 repo 按 editable 装成 `~/Dev/.venv/bin/doctools`。
`doctools <sub>` 与 `python3 <abs>/scripts/document/docx_cli.py <sub>` 进**同一个**
`docx_cli.main()`（`src/doctools/cli.py` 零分发逻辑，结构上不可能漂移）；
~/Work 130 处绝对路径**一处都不用改**。

- **版本号 SSOT = `src/doctools/__init__.py` 的 `__version__`**，`pyproject.toml` 用
  `dynamic = ["version"]` 读走。**禁止在 `pyproject.toml` 里补 `version = "…"`**。
  改完跑 `uv pip install -e tools/doctools --no-deps --reinstall`（`uv sync` 不重建 dist-info）。
- ⚠ **`--version` 故意不是 argparse 参数**（在 `main()` 手动拦第 0 位）：注册成 root action
  会改 surface 指纹。`--help` 靠 epilog 告知，epilog 不进指纹。实证见 `docs/history.md`。

## 目录结构

```
src/doctools/         # 可安装包壳：__version__ SSOT + console_script 入口（无业务逻辑；src-layout 必需）
scripts/
├── document/ (17)    # 入口（docx_cli/typeset/revise/bid_gate/renum/docx_fmt/md_tools/pptx_cli/pdf_cli/chart…）
│   └── sub/ (44)     # docx_cli 子命令实现（_groups 声明表 + _function_axis 职能表 + 业务模块 + _cli_common 样板）
└── data/ (5)         # 数据转换（xlsx_ + convert）
lib/                  # 公共模块：docx_xml / docx_safe_save / docx_parts / docx_surgical / docx_revise /
                      #   llm_client / cn_number / caption_re / styles / chapter_numbering / text_fixes /
                      #   schemas / soffice / clipboard / progress / common.sh
tools/                # 上表全部闸门
docs/                 # 从本文外迁的长叙事
```

（2026-07-30~31 的三轮折叠史见 `docs/history.md`。）

## 独立入口脚本（不走 `docx_cli`，直接敲）

不在 `docx_cli` 命令表里但**是活的**；登记在此 = 有文档点名，`script_graph` 不判孤儿。

| 脚本 | 干什么 | 怎么跑 |
|---|---|---|
| `scripts/document/bid_gate.py` | 标书终稿门检族（检测逻辑 SSOT 在 `bid_residue_lib.py`）。`run`=四门 driver · `scan`/`sweep`/`identity`/`print`=单门 · `deref`=交叉引用去耦合 | `bid_gate.py run <docx> --mode main\|pei [--rules Y] [--apply]`；单门 `scan\|sweep\|identity\|print <docx>`；`deref <docx> --check` |
| `scripts/document/md_to_audiobook.py` | md → 有声书（edge-tts，章节并发） | `uv run …/md_to_audiobook.py <md>`（PEP-723 自带依赖） |
| `scripts/document/docx_revise.py` | **修订注入：意见=ops.yaml 数据，禁在项目里现编注入脚本**（锚点唯一命中 fail-closed；引擎 `lib/docx_revise.py`） | `docx_revise.py <ops.yaml> [--dry-run]`（写法 `config/spec-examples/revise-ops-example.yaml`） |
| `scripts/document/renum.py` | 编号/题注位移与重排族。`chapter`=md 侧章号位移（/renumber skill 指向）· `tabfig`=md 侧表/图题注号对齐（--check 门 exit 2）· `figures`=docx 图号重排+引用同步（= docx_cli `renumber-fig`）。**三个子命令统一 exit 3 = 枚举为空**，见下节 | `renum.py chapter <chapters.yaml> [--apply]`；`tabfig <yaml\|目录> [--apply\|--check]`；`figures <docx> [--cn-section --kind 图\|表] [--dry-run\|--inplace]` |
| `scripts/document/docx_fmt.py` | docx 版式/字体/文本规范化族。`template`=套模板（docx_cli `template`）· `clone`=版式克隆（docx_cli `format`）· `fonts`=去等线（docx-font-guard hook 指向）· `text`=引号/标点/单位规范化（docx_cli `text-fmt`） | `docx_fmt.py template <docx> [-t 模板]`；`clone extract\|apply …`；`fonts <docx...> --check\|--apply`；`text [flags] <docx...>` |

**加新的独立入口脚本 → 必须在这张表里加一行**，否则 `script_graph` 判孤儿，下次清理就清了。

## 子命令：加子命令 = 加三行数据，不是加一个文件

声明全在 **`scripts/document/sub/_groups.py` 的 `GROUPS` 表**。`Opt` 里「怎么声明给
argparse」和「怎么转发给实现脚本」**必须挨着** —— 分居两处时「加了选项忘了转发」是
这类壳最常见的 bug（用户传了参数、脚本收不到、静默按默认值跑）。

**加/改一条子命令 = 三张表各加一行**（缺一道对应闸门就红并点名，双向：表里留了 CLI
已没有的条目同样判红）：

1. `sub/_groups.py` `GROUPS`（声明+转发）→ 跑 surface + forward_probe
2. `sub/_function_axis.py` `_ROWS`（职能标签）→ 跑 check_function_axis
3. `tests/smoke/_verb_specs.py` `_ROWS`（真敲一遍）→ 跑 check_smoke_coverage + pytest tests/smoke

三张表都用 tuple-of-tuples 不用 dict 字面量：dict 重复键不报错、后一条静默覆盖前一条。

**只跑 `cli_surface` 不够**：它证明接口没变，证明不了参数传过去还是原样 ——
`cli_forward_probe` 拦掉 `exec_script` 录 argv，那才是折叠真正动到的东西。

自带业务逻辑的模块（`outline` / `styles` / `health` / `slim` / `fix_styleset` /
`docx_para` / `combine` / `md_merge_track` / `audit_styleset` / `chapters_sync` /
`health_split` / `pipeline`）**不在 `_groups` 表里**，各自 `register()`。

**`docx_tools.py` 已拆薄**：extract / check / track 实现在 `sub/docx_{extract,check,track}.py`，
`docx_tools.py` 收薄为组合入口 + library re-export（外部有人当库 import，**名字面只增不减**）。
这三段走 CMD_TABLE fast-path，**也不进 `_groups` 表**。改 CLI 契约（子命令名/flag/默认值）
前先看外部调用面：cc-home `commands/docx.md`、zdys SKILL.md、`~/Work/CLAUDE.md` L161
（track review 的 include_ins+strict 默认全开是机器强制条款，**禁改**）。

### `track compare`（2026-08-03 落地 · 有回归门）

```bash
python3 scripts/document/docx_cli.py track compare 原.docx 改.docx -o 修订.docx [-a 作者]
```

段落级 `w:del`+`w:ins`，surgical 只重写 document.xml。**退出码不只有 0**：`0` 已产出 ·
`1` 无差异（不产出文件）· `2` 输入不合法/一边 0 段落 · `3` 已产出但有范围外差异
（表格/页眉页脚/脚注尾注，打 stderr 不静默吞）。
⚠ `track` 是 fast-path，**二级子命令不进 surface 指纹** —— 「空桩假装成功」在闸门体系里
是结构性盲区，只能靠真敲。全记录 `docs/track-compare.md`；回归门 `test_track_compare.py`（13 条）。

### 职能轴（每条子命令碰的是「什么」）

表 = `sub/_function_axis.py` `_ROWS`；用户侧只读入口 `docx_cli.py verbs [--fn X] [--json]`。
职能只有 6 个值，**别自创第七个**：`format` · `content` · `review` · `inspect` ·
`convert` · `dispatch`。当前 93 条可跑动词 + 8 条别名（`read`=extract · `diff`=compare ·
`styleset *`=audit-styleset *；别名不占表项，闸门按 `id(parser)` 自动认领）。

### 冒烟轴（每条子命令真敲一遍）

表 = `tests/smoke/_verb_specs.py` `_ROWS`（argv / 预期 rc / 改不改源件 / skip 理由）；
fixture 造件器 `tests/smoke/_fixture.py <目录>` **自带富度门**（10 类对象计数低于下界即
exit 1 —— fixture 变穷会让写盘动词整批变空跑还全绿，四处实证见 `docs/smoke-fixture.md`）。
当前真跑 91 / skip 2 / 共 93。

- **核心断言是 `mutates`，双向**：False = 源件 md5 必须没变（抓偷偷写盘）；True = 必须
  变了（抓静默空跑）。问的是**源件**：写 `-o`/`--report`/`--out-dir` 的动词源件必须一字节不动。
- ⚠ **`expect_rc` 不都是 0，且与 fixture 内容绑定**（`health gate`=2 是真 FAIL 不是回归、
  `para edit|fix-ppr`=4 来自写盘后复跑 gate），换 fixture 必须整列重测，**不能照抄邻居**。
  `health` 的 rc=2 与「参数错误 rc=2」撞码，runner 不能把 2 一律当失败。重测明细见
  `docs/smoke-fixture.md`。
- ⚠ 图题样式 `Image Caption`、表题 `表题`，**两个名字不能合并**；英文图题号**故意乱序**
  （顺序对了 remap 是恒等映射，`mutates` 断言等于没写）。为什么见 `docs/smoke-fixture.md`。
- ⚠ `mutates` 是布尔的，看不见「部分退化」—— `test_caption_re.py` 那道 pytest 门冒烟替不了。

### 装出去的 wheel 必须自包含（分发门 A/B/C）

`scripts/ lib/ config/ schemas/` 构建时 force-include 镜像进 `doctools/_bundled/`
（相对深度与工作树一致，40+ 处 `parents[N]` 算术原样成立）；工作树一个字节不动。
依赖 **`hq-devlib @ git+https://github.com/zengtianli/devtools.git`**（PEP 508 直接引用，
须 `allow-direct-references = true`；devtools 是 PRIVATE，第三方无凭证装不动；本机被
workspace `[tool.uv.sources]` 覆盖，行为零变化）。当前 **49/49 动词可用**，
`DECLARED_WORKING = 49`。

- ⚠ 本门 `git archive HEAD` 取源：打包相关改动没 commit 就跑必红，**是时序不是回归**。
- ⚠ 验异机行为**必须中和 `$HOME`**（`docx_cli.py` 有 `Path.home()/…` 兜底导入，不中和
  0/49 也能测成 49/49）；判据 B 带依赖装，**不是 `--no-deps`**。
- ⚠ `lib` 是**逐文件** force-include：`lib/llm_client.py` 是指向仓外的 symlink，进包 =
  把机器依赖打进去（判据 C 反向断言不许它在包里；该模块 wheel 场景由 hq-devlib 出）。
- ⚠ `parallel_contract` 三来源（总部工作树 insert(0) → `<根>/lib` **append** → hq-devlib
  包），`docx_cli` / `pdf_cli` 缺它都是 **fail-closed exit 2 不是降级**（pdf_cli 2026-08-04 改齐）。

假绿史（判据 A 曾漏跨机构建阻断、判据 C 曾只存在于文档、`--help` 曾弹 Finder）与
git-URL 三戒的全文：**`docs/wheel-dist.md`**。

### 全仓脚本关系图

```bash
python3 tools/script_graph.py --open     # 三视图：图谱（谁调谁）· 清单（每脚本层/引擎/动词）· 动词（93 条落点）
```

数字以实际输出为准（2026-08-04 时点：110 脚本 · 397 引用 · 0 孤儿）。孤儿判据 =
代码没人引用 + 没文档点名 + 非顶层入口，**三条证据全无才算**。动词轴解不出来 =
`exit 2` 不出图。实现坑（`Target.chain` / `g.flat`）见 `docs/history.md`。

## 开发约定

### 引用路径
- `scripts/xxx/` 下的 Python 脚本：`sys.path.insert(0, str(Path(__file__).parent.parent.parent / "lib"))`
- Shell 引用库：`source "$(dirname "$0")/../../lib/common.sh"`
- LLM 调用：`from llm_client import chat`

### 用 python-docx 存盘的脚本必须挂 surgical 收口（有守卫）

```python
import sys as _sys
from pathlib import Path as _Path
_sys.path.append(str(_Path(__file__).resolve().parents[N] / "lib"))   # N = 到仓根的层数
import docx_safe_save  # noqa: E402,F401
```

**append 不是 `insert(0)`** —— `lib/` 和 `sub/` 都有 `styles.py`，插 0 位会顶掉脚本自己那份。
收口把语义未变的部件按原字节还原（炸开面 60→1，由来见 `docs/history.md`）；`Document()`
新建文档自动不介入。守卫 = `check_docx_collar`；量炸开面 = `blast_radius.py run|diff`
（diff 判据是「老 vs 新」相对比较，**不是「新 vs 理想值」**，后者三种假红见该文件）。
逃生：`DOCX_GRAFT_OFF=1` · `DOCX_GRAFT_QUIET=1`。

### `renum.py` 枚举为空 = exit 3，不是 0 也不是 2（有回归门）

空集上连续性判据恒真，曾对「什么都没量到」发合规证。**3 与 2 必须分开**：2 = 真发现了
问题；空集 = 没能做出判定，混进 2 会让 `doc_dispatch.do_renum` 提前 break、
`typeset_pipeline` 步骤③ 回滚已写盘的居中（两处已放行 3）。空集判据问**枚举端**不问
计划端。回归门 `test_renum_empty_set.py`（7 条，含非空对照）；全文 `docs/predicates-ssot.md`。

### 题注/图表编号判据必走 `lib/caption_re.py`（有回归门）

**别再手写 `图\s*\d+[-–—]\d+` 这类正则**（合并前 8 种互不相同的短横子集互为盲区）。

```python
from caption_re import parse, finditer, pattern, RENUM_CN_CAPTION
n = parse(text, RENUM_CN_CAPTION.for_kind("图"))   # n.section / n.seq / n.raw / n.appendix
```

**加调用点 = 加一个具名 spec 声明，不是再写一条正则**；spec 间差异有一半是有意的，
**别顺手统一**（论证在模块 docstring 与 `docs/predicates-ssot.md`）。改判据必须跑
`test_caption_re.py`（55 条）+ 真 fixture 走 `renum figures` / `health diagnose` /
`caption pair` 三条链。

### 全文改写 docx 必走 `lib/docx_xml.py` 的元素级遍历

python-docx 的 `Document.paragraphs` **静默漏掉** `w:ins`/`w:del`/`w:hyperlink`/文本框/
嵌套表里的 run；批注/脚注/尾注整个 part 碰不到。凡「把全文某类字符改一遍」的引擎都用：

```python
from docx_xml import iter_text_roots, iter_paragraphs, para_own_runs, run_text, set_run_text
for label, root, flush in iter_text_roots(doc):      # 正文/批注/脚注/尾注/页眉页脚
    for p in iter_paragraphs(root):
        for r in para_own_runs(p):
            set_run_text(r, fix(run_text(r)))
    if flush: flush()                                 # 脚注/尾注是通用 Part,必须回写 blob
```

删除态文本载体是 `w:delText`，写成 `w:t` = 把删掉的字变回正文（`set_run_text` 已挡）。
回归门 `test_docx_text_formatter_scopes.py`。

### 中文数字转 int 只有 `lib/cn_number.py` 一份（有测试门）

`chinese_to_arabic`（严格，抛 ValueError）与 `cn_to_int`（宽松，返 None）**必须并存别合一**
—— 两类调用点分别靠异常/None 控流，合了必有一类静默坏。「万」不支持是有意的；正则是
另一根轴不归它管。回归门 `test_cn_number.py`（含「sub/ 下不许再有本地副本」）；
合并史 `docs/predicates-ssot.md`。

### lsof 占用检查 + 备份路径只有 `_cli_common` 一份（有对拍门）

```python
import _cli_common as _cc          # sub/ 自身 append 进 sys.path，别 insert(0)
occ = _cc.lsof_check(path)         # str|None；`--` 不是洁癖，删了会对 Word 开着的文件放行写盘
bak = _cc.find_next_backup(path)   # 只算路径；_cc.make_backup 才 copy2
```

⚠ `pipeline_lib.lsof_check` 等是**委派壳**（外部按名 import，名字面只增不减）。
`.bak-<时间戳>` 是另一套命名约定，**不归本机制管**。对拍门 = `tools/lsof_backup_ab.py`；
全文 `docs/predicates-ssot.md`。

### per-op 选项走后端声明，别写进 Swift

GUI 勾选框由 `doc_gui_backend.OPS[<op>]["options"]` 声明，Swift 只按 `type` 泛化渲染，
**Sources/ 里不许出现任何 option id 字面量**。契约铁律：未知 key / 非法值走信封
`{"ok": false, "error": …}` 且 **exit 0**（argparse 的 exit 2 + 空 stdout = Swift 只看到
「后端崩了」）。引擎配置走 `FormatConfig` 冻结 dataclass 显式穿参，**禁模块级可变全局**。
GUI 未勾的选项 = 「没传 = 保持默认」，只有显式 `=0` 才关。
⚠ 域过滤与引号规则不正交：被跳过的 run 仍要推进引号奇偶 counter 但不写回。回归门
`test_docx_text_formatter_safety.py::test_skipped_scope_does_not_flip_quote_direction`。

### 破坏性动作必须自己占一个动词

**不可逆操作不许搭在非破坏性动词的默认行为里**，也不许只靠 flag 才关得掉。判据：
**动词名字承诺什么，默认就只做什么**；超出的部分另立动词（如 `stripchrome` 标 `danger`）
或 opt-in flag。未知 flag 一律 `sys.exit(2)` 禁 fallthrough（`--help` 曾弹 Finder 写盘，
两案实录见 `docs/wheel-dist.md` 与 `docs/history.md`）；文本替换必须带词边界
（裸 `str.replace` 曾产出 小时候→h候）。

## Claude CLI 依赖脚本

| 脚本 | 功能 | 模型 |
|------|------|------|
| `document/md_tools.py frontmatter` | 批量生成 MD frontmatter | haiku |
| `document/sub/scan_sensitive_words.py` | 标书敏感词检测 | haiku |

`llm_client.py` 接口：`chat(system, message, model="haiku")` -> `claude -p --model <model>`
