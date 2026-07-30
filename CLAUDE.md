# doctools

文档处理与数据转换工具集，从 scripts 仓库拆分。**入口是 `/docx` skill 与命令行**（`raycast/` 整子树 2026-07-27 已归档：`commands/` 早已是空的，9 个 `doc_*.sh` 在自己的 `_archive/` 里）。

> **Python venv**：共享于 `~/Dev/.venv`（uv workspace member · 见 `~/Dev/CLAUDE.md` § uv workspace）。本 repo 不建独立 `.venv`。改 deps → 改 `pyproject.toml` + `cd ~/Dev && uv sync`。

## 目录结构

```
scripts/
├── document/ (13)    # 文档处理（docx_ + md_ + pptx_ + chart）
└── data/ (4)         # 数据转换（xlsx_ + convert）

lib/                  # 公共模块
├── display.py        # 终端输出（颜色、进度）
├── file_ops.py       # 文件操作工具函数
├── finder.py         # Finder 选择/输入获取
├── progress.py       # 进度条
├── docx_xml.py       # DOCX XML 操作
├── clipboard.py      # 剪贴板操作
├── env.py            # 环境变量
├── usage_log.py      # 使用日志
├── llm_client.py     # AI 调用（claude -p 封装）
└── common.sh         # Shell 公共函数

└── lib/              # run_python.sh 运行器
```

## 独立入口脚本（不走 `docx_cli`，直接敲）

这几个不在 `docx_cli` 的命令表里、也没有 skill 指向，但**是活的**。登记在这里有两个作用：
让人找得到；让 `tools/script_graph.py` 把它们算作「有文档点名」而不是孤儿。

| 脚本 | 干什么 | 怎么跑 |
|---|---|---|
| `scripts/document/bid_deref.py` | 标书正文交叉引用去耦合（合稿人会删/调章节，写死编号=断链） | `python3 scripts/document/bid_deref.py <docx>` |
| `scripts/document/gen_report.py` | 按参考报告逐章生成新县市报告（走 llm_client） | `python3 scripts/document/gen_report.py --help` |
| `scripts/document/md_to_audiobook.py` | md → 有声书（edge-tts，章节并发） | `uv run scripts/document/md_to_audiobook.py <md>`（PEP-723 自带依赖） |

**加新的独立入口脚本 → 必须在这张表里加一行**，否则它对全仓不可见：`script_graph` 会把它
判成孤儿，下次清理就把它清了。

## 全仓脚本关系图

```bash
python3 tools/script_graph.py --open     # 151 个脚本 · 谁调谁 · 双链可点
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
现在的入口是 `/docx` skill 与直接命令行。

## Claude CLI 依赖脚本

| 脚本 | 功能 | 模型 |
|------|------|------|
| `document/bullet_to_paragraph.py` | 要点转公文段落/表格 | haiku |
| `document/md_tools.py frontmatter` | 批量生成 MD frontmatter | haiku |
| `document/scan_sensitive_words.py` | 标书敏感词检测 | haiku |

`llm_client.py` 接口：`chat(system, message, model="haiku")` -> `claude -p --model <model>`
