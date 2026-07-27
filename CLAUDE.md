# doctools

文档处理与数据转换工具集，从 scripts 仓库拆分。主要通过 Raycast 调用。

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

raycast/
├── commands/         # 27 个 Raycast wrapper
└── lib/              # run_python.sh 运行器
```

## 开发约定

### 引用路径
- `scripts/xxx/` 下的 Python 脚本：`sys.path.insert(0, str(Path(__file__).parent.parent.parent / "lib"))`
- Shell 引用库：`source "$(dirname "$0")/../../lib/common.sh"`
- LLM 调用：`from llm_client import chat`

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

### Raycast 脚本
- `raycast/commands/` 下是 Shell wrapper（含 @raycast 元数据）
- Wrapper 通过 `run_python.sh` 调用实际脚本：`run_python "document/docx_text_formatter.py"`

## Claude CLI 依赖脚本

| 脚本 | 功能 | 模型 |
|------|------|------|
| `document/bullet_to_paragraph.py` | 要点转公文段落/表格 | haiku |
| `document/md_tools.py frontmatter` | 批量生成 MD frontmatter | haiku |
| `document/scan_sensitive_words.py` | 标书敏感词检测 | haiku |

`llm_client.py` 接口：`chat(system, message, model="haiku")` -> `claude -p --model <model>`
