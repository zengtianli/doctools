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
