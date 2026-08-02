"""doctools —— 版本号 SSOT + 可安装 CLI 的入口包。

这个包**只有两件事**：
  ① `__version__` —— 全仓唯一一份版本号。`pyproject.toml` 的 `[project]` 声明
     `dynamic = ["version"]`，由 hatchling 从本文件读走；`docx_cli.py --version`
     也读本文件。两边都是派生，**任何地方不许再写第二份字面量**。
  ② `cli:main` —— `[project.scripts]` 的 console_script 落点，转发给
     `scripts/document/docx_cli.py` 的 `main()`。

它**不是**新的实现层：业务逻辑一行都不在这里。
`python3 <abs>/scripts/document/docx_cli.py <sub> …` 这条老路径与 `doctools <sub> …`
并存，走的是同一个 `main()`，不是两套。
"""

__version__ = "0.1.0"

__all__ = ["__version__"]
