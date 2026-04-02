#!/usr/bin/env python3
"""
环境配置模块
动态检测项目路径，不再硬编码
"""

import os
import shutil
from pathlib import Path

# 项目根目录：lib/ 的上两级
PROJECT_ROOT = Path(__file__).resolve().parent.parent

# .assets 根目录
ASSETS_ROOT = Path(__file__).resolve().parent.parent

# 项目目录（复杂多文件项目）
PROJECTS_DIR = ASSETS_ROOT / "projects"

# Python 路径：优先使用当前解释器
PYTHON_PATH = shutil.which("python3") or "/usr/bin/env python3"

# 日志文件
USAGE_LOG = os.path.expanduser("~/.useful_scripts_usage.log")

# ── Raycast 环境变量加载 ──────────────────────────────
ENV_ZSH = Path.home() / "Documents/sync/zsh/config/env.zsh"


def load_env():
    """
    加载 env.zsh 中的环境变量。
    Raycast 不继承 shell 环境，需要主动加载。
    终端下环境变量已存在，setdefault 不会覆盖。
    """
    if not ENV_ZSH.exists():
        return
    import re
    for line in ENV_ZSH.read_text().splitlines():
        m = re.match(r'export (\w+)="([^"]*)"', line)
        if m:
            os.environ.setdefault(m.group(1), m.group(2))
