#!/bin/bash

# Required parameters:
# @raycast.schemaVersion 1
# @raycast.title View MD as HTML
# @raycast.mode fullOutput
# @raycast.packageName Document

# Optional parameters:
# @raycast.icon 📄
# @raycast.description 把 MD 文件或目录渲染成网页在浏览器中打开
# @raycast.argument1 { "type": "text", "placeholder": "文件或目录路径" }

SCRIPT_DIR="$(cd "$(dirname "$0")" && pwd)"
REPO_ROOT="$(cd "$SCRIPT_DIR/../.." && pwd)"
PYTHON=$(which python3)

"$PYTHON" "$REPO_ROOT/scripts/document/md_tools.py" to-html "$1"
