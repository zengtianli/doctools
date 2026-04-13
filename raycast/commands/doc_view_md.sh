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

source "$(dirname "$(realpath "$0")")/../lib/run_python.sh" && run_python "document/md_tools.py" to-html "$1"
