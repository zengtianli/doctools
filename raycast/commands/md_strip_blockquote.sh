#!/bin/bash
# @raycast.schemaVersion 1
# @raycast.title md-strip-blockquote
# @raycast.description 删除 Markdown 中所有 blockquote 行（Finder 选中文件或输入路径）
# @raycast.mode fullOutput
# @raycast.icon 📝
# @raycast.packageName Scripts
# @raycast.argument1 { "type": "text", "placeholder": "MD 文件/目录路径", "optional": true }

input="$1"

# 没有参数时，尝试获取 Finder 选中的文件
if [ -z "$input" ]; then
    input=$(osascript -e 'tell application "Finder" to get POSIX path of (selection as alias)' 2>/dev/null)
fi

if [ -z "$input" ]; then
    echo "❌ 请选中文件或输入路径"
    exit 1
fi

source "$(dirname "$0")/../lib/run_python.sh" && run_python "document/md_tools.py" strip "$input" --fix
