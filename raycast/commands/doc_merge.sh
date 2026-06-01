#!/bin/bash
# @raycast.schemaVersion 1
# @raycast.title 合并文档
# @raycast.mode fullOutput
# @raycast.icon 🔗
# @raycast.packageName Document Processing
# @raycast.description 选中多个同类(md/txt) → 合并成一个;格式自动认
# 实现：doc_dispatch merge @ ~/Dev/tools/doctools/scripts/document/doc_dispatch.py
source ~/Dev/tools/dev/lib/log_usage.sh 2>/dev/null
source "$(dirname "$(realpath "$0")")/../lib/run_python.sh"
files=()
if [ $# -gt 0 ]; then files=("$@"); else
  while IFS= read -r f; do [ -n "$f" ] && files+=("${f%/}"); done < <(get_finder_selection_multiple)
fi
[ ${#files[@]} -eq 0 ] && { echo "❌ 请在 Finder 选中文件再触发"; exit 1; }
run_python "document/doc_dispatch.py" merge "${files[@]}"
