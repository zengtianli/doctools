#!/bin/bash
# @raycast.schemaVersion 1
# @raycast.title 转换文档格式
# @raycast.mode fullOutput
# @raycast.icon 🔄
# @raycast.packageName Document Processing
# @raycast.description 选中文件 → 转成目标格式;源格式自动认(docx/pptx/csv/xlsx/txt/md 互转)
# @raycast.argument1 { "type": "dropdown", "placeholder": "目标格式", "data": [{"title":"→ Markdown","value":"md"},{"title":"→ Word(套模板)","value":"word"},{"title":"→ Excel","value":"xlsx"},{"title":"→ CSV","value":"csv"},{"title":"→ txt","value":"txt"}] }
# 实现：doc_dispatch convert @ ~/Dev/tools/doctools/scripts/document/doc_dispatch.py
source ~/Dev/tools/dev/lib/log_usage.sh 2>/dev/null
source "$(dirname "$(realpath "$0")")/../lib/run_python.sh"
TARGET="$1"
files=()
while IFS= read -r f; do [ -n "$f" ] && files+=("${f%/}"); done < <(get_finder_selection_multiple)
[ ${#files[@]} -eq 0 ] && { echo "❌ 请在 Finder 选中文件再触发"; exit 1; }
run_python "document/doc_dispatch.py" convert --to "$TARGET" "${files[@]}"
