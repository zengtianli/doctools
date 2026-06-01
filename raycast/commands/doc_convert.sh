#!/bin/bash
# @raycast.schemaVersion 1
# @raycast.title Convert Document
# @raycast.mode fullOutput
# @raycast.icon 🔄
# @raycast.packageName Document Processing
# @raycast.description Convert selected files to a target format; source format auto-detected
# @raycast.argument1 { "type": "dropdown", "placeholder": "Target format", "data": [{"title":"To Markdown","value":"md"},{"title":"To Word","value":"word"},{"title":"To Excel","value":"xlsx"},{"title":"To CSV","value":"csv"},{"title":"To txt","value":"txt"}] }
# 实现：doc_dispatch convert @ ~/Dev/tools/doctools/scripts/document/doc_dispatch.py
source ~/Dev/tools/dev/lib/log_usage.sh 2>/dev/null
source "$(dirname "$(realpath "$0")")/../lib/run_python.sh"
TARGET="$1"
files=()
while IFS= read -r f; do [ -n "$f" ] && files+=("${f%/}"); done < <(get_finder_selection_multiple)
[ ${#files[@]} -eq 0 ] && { echo "❌ Select file(s) in Finder first"; exit 1; }
run_python "document/doc_dispatch.py" convert --to "$TARGET" "${files[@]}"
