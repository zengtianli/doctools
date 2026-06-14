#!/bin/bash
# @raycast.schemaVersion 1
# @raycast.title Split / Merge
# @raycast.mode fullOutput
# @raycast.icon 🔀
# @raycast.packageName Document Processing
# @raycast.description Split one document or merge several — pick action, format auto-detected
# @raycast.argument1 { "type": "dropdown", "title": "Action", "data": [{"title":"Split","value":"split"},{"title":"Merge","value":"merge"}] }
# 实现：doc_dispatch split @ ~/Dev/tools/doctools/scripts/document/doc_dispatch.py
# 实现：doc_dispatch merge @ ~/Dev/tools/doctools/scripts/document/doc_dispatch.py
source ~/Dev/tools/dev/lib/log_usage.sh 2>/dev/null
source "$(dirname "$(realpath "$0")")/../lib/run_python.sh"

# $1 = dropdown value (split / merge) → 对应 doc_dispatch verb
case "$1" in
  split) verb="split" ;;
  merge) verb="merge" ;;
  *) echo "❌ Pick an action: Split or Merge"; exit 1 ;;
esac
shift  # 丢掉 dropdown 选择,剩余位置参数才是文件

# Finder 选择逻辑照搬两个源 wrapper(get_finder_selection_multiple + 去尾斜杠)
files=()
if [ $# -gt 0 ]; then files=("$@"); else
  while IFS= read -r f; do [ -n "$f" ] && files+=("${f%/}"); done < <(get_finder_selection_multiple)
fi
[ ${#files[@]} -eq 0 ] && { echo "❌ Select file(s) in Finder first"; exit 1; }
run_python "document/doc_dispatch.py" "$verb" "${files[@]}"
