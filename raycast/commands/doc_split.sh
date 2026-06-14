#!/bin/bash
# @raycast.schemaVersion 1
# @raycast.title Split Document
# @raycast.mode fullOutput
# @raycast.icon ✂️
# @raycast.packageName Document Processing
# @raycast.description Split md (by heading) or xlsx (by sheet) — format auto-detected
# 实现：doc_dispatch split @ ~/Dev/tools/doctools/scripts/document/doc_dispatch.py
source ~/Dev/tools/dev/lib/log_usage.sh 2>/dev/null
source "$(dirname "$(realpath "$0")")/../lib/run_python.sh"
files=()
if [ $# -gt 0 ]; then files=("$@"); else
  while IFS= read -r f; do [ -n "$f" ] && files+=("${f%/}"); done < <(get_finder_selection_multiple)
fi
[ ${#files[@]} -eq 0 ] && { echo "❌ Select file(s) in Finder first"; exit 1; }
run_python "document/doc_dispatch.py" split "${files[@]}"
