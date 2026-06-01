#!/bin/bash
# @raycast.schemaVersion 1
# @raycast.title Merge Documents
# @raycast.mode fullOutput
# @raycast.icon 🔗
# @raycast.packageName Document Processing
# @raycast.description Merge multiple same-type files (md/txt) into one — format auto-detected
# 实现：doc_dispatch merge @ ~/Dev/tools/doctools/scripts/document/doc_dispatch.py
source ~/Dev/tools/dev/lib/log_usage.sh 2>/dev/null
source "$(dirname "$(realpath "$0")")/../lib/run_python.sh"
files=()
if [ $# -gt 0 ]; then files=("$@"); else
  while IFS= read -r f; do [ -n "$f" ] && files+=("${f%/}"); done < <(get_finder_selection_multiple)
fi
[ ${#files[@]} -eq 0 ] && { echo "❌ Select file(s) in Finder first"; exit 1; }
run_python "document/doc_dispatch.py" merge "${files[@]}"
