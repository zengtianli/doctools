#!/bin/bash
# @raycast.schemaVersion 1
# @raycast.title Typeset to Word
# @raycast.mode fullOutput
# @raycast.icon 📝
# @raycast.packageName Document Processing
# @raycast.description md/docx -> templated Word (apply styles, fix text, center figure captions)
# 实现：doc_dispatch typeset @ ~/Dev/tools/doctools/scripts/document/doc_dispatch.py
source ~/Dev/tools/dev/lib/log_usage.sh 2>/dev/null
source "$(dirname "$(realpath "$0")")/../lib/run_python.sh"
files=()
if [ $# -gt 0 ]; then files=("$@"); else
  while IFS= read -r f; do [ -n "$f" ] && files+=("${f%/}"); done < <(get_finder_selection_multiple)
fi
[ ${#files[@]} -eq 0 ] && { echo "❌ Select file(s) in Finder first"; exit 1; }
run_python "document/doc_dispatch.py" typeset "${files[@]}"
