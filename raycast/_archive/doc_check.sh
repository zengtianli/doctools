#!/bin/bash
# @raycast.schemaVersion 1
# @raycast.title Check Document
# @raycast.mode fullOutput
# @raycast.icon 🔎
# @raycast.packageName Document Processing
# @raycast.description Pre-delivery check — Normalize selected files (quotes/punctuation/units/fonts) or Scan a folder for sensitive/risky wording
# @raycast.argument1 { "type": "dropdown", "placeholder": "Action", "data": [{"title":"Normalize","value":"clean"},{"title":"Scan","value":"scan"}] }
# 实现：doc_dispatch clean @ ~/Dev/tools/doctools/scripts/document/doc_dispatch.py
# 实现：doc_dispatch scan  @ ~/Dev/tools/doctools/scripts/document/doc_dispatch.py
source ~/Dev/tools/dev/lib/log_usage.sh 2>/dev/null
source "$(dirname "$(realpath "$0")")/../lib/run_python.sh"

VERB="$1"; shift

case "$VERB" in
  clean)
    # 取 Finder 多选文件(同 doc_normalize.sh);剩余位置参数为显式文件覆盖
    files=()
    if [ $# -gt 0 ]; then files=("$@"); else
      while IFS= read -r f; do [ -n "$f" ] && files+=("${f%/}"); done < <(get_finder_selection_multiple)
    fi
    [ ${#files[@]} -eq 0 ] && { echo "❌ Select file(s) in Finder first"; exit 1; }
    run_python "document/doc_dispatch.py" clean "${files[@]}"
    ;;
  scan)
    # 取 Finder 单选目录(同 doc_scan.sh)
    DIR="${1:-$(get_finder_selection_single)}"
    DIR="${DIR%/}"
    [ -z "$DIR" ] && { echo "❌ Select a folder in Finder first"; exit 1; }
    run_python "document/doc_dispatch.py" scan "$DIR"
    ;;
  *)
    echo "❌ Unknown action: $VERB (expected clean or scan)"; exit 1
    ;;
esac
