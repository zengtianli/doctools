#!/bin/bash
# @raycast.schemaVersion 1
# @raycast.title Scan Sensitive Words
# @raycast.mode fullOutput
# @raycast.icon 🔍
# @raycast.packageName Document Processing
# @raycast.description Scan md/docx in a selected folder for competitor names / risky wording (bid compliance)
# 实现：doc_dispatch scan @ ~/Dev/tools/doctools/scripts/document/doc_dispatch.py
source ~/Dev/tools/dev/lib/log_usage.sh 2>/dev/null
source "$(dirname "$(realpath "$0")")/../lib/run_python.sh"
DIR="${1:-$(get_finder_selection_single)}"
DIR="${DIR%/}"
[ -z "$DIR" ] && { echo "❌ Select a folder in Finder first"; exit 1; }
run_python "document/doc_dispatch.py" scan "$DIR"
