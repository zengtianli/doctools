#!/bin/bash
# @raycast.schemaVersion 1
# @raycast.title 敏感词扫描
# @raycast.mode fullOutput
# @raycast.icon 🔍
# @raycast.packageName Document Processing
# @raycast.description 选中一个目录 → 扫描里面 md/docx 的竞品名/过硬措辞(投标合规自检)
# 实现：doc_dispatch scan @ ~/Dev/tools/doctools/scripts/document/doc_dispatch.py
source ~/Dev/tools/dev/lib/log_usage.sh 2>/dev/null
source "$(dirname "$(realpath "$0")")/../lib/run_python.sh"
DIR="${1:-$(get_finder_selection_single)}"
DIR="${DIR%/}"
[ -z "$DIR" ] && { echo "❌ 请在 Finder 选中一个目录再触发"; exit 1; }
run_python "document/doc_dispatch.py" scan "$DIR"
