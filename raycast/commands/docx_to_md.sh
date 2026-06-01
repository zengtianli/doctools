#!/bin/bash
# @raycast.schemaVersion 1
# @raycast.title docx-to-md
# @raycast.mode fullOutput
# @raycast.icon 📄
# @raycast.packageName Document Processing
# @raycast.description DOCX转Markdown（使用markitdown）- 支持多选
source ~/Dev/tools/dev/lib/log_usage.sh
source "$(dirname "$(realpath "$0")")/../lib/run_python.sh" && run_shell "document/docx_to_md.sh" "$@"
