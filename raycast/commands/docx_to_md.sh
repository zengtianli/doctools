#!/bin/bash
# @raycast.schemaVersion 1
# @raycast.title docx-to-md
# @raycast.mode fullOutput
# @raycast.icon 📄
# @raycast.packageName DOCX
# @raycast.description DOCX转Markdown（使用markitdown）- 支持多选
source "$(dirname "$0")/../lib/run_python.sh" && run_shell "document/docx_to_md.sh" "$@"
