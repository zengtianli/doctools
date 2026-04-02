#!/bin/bash
# @raycast.schemaVersion 1
# @raycast.title docx-text-format
# @raycast.mode fullOutput
# @raycast.icon 📄
# @raycast.packageName Document Processing
# @raycast.description Fix text formatting in Word document (quotes, punctuation, units)
# @raycast.argument1 { "type": "text", "placeholder": "文件路径", "optional": true }
source "$(dirname "$0")/../lib/run_python.sh" && run_python "document/docx_text_formatter.py" "$@"
