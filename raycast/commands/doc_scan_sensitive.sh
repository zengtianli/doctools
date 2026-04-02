#!/bin/bash
# @raycast.schemaVersion 1
# @raycast.title doc-scan-sensitive
# @raycast.mode fullOutput
# @raycast.icon 🔍
# @raycast.packageName Document Processing
# @raycast.description Scan bid documents for sensitive words using Claude AI
# @raycast.argument1 { "type": "text", "placeholder": "Directory path", "optional": false }
source "$(dirname "$0")/../lib/run_python.sh" && run_python "document/scan_sensitive_words.py" "$1"
