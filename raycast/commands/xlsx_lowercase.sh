#!/bin/bash
# @raycast.schemaVersion 1
# @raycast.title xlsx-lowercase
# @raycast.description Convert all text in Excel to lowercase
# @raycast.mode fullOutput
# @raycast.icon 📊
# @raycast.packageName Scripts
source ~/Dev/devtools/lib/log_usage.sh
source "$(dirname "$(realpath "$0")")/../lib/run_python.sh" && run_python "data/xlsx_lowercase.py" "$@"
