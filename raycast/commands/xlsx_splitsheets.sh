#!/bin/bash
# @raycast.schemaVersion 1
# @raycast.title xlsx-split
# @raycast.description Split Excel workbook into separate sheets
# @raycast.mode fullOutput
# @raycast.icon 📊
# @raycast.packageName Scripts
source "$(dirname "$0")/../lib/run_python.sh" && run_python "data/xlsx_splitsheets.py" "$@"
