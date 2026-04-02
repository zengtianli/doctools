#!/bin/bash
# @raycast.schemaVersion 1
# @raycast.title xlsx-to-csv
# @raycast.description Convert Excel spreadsheet to CSV format
# @raycast.mode fullOutput
# @raycast.icon 📊
# @raycast.packageName Scripts
source "$(dirname "$0")/../lib/run_python.sh" && run_python "data/convert.py" xlsx-to-csv "$@"
