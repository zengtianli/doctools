#!/bin/bash
# @raycast.schemaVersion 1
# @raycast.title xlsx-from-csv
# @raycast.description Convert CSV file to Excel spreadsheet
# @raycast.mode fullOutput
# @raycast.icon 📊
# @raycast.packageName Scripts
source ~/Dev/devtools/lib/log_usage.sh
source "$(dirname "$(realpath "$0")")/../lib/run_python.sh" && run_python "data/convert.py" xlsx-from-csv "$@"
