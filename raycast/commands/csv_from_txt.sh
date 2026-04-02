#!/bin/bash
# @raycast.schemaVersion 1
# @raycast.title csv-from-txt
# @raycast.description Convert text file to CSV format
# @raycast.mode fullOutput
# @raycast.icon 📊
# @raycast.packageName Scripts
source "$(dirname "$0")/../lib/run_python.sh" && run_python "data/convert.py" csv-from-txt "$@"
