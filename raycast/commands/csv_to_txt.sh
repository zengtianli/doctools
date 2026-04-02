#!/bin/bash
# @raycast.schemaVersion 1
# @raycast.title csv-to-txt
# @raycast.description Convert CSV file to text format
# @raycast.mode fullOutput
# @raycast.icon 📊
# @raycast.packageName Scripts
source "$(dirname "$0")/../lib/run_python.sh" && run_python "data/convert.py" csv-to-txt "$@"
