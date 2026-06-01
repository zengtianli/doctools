#!/bin/bash
# @raycast.schemaVersion 1
# @raycast.title csv-to-txt
# @raycast.description Convert CSV file to text format
# @raycast.mode fullOutput
# @raycast.icon 📊
# @raycast.packageName Data Processing
source ~/Dev/tools/dev/lib/log_usage.sh
source "$(dirname "$(realpath "$0")")/../lib/run_python.sh" && run_python "data/convert.py" csv-to-txt "$@"
