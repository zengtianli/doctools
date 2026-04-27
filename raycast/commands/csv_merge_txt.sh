#!/bin/bash
# @raycast.schemaVersion 1
# @raycast.title csv-merge-txt
# @raycast.description Merge multiple text files into a single CSV
# @raycast.mode fullOutput
# @raycast.icon 📊
# @raycast.packageName Scripts
source ~/Dev/devtools/lib/log_usage.sh
source "$(dirname "$(realpath "$0")")/../lib/run_python.sh" && run_python "data/convert.py" csv-merge-txt "$@"
