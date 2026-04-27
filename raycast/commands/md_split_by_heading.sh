#!/bin/bash
# @raycast.schemaVersion 1
# @raycast.title md-split
# @raycast.description Split Markdown file by headings
# @raycast.mode fullOutput
# @raycast.icon ✂️
# @raycast.packageName Scripts
source ~/Dev/devtools/lib/log_usage.sh
source "$(dirname "$(realpath "$0")")/../lib/run_python.sh" && run_python "document/md_tools.py" split "$@"
