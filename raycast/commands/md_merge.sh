#!/bin/bash
# @raycast.schemaVersion 1
# @raycast.title md-merge
# @raycast.description Merge multiple Markdown files into one
# @raycast.mode fullOutput
# @raycast.icon 📝
# @raycast.packageName Scripts
source "$(dirname "$0")/../lib/run_python.sh" && run_python "document/md_tools.py" merge "$@"
