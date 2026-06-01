#!/bin/bash
# @raycast.schemaVersion 1
# @raycast.title pptx-table
# @raycast.description Apply table styling to PowerPoint presentation
# @raycast.mode fullOutput
# @raycast.icon 📽️
# @raycast.packageName Document Processing
source ~/Dev/tools/dev/lib/log_usage.sh
source "$(dirname "$(realpath "$0")")/../lib/run_python.sh" && run_python "document/pptx_tools.py" table "$@"
