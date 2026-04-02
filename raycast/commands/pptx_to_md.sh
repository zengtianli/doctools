#!/bin/bash
# @raycast.schemaVersion 1
# @raycast.title pptx-to-md
# @raycast.description Convert PowerPoint presentation to Markdown
# @raycast.mode fullOutput
# @raycast.icon 📽️
# @raycast.packageName Scripts
source "$(dirname "$0")/../lib/run_python.sh" && run_python "document/pptx_to_md.py" "$@"
