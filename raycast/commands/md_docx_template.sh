#!/bin/bash
# @raycast.schemaVersion 1
# @raycast.title md-to-docx-template
# @raycast.description Convert Markdown to Word document with template
# @raycast.mode fullOutput
# @raycast.icon 📄
# @raycast.packageName DOCX
source "$(dirname "$0")/../lib/run_python.sh" && run_python "document/md_docx_template.py" "$@"
