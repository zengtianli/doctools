#!/bin/bash
# @raycast.schemaVersion 1
# @raycast.title docx-apply-template
# @raycast.description Apply template styles to a Word document
# @raycast.mode fullOutput
# @raycast.icon 📄
# @raycast.packageName DOCX
source "$(dirname "$0")/../lib/run_python.sh" && run_python "document/docx_apply_template.py" "$@"
