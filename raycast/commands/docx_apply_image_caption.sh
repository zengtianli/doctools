#!/bin/bash
# @raycast.schemaVersion 1
# @raycast.title docx-image-caption
# @raycast.mode fullOutput
# @raycast.icon 📄
# @raycast.packageName Document Processing
# @raycast.description Apply ZDWP image caption style to Word document
source "$(dirname "$0")/../lib/run_python.sh" && run_python "document/docx_apply_image_caption.py" "$@"
