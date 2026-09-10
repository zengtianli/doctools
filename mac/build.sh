#!/bin/bash
set -euo pipefail
MODE="${1:---build}"
case "$MODE" in
  --build|--check|--install) [ "$#" -le 1 ] || exit 2 ;;
  --help) echo "Usage: $0 [--build|--check|--install] (default: build only)"; exit 0 ;;
  *) echo "Unknown option: $MODE" >&2; exit 2 ;;
esac
DIR="$(cd "$(dirname "$0")" && pwd -P)"
cd "$DIR"
PYTHON="$HOME/Dev/.venv/bin/python"
source "$HOME/Dev/tools/dev/lib/tools/macapp/xcode_env.sh"
xcode_env_use macosx
"$PYTHON" "$HOME/Dev/tools/dev/lib/tools/macapp/check_codingkeys.py" "$DIR"
BACKEND="$DIR/../scripts/document/doc_gui_backend.py"
[ -f "$BACKEND" ] && [ -f tests/decode_check.swift ]
CHECK_DIR="$(mktemp -d)"
trap 'rm -rf "$CHECK_DIR"' EXIT
cp tests/decode_check.swift "$CHECK_DIR/main.swift"
"$PYTHON" "$BACKEND" gui-ops > "$CHECK_DIR/ops.json"
swiftc -O -o "$CHECK_DIR/decode_check" "$CHECK_DIR/main.swift" Sources/Models.swift
"$CHECK_DIR/decode_check" "$CHECK_DIR/ops.json"
[ "$MODE" != --check ] || exit 0
DISPLAY_NAME="$("$PYTHON" -c 'import sys,yaml; print(yaml.safe_load(open(sys.argv[1]))["display_name"])' "$DIR/catalog.yaml")"
xcodebuild -project DocTools.xcodeproj -scheme DocTools -configuration Release build | tail -3
BUILT="$(xcodebuild -project DocTools.xcodeproj -scheme DocTools -configuration Release -showBuildSettings 2>/dev/null | awk -F' = ' '/ BUILT_PRODUCTS_DIR =/{print $2; exit}')"
APP="$BUILT/DocTools.app"
[ -d "$APP" ]
plutil -replace CFBundleName -string "$DISPLAY_NAME" "$APP/Contents/Info.plist"
plutil -replace CFBundleDisplayName -string "$DISPLAY_NAME" "$APP/Contents/Info.plist"
plutil -replace CFBundleIconFile -string AppIcon "$APP/Contents/Info.plist"
plutil -replace CFBundleVersion -string "$(git -C "$DIR" rev-list --count HEAD)" "$APP/Contents/Info.plist"
cp icon/AppIcon.icns "$APP/Contents/Resources/AppIcon.icns"
codesign --force -s - "$APP"
echo "Built: $APP"
[ "$MODE" = --install ] || exit 0
DEST="/Applications/$DISPLAY_NAME.app"
if [ -e "$DEST" ]; then
  RETIRED="$HOME/.Trash/app-rebuild-$(date +%Y%m%d-%H%M%S)-$$"
  mkdir -p "$RETIRED"
  mv "$DEST" "$RETIRED/"
fi
cp -R "$APP" "$DEST"
echo "Installed: $DEST"
