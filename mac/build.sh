#!/bin/bash
set -euo pipefail
MODE="${1:---build}"
case "$MODE" in
  --build|--check|--install) [ "$#" -le 1 ] || exit 2 ;;
  --help) echo "Usage: $0 [--build|--check|--install] (default: build only; --install also links ~/.local/bin/dockit)"; exit 0 ;;
  *) echo "Unknown option: $MODE" >&2; exit 2 ;;
esac
DIR="$(cd "$(dirname "$0")" && pwd -P)"
cd "$DIR"
PYTHON="$HOME/Dev/.venv/bin/python"
source "$HOME/Dev/tools/dev/lib/tools/macapp/xcode_env.sh"
xcode_env_use macosx
LIFECYCLE_VENDOR="${APP_LIFECYCLE_VENDOR:-$HOME/Dev/tools/dev/lib/tools/macapp/swift-shared/vendor-lifecycle.py}"
"$PYTHON" "$LIFECYCLE_VENDOR" --platform mac --shortcuts --target-source-dir "$DIR/Sources" --check
"$PYTHON" "$HOME/Dev/tools/dev/lib/tools/macapp/check_codingkeys.py" "$DIR"
BACKEND="$DIR/../scripts/document/doc_gui_backend.py"
[ -f "$BACKEND" ] && [ -f tests/decode_check.swift ]
CHECK_DIR="$(mktemp -d)"
trap 'rm -rf "$CHECK_DIR"' EXIT
cp tests/decode_check.swift "$CHECK_DIR/main.swift"
"$PYTHON" "$BACKEND" gui-ops > "$CHECK_DIR/ops.json"
# agent CLI:包装脚本语法 + 真实后端 --help(与装进 App 的是同一个文件)
sh -n bin/dockit
bin/dockit --help > /dev/null
# 三组使用独立产物与临时偏好域；12 核 Mini 同时编译，任一组失败都等其余收尾后拒绝构建。
CHECK_PIDS=()
(
  swiftc -O -num-threads 4 -o "$CHECK_DIR/decode_check" "$CHECK_DIR/main.swift" Sources/Models.swift
  "$CHECK_DIR/decode_check" "$CHECK_DIR/ops.json"
  swiftc -O -num-threads 4 -parse-as-library Sources/Models.swift Sources/BackendClient.swift Sources/ViewModel.swift \
    tests/PreferencesCheck.swift -o "$CHECK_DIR/preferences_check"
  "$CHECK_DIR/preferences_check"
) &
CHECK_PIDS+=("$!")
(
  # 命令与界面共用设置，隔离偏好文件，不碰本人偏好。
  swiftc -O -num-threads 4 -parse-as-library Sources/Models.swift Sources/BackendClient.swift Sources/ViewModel.swift \
    tests/SettingsFollowCheck.swift -o "$CHECK_DIR/settings_follow_check"
  "$CHECK_DIR/settings_follow_check" "$DIR/bin/dockit" "$CHECK_DIR/ops.json" "$CHECK_DIR/follow"
) &
CHECK_PIDS+=("$!")
(
  # 共用生命周期命令整条链，含离屏运行中的 App 跟随与背靠背命令。
  swiftc -O -num-threads 4 -suppress-warnings -parse-as-library Sources/*.swift -o "$CHECK_DIR/DocTools"
  DOCKIT_NATIVE="$CHECK_DIR/DocTools" "$PYTHON" tests/test_lifecycle_cli.py
) &
CHECK_PIDS+=("$!")
CHECK_STATUS=0
for CHECK_PID in "${CHECK_PIDS[@]}"; do
  if ! wait "$CHECK_PID"; then CHECK_STATUS=1; fi
done
[ "$CHECK_STATUS" -eq 0 ] || exit "$CHECK_STATUS"
[ "$MODE" != --check ] || exit 0
DISPLAY_NAME="$("$PYTHON" -c 'import sys,yaml; print(yaml.safe_load(open(sys.argv[1]))["display_name"])' "$DIR/catalog.yaml")"
xcodebuild -project DocTools.xcodeproj -scheme DocTools -configuration Release -derivedDataPath "$DIR/build/DerivedData" build | tail -3
BUILT="$DIR/build/DerivedData/Build/Products/Release"
APP="$BUILT/DocTools.app"
[ -d "$APP" ]
plutil -replace CFBundleName -string "$DISPLAY_NAME" "$APP/Contents/Info.plist"
plutil -replace CFBundleDisplayName -string "$DISPLAY_NAME" "$APP/Contents/Info.plist"
plutil -replace CFBundleIconFile -string AppIcon "$APP/Contents/Info.plist"
plutil -replace CFBundleVersion -string "$(git -C "$DIR" rev-list --count HEAD)" "$APP/Contents/Info.plist"
cp icon/AppIcon.icns "$APP/Contents/Resources/AppIcon.icns"
# dockit 进包(签名前),Contents/Resources/bin/dockit exec 共享 venv 里的同一个后端
mkdir -p "$APP/Contents/Resources/bin"
install -m 755 bin/dockit "$APP/Contents/Resources/bin/dockit"
codesign --force -s - "$APP"
# 装进包的那条链:包内 bin/dockit 把 config / update 交给自己所在包的 App 程序(只跑读命令,隔离偏好域)
DOCKIT_APP="$APP" "$PYTHON" tests/test_lifecycle_cli.py AssembledBundleTests
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
# agent 入口:~/.local/bin/dockit → 已装 App 包内的包装脚本(Chapter cli_entry 验这条链接)
mkdir -p "$HOME/.local/bin"
ln -sfn "$DEST/Contents/Resources/bin/dockit" "$HOME/.local/bin/dockit"
echo "Linked: $HOME/.local/bin/dockit"
