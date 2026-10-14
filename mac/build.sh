#!/bin/bash
set -euo pipefail
MODE="${1:---build}"
case "$MODE" in
  --build|--check|--install) [ "$#" -le 1 ] || exit 2 ;;
  --install-built) [ "$#" -eq 3 ] || { echo "--install-built <signed.app> <build-receipt.json>" >&2; exit 2; } ;;
  --help) echo "Usage: $0 [--build|--check|--install|--install-built <signed.app> <receipt.json>] (installation requires stable Developer ID signing)"; exit 0 ;;
  *) echo "Unknown option: $MODE" >&2; exit 2 ;;
esac
DIR="$(cd "$(dirname "$0")" && pwd -P)"
cd "$DIR"
PYTHON="$HOME/Dev/.venv/bin/python"
if [ "$MODE" = --install-built ]; then
  exec "$PYTHON" scripts/install-built.py "$2" "$3"
fi
if [ "$MODE" = --install ]; then
  if [[ "${CODE_SIGN_IDENTITY:-}" != "Developer ID Application:"* ]]; then
    echo "Use the existing certificate host to sign, then --install-built <signed.app> <receipt.json>; ad-hoc installation refused." >&2
    exit 69
  fi
  "$PYTHON" "$HOME/Apps/chapter/engine/app_sop.py" build-receipt --app doc-tools \
    --artifact "$DIR/build/DerivedData/Build/Products/Release/DocTools.app" \
    --source-glob 'Sources/**' --source-glob icon/AppIcon.icns --source-glob Info.plist \
    --source-glob DocTools.xcodeproj/project.pbxproj --source-glob build.sh --source-glob 'bin/**' \
    --source-glob catalog.yaml --source-glob scripts/install-built.py \
    --build-command './build.sh --build' --receipt-out build/accept/install-candidate.json --json
  exec "$PYTHON" "$HOME/Apps/chapter/engine/app_sop.py" build-receipt --app doc-tools \
    --artifact /Applications/DocKit.app \
    --source-glob 'Sources/**' --source-glob icon/AppIcon.icns --source-glob Info.plist \
    --source-glob DocTools.xcodeproj/project.pbxproj --source-glob build.sh --source-glob 'bin/**' \
    --source-glob catalog.yaml --source-glob scripts/install-built.py \
    --build-command './build.sh --install-built build/DerivedData/Build/Products/Release/DocTools.app build/accept/install-candidate.json' --json
fi
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
if [ -n "${CODE_SIGN_IDENTITY:-}" ]; then
  codesign --force --sign "$CODE_SIGN_IDENTITY" --options runtime --timestamp "$APP"
else
  codesign --force -s - "$APP"
fi
# 装进包的那条链:包内 bin/dockit 把 config / update 交给自己所在包的 App 程序(只跑读命令,隔离偏好域)
DOCKIT_APP="$APP" "$PYTHON" tests/test_lifecycle_cli.py AssembledBundleTests
echo "Built: $APP"
