import AppKit
import SwiftUI
import CoreFoundation

// =============================================================================
// 「配置与更新…」窗口与 `dockit config` / `dockit update` 共用的一处定义。
//
// 这几项属于 App 本身(记住的设置所在的偏好域、版本与构建号、发行渠道),所以命令放在本可执行文件里:
// `DocToolsMain.main()` 在第一个参数是 config 或 update 时只跑共用命令层 AppLifecycleCLI.swift
// (总部 swift-shared 的逐字节副本,不在这里改)并退出,不创建 NSApplication —— 没有窗口、不进 Dock、不弹任何提示。
// dockit 的 Python 后端只转调(doc_gui_backend.lifecycle_command),stdout、stderr、退出码原样带回。
//
// 运行中的窗口靠 `installApp` 里那一行 `AppLifecycleCLI.follow` 跟随命令的改动;产品不自己重放开关值。
// =============================================================================

@MainActor
enum Lifecycle {
    static let name = "DocKit"
    static let command = "dockit"
    static let productID = "cyou.tianli.DocTools"
    static let updateSource: AppUpdateSource = .privateCloud(channel: "private")
    static let suiteVariable = "DOCKIT_LIFECYCLE_SUITE"
    static let probeFlag = "--lifecycle-follow-probe"
    static let refusal = "APP_LIFECYCLE_SUPPORT_DIR is set (an isolated run): DOCKIT_LIFECYCLE_SUITE (a throwaway named preferences domain, not DocKit's own) is required too, so the run never reaches the owner's own settings."

    static var isolated: Bool { ProcessInfo.processInfo.environment["APP_LIFECYCLE_SUPPORT_DIR"] != nil }

    /// Where the remembered settings live: the app's own preferences. In an isolated run (tests/test_lifecycle_cli.py)
    /// the shared layer keeps to its temporary support and cloud directories, and the preferences must come from the
    /// named throwaway domain in DOCKIT_LIFECYCLE_SUITE; nil when it is missing, names DocKit's own domain, or is a
    /// path (a path-based domain does not travel between two running processes).
    static func preferences() -> UserDefaults? {
        guard isolated else { return .standard }
        guard let suite = ProcessInfo.processInfo.environment[suiteVariable], !suite.isEmpty, !suite.contains("/"),
              suite != productID, suite != Bundle.main.bundleIdentifier else { return nil }
        return UserDefaults(suiteName: suite)
    }

    static func makeConfiguration(_ preferences: UserDefaults) -> AppConfiguration {
        AppConfiguration(productID: productID,
                         defaultsKeys: AppViewModel.portablePreferenceKeys + AppShortcuts.portableKeys + ["appUpdateOptions"],
                         defaults: preferences, validateDefault: { key, value in
            switch key {
            case AppShortcuts.enabledKey:
                guard let number = value as? NSNumber, CFGetTypeID(number) == CFBooleanGetTypeID() else { throw AppUpdateError("快捷键开关必须是布尔值。") }
            case AppShortcuts.bindingKey:
                guard let object = value as? [String: Any], let key = object["key"] as? String,
                      let flags = object["modifiers"] as? NSNumber else { throw AppUpdateError("快捷键格式错误。") }
                try AppShortcuts.validate(.init(key: key, modifiers: .init(rawValue: flags.uintValue)))
            case "appUpdateOptions":
                let data = try JSONSerialization.data(withJSONObject: value)
                try AppUpdatePreferences.validate(JSONDecoder().decode(AppUpdateOptions.self, from: data))
            case AppViewModel.portablePreferenceKeys[0]:
                guard value is String else { throw AppUpdateError("上次操作必须是文字。") }
            case AppViewModel.portablePreferenceKeys[1]:
                guard value is [String: String] else { throw AppUpdateError("目标格式必须是文字映射。") }
            default: break
            }
        })
    }

    static var shortcuts: AppShortcuts?
    private static var launchObserver: NSObjectProtocol?
    private static var shortcutToggle = NSSwitch()
    private static var shortcutBinding = NSPopUpButton()
    private static var shortcutDetail = AppLifecycleUI.detailLabel()
    private static let settingsActions = SettingsActions()

    @MainActor private final class SettingsActions: NSObject {
        @objc func toggle(_ sender: NSSwitch) { shortcuts?.setEnabled(sender.state == .on); refreshShortcuts() }
        @objc func binding(_ sender: NSPopUpButton) {
            guard let key = sender.selectedItem?.representedObject as? String else { return }
            do { try shortcuts?.setBinding(.init(key: key, modifiers: .command)); refreshShortcuts() }
            catch { shortcutDetail.stringValue = error.localizedDescription }
        }
    }

    private static func refreshShortcuts() {
        shortcuts?.refresh()
        shortcutToggle.state = shortcuts?.enabled == true ? .on : .off
        shortcutBinding.selectItem(withTitle: "⌘" + (shortcuts?.binding.key.uppercased() ?? "R"))
        shortcutDetail.stringValue = shortcuts?.descriptionText ?? ""
    }

    static func product(_ preferences: UserDefaults) -> AppLifecycleCLI.Product {
        AppLifecycleCLI.Product(command: command, name: name, configuration: makeConfiguration(preferences), updateSource: updateSource)
    }

    /// The window's wiring, called by the app at launch and by the follow probe: the shared window, the line that lets
    /// a running window follow `dockit config sync` / `dockit config import` made in another process, and the signal
    /// that makes the main window re-read the remembered operation and target formats.
    @discardableResult
    static func installApp(_ preferences: UserDefaults) -> AppConfiguration {
        let configuration = makeConfiguration(preferences)
        shortcuts = AppShortcuts(defaults: preferences) {
            NotificationCenter.default.post(name: .consoleRefresh, object: nil)
        }
        AppLifecycleUI.shared.productGroups = {
            shortcutToggle.target = settingsActions; shortcutToggle.action = #selector(SettingsActions.toggle(_:))
            shortcutToggle.identifier = NSUserInterfaceItemIdentifier("dockit.refresh.enabled")
            shortcutBinding.removeAllItems()
            for key in ["r", "j", "l"] { shortcutBinding.addItem(withTitle: "⌘" + key.uppercased()); shortcutBinding.lastItem?.representedObject = key }
            shortcutBinding.target = settingsActions; shortcutBinding.action = #selector(SettingsActions.binding(_:))
            refreshShortcuts()
            return [AppLifecycleUI.group("快捷键", rows: [
                AppLifecycleUI.row("刷新操作列表", detail: shortcutDetail, accessory: shortcutToggle),
                AppLifecycleUI.row("刷新组合", detail: AppLifecycleUI.detailLabel("只在 DocKit 内生效；系统与编辑组合保留，菜单冲突会拒绝。⌘K 搜索功能。"), accessory: shortcutBinding)])]
        }
        AppLifecycleUI.shared.willShow = { refreshShortcuts() }
        configuration.onChange = {
            refreshShortcuts()
            NotificationCenter.default.post(name: .dockitPreferencesChanged, object: nil)
        }
        AppLifecycleUI.install(name: name, configuration: configuration, updateSource: updateSource)
        AppLifecycleCLI.follow(configuration)
        if NSApp.activationPolicy() != .prohibited {
            launchObserver = NotificationCenter.default.addObserver(forName: NSApplication.didFinishLaunchingNotification, object: nil, queue: .main) { _ in
                MainActor.assumeIsolated {
                    shortcuts?.installMenu()
                    if let item = NSApp.mainMenu?.items.first?.submenu?.items.first(where: { $0.title == "配置与更新…" }) {
                        item.keyEquivalent = ","; item.keyEquivalentModifierMask = .command
                    }
                }
            }
        }
        return configuration
    }

    static func shortcutCommand(_ words: [String]) -> Int32 {
        guard let preferences = preferences() else { fputs(refusal + "\n", stderr); return 1 }
        let adapter = AppShortcuts(defaults: preferences, action: {})
        let body: [String: Any] = ["ok": words == ["shortcuts", "status", "--json"], "enabled": adapter.enabled,
                                 "binding": adapter.binding.dictionary, "description": adapter.descriptionText,
                                 "scope": "application menu", "search": "⌘K", "settings": "⌘,"]
        print(String(decoding: try! JSONSerialization.data(withJSONObject: body, options: [.sortedKeys]), as: UTF8.self))
        return body["ok"] as? Bool == true ? 0 : 2
    }

    /// `dockit config …` and `dockit update …`, words starting at the verb. Called before any NSApplication exists.
    static func runCommand(_ words: [String]) -> Int32 {
        guard let preferences = preferences() else {
            let command = words.prefix(2).filter { !$0.hasPrefix("-") }.joined(separator: " ")
            let body: [String: Any] = ["ok": false, "command": command, "error": ["code": "isolation_incomplete", "message": refusal]]
            if words.contains("--json"), let data = try? JSONSerialization.data(withJSONObject: body, options: [.sortedKeys]) {
                print(String(decoding: data, as: UTF8.self))
            } else { fputs(refusal + "\n", stderr) }
            return 1
        }
        return AppLifecycleCLI.run(words, product: product(preferences))
    }
}

/// `--lifecycle-follow-probe <state file>`: this executable as the running app of tests/test_lifecycle_cli.py.
/// Only in an isolated run. Activation policy `.prohibited`: no Dock icon, no menu bar, nothing can be ordered in.
/// It runs the production `Lifecycle.installApp`, hosts the real main view on the real view model (never shown),
/// builds the shared window the way its menu item builds it (never shown), and writes what the configuration, the
/// window's own switch and the view model hold to the state file for the test to poll.
@MainActor
enum LifecycleProbe {
    private static var keep: [Any] = []
    private static var changes = 0
    private static var tick = 0
    private static var write: (() -> Void)?

    static func run(_ words: [String]) -> Never {
        guard Lifecycle.isolated, let preferences = Lifecycle.preferences(),
              let index = words.firstIndex(of: Lifecycle.probeFlag), words.count > index + 1 else {
            fputs("\(Lifecycle.probeFlag) <state file> runs only in an isolated run. " + Lifecycle.refusal + "\n", stderr)
            exit(64)
        }
        let state = URL(fileURLWithPath: words[index + 1])
        let application = NSApplication.shared
        application.setActivationPolicy(.prohibited)
        let configuration = Lifecycle.installApp(preferences)
        keep.append(NotificationCenter.default.addObserver(forName: .dockitPreferencesChanged, object: nil, queue: .main) { _ in
            MainActor.assumeIsolated { changes += 1 }
        })
        let model = AppViewModel(preferences: preferences)
        let window = NSWindow(contentRect: NSRect(x: -10000, y: -10000, width: 1000, height: 680),
                              styleMask: [.titled, .closable, .resizable], backing: .buffered, defer: false)
        window.isReleasedWhenClosed = false
        let view = NSHostingView(rootView: ContentView(viewModel: model, autoLoad: false))
        window.contentView = view
        keep.append(window)
        Task { @MainActor in
            await model.loadOps()
            guard !model.ops.isEmpty else {
                fputs("operations did not load: \(model.banner?.text ?? "no reason given")\n", stderr)
                exit(1)
            }
            let built: [String: Bool]
            do { built = try AppLifecycleUI.shared.offscreenSnapshot(to: state.deletingPathExtension().appendingPathExtension("png")) }
            catch { fputs("lifecycle window: \(error.localizedDescription)\n", stderr); exit(1) }
            let control = NSApp.windows.lazy.filter { $0 !== window }.compactMap { cloudSwitch(in: $0.contentView) }.first
            let keys = AppViewModel.portablePreferenceKeys
            write = {
                view.layoutSubtreeIfNeeded()   // the hidden main view keeps updating, as a visible one does
                tick += 1
                let shown = (control as? NSButton)?.state ?? (control as? NSSwitch)?.state
                let seen: [String: Any] = [
                    "enabled": configuration.enabled, "status": configuration.status, "changes": changes, "tick": tick,
                    "window_switch": shown.map { $0 == .on } ?? NSNull(), "window_built": built,
                    "windows_on_screen": NSApp.windows.filter(\.isVisible).count,
                    "policy_prohibited": NSApp.activationPolicy() == .prohibited,
                    "operations": model.ops.map(\.id),
                    "selected_operation": model.selectedOpID ?? NSNull(),
                    "selected_target": model.selectedTargetID ?? NSNull(),
                    "app_reads_operation": preferences.string(forKey: keys[0]) ?? NSNull(),
                    "app_reads_targets": preferences.dictionary(forKey: keys[1]) as? [String: String] ?? [:]]
                try? JSONSerialization.data(withJSONObject: seen, options: [.sortedKeys]).write(to: state, options: .atomic)
            }
            Timer.scheduledTimer(withTimeInterval: 0.05, repeats: true) { _ in MainActor.assumeIsolated { write?() } }
        }
        application.run()
        exit(0)
    }

    /// The「使用 iCloud 记住配置」control of the shared window: a checkbox in the copy vendored here, a switch in the current shared source.
    private static func cloudSwitch(in view: NSView?) -> NSControl? {
        guard let view else { return nil }
        if let button = view as? NSButton, button.title == "使用 iCloud 记住配置" { return button }
        func containsTitle(_ node: NSView) -> Bool {
            (node as? NSTextField)?.stringValue == "使用 iCloud 记住配置" || node.subviews.contains(where: containsTitle)
        }
        if view.subviews.contains(where: containsTitle), let toggle = view.subviews.compactMap({ $0 as? NSSwitch }).first { return toggle }
        for child in view.subviews { if let found = cloudSwitch(in: child) { return found } }
        return nil
    }
}

/// The real settings controls and application-menu action, rendered without ordering any window.
@MainActor
enum ProductBasicsSelfTest {
    static func run() async {
        var checks: [String: Bool] = [:]
        let output = URL(fileURLWithPath: ProcessInfo.processInfo.environment["SOP_OUT_DIR"]!)
        do {
            guard let preferences = Lifecycle.preferences() else { throw AppUpdateError(Lifecycle.refusal) }
            let model = AppViewModel(preferences: preferences)
            await model.loadOps()
            guard let operation = model.ops.last else { throw AppUpdateError("操作目录为空。") }
            model.selectedOpID = operation.id; model.onOpChanged()
            _ = Lifecycle.installApp(preferences)
            NSApp.mainMenu = NSMenu()
            let search = NSMenuItem(title: "搜索功能…", action: nil, keyEquivalent: "k")
            search.keyEquivalentModifierMask = .command; NSApp.mainMenu!.addItem(search)
            guard let shortcuts = Lifecycle.shortcuts else { throw AppUpdateError("快捷键未接入。") }
            shortcuts.installMenu()
            let main = NSWindow(contentRect: NSRect(x: -20000, y: -20000, width: 1000, height: 680),
                                styleMask: [.titled, .closable], backing: .buffered, defer: false)
            main.isReleasedWhenClosed = false
            main.contentView = NSHostingView(rootView: ContentView(viewModel: model, autoLoad: false))
            main.contentView!.layoutSubtreeIfNeeded()
            try await Task.sleep(for: .milliseconds(200))
            for (name, appearance) in [("settings-light", NSAppearance.Name.aqua), ("settings-dark", .darkAqua),
                                       ("updates-light", .aqua), ("updates-dark", .darkAqua)] {
                let result = try AppLifecycleUI.shared.offscreenSnapshot(to: output.appendingPathComponent(name + ".png"),
                    appearance: appearance, showingUpdateOptions: name.hasPrefix("updates"))
                for (key, value) in result { checks[name + "." + key] = value }
            }
            func descendants(_ view: NSView) -> [NSView] { [view] + view.subviews.flatMap(descendants) }
            guard let toggle = NSApp.windows.compactMap(\.contentView).flatMap(descendants).first(where: {
                $0.identifier?.rawValue == "dockit.refresh.enabled"
            }) as? NSSwitch else { throw AppUpdateError("真实设置开关不存在。") }
            func toggleTo(_ state: NSControl.StateValue) {
                toggle.state = state
                _ = (toggle.target as? NSObject)?.perform(toggle.action!, with: toggle)
            }
            toggleTo(.off)
            checks["settings.switch_action"] = !shortcuts.enabled && preferences.bool(forKey: AppShortcuts.enabledKey) == false
            checks["shortcuts.disabled_equivalent"] = shortcuts.menuItem?.keyEquivalent == ""
            toggleTo(.on)
            checks["shortcuts.restored_equivalent"] = shortcuts.enabled && shortcuts.menuItem?.keyEquivalent == "r"
            checks["settings.saved_operation"] = preferences.string(forKey: AppViewModel.portablePreferenceKeys[0]) == operation.id
            let original = shortcuts.binding.dictionary as NSDictionary
            do { try shortcuts.setBinding(.init(key: "q", modifiers: .command)); checks["shortcuts.reserved_refused"] = false }
            catch { checks["shortcuts.reserved_refused"] = shortcuts.binding.dictionary as NSDictionary == original }
            do { try shortcuts.setBinding(.init(key: "k", modifiers: .command)); checks["shortcuts.conflict_refused"] = false }
            catch { checks["shortcuts.conflict_refused"] = shortcuts.binding.dictionary as NSDictionary == original }
            try shortcuts.setBinding(.init(key: "j", modifiers: .command))
            checks["shortcuts.custom_binding"] = shortcuts.menuItem?.keyEquivalent == "j"
            try shortcuts.setBinding(.init(key: "r", modifiers: .command))
            // Same callback and SwiftUI notification subscriber as the ordinary refresh menu.
            model.banner = .error("刷新验收占位")
            shortcuts.performAction()
            for _ in 0..<100 {
                if model.banner == nil && !model.isLoadingOps { break }
                try await Task.sleep(for: .milliseconds(30))
            }
            checks["shortcuts.actual_refresh_action"] = model.banner == nil && !model.ops.isEmpty && !model.isLoadingOps
            // The production checker, with its presentation flag off; automatic timers are disabled in isolation.
            AppLifecycleUI.shared.checkForUpdates(present: false)
            let bundleID = Bundle.main.bundleIdentifier!
            for _ in 0..<100 {
                if !(AppUpdatePreferences.defaults(bundleID).dictionary(forKey: "appUpdateState")?["last_error"] as? String ?? "").isEmpty { break }
                try await Task.sleep(for: .milliseconds(30))
            }
            checks["updates.actual_check_failure"] = !(AppUpdatePreferences.defaults(bundleID).dictionary(forKey: "appUpdateState")?["last_error"] as? String ?? "").isEmpty
            preferences.synchronize()
            checks["no_visible_windows"] = NSApp.windows.allSatisfy { !$0.isVisible && !$0.isKeyWindow }
            checks["no_activation_or_dock"] = NSApp.activationPolicy() == .prohibited && !NSApp.isActive
            main.close()
        } catch {
            checks["execution_completed"] = false; fputs("Product basics UI: \(error)\n", stderr)
        }
        let result: [String: Any] = ["checks": checks, "scope": "Real hidden settings controls, menu binding validation and refresh subscriber; no synthesized events."]
        let data = try! JSONSerialization.data(withJSONObject: result, options: [.sortedKeys, .prettyPrinted])
        try! data.write(to: output.appendingPathComponent("ui.json"), options: .atomic)
        print(String(decoding: data, as: UTF8.self))
        exit(!checks.isEmpty && checks.values.allSatisfy { $0 } ? 0 : 1)
    }
}
