import AppKit
import SwiftUI

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
        AppConfiguration(productID: productID, defaultsKeys: AppViewModel.portablePreferenceKeys, defaults: preferences)
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
        configuration.onChange = {
            NotificationCenter.default.post(name: .dockitPreferencesChanged, object: nil)
        }
        AppLifecycleUI.install(name: name, configuration: configuration, updateSource: updateSource)
        AppLifecycleCLI.follow(configuration)
        return configuration
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
        if let toggle = view as? NSSwitch { return toggle }
        for child in view.subviews { if let found = cloudSwitch(in: child) { return found } }
        return nil
    }
}
