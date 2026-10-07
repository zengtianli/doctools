import SwiftUI
import AppKit

// =============================================================================
// DocTools — App 入口
//
// 骨架蒸馏自 ssot-console（2026-06-11 活体）。形态铁律：
//   · WindowGroup 只放 ContentView；初始尺寸用 .defaultSize 给场景，
//     min 尺寸约束加在 NavigationSplitView 的 detail 根上（见 ContentView），
//     不在 WindowGroup / SplitView 上 .frame(min…)——那是范本踩过的布局崩坑。
//   · 菜单命令经 NotificationCenter 广播（⌘R 刷新），视图 onReceive 消费，
//     不直接持有 ViewModel —— 多窗口/多 tab 下天然解耦。
// =============================================================================

// MARK: - 跨视图通知（菜单命令 → 当前视图）

extension Notification.Name {
    static let consoleRefresh = Notification.Name("consoleRefresh")
    static let dockitPreferencesChanged = Notification.Name("dockitPreferencesChanged")
}

struct DocToolsApp: App {
    init() {
        // 「配置与更新…」窗口与 `dockit config` / `dockit update` 用同一个工厂(ProductLifecycle.swift)。
        Lifecycle.installApp(Lifecycle.preferences() ?? .standard)
    }

    var body: some Scene {
        WindowGroup {
            ContentView()
        }
        .defaultSize(width: 1000, height: 680)
        .commands {
            CommandMenu("操作") {
                Button("搜索功能…") {
                    NotificationCenter.default.post(name: .tlPaletteToggle, object: nil)
                }
                .keyboardShortcut("k", modifiers: .command)   // ⌘K 命令面板
                Button("刷新") {
                    NotificationCenter.default.post(name: .consoleRefresh, object: nil)
                }
                .keyboardShortcut("r", modifiers: .command)
            }
        }
    }
}

/// Self-test starts AppKit without a WindowGroup, menu or activation.
@main
enum DocToolsMain {
    @MainActor static func main() {
        let words = Array(CommandLine.arguments.dropFirst())
        // `dockit config …` / `dockit update …`:后端把这两个动词整段交到这里,共用命令层在窗口自己的配置与更新源上
        // 执行后直接退出。此时还没有 NSApplication:没有窗口、不进 Dock、不弹提示;已在运行的实例不被打扰,
        // 它经 AppLifecycleCLI.follow 自己跟随。
        if AppLifecycleCLI.handles(words.first) { exit(Lifecycle.runCommand(words)) }
        if words.contains(Lifecycle.probeFlag) { LifecycleProbe.run(words) }
        if LaneSignal.quiet {
            quietMain()
        } else if CommandLine.arguments.contains("--ui-self-test") {
            let application = NSApplication.shared
            application.setActivationPolicy(.prohibited)
            Task { await UISelfTest.run() }
            application.run()
        } else {
            DocToolsApp.main()
        }
    }

    @MainActor private static func quietMain() {
        let application = NSApplication.shared
        application.setActivationPolicy(.accessory)
        LaneSignal.enterQuietIfAsked()
        let model = AppViewModel(usesPortablePreferences: false)
        let window = NSWindow(contentRect: NSRect(x: -10000, y: -10000, width: 1000, height: 680),
                              styleMask: [.titled, .closable, .resizable], backing: .buffered, defer: false)
        window.isReleasedWhenClosed = false
        let view = NSHostingView(rootView: ContentView(viewModel: model, autoLoad: false))
        window.contentView = view
        let timeout = Timer.scheduledTimer(withTimeInterval: 30, repeats: false) { _ in NSApp.terminate(nil) }
        Task { @MainActor in
            await model.loadOps()
            guard !model.ops.isEmpty, model.banner?.kind != .error else {
                NSApp.terminate(nil)
                return
            }
            DispatchQueue.main.async {
                view.layoutSubtreeIfNeeded()
                guard let bitmap = view.bitmapImageRepForCachingDisplay(in: view.bounds) else {
                    NSApp.terminate(nil)
                    return
                }
                view.cacheDisplay(in: view.bounds, to: bitmap)
                timeout.invalidate()
                LaneSignal.ready("main")
            }
        }
        application.run()
        withExtendedLifetime(window) {}
    }
}
