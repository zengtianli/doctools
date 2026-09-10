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
}

@main
struct DocToolsApp: App {
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
