import SwiftUI
import AppKit

// ═══════════════════════════════════════════════════════════════════════════
// TL macapp 舰队 · 共享命令面板 CommandPalette —— ⌘K 搜索定位 app 内功能
//
// SSOT(唯一真相)= ~/Dev/tools/dev/lib/tools/macapp/swift-shared/CommandPalette.swift
//   各 app 的 Sources/CommandPalette.swift 是 vendored 副本(逐字节一致)。
//   改动只改 SSOT,再跑 `swift-shared/vendor-palette.sh <app-dir>` 重新分发;
//   禁就地改副本(drift 守卫 `vendor-palette.sh --check` 会红)。
//
// host 接入三步:
//   1) 根视图加   .commandPalette(items: <[PaletteItem]>, isPresented: $showPalette)
//   2) 加         @State private var showPalette = false
//   3) App.swift 的 .commands 里加一条 ⌘K 菜单(post .tlPaletteToggle),或直接用 PaletteCommands()
//   每个 PaletteItem.run = 选中后跳到该功能(通常 = 设置 selection / tab)。
// ═══════════════════════════════════════════════════════════════════════════

/// 一条可搜索的「功能」。run 闭包 = 选中后的跳转动作。
struct PaletteItem: Identifiable {
    let id: String
    let title: String           // 主名(中文,左侧主行)
    var subtitle: String = ""   // 分组/路径/描述(右侧灰字)
    var icon: String = ""       // SF Symbol 名 或 emoji(运行时自动识别)
    var keywords: String = ""   // 额外搜索词(key/英文别名/拼音,不显示、只参与匹配)
    let run: () -> Void
    var haystack: String { "\(title) \(subtitle) \(keywords)".lowercased() }
}

/// 匹配打分:标题前缀 > 标题子串 > 语料子串 > 子序列模糊;nil = 不命中。
enum PaletteMatch {
    static func score(_ item: PaletteItem, query raw: String) -> Int? {
        let q = raw.trimmingCharacters(in: .whitespaces).lowercased()
        guard !q.isEmpty else { return 0 }
        let title = item.title.lowercased()
        if title.hasPrefix(q) { return 1000 }
        if title.contains(q) { return 800 - max(0, title.count - q.count) }
        if item.haystack.contains(q) { return 500 }
        if isSubsequence(q, item.haystack) { return 200 }
        return nil
    }
    /// q 的字符按序出现在 s 中(latin 模糊匹配;CJK 退化为顺序包含,无害)。
    static func isSubsequence(_ q: String, _ s: String) -> Bool {
        var i = s.startIndex
        for ch in q {
            var hit = false
            while i < s.endIndex { let c = s[i]; i = s.index(after: i); if c == ch { hit = true; break } }
            if !hit { return false }
        }
        return true
    }
}

extension Notification.Name {
    /// App 菜单 ⌘K → post 这个 → 面板 toggle。
    static let tlPaletteToggle = Notification.Name("TLCommandPaletteToggle")
}

/// 面板状态(ref type:供 NSEvent 键盘监听稳定捕获,避免 struct @State 闭包陈旧)。
final class PaletteModel: ObservableObject {
    @Published var query: String = "" { didSet { recompute() } }
    @Published private(set) var results: [PaletteItem] = []
    @Published var sel: Int = 0
    private let all: [PaletteItem]

    init(_ items: [PaletteItem]) { all = items; results = items }

    func recompute() {
        let q = query.trimmingCharacters(in: .whitespaces)
        if q.isEmpty {
            results = all
        } else {
            results = all
                .compactMap { it in PaletteMatch.score(it, query: q).map { (it, $0) } }
                .sorted { $0.1 > $1.1 }
                .map(\.0)
        }
        sel = 0
    }
    func move(_ d: Int) {
        guard !results.isEmpty else { return }
        sel = (sel + d + results.count) % results.count
    }
    var current: PaletteItem? { results.indices.contains(sel) ? results[sel] : nil }
}

/// ⌘K 命令面板浮层。
struct CommandPalette: View {
    @Binding var isPresented: Bool
    @StateObject private var model: PaletteModel
    @FocusState private var focused: Bool
    @State private var monitor: Any?

    init(items: [PaletteItem], isPresented: Binding<Bool>) {
        _isPresented = isPresented
        _model = StateObject(wrappedValue: PaletteModel(items))
    }

    var body: some View {
        VStack(spacing: 0) {
            field
            Divider()
            if model.results.isEmpty { empty } else { list }
        }
        .frame(width: 560)
        .background(.regularMaterial, in: RoundedRectangle(cornerRadius: 14))
        .overlay(RoundedRectangle(cornerRadius: 14).strokeBorder(Color.primary.opacity(0.10)))
        .shadow(color: .black.opacity(0.28), radius: 28, y: 12)
        .onAppear {
            DispatchQueue.main.async { focused = true }
            startMonitor()
        }
        .onDisappear { stopMonitor() }
    }

    private var field: some View {
        HStack(spacing: 9) {
            Image(systemName: "magnifyingglass").foregroundStyle(.secondary)
            TextField("搜索功能…", text: $model.query)
                .textFieldStyle(.plain)
                .font(.title3)
                .focused($focused)
                .onSubmit(run)
            if !model.query.isEmpty {
                Button { model.query = "" } label: { Image(systemName: "xmark.circle.fill") }
                    .buttonStyle(.borderless).foregroundStyle(.tertiary)
            }
            Text("\(model.results.count)").font(.caption).monospacedDigit().foregroundStyle(.tertiary)
        }
        .padding(.horizontal, 15).padding(.vertical, 13)
    }

    private var list: some View {
        ScrollViewReader { proxy in
            ScrollView {
                LazyVStack(spacing: 1) {
                    ForEach(Array(model.results.enumerated()), id: \.element.id) { idx, item in
                        row(item, active: idx == model.sel)
                            .id(idx)
                            .contentShape(Rectangle())
                            .onTapGesture { model.sel = idx; run() }
                            .onHover { if $0 { model.sel = idx } }
                    }
                }
                .padding(6)
            }
            .frame(maxHeight: 360)
            .onChange(of: model.sel) { _ in
                withAnimation(.easeOut(duration: 0.12)) { proxy.scrollTo(model.sel, anchor: .center) }
            }
        }
    }

    private func row(_ item: PaletteItem, active: Bool) -> some View {
        HStack(spacing: 11) {
            icon(item.icon).frame(width: 22)
            Text(item.title).foregroundStyle(.primary).lineLimit(1)
            Spacer(minLength: 8)
            if !item.subtitle.isEmpty {
                Text(item.subtitle).font(.caption).foregroundStyle(.secondary).lineLimit(1)
            }
        }
        .padding(.horizontal, 11).padding(.vertical, 8)
        .background(RoundedRectangle(cornerRadius: 8).fill(active ? Color.accentColor.opacity(0.16) : .clear))
    }

    @ViewBuilder private func icon(_ s: String) -> some View {
        if s.isEmpty {
            Image(systemName: "circle.dotted").foregroundStyle(.secondary)
        } else if NSImage(systemSymbolName: s, accessibilityDescription: nil) != nil {
            Image(systemName: s).foregroundStyle(.tint)
        } else {
            Text(s)
        }
    }

    private var empty: some View {
        VStack(spacing: 8) {
            Image(systemName: "magnifyingglass").font(.title2).foregroundStyle(.tertiary)
            Text("没有匹配的功能").font(.callout).foregroundStyle(.secondary)
        }
        .frame(maxWidth: .infinity).frame(height: 120)
    }

    private func run() {
        guard let item = model.current else { return }
        isPresented = false
        DispatchQueue.main.async(execute: item.run)
    }

    private func startMonitor() {
        monitor = NSEvent.addLocalMonitorForEvents(matching: .keyDown) { e in
            switch e.keyCode {
            case 125: model.move(1); return nil      // ↓
            case 126: model.move(-1); return nil     // ↑
            case 36, 76: run(); return nil           // ⏎ / 小键盘 enter
            case 53: isPresented = false; return nil // esc
            default: return e
            }
        }
    }
    private func stopMonitor() {
        if let m = monitor { NSEvent.removeMonitor(m); monitor = nil }
    }
}

extension View {
    /// host 一行接入:.tlPaletteToggle 通知(=菜单 ⌘K)弹出命令面板浮层。
    func commandPalette(items: [PaletteItem], isPresented: Binding<Bool>) -> some View {
        overlay {
            if isPresented.wrappedValue {
                ZStack(alignment: .top) {
                    Color.black.opacity(0.16).ignoresSafeArea()
                        .onTapGesture { isPresented.wrappedValue = false }
                    CommandPalette(items: items, isPresented: isPresented)
                        .padding(.top, 64)
                }
            }
        }
        .onReceive(NotificationCenter.default.publisher(for: .tlPaletteToggle)) { _ in
            isPresented.wrappedValue.toggle()
        }
    }
}

/// App.swift 的 .commands 直接塞 PaletteCommands():加「搜索功能 ⌘K」菜单项。
/// (已有自定义 .commands 块的 app 可改为在原块里加一条同样的 Button,见各 app 接法。)
struct PaletteCommands: Commands {
    var body: some Commands {
        CommandGroup(after: .toolbar) {
            Button("搜索功能…") { NotificationCenter.default.post(name: .tlPaletteToggle, object: nil) }
                .keyboardShortcut("k", modifiers: .command)
        }
    }
}
