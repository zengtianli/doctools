import SwiftUI
import UniformTypeIdentifiers
import AppKit

// =============================================================================
// DocTools — ContentView（蒸馏自 ssot-console ContentView.swift 2026-06-11）
//
// 设计语言（macOS 原生产品形态，坑单以代码形态固化）：
//   · 语义色零硬编码：.primary/.secondary/.tertiary + controlBackgroundColor /
//     windowBackgroundColor / separatorColor；状态色 green/orange/red 只点缀 icon 与 badge。
//   · 卡片：RoundedRectangle(10) 填 controlBackgroundColor + 0.5pt separator 描边（card()）。
//   · 侧栏 .listStyle(.sidebar)、List 当根、控件放 Section；动作进 .toolbar；
//     空态 ContentUnavailableView；代码/路径一律 .monospaced。
//   · 【禁 fixedSize(horizontal:false, vertical:true)】窗口最小尺寸探测用**零宽提案**，
//     fixedSize(v:true) 强迫长文本在零宽下逐字换行，150 字符中文/长命令 ≈ 2500px 最小高，
//     全卡累加 → 窗口高度锁死数千 px 不可调、且 min 随选中内容变。正常布局给 nil 高度
//     提案，长文本天然换行不截断 —— 根本不需要 fixedSize。诊断口诀：AX `set size`
//     看回弹值是否随内容变。
//   · detail 根 .frame(minWidth: 600) 兜底（极窄宽度逐字换行的二道防线）。
// Swift 只渲染，业务逻辑全委托 Python 后端。
// =============================================================================

// MARK: - 通用视觉组件

/// 卡片语言：RoundedRectangle(10) + controlBackgroundColor + 0.5pt separator 描边。
struct CardBackground: ViewModifier {
    var padding: CGFloat = 13
    func body(content: Content) -> some View {
        content
            .padding(padding)
            .background(RoundedRectangle(cornerRadius: 10)
                .fill(Color(nsColor: .controlBackgroundColor)))
            .overlay(RoundedRectangle(cornerRadius: 10)
                .stroke(Color(nsColor: .separatorColor), lineWidth: 0.5))
    }
}

extension View {
    func card(padding: CGFloat = 13) -> some View { modifier(CardBackground(padding: padding)) }
}

/// 常驻错误/警告/信息 banner：icon + 人话 + 可关闭。detail 根级常驻槽位。
struct StatusBanner: View {
    let msg: BannerMsg
    var onClose: () -> Void

    private var color: Color {
        switch msg.kind {
        case .error: return .red
        case .warning: return .orange
        case .info: return .secondary
        }
    }
    private var icon: String {
        switch msg.kind {
        case .error: return "exclamationmark.octagon.fill"
        case .warning: return "exclamationmark.triangle.fill"
        case .info: return "info.circle.fill"
        }
    }

    var body: some View {
        HStack(alignment: .firstTextBaseline, spacing: 8) {
            Image(systemName: icon).foregroundStyle(color)
            Text(msg.text)
                .font(.callout)
                .textSelection(.enabled)
            Spacer(minLength: 8)
            Button { onClose() } label: {
                Image(systemName: "xmark.circle.fill").foregroundStyle(.tertiary)
            }
            .buttonStyle(.borderless)
            .help("关闭")
        }
        .padding(.horizontal, 12).padding(.vertical, 8)
        .background(RoundedRectangle(cornerRadius: 8).fill(color.opacity(0.08)))
        .overlay(RoundedRectangle(cornerRadius: 8).stroke(color.opacity(0.3), lineWidth: 0.5))
    }
}

// MARK: - Root

struct ContentView: View {
    @StateObject private var vm = AppViewModel()
    @State private var showPalette = false   // ⌘K 命令面板浮层开关

    var body: some View {
        NavigationSplitView {
            // 侧栏：操作列表（List 当根，范本坑单：别在 SplitView 外再包容器）。
            List(selection: $vm.selectedOpID) {
                Section("操作") {
                    ForEach(vm.ops) { op in
                        Label {
                            VStack(alignment: .leading, spacing: 1) {
                                Text(op.title)
                                Text(op.subtitle)
                                    .font(.caption2)
                                    .foregroundStyle(.secondary)
                                    // 2 行(2026-07-27):副标题被截在一行时,截掉的恰好是**消歧信息**
                                    // ——「纯文本修复…(只想动引号 → 用「引号统一」)」正好断在提示前,
                                    // 而 13 个动词里 规范化/引号统一/字体统一/清页眉页脚 是近邻,
                                    // 消歧提示被截 = 拆动词的收益白丢。
                                    // ⚠ 只用 lineLimit,**不加 fixedSize(v:true)** —— 见本文件头坑单:
                                    // 零宽提案下它会让长文本逐字换行,把窗口最小高撑到 2500px。
                                    // lineLimit(2) 有上界,天然免疫那个放大器。
                                    .lineLimit(2)
                            }
                        } icon: {
                            Image(systemName: op.icon)
                        }
                        .help(op.subtitle)   // 悬停看全文:2 行仍放不下的长句由 tooltip 兜底
                        .tag(op.id)
                    }
                }
            }
            .listStyle(.sidebar)
            .navigationTitle((Bundle.main.object(forInfoDictionaryKey: "CFBundleDisplayName") as? String ?? "DocKit"))
            .navigationSplitViewColumnWidth(min: 220, ideal: 260, max: 360)
            .overlay {
                if vm.isLoadingOps && vm.ops.isEmpty { ProgressView() }
            }
        } detail: {
            DetailView(vm: vm)
        }
        .task { await vm.loadOps() }
        .onChange(of: vm.selectedOpID) { vm.onOpChanged() }
        .onReceive(NotificationCenter.default.publisher(for: .consoleRefresh)) { _ in
            Task { await vm.loadOps() }
        }
        .toolbar {
            ToolbarItemGroup {
                Button { vm.clearFiles() } label: {
                    Label("清空", systemImage: "trash")
                }
                .help("清空待处理文件")
                .disabled(vm.files.isEmpty || vm.isRunning)
                Button { Task { await vm.loadOps() } } label: {
                    Label("刷新", systemImage: "arrow.clockwise")
                }
                .help("重读操作列表（⌘R）")
                .disabled(vm.isLoadingOps)
            }
        }
        // ⌘K 命令面板：搜索定位本 app 的文档操作，回车 = 切到该操作（纯增量接入）。
        .commandPalette(items: paletteItems, isPresented: $showPalette)
    }

    // MARK: - ⌘K 命令面板条目

    /// 把后端列出的文档操作（vm.ops）map 成可搜索条目；
    /// run = 用现成的 selectedOpID 机制切到该操作（与侧栏点选等价，触发 onOpChanged）。
    private var paletteItems: [PaletteItem] {
        vm.ops.map { op in
            PaletteItem(
                id: op.id,
                title: op.title,                 // 操作中文名
                subtitle: op.subtitle,           // 一行说明（分组/类别）
                icon: op.icon,                   // 操作已有的 SF Symbol
                keywords: paletteKeywords(for: op)  // 英文 verb/别名，拼写也能搜到
            ) {
                // 跳转 = 选中该操作（NavigationSplitView 侧栏 selection），同侧栏点击一致。
                vm.selectedOpID = op.id
            }
        }
    }

    /// 英文搜索别名 —— **由后端 gui-ops 的 aliases 字段声明**（2026-07-27 下沉）。
    /// 原来是 Swift 里一张硬编码词典，只覆盖 6 个动词：新增 op 在 ⌘K 里搜英文一律命不中，
    /// 而补别名要重编 .app —— 与「加 op 只改 Python」的架构目标直接冲突。
    private func paletteKeywords(for op: DocOp) -> String {
        [op.verb, op.id, op.aliases].filter { !$0.isEmpty }.joined(separator: " ")
    }
}

// MARK: - Detail

struct DetailView: View {
    @ObservedObject var vm: AppViewModel

    var body: some View {
        VStack(alignment: .leading, spacing: 12) {
            // 常驻 banner 槽：无选中项也能看到错误（范本铁律）。
            if let b = vm.banner {
                StatusBanner(msg: b) { vm.banner = nil }
            }
            if let op = vm.selectedOp {
                ScrollView {
                    VStack(alignment: .leading, spacing: 12) {
                        opHeader(op)
                        DropZone(op: op, vm: vm)
                        if !vm.files.isEmpty { fileListCard(op) }
                        runBar(op)
                        if !vm.results.isEmpty { resultsCard }
                    }
                    .frame(maxWidth: .infinity, alignment: .topLeading)
                }
                statusLine
            } else {
                ContentUnavailableView(
                    "选择一个操作",
                    systemImage: "sidebar.left",
                    description: Text("从左侧选择规范化 / 转换 / 拆分 / 合并 / 套模板 / 扫描。"))
                    .frame(maxWidth: .infinity, maxHeight: .infinity)
            }
        }
        .padding(16)
        // minWidth 兜底：窗口最小尺寸探测用零宽提案，长文本在极窄宽度下逐字换行会把
        // 最小高度撑爆（fixedSize(v:true) 是放大器，本模板全文禁用）。
        .frame(minWidth: 600, maxWidth: .infinity, maxHeight: .infinity, alignment: .topLeading)
        .navigationTitle(vm.selectedOp?.title ?? (Bundle.main.object(forInfoDictionaryKey: "CFBundleDisplayName") as? String ?? "DocKit"))
    }

    // MARK: 操作头 + 目标格式选择器（convert 专用）

    private func opHeader(_ op: DocOp) -> some View {
        VStack(alignment: .leading, spacing: 8) {
            HStack(spacing: 8) {
                Image(systemName: op.icon).font(.title3)
                    .foregroundStyle(op.danger ? Color.red : Color.accentColor)
                Text(op.title).font(.headline)
                if op.danger {
                    Text("破坏性").font(.caption2).bold()
                        .padding(.horizontal, 5).padding(.vertical, 1)
                        .background(Color.red.opacity(0.15), in: Capsule())
                        .foregroundStyle(.red)
                }
                Spacer()
                if vm.isRunning { ProgressView().controlSize(.small) }
            }
            Text(op.subtitle).font(.callout).foregroundStyle(.secondary)
            if !op.exts.isEmpty {
                Text("支持源格式：" + op.exts.joined(separator: " · "))
                    .font(.caption).foregroundStyle(.tertiary)
            }
            if op.needsTarget {
                Divider()
                HStack(spacing: 8) {
                    Text("目标格式").font(.callout)
                    Picker("", selection: Binding(
                        get: { vm.selectedTargetID ?? op.targets.first?.id ?? "" },
                        set: { vm.selectedTargetID = $0 })) {
                        ForEach(op.targets) { t in Text(t.title).tag(t.id) }
                    }
                    .labelsHidden()
                    .pickerStyle(.segmented)
                    .fixedSize()   // 仅短标签的分段控件，安全（禁的是长文本 v:true）
                    Spacer()
                }
            }
            if op.hasOptions {
                Divider()
                optionsPanel(op)
            }
        }
        .card()
    }

    // MARK: 勾选项面板（后端 gui-ops 声明驱动，Swift 只按 type 泛化渲染）
    //
    // ⚠ 这里不许出现任何 option id 字面量。要加一个勾选框 = 只改 Python 的
    //   doc_gui_backend.OPS[...]["options"]，不重编 .app。

    private func optionsPanel(_ op: DocOp) -> some View {
        VStack(alignment: .leading, spacing: 10) {
            HStack {
                Text("选项").font(.callout).bold()
                Spacer()
                Button("恢复默认") { vm.resetOptionsToDefaults() }
                    .buttonStyle(.link).font(.caption)
                    .disabled(vm.changedOptions.isEmpty)
            }
            ForEach(op.groupedOptions, id: \.group.id) { entry in
                optionGroup(entry.group, entry.items, op: op)
            }
        }
    }

    private func optionGroup(_ g: OpOptionGroup, _ items: [OpOption], op: DocOp) -> some View {
        // applies_to 非空且当前文件都不是这些后缀 → 灰显（拖了 pptx 时「改哪些范围」无意义）
        let applicable = g.appliesTo.isEmpty || vm.files.isEmpty
            || vm.files.contains { g.appliesTo.contains(($0.path as NSString).pathExtension.lowercased()) }
        return VStack(alignment: .leading, spacing: 4) {
            HStack(spacing: 6) {
                Text(g.title).font(.caption).bold()
                    .foregroundStyle(g.danger ? Color.red : .secondary)
                if !applicable, !g.appliesTo.isEmpty {
                    Text("(仅 " + g.appliesTo.joined(separator: "/") + ")")
                        .font(.caption2).foregroundStyle(.tertiary)
                }
            }
            // 按后端声明的 type 分派渲染。这里依旧不出现任何 option id 字面量 ——
            // 加一个新旋钮 = 只改 Python；加一个新**类型**才需要动这里。
            ForEach(items) { o in
                if o.isFile {
                    fileOptionRow(o, danger: g.danger, enabled: applicable)
                } else {
                    Toggle(isOn: Binding(
                        get: { vm.optionValues[o.id] ?? o.defaultOn },
                        set: { vm.optionValues[o.id] = $0 })) {
                        HStack(spacing: 5) {
                            Text(o.title).font(.callout)
                                .foregroundStyle(g.danger ? Color.red : .primary)
                            if !o.note.isEmpty {
                                Text(o.note).font(.caption2).foregroundStyle(.tertiary)
                            }
                        }
                    }
                    .toggleStyle(.checkbox)
                    .disabled(!applicable)
                }
            }
        }
        .opacity(applicable ? 1 : 0.45)
    }

    /// type=file 的选项行：标题 + 说明 + 「选择…」+ 已选文件名（可清除）。
    private func fileOptionRow(_ o: OpOption, danger: Bool, enabled: Bool) -> some View {
        let picked = vm.optionPaths[o.id] ?? ""
        return VStack(alignment: .leading, spacing: 3) {
            HStack(spacing: 5) {
                Text(o.title).font(.callout).foregroundStyle(danger ? Color.red : .primary)
                if o.required {
                    Text("必选").font(.caption2).foregroundStyle(picked.isEmpty ? Color.red : .secondary)
                }
                if !o.note.isEmpty {
                    Text(o.note).font(.caption2).foregroundStyle(.tertiary)
                }
            }
            HStack(spacing: 8) {
                Button("选择…") { pickOptionFile(o) }.controlSize(.small)
                if picked.isEmpty {
                    Text("未选").font(.caption).foregroundStyle(.tertiary)
                } else {
                    Text((picked as NSString).lastPathComponent)
                        .font(.caption).foregroundStyle(.secondary)
                        .lineLimit(1).truncationMode(.middle)
                        .help(picked)                      // 悬停看全路径
                    Button {
                        vm.optionPaths[o.id] = nil
                    } label: { Image(systemName: "xmark.circle.fill") }
                        .buttonStyle(.plain).foregroundStyle(.tertiary)
                        .help("清除")
                }
            }
        }
        .disabled(!enabled)
    }

    private func pickOptionFile(_ o: OpOption) {
        let panel = NSOpenPanel()
        panel.allowsMultipleSelection = false
        panel.canChooseDirectories = false
        // 后端声明限定后缀。allowedFileTypes 已软弃用 → 走 UTType；
        // 映射不出来时**不静默放行全部**，退回旧 API 保住限制（宁可用弃用 API，
        // 也不要一个"看起来在过滤、实际什么都能选"的选择器）。
        if !o.exts.isEmpty {
            let types = o.exts.compactMap { UTType(filenameExtension: $0) }
            if types.count == o.exts.count {
                panel.allowedContentTypes = types
            } else {
                panel.allowedFileTypes = o.exts
            }
        }
        panel.prompt = "选择"
        panel.message = o.title
        if panel.runModal() == .OK, let u = panel.url {
            vm.optionPaths[o.id] = u.path
        }
    }

    // MARK: 拖放区 / 文件列表

    private func fileListCard(_ op: DocOp) -> some View {
        VStack(alignment: .leading, spacing: 6) {
            HStack {
                Text(op.wantsDir ? "待扫描目录" : "待处理文件（\(vm.files.count)）")
                    .font(.subheadline.bold())
                Spacer()
            }
            Divider()
            ForEach(vm.files) { f in
                HStack(spacing: 8) {
                    Image(systemName: f.isDir ? "folder.fill" : "doc.fill")
                        .foregroundStyle(.secondary)
                    Text(f.name).font(.callout).lineLimit(1).truncationMode(.middle)
                    if !f.ext.isEmpty {
                        Text(f.ext.uppercased())
                            .font(.caption2.bold())
                            .padding(.horizontal, 5).padding(.vertical, 1)
                            .background(Capsule().fill(Color.accentColor.opacity(0.15)))
                    }
                    Spacer()
                    Text(f.path).font(.caption.monospaced())
                        .foregroundStyle(.tertiary).lineLimit(1).truncationMode(.head)
                    Button { vm.remove(f) } label: {
                        Image(systemName: "xmark.circle.fill").foregroundStyle(.tertiary)
                    }
                    .buttonStyle(.borderless)
                }
            }
        }
        .card()
    }

    // MARK: 运行栏

    private func runBar(_ op: DocOp) -> some View {
        HStack(spacing: 12) {
            Button { pickFiles(op) } label: {
                Label(op.wantsDir ? "选择目录" : "选择文件", systemImage: "plus")
            }
            .disabled(vm.isRunning)
            Spacer()
            if !vm.summary.isEmpty {
                Text(vm.summary).font(.callout).foregroundStyle(.secondary)
            }
            Button { Task { await vm.run() } } label: {
                Label(vm.isRunning ? "执行中…" : "执行", systemImage: "play.fill")
                    .frame(minWidth: 80)
            }
            .keyboardShortcut(.return, modifiers: [.command])
            .buttonStyle(.borderedProminent)
            .disabled(!vm.canRun)
        }
        .card(padding: 10)
    }

    // MARK: 结果

    private var resultsCard: some View {
        VStack(alignment: .leading, spacing: 8) {
            Text("结果").font(.headline)
            Divider()
            ForEach(vm.results) { r in
                VStack(alignment: .leading, spacing: 4) {
                    HStack(spacing: 8) {
                        Image(systemName: r.ok ? "checkmark.circle.fill" : "xmark.circle.fill")
                            .foregroundStyle(r.ok ? Color.green : Color.red)
                        Text(r.name).font(.callout.bold())
                        Spacer()
                        Text(r.message).font(.caption).foregroundStyle(.secondary)
                            .lineLimit(2).truncationMode(.tail)
                    }
                    if !r.outputs.isEmpty {
                        ForEach(r.outputs, id: \.self) { out in
                            HStack(spacing: 6) {
                                Image(systemName: "arrow.turn.down.right")
                                    .font(.caption2).foregroundStyle(.tertiary)
                                Text((out as NSString).lastPathComponent)
                                    .font(.caption.monospaced())
                                Spacer()
                                Button("在 Finder 显示") { vm.reveal(out) }
                                    .buttonStyle(.link).font(.caption)
                            }
                            .padding(.leading, 24)
                        }
                    }
                }
                .padding(.vertical, 3)
            }
            if !vm.lastLog.isEmpty {
                DisclosureGroup("后端日志") {
                    Text(vm.lastLog)
                        .font(.caption.monospaced())
                        .foregroundStyle(.secondary)
                        .textSelection(.enabled)
                        .frame(maxWidth: .infinity, alignment: .leading)
                }
                .font(.caption)
            }
        }
        .card()
    }

    private var statusLine: some View {
        HStack(spacing: 6) {
            if vm.isRunning { ProgressView().controlSize(.small) }
            Text(vm.statusText)
                .font(.caption).foregroundStyle(.secondary).lineLimit(1)
            Spacer()
        }
    }

    // MARK: 文件选择器

    private func pickFiles(_ op: DocOp) {
        let panel = NSOpenPanel()
        panel.allowsMultipleSelection = !op.wantsDir
        panel.canChooseDirectories = op.wantsDir
        panel.canChooseFiles = !op.wantsDir
        if panel.runModal() == .OK {
            if op.wantsDir {
                vm.clearFiles()  // scan 单目录，先清旧
            }
            vm.addPaths(panel.urls.map(\.path))
        }
    }
}

// MARK: - 拖放区

struct DropZone: View {
    let op: DocOp
    @ObservedObject var vm: AppViewModel
    @State private var hovering = false

    var body: some View {
        VStack(spacing: 8) {
            Image(systemName: op.wantsDir ? "folder.badge.plus" : "square.and.arrow.down")
                .font(.system(size: 30))
                .foregroundStyle(hovering ? Color.accentColor : .secondary)
            Text(op.wantsDir ? "把一个目录拖到这里" : "把文件拖到这里")
                .font(.callout).foregroundStyle(.secondary)
            Text("或用下方「\(op.wantsDir ? "选择目录" : "选择文件")」按钮")
                .font(.caption2).foregroundStyle(.tertiary)
        }
        .frame(maxWidth: .infinity)
        .padding(.vertical, 22)
        .background(RoundedRectangle(cornerRadius: 10)
            .fill(Color(nsColor: .controlBackgroundColor)))
        .overlay(RoundedRectangle(cornerRadius: 10)
            .strokeBorder(style: StrokeStyle(lineWidth: 1.2, dash: [6, 4]))
            .foregroundStyle(hovering ? Color.accentColor : Color(nsColor: .separatorColor)))
        .onDrop(of: [.fileURL], isTargeted: $hovering) { providers in
            handleDrop(providers)
        }
    }

    private func handleDrop(_ providers: [NSItemProvider]) -> Bool {
        var collected: [String] = []
        let group = DispatchGroup()
        for p in providers {
            group.enter()
            _ = p.loadObject(ofClass: URL.self) { url, _ in
                if let url { collected.append(url.path) }
                group.leave()
            }
        }
        group.notify(queue: .main) {
            // scan 只吃单目录：拖入只取第一个
            if op.wantsDir {
                vm.clearFiles()
                if let first = collected.first { vm.addPaths([first]) }
            } else {
                vm.addPaths(collected)
            }
        }
        return true
    }
}
