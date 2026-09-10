import Foundation
import SwiftUI

// =============================================================================
// DocTools — ViewModel（蒸馏自 ssot-console ViewModel.swift 2026-06-11）
//
// 形态铁律：
//   · @MainActor ObservableObject；所有后端调用 async，错误一律进 banner（人话），
//     不弹 alert、不 print 了事 —— detail 根级常驻 StatusBanner 槽保证无选中也可见。
//   · busy 标志按动作拆独立 @Published；这里 isLoadingOps（启动列操作）/ isRunning（跑操作）拆开。
//   · Swift 零业务：拖入文件 → 选操作 → 调后端 → 显结果，全部委托 doc_gui_backend.py。
// =============================================================================

/// 常驻状态/错误 banner（detail 根级一个，无选中项也能看到）。
struct BannerMsg: Equatable {
    enum Kind { case error, warning, info }
    var kind: Kind
    var text: String

    static func error(_ t: String) -> BannerMsg { .init(kind: .error, text: t) }
    static func warning(_ t: String) -> BannerMsg { .init(kind: .warning, text: t) }
    static func info(_ t: String) -> BannerMsg { .init(kind: .info, text: t) }
}

/// 一个拖入/选取的文件条目（用绝对路径去重）。
struct InputFile: Identifiable, Hashable {
    let id: String         // = path（去重键）
    var path: String { id }
    var name: String { (path as NSString).lastPathComponent }
    var ext: String { (path as NSString).pathExtension.lowercased() }
    var isDir: Bool {
        var d: ObjCBool = false
        FileManager.default.fileExists(atPath: path, isDirectory: &d)
        return d.boolValue
    }
}

@MainActor
final class AppViewModel: ObservableObject {
    @Published var banner: BannerMsg?
    @Published var isLoadingOps = false
    @Published var isRunning = false

    @Published var ops: [DocOp] = []
    @Published var selectedOpID: String?
    @Published var selectedTargetID: String?     // 单选槽（convert/renum/bidfinal）

    /// per-op 勾选项状态：option id → 开/关。切操作时按后端声明的 default 复位。
    /// key 全部来自后端声明，Swift 不认识任何具体 id。
    @Published var optionValues: [String: Bool] = [:]

    /// per-op 文件型选项状态：option id → 绝对路径（type=file，如「范式 docx」）。
    /// 与 optionValues 分开存是因为值域不同（路径 vs 布尔），不是因为它特殊 ——
    /// Swift 依旧不认识任何具体 id，只按后端声明的 type 分派。
    @Published var optionPaths: [String: String] = [:]

    @Published var files: [InputFile] = []
    @Published var results: [FileResult] = []
    @Published var lastLog: String = ""
    @Published var summary: String = ""          // "成功 N/M" 之类
    @Published var statusText: String = "拖入文件或点「选择文件」，再挑一个操作。"

    private let backend = BackendClient()

    var selectedOp: DocOp? { ops.first { $0.id == selectedOpID } }

    var canRun: Bool {
        guard let op = selectedOp, !isRunning, !files.isEmpty else { return false }
        if op.needsTarget && (selectedTargetID?.isEmpty ?? true) { return false }
        // 必填选项没填 → 不让点「执行」。后端也会拦（契约层 required 校验），
        // 但让按钮直接灰掉比跑一趟再报错更诚实。
        for o in op.options where o.required {
            let filled = o.isFile ? !(optionPaths[o.id] ?? "").isEmpty : true
            if !filled { return false }
        }
        return true
    }

    // MARK: - 启动：列操作

    func loadOps() async {
        isLoadingOps = true
        defer { isLoadingOps = false }
        do {
            let r = try await backend.ops()
            ops = r.ops
            if selectedOpID == nil { selectedOpID = ops.first?.id }
            syncTargetDefault()
            if banner?.kind == .error { banner = nil }
            statusText = "已就绪 · 共 \(ops.count) 个操作。拖入文件开始。"
        } catch is CancellationError {
        } catch {
            banner = .error("加载操作列表失败：\(error.localizedDescription)")
            statusText = "后端不可达（UI 不崩，先排查 uv / 路径）。"
        }
    }

    /// 切操作时：clean→convert 等切换后，把 target 复位到该操作的第一个目标（或清空）。
    func onOpChanged() {
        results = []; lastLog = ""; summary = ""
        syncTargetDefault()
        resetOptionsToDefaults()
    }

    /// 把勾选项复位成后端声明的默认值（切操作、以及「恢复默认」按钮都走这里）。
    func resetOptionsToDefaults() {
        guard let op = selectedOp else { optionValues = [:]; optionPaths = [:]; return }
        optionValues = Dictionary(uniqueKeysWithValues:
            op.options.filter { !$0.isFile }.map { ($0.id, $0.defaultOn) })
        optionPaths = [:]
    }

    /// 要传给后端的选项（统一成字符串，bool → "1"/"0"，file → 绝对路径）。
    /// bool 与默认值相同的不传（命令行短、语义清晰）；file 只要填了就传。
    var changedOptions: [String: String] {
        guard let op = selectedOp else { return [:] }
        var out: [String: String] = [:]
        for o in op.options {
            if o.isFile {
                let p = optionPaths[o.id] ?? ""
                if !p.isEmpty { out[o.id] = p }
            } else if let v = optionValues[o.id], v != o.defaultOn {
                out[o.id] = v ? "1" : "0"
            }
        }
        return out
    }

    private func syncTargetDefault() {
        guard let op = selectedOp, op.needsTarget else { selectedTargetID = nil; return }
        if selectedTargetID == nil || !op.targets.contains(where: { $0.id == selectedTargetID }) {
            selectedTargetID = op.targets.first?.id
        }
    }

    // MARK: - 文件增删

    func addPaths(_ paths: [String]) {
        var seen = Set(files.map(\.id))
        for p in paths where !seen.contains(p) {
            files.append(InputFile(id: p)); seen.insert(p)
        }
        results = []; summary = ""
        statusText = "\(files.count) 个待处理。"
    }

    func remove(_ f: InputFile) {
        files.removeAll { $0.id == f.id }
        statusText = files.isEmpty ? "已清空。" : "\(files.count) 个待处理。"
    }

    func clearFiles() {
        files = []; results = []; lastLog = ""; summary = ""
        statusText = "已清空。拖入文件或点「选择文件」。"
    }

    // MARK: - 跑操作

    func run() async {
        guard canRun, let op = selectedOp else { return }
        isRunning = true
        defer { isRunning = false }
        results = []; summary = ""; lastLog = ""
        statusText = "正在执行「\(op.title)」…"
        let target = op.needsTarget ? selectedTargetID : nil
        let paths = files.map(\.path)
        do {
            let r = try await backend.run(op: op.id, target: target,
                                          options: changedOptions, files: paths)
            results = r.results
            lastLog = r.log
            if op.wantsDir {
                summary = r.results.first?.ok == true ? "扫描完成" : "扫描出错"
            } else {
                summary = "成功 \(r.succeeded)/\(r.total)"
                if !r.skippedMissing.isEmpty {
                    summary += " · 跳过不存在 \(r.skippedMissing.count)"
                }
            }
            statusText = summary
            banner = nil
        } catch is CancellationError {
            statusText = "已取消。"
        } catch {
            banner = .error("执行失败：\(error.localizedDescription)")
            statusText = "执行失败（详见上方 banner）。"
        }
    }

    /// 在 Finder 中显示产出（揭示文件/目录）。
    func reveal(_ path: String) {
        let url = URL(fileURLWithPath: path)
        NSWorkspace.shared.activateFileViewerSelecting([url])
    }
}
