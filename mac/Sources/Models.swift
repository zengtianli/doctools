import Foundation

// =============================================================================
// DocTools — Codable 契约（镜像 Python 后端 JSON）
//
// 蒸馏自 ssot-console Models.swift（2026-06-11 活体）。契约约定：
//   · 所有 gui-* 子命令一律 exit 0；成功 {"ok": true, ...}，失败 {"ok": false, "error": "人话"}。
//   · Decoder 用 .convertFromSnakeCase：display_path → displayPath 等自动映射，
//     后端字段一律 snake_case，Swift 侧不写 CodingKeys 改名。
//   · 所有展示性字段一律 decodeIfPresent + 默认值 —— 后端少回一个字段
//     不准整个解码失败（范本铁律：列表里一条坏数据不准拖垮整页）。
//     Decodable init 写在 extension 里，保留 memberwise init 给 Swift 侧自行构造。
// =============================================================================

// MARK: - 信封探针

/// runDecoding 的第一道探针：{ok, error}。ok == false 即抛 error 人话文本。
struct BackendProbe: Decodable {
    let ok: Bool?
    let error: String?
}

/// 旧式失败信封（非零 exit + {"error"}）兼容：操作类子命令允许这种形态。
struct BackendErrorEnvelope: Codable {
    let error: String
}

// MARK: - 解码小工具（decodeIfPresent + 默认值的简写）

extension KeyedDecodingContainer {
    func str(_ key: Key, _ fallback: String = "") -> String {
        (try? decodeIfPresent(String.self, forKey: key)) ?? nil ?? fallback
    }
    func int(_ key: Key, _ fallback: Int = 0) -> Int {
        (try? decodeIfPresent(Int.self, forKey: key)) ?? nil ?? fallback
    }
    func bool(_ key: Key, _ fallback: Bool = false) -> Bool {
        (try? decodeIfPresent(Bool.self, forKey: key)) ?? nil ?? fallback
    }
    func strOpt(_ key: Key) -> String? {
        (try? decodeIfPresent(String.self, forKey: key)) ?? nil
    }
}

// MARK: - gui-ops：操作目录（UI 从此渲菜单）

/// convert 专用的目标格式选项。
struct OpTarget: Identifiable, Hashable {
    let id: String       // "md" / "word" / "xlsx" / "csv" / "txt"
    let title: String    // "Markdown" / "Word" …
}
extension OpTarget: Decodable {
    private enum CodingKeys: String, CodingKey { case id, title }
    init(from decoder: Decoder) throws {
        let c = try decoder.container(keyedBy: CodingKeys.self)
        id = c.str(.id); title = c.str(.title)
    }
}

/// 一个可勾选的 per-op 选项（后端 gui-ops 声明，Swift 按 type 泛化渲染）。
/// ⚠ 本文件与 ContentView 都**不许出现任何 option id 字面量** —— 一旦出现，
/// 「加选项只改 Python、不重编 app」这条就破了。
struct OpOption: Identifiable, Hashable, Decodable {
    let id: String            // "rule.quotes" / "scope.comments" / "ref"
    let group: String         // 归属分组 id
    let type: String          // "bool"（勾选框）/ "file"（文件选择器）
    let title: String
    let note: String          // 副标题（可空）
    let defaultOn: Bool
    let exts: [String]        // type=file 时限定可选后缀（空 = 不限）
    let required: Bool        // 必填：空值时禁用「执行」

    var isFile: Bool { type == "file" }

    private enum CodingKeys: String, CodingKey {
        case id, group, type, title, note, `default`, exts, required
    }
    init(from decoder: Decoder) throws {
        let c = try decoder.container(keyedBy: CodingKeys.self)
        id = c.str(.id); group = c.str(.group, "")
        type = c.str(.type, "bool"); title = c.str(.title)
        note = c.str(.note, "")
        defaultOn = (try? c.decodeIfPresent(Bool.self, forKey: .default)) ?? nil ?? true
        exts = (try? c.decodeIfPresent([String].self, forKey: .exts)) ?? nil ?? []
        required = (try? c.decodeIfPresent(Bool.self, forKey: .required)) ?? nil ?? false
    }
}

/// 选项分组（"修哪些内容" / "改哪些范围" / "破坏性"）。
struct OpOptionGroup: Identifiable, Hashable, Decodable {
    let id: String
    let title: String
    let danger: Bool          // true → 红色标题 + 默认折叠
    let appliesTo: [String]   // 只对这些后缀有意义（空 = 全部）

    // ⚠ decoder 开了 .convertFromSnakeCase：JSON 的 applies_to 到这里已变成 appliesTo，
    // CodingKey 写 snake_case 会永远匹配不上（静默解成空数组）。
    private enum CodingKeys: String, CodingKey { case id, title, danger, appliesTo }
    init(from decoder: Decoder) throws {
        let c = try decoder.container(keyedBy: CodingKeys.self)
        id = c.str(.id); title = c.str(.title)
        danger = (try? c.decodeIfPresent(Bool.self, forKey: .danger)) ?? nil ?? false
        appliesTo = (try? c.decodeIfPresent([String].self, forKey: .appliesTo)) ?? nil ?? []
    }
}

/// 一个文档操作（clean / convert / split / merge / typeset / scan）。
struct DocOp: Identifiable, Hashable {
    let id: String          // 操作 id，传给 gui-run --op
    let verb: String        // 底层 doc_dispatch 动词
    let title: String       // 中文标题
    let aliases: String     // 英文/别名搜索词（后端声明；⌘K 面板用）
    let subtitle: String    // 一行说明
    let icon: String        // SF Symbol
    let exts: [String]      // 支持的源后缀（拖入提示用）
    let kind: String        // "files"（多文件）/ "dir"（单目录）
    let targets: [OpTarget] // 单选参数槽（convert 目标格式 / renum 范围 / bidfinal 模式）
    let danger: Bool        // 破坏性动词（原地覆写 / 不可撤销）→ UI 标红
    let optionGroups: [OpOptionGroup]  // 勾选项分组（后端声明，可空）
    let options: [OpOption]            // 勾选项（后端声明，可空）

    var needsTarget: Bool { !targets.isEmpty }
    var hasOptions: Bool { !options.isEmpty }
    /// 按声明顺序分组；只保留真有选项的组。
    var groupedOptions: [(group: OpOptionGroup, items: [OpOption])] {
        optionGroups.compactMap { g in
            let items = options.filter { $0.group == g.id }
            return items.isEmpty ? nil : (g, items)
        }
    }
    var wantsDir: Bool { kind == "dir" }
}
extension DocOp: Decodable {
    private enum CodingKeys: String, CodingKey {
        // ⚠ option_groups 必须写成 optionGroups —— decoder 的 .convertFromSnakeCase
        // 已经把 JSON key 转过了；写 snake_case = 该字段永远解不出来（checkbox 全空）
        case id, verb, title, subtitle, icon, exts, kind, targets, options, optionGroups, danger, aliases
    }
    init(from decoder: Decoder) throws {
        let c = try decoder.container(keyedBy: CodingKeys.self)
        id = c.str(.id); verb = c.str(.verb)
        title = c.str(.title); subtitle = c.str(.subtitle)
        aliases = c.str(.aliases)
        icon = c.str(.icon, "doc"); kind = c.str(.kind, "files")
        exts = (try? c.decodeIfPresent([String].self, forKey: .exts)) ?? nil ?? []
        targets = (try? c.decodeIfPresent([OpTarget].self, forKey: .targets)) ?? nil ?? []
        options = (try? c.decodeIfPresent([OpOption].self, forKey: .options)) ?? nil ?? []
        optionGroups = (try? c.decodeIfPresent([OpOptionGroup].self, forKey: .optionGroups)) ?? nil ?? []
        danger = (try? c.decodeIfPresent(Bool.self, forKey: .danger)) ?? nil ?? false
    }
}

/// `gui-ops` → {"ok": true, "ops": [...]}
struct OpsResult: Decodable {
    let ok: Bool
    let ops: [DocOp]
    private enum CodingKeys: String, CodingKey { case ok, ops }
    init(from decoder: Decoder) throws {
        let c = try decoder.container(keyedBy: CodingKeys.self)
        ok = c.bool(.ok, true)
        ops = (try? c.decodeIfPresent([DocOp].self, forKey: .ops)) ?? nil ?? []
    }
}

// MARK: - gui-run：逐文件结果

/// 单个输入的处理结果。outputs = 产出文件/目录绝对路径。
struct FileResult: Identifiable, Hashable {
    let id = UUID()
    let input: String      // 输入绝对路径（或 merge 的 "a + b"）
    let name: String       // 展示名
    let ok: Bool
    let outputs: [String]  // 产出绝对路径
    let message: String    // 一行人话结果
}
extension FileResult: Decodable {
    private enum CodingKeys: String, CodingKey { case input, name, ok, outputs, message }
    init(from decoder: Decoder) throws {
        let c = try decoder.container(keyedBy: CodingKeys.self)
        input = c.str(.input); name = c.str(.name)
        ok = c.bool(.ok)
        outputs = (try? c.decodeIfPresent([String].self, forKey: .outputs)) ?? nil ?? []
        message = c.str(.message)
    }
}

/// `gui-run` → {"ok", "op", "results":[...], "succeeded", "total", "log", "skipped_missing"?}
struct RunResult: Decodable {
    let ok: Bool
    let op: String
    let results: [FileResult]
    let succeeded: Int
    let total: Int
    let log: String
    let skippedMissing: [String]
    private enum CodingKeys: String, CodingKey {
        case ok, op, results, succeeded, total, log, skippedMissing
    }
    init(from decoder: Decoder) throws {
        let c = try decoder.container(keyedBy: CodingKeys.self)
        ok = c.bool(.ok, true)
        op = c.str(.op)
        results = (try? c.decodeIfPresent([FileResult].self, forKey: .results)) ?? nil ?? []
        succeeded = c.int(.succeeded)
        total = c.int(.total)
        log = c.str(.log)
        skippedMissing = (try? c.decodeIfPresent([String].self, forKey: .skippedMissing)) ?? nil ?? []
    }
}
