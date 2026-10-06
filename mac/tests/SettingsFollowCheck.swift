// 命令与界面共用设置的离屏核对 —— build 前置门(./build.sh --check)。
//
// dockit settings 写进一份临时偏好文件,真实 AppViewModel 按它选中;界面改了选择,dockit settings 读得到;
// 命令再改回去,下一次打开的界面跟着变。全程不开窗口,不碰本人的 DocKit 偏好:
// 偏好域是 <临时目录>/prefs.plist(DOCKIT_DEFAULTS_DOMAIN),操作与目标取自真实 gui-ops,不写死 id。
//
// 用法: settings_follow_check <dockit> <ops.json> <临时目录>
//      (内部另以 --window <偏好域> <ops.json> [<目标> <操作>] 自起一个进程,充当界面的一次启动)
import Foundation

@main
enum SettingsFollowCheck {
    static var dockit = ""
    static var domain = ""

    static func settings(_ args: String...) -> (code: Int32, payload: [String: Any]) {
        let process = Process()
        process.executableURL = URL(fileURLWithPath: dockit)
        process.arguments = ["settings"] + args + ["--json"]
        var environment = ProcessInfo.processInfo.environment
        environment["DOCKIT_DEFAULTS_DOMAIN"] = domain + ".plist"
        process.environment = environment
        let output = Pipe()
        process.standardOutput = output
        process.standardError = FileHandle.nullDevice
        process.standardInput = FileHandle.nullDevice
        try! process.run()
        let data = output.fileHandleForReading.readDataToEndOfFile()
        process.waitUntilExit()
        return (process.terminationStatus, (try? JSONSerialization.jsonObject(with: data)) as? [String: Any] ?? [:])
    }

    static func remembered() -> (operation: String?, targets: [String: String]) {
        let read = settings()
        if read.code != 0 || read.payload["ok"] as? Bool != true { fail("dockit settings 读回失败(退出 \(read.code))") }
        let body = read.payload["settings"] as? [String: Any] ?? [:]
        return (body["last_operation"] as? String, body["target_formats"] as? [String: String] ?? [:])
    }

    /// 界面的一次启动:另起一个进程,新开 AppViewModel 按偏好选中(loadOps() 读偏好走的就是
    /// reloadPortablePreferences())。给了 <目标> <操作> 时,再照界面的做法改选择:分段选目标,侧栏点操作。
    /// 必须是独立进程:同一进程里 UserDefaults 读过一次就不再看见别的进程后来写的值。
    @MainActor static func windowProcess(_ args: [String]) throws -> Never {
        let preferences = UserDefaults(suiteName: args[0])!
        let vm = AppViewModel(preferences: preferences)
        vm.ops = try loadOps(args[1])
        vm.selectedOpID = vm.ops.first?.id
        vm.reloadPortablePreferences()
        print("\(vm.selectedOpID ?? "-") \(vm.selectedTargetID ?? "-")")
        if args.count == 4 {
            vm.selectedTargetID = args[2]
            vm.selectedOpID = args[3]
            vm.onOpChanged()
            preferences.synchronize()
        }
        exit(0)
    }

    static func window(_ opsPath: String, change: [String] = []) -> (operation: String, target: String) {
        let process = Process()
        process.executableURL = URL(fileURLWithPath: CommandLine.arguments[0])
        process.arguments = ["--window", domain, opsPath] + change
        let output = Pipe()
        process.standardOutput = output
        process.standardInput = FileHandle.nullDevice
        try! process.run()
        let data = output.fileHandleForReading.readDataToEndOfFile()
        process.waitUntilExit()
        let words = String(decoding: data, as: UTF8.self).split(separator: " ").map { String($0).trimmingCharacters(in: .newlines) }
        precondition(process.terminationStatus == 0 && words.count == 2, "界面进程没有正常结束")
        return (words[0], words[1])
    }

    static func loadOps(_ path: String) throws -> [DocOp] {
        let decoder = JSONDecoder()
        decoder.keyDecodingStrategy = .convertFromSnakeCase   // ← 必须同 BackendClient.swift
        return try decoder.decode(OpsResult.self, from: Data(contentsOf: URL(fileURLWithPath: path))).ops
    }

    static func fail(_ message: String) -> Never {
        print("✗ " + message)
        exit(1)
    }

    @MainActor static func main() throws {
        let args = CommandLine.arguments
        if args.count >= 4, args[1] == "--window" { try windowProcess(Array(args.dropFirst(2))) }
        guard args.count == 4 else {
            FileHandle.standardError.write(Data("usage: settings_follow_check <dockit> <ops.json> <dir>\n".utf8))
            exit(2)
        }
        dockit = args[1]
        domain = URL(fileURLWithPath: args[3]).appendingPathComponent("prefs").path
        try FileManager.default.createDirectory(atPath: args[3], withIntermediateDirectories: true)
        let ops = try loadOps(args[2])
        guard let first = ops.first,
              let op = ops.last(where: { $0.id != first.id && $0.needsTarget && $0.targets.count >= 2 }) else {
            fail("操作目录里没有可用来核对的操作(要一个非首位、至少两个目标的操作)")
        }
        let (initial, chosen) = (op.targets[0].id, op.targets[1].id)

        // 1. 空偏好:界面停在默认的第一个操作,命令读到未记录
        var opened = window(args[2])
        var now = remembered()
        if opened.operation != first.id || now.operation != nil { fail("空偏好下界面选中 \(opened.operation),命令读到 \(now.operation ?? "未记录")") }

        // 2. 命令改设置 → 界面跟着变;同一次打开里界面再改选择(分段选目标,侧栏点回第一个操作)
        if settings("set", "last_operation", op.id).code != 0 || settings("set", "target_formats.\(op.id)", chosen).code != 0 {
            fail("dockit settings set 没有成功")
        }
        opened = window(args[2], change: [initial, first.id])
        if opened != (op.id, chosen) { fail("命令写了 \(op.id)/\(chosen),界面选中的是 \(opened.operation)/\(opened.target)") }

        // 3. 界面改的 → 命令读到新值
        now = remembered()
        if now.operation != first.id || now.targets[op.id] != initial {
            fail("界面改成 \(first.id) 与 \(op.id)=\(initial),命令读到 \(now.operation ?? "未记录") 与 \(now.targets)")
        }

        // 4. 命令改回去 → 下一次打开的界面跟着变
        if settings("set", "last_operation", op.id).code != 0 { fail("dockit settings set 没有成功") }
        opened = window(args[2])
        if opened != (op.id, initial) { fail("命令改回 \(op.id),界面选中的是 \(opened.operation)/\(opened.target)") }

        // 5. 不认识的值:退出码 2、带 error,偏好不变
        let rejected = settings("set", "last_operation", "no-such-operation")
        now = remembered()
        if rejected.code != 2 || rejected.payload["ok"] as? Bool != false || rejected.payload["error"] == nil
            || now.operation != op.id || now.targets[op.id] != initial {
            fail("不认识的操作应以 2 退出且不改偏好,实际退出 \(rejected.code),偏好 \(now.operation ?? "未记录")")
        }

        print("PASS command-written settings reach the view model and view-model changes reach dockit settings (\(op.id): \(chosen) → \(initial))")
    }
}
