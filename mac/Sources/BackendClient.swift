import Foundation

// =============================================================================
// BackendClient — 外部 Python CLI 的薄封装
//
// 核心蒸馏自 ssot-console BackendClient 2026-06-11，升级时 diff 对齐范本
// （~/Apps/mac/ssot-console/Sources/BackendClient.swift）。
//
// Swift 层只是 GUI 壳：不解析 YAML、不写 SQL、不重写业务逻辑。一切真实工作经
// `Foundation.Process` 委托给后端，stdout 按 JSON (Decodable) 解码。
//
// 信封协议：所有 gui 消费的子命令一律 exit 0；
//   成功 {"ok": true, ...}，失败 {"ok": false, "error": "人话"}。
// runDecoding 先解 {ok, error} 探针 → ok==false 即抛 error 文本；否则解目标类型；
// 目标类型解码失败再 fallback 探针（不把 raw JSON 糊给用户）。
//
// 已固化的坑（每条都在范本上真踩过，删任何一条前先想清楚）：
//   · GUI app 只继承 launchd 极简 PATH → env["PATH"] 前插 /opt/homebrew/bin 等，
//     否则 /usr/bin/env uv 找不到（exit 127）。
//   · 子进程输出 >64KB 管道塞死 → stdout/stderr 各自后台队列并发 drain
//     （readDataToEndOfFile），禁在 terminationHandler 里读。
//   · stdin 大输入对称死锁 → 必须 process.run() 之后再写 stdin、写完 close；
//     用可抛的 write(contentsOf:)，子进程早退（broken pipe）只静默失败不崩 app。
//   · 超时 → terminate + 抛人话错误；Task 取消 → terminate + CancellationError。
// =============================================================================

enum BackendError: LocalizedError {
    case launchFailed(String)
    case nonZeroExit(code: Int32, stderr: String, stdout: String)
    case decodeFailed(String)
    case emptyOutput
    case backend(String)            // 信封 {ok:false, error}
    case timeout(TimeInterval)

    var errorDescription: String? {
        switch self {
        case .launchFailed(let m):
            return "无法启动后端进程: \(m)"
        case .nonZeroExit(let c, let e, let out):
            // 操作类子命令失败时打 {"error": "..."} 到 stdout（非零 exit）。
            if let data = out.data(using: .utf8),
               let env = try? JSONDecoder().decode(BackendErrorEnvelope.self, from: data) {
                return "后端错误: \(env.error)"
            }
            if c == 127 {  // env 找不到 uv
                return "后端启动失败 (exit 127): 找不到 uv —— 请确认 uv 在 PATH 中。\(e)"
            }
            let detail = !e.isEmpty ? e : (out.isEmpty ? "(无输出)" : out)
            return "后端退出码 \(c): \(detail)"
        case .decodeFailed(let m):
            return "后端响应解析失败: \(m)"
        case .emptyOutput:
            return "后端没有输出"
        case .backend(let m):
            return m
        case .timeout(let t):
            return "后端超时（\(Int(t)) 秒未完成），已终止该进程。可重试或去终端跑同样命令排查。"
        }
    }
}

/// 引用盒：并发 drain 队列把数据交还 waiter 闭包。
/// `@unchecked Sendable` 在此安全：每个盒只被一个 drain task 写，
/// waiter 在 `group.wait()` 建立 happens-before 之后才读。
private final class DataBox: @unchecked Sendable {
    var data = Data()
}

/// 布尔标志盒（超时 / 取消），同样靠队列序保证可见性。
private final class FlagBox: @unchecked Sendable {
    var on = false
}

actor BackendClient {
    // MARK: - 后端脚本路径（脚本路径是契约的一部分）
    //
    // 脚手架初始指向 app 目录内的 backend_demo.py（开箱即跑）。
    // 正式后端按平台-子公司模型放 ~/Dev/tools/dev/lib/tools/ 下（能力留在总部，
    // 可被 CC/raycast/cron 复用，app 只是视图层），然后改这一个常量。
    // 多后端 = 多 script 常量，runDecoding(args:script:) 路由（参照 ssot-console 三后端）。

    // 正式后端 = doctools 子公司的 GUI 适配器（套 JSON 信封壳，复用 doc_dispatch 路由，
    // 零重写业务）。经 `uv run --project ~/Dev` 跑（脚本只用 stdlib + import 同目录 doc_dispatch）。
    static let defaultScriptPath =
        ("~/Dev/tools/doctools/scripts/document/doc_gui_backend.py" as NSString).expandingTildeInPath

    /// 默认 120s；典型 clean/convert 秒级，但 typeset 套模板 + soffice 升级老 doc 可能数十秒。
    static let defaultTimeout: TimeInterval = 120
    /// 重活（多文件 typeset / soffice 冷启动）用更长超时。
    static let longTimeout: TimeInterval = 600

    let scriptPath: String

    init(scriptPath: String = BackendClient.defaultScriptPath) {
        self.scriptPath = scriptPath
    }

    // MARK: - DocTools API（gui-ops 列操作 / gui-run 跑操作）

    /// 列出可用操作（clean / convert / split / merge / typeset / scan）。
    func ops() async throws -> OpsResult {
        try await runDecoding(args: ["gui-ops"])
    }

    /// 跑一个操作。op = 操作 id；target = 单选槽（convert 目标格式 / renum 范围 /
    /// bidfinal 模式，无则 nil）；options = 勾选项（后端声明的 id → 开关，只传改动过的）；
    /// files = 文件绝对路径数组（scan 传单个目录）。typeset 走 longTimeout。
    // options 的值统一是字符串：bool 型已由 ViewModel 折成 "1"/"0"，file 型是绝对路径。
    // 后端按自己声明的 type 解读 —— Swift 不需要知道哪个 id 是什么类型。
    func run(op: String, target: String?, options: [String: String] = [:],
             files: [String]) async throws -> RunResult {
        var args = ["gui-run", "--op", op]
        if let target, !target.isEmpty { args += ["--to", target] }
        for k in options.keys.sorted() { args += ["--opt", "\(k)=\(options[k]!)"] }
        args.append("--files"); args += files
        let timeout = (op == "typeset") ? Self.longTimeout : Self.defaultTimeout
        return try await runDecoding(args: args, timeout: timeout)
    }

    // MARK: - 实际执行的命令行
    //
    //   /usr/bin/env uv run --project ~/Dev <scriptPath> <args...>
    //
    // `uv` 让后端跑在共享 ~/Dev venv 里（后端只用 stdlib 时也无害）。

    private static let devRoot =
        ("~/Dev" as NSString).expandingTildeInPath

    private func buildArguments(script: String, _ args: [String]) -> [String] {
        ["uv", "run", "--project", BackendClient.devRoot, script] + args
    }

    private func runDecoding<T: Decodable>(args: [String], stdin: String? = nil,
                                           script: String? = nil,
                                           timeout: TimeInterval = BackendClient.defaultTimeout)
    async throws -> T {
        let output = try await runProcess(args: args, stdin: stdin,
                                          script: script ?? scriptPath, timeout: timeout)
        guard let data = output.data(using: .utf8), !data.isEmpty else {
            throw BackendError.emptyOutput
        }
        let decoder = JSONDecoder()
        decoder.keyDecodingStrategy = .convertFromSnakeCase
        // 1) 信封探针：ok == false → 直接抛后端给的人话。
        let probe = try? decoder.decode(BackendProbe.self, from: data)
        if probe?.ok == false {
            throw BackendError.backend(probe?.error ?? "后端返回失败（未给出原因）")
        }
        // 2) 解目标类型。
        do {
            return try decoder.decode(T.self, from: data)
        } catch {
            // 3) fallback 探针：目标解码失败但信封里有 error 文本 → 用人话，不糊 raw JSON。
            if let err = probe?.error, !err.isEmpty {
                throw BackendError.backend(err)
            }
            throw BackendError.decodeFailed(
                "\(error.localizedDescription) — raw: \(output.prefix(300))")
        }
    }

    /// 跑进程、捕获 stdout/stderr/exit。
    ///
    /// · stdout/stderr 在子进程运行期间于后台队列并发 drain（防 64KB 管道死锁）。
    /// · stdin 在 run() 之后写入（reader 已存在），`write(contentsOf:)` 可抛
    ///   —— 子进程早退时不再触发 ObjC 异常崩 app。
    /// · 超时：到点 terminate 进程并抛 .timeout。
    /// · 取消（Task.cancel）：terminate 进程并抛 CancellationError。
    private func runProcess(args: [String], stdin: String?, script: String,
                            timeout: TimeInterval) async throws -> String {
        let process = Process()
        process.executableURL = URL(fileURLWithPath: "/usr/bin/env")
        process.arguments = buildArguments(script: script, args)

        // GUI 启动的 app 只继承 launchd 的最小 PATH，前置常用工具目录让 uv 可解析。
        var env = ProcessInfo.processInfo.environment
        let extraPaths = ["/opt/homebrew/bin", "\(NSHomeDirectory())/.local/bin", "/usr/local/bin"]
        let currentPath = env["PATH"] ?? "/usr/bin:/bin:/usr/sbin:/sbin"
        env["PATH"] = (extraPaths + [currentPath]).joined(separator: ":")
        process.environment = env

        let outPipe = Pipe()
        let errPipe = Pipe()
        process.standardOutput = outPipe
        process.standardError = errPipe
        let inPipe: Pipe? = (stdin != nil) ? Pipe() : nil
        if let inPipe { process.standardInput = inPipe }

        let canceled = FlagBox()
        let timedOut = FlagBox()

        return try await withTaskCancellationHandler {
            try await withCheckedThrowingContinuation { (cont: CheckedContinuation<String, Error>) in
                let ioQueue = DispatchQueue(
                    label: "cyou.tianli.DocTools.backend-io", attributes: .concurrent)
                let group = DispatchGroup()
                let outBox = DataBox()
                let errBox = DataBox()

                do {
                    try process.run()
                } catch {
                    cont.resume(throwing: BackendError.launchFailed(error.localizedDescription))
                    return
                }

                // stdin 在 run() 之后写（reader 已存在），写完 close 让子进程见 EOF。
                // 用可抛的 write(contentsOf:)：子进程早退（broken pipe）只静默失败，不崩 app。
                if let inPipe, let stdin {
                    ioQueue.async {
                        let h = inPipe.fileHandleForWriting
                        try? h.write(contentsOf: Data(stdin.utf8))
                        try? h.close()
                    }
                }

                // 并发 drain：各自线程阻塞读到子进程关管道，管道永远不会被写满。
                group.enter()
                ioQueue.async {
                    outBox.data = outPipe.fileHandleForReading.readDataToEndOfFile()
                    group.leave()
                }
                group.enter()
                ioQueue.async {
                    errBox.data = errPipe.fileHandleForReading.readDataToEndOfFile()
                    group.leave()
                }

                // 超时守卫：到点 terminate（waitUntilExit 随之返回，timedOut 标志驱动抛错）。
                ioQueue.asyncAfter(deadline: .now() + timeout) {
                    if process.isRunning {
                        timedOut.on = true
                        process.terminate()
                    }
                }

                // 等退出 + 双 drain 完成，恰好 resume 一次。
                ioQueue.async {
                    process.waitUntilExit()
                    group.wait()
                    if canceled.on {
                        cont.resume(throwing: CancellationError())
                        return
                    }
                    if timedOut.on {
                        cont.resume(throwing: BackendError.timeout(timeout))
                        return
                    }
                    let out = String(decoding: outBox.data, as: UTF8.self)
                    let err = String(decoding: errBox.data, as: UTF8.self)
                    if process.terminationStatus == 0 {
                        cont.resume(returning: out)
                    } else {
                        cont.resume(throwing: BackendError.nonZeroExit(
                            code: process.terminationStatus, stderr: err, stdout: out))
                    }
                }
            }
        } onCancel: {
            canceled.on = true
            if process.isRunning { process.terminate() }
        }
    }
}
