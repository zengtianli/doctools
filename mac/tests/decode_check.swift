// 解码契约自检 —— build 前置门。
//
// 存在的理由(2026-07-26 真事故):后端 gui-ops 声明了 option_groups,Swift 侧
// CodingKey 也写了 option_groups,看着完全对应 —— 但 BackendClient 的 decoder 开着
// `.convertFromSnakeCase`,JSON key 到手时已被转成 optionGroups,于是那个 CodingKey
// 永远匹配不上,分组静默解成空数组:面板只渲染出「选项」标题,一个 checkbox 都没有。
// 当时用默认 JSONDecoder 写的自测是绿的 —— **测试没复现真实解码路径,等于没测**。
//
// 所以本文件的铁律:decoder 配置必须与 BackendClient.swift 逐字一致。
import Foundation

let path = CommandLine.arguments.count > 1 ? CommandLine.arguments[1] : "ops.json"
let data = try! Data(contentsOf: URL(fileURLWithPath: path))

let decoder = JSONDecoder()
decoder.keyDecodingStrategy = .convertFromSnakeCase   // ← 必须同 BackendClient.swift
let r = try! decoder.decode(OpsResult.self, from: data)

var bad = 0
guard !r.ops.isEmpty else { print("✗ gui-ops 返回 0 个操作"); exit(1) }
for op in r.ops {
    guard !op.options.isEmpty else { continue }
    let rendered = op.groupedOptions.reduce(0) { $0 + $1.items.count }
    if op.groupedOptions.isEmpty || rendered != op.options.count {
        bad += 1
        print("✗ \(op.id): 声明 \(op.options.count) 项,实际能渲染 \(rendered) 项(分组 \(op.optionGroups.count))")
    } else {
        print("✓ \(op.id): \(op.options.count) 项 / \(op.groupedOptions.count) 组")
    }
}
if bad > 0 { print("✗ \(bad) 个操作的选项渲染不出来 —— 多半是 CodingKey 与 snake_case 策略打架"); exit(1) }
print("✓ 解码契约自检通过(\(r.ops.count) 个操作)")
