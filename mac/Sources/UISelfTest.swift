import AppKit
import SwiftUI

/// Runs inside the built app. No event synthesis, ordered windows or clipboard access.
@MainActor
enum UISelfTest {
    static func run() async {
        var checks: [String: Bool] = [:]
        var captures: [[String: Any]] = []
        let output = URL(fileURLWithPath: ProcessInfo.processInfo.environment["SOP_OUT_DIR"]
                         ?? FileManager.default.temporaryDirectory.appendingPathComponent("dockit-ui-self-test").path)
        do {
            try FileManager.default.createDirectory(at: output, withIntermediateDirectories: true)
            let vm = AppViewModel()
            await vm.loadOps() // Same action as the refresh toolbar/menu.
            checks["backend_refresh"] = !vm.ops.isEmpty && !vm.isLoadingOps && vm.banner == nil
            guard let first = vm.ops.first else { throw Failure("No backend operations") }
            let originalIDs = vm.ops.map(\.id)
            vm.selectedOpID = vm.ops.last?.id
            await vm.loadOps()
            checks["refresh_preserves_selection"] = vm.ops.map(\.id) == originalIDs
                && vm.selectedOpID == vm.ops.last?.id && !vm.isLoadingOps
            vm.addPaths(["/tmp/dockit-ui-fixture.docx", "/tmp/dockit-ui-fixture.docx"])
            checks["file_deduplication"] = vm.files.count == 1
            vm.clearFiles()
            checks["clear_files"] = vm.files.isEmpty && vm.results.isEmpty && !vm.canRun

            let content = ContentView(viewModel: vm, autoLoad: false)
            let items = content.paletteItems
            let palette = PaletteModel(items)
            checks["palette_covers_operations"] = items.map(\.id) == originalIDs
            palette.query = first.title
            checks["palette_search"] = palette.current?.id == first.id
            palette.current?.run()
            vm.onOpChanged()
            checks["palette_selection"] = vm.selectedOpID == first.id
            palette.query = ""
            palette.move(-1)
            checks["palette_navigation_wrap"] = palette.sel == items.count - 1
            palette.move(1)
            checks["palette_navigation_return"] = palette.sel == 0
            palette.query = "nonexistent-acceptance-operation-98127"
            checks["palette_empty_result"] = palette.results.isEmpty && palette.current == nil

            vm.banner = .warning("界面自检：这是一条可关闭的提示。")
            captures.append(try await capture(content, size: NSSize(width: 1000, height: 680),
                                             to: output.appendingPathComponent("native-main.png")))
            vm.dismissBanner() // Same callback as StatusBanner's close button.
            checks["banner_close"] = vm.banner == nil
            captures.append(try await capture(content, size: NSSize(width: 900, height: 640),
                                             to: output.appendingPathComponent("native-compact.png")))
            let panel = CommandPalette(items: items, isPresented: .constant(true)).padding(32)
            captures.append(try await capture(panel, size: NSSize(width: 640, height: 500),
                                             to: output.appendingPathComponent("native-palette.png")))
            guard let privacyOp = vm.ops.first(where: { $0.options.contains { $0.group == "privacy" } }),
                  let consent = privacyOp.options.first(where: { $0.group == "privacy" && !$0.isFile }) else {
                throw Failure("No backend-declared privacy option")
            }
            vm.selectedOpID = privacyOp.id
            vm.onOpChanged()
            checks["privacy_default_off"] = !consent.defaultOn && vm.optionValues[consent.id] == false
                && vm.changedOptions[consent.id] == nil
            vm.optionValues[consent.id] = true
            checks["privacy_opt_in_forwarded"] = vm.changedOptions[consent.id] == "1"
            vm.onOpChanged()
            checks["privacy_selection_resets_consent"] = vm.optionValues[consent.id] == false
                && vm.changedOptions[consent.id] == nil
            vm.optionValues[consent.id] = true
            vm.resetOptionsToDefaults()
            checks["privacy_restore_default_off"] = vm.optionValues[consent.id] == false
                && vm.changedOptions[consent.id] == nil
            captures.append(try await capture(content, size: NSSize(width: 1000, height: 680),
                                             to: output.appendingPathComponent("native-privacy.png")))
            checks["rendered_four_views"] = captures.count == 4
            checks["no_visible_or_key_window"] = NSApp.windows.allSatisfy { !$0.isVisible && !$0.isKeyWindow }
            checks["no_activation"] = !NSApp.isActive && NSApp.activationPolicy() == .prohibited
        } catch {
            checks["self_test_completed"] = false
            fputs("UI self-test: \(error)\n", stderr)
        }
        let ok = !checks.isEmpty && checks.values.allSatisfy { $0 }
        let result: [String: Any] = ["ok": ok, "checks": checks, "captures": captures,
            "summary": "App 内真实 SwiftUI 视图和命令面板离屏渲染；刷新、选择、搜索、清空、提示关闭路径自检",
            "scope": "In-process rendering and state assertions, including backend-declared privacy consent defaults/reset; scan is never executed. No synthetic input, window activation, installed app or document mutation.",
            "rendering_limit": "Hidden non-key windows; key-window focused selection appearance is not covered. PNG is composited on the system window background without altering selection state."]
        do {
            let data = try JSONSerialization.data(withJSONObject: result, options: [.prettyPrinted, .sortedKeys])
            try data.write(to: output.appendingPathComponent("native_ui.detail.json"), options: .atomic)
            print(String(decoding: data, as: UTF8.self))
        } catch { fputs("Cannot write UI detail: \(error)\n", stderr); exit(1) }
        exit(ok ? 0 : 1)
    }

    private struct Failure: Error { let message: String; init(_ message: String) { self.message = message } }

    private static func capture<V: View>(_ view: V, size: NSSize, to path: URL) async throws -> [String: Any] {
        let window = NSWindow(contentRect: NSRect(origin: NSPoint(x: -20000, y: -20000), size: size),
                              styleMask: [.borderless], backing: .buffered, defer: false)
        window.isReleasedWhenClosed = false
        window.appearance = NSAppearance(named: .aqua)
        window.backgroundColor = .windowBackgroundColor
        let host = NSHostingView(rootView: view.frame(width: size.width, height: size.height))
        window.contentView = host
        defer { window.close() }
        try await Task.sleep(for: .milliseconds(150))
        host.layoutSubtreeIfNeeded()
        host.displayIfNeeded()
        guard !window.isVisible, !window.isKeyWindow,
              let rawBitmap = host.bitmapImageRepForCachingDisplay(in: host.bounds) else { throw Failure("Hidden rendering unavailable") }
        host.cacheDisplay(in: host.bounds, to: rawBitmap)
        guard let bitmap = NSBitmapImageRep(bitmapDataPlanes: nil, pixelsWide: rawBitmap.pixelsWide,
                    pixelsHigh: rawBitmap.pixelsHigh, bitsPerSample: 8, samplesPerPixel: 4,
                    hasAlpha: true, isPlanar: false, colorSpaceName: .deviceRGB, bytesPerRow: 0, bitsPerPixel: 32),
              let context = NSGraphicsContext(bitmapImageRep: bitmap) else { throw Failure("Opaque rendering unavailable") }
        let pixels = NSRect(x: 0, y: 0, width: bitmap.pixelsWide, height: bitmap.pixelsHigh)
        guard let background = window.backgroundColor.usingColorSpace(.deviceRGB),
              let rendered = rawBitmap.cgImage else { throw Failure("No rendered image or background color") }
        let canvas = context.cgContext
        canvas.setFillColor(CGColor(red: background.redComponent, green: background.greenComponent,
                                   blue: background.blueComponent, alpha: 1))
        canvas.fill(pixels)
        canvas.setBlendMode(.normal)
        canvas.draw(rendered, in: pixels)
        guard let pixelData = bitmap.bitmapData else { throw Failure("No rendered pixels") }
        let opaque = (0..<bitmap.pixelsHigh).allSatisfy { y in
            (0..<bitmap.pixelsWide).allSatisfy { x in pixelData[y * bitmap.bytesPerRow + x * 4 + 3] == 255 }
        }
        guard opaque else { throw Failure("Transparent pixels remain after background compositing") }
        guard let data = bitmap.representation(using: .png, properties: [:]), data.count > 10000,
              bitmap.pixelsWide >= Int(size.width), bitmap.pixelsHigh >= Int(size.height) else {
            throw Failure("Empty or undersized render: \(path.lastPathComponent)")
        }
        // Reject a blank/solid image, not merely a nonempty PNG container.
        var colors = Set<String>()
        for y in stride(from: 0, to: bitmap.pixelsHigh, by: 7) {
            for x in stride(from: 0, to: bitmap.pixelsWide, by: 7) {
                if let c = bitmap.colorAt(x: x, y: y)?.usingColorSpace(.deviceRGB) {
                    colors.insert("\(Int(c.redComponent * 255)),\(Int(c.greenComponent * 255)),\(Int(c.blueComponent * 255))")
                }
            }
        }
        guard colors.count > 20 else { throw Failure("Blank render: \(path.lastPathComponent)") }
        try data.write(to: path, options: .atomic)
        return ["file": path.lastPathComponent, "width": bitmap.pixelsWide, "height": bitmap.pixelsHigh,
                "png_bytes": data.count, "sampled_colors": colors.count, "opaque": opaque]
    }
}
