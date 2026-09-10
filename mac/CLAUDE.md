# DocKit · doctools 桌面端

共享文档处理能力与业务闸门见 [父项目](../CLAUDE.md)，此目录仅维护 SwiftUI 客户端、图标及本机构建；源码由 doctools 父仓管理。

- `BackendClient.swift` 调用父项目 `scripts/document/doc_gui_backend.py`，经 `uv run --project ~/Dev` 使用共享环境。
- `gui-ops` 声明操作、选项及分组；Swift 动态渲染，不硬编码业务选项。
- `gui-run` 返回 JSON 信封：成功 `ok: true`，业务错误 `ok: false, error: ...`；保留现有参数与错误语义。
- JSON 使用 snake_case，真实 decoder 为 `.convertFromSnakeCase`，Models 的 CodingKeys 使用 camelCase。
- 修改模型或后端协议必须通过真实 `gui-ops` 解码检查，测试 decoder 与 BackendClient 保持一致。
- bundle ID `cyou.tianli.DocTools`、Xcode project/scheme `DocTools` 保持；显示名以 catalog 为准。
- `./build.sh --check` 运行 CodingKeys 与真实后端解码门；`./build.sh` 本机 Release 构建并签名，不装机。
- 只有显式 `./build.sh --install` 才替换 `/Applications/DocKit.app`，旧安装进入 Trash；工具链复用共享 Xcode 选择器。
- 旧 `~/Apps/mac/doc-tools` 暂为并发会话兼容链接；当前真身为本目录，后续只维护父仓。
