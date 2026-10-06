# DocKit · doctools 桌面端

共享文档处理能力与业务闸门见 [父项目](../CLAUDE.md)，此目录仅维护 SwiftUI 客户端、图标及本机构建；源码由 doctools 父仓管理。

- `BackendClient.swift` 调用父项目 `scripts/document/doc_gui_backend.py`，经 `uv run --project ~/Dev` 使用共享环境。
- `gui-ops` 声明操作、选项及分组；Swift 动态渲染，不硬编码业务选项。
- `gui-run` 返回 JSON 信封：成功 `ok: true`，业务错误 `ok: false, error: ..., error_code: ...`；保留现有参数与错误语义，`gui-*` 一律 exit 0。
- agent 命令行 `bin/dockit`（POSIX sh，build.sh 签名前复制进 `Contents/Resources/bin/`，`--install` 链 `~/.local/bin/dockit`）exec 同一个 `doc_gui_backend.py` 的 `ops` / `run` / `status` / `doctor` / `settings`；后端改动不重建 App 即生效，改 `bin/dockit` 本身才需重装。`status` 不带参数是读回命令（已装版本、界面记住的设置、后台任务）；`settings` 经 `/usr/bin/defaults` 读写 App 偏好域里的 `dockit.lastOperation`、`dockit.targetFormats`，键名与 `ViewModel.portablePreferenceKeys` 必须一致（`test_dockit_cli.py` 守着）。界面功能对照登记在 `project.yaml` 的 `sop.agent_cli`，其中 human 项的名字与 `dockit --help`「仅在窗口中」一节逐字对应：界面加减功能时两处一起改。GUI 与 CLI 的差异只允许在出口：退出码、view 默认不开浏览器、scan 带结构化 `findings`、破坏性操作与覆盖已有同名文件要 `--yes`（GUI 是红标确认与「允许覆盖已有同名文件」勾选项）、`run --background` + `status` 给长任务查询。
- 写操作按源文件所在目录取分层非阻塞 flock（本目录 EX、每级上级 SH；`~/Library/Caches/cyou.tianli.DocTools/locks`，`DOCKIT_CACHE_DIR` 可改到沙盒），同一目录树并发返回 `busy`，防止快照差分把产出记到别的运行头上；锁文件释放时确认无人持有即删除。后台任务记录在同一缓存的 `jobs/`，保留 7 天。
- JSON 使用 snake_case，真实 decoder 为 `.convertFromSnakeCase`，Models 的 CodingKeys 使用 camelCase。
- 修改模型或后端协议必须通过真实 `gui-ops` 解码检查，测试 decoder 与 BackendClient 保持一致。
- bundle ID `cyou.tianli.DocTools`、Xcode project/scheme `DocTools` 保持；显示名以 catalog 为准。
- `./build.sh --check` 运行 CodingKeys 与真实后端解码门；`./build.sh` 本机 Release 构建并签名，不装机。
- 只有显式 `./build.sh --install` 才替换 `/Applications/DocKit.app`，旧安装进入 Trash；工具链复用共享 Xcode 选择器。
- 当前真身为本目录，后续只维护父仓；旧的 `~/Apps/mac/doc-tools` 兼容入口已退役。
