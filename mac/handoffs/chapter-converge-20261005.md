# DocKit 私有 Mac 组件 · 2026-10-05

本组件已在原 Chapter 固定流程真实完成；DocKit 产品线还要合并公开版结果，不能据本组件宣称整行通过。

- 仅接总部 `LaneSignal.swift`、实际 `loadOps()` 与从不上屏的 NSHostingView 离屏绘制就绪信号。默认交互与 doctools 业务未改；静默实例不使用本人偏好，不处理输入文档。源码23ee77c，总部 lifecycle 同步491e0d1。
- 第一次原 build.sh 的 vendor 改变输入，来源闸拒写 receipt，旧回执保留。同步固定来源后原增量构建/装机成功：1.0.1(323)，可执行文件21e6207d…、构建输入68860382…、实际verify=true。receipt 输入仍为原清单，无新平行构建实现。
- 首次原 sim_lane 静默检查0.723s、os_log、UIElement、window=false、focus=false、front_unchanged=true；证明后才登记 in_use:true。
- 正式test9348dc7743ea49c0af00a34045479413与当前code_key一致。固定验收 functionality bda78c6ee41849869ca8ae82a7efbf01、recovery54768551889242f08d73a55ec1f1b76f、privacy9d8ed4d2561b422ba743d58e875d45a4、native_ui2e3a8962906f4e189af2ae2978403734、cli_entry2c75b26dc98f449fb3095f1f9ff4d1ce全部通过。既有图标有效证据复用。
- perf e7f3e4d435e64575845f57ae1bc9b639实际通过：五次os_log首屏就绪[579,509,491,548,523]ms、中位数523ms；静置45秒后采60秒，50MiB、CPU0%；安装2,818,048 bytes。采样全程前台未变、无窗口与抢焦点。测量输入d45ce87b…，原件`~/Library/Caches/app-lightweight/measurement-candidates/323f847dccee476c9b96b446947cf08e.raw.json`、SHA a19f21b5…。
- 最终 `sop run --check-only --retry`：current_passed、delivery_state=complete、coverage=[]、gaps=[]。未推送、发布、重录或重做推广。

## 2026-10-07 追加：界面功能对照与命令补齐（agent_cli）

约定见 app 技能 `references/agent-cli.md`，缺口清单见 Chapter 的 `docs/PRD-agent-cli.md`。

做了什么：

- 从 `Sources/` 五个界面文件逐项列出 27 项功能，登记在 `project.yaml` 的 `sop.agent_cli`：命令 13、只在窗口 9、暂缺 5。读回命令是 `dockit status`。
- `dockit status` 不带参数时除后台任务外，另报已装 `DocKit.app` 的版本与构建号、界面记住的设置。原有的 `jobs` 字段没动，只多了 `app` 和 `settings`。
- 新增 `dockit settings`（只读）和 `dockit settings set <键> <值>`。读写的就是界面那两个偏好键（`dockit.lastOperation`、`dockit.targetFormats`），经 `/usr/bin/defaults` 走同一个偏好域。值先按操作目录校验，未知操作、目标或设置项以 2 退出且不写入。偏好域里的窗口位置、最近目录等内容，读命令不输出，写命令不改。
- `dockit --help` 补齐四样：读命令与写命令分节、每条命令的 `--json` 输出形状、退出码表、「仅在窗口中」清单。登记里每个 human 项的名字都在这一节里逐字出现，公开版登记的 10 个 human 项也在。
- README 中英文、本目录与父仓 `CLAUDE.md`、`catalog.yaml` 同步了新命令。

对照结果：

- 命令 13 项：操作列表与说明、刷新、选择文件或目录、目标格式、选项勾选、选项里选择参考文件、执行、执行中的等待、结果与错误原因、后端日志、状态行、记住上次操作与目标格式、版本与构建号。
- 只在窗口 9 项：搜索功能面板、拖入文件或目录、逐个移除待处理文件、清空待处理文件、清除已选参考文件、恢复默认选项、在 Finder 显示产出、关闭提示条、打开「配置与更新…」窗口。
- 暂缺 5 项：使用 iCloud 记住配置、导出配置、导入配置、检查更新、升级到新版。原因都是共享生命周期模块暂无命令入口，由管该模块的单元统一做，本组件没有各自实现。

怎么验的：

- `python3 -m pytest -q scripts/document/tests`：327 通过、1 跳过。其中 `test_dockit_cli.py` 50 项，新增 5 个用例，另给帮助用例加了 3 组参数。新用例把偏好指到临时 plist、把已装版本指到临时包，不碰本人偏好和已装 App；其中一条钉住登记里的 human 项必须出现在帮助的「仅在窗口中」。共享 `~/Dev/.venv` 没有 pytest，沿用系统 `python3`。
- `./build.sh --check` 通过（14 个操作解码、偏好隔离检查、包装脚本帮助）。
- 已装的 `dockit` 实跑：`dockit status --json` 退出 0，读到 1.0.1 (323)；`dockit settings --no-such-flag --json` 退出 2 并带 `error`。实跑前后导出本人偏好域逐字节比对，没有变化；缓存目录文件清单没有变化。
- 用临时目录里的样例 Markdown 实跑 `dockit run convert --to word`：先 `--dry-run`，再执行得到 docx，重跑按覆盖门以 `would_overwrite` 拒绝。
- `chapter sop accept --app doc-tools-doctools --check agent_cli`：登记与帮助没有问题（problems 为空），因暂缺 5 项判未通过。

没做的：

- 没有重新构建和装机。`bin/dockit` 没改，它直接执行工作树里的后端，所以新命令对已装的 `dockit` 立即生效；App 包本身没有变化。
- 没有对本人真实偏好运行 `settings set`，写入只在临时 plist 上验过。界面在运行中不会即时反映命令写入的值，要到下次启动才按新值选中，这一点没有开窗口验证。
- 五项共享生命周期功能的命令没做。
- `scripts/accept/` 的功能、恢复、隐私、原生界面四项固定验收没有重跑；后端改动后它们的证据需要 Chapter 按新输入重跑。
- Chapter 这次写的 `perf/acceptance/agent_cli.*` 与 `perf/delivery-evidence.json` 没有提交。

## 2026-10-07 追加（二）：核验意见、装机判断与离屏核对

上一节的结果经独立核验，这一轮逐条处理，并判断要不要装机。

没有装机，已装的 `/Applications/DocKit.app` 一个字节没动：

- 装着的是私有 Homebrew 发行的公证版 1.0.1 (323)：`spctl -a -vv` 报 `Notarized Developer ID`，`codesign -dv` 见 Developer ID 与 `Notarization Ticket=stapled`，记录在 `build/notarized/private-brew-20261006/`。
- `build.sh` 只做临时签名，`--install` 会把这份公证版移进废纸篓、换成临时签名的包，签名等级变低。本轮不许公证和发行，做不到同等级，所以不装。
- 也不需要装：323 之后进包的输入只有 `catalog.yaml` 的一行说明变了，`Sources/` 和 `bin/dockit` 没变，包内包装脚本与源文件逐字节相同。它执行工作树里的后端，上一节的 `status`、`settings` 对已装的 `dockit` 早已生效。
- 没有另做一份构建留在 `build/`：构建会先跑生命周期 vendor，而共享模块的 `AppLifecycleUI.swift` 已比本仓副本新（`vendor-lifecycle.py --check` 报 drift），构建会把别的单元还在改的原版拷进 `Sources/`。

核验意见的处理：

- 对照补了一行「侧栏点选操作」→ `dockit run`（操作就是 `run` 的 `<op>` 参数）。现在 28 项：命令 14、只在窗口 9、暂缺 5。
- 失败输出仍是平铺的 `{"ok":false,"error":…,"error_code":…}`，没有改成嵌套的 error 对象：这是 dockit 既有的稳定形状，帮助里写明了，界面的 Swift 解码也读它，约定允许保持。
- 五项共享生命周期功能仍记暂缺，原因不变。共享模块源目录里已有命令层 `AppLifecycleCLI.swift`，本组件还没接；接入要改 Swift 并重新发行。

新增的离屏核对（`./build.sh --check` 里多一步，`tests/SettingsFollowCheck.swift`）：

- `dockit settings set` 写进临时偏好文件，真实 `AppViewModel` 新开一次就按它选中；界面改了目标和操作，`dockit settings` 读到新值；命令改回去，下一次打开的界面跟着变；不认识的值以 2 退出且偏好不变。操作与目标取自真实 `gui-ops`，不写死 id。
- 每次「打开界面」另起一个进程。试过放在同一进程里：`UserDefaults` 读过一次就看不到别的进程后来写的值。这也说明窗口开着时命令改的设置不会即时反映，要到下次启动才选中。
- `tests/PreferencesCheck.swift` 原来每跑一次在 `~/Library/Preferences` 留一份空的 `dockit-fixture-<UUID>.plist`，本机已有 10 份。改成把偏好域放在临时目录的文件上，用完删目录；改后跑了三次 `--check`，没有再多。旧的 10 份没有删，留给本人决定。

怎么验的：

- `./build.sh --check` 通过，含新的一步。`python3 -m pytest -q scripts/document/tests`：327 通过、1 跳过。`script_graph.py` 退出 0，`git diff --check` 干净。
- 只用已装的 `dockit` 做完一项界面功能并读回：scratchpad 里的样例 Markdown，`run convert --to word --dry-run`，再 `--background` 执行，`status <job>` 从 running 读到 done、`all_ok` 为真、产出 docx 在盘上；重跑按覆盖门以 2 退出（`would_overwrite`）。锁和任务记录指到了 scratchpad。
- 错误参数 7 种都以 2 退出，`--json` 输出带 `error` 与 `error_code`：未知选项（status、settings、doctor）、未知操作、未知目标、未知设置项、未知任务。
- `chapter sop accept --app doc-tools-doctools --check agent_cli` 再跑，`chapter agent-cli --json` 读到：28 项、命令 14、只在窗口 9、暂缺 5，problems 为空，状态仍是暂缺。
- 前后各导出一次 `cyou.tianli.DocTools` 与 `io.github.zengtianli.DocTools` 偏好，逐字节相同；缓存目录清单、两个已装 App 与 `~/.local/bin/dockit` 链接的修改时间都没变；没有留下进程，没有监听端口。

没做的：

- 没有装机，也没有新构建（原因见上）。下次随正常发行流程重新签名公证时再带上。
- 窗口开着时不会即时跟随命令写的设置。要改得在 Swift 里加跨进程通知，改完要重新发行，这轮没动。
- 五项共享生命周期功能的命令没接。
- `build.sh`、`tests/*.swift`、`catalog.yaml` 都在构建回执的输入里，回执已与当前源码不一致；已装可执行文件在公证时重新签名，哈希也与回执里的不同。要等下一次发行构建刷新。
- `scripts/accept/` 四项固定验收没有重跑。
- Chapter 写的 `perf/acceptance/agent_cli.*` 与 `perf/delivery-evidence.json` 没有提交。
