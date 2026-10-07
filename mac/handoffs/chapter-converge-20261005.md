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

## 2026-10-07 追加（三）：接「配置与更新」命令层，公证后装机

本轮授权是本人一句「好，继续做完，铺开」。边界由协调会话裁定，不是本人逐项同意的。

边界：

- 派活原文写「~/Dev/tools/doctools 是引擎，不归你改」，与 Chapter 登记（本组件目录就是 `~/Dev/tools/doctools/mac`）对不上。问过协调会话，裁定可以改三处：`mac/**`、`scripts/document/doc_gui_backend.py`、`scripts/document/tests/test_dockit_cli.py`。
- `doc_gui_backend.py` 只加了一段：参数解析之前的转调函数、对应的帮助行、`main()` 开头两处判断。文档处理、`ops`／`run`／`status`／`doctor`／`settings` 的行为、字段和退出码没有改。
- `project.yaml` 只动了 `sop.agent_cli`。`perf/` 下 Chapter 写的验收证据没动、没提交。父仓 `CLAUDE.md` 没动（见「没做的」）。
- 同一单元还有另一个实例在做公开仓 `~/Apps/oss/doc-tools-oss`，两边用同一个声明身份，`claims.py` 互相拦不住；靠协调会话分工：它不碰 doctools，我不碰公开仓。

做了什么：

- App 侧：`Sources/AppLifecycleCLI.swift` 是总部 `swift-shared` 的逐字节副本（sha256 `c869edf8…`）。`Sources/ProductLifecycle.swift` 的 `Lifecycle` 是唯一工厂，窗口和命令共用。`DocToolsMain.main()` 在第一个参数是 `config` 或 `update` 时只跑共用命令层并退出，不创建 NSApplication。`installApp` 里加了 `AppLifecycleCLI.follow`，运行中的窗口跟随命令的改动。
- 命令行侧：`dockit config …`、`dockit update …` 整段转给 App 可执行文件，输出和退出码原样带回。App 取法：`DOCKIT_APP_BUNDLE`（只给测试）→ `bin/dockit` 导出的入口所在包 → `/Applications/DocKit.app`。
- 防开窗的门：只有包的 Info.plist 里 `DocKitCommandVerbs` 列了这个词才转调，否则报 `app_outdated`、不启动。原因：不认这些词的旧版会把它当普通启动，把窗口打开。后端是工作树里的活文件，这道门先于转调接线写好。
- `dockit public <词…>`：把词原样交给公开包的可执行文件（默认 `/Applications/DocKit Public.app`，bundle id 须是 `io.github.zengtianli.DocTools`），环境带 `DOCKIT_CLI_NAME="dockit public"`，同一道门。公开包那一侧由另一个实例做。
- 帮助：顶层帮助的读、写两节加了共用层那几行，退出码和 `--json` 形状各补了说明，「暂无命令」改成升级到新版和同步状态那句话。`dockit config --help` 在没装 App 时由后端给出同一份文字。
- 登记：使用 iCloud 记住配置 → `dockit config sync`，导出配置 → `dockit config export`，导入配置 → `dockit config import`，检查更新 → `dockit update check`。新增一行「iCloud 配置同步状态」记暂缺，「升级到新版」仍记暂缺，原因都写在登记里。
- README 中英文、本目录 `CLAUDE.md` 同步了这组命令。

怎么验的：

- `python3 -m pytest -q scripts/document/tests`：333 通过、1 跳过（改动前 327）。新增 6 条都用 sh 桩当 App 程序。
- `./build.sh --check` 通过。它现在多一步：用 swiftc 编出 App 程序跑 `tests/test_lifecycle_cli.py`（9 条，约 45 秒）。
- `tests/test_lifecycle_cli.py` 走真实进程：sh 薄壳 → 后端 → App 程序。程序再起一份充当运行中的 App（激活策略 prohibited，生产的 `installApp`，真实 `ContentView` 和共用窗口都建出但不显示）。内容：拨三轮开关、on/off 背靠背三次、on/off/on 一次、导入后主视图换到导入的操作、同步开着时连续三次导入加一对背靠背导入。每次判定都另起进程读存下来的值。屏幕上窗口数全程为 0。
- 分辨力对照，两种坏写法各编一份跑同一套用例，都被抓到：
  - 共用层换成总部提交 4faeaca 那版：败在「on, off back to back stays off (1)」，隔开拨的三轮照过。
  - 产品的 `onChange` 把启动时手里的旧设置存回去：两条导入用例都失败。
- 公证版上的整套用例：签名把程序和 Info.plist 绑在一起，拷进测试包会被系统直接终止（实测退出 -9）。所以加了 `DOCKIT_APP_IN_PLACE`，就地跑那个包本身，偏好仍用一次性域。在随后装机的那份公证包上 9 条全过，跑完本人两个偏好域逐字节未变。
- 装机版实测：`dockit config status --json` 退出 0（同步关、可迁移项 0 个、App 未运行）；`dockit config status --no-such --json` 退出 2、`error.code` 为 `usage`；`dockit config sync on --json` 不带 `--yes` 退出 2（`confirmation_required`）；`dockit config sync on --dry-run --json` 报 `would_change: true`，没有改动。
- `chapter sop accept --app doc-tools-doctools --check agent_cli` 后读回：29 项，命令 18、只在窗口 9、暂缺 2，problems 为空，状态仍是暂缺。
- `scripts/accept/native_ui.sh` 直接跑了一次作回归，通过；没有经 Chapter 写证据。

装机：

- 装前 `/Applications/DocKit.app` 是 1.0.1 (323)，`spctl` 报 Notarized Developer ID。
- 新包 1.0.1 (329)，源码提交 8a15b10。`build.sh` 构建后另存到 `build/notarized/agentcli-20261007/DocKit.app`，用同一身份（Developer ID Application: tianli Zeng，B9LJH93LA4）加 hardened runtime 和时间戳签名，公证 Accepted（id `c3cf17c3-508f-47ee-908e-9035c0c203b8`），已装订。`spctl -a -vv` 报 Notarized Developer ID，`stapler validate` 通过。
- 10-07 13:33:09 装机：旧包移到 `~/.Trash/app-rebuild-20261007-133309-44495/`，新包用 ditto 拷入，`~/.local/bin/dockit` 重新链接。这是 `build.sh --install` 的那几步，只是拷的是公证包，因为 `--install` 自己装的是临时签名的包。装后逐文件与公证包相同。
- 装前装后：两个偏好域导出逐字节相同，数据文件清单（iCloud 里的三份发行文件）和缓存目录相同。原有命令的输出只有 `status` 的 `app.build` 由 323 变 329；帮助只少了旧的「暂无命令」那一行，其余是新增。
- 证据和装机记录在 `build/notarized/agentcli-20261007/`（`install-record.json`、`evidence/`），该目录不入库。

没做的、没验证的：

- 没有在真实窗口开着时跑过。探针的窗口从未显示，本人开着 DocKit 时拨开关、导入会怎样只有离屏证据。
- 没有对本人的偏好域和 iCloud Drive 跑 `config sync on --yes`、`config import`、`update check`。生产里命令和窗口都读写 `.standard`，这条跨进程路径没有实测；测试里两边用的是同一个具名测试域。
- 「配置与更新」窗口没有换版。`Sources/AppLifecycleUI.swift` 仍是旧副本，构建时把 `APP_LIFECYCLE_VENDOR` 指到不存在的文件跳过了分发。以后不带这个变量跑 `./build.sh`，四份共用文件会一起刷新成总部现版，窗口随之变化；要不要换由本人定。
- 私有 Homebrew 的 cask、发行资产和 iCloud 里的私有更新记录都没动，仍是 1.0.1-323。构建回执没有刷新，没有做性能测量。
- `scripts/accept/` 的功能、恢复、隐私三项没有重跑；Chapter 里这四项的证据输入已变。
- 60 秒超时那条分支没有用真的挂起去测。
- `dockit public`：写这一节时现装的公开版是 1.1.3 (39)，没有那个键，实测报 `app_outdated`、没有被启动。10-07 13:31 公开仓那一半装上了 1.1.3 (46)，`Info.plist` 带 `DocKitCommandVerbs=[status, settings, config, update, help]`，之后的联测（协调会话 10-07 更正，原句只写到桩测）：`dockit public status --json` 退出 0，`dockit public status --no-such --json` 退出 2，`dockit public config status --json` 退出 0；独立核验者另用进程审计确认词表之外的首词（`nosuchword`、`run`、`ops`、`--version`）都报 `app_outdated` 且一次进程都没起。
- 装机版窗口的证据（协调会话 10-07 补）：对 `/Applications/DocKit.app` 1.0.1 (329) 就地跑进程内离屏自检 `--ui-self-test`，18 项通过、4 张离屏图，偏好域前后逐字节相同；结果在 `build/notarized/agentcli-20261007/evidence/ui-self-test-installed/`。真实窗口仍没有人打开过。装前装后的偏好与数据留底也补进了同目录的 `before/`、`after/`。
- 父仓 `CLAUDE.md` 的独立入口表里 `mac/bin/dockit` 那一行还没写 `config`／`update`／`public`，不在本轮边界内。
- `dockit settings set` 改的值，开着的窗口仍要到下次启动才选中，这轮没动。

## 2026-10-07 追加（四）：第二轮 — `update install` 与同步状态那句话（接续被中断的单元）

本轮授权是本人一句「继续全部做完。按照你的意思」（针对主线的五条意见）。本节由接续者写：前一个执行者 17:46 开工、18:05 前后随主会话重启被结束，没有留下回报；它的改动都在工作区，没有提交、没有构建、没有装机。

进度（做一步写一步，未写「完成」的都还没做）：

- 18:20 声明边界（`claims.py --pid 68260 --session agentcli2-dockit`，两仓共 22 个具体文件）。此前三次被拒：前一个执行者用旧主进程号留下的同名租约没放，由主线处理后拿到。
- 18:21 重新留底：`build/notarized/agentcli2-20261007/evidence/before-resume/`（装机版 1.0.1 (329)、Notarized Developer ID、两个偏好域导出、iCloud 里三份发行文件的哈希、包内文件哈希、只读命令输出；两个 App 都没在运行）。前一个执行者 17:49 的留底在同目录 `before/`。
- 四份共用副本（`AppLifecycle`／`AppConfiguration`／`AppLifecycleUI`／`AppLifecycleCLI`）与总部现版逐字节相同，不用再刷新。仓里没有 `.sync-conflict-*` 文件，没有东西要移。
- 前一个执行者留下的改动（后端转调与帮助、两份测试、登记两项改 `command`、README 中英文、`CLAUDE.md`）逐处读过，沿用；测试结果见下。
- 18:30 `./build.sh --check` 通过（经 2 号槽排队跑）：刷新后的四份共用副本配上原有接线编得过；`tests/test_lifecycle_cli.py` 12 条里 11 过、1 条按设计跳过（要给组装好的包才跑）。新用例含在临时目录里把一个测试包真的换成发行记录里的新版（旧包进隔离的废纸篓目录、不重开、记住的设置不变）。
- 18:32 `python3 -m pytest -q scripts/document/tests`：334 过、1 跳过。前一个执行者 17:52 改动前的留底是 331 过、2 败、1 跳过：`test_lifecycle_help_lines_are_listed_from_one_place` 是 16:57 副本刷新后帮助行没跟上（已随本轮改好），`test_sigterm_stops_engine_group_and_releases_locks` 这次通过（当时机器负载很高，是计时类偶发；引擎侧，本轮没碰）。
- 18:33 源码、测试、帮助、登记本地提交 `f27b692`（7 个文件，限定路径，未推送）。构建号取提交数，这一版是 334。
- 18:38 `APP_LIFECYCLE_VENDOR=/nonexistent/skip-vendor ./build.sh` 从提交 `f27b692` 构建通过（含检查段 12 条用例与组装后的只读用例）。跳过分发是为了让包里的四份共用副本就是测过的那一版；构建前后四份都与总部现版逐字节相同。
- 18:45 新包另存到 `build/notarized/agentcli2-20261007/DocKit.app`：1.0.1 (334)，同一身份（Developer ID Application: tianli Zeng，B9LJH93LA4）加 hardened runtime 与时间戳签名，公证 Accepted（id `31fe51a9-d9b2-4ebf-836a-8b2b314c3f97`），已装订；`spctl -a -vv` 报 Notarized Developer ID，`stapler validate` 通过。可执行文件 sha256 `eb29baa0b74ed2d0…`。只做了提交与装订，没有碰任何发行渠道。
- 18:51 装机前，在这份公证包上就地跑整套隔离用例（`DOCKIT_APP_IN_PLACE`，经 2 号槽）：11 条里 10 过、1 条按设计跳过（真实替换那条只在临时包上做，不在就地的包上做）。跑完包的程序与 Info.plist 哈希没变，本人偏好域导出的哈希没变，`spctl` 仍是 Notarized Developer ID。
- 18:51:46 装机（DocKit 当时没在运行，没有重启任何东西）：旧包 1.0.1 (329) 移到 `~/.Trash/dockit-1.0.1-329-20261007-185146/DocKit.app`，公证包用 ditto 拷入 `/Applications/DocKit.app`，逐文件与公证包相同；`~/.local/bin/dockit` 原本就指向包内入口，没有动。装后 `spctl -a -vv` 报 Notarized Developer ID，`stapler validate` 通过，可执行文件 sha256 `eb29baa0b74ed2d0…`（装前 `b68f1ce5dfe5401b…`）。
- 装后比对（`evidence/after-self-334/` 对 `before-resume/`）：两个偏好域导出、偏好文件大小与修改时间、iCloud 里三份发行文件、缓存目录、公开版整包文件哈希都逐字节相同；16 条只读命令的退出码全部相同。输出有差别的只有预期的四处：`status` 的 `app.build` 329 → 334；`config status --json` 多了 `sync_status`；文字输出多一行「同步状态：…」；`config --help` 换成新版（少了「暂无命令」与「同步状态那句话由运行中的 App 持有」两行，多了 `update install` 各行）。
- 装机版只读验证：`dockit --help` 有 `update install --yes` 一行、没有「暂无命令」；`dockit config status --json` 带 `sync_status{text:"iCloud 配置同步已关闭", at:null, from:"derived", live:false}`；`dockit update install --no-such --json` 与 `dockit update install extra --json` 都退出 2、`error.code` 为 `usage`。没有对装机版跑 `update check`、`update install --yes`、`config sync`、`config import`。
- 18:52 `chapter sop accept --app doc-tools-doctools --check agent_cli` 通过；`chapter agent-cli --json --app doc-tools-doctools` 读回：`passed`，29 项，命令 20、只在窗口 9、暂缺 0，problems 为空（改前是暂缺：命令 18、只在窗口 9、暂缺 2）。
- 18:53 对装上的 `/Applications/DocKit.app` 1.0.1 (334) 就地跑进程内离屏自检 `--ui-self-test`（经 2 号槽）：18 项通过、4 张离屏图，偏好域前后逐字节相同，没有留下进程；结果在 `build/notarized/agentcli2-20261007/evidence/ui-self-test-installed-334/`。
- 装机记录 `build/notarized/agentcli2-20261007/install-record.json`，证据在同目录 `evidence/`（不入库）。声明已释放。

本产品事项（派活点名的）：

- 「`update install` 这个词在转调与把关两处都放行」：把关按首词判定，`update` 上一轮就在两个包的 `DocKitCommandVerbs` 里，所以 Info.plist 不用改；转调原样带过去，只给它单独的时限（720 秒，其余 60 秒）。`test_update_install_is_forwarded_with_its_own_time_limit` 用会睡的桩把 `dockit update install` 与 `dockit public update install` 两条都实测了。
- 只读审计「只能人来」9 项：没有改动。

没做的、没验证的：

- 没有对装机版、本人的偏好域和 iCloud Drive 跑 `update install --yes`、`update check`、`config sync on|off`、`config import`。真实替换只在临时目录里的测试包上做过（本机一次性签名、文件型发行记录、不重开）；从本人 iCloud 里的私有发行记录走一次真实升级没有证据。
- 「配置与更新…」窗口这一版换成了总部现版：上一轮构建时跳过分发、留着旧副本，这一轮四份共用文件按约定同版（副本是别的会话 16:57 刷新并提交的 `5bff202`）。窗口只有离屏证据：测试里的运行中 App 按菜单项的方式把它建出来（在公证包上就地跑时也建出来了）；18 项界面自检不包含这个窗口。真实窗口没有人打开过。
- 私有 Homebrew 的 cask、发行资产和 iCloud 里的私有更新记录都没动，仍是 1.0.1-323；现在装着的 334 比更新源新。构建回执没有刷新，没有做性能测量。
- `scripts/accept/` 的功能、恢复、隐私三项没有重跑；`native_ui.sh` 没有经 Chapter 跑。Chapter 里这几项的证据输入已变。
- 转调的 720 秒超时那条分支只用缩短时限的桩测过，没有用真的挂起去测。
- 父仓 `CLAUDE.md` 的独立入口表里 `mac/bin/dockit` 那一行仍没写 `config`／`update`／`public`，不在本轮边界内。
- 产品 `build/` 下有多份 `DocKit.app` 副本，Spotlight 搜得到：`notarized/private-brew-20261006`（323）、`notarized/agentcli-20261007`（329，上一轮的公证包）、`notarized/agentcli2-20261007`（334，这一轮的公证包，与装机版逐文件相同）、`icon-refresh`、`native-icon-refresh`。它们是构建留存，不是装机换下来的旧包（那个在废纸篓），这一轮没有动；要不要清由本人定。
