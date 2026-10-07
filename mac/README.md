# DocKit

[English](README_EN.md)

App 菜单新增「配置与更新…」和「检查更新…」。记住上次操作与各操作的目标格式，可导出、导入，并选择开启 iCloud 配置同步（默认关闭）。勾选规则仍按操作默认值重新确认；文档、外发同意、模板路径与运行记录不参与同步。本机版从自己的 iCloud 私有更新目录读取发行清单。

[doctools](../README.md) 的 SwiftUI 桌面端，操作与选项从同一文档后端 `gui-ops` 加载。源码、文档引擎与构建由父仓统一维护；本客户端依赖本机 `~/Dev/.venv`，不是独立分发版。

```sh
./build.sh --check    # CodingKeys + 真实后端解码
./build.sh            # Release 构建、签名，不装机
./build.sh --install  # 明确要求时安装到 /Applications/DocKit.app
```

显示名取自 catalog，bundle ID 保持 `cyou.tianli.DocTools`。构建使用共享 Xcode 选择器。JSON 与客户端开发约定见 [CLAUDE.md](CLAUDE.md)。

固定验收在 `scripts/accept/`，通过 Chapter 运行并由其写入证据：

```sh
~/Dev/.venv/bin/python ~/Apps/chapter/engine/app_sop.py accept --app doc-tools-doctools --check functionality --check recovery --check privacy --check native_ui --json
```

功能与恢复测试只使用临时文档。原生界面验收会构建当前源码，并调用 App 的 `--ui-self-test` 离屏渲染真实视图；不装机、不抢焦点，也不操作剪贴板。它证明本地构建的行为，装机图标仍需本人在 Chapter 确认。

规范化、引号等操作由本机文档引擎执行。“敏感词扫描”使用外部 Claude 服务，会发送扫描目录内 `.md` 文件的文件名与正文（不读取 docx）；每次选择该操作时须明确勾选外发同意。验收只检查拒绝路径与本地处理，不发送文档，也不证明外部服务可用。

## 命令行（给 agent 用）

GUI 给人用，`dockit` 给 agent 用。两者调用同一个后端 `doc_gui_backend.py` 的 `gui_run`：操作目录、选项校验、外发同意门、覆盖门和产出归因只有一份实现。`dockit` 是 App 包内的 POSIX 脚本（`DocKit.app/Contents/Resources/bin/dockit`），执行共享环境 `~/Dev/.venv` 里的工作树后端，所以后端改动不用重建 App。`./build.sh --install` 会把它链接到 `~/.local/bin/dockit`；找不到 venv 或后端时退出码为 69。

```sh
dockit ops                                          # 14 个操作，带「破坏性」「外发」「只诊断」「覆盖需确认」标记
dockit ops clean                                    # 输入格式、目标、选项、默认值和示例
dockit ops --json                                   # 与 GUI 的 gui-ops 相同：{"ok":true,"ops":[...]}
dockit run clean --opt scope.table=0 --json /abs/报告.docx
dockit run convert --to md /abs/a.docx /abs/b.pdf
dockit run convert --to word --dry-run /abs/b.md     # 先看会不会覆盖已有的 b.docx
dockit run formatclone --opt ref=/abs/范式.docx /abs/草稿.md
dockit run renum --to tabfig /abs/报告.docx
dockit run bidfinal --to pei --json /abs/标书.docx   # results[].verdict：pass / red / error
dockit run view /abs/说明.md                         # 报告 HTML 路径；加 --open 才打开浏览器
dockit run lowercase --dry-run /abs/数据.xlsx        # 破坏性操作先看计划，确认后再加 --yes
dockit run typeset --background --json /abs/报告.md  # 长任务立刻返回 job id
dockit status <job> --json                          # running / done / interrupted / lost，结束后附完整结果
dockit status --json                                # 读回当前状态：已装版本与构建号、界面记住的设置、最近的后台任务
dockit settings --json                              # 界面记住的上次操作与各操作的目标格式
dockit settings set target_formats.convert md       # 改其中一项；另一个键是 last_operation <op>
dockit doctor                                       # venv、uv、soffice、markitdown、pdftotext、claude 等依赖
dockit config status --json                         # 「使用 iCloud 记住配置」开关、开关下面那句同步状态（sync_status）、可迁移的配置项、App 是否在运行
dockit config export -o /abs/dockit-config.json     # 导出配置；config import <file> --yes 导入，config sync on|off --yes 拨开关
dockit update check --json                          # 当前版本、此渠道最新版本、有没有新版、怎么升级
dockit update install --dry-run --json              # 升级到新版会做什么；确认后 update install --yes（与窗口「升级到新版…」同一个安装器）
```

读命令是 `ops`、`status`、`doctor`、`settings`、`run --dry-run`、`config status` 和 `update check`，不写任何文件、偏好或任务记录；写命令是 `run`、`settings set`、`config export`（不改设置，只写你指定的那个文件）、`config import`、`config sync` 和 `update install`。`dockit --help` 列出这两组命令、每条命令的 `--json` 输出形状、退出码表和只在窗口里的动作。

退出码：0 表示成功；1 表示业务失败（有输入失败或被跳过、门检有红门、doctor 有缺项）；2 表示用法或校验错误（未知操作、选项或目标，缺必填项，缺外发同意，破坏性操作缺 `--yes`，会覆盖已有同名文件却没给 `--yes`，没有有效文件）；75 表示同一目录树（含上级或下级目录）正有另一个 DocKit 操作在运行，稍后重试即可；128+N 表示被信号 N 中断（例如 agent 超时发出的 SIGTERM 得到 143），正在运行的引擎进程组会一并终止；69 表示找不到 venv 或后端。

`--json` 输出与 GUI 相同的结果信封：`ok`、`all_ok`、`results[].outputs`（绝对路径）、`results[].modified_in_place`、`succeeded`/`total`、`skipped_missing`、`log`，失败时另带 `error` 和 `error_code`。`ok` 是请求级的：请求被处理了就是 true，个别输入失败时也是 true；判断是否全部成功看 `all_ok`、逐个输入的 `results[].ok` 或退出码。命令行的用法错误加 `--json` 时同样返回 JSON。

长任务：`run` 默认同步，套模板或 soffice 冷启动可能需要几分钟，调用时留 10 分钟超时。`--background` 先在前台做完全部校验（校验错误照常以 2 退出），再在后台运行并立刻返回 `job.id`；`dockit status <id>` 查询，任务结束后附完整结果信封，退出码与同步运行时相同。任务记录在 `~/Library/Caches/cyou.tianli.DocTools/jobs`，保留 7 天。

外发、覆盖和破坏性操作：

- `scan` 会把目录内 `.md` 文件的文件名与正文经 `claude -p` 发送至 Claude。不带 `--opt privacy.cloud_consent=1` 时，命令在列目录之前就以 2 退出。只有本人明确同意后才加这个选项；它会启动 `claude -p`，不要在 Claude Code 会话内部调用（llm_client 有嵌套限制）。发现项在 `results[0].findings`；目录里没有 `.md` 时不启动引擎、不发送任何内容，`findings` 为 `[]`、`md_files` 为 0。
- 覆盖门：`convert`、`typeset`、`merge`、`split` 的部分产出与已有文件可能同名（`b.md` 转成 `b.docx`、`merged.md`、`merged.csv`、按工作表命名的 `b_Sheet1.xlsx` 等），引擎会直接覆盖且不留备份。目标已存在时这几个操作默认拒绝（`error_code: would_overwrite`，`overwrites` 列出路径）；确认覆盖加 `--yes`，GUI 里勾选「允许覆盖已有同名文件」。`--dry-run` 同样报出会覆盖哪些文件。带 DocKit 后缀的产出（`_fixed`、`_styled`、`_序号修正`、`_lower`、`_split` 目录）重跑时覆盖上一次的产出；`formatclone` 的 `_成品` 永不覆盖。
- `fontunify` 和 `clean` 处理 pptx 时原地改写原文件，结果里用 `modified_in_place` 标出。第一次备份为 `<名>.pptx.backup`；已有 `.backup` 时另存 `<名>.bak-N-日期.pptx`，不会用改过的文件顶掉原件。`fontunify`、`lowercase`、`stripchrome` 须加 `--yes`；`--dry-run` 只列出将处理的文件，不执行、不写盘。
- 其余产出写在源文件旁边，原件不动。同一目录树（含上级或下级目录）同时只允许一个 DocKit 写操作，所以一次运行的产出不会记到另一次 DocKit 运行名下；其他程序同时往该目录写的文件仍可能被算作产出。

覆盖范围：

| GUI 能力 | 命令行 |
|---|---|
| 侧栏操作列表、说明、支持格式、破坏性标记 | `dockit ops [--json]` |
| 目标单选、选项分组、默认值、必填项、「允许覆盖已有同名文件」 | `dockit ops <op>`；`run --to T`、`--opt K=V`、`--yes` |
| 「执行」与结果区（逐文件成败、产出、成功 N/M、跳过的文件、日志） | `dockit run … [--json] [-v]` |
| 执行中的等待 | `run --background` 与 `dockit status <id>` |
| 「已就绪 / 后端不可达」 | `dockit doctor [--json]` |
| 敏感词扫描的外发同意 | `--opt privacy.cloud_consent=1`（同一道门） |
| 记住上次操作与各操作的目标格式 | `dockit settings [--json]`；`dockit settings set <键> <值>` |
| 「配置与更新」窗口里的版本与构建号 | `dockit status [--json]` 的 `app` |
| 「配置与更新」窗口：使用 iCloud 记住配置、导出配置、导入配置 | `dockit config sync on\|off --yes [--dry-run]`、`config export -o <file> [--force]`、`config import <file> --yes`；`config status` 读回 |
| 「配置与更新」窗口：开关下面那句 iCloud 配置同步状态 | `dockit config status [--json]` 的 `sync_status{text, at, from, live}` |
| 「配置与更新」窗口：检查更新 | `dockit update check [--json]` |
| 「配置与更新」窗口：升级到新版 | `dockit update install --yes [--dry-run] [--json]`；`update check` 读回 |

`settings` 读写的就是 App 的那两个偏好键（`dockit.lastOperation`、`dockit.targetFormats`），不另存一份；值先按操作目录校验，未知操作或目标以 2 退出、不写入。DocKit 在下次启动时按新值选中。偏好域里的窗口位置、最近目录等其他内容，读命令不输出，写命令不改动。

只在窗口里：拖入文件或目录、清空待处理文件、逐个移除待处理文件、恢复默认选项、清除已选参考文件、在 Finder 显示产出、关闭提示条、搜索功能面板（⌘K）、打开「配置与更新…」窗口，以及 `--ui-self-test` 离屏自检。命令行分别用绝对路径参数、`ops` 列出的默认值和 JSON 里的产出路径替代。

`config` 与 `update` 是「配置与更新…」窗口里的几项，属于 App 本身（偏好域、版本、发行渠道）：由 `DocKit.app` 的可执行文件执行（各产品共用的命令层，不创建窗口、不进 Dock），`dockit` 只把整段参数转过去，输出与退出码原样带回；已在运行的 DocKit 自己跟随命令的改动。这两组命令的 `--json` 是共用命令层的形状，失败为 `{"ok":false,"command":…,"error":{"code","message"}}`，与其余命令平铺的 `error`／`error_code` 不同；退出码只有 0、1、2，全部用法与各 `error.code` 见 `dockit config --help`。没能交给 App 可执行文件时退出 1：`app_missing`（找不到 `DocKit.app`）、`app_outdated`（装着的是不带这组命令的旧版，不会启动它）、`app_failed`、`timeout`。`config import`、`config sync` 与 `update install` 要 `--yes`；`update check` 与 `update install` 读本人 iCloud Drive 里的发行记录。`update install` 没有新版时退出 0（`installed` 为 false）；有新版时验证发行包与签名、替换当前 App（运行中的先退出、换好再重开），旧包移到废纸篓，替换失败回滚；它要下载、验证、替换，`dockit` 等它最多 720 秒（其余命令 60 秒），超时后先 `dockit status` 读回版本，不要直接重发。

「配置与更新…」窗口的每一项都有命令，没有暂缺项。同步状态那句话的来源写在 `sync_status.from`：`app`（运行中的 DocKit 此刻显示的）、`record`（App 没在运行，上一次同步留下的那句）、`derived`（没有记录，按开关给初值）。逐项对照登记在 `project.yaml` 的 `sop.agent_cli`。

<!-- lightweight:start -->
## 资源占用

| 安装后占用 | 空闲内存 | 空闲 CPU | 读取操作列表（生产后端，含 Python 进程启动） |
|---|---|---|---|
| **2.3 MB** | **46 MB** | **0.05%** | **29 ms** |

原生 SwiftUI 客户端，文档能力复用本机共享 Python 引擎；空闲时后端进程退出。

<sub>v1.0 (298) · Mac16,12 / Apple M4 / macOS 27.2 · 已安装自用版首页，载入 14 项真实操作与动态选项；未打开或处理用户文档。 · 2026-09-26。数字来自所列设备实测，版本更新后重新测量。内存口径为 phys_footprint；CPU 为 60 秒采样窗内 CPU 时间 ÷ 墙钟；大小按十进制 MB。原始数据见 [perf/lightweight.json](perf/lightweight.json)。</sub>
<!-- lightweight:end -->
