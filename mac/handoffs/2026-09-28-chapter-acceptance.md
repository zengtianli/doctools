# DocKit · Chapter 固定验收（2026-09-28）

本轮仅在 doctools 父 Git 仓内修改。主 agent 集成并提交；子 agents 分别负责 functionality、recovery、privacy、native_ui 文件。没有修改共享模块，没有 push、发版、装机、部署或调用真实云端扫描。

## 已完成

- `project.yaml` 登记四项 `sop.accept`，共用工具为 `scripts/accept/_common.py`；补上 GUI 后端、调度器及文本引擎的输入绑定。
- functionality：14 项操作 / 20 个动态选项契约；真实 GUI 后端完成 DOCX 默认规范化、选择性规范化、Markdown 引号、Markdown 合并四条链，检查内容与原件哈希。
- recovery：11 项，涵盖错误信封、损坏文件、混合批次继续、后续请求恢复及原件保护。发现并修复 `docx_fmt.py text` 吞掉处理失败退出码；坏文件返回 1，正常文件继续处理。
- privacy：GUI 云端扫描声明发送文件名与正文至 Claude，外发同意默认关闭；未同意时在枚举目录和启动扫描前拒绝。固定脚本以操作系统禁网和禁止非 Python 子进程的负对照证明隔离，未发送文档。真实本地 clean / quotes 通过禁网验收，合成正文不进入 GUI 日志信封。
- native_ui：App 自身 `--ui-self-test` 入口，禁止激活、没有 WindowGroup 或可见窗口；真实 ContentView / CommandPalette / ViewModel 验证刷新、选择、搜索、导航、清空、提示关闭与隐私选项复位，输出四张离屏 PNG。截图背景已作不透明合成，主 agent 已查看主界面和隐私面板。
- 四项最终均经 `app_sop accept` 返回 `passed`，证据完全由 app_sop 写入 `perf/delivery-evidence.json`，原始日志与截图在 `perf/acceptance/`。
- 主 agent 验证：文本引擎相关 pytest **46 passed**；`./build.sh --check` 真实解码 14 项 / 20 选项通过；生产 ViewModel 的 `tests/run-check.swift` 重复执行保护通过；CLI surface 改前后相同（126 为含中间节点的子命令数）；forward probe 67 条通过；docx surgical 收口闸通过；`git diff --check` 通过。

## 复现

在本目录执行：

```sh
~/Dev/.venv/bin/python ~/Apps/chapter/engine/app_sop.py accept --app doc-tools-doctools --check functionality --check recovery --check privacy --check native_ui --json
./build.sh --check
```

在父仓执行回归：

```sh
cd /Users/tianli/Dev/tools/doctools
python3 -m pytest -q scripts/document/tests/test_docx_text_formatter_exit.py scripts/document/tests/test_docx_text_formatter_safety.py scripts/document/tests/test_docx_text_formatter_scopes.py
```

共享 `~/Dev/.venv` 当前没有 pytest，回归使用已有 pytest 的系统 `python3`；没有安装依赖。

## 范围与本人材料

- 原生截图覆盖非 key 的离屏窗口；侧栏选中行在该渲染条件下呈黑块，不能据此判断实际焦点选中样式。未激活窗口验证，不把它算作实体键鼠/焦点外观验证。
- 功能覆盖为上述四条链，不是所有格式和 14 个操作逐一实测；恢复不含断电、磁盘故障、强杀；隐私不含第三方服务留存策略或实际云端扫描可用性。
- installed_icon 仍由本人在 Chapter 确认。材料：`icon/AppIcon.png`、`icon/AppIcon.icns`、`icon/provenance.json`，以及 `perf/acceptance/native-{main,compact,palette,privacy}.png`。未替本人写装机图标证据。
- 开工已有未跟踪的 `perf/delivery-evidence.json` 与 icon_review 验收文件。保留原内容，delivery-evidence 仅允许 app_sop 合并本轮结果；这些原有文件不纳入本次提交。

## 受阻与 CLI 接手

1. Chapter 测试状态同步：本轮最后运行 `app_sop run --test-only` 返回 exit 75 / busy（另一轮 app_sop 正在运行）。已有固定验收证据及本地测试通过，不重复争抢共享队列。待队列空闲后，在本目录执行：

   ```sh
   ~/Dev/.venv/bin/python ~/Apps/chapter/engine/app_sop.py run --app doc-tools-doctools --test-only --json
   ~/Dev/.venv/bin/python ~/Apps/chapter/engine/app_sop.py run --app doc-tools-doctools --check-only --json
   ```

2. 当前源码相对已安装 1.0 (298) 有变化；原性能数据仍只代表 2026-09-26 的旧安装版。本轮禁止装机且要求跳过长时间空闲采样，因此没有更新性能数字或旧 build receipt。若后续决定装机并补新版本实测，在本目录执行下列命令；采样须满足接电源、空闲和负载门，不改旧日期伪装新数据：

   ```sh
   ./build.sh --install
   ~/Dev/.venv/bin/python ~/Apps/.claude/skills/app-lightweight/scripts/batch_measure.py doc-tools-doctools
   ~/Dev/.venv/bin/python ~/Apps/chapter/engine/app_sop.py run --app doc-tools-doctools --check-only --json
   ```

Chapter 已排队的只读复检可继续运行。本轮未声明四个产品维度整体完成。
