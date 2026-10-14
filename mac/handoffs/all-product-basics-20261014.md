# DocKit 四项基础能力 · 2026-10-15

## 最新范围与完成

组件 `doc-tools`，源码 `/Users/tianli/Dev/tools/doctools/mac`；只处理 product_basics，并按 Chapter 最新明确授权完成一次本机迭代安装。当前派出者 claims 为 `pid 48978 / chapter-task:393cf037ba19560db001948ba5282bb0`，写前 check，复用原跨机 lease，不另开声明、不释放。入场时父仓源码、project.yaml、旧 build-receipt 与 delivery-evidence 的其他修改、暂存和后台保留；旧装机回执原件先另存再由 Chapter 写入新来源。

最终已装为 **1.0.1 (345)**，同 bundle `cyou.tianli.DocTools`、同 Developer ID team `B9LJH93LA4`，满足旧包指定签名要求。默认已装固定验收退出 0，14 个必测行为逐项为严格 true；无用户待决策、无当前受阻。已装原件与持证机签名候选的全部包成员一致，用户偏好、真实云文件、原 CLI 链接、服务配置和后台进程未变化，未启动正常窗口或重启常驻实例。

固定标准唯一源：`~/Apps/chapter/docs/PRD-product-basics.md` 与 `engine/product-basics.json`。

## 实现与复用

- 设置与配置：原 Lifecycle 工厂承接操作/目标格式，新增刷新快捷键和共享更新选项白名单；真实开关动作、新进程读回、坏导入保留均通过。
- 快捷键：⌘, 配置与更新、⌘K 搜索、默认 ⌘R 刷新；共享 AppShortcuts 支持开关/组合持久化、系统保留键和菜单冲突校验；实际刷新菜单回调与 SwiftUI subscriber、禁用和恢复均通过，共享菜单挂好后才接线。
- iCloud：原共享配置同步开关、状态与导入/导出；实际隔离目录推送/接收、独立字段合并、同字段本地优先、损坏远端恢复、不可用本地回落和关闭保留均通过；外发同意、文档路径、账号和运行记录不迁移。
- 更新：原私有 iCloud 发行源与共享更新事务，检查状态和错误恢复、更新选项新进程读回、真实临时包升级并保留配置均通过；自动行为默认关闭。
- `sop.product_basics`、`sop.accept.product_basics` 和新增设置的 agent_cli 对照已登记；固定脚本默认验 `/Applications/DocKit.app` 与 `perf/build-receipt.json`，候选 `--app/--receipt` 入口仍保留，二者不会混称。
- 五份共享 Swift 原版由总部 vendor `--check` 逐字节核对；未复制生产引擎。签名只调用既有私有 tap 的 identity/run/snapshot/signed_code_digest 等函数，不调用公证或发布函数。

## 实际构建、签名与安装原件

- 原快构建 `./build.sh --build`：CodingKeys、真实后端 14 个操作解码、偏好及 CLI→真实视图双向跟随通过；生命周期 12 条用例中组装包用例先 skip，完成组装后另实际执行该 1 条并 OK，Release build succeeded。
- 最终签名候选：`build/iteration-authorized-20261015-48978-r3/signed/DocKit.app`；回执 `build/accept/product-basics-signed-20261015-48978-r3.json` 由 Chapter 包住实际构建/持证机签名生成，26 项输入含原父仓已有修改、安装接线和签名适配器；签名输入 zip/生产签名脚本/传输全部 SHA 核验。
- 持证机为原 macbook；复用原派出者跨机 lease，登录域一次性签名作业 `cyou.tianli.DocKit.sign-once.48978.20261015` 已 bootout，最终回读不存在；没有新增持久登录项或改钥匙串/账号。
- 实际安装沿 `./build.sh --install-built build/iteration-authorized-20261015-48978-r3/signed/DocKit.app build/accept/product-basics-signed-20261015-48978-r3.json`，外层由 `app_sop build-receipt --artifact /Applications/DocKit.app` 包住，退出 0，仅实际替换一次。
- 安装前先严格来源/签名/同 team/旧指定要求/进程/链接校验，完整暂存再替换、失败回滚；新增 `scripts/install-built.py` 是本产品接线，复用 Chapter 来源核验与总部 distribution_grade。默认开发候选和 `--install` 无 Developer ID 的入口确实拒绝（69），原包字节不变；正式候选 check-only 实际通过。
- 当前装机来源：`perf/build-receipt.json`，2026-10-15 02:40:40；回执 SHA `9a5d83307a673f7acd3e4dcc764d52bba53a33f21c7d651804b6986ec3f1e5bb`；source SHA `76bb8c6fdc2e3fb61c723438e1d6dedbb0dc05079fcb94193d6ca1231a3fbe7d`；已装可执行 SHA `39c8bfc7c62a2f933b17a17c465e2ecf3de15cba9c51d881063fe718d61d0203`；最终 `verify_build_receipt` true。
- 旧已装 1.0.1 (340) 的完整包已保留于 `/Users/tianli/.Trash/app-rebuild-20261015-024040-71763/DocKit.app`，全部成员哈希与安装前相同；旧默认回执保留于 `build/iteration-authorized-20261015-48978-r3/previous-installed-receipt.json`。安装前后事实和最终读回在同目录的 `installation-before.json`、`installation-after.json`、`final-readback.json`；真实构建与签名日志也在该目录。

## 已装固定验收

实际执行默认入口，无候选覆盖：

```sh
cd /Users/tianli/Dev/tools/doctools/mac
env -u SOP_PRODUCT_BASICS_APP -u SOP_PRODUCT_BASICS_RECEIPT \
  chapter sop accept --app doc-tools --check product_basics --json
```

退出 0、product_basics passed。正式原件 `perf/acceptance/run-flkcbgls/product_basics.json`，SHA `46f41163c3750e2fb7149549a3dc8824fdaeb08294ff698155bae6dba690d3be`；Chapter 自动写入 `perf/delivery-evidence.json`，未手写通过证据。原始 14 checks、4 项 owner_data_unchanged、逐条真实 CLI 响应、UI 数值/状态、浅深 4 张离屏图及真实临时升级日志在 `perf/acceptance/product-basics-91d7522a2b3c47bab7fe8fb35d1204a9/`。

已逐项回读全部必测布尔、原件/回执/UI/升级日志 SHA 和用户数据前后哈希；升级用例实际 OK、无 skip；4 张图由 agent 检查。最终 `chapter basics --app doc-tools --json` fresh passed。安装前后没有 DocKit 正常实例/生产后端进程，未改变该状态；其他后台保留。

## 失败原件与历史边界

早期候选 343/344 的回执与验收仍保留（包括 run-nyw0p84k、run-g9rwec41），它们是候选验收，不代表已装。此前 84812 安装阶段在读取指定签名要求时 IndexError、未替换包，收到 32510 的“不安装”边界后撤回安装接线；本轮 48978 再获一次实际安装授权才继续，修本产品入口后完成安装。

本轮第一签名派活在远端任务父目录已存在时提前停止，尚未签名或安装；改成核新隔离子目录与输入 SHA 后签名完成。正式安装前 check-only 又正确拦住未加“=”的内联要求参数；本机 codesign 手册及真实旧指定要求校验确认格式，修正后重新实际构建/签名并通过预检。生产包只在最终正确事务中替换一次；首次临时升级测试 skip 的早期固定脚本问题也已修，最终真实升级无 skip。失败日志和前轮原件保留，不减检查或循环同因输入。

历史派活脚本含过期 PID，后续不得直接重放；无过期 PID 的原签名适配器只作为 SHA 绑定的函数接线复用。未来任务须用当轮派出者身份和新输入重新生成来源，不再消费本轮一次性安装授权。

## 注意与剩余事实

- 所有 UI 验收都在 App 进程内 prohibited 激活策略与私有 re-identified 副本中离屏完成，无 visible/key 窗口、不进 Dock、不合成键鼠或碰剪贴板；真实安装不启动正常 App。
- iCloud 验证的是隔离本机目录传输与恢复；苹果服务器和跨设备传输未验，不能据此声称另一台已收到。
- 本轮是明确授权的 Developer ID 本机迭代，未公证、推送、公开发布、部署或修改真实更新源；正式旧包仍完整可回溯，公开发行走原独立流程。
- 父仓“未装 App”fallback 帮助仍是旧共享文字；已装 config/update 完整帮助转调当前原件，新增选项实际可用，本轮不改父仓。
- 当前 product_basics 和授权一次安装均完成，无用户待决策、无当前受阻；其他旧 SOP、公开变体和全量性能未展开。
