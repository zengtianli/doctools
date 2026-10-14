# DocKit 四项基础能力 · 2026-10-15

## 范围与身份

本轮组件 `doc-tools`，源码 `/Users/tianli/Dev/tools/doctools/mac`；最终边界只处理 product_basics 候选、不安装。沿 Chapter 最新派出者 claims 身份 `pid 32510 / chapter-task:393cf037ba19560db001948ba5282bb0` 写前 check，不另开声明、不释放；此前阶段沿当时派出者 55361/84812。入场时父仓源码、project.yaml、旧 build-receipt 与 delivery-evidence 已有其他修改，全部保留；提交仅含本轮源文件/登记块/证据。

固定标准：`~/Apps/chapter/docs/PRD-product-basics.md` 与 `engine/product-basics.json`。源码改动为 Lifecycle 产品接线与离屏入口、菜单快捷键、Xcode 源登记、build.sh、固定 product_basics 验收及文档。四份生命周期模块和新增 AppShortcuts 由总部 vendor 同步，未修改或复制生产引擎。

## 已完成与实际输入

- 设置窗口增加可持久化的刷新开关/组合（默认 ⌘R，⌘K 搜索，⌘, 配置与更新）；总部适配器拒绝系统保留组合与应用菜单冲突。
- 原配置工厂继续承接操作/目标格式，加入快捷键与共享更新选项的白名单、导入校验和 iCloud 迁移；外发同意、文档、模板路径、账号和运行记录不迁移。
- 更新沿原本人 iCloud 私有渠道和共享检查/安装事务，设置窗及 `dockit update options` 同源；自动行为默认关闭。
- `sop.product_basics` 四项及 `sop.accept.product_basics` 已登记，新设置在 agent_cli 对照里同步登记（快捷键可导入配置，更新选项有原共享命令）。固定脚本默认验已装包，可用 `--app/--receipt` 或对应 SOP 环境明确验候选；原件身份/版本/签名/可执行 SHA 与回执、全部构建输入哈希必须相符才运行。
- `build.sh --check` 真实后端 14 项解码、偏好/CLI→视图双向跟随通过；生命周期 12 项用例通过（1 项组装包用例由构建完成后另验）。Release 构建入口也执行这些门和组装包入口用例，退出 0。
- 最后补齐菜单启动顺序：共享层挂好菜单后再配置 ⌘,，刷新项并入原「操作」菜单；离屏自检直接调用同一 `installMenus`，断言设置组合和唯一菜单归属均 true。
- 最终候选：`build/DerivedData/Build/Products/Release/DocTools.app`，1.0.1 (344)，bundle `cyou.tianli.DocTools`，临时签名开发候选，未安装；2026-10-15 02:23:48 的实际快构建退出 0。
- 最终来源回执：`build/accept/product-basics-candidate-20261015-r2.json`，23 项实际生产输入（含父仓已有修改），source SHA `1614e53e27e2d30442540f0a556c1228f45cca54c2e44948a793be54abde31f4`；可执行 SHA `c6544af269f2a0f1c9a64b39fc7f4cd7209fab5cbb1ced56140e98d078444f4f`；回执 SHA `7d46e191fc56f09357bdcf587bf8f05e0d2df187b452fcfcc4f78634d219237d`。回执由 app_sop build-receipt 包住 `./build.sh` 的实际成功构建生成；保留旧 `perf/build-receipt.json` 和早期 343 回执。

## 固定验收结果

实际运行：

```sh
cd /Users/tianli/Dev/tools/doctools/mac
SOP_PRODUCT_BASICS_APP="$PWD/build/DerivedData/Build/Products/Release/DocTools.app" \
SOP_PRODUCT_BASICS_RECEIPT="$PWD/build/accept/product-basics-candidate-20261015-r2.json" \
chapter sop accept --app doc-tools --check product_basics --json
```

退出 0、`product_basics: passed`。最终正式原件 `perf/acceptance/run-g9rwec41/product_basics.json`，SHA `db282f182d8ea71941bc9068959d3329b0995c89540004f3342e148e28dde612`；Chapter 自动写入 `perf/delivery-evidence.json`，未手写该证据。原始 checks、逐条 CLI 响应、控件结果、浅/深 4 张离屏图及升级日志在 `perf/acceptance/product-basics-7b4d7577280c42e587c755b6c7d85a14/`。最终登记的更新 readback 是单条可直接执行的 `dockit update check --json`；此前阶段原件全部保留（包括 `run-nyw0p84k`、`run-0fijilv5`），以本段最终包与原件为准。

四项 14 个必测名全部为严格 true，另有 4 项 owner_data_unchanged：设置真实开关动作/新进程读回/坏导入保留；菜单真实刷新回调与 SwiftUI subscriber/关闭恢复/保留键冲突；云端写入接收/不可用本地回落/同字段本地优先且独立字段合并/坏远端恢复/关闭保留；测试发行源比较/坏源恢复/更新选项持久化/真实临时包替换并保留配置。原件 SHA 与用户偏好、真实配置文件和已装字节的前后哈希一致。

首次正式验收正确判 failed：临时升级用例误用了 IN_PLACE 保护入口而被 skip。仅改固定脚本接线为原临时包事务后重验，升级用例实际执行且退出 0，无 skip；旧失败原件保留。没有减检查或用顶层 ok 代替 checks。

最终回读逐项核了 14 个 policy 必测布尔、4 项用户数据不变、全部原件/回执/UI/升级日志 SHA，`app_sop.verify_build_receipt` 为 true；升级日志为真实临时包事务 1 条实际用例 OK、无 skip。新的 4 张浅/深离屏图也由 agent 检查，设置和更新内容可读。

## 已撤回的中途安装阶段

中途 Chapter 曾明确授权一次本机迭代快装（身份 84812）；沿原持证机和私有 tap 签名函数产出 Developer ID 候选，无公证/发布。SSH 审计上下文的签名失败，改成同一持证机登录域的一次性签名作业后成功，作业已移除。实际安装在解析旧包 designated requirement 的前置检查处 `IndexError` 退出，尚未复制或替换已装包。收到最新“不安装”边界后停止该阶段，撤回本轮新增安装脚本改动，`build.sh` 返回已提交的原快构建入口，再生成上述最终独立候选回执和验收。

历史签名/失败原件保留在 `build/iteration-install-20261015/`、`build/iteration-final-20261015/`；失败日志已原样保存为 `build/iteration-final-20261015/installation-failed.log`。`build/accept/product-basics-installed-20261015.json` 名字来自当时计划，它绑定历史候选，不能当已装回执，且其构建输入包含现已撤回的安装脚本。后续不得重放这些过期身份/输入的脚本。当前已装仍为 1.0.1 (340)，可执行 SHA `ec506d7d0b8e68a04e8fc24b8dc3c14fd6fcf279c9ab635fddf21f3e33785ae0`，与安装前原件逐字节一致，正式签名和原入口保留；不将本轮候选回执写入默认装机回执。

## 边界与接续

- UI 全部进程内禁止激活（prohibited）、无可见/key 窗口，不进 Dock、不合成键鼠、不碰剪贴板。四图已由 agent 目视检查，更新图为同一设置窗滚动到底部的离屏视图。
- iCloud 已验隔离本机目录传输、冲突和错误恢复；苹果服务/跨设备传输未验，不能据此声称另一台已收到。
- 已装版保持原样，本轮没有装机、公证、推送、发布、部署、改更新源或操作账号/真实文档/既有后台。候选通过只代表代码阶段；Chapter 后续独立 install 阶段须按原正式签名安装边界执行，并对已装原件重跑默认固定验收。
- 父仓后端顶层与“未装 App”fallback 帮助仍为旧共享文字（未列新 update options，旧文案称 update check 完全只读）；本轮范围内不改父仓。已组装包的 config/update 完整帮助逐字转调当前 App，新增选项实际可用；回归分别核原 fallback 与当前原件，而非假称两版文字一致。
- 当前范围已完成，无用户待决策事项。其他旧 SOP、全量性能与公开变体未展开。
