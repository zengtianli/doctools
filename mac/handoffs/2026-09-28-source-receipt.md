# DocKit 来源登记与装机闭环（2026-09-28）

本轮承接 `2026-09-28-chapter-acceptance.md`。用户已明确长期授权本产品装机与推送；本轮实际安装 1.0 (304)，未改变公开范围、发版或部署。未修改共享模块。

## 已处理

- 上轮误把 `source_extra_files` 写成相对路径，把 `test_extra_files` 写成 glob；现行 app_sop 两个字段只接受逐个绝对文件。已将三个生产后端文件改为绝对路径，仓内测试改用 `test_inputs: [tests/**, scripts/accept/**]`。
- 将构建时实际读取的 `catalog.yaml` 纳入 `sop.source`；构建来源现覆盖 15 个文件（含三个后端输入）。`code_key` 已可正常生成，不再报“无法确认测试对应的业务源码输入”。
- 装机前回读唯一安装 `/Applications/DocKit.app` 为 1.0 (298)，未发现运行中的 DocKit；旧安装与当前源码不匹配。
- 使用官方 `app_sop build-receipt` 包裹既有 `./build.sh --install`，真实构建、签名、装机并自动写 receipt；没有手填构建证据。
- 装机后回读 1.0 (304)，可执行 SHA256 `a6c3aae1a50ec614ae669effba7775baebf64043148f7c278d5562e892b91fb4`；源码摘要 `c3474699e54d143aa908841429c48ef073a069dabe7b9ad2bf0b464ee43f8ee5`；`verify_build_receipt` 返回 true。
- 已核验旧 298 包完整保留在 `~/.Trash/app-rebuild-20260928-133032-43613/DocKit.app`，可执行哈希仍为 `6dbd0a8b1f1eea28af7caf7f86b0127805aada5626a80ad552a5f08b5f379fd1`。应用数据、bundle ID 和系统服务配置未改。
- 为新安装补跑 functionality / recovery / privacy / native_ui 四项固定验收，app_sop 全部返回 passed；证据自动合并到原有 delivery-evidence，不手写通过状态。
- `./build.sh --check` 通过：14 项操作、20 项动态选项真实解码；原生自检 18 项断言和四张离屏截图通过。没有激活应用窗口或访问剪贴板。
- native_ui 固定脚本只增加日志路径脱敏，保留完整 stdout/stderr 与退出码，实际运行后由 app_sop 重写日志和哈希；已跟踪的最新 perf 证据不再包含本机 home/临时目录路径。

## 复核与原件

- `perf/build-receipt.json` 是本次真实装机构建的官方收据，绑定源码提交 `1a91e15`；后续证据/文档提交不改变它已核验的源码内容。
- `build/chapter-accept-installed-20260928.json`、`build/decode-check-installed-20260928.log`、`build/chapter-audit-installed-20260928.json` 为本轮原始输出；build 目录忽略入库，可按下列命令重现。
- 无锁只读 `app_sop audit --offline` 已确认 build-receipt / install 都是 ok，当前版本 1.0 (304)；未同步的 test 变为“当前代码还没跑过测试”，原输入登记错误已消失。
- 原有未跟踪的 icon_review 文件与 delivery-evidence 文件不纳入本轮提交；后者仅由 app_sop 合并四项验收。

## 推送范围

`origin` 仍为既有 `zengtianli/doctools` PUBLIC 仓库，mac 组件在远端原已存在。本轮不改变可见性。推送前已读取 `git log --stat @{u}..HEAD`；原领先三笔均为已完成的 DocKit 登记、验收及对应后端修复，无其他组件待完成改动。

已读取 `.github/workflows/gates.yml` 与项目提交钩子：main push 仅触发 gates 的 Python 3.12 / 3.13 闸门、单元测试和冒烟；没有 ci_scripts、Xcode Cloud 或部署监听配置，没有自动发版/装机/部署动作。源码与证据将由主 agent 统一推送；CI 结果单独回读，不把推送成功说成 CI 已通过。

## 需要本人

installed_icon 仍由本人在 Chapter 确认。当前已安装 1.0 (304)，包内 ICNS 哈希与正式原图链一致；材料为 `icon/AppIcon.png`、`icon/provenance.json` 和 `perf/acceptance/native-*.png`。未替本人写图标通过。

## 受阻 / 自动接续

1. `app_sop run --test-only` 返回 exit 75 / busy；后续只读探测确认全局锁仍占用，未反复调用同一失败命令。源码登记与本地测试已通过，但 Chapter 测试状态尚待锁释放后写入。队列空闲后，在本目录执行：

   ```sh
   ~/Dev/.venv/bin/python ~/Apps/chapter/engine/app_sop.py run --app doc-tools-doctools --test-only --json
   ```

2. 当前性能数字仍明确属于旧安装 1.0 (298)；安装已更新，因此 304 的性能不能复用为新实测。本轮按快速处理要求跳过长时间空闲采样。满足接电源、空闲和负载门后，在本目录执行（不加 --now，不放宽门槛）：

   ```sh
   ~/Dev/.venv/bin/python ~/Apps/chapter/engine/app_sop.py run --app doc-tools-doctools --stage perf --json
   ```

完成后统一只读复检：

```sh
~/Dev/.venv/bin/python ~/Apps/chapter/engine/app_sop.py run --app doc-tools-doctools --check-only --json
```
