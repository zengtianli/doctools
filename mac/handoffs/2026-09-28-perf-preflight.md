# DocKit 304 性能重测预检（2026-09-28 14:32）

本轮唯一待办是重新实测性能；复用现有 `scripts/accept/_common.py`，没有缺失固定验收脚本，也没有重做已通过验收。一个子 agent 只读核对测量流程，主 agent 实际核验安装与标准空闲门。未改共享模块。

## 已核验

- `/Applications/DocKit.app` 仍为 1.0 (304)，可执行 SHA256 为 `a6c3aae1a50ec614ae669effba7775baebf64043148f7c278d5562e892b91fb4`。
- `app_sop.verify_build_receipt` 返回 true：装机可执行文件、图标和当前构建输入匹配，无需再次构建或装机。
- 直接调用现行 `app_sop.steady()` 的实际结果：`false`，原因“用户 0 秒前有操作”，检查时间为 2026-09-28 14:32:28 +08:00。
- 预检原始输出存于忽略入库的 `build/perf-preflight-20260928.json`；这只是准入检查，不是性能测量或验收通过证据。

## 自动接续

标准要求接电源、用户至少连续 600 秒无操作、负载低于核数、没有其他构建；当前空闲条件不满足。依用户“快速处理、不等待长时间空闲采样”的要求，本轮没有启动应用、采样、反复检查或降低门槛，`perf/lightweight.json` 保留原 1.0 (298) 的真实测量日期和数字。

该阻塞可由 Chapter 在条件满足后自动接续；用户也可在本目录执行：

```sh
~/Dev/.venv/bin/python ~/Apps/chapter/engine/app_sop.py run --app doc-tools-doctools --stage perf --retry --json
```

`--retry` 用于外部空闲条件恢复后的明确重试，不绕过空闲门；不要使用底层测量脚本绕开准入检查。测量需沿现有隔离候选文件流程，完成采样和清理后再写性能证据。

已入队的两项只读重检交由 Chapter 执行，不重复操作；装机图标仍由本人在 Chapter 确认，当前安装、正式图标及既有四张界面截图已经备齐。

本轮仅新增本交接文件，不改业务、性能数字、测试证据或安装。若同步到既有远端，main 推送只触发原 gates 闸门 CI，不发版、不部署、不改变公开范围。

## 2026-09-29 03:48 复查

- `app_sop.py run --stage perf --retry` 返回 `busy`：另一轮 app_sop（unrevoke-mac 检查）持有全局锁；同时另一会话在跑 notifhub 模拟器测量。
- 直接调用 `app_sop.steady()`：`(False, '负载 11.0 ≥ 10')`（5/15 分钟负载 73.7/62.1），接电源满足。
- 未采样、未改性能数字、未构建装机；`perf/lightweight.json` 仍为 1.0 (298) 的真实测量。接手命令同上，由 Chapter 在空闲门满足后自动补测。
- 工作区里 `perf/acceptance/icon_review.*`、`perf/delivery-evidence.json`、`perf/installed-icon-review.json` 为 Chapter 写入的未跟踪证据，本轮未改、未提交。
