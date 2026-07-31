# docx 工具链改造 · 路线图（冷启动自包含 · 2026-07-31 立）

> 读者 = 无上文的冷启动会话。进场三步：`cd ~/Dev/tools/doctools` → 读 `CLAUDE.md` →
> 跑下方「验收命令块」拿基线。改造史与数字口径见 `handoffs/before-after.md`（含 §十一）。

## 一 · 项目定位

把 doctools 的 docx 轴从「一堆各自为政的脚本」改造成「**数据驱动 + 少数引擎 + 闸门守护**」：

| 轴 | 引擎 | 数据 | 状态 |
|---|---|---|---|
| 版面（字体/编号/页眉脚/行距…） | `scripts/document/typeset_apply.py`（29 动作） | spec yaml（`config/spec-examples/report-generic.yaml`） | ✅ 已建 |
| 内容修订（意见落实/修订标记/批注） | `scripts/document/docx_revise.py` + `lib/docx_revise.py` | ops yaml（`config/spec-examples/revise-ops-example.yaml`） | ✅ 已建 |
| 存盘安全 | `lib/docx_safe_save.py` 收口 + `~/.claude/hooks/docx-surgical-guard.sh` 硬拦 | — | ✅ 59 脚本全挂 |
| 子命令声明 | `scripts/document/sub/_groups.py` 一张表 | Opt/Target 数据行 | ✅ 20 壳已折 |

## 二 · 已完成（别重做，数字都有尺子）

1. surgical 收口：59 脚本/11 repo 全接 `docx_safe_save`，守卫 36/36，PreToolUse 硬拦（exit 2）。
2. spec 排版引擎：29 动作、`path-pre` 相位、计数键 fail-closed、readonly/tail_path。
3. 20 个转发壳 → `_groups.py` 声明表；**两道闸门**：`tools/cli_surface.py`（接口指纹）+
   `tools/cli_forward_probe.py`（转发 argv 录音）。
4. 两轮退役：117→98→96（原口径；`before-after.md` §十一记着尺子），docx_cli 子命令树 128→125。
5. 修订注入引擎 docx_revise：四种 op、锚点唯一命中 fail-closed、within 标题式消歧跳目录、
   落位自检、部件断言。对真件（qual-supply）端到端验证 + 反向验证 4/4。
6. 测试 **80 绿**（`scripts/document/tests` + `sub/tests`）；`script_graph` 0 孤儿。

## 三 · 本项目铁则（每改必守）

- **两道闸门每改必跑**：改前 dump surface/probe 基线，改后 diff **必须恰好等于预期变更**，零意外。
- 加子命令 = `_groups.py` 加一行数据；加独立入口脚本 = `CLAUDE.md` 独立入口表加一行。
- 版面改动进 spec yaml、内容修订进 ops yaml——**禁在项目/会话里现编注入脚本**。
- 守卫/自检 fail-closed（空集不报绿），写完立刻反向验证。
- 退役任何功能 = 用户拍板；删除走 `~/.Trash/<slug>-<date>/` 带 MANIFEST。

## 四 · 验收命令块（可复制，改完全跑）

```bash
cd ~/Dev/tools/doctools
python3 -m pytest scripts/document/tests scripts/document/sub/tests -q   # 80 passed
python3 tools/cli_surface.py > /tmp/surface.json                          # 与改前 diff
python3 tools/cli_forward_probe.py > /tmp/probe.json                      # 50 条全正常
python3 tools/script_graph.py                                             # 0 个没人引用
python3 tools/check_docx_collar.py                                        # 收口全挂
```

## 五 · 待改造 backlog（按优先级；判断已下过，别重推导）

| # | 事 | 体量 | 已下的判断 / 验收标准 |
|---|---|---|---|
| P1 | **bid_\* 家族 7 文件 → 1**（`bid_gate.py` 子命令族，共 1,818 行） | 中 | driver `bid_final` subprocess 串 4 道门 → 直接函数调用；`bid_residue_lib` 保持 SSOT。**风险=活着的标书门检**：验收 = 拿真标书件改前改后干跑（`doc_dispatch.py bidfinal`），诊断输出逐字一致 |
| P2 | **`docx_tools.py` 1705 行单体拆解**：extract / check / track-changes 三个互不相干的能力挤一个文件 | 中 | 拆成 sub/ 实现模块 + `_groups` 行（走 20 壳同款两道闸门流程）；`docx_cli` 的 extract/check/… 转发随之改指向 |
| P3 | **docx_cli legacy 层收编评估**：剩余 10 个旧脚本仍是「standalone argparse + docx_cli 转发」双入口 | 大 | 逐个评估收进 sub/（外部有直接调用的保留 standalone，见 CLAUDE.md 独立入口表 + `~/Dev/tools/cc-home/skills` 里的引用） |
| P4 | **strip_\*/audit_\* 家族 main() 样板统一**（各 ~50 行重复的 argparse/backup/report 骨架） | 小 | 抽 `sub/_cli_common.py`；surface 指纹不得变 |
| P5 | **docx_revise 扩 op**（表格加列/数据系列延展等） | 按需 | **需求驱动，别预建**——qual-supply 下一轮真用到再加，加时先写单测 |
| P6 | `fix_styleset` 包内相对 import 重构（2007 行，`load_step` 加载不到） | 大 | 收益低，只在要接 spec 时做 |

**不做清单（已核实，别再查一遍）**：`typeset_pipeline` vs `typeset_apply` 不合并（韧性 driver vs
声明式引擎，职责不同）；`docx_apply_template` vs `docx_format_clone` 不重复（直排版问题，
format_clone docstring 有分工说明）；四对疑似重复血统均非重复；34 个无 pipeline 接口的脚本
各有理由（`before-after.md` §四表）。

## 六 · 停止线

- surface diff 出现预期之外的变化 → 停，先解释每一条再继续。
- 标书门检（bid_\*）、交付件相关改动没有真件回归 → 不许 commit。
- 想退役/删除任何在用功能 → 停，AskUserQuestion。
- 同一处第 3 个 bug → 停手找缺的那根轴（铁律 #11 下半条）。
