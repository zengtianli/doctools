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
| P2 | ~~**`docx_tools.py` 1705 行单体拆解**~~ ✅ 2026-07-31 完成：三段各落 `sub/docx_{extract,check,track}.py` 实现模块（自带 main() 可独立敲，argparse 声明在各自 `add_*_parser()` 只写一遍），`docx_tools.py` 收薄为 372 行组合入口（batch 层 + 组合 argparse + library re-export，兼容外部 `from docx_tools import extract_paragraphs` 用法）。**原路线「+ `_groups` 行」经证据核实不可行已放弃**（主会话拍板）：extract/check/track 在 docx_cli 走 CMD_TABLE fast-path（带 read/diff aliases 与 snapshot/compare 前置 token 注入），`_groups` 的 Opt/Target 模型两者都装不下，强行进表 surface 必变。`docx_cli` 一行未动、转发仍指 docx_tools。顺手修复：docx_tools 裸跑 `ModuleNotFoundError: file_ops`（doctools/lib/file_ops.py 2026-05-21 已删、canonical 在 dev/lib，docx_cli 路径因先插 dev/lib 一直掩盖着）——`sub/docx_track.py` 自带 dev/lib append。工单 = 会话 scratchpad order-p2.json | 中 | 已验收：surface/probe diff 空 · pytest 82 · script_graph 0 孤儿 · collar 36 全挂 · 端到端对拍 39 采样（extract 裸/--info/--json/--split-chapters/-o、snapshot、compare、tc read md/json、tc review 含 strict-miss/include-ins 语义、docx_cli read/diff/snapshot/track、--batch 3 行 JSONL、library import）改前改后逐字一致，review 输出 docx 剥 w:date 后逐部件相同 |
| P3 | ~~**docx_cli legacy 层收编评估**~~ ✅ 2026-07-31 完成：10 个逐一评估，**8 保留 standalone**（docx_tools / docx_apply_template / docx_text_formatter / md_docx_template / docx_renumber_figures / docx_format_clone / docx_apply_image_caption 有铁证级外部调用方；md_tools 是 doc_dispatch 6 处 subprocess 的活契约，用户拍板保留），**2 收编进 sub/**（fix_superscript_refs + scan_sensitive_words，全生态 grep 无 standalone 活调用；docx_cli 转发目标改 `sub/<stem>`，CMD_TABLE 与声明段一字未动，surface/probe diff 均空）。工单 = 会话 scratchpad order-p3.json | 大 | 已验收：surface/probe diff 空 · pytest 80→82（补 fix_superscript_refs apply() 两测）· script_graph 0 孤儿 · collar 全挂 · 真件 typeset --dry-run 报告与改前一致 |
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
