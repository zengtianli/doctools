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
6. 测试 **80 绿**（`scripts/document/tests` + `sub/tests`）；`script_graph` 0 孤儿。（2026-07-31 P3 补测后 = **82 绿**）
7. **collar 第二判据存量清零**（2026-07-31，81e02e0）：25 个 zipfile 写盘脚本按四类接线
   （19 A-preserve / 3 B-diff_parts+声明删除集 / 1 C-对模板基线 / 2 D-engine 一处覆盖全调用方），
   每处 stub 留痕 + 真跑无假红 + 反向注入全拦；`check_docx_collar` **默认已翻全仓扫描**
   （`--all` 降为兼容 no-op），裸 ZipFile 探针反向验证即红。分类工单 = 会话 scratchpad collar-classify.json。

## 三 · 本项目铁则（每改必守）

- **两道闸门每改必跑**：改前 dump surface/probe 基线，改后 diff **必须恰好等于预期变更**，零意外。
- 加子命令 = `_groups.py` 加一行数据；加独立入口脚本 = `CLAUDE.md` 独立入口表加一行。
- 版面改动进 spec yaml、内容修订进 ops yaml——**禁在项目/会话里现编注入脚本**。
- 守卫/自检 fail-closed（空集不报绿），写完立刻反向验证。
- 退役任何功能 = 用户拍板；删除走 `~/.Trash/<slug>-<date>/` 带 MANIFEST。

## 四 · 验收命令块（可复制，改完全跑）

```bash
cd ~/Dev/tools/doctools
python3 -m pytest scripts/document/tests scripts/document/sub/tests -q   # 82 passed
python3 tools/cli_surface.py > /tmp/surface.json                          # 与改前 diff
python3 tools/cli_forward_probe.py > /tmp/probe.json                      # 67 条全正常(内嵌预期 argv,改错转发即红)
python3 tools/script_graph.py                                             # 0 个没人引用
python3 tools/check_docx_collar.py                                        # 收口全挂
```

## 五 · 待改造 backlog（按优先级；判断已下过，别重推导）

| # | 事 | 体量 | 已下的判断 / 验收标准 |
|---|---|---|---|
| P1 | ~~**bid_\* 家族 7 文件 → 1**~~ ✅ 2026-07-31 完成：6 入口合并为 `bid_gate.py` 子命令族（`run`=driver 直调四门 / `scan`/`sweep`/`identity`/`print` 单门 / `deref`），`bid_residue_lib.py` 零改动保持 SSOT；identity/print 的内嵌 `sys.exit` 全改 return；banner 保留旧脚本名字面量（逐字验收的钦定代价）；`sweep --apply` 与 `deref` 写盘挂 `assert_parts_intact`（反向验证过）。旧 6 件 → `~/.Trash/bid-gate-merge-20260731/`（MANIFEST+sha256）。**全生态指针一次改完**：work_ops（BID_FINAL+argv 加 run token）/ hq_capabilities / doc_dispatch / bids 顶层+项目层+_template（FACTS tpl、rules yaml、SOP）/ skills×3 / tl-report / cockpit；oss vendored 未碰。工单 = 会话 scratchpad order-p1.json | 中 | 已验收：真件（wenzhou 飞-zlt727）main exit 0 / pei exit 2 两基线 + 单门 5 场景 + doc_dispatch bidfinal 链路，新旧引擎 stdout **逐字节相同**；surface/probe diff 空；三向对抗核验后项目层残留 8 处已清（3c17750） |
| P2 | ~~**`docx_tools.py` 1705 行单体拆解**~~ ✅ 2026-07-31 完成：三段各落 `sub/docx_{extract,check,track}.py` 实现模块（自带 main() 可独立敲，argparse 声明在各自 `add_*_parser()` 只写一遍），`docx_tools.py` 收薄为 372 行组合入口（batch 层 + 组合 argparse + library re-export，兼容外部 `from docx_tools import extract_paragraphs` 用法）。**原路线「+ `_groups` 行」经证据核实不可行已放弃**（主会话拍板）：extract/check/track 在 docx_cli 走 CMD_TABLE fast-path（带 read/diff aliases 与 snapshot/compare 前置 token 注入），`_groups` 的 Opt/Target 模型两者都装不下，强行进表 surface 必变。`docx_cli` 一行未动、转发仍指 docx_tools。顺手修复：docx_tools 裸跑 `ModuleNotFoundError: file_ops`（doctools/lib/file_ops.py 2026-05-21 已删、canonical 在 dev/lib，docx_cli 路径因先插 dev/lib 一直掩盖着）——`sub/docx_track.py` 自带 dev/lib append。工单 = 会话 scratchpad order-p2.json | 中 | 已验收：surface/probe diff 空 · pytest 82 · script_graph 0 孤儿 · collar 36 全挂 · 端到端对拍 39 采样（extract 裸/--info/--json/--split-chapters/-o、snapshot、compare、tc read md/json、tc review 含 strict-miss/include-ins 语义、docx_cli read/diff/snapshot/track、--batch 3 行 JSONL、library import）改前改后逐字一致，review 输出 docx 剥 w:date 后逐部件相同 |
| P3 | ~~**docx_cli legacy 层收编评估**~~ ✅ 2026-07-31 完成：10 个逐一评估，**8 保留 standalone**（docx_tools / docx_apply_template / docx_text_formatter / md_docx_template / docx_renumber_figures / docx_format_clone / docx_apply_image_caption 有铁证级外部调用方；md_tools 是 doc_dispatch 6 处 subprocess 的活契约，用户拍板保留），**2 收编进 sub/**（fix_superscript_refs + scan_sensitive_words，全生态 grep 无 standalone 活调用；docx_cli 转发目标改 `sub/<stem>`，CMD_TABLE 与声明段一字未动，surface/probe diff 均空）。工单 = 会话 scratchpad order-p3.json | 大 | 已验收：surface/probe diff 空 · pytest 80→82（补 fix_superscript_refs apply() 两测）· script_graph 0 孤儿 · collar 全挂 · 真件 typeset --dry-run 报告与改前一致 |
| P4 | ~~**strip_\*/audit_\* 家族 main() 样板统一**~~ ✅ 2026-07-31 完成：抽 `sub/_cli_common.py`（111 行：argparse 骨架/backup/report/lsof 占用检查），strip_×5 + audit_×5 收编（`audit_styleset` 排除——自带 register() 多 target 结构，样板重合度低）；`import _cli_common` 带 self-dir append（`_dispatch._load` 不进 sys.path，漏了会静默 rc=2）；退出码不一致**保持现状**（消费点未核，不顺手统一）。工单 = 会话 scratchpad order-p4.json | 小 | 已验收：surface/probe diff 空 · 每脚本 `--help` 逐字不变 · slim 整链真跑（72 部件 sha256 全同，仅 zip mtime 噪音）· collar 绿 |
| P5 | **docx_revise 扩 op**（表格加列/数据系列延展等） | 按需 | **需求驱动，别预建**——qual-supply 下一轮真用到再加，加时先写单测 |
| P6 | `fix_styleset` 包内相对 import 重构（2007 行，`load_step` 加载不到） | 大 | 收益低，只在要接 spec 时做 |
| P7 | ~~**pdf 族 3 文件 → 1**~~ ✅ 2026-07-31 完成：`pdf_pipeline_lib.py`（纯单消费方，pipeline 引擎段整段）+ `pdf_to_docx.py`（to-docx 引擎段，python-docx/docx_safe_save 改惰性 `_ensure_docx_deps()`，read 等子命令实测不再加载 python-docx）并入 `pdf_cli.py`（755→1971 行，bid_gate 单文件范式）；CLI surface 零改动。banner 主会话拍板 A：`convert to-docx` stdout 统一为原 pdf_to_docx 中文格式（doc_dispatch 机器链逐字节保真，pdf_cli convert 英文 banner 仅人眼消费承担 cosmetic 变化）。真重复去重 3 处（pdfimages 包装 ×3→`_run_pdfimages` / 3072 阈值→`MIN_IMAGE_BYTES` / sanitize 共底 `_sanitize_core`；outline 遍历 ×2 按工单不动）。外部指针：doc_dispatch:263 argv / hq_capabilities / Apps doc-tools catalog / wiki memory 两副本。旧 2 件 → `~/.Trash/consolidation-20260731/pdf/`（MANIFEST+sha256）。extract image 无噪音过滤/img-NNN 命名违反铁律 #6 = 已知现状，记 backlog 不借折叠改行为（OQ2）。工单 = 会话 scratchpad cons-order-pdf.json | 中 | 已验收：24 命令 stdout/stderr/rc 对拍（14-pipepar 真跑 spawn 并行）+ convert/dispatch 产物逐部件 sha256 全同 + 收口惰性 import 过守卫 regex 且反向验证转红 |
| P8 | ~~**sub/ 12 族折叠**~~ ✅ 2026-07-31 完成：strip×7/audit×6/freeze×2/image×3/table×4/caption×2/chapter×3/renumber×2/blocks×3/split×2/typeset_ops×5/biddiff×3 共 42 旧件 → 12 个子命令族文件（sub/ 72→43）；`_groups` Target 加 impl_argv、`_dispatch.exec_script` 加 RLock（修 health 线程池 sys.argv 竞态——折叠前就在静默产错诊断，真 bug 修复）；typeset_apply Action 加 script/entry/tail_entry 三字段（spec 词汇 SSOT 不动）。工单 = cons-order-subf.json | 大 | surface 逐字节空 · probe diff 恰=映射 9+N 条 · 28 条 standalone 对拍 26 逐字节同/2 条仅迭代序与 SyntaxWarning 噪音 · slim/typeset 整链部件 sha 全同 |
| P9 | ~~**入口散件 9→2+1**~~ ✅ 2026-07-31 完成：renum.py（chapter/tabfig/figures）+ docx_fmt.py（template/clone/fonts/text）+ image_caption 平移 sub/；chart/write_gate/docx_tools/四大引擎/dispatch/GUI 保留。46 条对拍逐字节全清。工单 = cons-order-dent.json | 中 | 同上闸门 + govern-ledger 3 条机检指向顺手修 |
| P10 | ~~**pptx 4→1**~~ ✅ 2026-07-31 完成：pptx_cli.py 单骨架（顶层 stdlib-only 双解释器契约改写为按子命令），32 用例对拍；to-md 的死 flag `-o`（声明但恒忽略）删除=从静默忽略变 argparse 硬错，**主会话追认**（诚实报错优于假接受）。工单 = cons-order-pptx.json | 中 | 13 子命令 stdout/exit/产物对拍 + /usr/bin/python3 stdlib 契约实测 |
| P11 | ~~**md 3→2**~~ ✅ 2026-07-31 完成：md_docx_template 并入 md_tools 子命令 md2docx；md_to_audiobook 保留（PEP-723 依赖隔离实测为真）。`to-docx`(pandoc 版) 保留原样——**已知它绕开 docx-gen-guard 字面匹配，属守卫旁路，建议下轮用户拍板退役**。工单 = cons-order-md.json | 小 | 7 条基线对拍全绿 |
| P12 | ~~**折叠断链修复轮**~~ ✅ 2026-07-31 三向核验抓出后当场修：qual-supply 23 shim + bid-diff 5 shim 重写为 **import 透明型**（spec 直载家族文件+公有名别名回旧名；初版 execv 会被 run_pipeline 的 importlib 一进就逃逸，当日发现当日改；验证=28/28 import 不逃逸 + 两处 run_pipeline 端到端 rc=0；fix_heading_disorder/reorder 两条 2026-05-26 起依赖已死，忠实保留断态）；bids 4 驱动脚本/eco-flow report_gate/yazi keymap/panan split.sh 改指；collar 第一判据正则精修（docx_parts 不再误判为 python-docx，双向反向验证）；cli_forward_probe 重建为 67 条内嵌预期 argv（录音机→比对器，反向验证：改错转发必红） | 中 | 每处真跑 --help/干跑验证 + 全生态旧名 grep 复核 |

**不做清单（已核实，别再查一遍）**：`typeset_pipeline` vs `typeset_apply` 不合并（韧性 driver vs
声明式引擎，职责不同）；`docx_apply_template` vs `docx_format_clone` 不重复（直排版问题，
format_clone docstring 有分工说明）；四对疑似重复血统均非重复；34 个无 pipeline 接口的脚本
各有理由（`before-after.md` §四表）。

## 六 · 停止线

- surface diff 出现预期之外的变化 → 停，先解释每一条再继续。
- 标书门检（bid_\*）、交付件相关改动没有真件回归 → 不许 commit。
- 想退役/删除任何在用功能 → 停，AskUserQuestion。
- 同一处第 3 个 bug → 停手找缺的那根轴（铁律 #11 下半条）。
