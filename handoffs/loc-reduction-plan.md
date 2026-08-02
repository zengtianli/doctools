# doctools 代码总量还能降多少 · 方案 · 2026-08-02

> 冷启动自包含。仓库 `/Users/tianli/Dev/tools/doctools`，所有命令以此为 cwd。
> 本轮**只出方案不落地**。产出方式：5 镜头并行扫 + 7 条对抗核验（存活 6 / 被否 1）+ 主会话独立复核。

## 零 · 先说结论，含不好听的那部分

**还能净删约 300 行，占全仓 0.6%。文件数一个不减（做完仍是 99 个脚本）。**

这个数字小得配不上"重构"两个字。所以先把**天花板**摆出来，免得下次有人再来问一遍：

我写了个 ast 克隆检测器（把标识符与字面量抹平只留结构骨架）扫全仓 1329 个 ≥6 行的函数
（脚本 `handoffs/_clone_scan.py`，可复跑）：

| 口径 | 克隆组 | 理论净删上限 | 占全仓 53,643 行 |
|---|---:|---:|---:|
| 逐字节相同的函数 | 5 组 | **47 行** | 0.09% |
| 结构相同（只差数据）的函数 | 18 组 | **421 行** | 0.79% |

**这个上限的边界必须说清**：它只看得见「整个函数是克隆」，看不见 ① 函数内部片段重复
② if/elif 分支族（每个分支不是函数）③ 多一个 if 就判为结构不同的近似克隆。
所以 421 是**复制粘贴这一类的下界**，不是全部空间 —— 但五个镜头连着扫完，
真正过了对抗核验的也就 300 行出头，说明剩下的空间确实不在这一类里。

**真正的价值不在行数，在它顺手修掉的判据分歧**（W2 那条：同一个"文件被占用吗"有 4 种语义）。
如果只盯着行数看，这三条都不值得做。

---

## 一 · 施工单（三件事，按 ROI 排序，各自独立可交付、独立 commit）

顺序 **W1 → W3 → W2**：W1 会给 `blocks.py` 补上 `_cli_common` 的 bootstrap，W2 可直接复用不重复加。

### W1 · 12 个家族 `main()` 分发器 → `_cli_common.family_main()`

| 项 | 内容 |
|---|---|
| 抽什么 | `sub/` 下 12 个家族文件里**结构逐字相同**的 19 行子命令分发器（argv 拆分 → `SUBCOMMANDS` 查表 → `sys.argv` 保存/替换/还原 → rc 归一化） |
| 从哪抽 | `sub/{audit,biddiff,blocks,caption,chapter,freeze,image,renumber,split,strip,table,typeset_ops}.py` |
| 净删 | **~150 行**（两轮独立核验报 161 / 154；按本仓 4 行 bootstrap 约定取 142 为下界） |
| 落到哪 | `sub/_cli_common.py` 追加 `family_main(subcommands, argv=None, *, file=__file__, usage_args="<args…>")` 约 30 行 |
| 现场枚举 | `grep -rln '^SUBCOMMANDS' scripts/document/sub/` —— 落笔时正好 12 个，**没有第 13 个** |

**差异只有两处**（主会话独立复核过：把 12 段 `main()` 抽出做归一 diff，只有 usage 尾巴一行不同）：

1. usage/报错串里的族名 —— 12/12 等于文件 stem，`Path(file).stem` 派生
2. usage 尾巴：`<docx> [flags…]`（audit / freeze / strip 三个）vs `<args…>`（其余 9 个）

**弄反第 2 条，`--help` 输出就变了，而五道闸门照样全绿。**

⚠ **`cli_surface` 与 `cli_forward_probe` 在这一层是瞎的** —— probe 把 `exec_script` 换成录音机，
根本进不到家族 `main()`。只跑那两道就宣称绿 = 假绿。**必须跑 §2.1 的 144 例 CLI 对拍。**

风险，按先炸顺序：

| # | 会炸什么 |
|---|---|
| 1 | **bootstrap 插入位置**：`typeset_ops.py` / `image.py` 的 `_sys` / `_Path` 别名定义在 `import argparse` 块**之后**十几行，锚错就当场 `NameError`。必须锚在各文件已有的 `_sys.path.append(... / "lib")` 那一行之后。`biddiff.py` 连别名都没有，只能用裸 `sys` / `Path`，写法与其余不同 |
| 2 | **`sub/` 会被多推进 6 个模块的 sys.path**。`lib/styles.py` 与 `sub/styles.py` 同名 —— 今天安全（`grep -rn '^ *import styles\|^ *from styles' scripts/ lib/ tools/` 零命中，三处要 `lib/styles` 的都走 `spec_from_file_location`），但这是本条唯一扩大的攻击面，**必须写进 `family_main` 的 docstring** |
| 3 | 族名从 `__file__` 派生 = 文件改名会连带改 usage 串，同样写进 docstring |
| 4 | traceback 多一层 `family_main` 帧。已知无人按栈形状断言（202 个测试零命中），但 `blocks.py` 的 `reorder` / `fix-heading-disorder` **本来就在入口引爆** `ModuleNotFoundError: No module named 'apply_body_styles'`（既存 bug，与本条无关），对拍时会看到这两例只差栈帧 —— 属预期内 |
| 5 | `_cli_common` 铁则：**禁 import docx、禁 import 期副作用**（`check_docx_collar` 枚举全仓，惊动它 = 假红）。`family_main` 只用 `sys` + `pathlib`，不违 |

---

### W3 · `fix_styleset.py` 8 处全零 `allowed_deltas` → 整块删

排在 W2 之前是因为它单文件、纯删除、零 import 管线、零跨模块。

| 项 | 内容 |
|---|---|
| 抽什么 | 8 处 `allowed_deltas={…全 0…}` kwarg。`sub/shape_contract.py` 三个分支的缺省全是 0，所以「写 0」≡「不写」 |
| 从哪抽 | `fix_styleset.py` 行 434 / 594 / 689 / 822 / 884 / 1001 / 1138（7 处独立 kwarg，整段删）+ 1679（内联，塌成 `violations = diff_structure(before, after)`）。**行号以现场 grep 为准** |
| 净删 | **67 行** |
| 落到哪 | 无新模块，纯删除 |

**主会话独立复核（ast 实测，不是目测）**：8 处全部 `literal_eval` 通过，非零值 **一个都没有**，
键数 5/6/7 三种形状；`shape_contract.py:202/218/226/247` 四处缺省确认全是 `allowed.get(f, 0)`。

⚠ **202 个 pytest 里一条都没覆盖 `shape_contract` / `fix_styleset`** —— 删错了套件照样 202 passed。
必须跑 §2.3 的 24000 例等价对拍。

**必须同一个 commit 一起做的三件事**，否则这条等于用 67 行死代码换 3 行新死代码 + 一段说谎的 docstring：

1. 删完后 `_shape_gate(...)` 的 `allowed_deltas` 形参**零调用者**（7 个调用点一个不落全传）→ 顺手摘掉形参
2. `fix_styleset.py` 模块 docstring 第 22–27 行逐子命令写着「style-rebrand: 全 0 / style-pool-cleanup: 段数 0 …」，删完就成了描述不存在的配置 → 一并改
3. **补一条钉住缺省的回归断言**（约 12 行，落 `scripts/document/tests/`）：断言 `diff_structure` 在「字段未列入 `allowed_deltas`」时容差为 0。
   这 8 处原来把「零容忍」显式写在调用点，删掉后全靠隐式缺省兜；哪天有人把缺省放松成「未列入 = 不检查」，
   8 道 gate 会**静默从 fail-closed 变 fail-open**，调用点没有任何本地证据能提示。
   **没有这条断言不许合这个 commit。**

---

### W2 · `lsof` 占用检测 + `.bak-N-日期` 备份路径 → 收编到已有的 `_cli_common`

| 项 | 内容 |
|---|---|
| 净删 | **~100 行**（镜像实测 113，按 4 行 bootstrap 约定 −14） |
| 落到哪 | `sub/_cli_common.py`（`lsof_check` / `find_next_backup` 已存在，只新增 `lsof_check_lines`） |

#### 这条真正的价值：`lsof_check` 有 10 份副本，分成 4 种语义

主会话独立复核（`ast.unparse` 逐份 dump 后比对），**没有任何两份逐字节相同**：

| 变体 | 份数 | 实现 | 落位 |
|---|---:|---|---|
| **A** | 5 | `lsof <path>`，rc==0 且 stdout 非空即判占用，返 **raw stdout（带尾换行）** | `fix_styleset:87` `outline:65` `slim:112` `styles:83` `pipeline_lib:86` → 指向 `_cc.lsof_check` |
| **B** | 3 | `lsof -- <path>`（多一个 `--`），且要求 `len(lines) > 1` 才算占用，返 **已 strip 的 join** | `add_header_footer:52` `blocks:1335` `caption:393` → 新增 `_cc.lsof_check_lines()` |
| **C** | 1 | 返 **bool**、`except Exception` 全吞、**不看 returncode** | `caption.py:181 lsof_check_bool` → **不动** |
| **D** | 1 | **无 timeout** | `pptx_cli.py:867 _lsof_guard` → **不动** |

⚠ **现存的那个"正统" SSOT（`_cli_common.py:35`）恰好是较弱的变体 A**：没有 `--`，
挡不住以 `-` 开头的文件名。B 严格更健壮。

**但本方案不合并 A 与 B**，理由是行为保全：合并会丢掉 B 的「只有表头就当空闲」那道守卫，
并给 `[error] docx 被进程占用:\n{lsof}` 多一个尾换行。
**「给 A 补上 `--`」是一次独立的行为变更，该单独占一个 commit、单独写理由，不许搭在这次收编里** ——
这正是本仓「破坏性动作必须自己占一个动词」那条铁则的同一形状。

#### 明确不动的范围（越界即中止）

| 不动 | 为什么 |
|---|---|
| `caption.py:181 lsof_check_bool` | 变体 C。指向 `_cc.lsof_check` 会让 `caption.py:323` 的占用判定**直接翻向**，且自动覆盖率为 0 |
| `pptx_cli.py:867 _lsof_guard` | 变体 D，无 timeout |
| `image.py:890` 的同形内联 `.bak-{n}-{date}` | 在 `apply_patch` 的 backup 布尔分支里，形状不同且**未经等价核验**，本方案不认领 |
| 全仓一大票 `.bak-<时间戳>` 写法（`table:856/1056` `normalize_fonts:162` `md_merge_impl:259` `docx_fmt:890` `bid_gate:217/1175`…） | **另一套命名约定**，不是同一个机制，扫进来就是行为变更 |

#### 先炸点（主会话已复核）

`pipeline_lib.py:661` 的 `__all__` 列着 `lsof_check` / `make_backup_path`，被
`scripts/document/typeset_apply.py:98` 直接 `from sub.pipeline_lib import load_step, lsof_check, make_backup_path`，
qual-supply 那边还有个 `from sub.pipeline_lib import *` 的 shim。
**必须留模块级别名/委派 def**（仓里 `caption.py:190` / `renumber.py:317` 已是这个形状）—— 删名字这两处当场 ImportError，
而 `cli_surface` / `probe` / pytest **全都看不见**。

同 commit 必改：`_cli_common.py:10-11` 白纸黑字写着「语义抄 canonical 实现 `pipeline_lib.lsof_check`（不 import 它）」——
方向反转后这句话变成谎话，下一个人会照它去 `pipeline_lib` 找真身。

顺手清理：两函数删掉后 `import subprocess`（4 处）与 `from datetime import date`（outline / pipeline_lib）成死 import，
清掉再多删 ~6 行。

---

## 二 · 验收命令（照抄能跑）

### 2.0 基线（每条开工前跑一次，产物留着）

```bash
cd /Users/tianli/Dev/tools/doctools
mkdir -p /tmp/doctools-gate
python3 tools/cli_surface.py              > /tmp/doctools-gate/surface.before.json
python3 tools/cli_forward_probe.py --json > /tmp/doctools-gate/probe.before.json
python3 -m pytest scripts/document/tests scripts/document/sub/tests -q | tail -1
```

### 2.1 收尾五道门（每条施工完都要全绿）

```bash
cd /Users/tianli/Dev/tools/doctools
python3 tools/cli_surface.py > /tmp/doctools-gate/surface.after.json
diff /tmp/doctools-gate/surface.{before,after}.json >/dev/null 2>&1; echo "surface diff rc=$?"   # 必须 0
python3 tools/cli_forward_probe.py  >/dev/null 2>&1; echo "probe=$?"      # 必须 0
python3 tools/check_docx_collar.py  >/dev/null 2>&1; echo "collar=$?"     # 必须 0
python3 tools/check_function_axis.py >/dev/null 2>&1; echo "axis=$?"      # 必须 0
python3 tools/check_external_refs.py >/dev/null 2>&1; echo "xref=$?"      # 必须 0
python3 -m pytest scripts/document/tests scripts/document/sub/tests -q | tail -1   # 202 passed
```

### 2.2 W1 专用：144 例家族 CLI 对拍（12 族 × 12 例）

闸门在这一层是瞎的，这是唯一能抓住 W1 回归的东西。

```bash
cd /Users/tianli/Dev/tools/doctools
python3 handoffs/_loc_plan_harness/family_ab.py "$PWD" /tmp/doctools-gate/fam.before.json
#  ……施工……
python3 handoffs/_loc_plan_harness/family_ab.py "$PWD" /tmp/doctools-gate/fam.after.json
diff /tmp/doctools-gate/fam.{before,after}.json >/dev/null 2>&1; echo "family diff rc=$?"
```

判据：`diff` 必须**完全空**（norm 已抹掉绝对路径 / traceback 行号 / 对象地址）。
栈帧多一层 `_cli_common.py in family_main` 那一行仍会显出来 —— **只允许这一种差异，
且仅限 `blocks|['reorder']` 那类本来就崩的既存 case，rc / 异常类型 / 异常文本必须一字不差**。

补一条真调用路径（`docx_cli` 走 `_dispatch` 的 `spec_from_file_location` 加载，族名不能退化成 alias 名）：

```bash
cd /Users/tianli/Dev/tools/doctools/scripts/document/sub
python3 -c "import _dispatch; [print(f, _dispatch.exec_script(f,['bogus'])) for f in ['chapter','audit','image','strip']]"
# 期望：打印 [chapter]/[audit]/[image]/[strip] unknown subcommand… 且 rc 全 2
```

### 2.3 W2 / W3 专用对拍

```bash
cd /Users/tianli/Dev/tools/doctools
python3 handoffs/_loc_plan_harness/lsof_backup_ab.py      # rc=0 即三项全过
python3 handoffs/_loc_plan_harness/allowed_deltas_ab.py   # rc=0 即 24000 例全等价
python3 -c "import sys; sys.path.insert(0,'scripts/document'); \
  from sub.pipeline_lib import load_step, lsof_check, make_backup_path; print('import OK')"
```

## 二·五 · 四个对拍脚本**已实跑通过**，不是「看着能跑」

| 脚本 | 落位 | 实跑输出 |
|---|---|---|
| `family_ab.py` | `handoffs/_loc_plan_harness/` | `cases=144` |
| `lsof_backup_ab.py` | 同上 | `backup mismatches = 0` · `A distinct = 1 \| B distinct = 1` · `A.strip() == B : True` · ✓ 全过 |
| `allowed_deltas_ab.py` | 同上 | `comparisons = 24000 · differences = 0` |
| `_clone_scan.py` | `handoffs/` | `函数总数 1329 · 克隆组 18 · 理论净删上限 421 行` |

⚠ 两个坑已经踩过并写进脚本注释：

1. `lsof_backup_ab` 必须按 `sub.<name>` **包路径**导入 —— 用 `spec_from_file_location`
   会因 `fix_styleset.py` 的相对导入（`from .shape_contract import …`）当场
   `ImportError: attempted relative import with no known parent package`。
2. `allowed_deltas_ab` 必须 import 生产的 `shape_contract`，禁重写替身（铁律 #2）。

---

## 三 · 停止线

| 条件 | 动作 |
|---|---|
| 实测净删 < 估计值的一半 | 放弃该条，把实测数写回本文件的被否节 |
| §2.2 的 144 例对拍出现**任何**非预期差异 | 立即回滚该条，不许"看着差不多"放行 |
| W2 触碰到变体 C / D，或扫进 `.bak-<时间戳>` 那一套 | 立即中止 —— 那是行为变更，不是收编 |
| W3 没有先补「缺省容差为 0」的回归断言 | 不许合 commit |
| 任一条需要改 `cli_surface` 指纹才能过 | 立即中止 —— 接口不变是本轮的前提，不是可谈判项 |

---

## 四 · 被否节（价值不低于施工单 —— 它防止下次有人再提一遍）

### 4.1 「`_common_setup` + `_save_with_backup` + `_emit_report` 写盘三连抽成公共函数」— 否

声称净删 69 行，实测 **14 行**，且有一条致命理由：

**抽走 `_save_with_backup` 会把 `styles.py` 和 `outline.py` 从 surgical 收口守卫的名册上摘掉。**

主会话独立复核：

```
grep -c '\.save(' sub/styles.py   → 1     ← 就在 _save_with_backup 里面
grep -c '\.save(' sub/outline.py  → 1     ← 同上
```

`check_docx_collar` 的判据是 `IMPORTS_DOCX and ".save(" in src`。
把这唯一一处 `.save(` 抽走 → 两个文件**整体退出守卫名册**，
fail-closed 静默变 fail-open，而五道闸门照样全绿。

> **这是一条结构性教训，不只是否掉一个候选**：现行 collar 判据把「文件里有没有 `.save(`」
> 当作「这个文件要不要挂收口」。于是**任何把存盘逻辑抽成公共函数的重构，都会顺手解除该文件的守卫**。
> 这一类重构在守卫判据改掉之前，全部不安全。要动这一层，**先改守卫（改成跟着调用链走，
> 而不是跟着字面量走），再谈抽取**。

### 4.2 其余 11 条未进核验的候选

5 个镜头共报 18 条，去重排序后只核验了净删 > 0 的前 7 条。剩下 11 条**净删 ≤ 0 或排序靠后**，
**没有经过对抗核验，不得当作已验证结论引用**。它们的共同形状是「省下的行数抵不过新增的参数与 import」。

---

## 五 · 诚实的总账

| | 数 |
|---|---:|
| 全仓（scripts + lib + tools，含测试） | 53,643 行 |
| 三条全做完净删 | **~317 行** |
| 占比 | **0.59%** |
| 脚本数变化 | **0**（99 → 99） |
| 新增文件 | **0**（全部落进已有的 `_cli_common.py`） |

**做不做，取决于你怎么看这 317 行。**

- 只看行数：不值得。0.6% 的收益换三次带对拍的施工。
- 看判据分歧：值得。W2 顺手把「文件被占用吗」的 4 种语义收敛掉，W3 顺手把
  8 处「静默 fail-open 的入口」钉死成一条断言。这两样是 bug 源，行数只是副产品。
- **最有价值的那条其实是被否的 4.1** —— 它暴露了 collar 守卫的判据缺陷：
  「抽公共存盘函数会自动解除守卫」。这个缺陷现在没人知道，也没有任何闸门会报。
  **建议优先级高于上面三条施工项。**
