---
title: doctools 现状 · 99 个脚本 / 93 条动词 / 6 道闸门
slug: doctools-status
project: doctools
date: 2026-08-02
kpi:
  - label: 脚本
    value: "99"
    note: 工具本体 89 + 测试 10
  - label: 可跑动词
    value: "93"
    note: CLI 面 126 节点去别名后
  - label: 只碰版式的动词
    value: "41"
    note: 占 93 条的 44%
  - label: 闸门
    value: "6 / 6 绿"
    note: 本页数字全部来自刚跑的这 6 条命令
---

## 一、脚本数：99

| 层 | 数 | 是什么 |
|---|---|---|
| `scripts/document/` | 17 | 入口脚本（`docx_cli` 总入口 + 16 个直敲入口） |
| `scripts/document/sub/` | 44 | 子命令实现（含 `_groups` 声明表 + `_function_axis` 职能表） |
| `scripts/data/` | 5 | 数据转换（xlsx 族 + convert） |
| `lib/` | 16 | 公共库（收口 / 手术引擎 / 元素级遍历 / 判据 SSOT） |
| `tools/` | 7 | 闸门与图（collar · surface · probe · axis · xref · graph · blast_radius） |
| `tests/` ×2 | 10 | 9 个测试文件 + 1 个 `__init__.py`，共 202 个用例 |

另有 `handoffs/_inventory.py`（盘点脚本自身），不计入 99，`script_graph` 也不数它。

## 二、CLI 面：126 条子命令 = 93 条真动词 + 8 别名 + 组父节点

**三个数都对，报的时候必须说明口径**：`cli_surface` 数的是节点总数 **126**（含 `audit` / `strip` 这类组父节点）；叶子路径 101；按 `id(parser)` 去重后 **93** 条可跑动词，其余 8 条是别名（`read`=extract · `diff`=compare · `styleset *`=audit-styleset *）。

### 职能轴分布

| 职能 | 条数 | 碰的是什么 |
|---|---|---|
| `format` | 41 | 只碰版式：样式 / 字体 / 段落属性 / 页眉页脚 / 分节 / 编号与题注号 / 大纲层级 / 表格边框 / 装帧 / 样式集 |
| `inspect` | 28 | 只读检查审计，不写盘 |
| `content` | 15 | 碰正文写了什么：删段 / 改文字 / 合并拆分 / 回写章节 / 挪段落块 / 资源重挂 |
| `review` | 6 | 审阅修订：`w:ins` / `w:del` / 批注 / 版本对照 |
| `convert` | 2 | 跨格式转换 |
| `dispatch` | 1 | 调度编排 |

**「只处理格式的脚本有哪些」现在是一条命令查得出的事实，不是一份会过期的名单：**

```bash
python3 scripts/document/docx_cli.py verbs --fn format   # 41 条
python3 scripts/document/docx_cli.py verbs               # 全部 93 条按 6 档分组
```

职能只有这 6 个值，别自创第七个。表在 `sub/_function_axis.py`，`tools/check_function_axis.py` 双向对账——**新加子命令不打标签判红，表里留了 CLI 已没有的条目同样判红**。

## 三、闸门状态：6 道全绿（本页数字即由它们产出）

| 闸门 | 管什么 | 刚跑的结果 |
|---|---|---|
| `check_docx_collar.py` | python-docx 存盘必挂 surgical 收口；zipfile 路线必挂部件完整性断言 | ✓ 23/23 · 16/16 |
| `check_function_axis.py` | 职能轴表 ↔ CLI 双向对账 | ✓ 93 条全有标签 |
| `cli_forward_probe.py` | 每条子命令**真正转发出去的 argv** | ✓ `bad: []`（67 条内嵌预期） |
| `check_external_refs.py` | 全生态引用存在性（已挂 pre-commit `--changed-only`） | ✓ 219 处 / 39 条不同路径全部存在 |
| `pytest` | 回归 | ✓ 202 passed |
| `script_graph.py` | 谁调谁 · 孤儿判定 | ✓ 99 脚本 · 362 引用 · **0 孤儿** |

`cli_surface.py` 是第 7 条，但它是**指纹比对器**（改前改后 diff 必须为空），不是能单独判绿的门。

## 四、两处与文档的口径差，已核实非问题

| 差在哪 | 真相 |
|---|---|
| README 写「222 处仓外引用」，现跑 219 | 2026-08-01 eco-flow 顶层三层收敛（commit `9e64527` / `4d02d30`）后引用路径合并所致。**闸门本身就是发现这件事的东西**，现在 39 条路径全部指向存在的文件 |
| 「测试 10」 | 含 `sub/tests/__init__.py`；真实测试文件 9 个、用例 202 个 |

## 五、这条改造线的账：133 → 92 → 99

| 阶段 | 脚本数 | 发生了什么 |
|---|---|---|
| 折叠前 | 133 | 一个动作一个文件 |
| 2026-07-31 折叠后 | 92 | 42 旧件并成 12 个子命令族；20 个只做 argparse 转发的 group 模块（2265 行）折成 `_groups.py` 一张表 |
| 现在 | 99 | **涨的 7 个不是新功能** |

后半段涨的 7 个是：从抄了 N 份的判据里下沉出的两个 SSOT（`lib/cn_number.py`、`lib/caption_re.py`）+ 两道新闸门（`check_function_axis.py`、`check_external_refs.py`）+ 对应测试。

净效果：**代码净删 249 行，测试 82 → 202 个用例。**

### 文件层面已经减不动了

上一轮 8 条折叠提议经对抗核验**全部否掉，存活 0**。声称的重复行数逐条核实：`12→0` · `20→0` · `200→55` · `90→58` · `80→33`。

**真正的冗余在低一层**——同一判据被抄 N 份然后各自漂移：

| 被抄的判据 | 抄了几份 | 漂成什么样 |
|---|---|---|
| 中文数字转 int | 5 份 | 三档能力。`十六` 在 caption/styles 侧返 None，`一百零五` 在 blocks 侧返 None |
| 题注/图表编号正则 | 12 个文件 · 30 条 pattern · 20 个判据点 | 光「短横」就有 **8 种互不相同的字符子集，没有一处是全集** |
| zipfile 重打包 | 6 处手抄 | 已迁回 `lib/docx_surgical.py` |

## 六、目录不重构（2026-08-01 已拍板）

按格式轴（docx / pdf / md / xlsx / pptx）分目录被否，两条硬数字：

| 理由 | 数字 |
|---|---|
| 分不均 | 84 个非测试脚本里 docx 占 **56**——`docx/` 装 56 个，`pdf/` 和 `pptx/` 各装 1 个 |
| 搬完没法验收 | 222 处外部引用有 **130 处在 `~/Work` / `~/Apps`**，正是 `/refactor dir` 明文排除的地方（`paths.py:884` 的 roots 只有 `~/Dev`，`SCAN_EXCLUDE_DIRS` 含 `"Work"`） |

替代方案已落地：**目录轴回答「实现在哪」，职能轴回答「这条命令碰的是什么」，两条轴正交，所以职能不落成目录，落成一张表 + 一道机检。**

## 七、仍然挂账的一条

`scripts/document/sub/strip.py:1147` 的 `CAPTION_PATTERN = re.compile(r'^(表|图)\s*\d+\.\d+-\d+')` 与迁移前的 `styles._CAPTION_PATTERN` 逐字节相同，但**没有迁进 `lib/caption_re.py`**。

后果：SSOT 那一轮之后留下一处新的判据分歧——`strip` 只认 ASCII 短横，`styles` 现在认 5 种短横。

**没有自行修**：这是写盘动词的作用范围问题，与上一轮那条 [high] 级发现（`caption pair` 偷偷把 `表3.1-1` 压平成 `表1-1`）同类，改之前需要你确认。
