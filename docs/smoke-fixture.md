# 冒烟轴与 fixture 富度全记录（从 CLAUDE.md 外迁 · 2026-08-04）

> 常驻结论（表在哪 / mutates 语义 / expect_rc 警告 / 加子命令三处加行）在 `CLAUDE.md`
> §「冒烟轴」；本页是 fixture 富度判据、2026-08-03 重测明细与 EMPTY_RC 覆盖史。

## 为什么立冒烟轴（2026-08-02）

前面几道闸门（`cli_surface` 接口指纹 · `cli_forward_probe` 转发 argv ·
`check_function_axis` 职能标签）**没有一道真的执行过任何一条子命令**。三道全绿，
仍然可能每一条敲下去都是坏的。立表当天实跑 93 条就撞见一例：
`chapter delete --prefix` 声明了、也忠实转发了，可 `chapter.py` 只认
`--h1/--h1-text` —— 这个选项从写下那天起不可能生效，前三道闸门全绿。

## fixture 的富度是机器判据，不是注释（2026-08-03 立）

`mutates` 抓的是「写盘动词其实没写」，但它抓不到**上一层**的空跑：动词写没写盘，
取决于 fixture 里有没有它认得的对象。2026-08-03 实测出的四处零对象面：

| 症状 | 后果（rc 全是 0，看输出也像干了活） |
|---|---|
| 6 个题注段全套 `Normal` | `renumber h4-figures` 打 `图=0 表=0`；`caption` 族与 `style caption` 一个对象都扫不到 |
| 题注短横只有 ASCII 一种 | `caption_re` 声明认 5 种，只测 1 种；「退回只认 ASCII」这类回归（2026-08-02 `strip.py` 真的发生过）冒烟看不见 |
| 题注全是两段式 `图1-1` | `SECTIONED_CAPTION`（`style caption` / `strip outlinelvl` 的写盘范围，只认三段式）恒为空集 |
| 中文题注再多也喂不到 `renumber-fig` 默认模式 | 中英两条编号线在 `caption_re` 里零共享，默认模式只认 `Figure N` |
| 草稿写成 `## 意见 N` | `revise-rules gen` 的 `SECTION_RE` 只认 `## 改动 N` → `parse_md` 直接 `return []` → 写出一个 `[]` 还 rc=0 |

所以 `_fixture.py` 自带 `selfcheck()`：题注样式 / 五种短横 / 三段式 / `outlineLvl`
病灶 / 内嵌图 / 英文图题 / 草稿节数逐条数，低于下界即 exit 1。
**下界留了余量**（加东西不会碰它，砍掉一整类对象才会）。

⚠ 两处判据看着像洁癖，其实都是被实现逼出来的，别顺手"统一"掉：

- **图题用 `Image Caption`、表题用 `表题`，两个名字不能合并成一个。**
  `styles_registry.yaml` 把 `Caption` 同时列进 FIG 与 TABLE 两族，而
  `styles._build_h4fig_plan` 是 `elif` 链、先判图后判表 —— 表题若也叫 `Caption`
  会被图分支先认领、再被 `startswith("图")` 否掉，`plan_tbl` 恒为 0。
  另外 `caption_re.is_fig_caption_style("Caption")` 是 **False**，用它会让
  `shape_contract.caption_figure_count` 恒为 0。
- **英文图题号故意是乱的（`Figure 3` 在 `Figure 1` 前面）。** 顺序对了 remap 就是
  恒等映射，`renumber-fig --inplace` 跑完源件一个字节不变，`mutates` 断言等于没写。

## 2026-08-03 换 fixture 之后整列重测的结果（10 行变了，逐条真敲测出来的）

用一次性驱动（与 runner 同一造件路径、同一 copytree、同一 subprocess 参数）跑完 91 条，
只有这 10 条与旧声明不符 —— **全部属于「fixture 变强 → 动词现在有对象可处理」，
没有一条是 fixture 揭出来的崩溃**：

| 变的是什么 | 哪几条 | 实测证据 |
|---|---|---|
| `mutates` **False → True**（7 条） | `style caption` · `renumber h4-figures` · `renumber-fig` · `caption number` · `caption number-by-style` · `outline normalize-arabic` · `strip outlinelvl` | 分别打出 `tables=1 figures=1` / `图=4 表=3` / `remap {3:1,1:2}` / `chapters_detected=[4]` / `chapters=[1,2,3,4]` / `change_count=1` / `processed=2 removed=2`，旧 fixture 上全是 0 |
| `expect_rc` **0 → 2**（1 条） | `health gate` | FAIL，`failed=['caption-table-pairing','caption-count-consistency']` |
| `expect_rc` **0 → 4**（2 条） | `para edit` · `para fix-ppr` | 编辑本身成功（stdout 有 `EDITED PARA 30 …` / `FIXPPR PARA 12 …`），4 来自写盘后 `_finish_gate()` 复跑 health gate 拿到 FAIL |

**这三条 rc 编码的是同一个 fixture 事实**，会一起随 fixture 翻。`health gate` 的 FAIL 是
**真判定不是回归**：旧 fixture 的 PASS 是假绿 —— 那两个 check 都有「零 style 题注 → SKIP」
守卫，题注全 `Normal` 时压根没跑。现在报的是「图: 编号集 4 ≠ caption 段 7」「orphan_captions=1」，
而差额那几段正是 `caption number` 唯一的作业对象。**要把 rc 变回 0 就得删掉那些无编号题注 =
把 `caption number` 重新变成空跑**，所以选择保留对象、如实记 FAIL。

`para` 那两条**没有**加 `--no-gate` 把 4 抹平成 0（该 flag 存在、实测 `GATE SKIP (--no-gate)`
+ rc=0 + 源件仍变）：加了就再没有任何一条 smoke 覆盖 `_finish_gate` 的 FAIL 分支，
而那是这两条写盘动词唯一的安全网。

## `renum.EMPTY_RC=3` 的回归门为什么单开一个文件

**`scripts/document/tests/test_renum_empty_set.py`**（7 条）。冒烟轴按设计是
「一条动词一份标准 fixture」，装不下「同一条动词在空集上的行为」—— 而加强后的 fixture
连 `--cn-section --kind 表 --check` 都非空（实测 rc=2「重复号 `{'1': [1]}`」），给
`renumber-fig` 加第二行又会被 `_verb_specs._build()` 判重复动词。所以它不进 `_verb_specs`。

⚠ 立这个文件的直接原因是**它一度真的零覆盖**：`EMPTY_RC` 与加强 fixture 是同一轮
落的，于是新判据刚写完就被自己那轮的 fixture 绕过 —— 谁把 `_empty_set_exit` 改回
`return`，92 条 smoke 一条都不会红。反向验证：注入 `return` → 本文件 6 红 / 恢复后 7 绿。
第 7 条是**非空对照**（有 2 条图题注的文档不许被判成空集），没有它，把
`_empty_set_exit` 挪到函数开头无条件调用也能让前 6 条全绿。

## `mutates` 是布尔的，看不见「部分退化」

2026-08-03 把 `caption_re.DASH_CHARS` 注回只认 ASCII（本仓真发生过的那类回归），
`strip outlinelvl` 的 `processed` 从 2 掉到 1，**md5 照样变、smoke 全绿 92 passed**；
红的是 `test_caption_re.py`（15 failed）。所以那道 pytest 门不能省，冒烟替不了它。
