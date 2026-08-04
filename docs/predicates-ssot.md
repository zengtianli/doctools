# 共享判据 SSOT 全记录（从 CLAUDE.md 外迁 · 2026-08-04）

> 常驻结论（API 用法 + 禁令 + 回归门命令）在 `CLAUDE.md` 各小节；本页是合并史、实测
> 后果与「为什么不许顺手统一」的完整论证。涵盖 `caption_re` / `cn_number` /
> `lsof_check` / `renum.EMPTY_RC` 四件。

## `lib/caption_re.py`：题注/图表编号判据（2026-08-01 合并）

合并前全仓 12 个文件 ~30 条 pattern、20 个互相竞争的判据点，光「短横」就有
**8 种互不相同的字符子集**，没有一处是全集（`bid_gate` 有 U+2011 无全角 `－`，
`renum` 恰好反过来，两个门互为盲区）。实测后果：`renum figures` 把已编号的 `图1‑2`
当无号题注 prepend 出 `图1-3 图1‑2 …`，而自检用**同一条瞎正则**读回，对着写坏的
文档打印「✓ 每节连续」。

**不是一条正则打天下** —— 差异有一半是有意的（`bid_gate` 的右界断言防吃正文内联引用、
`bid_residue` 的非锚定是为扫交叉引用、`styles._CAPTION_PATTERN` 强制三段是写盘范围、
`table`/`image` 只看首字是为给无编号题注命名）。所以形状是「一套字符类 + 一个构造器 +
若干具名 spec」：**加调用点 = 加一个 spec 声明，不是再写一条正则**。

想「顺手统一」样式名/关键词三套判据 → **先别**：那是给写盘动词扩范围，模块 docstring
写了为什么不合并。

⚠ `cli_surface` / `cli_forward_probe` **看不见这层**（那两个只管 argv），改判据必须跑
`pytest scripts/document/tests/test_caption_re.py`（55 条，含 9 种变异实证能抓）；
改完还要用真 fixture 走一遍 CLI（`renum figures` / `health diagnose` / `caption pair`
是三条已知会现形的链）。

**2026-08-02 补一处漏网**：`sub/strip.py` 的 `CAPTION_PATTERN` 与迁移前的
`styles._CAPTION_PATTERN` 逐字节相同却没跟着搬，于是归并那轮**自己造出一处新分歧** ——
strip 只认 ASCII `-`，styles 侧已认 5 种。现共用 `STRIP_OUTLINELVL_CAPTION`
（= `SECTIONED_CAPTION` 同一对象）。实测 `strip outlinelvl` 在 6 种编号的 fixture 上
从 processed 2/7 变 7/7（U+2011 / U+2013 / U+2014 / U+FF0D / 全角句点 5 条原本漏网、
`w:outlineLvl` 继续污染 Word 导航窗格）。教训：**「逐字节相同」正是最容易被跳过的那种，
归并收尾必须按 grep 结果逐条销号，不能靠「看起来都搬完了」**。

## `lib/cn_number.py`：中文数字转 int（2026-08-01 合并）

合并前全仓 **5 处** 各写各的（chapter / outline / blocks / caption / styles），分三档能力，
同一个输入给三种答案：`十六` 在 caption/styles 侧返 None、`一百零五` 在 blocks 侧返 None。
后果不是学术问题 —— `caption number` 的章计数器解析失败就**不换章**，第 16 章往后的表图
继续按上一章编（表15-7、表15-8…）；`styles` 那侧上层写的是
`_parse_chapter_from_text(t) or (chapter + 1)`，**静默拿「上一章+1」顶上**。
三个入口都挂在 `typeset_apply.py` 步骤表里，/typeset 一条龙每次都在跑。

**两个 API 必须并存，别合成一个** —— chapter/outline 三处靠 `except ValueError` 控流
（收 None 会拿着 None 往下算、写出「None、标题」且不报错），blocks/caption/styles 三处靠
None 分支（抛异常会直接崩）。`cn_to_int` 就是 `chinese_to_arabic` 外包一层 try，
语义不会再分叉。

能力上界 = 十/百/千 + 「十X」省略一 + 〇/两 + 阿拉伯直通；**「万」不支持是有意的**（旧的
三档没一档支持，加它要改累加器结构）。**本模块不管正则** —— 各调用点的章标题字符类没统一
（caption/styles 侧不含 `百`/`零`，所以「一百零五、」压根匹配不到），那是另一根轴。

## `_cli_common.lsof_check` / 备份路径（2026-08-02 收敛）

收敛前全仓 **10 份 `lsof_check`，四种语义，没有任何两份逐字相同**：A 裸 `lsof <path>`
（5 份）· B `lsof -- <path>` 且要求行数>1（3 份）· C 返 bool、不看 returncode、
`except Exception` 全吞（1 份）· D 无 timeout（1 份）。用户拍板**取最健壮的 B**。

**`--` 不是洁癖**：实测 `lsof -hold.docx` 被解析成 `-h`（帮助）→ **rc=0、stdout 空** →
老的 A/C/D 三派一致判「空闲」，对着 Word 正开着的文件放行写盘。returncode 检查救不了它，
只有 `--` 能。所以别顺手把它删掉当「多余参数」。

⚠ `pipeline_lib.lsof_check` / `make_backup_path` 现在只是**委派壳**，别再去那里找真身 ——
它们留着是因为 `pipeline_lib.__all__` 列着、`typeset_apply.py:98` 按名 import、外仓还有
`import *` 的 shim；**名字面只增不减**。

**不归本机制管**：`.bak-<时间戳>` 是**另一套命名约定**（`table` / `normalize_fonts` /
`md_merge_impl` / `docx_fmt` / `bid_gate` / `lib/docx_surgical.make_backup`），
扫进来就是行为变更。

## `renum.py` 枚举为空 = exit 3（2026-08-03 立）

三个子命令原来在**一个对象都没枚举到**时全部打通过语并 exit 0。根因是连续性判据在
空集上**恒真**，于是自检对着「什么都没量到」发合规证：

| 路径 | 空集上原本打的 | 恒真的那个式子 |
|---|---|---|
| `figures`（默认英文线） | `✓ 连续 1..N` | `[] == list(range(1, 1))` |
| `figures --cn-section` | `✓ 每节连续 1..k` | `all(…)` over `{}` |
| `figures --cn-section --check` | `✓ 图序号与居中均合规` | `any(4 个空桶)` = False |
| `tabfig` | `✅ 表/图编号与章号全部对齐` | 循环体一次都没进 |
| `chapter` | `磁盘已与 config 一致 (no-op)` | `all(())`（`sequence: []`） |

现在这五处一律 **exit 3 + `✗ 未发现任何… —— 枚举为空`**（`renum.EMPTY_RC`）。

**3 而不是 2**：该文件的 2 已经被三种「真发现了问题」占着（`--check` 报断号/重号 ·
检测到重复图号 · 重编号后仍不连续）。空集是**没能做出任何判定**，与之正交；混进 2
会让上游把「这份文档没有图题」误读成「图号有问题」，实测两处会真坏 ——
`doc_dispatch.do_renum` 的 `图` 那轮返非 0 就 `break`，`--kind 表` 整轮不跑；
`typeset_pipeline` 步骤③ 见 `rc!=0` 即回滚，把 `--fix-center` 已经写进盘的居中一起撤掉。
两处都已放行 3（后者用 per-step `keep_rcs`，对账卡上显示 `○ 空集(rc=3)` 而非 `✅ 完成`）。
选 3 也对齐本仓既有先例：`para scan-ppr` 用 3 表示「非错误、但需要人看」。

⚠ **写盘与空集正交**：`--fix-center` 时先写盘再报空集（居中是真活），其余情况在写盘前
就退出 —— 产一份逐字节复制的 `.renumbered.docx` 只会让下游以为「重排过了」。

⚠ 空集判据必须问**枚举端**不能问**计划端**：`renumber_cn_section` 为此多返一个
`n_found`。`len(plan)` 在 `--no-supplement` / 无号题注定位不到节号时同样是 0，
「枚举到了但没排进计划」与「一个题注都没有」在 plan 上长得一模一样。

> `tests/smoke/_verb_specs.py` 的 `renumber-fig` 行曾照着空集绿写下 `expect_rc=0`——
> 2026-08-03 换强 fixture 后已逐条重测改正（现 rc=0/mutates=True 是真作业不是空跑）；
> 空集行为的覆盖在 `test_renum_empty_set.py`，见 `docs/smoke-fixture.md`。
