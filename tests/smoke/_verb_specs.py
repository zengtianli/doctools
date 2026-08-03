#!/usr/bin/env python3
"""_verb_specs.py — 每条 docx_cli 子命令的**冒烟契约** SSOT（2026-08-02 立）。

一条动词一行数据：怎么敲、预期 rc 是几、跑完源 docx 该不该变。
`tests/smoke/test_verbs_smoke.py` 逐行参数化去真跑，`tools/check_smoke_coverage.py`
拿这张表跟真实 CLI 对账（两个方向都 fail-closed）。

## 为什么要这张表

`cli_surface` 证明「接口没变」，`cli_forward_probe` 证明「argv 转发没变」，
`check_function_axis` 证明「每条都有职能标签」—— 三道闸门全绿，仍然可能
**每一条敲下去都是坏的**：它们没有一个真的执行过任何一条子命令。

2026-08-02 立表时实跑 93 条，当场撞见的就有：
`chapter delete --prefix` 声明了、也忠实转发了，可 `chapter.py` 只认
`--h1/--h1-text` —— 这个选项从写下那天起就不可能生效，前三道闸门全绿。

## 字段

    verb        动词路径，必须与 `docx_cli.py verbs` 输出的 path 逐字一致（闸门按它对账）
    argv        完整 argv 模板，占位符见下表；元组不是列表 —— 表是数据不是可变状态
    expect_rc   预期退出码。**不都是 0**，见下面「rc 不是 0 的那些」
    mutates     跑完之后**源 fixture docx 本身**是否应当改变
    note        为什么是这个 argv / 这个 rc；踩过的坑写在这里
    skip_reason 非空 = 这条不真跑（pytest.mark.skip），理由会被打印出来

### 占位符

    {DOCX}      主 fixture docx（每条动词独享新鲜副本）
    {DOCX2}     它的 v2（改了两段 + 多一段），供 diff / compare / relink 用
    {TEMPLATE}  只有样式与页眉页脚的模板 docx
    {DIR}       空输出目录        {OUT}/{OUTDOCX}/{OUTJSON}  输出文件
    {MD}/{MD2}/{MDDIR}   markdown 造件       {DRAFTS}   改动草稿目录
    {BIDDIR}    含敏感词的 md 目录            {DECISION}/{PLAN}  决策/搬移 plan JSON

### `mutates` 为什么值钱

它是**双向**断言，两个方向抓的是两类相反的 bug：

    mutates=False 却变了  → 只读动词偷偷写盘（`audit` 系、`snapshot`、`extract`…）
    mutates=True  却没变  → 写盘动词其实没干活（静默空跑，rc 照样 0）

注意它问的是「**源** docx 变没变」，不是「有没有产出文件」。`template` / `audit
images` / `table extract` 都写盘，但写的是 `-o` / `--report` / `--out-dir`，
源件必须一个字节都不动 —— 那恰恰是最该被钉死的一条。

### rc 不是 0 的那些（别顺手"修"成 0）

    audit-styleset <3 条>       rc = severity（fail=1）
    health diagnose/full/gate   rc = 健康度 / gate 结果（0 健康 / 1 警告 / 2 错误·FAIL）
    para scan-ppr               rc = 3 if suspects else 0
    para edit / para fix-ppr    rc = 4 —— 写盘**成功**后 `_finish_gate()` 复跑
                                `health gate` 拿到 FAIL 才返的（见下）
    renumber-fig（本表未覆盖）   rc = 3 = renum.EMPTY_RC「枚举为空，空集不报绿」

这些 rc **与 fixture 内容绑定**：换 fixture 就要重新实测这一列，不能照抄。
⚠ `health diagnose` / `health gate` 的 rc=2 与 docx_cli 的「参数错误 rc=2」撞码，
runner 不能把 rc=2 一律当失败。

#### 三条 rc 编码的是**同一个 fixture 事实**（2026-08-03 实测）

`health gate` = FAIL、`para edit` = 4、`para fix-ppr` = 4 —— 后两条的 4 完全来自
`sub/docx_para.py::_finish_gate()`：编辑本身成功了（stdout 有
`EDITED PARA 30 | …` / `FIXPPR PARA 12 …`，且 `mutates=True` 那条断言仍成立），
只是写盘后复跑的 gate 报 FAIL。gate 报的是**真判定不是回归**：

    caption-table-pairing     orphan_captions=1 · orphan_tbls=1
    caption-count-consistency 图: 编号集 4 ≠ caption 段 7 · 表: 编号集 3 ≠ caption 段 4

差额那几段正是 `caption number` 唯一的作业对象（无编号题注）+ 2 条英文题注。
**要把这三条变回 0，就必须把那些无编号题注从 fixture 里删掉** —— 那等于把
`caption number` 重新变成空跑。旧 fixture 的 PASS 是假绿：题注全 `Normal` 时
两个 check 都有「零 style 题注 → SKIP」守卫，压根没跑。故选择保留作业对象、
如实记 FAIL。

这两条 `para` 行**没有**加 `--no-gate` 把 rc 抹平成 0（该 flag 存在且实测可用：
`GATE SKIP (--no-gate)` + rc=0 + 源件仍变）。理由是加了就再没有任何一条 smoke
覆盖 `_finish_gate` 的 FAIL 分支，而那是这两条动词写盘后唯一的安全网。
代价是这三行会一起随 fixture 翻，翻了会响、不会静默 —— 这是可接受的耦合。

## 加一条子命令要改哪些地方

    1. sub/_groups.py 或各模块 register()      —— CLI 本身
    2. sub/_function_axis.py 的 _ROWS          —— 职能标签（check_function_axis 对账）
    3. **本表**                                —— 冒烟契约（check_smoke_coverage 对账）

漏第 3 步 → `python3 tools/check_smoke_coverage.py` 转红并点名。
反过来，表里留了 CLI 已经没有的动词同样判红 —— 两个方向都堵死，
这张表才不会变成「还在、只是不准了」。

## 为什么是 tuple-of-tuples，不是 dict 字面量

dict 字面量里写重复键不报错、后一条静默覆盖前一条 —— 那正是这张表最该被拦住的
一类手误（同一动词写两遍、只有后一条生效，覆盖率却看起来是满的）。
本仓 `sub/_groups.py` 的 GROUPS 与 `sub/_function_axis.py` 的 _ROWS 同样理由、同样形状。
"""
from __future__ import annotations

# (verb, argv, expect_rc, mutates, note, skip_reason)
VerbSpec = tuple[str, tuple[str, ...], int, bool, str, str]

_ROWS: tuple[VerbSpec, ...] = (
    # ── format —— 只碰版式 ────────────────────────────────────────────────
    ("style body", ("style", "body", "{DOCX}",), 0, False,
     "",
     ""),
    ("style table", ("style", "table", "{DOCX}",), 0, True,
     "",
     ""),
    ("style caption", ("style", "caption", "{DOCX}",), 0, True,
     "2026-08-03 fixture 加强后实测 tables=1 figures=1 skip=0（旧 fixture 是 0/0 空跑）。"
     "⚠ 图题/表题**故意用两个不同的样式名**（Image Caption / 表题）：styles._build_h4fig_plan "
     "是 elif 链先判图后判表，同名会让表分支恒为 0；且题注不能直接叫内置 Caption，"
     "否则 FIGURE_STYLE_PRIORITY 命中 no_change_skip、apply 分支一次都走不到",
     ""),
    ("fix clear-direct-format", ("fix", "clear-direct-format", "{DOCX}", "--inplace",), 0, True,
     "",
     ""),
    ("fix role-fill", ("fix", "role-fill", "{DOCX}", "--inplace",), 0, True,
     "",
     ""),
    ("fix style-create", ("fix", "style-create", "{DOCX}", "--base", "Normal", "--new-id", "SmokeStyle", "--new-name", "冒烟测试样式", "--inplace",), 0, True,
     "--base/--new-id/--new-name 三个都 required，缺一 argparse exit 2",
     ""),
    ("fix style-pane-filter", ("fix", "style-pane-filter", "{DOCX}", "--inplace",), 0, True,
     "",
     ""),
    ("fix style-pool-cleanup", ("fix", "style-pool-cleanup", "{DOCX}", "--inplace",), 0, True,
     "",
     ""),
    ("fix style-rebrand", ("fix", "style-rebrand", "{DOCX}", "--from", "Normal", "--to", "BodyText", "--inplace",), 0, True,
     "",
     ""),
    ("fix style-rename", ("fix", "style-rename", "{DOCX}", "--from", "caption", "--to", "图题", "--inplace",), 0, True,
     "--from 匹配的是 w:name 的 val：内置 Caption 样式 name 是小写 'caption'，写 'Caption' 会 renamed=0 静默无活",
     ""),
    ("fonts normalize", ("fonts", "normalize", "{DOCX}",), 0, True,
     "原地改 + 留 .bak-<时间戳>；fixture 里那个「等线」run 就是给它的活",
     ""),
    ("template", ("template", "{DOCX}", "-t", "{TEMPLATE}", "-o", "{OUTDOCX}",), 0, False,
     "必须传 -t：不传会去读 templates/template.docx —— 那是指向 ~/Work 真实交付件的 symlink",
     ""),
    ("format", ("format", "extract", "{DOCX}", "-o", "{OUTJSON}",), 0, False,
     "docx_fmt clone：extract 出 profile.json；apply 需要 --ref，同样会踩 template symlink",
     ""),
    ("chrome", ("chrome", "--validate", "{DOCX}", "{TEMPLATE}",), 0, False,
     "构建模式要求 pStyle=1 的院模板章标题，普通 docx 直接 rc=1 且 stdout/stderr 全空（静默失败）；故走只读的 --validate",
     ""),
    ("header-footer add", ("header-footer", "add", "{DOCX}", "--header", "某某水利工程可行性研究报告", "--footer-prefix", "第", "--no-backup",), 0, True,
     "--header 与 --footer-prefix 都是 required",
     ""),
    ("text-fmt", ("text-fmt", "{DOCX}", "--in-place", "--quotes", "--punct",), 0, True,
     "--help/-h 走 docx_fmt.TEXT_USAGE → rc=0（2026-08-02 前是「未知参数」rc=2）；"
     "未知 flag 仍 rc=2。--in-place 才写盘",
     ""),
    ("slim", ("slim", "{DOCX}", "-o", "{OUTDOCX}",), 0, False,
     "默认 safe 模式；给 -o 就不动源件",
     ""),
    ("renumber headings", ("renumber", "headings", "{DOCX}", "--no-backup",), 0, True,
     "",
     ""),
    ("renumber h4-figures", ("renumber", "h4-figures", "{DOCX}", "--no-backup",), 0, True,
     "2026-08-03 fixture 加强后实测 H4=1 图=4 表=3（旧 fixture 是 图=0 表=0 空跑）",
     ""),
    ("renumber-fig", ("renumber-fig", "{DOCX}", "--inplace",), 0, True,
     "非 --cn-section 走英文线。2026-08-03 fixture 加强后实测「图题数 2 · 现号 [3,1] → "
     "remap {3:1,1:2} · ✓ 连续 1..N」，rc=0 且源件变。⚠ 英文图号**故意乱序**：顺序对了 "
     "remap 退化成恒等映射、跑完源件一个字节不变，这条就假绿。"
     "⚠ 本表**没有覆盖** renum.EMPTY_RC=3 那条空集分支 —— 现 fixture 连 "
     "`--cn-section --kind 表 --check` 都非空（实测 rc=2「重复号 {'1':[1]}」）。"
     "要覆盖 3 得另造一份零题注 fixture，不能给本动词加第二行（_build() 会判重复动词）",
     ""),
    ("caption number", ("caption", "number", "{DOCX}", "--no-backup",), 0, True,
     "2026-08-03 fixture 加强后实测 tables_numbered=1 figures_numbered=1 chapters_detected=[4]。"
     "另有 3 条 manual_review 落 no-chapter-context（1 条中文 + 2 条英文题注）—— "
     "根因在实现侧：caption.is_h1_chapter() 只认 `一、` 不认 `第N章`，fixture 补不了",
     ""),
    ("caption number-by-style", ("caption", "number-by-style", "{DOCX}", "--no-backup",), 0, True,
     "2026-08-03 fixture 加强后实测 tables=1 figures=1 manual=0 chapters=[1,2,3,4]",
     ""),
    ("caption pair", ("caption", "pair", "{DOCX}", "--decision", "{DECISION}", "--no-backup",), 0, True,
     "--decision 是 required，且过 v1 schema 校验",
     ""),
    ("outline promote-h1", ("outline", "promote-h1", "{DOCX}", "--no-backup",), 0, False,
     "fixture 的 H1 本来就规整 → promote_count=0，rc 仍 0",
     ""),
    ("outline demote-h2", ("outline", "demote-h2", "{DOCX}", "--no-backup",), 0, False,
     "同上：candidates=0",
     ""),
    ("outline normalize-arabic", ("outline", "normalize-arabic", "{DOCX}", "--no-backup",), 0, True,
     "2026-08-03 fixture 加强后实测 change_count=1（`四、补充图件与指标` → `1 补充图件与指标`）。"
     "⚠ 顺带暴露一处实现侧不一致（不是本表的问题，不许在这里改期望值绕过）："
     "本动词把 `四、` 转成阿拉伯之后，`caption number` 的 caption.is_h1_chapter() 就再也认不出这一章 —— "
     "两条动词在 typeset_apply.py 步骤表里是先后关系",
     ""),
    ("freeze headings", ("freeze", "headings", "{DOCX}", "--no-backup",), 0, True,
     "",
     ""),
    ("freeze fields", ("freeze", "fields", "{DOCX}", "--no-backup",), 0, True,
     "",
     ""),
    ("chapter convert-arabic", ("chapter", "convert-arabic", "{DOCX}", "--no-backup",), 0, True,
     "",
     ""),
    ("image-caption", ("image-caption", "{DOCX}", "Caption",), 0, True,
     "第二个位置参数是样式名；不给就默认找 'ZDWP图名'，普通 docx 直接 rc=1。它把样式名也当文件扫一遍、打一行「⚠️ 文件不存在: Caption」，但成功计数仍是 1/1。"
     "⚠ 本行 argv 里没有任何 `-` 开头的 token，所以不受 2026-08-02 那道「未知 flag 一律 rc=2、禁 fallthrough」闸门影响；"
     "实测 `image-caption <docx> --dry-run` 现在是 rc=2（旧行为是静默忽略该 flag 照常写盘）",
     ""),
    ("fix-ref", ("fix-ref", "{DOCX}", "-o", "{OUTDOCX}",), 0, False,
     "fixture 里引用已是上标 → 「无需修改」，rc 0",
     ""),
    ("legacy fix-heading-disorder", ("legacy", "fix-heading-disorder", "{DOCX}", "--no-backup",), 0, False,
     "DEPRECATED，会往 stderr 打 WARNING，但能跑",
     ""),
    ("strip outlinelvl", ("strip", "outlinelvl", "{DOCX}", "--no-backup",), 0, True,
     "2026-08-03 fixture 加强后实测 captions_processed=2 outlinelvl_removed=2（旧 fixture 是 0/0）。"
     "这条正是 caption_re 短横字符类的哨兵：把 DASH_CHARS 改回只认 ASCII，processed 立刻掉到 1",
     ""),
    ("strip style-outlinelvl", ("strip", "style-outlinelvl", "{DOCX}", "--no-backup",), 0, False,
     "",
     ""),
    ("strip doc-protection", ("strip", "doc-protection", "{DOCX}", "--no-backup",), 0, False,
     "",
     ""),
    ("audit-styleset restore", ("audit-styleset", "restore", "{DOCX}", "--dry-run", "--no-llm",), 0, False,
     "重活：--dry-run + --no-llm 才是确定性的；不加 --no-llm 会调模型",
     ""),
    ("health fix", ("health", "fix", "{DOCX}", "--dry-run",), 0, False,
     "--dry-run 只规划",
     ""),
    ("health full", ("health", "full", "{DOCX}", "--dry-run",), 2, False,
     "rc = 健康度(0 健康/1 警告/2 错误)，与 fixture 内容绑定，**不是**参数错误",
     ""),
    ("para fix-ppr", ("para", "fix-ppr", "{DOCX}", "--para", "12", "--clone-from", "13", "--no-backup",), 4, True,
     "--para/--clone-from 用的是**含表格单元格**的段索引空间（与 doc.paragraphs 不同）。"
     "⚠ rc=4 **不是失败**：stdout 有 `FIXPPR PARA 12 (克隆源 PARA 13)` + before/after，"
     "改是改成了；4 来自写盘后 `_finish_gate()` 复跑 health gate 拿到 FAIL（详见本文件顶部那节）",
     ""),
    ("table borders", ("table", "borders", "{DOCX}", "--no-backup",), 0, True,
     "fixture 需含 ≥1 张表",
     ""),
    ("table center", ("table", "center", "{DOCX}", "--no-backup",), 0, True,
     "同上",
     ""),
    # ── content —— 碰正文写了什么 ────────────────────────────────────────────
    ("strip bookmarks", ("strip", "bookmarks", "{DOCX}", "--no-backup",), 0, True,
     "",
     ""),
    ("strip orphan-media", ("strip", "orphan-media", "{DOCX}", "--no-backup",), 0, False,
     "fixture 无孤儿媒体 → 扫完无活，rc 仍 0",
     ""),
    ("strip empty-captions", ("strip", "empty-captions", "{DOCX}", "--no-backup",), 0, False,
     "同上",
     ""),
    ("table delete-rows", ("table", "delete-rows", "{DOCX}", "--table-index", "0", "--rows", "3:3",), 0, True,
     "--rows 是 FROM:TO 闭区间 0-based，写成 '3' 会 ValueError unpack",
     ""),
    ("image relink", ("image", "relink", "{DOCX}", "--source", "{DOCX2}", "--no-backup",), 0, False,
     "--source 给另一份 docx；本 fixture 两份图片同源 → 无需重挂",
     ""),
    ("blocks reorder", ("blocks", "reorder", "{DOCX}", "--dry-run",), 0, False,
     "",
     ""),
    ("blocks relocate", ("blocks", "relocate", "{DOCX}", "--plan", "{PLAN}", "--dry-run",), 0, False,
     "--plan 是 required；plan.moves[].source_heading_text 必须在文档里找得到，否则该 move 记成 skip —— 但整体**仍 rc=0**，所以空 plan 看起来和干了活一样",
     ""),
    ("chapter delete", ("chapter", "delete", "{DOCX}", "--h1-text", "第3章", "--no-backup",), 0, True,
     "⚠ 用的是 --h1-text（经 rest 透传）。CLI 自己声明的 --prefix 会被原样转发给 chapter.py，而 chapter.py 只认 --h1/--h1-text → 必然 rc=2。--h1 收的是编号(^N\\s+)，不是标题文本",
     ""),
    ("chapter delete-empty-h1", ("chapter", "delete-empty-h1", "{DOCX}", "--no-backup",), 0, False,
     "fixture 无空 H1 → 无活，rc 仍 0",
     ""),
    ("split by-h1", ("split", "by-h1", "--docx", "{DOCX}", "--out-dir", "{DIR}",), 0, False,
     "--docx/--out-dir 都 required",
     ""),
    ("split body-replace", ("split", "body-replace", "--shell", "{DOCX}", "--content", "{DOCX2}", "--out", "{OUTDOCX}",), 0, False,
     "--shell/--content/--out 三个都 required",
     ""),
    ("combine", ("combine", "--docx", "{DOCX}", "{DOCX2}", "--out", "{OUTDOCX}",), 0, False,
     "--docx 收多份，--out required",
     ""),
    ("chapters-sync", ("chapters-sync", "{DOCX}", "--chapters-dir", "{MDDIR}",), 0, False,
     "不加 --execute = DRY-RUN，不动盘",
     ""),
    ("md-merge", ("md-merge", "{MD}", "{DOCX}", "--start-anchor", "第2章 水文分析", "--end-anchor", "第3章 结论与建议", "--no-backup",), 0, False,
     "位置参数是 <md> <docx>；不给 start/end 索引就必须给 --start-anchor/--end-anchor，否则 rc=2",
     ""),
    ("para edit", ("para", "edit", "{DOCX}", "--para", "30", "--replace", "50 年一遇", "100 年一遇", "--no-backup",), 4, True,
     "同上索引空间：'按照规范要求' 在 doc.paragraphs 里是 18，在这里是 30 —— 索引拿错会 rc=2 且报「未在段 N 找到」。"
     "⚠ rc=4 **不是失败**：stdout 有 `EDITED PARA 30 | -\"50 年一遇\" +\"100 年一遇\"`，"
     "4 来自写盘后 `_finish_gate()` 的 health gate FAIL（详见本文件顶部那节）",
     ""),
    # ── review —— 审阅修订 ────────────────────────────────────────────────
    ("strip revisions", ("strip", "revisions", "{DOCX}", "--no-backup",), 0, False,
     "",
     ""),
    ("track", ("track", "read", "{DOCX}",), 0, False,
     "只跑 read 分支。`track compare` 2026-08-03 已落地（段落级 diff → w:ins/w:del，实现在 sub/docx_track.py），"
     "不进本表也不加行：`track` 是 CMD_TABLE fast-path，`docx_cli.py verbs` 里只有 `track` 一个动词路径，"
     "read/review/compare 都不是动词 —— 加一行 `track compare` 会让 check_smoke_coverage 判「表里有 CLI 没有的动词」转红。"
     "且 compare 要两份 docx + 一个输出路径、rc 有 0/1/2/3 四种含义，一行一 argv 装不下。"
     "它的门是 scripts/document/tests/test_track_compare.py（13 条）",
     ""),
    ("md-merge-track", ("md-merge-track", "--src", "{DOCX}", "--md", "{MD}", "--anchor", "第1章 概述", "--out", "{OUTDOCX}",), 0, False,
     "--src/--md/--anchor 三个 required；--anchor 与 fixture 标题绑定",
     ""),
    ("revise-rules gen", ("revise-rules", "gen", "--drafts-dir", "{DRAFTS}", "--docx", "{DOCX}", "--out", "{OUTJSON}",), 0, False,
     "⚠ 默认 glob 是 `*-改动草稿.md`（带短横），而 compare-ref ref 的默认是 `*改动草稿.md`。fixture 文件名 `第1章-改动草稿.md` 是同时满足两者的写法；不带短横这条就静默变成「待处理 MD: 0 份」+ rc=0",
     ""),
    ("seqdiff seq", ("seqdiff", "seq", "--src", "{DOCX}", "--dst", "{DOCX2}", "--out", "{OUT}",), 0, False,
     "--src/--dst/--out 都 required；统计打在 stderr",
     ""),
    ("compare", ("compare", "{DOCX}", "{DOCX2}", "-o", "{OUT}",), 0, False,
     "位置参数 before after",
     ""),
    # ── inspect —— 只读检查 ───────────────────────────────────────────────
    ("table extract", ("table", "extract", "--docx", "{DOCX}", "--out-dir", "{DIR}",), 0, False,
     "docx 走 --docx，不是位置参数",
     ""),
    ("image extract", ("image", "extract", "{DOCX}", "--out-dir", "{DIR}",), 0, False,
     "按邻近 caption 语义命名，需要图 + 图题注",
     ""),
    ("seqdiff image", ("seqdiff", "image", "--src", "{DOCX}", "--dst", "{DOCX2}", "--out", "{OUT}",), 0, False,
     "同上三个 required；两份 fixture 有逐字节重复的图 → 报「重复 2 张」",
     ""),
    ("compare-ref ref", ("compare-ref", "ref", "--drafts-dir", "{DRAFTS}", "--ref", "{DOCX}", "--out", "{OUT}",), 0, False,
     "同上 glob 差异；空目录同样 rc=0（「改为段 0 条」），空跑与真跑同码",
     ""),
    ("audit bookmarks", ("audit", "bookmarks", "{DOCX}",), 0, False,
     "无书签也 rc=0（全 0 的空跑），所以 fixture 在两个 H1 上挂了 _Toc_*",
     ""),
    ("audit captions", ("audit", "captions", "{DOCX}",), 0, False,
     "",
     ""),
    ("audit fields", ("audit", "fields", "{DOCX}",), 0, False,
     "只扫 word/document.xml —— 只有页脚有域的话它读不到，所以 fixture 正文里塞了 REF+TOC",
     ""),
    ("audit headings", ("audit", "headings", "{DOCX}",), 0, False,
     "",
     ""),
    ("audit images", ("audit", "images", "{DOCX}", "--report", "{OUT}",), 0, False,
     "必须给 --report，否则默认写 /tmp/audit-images-<stem>.json（sandbox 外裸写）",
     ""),
    ("audit table-pairing", ("audit", "table-pairing", "{DOCX}",), 0, False,
     "需要表 + 表题注，否则空跑",
     ""),
    ("audit-styleset body-style-concentration", ("audit-styleset", "body-style-concentration", "{DOCX}",), 1, False,
     "同上 rc=severity：body 集中度 0% < 阈值 → fail",
     ""),
    ("audit-styleset role-coverage", ("audit-styleset", "role-coverage", "{DOCX}",), 1, False,
     "同上 rc=severity：mandatory 角色缺失 → fail",
     ""),
    ("audit-styleset style-coherence", ("audit-styleset", "style-coherence", "{DOCX}",), 1, False,
     "rc = severity(pass/warn=0, fail=1)。fixture 用 python-docx 默认样式、缺 ZDWP 系列 → fail。换成院模板 docx 这条会变 0 —— expect_rc 与 fixture 绑定",
     ""),
    ("audit-styleset style-pane-filter", ("audit-styleset", "style-pane-filter", "{DOCX}",), 0, False,
     "同上 rc=severity：此条落 warn → 0",
     ""),
    ("audit-styleset style-pool-cleanliness", ("audit-styleset", "style-pool-cleanliness", "{DOCX}",), 0, False,
     "同上 rc=severity：此条落 warn → 0",
     ""),
    ("check", ("check", "snapshot", "{DOCX}",), 0, False,
     "`check` 是组，裸敲 `check <docx>` = rc 2（invalid choice），必须带 snapshot|compare",
     ""),
    ("snapshot", ("snapshot", "{DOCX}",), 0, False,
     "",
     ""),
    ("extract", ("extract", "{DOCX}",), 0, False,
     "只读导出正文",
     ""),
    ("section read", ("section", "read", "{DOCX}", "--list",), 0, False,
     "--list 枚举标题，与正文内容无关，比带 query 更适合 smoke",
     ""),
    ("scan-sensitive", ("scan-sensitive", "{BIDDIR}", "--json",), 0, False,
     "",
     "真跑要对目录里每个 md 调 LLM（lib/llm_client.py → claude -p），非确定性 + 每次执行有真实模型成本，不进确定性 smoke。（2026-08-02 实测确认能跑：对含敏感词的 bid.md 目录 rc=0，返 4 条 severity=high）"),
    ("health diagnose", ("health", "diagnose", "{DOCX}", "--report", "{OUT}",), 2, False,
     "同上 rc=健康度；⚠ 与 docx_cli 的「参数错误 rc=2」撞码。另外它会往 /tmp/docx_health_<stem>/ 落分项 JSON，不受 --report 控制",
     ""),
    ("health gate", ("health", "gate", "{DOCX}",), 2, False,
     "rc 由 gate 结果定（PASS=0 / FAIL=2）；不给 --report 也会往 /tmp/docx_health_<stem>/ 落盘。"
     "2026-08-03 fixture 加强后是 FAIL=2，failed=['caption-table-pairing','caption-count-consistency'] —— "
     "**真判定不是回归**（旧 fixture 的 PASS 是假的：题注全 Normal 时两个 check 都走「零 style 题注 → SKIP」）。"
     "⚠ 与 docx_cli 的「参数错误 rc=2」撞码；也是 para edit / para fix-ppr 那两条 rc=4 的来源",
     ""),
    ("health-split", ("health-split", "{DOCX}", "--out-dir", "{DIR}", "--no-backup", "--no-fix",), 0, False,
     "⚠ 必须 --no-backup，默认备份目标是 ~/Archives/docx-backups/<today>/（sandbox 外）；--no-fix 也几乎必须：不带它会触发 styleset restore，非院模板 docx 直接 rc=1",
     ""),
    ("para locate", ("para", "locate", "{DOCX}", "第2章 水文分析", "--json",), 0, False,
     "第二个位置参数与 fixture 正文强绑定；查不到时 hits 为空但仍 rc=0（不是 fail-closed）",
     ""),
    ("para inspect", ("para", "inspect", "{DOCX}", "--para", "1", "--json",), 0, False,
     "--para 1 是最不挑 fixture 的取值",
     ""),
    ("para scan-ppr", ("para", "scan-ppr", "{DOCX}", "--json",), 3, False,
     "rc = `3 if suspects else 0`（docx_para.py）。本 fixture 4 条 suspects → 3，与内容绑定",
     ""),
    ("para render", ("para", "render", "{DOCX}", "--warm",), 0, False,
     "",
     "--warm 要起 LibreOffice(soffice) 转 PDF，约 6s，且缓存落 ~/.cache/doctools/ —— 外部二进制依赖 + sandbox 外写盘，不进确定性 smoke。（2026-08-02 实测确认能跑：rc=0，「WARM fixture.docx → 5 页缓存于 …」）"),
    ("verbs", ("verbs", "--json",), 0, False,
     "只读动词清单，smoke 里它顺带证明 CLI 自省没坏",
     ""),
    # ── convert —— 跨格式 ────────────────────────────────────────────────
    ("md-to-docx", ("md-to-docx", "{MD}", "-t", "{TEMPLATE}", "-o", "{OUTDOCX}",), 0, False,
     "必须传 -t 断掉 template symlink 依赖。⚠ 它的 `extract` 子模式会把 heading_styles.xml/styles_info.txt **无条件写进当前工作目录**且没有 -o，故 smoke 只跑转换模式",
     ""),
    ("md", ("md", "format", "{MD}",), 0, False,
     "只跑确定性的 `md format`。⚠ `md frontmatter` 真跑会逐文件调 LLM 并原地改写，故不进 smoke",
     ""),
    # ── dispatch —— 调度编排 ──────────────────────────────────────────────
    ("pipeline run", ("pipeline", "run", "{DOCX}", "--steps", "health-diagnose", "--no-backup", "--report-dir", "{DIR}",), 0, False,
     "--steps 只认 7 个内置步名（audit-styleset-all / split-by-h1 / health-diagnose / image-extract / table-extract / slim-safe / slim-aggressive）或 --step-dir 下的模块；给个不存在的名字会 fallthrough 到 import scripts.<name> → 每份 docx 记 ERROR 但**整体 rc=1**",
     ""),
)


# ─── 派生视图（别在别处再抄一遍这张表） ──────────────────────────────────

def _build() -> "dict[str, dict]":
    out: dict[str, dict] = {}
    for verb, argv, rc, mutates, note, skip in _ROWS:
        if verb in out:                      # tuple 表不会自动去重，得自己拦
            raise ValueError(f"_verb_specs: 动词重复出现两次: {verb!r}")
        out[verb] = {
            "verb": verb, "argv": argv, "expect_rc": rc,
            "mutates": mutates, "note": note, "skip_reason": skip,
        }
    return out


SPECS: "dict[str, dict]" = _build()
VERBS: tuple[str, ...] = tuple(SPECS)
RUNNABLE: tuple[str, ...] = tuple(v for v in VERBS if not SPECS[v]["skip_reason"])
SKIPPED: tuple[str, ...] = tuple(v for v in VERBS if SPECS[v]["skip_reason"])
