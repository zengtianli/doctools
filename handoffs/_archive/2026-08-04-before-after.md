---
title: docx 工具链 before / after
slug: before-after
project: doctools
date: 2026-07-30
out: /Users/tianli/Dev/tools/doctools/reports/before-after.html
---

> **全部做完了**（2026-07-30 收尾）。下面每个数字都是实跑出来的，不是估的；
> 涉及「多了/少了」的一律用同一把尺子在折前折后各跑一遍，不拿两个口径对比。

---

## 零 · 一句话

| | before | after |
|---|---|---|
| 改一份报告的版面 | 手敲 **20 条命令**，每条自己开文件自己存盘 | 写**一份 yaml**，跑一条命令 |
| python-docx 存一次盘 | 重写 **~60 个部件**（语义只变 1 个） | 重写 **1~2 个** |
| 跑完一轮排版 | 存盘 **15 次** · 进程 **20 个** · **100 秒** | 存盘 **6 次** · 进程 **1 个** · **33 秒** |
| 加一条横切改动（如收口） | 改 **35 处** | 改 **1 处** |
| **涉 docx 的脚本** | **117 个** / 42,572 行 | **98 个** / 42,059 行 → 退孤儿后 **96 个**（§十一） |
| 加一个子命令 | 新建一个 group 模块文件 | **加一行数据** |
| 没挂收口就改 docx | 打一行提醒，照跑 | **hook 硬拦 exit 2** |

---

## 一 · 存盘：拿错尺子的那次

### before

python-docx 打开一份 301 部件 / 75 个嵌入公式的真报告，**什么都不改**直接存回：

```
zip 条目数     301 → 301      一个没少   ← 只看这个会判成「没损失」
字节变了的部件  60 个
语义真变了的    1 个           ← 另外 59 个纯属被重新序列化
```

**条目数相同 ≠ 没损失。** 真该问的是「改一个字号碰了多少个不该碰的文件」：60 个，
每一个都是一次「Word 可能渲染得不一样」的机会。

### after

`lib/docx_safe_save.py` 补在 python-docx 唯一的存盘咽喉（`OpcPackage.open/save`，
`inspect.getsource` 实测出来的），存盘时：

```
1. 原样写到同目录临时件（不碰目标）
2. graft_unchanged：逐部件 XML 规范化(C14N)对比，语义没变的按原件字节还原
3. 原子 replace 到目标
```

graft 在 replace **之前** —— 判定「部件整个丢了」时目标文件还是原样。

| 场景 | before | after |
|---|---:|---:|
| 什么都不改，打开即存回 | 60 个部件被重写 | **0**（输出与原件逐字节相同） |
| 改一个字体 | 59 个意外重写 | **1** |
| 给 17 个节加页眉页脚 | 59 | 26（新增 10 个 header/footer 是本职） |

**59 个脚本一行业务逻辑没改**，每个只加 4 行 import。

---

## 二 · 排版：从串命令到一份 spec

### before

`typeset_pipeline.py` 是 subprocess driver：每步 fork 一个脚本，各自开文件、各自存盘。
实际用起来是手敲一串：

```bash
python3 sub/strip_revisions.py 报告.docx
python3 sub/strip_bookmarks.py 报告.docx
python3 sub/strip_empty_captions.py 报告.docx
python3 sub/renumber_headings.py 报告.docx
python3 sub/number_captions.py 报告.docx
python3 sub/add_header_footer.py 报告.docx --header … --footer-prefix …
python3 sub/freeze_heading_numbers.py 报告.docx --levels 1,2,3,4
python3 sub/normalize_fonts.py 报告.docx --body-cjk 宋体 --heading-cjk 黑体
…还有 12 条
```

顺序错了**不会报错**，只会安静地产出错的版面（例如章号还没定就编题注号，编出来的下一步就作废）。

### after

```bash
python3 scripts/document/typeset_apply.py 报告.docx \
    --spec config/spec-examples/report-generic.yaml
```

spec 长这样（节选）：

```yaml
version: 1
actions:
  strip_revisions: {revision_mode: accept-all}
  strip_bookmarks: {bookmark_prefixes: "_Toc,_Ref,_Hlk"}
  renumber_headings: {h1_base: 1}
  number_captions: true
  add_header_footer:
    header: 金华江流域生态流量分类管控保障方案
    footer_prefix: 浙江省水利水电勘测设计院
    font_size: 10.0
  normalize_fonts: {body_cjk: 宋体, heading_cjk: 黑体, latin: Times New Roman}
```

**顺序不给 spec 控** —— 有真依赖（先冻结域再改字体、先配对题注再编号），顺序写死在引擎的
`ACTIONS` 表里，每条带 `why_here` 说明理由。spec 只管「跑不跑」和参数。

| | before | after |
|---|---:|---:|
| 进程数 | 20 | **1** |
| 存盘次数 | 15 | **6** |
| 墙钟 | 100 s | **33 s** |
| 改写部件（关掉页眉页脚这一档） | 4 | **2** |
| 顺序排错 | 静默出错版面 | 表里定死，配不出来 |
| 拼错一个选项 | 静默忽略 | **rc=1 报错** |

> 更正一句我先前说错的：**「炸开面大幅收窄」在本仓不成立**。老链每个脚本自己也挂了收口，
> 单次存盘早被压到地板（4 vs 2）。真实收益是**少 9 次存盘、少 19 个进程、快 3 倍**。

---

## 三 · 七件事：不是一类，是两类

这条是探路时发现的，**改变了方案**。判据是**参数长什么样**：

| | 特征 | 例子 | 能 spec 化 |
|---|---|---|---|
| **设值类** | 参数就是值（字体/字号/颜色/对齐/文字/层级） | `normalize_fonts` `add_header_footer` `set_table_borders` `freeze_heading_numbers` | 能 |
| **识别类** | 参数只有 `--dry-run/--no-backup/--report`，**一个值都没有** | `number_captions` `pair_table_captions` `set_table_align` | **不能，只能开关** |

`number_captions` 397 行里全是 `has_nearby_table` / `has_nearby_drawing` / `parse_chapter` ——
它在**猜**「这一段是不是表名」。这是识别问题不是配置问题，没有值可配。
`pair_table_captions` 更直白：它要一个 `--decision` JSON，等于说「我猜不准，你先给判断结果」。

所以 spec 的边界是 **配置设值类 + 开关识别类**，不是「一份 spec 取代所有脚本」。

---

## 四 · 完成度

| | 状态 |
|---|---|
| 存盘收口 | ✅ 59 个脚本 / 11 个 repo 全部接入，守卫 36/36 绿，commit 期硬拦 |
| spec 引擎 | ✅ **29 条动作**（27 可用 / 2 标 unavailable），真报告端到端跑通 |
| 脚本接入 pipeline 接口 | ✅ 会写盘的脚本里，有接口的 **28 → 41**，没接口的 **46 → 34** |
| 脚本数缩减 | ✅ **117 → 98**，20 个转发壳折成一张声明表 |
| surgical 守卫 | ✅ 提醒改**硬拦**（exit 2），反向验证 11/11 |

> 口径说明（上一版这里的「67 / 37 / 30」是错的，它用的尺子和现在对不上，已作废）：
> 「会写盘」= 源码里出现 `.save(` 或 `ZipFile`；同一段代码在折前 commit（`2fc6879`）
> 和现在各跑一遍得到 74 → 75 个（分母基本没动），有接口的 28 → 41。

**剩下 34 个没接口的，都有各自的理由**（不是漏了）：

| 类别 | 例子 | 为什么不接 |
|---|---|---|
| 不是 docx | `convert` `xlsx_*` `pptx_*` `pdf_to_docx` `md_to_audiobook` | 根本不在这条轴上 |
| driver | `doc_dispatch` `docx_tools` `typeset_pipeline` `typeset_apply` | 它们是调别人的那一层 |
| 库 | `pipeline_lib` `bid_residue_lib` `shape_contract` `styles` `health` | 没有「对一份 docx 做一件事」的语义 |
| 不是单文档进出 | `split_by_h1` `combine` `body_replace` `extract_tables` `image_extract` `port_sections` | 一进多出 / 多进一出 |
| 要外部参照件 | `docx_apply_template` `docx_format_clone` `md_docx_template` `docx_chrome` | 是「照模板造新件」不是「调这份的版面」 |

**曾经只能串命令、本轮接进来的四个**：

| 脚本 | 怎么接的 |
|---|---|
| `center_images` · `line_spacing` | 抽出共用实现 + `apply_path`，排在 path 段末尾 |
| `restyle` · `sync_toc` | 同上，但放进**新开的 `path-pre` 段** —— 见下面 §五 C |

**还进不来的（各有原因，不是漏了）**：

| 脚本 | 为什么 |
|---|---|
| `fix_styleset` | 包内相对 import，`load_step` 加载不到（改动前后都这样） |
| `outline` | 三个互斥子动作，与 `renumber_headings` 抢同一件事 |
| `docx_text_formatter` | 唯一一个选项不是标量的（scope 白名单），spec 表达不出来 |
| `delete_chapter` · `delete_table_rows` | **刻意不收**：破坏性动作混进「调版面」的配置里是反模式 |
| `fix_heading_disorder` · `reorder_heading_blocks` | 缺 `apply_body_styles` 依赖，引擎已标 `unavailable`，被启用时明确报错 |
| `split_by_h1` · `combine` · `body_replace` | 输入输出不是「单文档进出」（一进多出 / 多进一出 / 壳+料→新文件） |

---

## 五 · 顺带挖出来的五个坑（都不是自己发现的，是闸门和核验逼出来的）

### A 成品里留着真实评审批注，Word 看不见，闸门也抓不到

`strip_revisions` 有两半（清正文锚点 + 清 `comments.xml`），引擎只声明了前一半。

| | comments.xml | 批注定义 | 正文锚点 |
|---|---|---:|---:|
| 原件 | 4178 B | 2 | 2 |
| 逐条敲 CLI | 140 B | 0 | 0 |
| spec 引擎（修前） | 4178 B | **2** | 0 |

留下的是 `作者=dh9304 / 2025-12-08 / 内部核算口径`。没锚点 = Word 里看不见，
`bid_residue_scan` 修前修后都报 FAIL 75、一个字不差 —— **没有任何闸门抓得到，随包发出去**。

修法：`Action.tail_path`，两半绑定，spec 里仍只有一个开关。修后 **批注 0 · 锚点 0**。

### B 纯只读的体检 spec 也把 59MB 交付件整包重写 + 落一个 59MB 备份

修法：`Action.readonly` + 跳过 lsof / 备份 / 存盘三处。
修后 **rc=0 · md5 同 · mtime 同 · bak 0**。

### C `restyle` / `sync_toc` 放进排版链里会**必然失效**，而且不报错

这两个是**按段落文本**去 golden 里找对应段的。而链上 `convert_chapter_format` /
`renumber_headings` / `number_captions` / `fix_superscript_refs` 每一个都会改文本 ——
排在它们后面，匹配率塌到接近零，然后安静地说一句「无需修改」。

修法：新开 **`path-pre` 段**（整条链之前，zip 级）。别为了省一个 phase 把它们塞进
`path` —— 那等于让它们必然失效。

### D `strip_bookmarks` 删了 77 个书签，摘要写「只读/无 changed 计数」

因为报告只认 `changed` 这一个键，而 `strip_*` 一族用的是领域名（`bookmarks_removed`）。
报告说不出自己改了多少，就等于没有报告。

修法：`Action.count_keys` 显式声明；**会写盘却报不出计数一律 rc=1**（反向验过：
故意把键写错 → rc=1 并列出脚本实际返回的键）。不靠启发式去猜哪个是主计数 ——
猜错的表现是**报一个错的数**，比不报还糟。

### E `table delete-rows` 从写下那天起就是坏的

它只 `set_defaults(_sub_target=...)` 没设 `func=`，敲下去只得到
`[docx_cli.py] no handler for table (incomplete subcommand?)`；旁边那个本该接住它的
`handle()` 从来没有任何人调用过。**18 份手写的转发代码里藏着一个死子命令，没有任何
闸门看得见。** 收进声明表之后每个 target 必然拿到 `func`，现在能跑了。

---

## 六 · 20 个转发壳 → 一张表

`sub/` 下曾有 20 个文件只干一件事：声明 argparse 子命令，然后 `exec_script` 转发。
它们不碰 docx、没有业务逻辑，2265 行的信息量 = 一张「组名 / 子命令 / 实现脚本 /
选项 / 怎么转发」的表。

代价不是「文件多」这么轻，是**同一件事写了 20 遍**：加一个标准选项要改 20 处；
转发规则变一次要核对 20 处；新加组只能复制粘贴上一个组。而且错了没人发现（见 §五 E）。

现在 `sub/_groups.py` 一张 `GROUPS` 表。`Opt` 里**同时**写「怎么声明给 argparse」和
「怎么转发给实现脚本」—— 这两件事原来分居 `register()` 与 `_run()` 两处，于是
「加了选项忘了转发」是这类壳最常见的 bug：用户传了参数、脚本收不到、静默按默认值跑。

四处不规则的也建模成了数据，没有留回调（一留回调，表就会重新长回代码）：

| | 怎么表达的 |
|---|---|
| `table borders --center` 会接着跑第二个脚本 | `chain=("center", "center")` |
| `chrome --validate` 是验证模式，其余选项一概不发 | `fwd="short_circuit"` |
| `section read --list` 顶掉位置参数 `query` | `suppressed_by="list_headings"` |
| `split body-replace` 的互斥对（默认真值那支不转发） | `mx="h1"` + `fwd="when_false"` |

### 怎么证明接口没坏：两道闸门，缺一不可

| 闸门 | 证明什么 | 结果 |
|---|---|---|
| `tools/cli_surface.py` | 子命令树指纹（选项名 / dest / nargs / 默认值 / required / choices / 子命令组的 dest+required+metavar / 互斥组） | 折前折后 **128 个逐字节相同** |
| `tools/cli_forward_probe.py` | 拦掉 `exec_script`，录下每条子命令**真正转发出去的 argv** | 50 个样本，折前能跑的 **48 条完全一致** |

**只跑第一个不够**：接口对了 argv 错了 = 命令能敲、跑出来的东西不对，且不报错。

加强闸门本身也立刻见效：给 `cli_surface` 补上「子命令组的 metavar」当场就抓到我把
`header-footer` 的 `<action>` 写成了 `<target>` —— 这种差异不报错、只在 `--help` 里
显示不同，肉眼永远看不出来。

被折的 20 个文件在 `~/.Trash/doctools-shim-fold-20260730/`（带 MANIFEST），没有 `rm`。

---

## 七 · surgical 守卫：提醒 → 硬拦

`docx-surgical-advisory.sh` → `docx-surgical-guard.sh`（改名不是洁癖：一个名字叫
advisory、行为是 block 的守卫，下次有人读名字做判断时必然判错）。

**两档，不是一档**：

| | 什么情况 | 行为 |
|---|---|---|
| `exit 2` | 读到那个 `.py`，它 import python-docx + 调 `.save()` + 没 `import docx_safe_save`；或命令串里内联了 python-docx 而没带收口 | **拦** |
| `exit 0` + 警告 | 命令里有 `.py` 但打不开（相对路径 / cwd 不同） | 出声，放行 |

「判不了」不一起拦，是这次唯一值得写清楚的取舍：**拦一条其实无辜的命令，代价不是
多打一行字，是用户会把 `SURGICAL_OK=1` 写进环境永久关掉这个守卫** —— 那时它对所有
真风险一起瞎。底线是任何一档都不静默。

**去掉了 advisory 时代的去重**：硬拦下去重 = 第二条同样危险的命令被静默放行，
而那正是最该拦的时候。（2026-07-28 还实测过一次去重把守卫彻底弄哑：sid 取不到 →
键恒为 `unknown` → `/tmp` 里一个 `.seen` 文件让全机所有会话永久静默。）

反向验证 **11/11**：该拦的 4 条（含同一条连跑两次，证明无去重漏洞）、不该拦的 6 条、
判不了的 1 条。另拿 doctools **4 个真脚本走生产路径**实测放行，再把其中一个的收口那行
摘掉 → 立刻 rc=2。夹具通过不等于生产路径通过。

---

## 八 · 现在怎么用

```bash
# 排版：一份 spec 跑完
python3 ~/Dev/tools/doctools/scripts/document/typeset_apply.py 报告.docx \
    --spec ~/Dev/tools/doctools/config/spec-examples/report-generic.yaml

python3 …/typeset_apply.py --list          # 29 条动作 + 每条为什么排在这个位置
python3 …/typeset_apply.py --dump-schema   # spec 全量字段说明（从 ACTIONS 派生）

# 量任一条命令的炸开面（surgical 的标准就是这个数字）
python3 ~/Dev/tools/doctools/tools/blast_radius.py run 报告.docx -- <命令，{docx} 占位>

# 守卫：谁用 python-docx 存盘却没挂收口
python3 ~/Dev/tools/doctools/tools/check_docx_collar.py

# 全仓 134 个脚本谁调谁（可点双链页面）
python3 ~/Dev/tools/doctools/tools/script_graph.py --open
```

逃生：`DOCX_GRAFT_OFF=1` 退回裸存盘 · `DOCX_GRAFT_QUIET=1` 不打 stderr 那行。

---

## 九 · 还没做的

**§七 那四条已全部清完。** 剩下的都是「知道为什么不做」而不是「还没轮到」：

| | 为什么不做 |
|---|---|
| `fix_styleset` 接进 spec | 包内相对 import，`load_step` 加载不到。要接得先动它的 import 结构，那是改一个跑得好好的 2007 行脚本的业务逻辑，不值 |
| `outline` 接进 spec | 三个互斥子动作，与 `renumber_headings` 抢同一件事。硬接进去 = 让用户配得出两个互相打架的动作 |
| `delete_chapter` / `delete_table_rows` 接进 spec | **刻意不收**：破坏性动作混进「调版面」的配置里是反模式 |
| 合并实现脚本 | 查过四对「疑似重复血统」，**没有一对是真重复**（两个操作 md 不操 docx，两个分工写在文件里）。合并它们只会把 42,059 行搬个地方 |

真要再缩，只剩「退役某些功能」这一条路 —— 那是产品决定，不是重构。

---

## 十 · 每一条数字怎么来的

| 断言 | 怎么验的 |
|---|---|
| 128 个子命令折前折后相同 | `tools/cli_surface.py` 在 `git worktree`（折前 commit）与 HEAD 各跑一次，`diff` 为空 |
| 48 条转发 argv 相同 | `tools/cli_forward_probe.py` 同上，另 2 条是折前根本跑不起来的死子命令 |
| 117 → 98 个脚本 | 同一段 AST 扫描在两个 commit 各跑一次 |
| 28 → 41 个有接口 | 同上（口径写在 §四） |
| 计数键闸门有效 | 故意把 `count_keys` 写成不存在的键 → rc=1 且列出脚本实际返回的键 |
| 守卫真能拦 | 11 条夹具 + 4 个真脚本走生产路径 + 摘掉收口那行立刻 rc=2 |
| 没坏东西 | 74 个单测全绿 · 收口守卫 36/36 · 134 脚本 0 孤儿 · 真报告端到端 rc=0 |

---

## 十一 · 第二轮（2026-07-30 当天）：退役 5 个零消费孤儿

上面 §九 说「真要再缩只剩退役功能这条路」—— 用户看完拍板走了这条路。换了四个角度
重查（入口分发层重叠 / 老单体 vs sub 命令 / 全域零引用 / 血统重复），找到 5 个
**全域（Dev/Work/Apps/Archives/Money/VPS）零消费**的孤儿，批准后退役：

| 脚本 | 行数 | 判死证据 |
|---|---|---|
| `report_quality_check.py` | 953 | 全域零引用；2026-05-21 后零改动；它 import 的 bullet_to_paragraph 同退，不同退就成「依赖已进 Trash 的静默降级半残件」 |
| `gen_report.py` | 484 | eco-flow 报告线在 Work/Archives 零痕迹；2026-04-16 生，此后只有机械改 |
| `bullet_to_paragraph.py` | 364 | 2026-05-22 后零改动；唯一消费者是同退的 report_quality_check |
| `review_deep.py` | 317 | 同 eco-flow 线；唯一入口 docx_cli review 转发 |
| `docx_qa.py` | 301 | 自称 `/docx finalize` 后端 —— `/docx` skill 已不存在；全域零引用 |

**账（口径写死，免得下次又对不上）**：

| 尺子 | before | after |
|---|---|---|
| §零 同尺（`scripts/**` 非测试、原 docx 判据，inv.json 存档在会话 scratchpad） | 98 | **96**（5 个里只有 docx_qa / review_deep 在这个分母里，另 3 个原判 docx=False：主要操 md） |
| 全仓透明尺（`find . -name '*.py'` 非测试 + `grep -li docx`） | 113 | **108** |
| 仓内实际文件 / 行数 | — | **−5 个 / −2,419 行** |
| docx_cli 子命令树（cli_surface 嵌套计数） | 128 | **125**（摘 bullet / quality-check / review 三条 legacy 转发） |

**验证**：cli_surface 折前折后 diff = 恰好只少这 3 个子命令、零误伤；forward probe 50/50；
`script_graph` 134 → 129 个脚本 0 孤儿；收口守卫 36/36；74 单测全绿。
去向：`~/.Trash/doctools-orphans-20260730/`（带 MANIFEST），git 历史可捞回。

顺手修正的一处 §十 遗留：「117 → 98 的尺子」当时没写进文档，本轮从会话 scratchpad 的
`inv.json` 反查出来了 —— 范围 `scripts/**`、非测试、按源码 docx 特征判旗标。教训同 §四：
**数字必须连尺子一起落盘**，只落数字的口径三天就丢。
