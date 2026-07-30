---
title: docx 工具链 before / after
slug: before-after
project: doctools
date: 2026-07-30
out: /Users/tianli/Dev/tools/doctools/reports/before-after.html
---

> 先回答「全部做好了吗」：**没有全部。** 两件事做完了（存盘收口、spec 引擎），
> 一件事只做了一半（脚本接入）。下面每个数字都是本轮实跑出来的，不是估的。

---

## 零 · 一句话

| | before | after |
|---|---|---|
| 改一份报告的版面 | 手敲 **20 条命令**，每条自己开文件自己存盘 | 写**一份 yaml**，跑一条命令 |
| python-docx 存一次盘 | 重写 **~60 个部件**（语义只变 1 个） | 重写 **1~2 个** |
| 跑完一轮排版 | 存盘 **15 次** · 进程 **20 个** · **100 秒** | 存盘 **6 次** · 进程 **1 个** · **33 秒** |
| 加一条横切改动（如收口） | 改 **35 处** | 改 **1 处** |

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

## 四 · 完成度：如实说

| | 状态 |
|---|---|
| 存盘收口 | ✅ **做完**。59 个脚本 / 11 个 repo 全部接入，守卫 36/36 绿，commit 期硬拦 |
| spec 引擎 | ✅ **做完**。25 条动作（23 可用 / 2 标 unavailable），一份真报告端到端跑通 |
| 脚本接入 pipeline 接口 | ⚠️ **一半**。仓里 **67 个会写盘的脚本，37 个有 `apply`/`apply_path`，30 个没有** |

**仍然只能串命令的**（各有原因，不是漏了）：

| 脚本 | 为什么进不来 |
|---|---|
| `center_images` · `line_spacing` · `restyle` · `sync_toc` | 走 lxml 直读 zip，两个接口一个都没有 |
| `fix_styleset` | 包内相对 import，`load_step` 加载不到（改动前后都这样） |
| `outline` | 三个互斥子动作，与 `renumber_headings` 抢同一件事 |
| `docx_text_formatter` | 唯一一个选项不是标量的（scope 白名单），spec 表达不出来 |
| `delete_chapter` · `delete_table_rows` | **刻意不收**：破坏性动作混进「调版面」的配置里是反模式 |
| `fix_heading_disorder` · `reorder_heading_blocks` | 缺 `apply_body_styles` 依赖，引擎已标 `unavailable`，被启用时明确报错 |
| `split_by_h1` · `combine` · `body_replace` | 输入输出不是「单文档进出」（一进多出 / 多进一出 / 壳+料→新文件） |

---

## 五 · 顺带修好的两件事（对抗核验抓的，不是自己发现的）

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

---

## 六 · 现在怎么用

```bash
# 排版：一份 spec 跑完
python3 ~/Dev/tools/doctools/scripts/document/typeset_apply.py 报告.docx \
    --spec ~/Dev/tools/doctools/config/spec-examples/report-generic.yaml

python3 …/typeset_apply.py --list          # 25 条动作 + 每条为什么排在这个位置
python3 …/typeset_apply.py --dump-schema   # spec 全量字段说明（从 ACTIONS 派生）

# 量任一条命令的炸开面（surgical 的标准就是这个数字）
python3 ~/Dev/tools/doctools/tools/blast_radius.py run 报告.docx -- <命令，{docx} 占位>

# 守卫：谁用 python-docx 存盘却没挂收口
python3 ~/Dev/tools/doctools/tools/check_docx_collar.py

# 全仓 151 个脚本谁调谁（可点双链页面）
python3 ~/Dev/tools/doctools/tools/script_graph.py --open
```

逃生：`DOCX_GRAFT_OFF=1` 退回裸存盘 · `DOCX_GRAFT_QUIET=1` 不打 stderr 那行。

---

## 七 · 还没做的（按性价比排）

1. **30 个没接口的脚本** —— 其中 `center_images` / `line_spacing` / `restyle` / `sync_toc`
   是真排版动作，值得接；其余多数是 driver、库、或本来就不该接的。
2. **报告计数键名不统一** —— `strip_bookmarks` 删了 77 个书签，摘要却打「只读/无 changed 计数」。
3. **两处历史 `apply` 签名违规**（`docx_qa` / `md_merge_impl`）—— 非本轮引入，但本轮把
   「顶层 apply = 统一契约」变成全仓约定后，它们成了显式反例，`load_step` 会 TypeError。
4. **`docx-surgical-advisory` 从提醒改成硬拦** —— 破坏性变更，等你点头。
