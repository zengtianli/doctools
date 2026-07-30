# docx 排版：从「串命令」到「一份 spec」 · 进行中

> 状态：workflow `wf_e7d20578-92d` 运行中（9 agent / 四相）。本文件先落骨架，跑完回填结果。

## 一 · 这件事的起点是用户的一个判断

原话：「写报告标书这些内容，其实主要是思维层面的。docx 的操作就 header footer、序号、
text format、图片、表格、图片表格题注，还有审阅模式，就这些啊。内容方面就很多了……
内容想好，config json yaml txt md 搞好就行了，最后呈现是 docx 就行了。**不用另起炉灶
搞 docx 针对每个内容**。」

翻译成架构：

```
内容层    md / yaml / json          ← 无限种（投标、报告、论文…）
   ↓
版式层    一份 spec                  ← 有限：七件事
   ↓
渲染层    一个引擎，读 spec 渲成 docx
```

**新增一类交付 = 新写一份 spec，不是新写一套脚本。**

## 二 · 核过的账（不是感觉，是数出来的）

`scripts/document` 124 个脚本 / 46,764 行，按用户列的七件事归类：

| 操作 | 脚本 | 行数 |
|---|---:|---:|
| 页眉页脚 | 2 | 318 |
| 序号编号 | 12 | 3,800 |
| 文字格式 | 11 | 5,651 |
| 图片 | 8 | 2,693 |
| 表格 | 7 | 2,190 |
| 题注 | 3 | 371 |
| 审阅修订 | 6 | 1,336 |
| **合计** | **49** | **16,359** |

用户没列但客观存在的第八件：**结构增删**（拆章/合章/删章/搬块）15 个 / 5,093 行。

## 三 · 探路发现：七件事不是一类，是两类

这条改变了方案。判据是**参数长什么样**：

| | 特征 | 例子 | 能否 spec 化 |
|---|---|---|---|
| **设值类** | 参数就是值（字体/字号/颜色/对齐/文字/层级） | `normalize_fonts` `add_header_footer` `set_table_borders` `freeze_heading_numbers` | 能 |
| **识别类** | 参数只有 `--dry-run/--no-backup/--report`，**一个值都没有** | `number_captions` `renumber_headings` `pair_table_captions` `set_table_align` | **不能** |

识别类不可配的原因很实在：`number_captions` 397 行里全是 `has_nearby_table` /
`has_nearby_drawing` / `next_nonempty_idx` / `parse_chapter` —— 它在**猜**「这一段是不是表名」。
这是识别问题不是配置问题，没有值可配。`pair_table_captions` 更直白：它要一个 `--decision`
JSON，等于说「我猜不准，你先给判断结果」。

所以 spec 的边界是：**配置设值类 + 开关识别类**，不是「一份 spec 取代所有脚本」。

## 四 · 一个意外的好消息

**17 个脚本已经是「引擎 + CLI 壳」形态**（有独立的 `apply(doc, args)`，`main()` 只是壳）。
意味着引擎不用 subprocess 串命令，可以直接 import 调 `apply()` —— 同一进程、同一份 doc、
**开一次存一次**。这跟前两天做的 surgical 收口同向：少开合一次 = 少一次重打包机会。

## 五 · 工单（实测出来的，不是估的）

27 个会写盘的排版脚本：

| | 数量 | 处理 |
|---|---:|---|
| 已有 `apply()` | 12 | 当样板 |
| 需补 `apply()` | 12 | 本轮补 |
| **不该套这个模子** | 3 | `split_by_h1`（一进多出）· `combine`（多进一出）· `body_replace`（壳+料→新文件） |

## 六 · 本轮的硬要求（写死在工单里）

- **顺序不给 spec 控** —— 有依赖（先冻结域再改字体、先配对题注再编号），顺序写死在引擎里，
  理由写进注释。spec 只管「跑不跑」和参数。
- **spec 里出现不认识的键 → 报错退出，禁静默忽略**。静默忽略配置项是最难查的一类 bug：
  用户以为配了，其实没生效。
- 空 spec / 文件不存在 / yaml 语法错 → 一律非 0 退出。
- `apply()` 契约：**不开文件、不存盘、不备份、不 sys.exit**，那些都是调用方的事。
- 本轮**只做结构搬移，不改业务逻辑**；不删任何脚本；agent 禁 git 操作。

## 七 · 待回填

- [ ] schema：设值类几节、识别类几个开关
- [ ] 12 个 `apply()` 的实际完成情况与未做项
- [ ] `typeset_apply.py` 的对拍结果（等价性 / 开合次数 / 炸开面）
- [ ] **bid_* 案例的答案**：「新增一类交付 = 新写一份 spec」在标书上成立到什么程度（百分比）
- [ ] 对抗核验的总判断：能用吗、边界在哪

---

## 八 · spec 层落地（2026-07-30）

### 产物

| 文件 | 是什么 |
|---|---|
| `scripts/document/typeset_apply.py` | 引擎。`ACTIONS` 表 = 顺序 + kind + phase + 选项默认值 + 每条的 `why_here`，是本层唯一事实源 |
| `config/docx_spec.schema.yaml` | **派生物**，`typeset_apply.py --dump-schema > 它`。别手改——手写 schema 必然和 ACTIONS 漂开，而「文档说有这个选项、引擎其实不认」正是 spec 层最要命的谎 |
| `config/spec-examples/report-generic.yaml` | 通用技术报告版式的最小可用 spec（19 个动作） |

旧的 `config/docx_spec.example.yaml` 已被上面两份取代 → `~/.Trash/doctools-spec-example-20260730/`。

### (a) 薄层 vs (b) 自带执行：选 (b)，但只自带执行循环

加载 / lsof / 备份命名三件事**复用** `sub/pipeline_lib.py`，没重写。没走
`run_pipeline` 的三条具体理由（写进模块 docstring）：① 它对全部 step 共用同一个扁平
Namespace，选项会跨动作串味，而 spec 层的全部意义就是「选项隔离 + 未知选项报错」；
② 它把 step 异常吞成 `{"error":…}` **继续跑、照常存盘**，排版链上这等于产出半新半旧的
交付件；③ 它只有 doc→path 两段，表达不了「`normalize_fonts` 必须跑在 `freeze_*` 之后」。

### 两段 → 三段（doc-pre → save → path → 重开 → doc-post）

起因是上一轮的实测阻断：三个脚本补了 `apply()` 后 `load_step` 优先取它，kind 从 path
翻成 doc，与 ACTIONS 声明对不上，引擎 fail-closed exit 2。**没有关掉那个守卫**（反向验证
过：把 kind 谎报回 path，引擎仍然 exit 2 并给出同一条消息），而是消掉了漂移本身：

- `normalize_fonts` → `doc-post`（唯一真需要重开的：它必须看见 freeze 冻出来的编号 run）
- `set_table_borders` / `set_table_align` → `doc-pre`（与文本轴无依赖，白捡，不用重开）

比「让 kind 具备强制力、继续走 apply_path」少 3 次文件开合。

### 调查者六问的回答（全部写进代码注释，不悬着）

| 问 | 答 |
|---|---|
| Q1 kind 漂移怎么办 | 三段模型，见上 |
| Q2 `h1_base` 接不接 | **接了**。这不是新业务逻辑，是 apply() 抽取时**漏搬**的参数（main 一直是 `plan_renumber(doc, h1_base)`）。实测 `h1_base=10` → 首两章变 10/11，缺省与漏搬前逐字相同 |
| Q3 图片节要不要收 | 收 `docx_apply_image_caption`（位置照建议：`number_captions` 之后、`add_header_footer` 之前）。`center_images` 仍不收——两个接口都没有 |
| Q4 `strip_revisions` 双 kind | 不做。为一个脚本把「一动作一接口」改成「一动作两接口」，kind 校验/phase 归属/报告合并三处都要改，换来一个 bool |
| Q5 yaml 分节 | 不做，`actions` 保持扁平。分节不承载执行含义，多一层只多一处能拼错的地方；分节用注释表达 |
| Q6 顶层 `profile` 继承 | 不做。会打破「选项不跨动作泄漏」这条刻意设计，且当前收编的动作**一个都不吃 profile** |

另外 `fix_styleset` 实测在本引擎的加载路径下 `ImportError: attempted relative import`，
`load_step` 根本加载不到 —— 收编前要先让它能独立加载，那是改它的业务代码，本轮不动。

### 自验（夹具复制到 scratchpad，原件未动）

夹具 = `00金华江流域…统稿12.8-缙云参考.docx`（58.9MB / 301 部件 / 1254 段 / 57 表 / 17 节）。

- `--dry-run`：rc=0，跑完 md5 与原件一致（dry-run 真的没写盘）
- 真跑：rc=0；段 1254→1246（`strip_empty_captions` 删了 8 个空题注段）、表 57→57、节 17→17；
  python-docx 能重新打开；H1 变成 `1 区域概况 / 2 … / 6 …`
- **炸开面**（`tools/blast_radius.py run`）：
  - 全量 spec（19 动作，含页眉页脚）：301 → 311 部件，删 0 · **新增 10** · **改写 28**。
    新增 10 = `add_header_footer` 给 17 个节补的 header/footer 部件；改写里 24 个是
    既有 header/footer + `[Content_Types].xml` + `document.xml.rels`，全部可归因到这一个动作。
  - 同一份 spec 关掉 `add_header_footer`（18 动作）：**301 → 301，删 0 · 新增 0 · 改写 2**
    （`word/document.xml` + `word/styles.xml`）。**这就是引擎本身的炸开面**：18 个排版动作
    跑完只碰 2 个部件。对照：subprocess driver 一步一存盘，光是「什么都不改地开合一次」
    python-docx 就要重写 60 个部件。
- 7 条 fail-closed 全部 rc=1：顶层拼错 / 不认识的动作 / 不认识的选项 / 等线字体 /
  一个动作都没启用 / 空文件 / spec 不存在；`unavailable` 动作被启用时报出具体原因。
- 测试：`74 passed`（与改动前同）。

---

## 九 · 对抗核验抓到的两个问题，主会话已修（2026-07-30 收尾）

核验那一棒的价值全在这里 —— 前面四棒都自报「跑通」，它抓出两条**自述里没有的**问题。
两条都由主会话（不是 agent）修，因为都涉及「该不该做」的判断，不是机械执行。

### A（会咬人）成品里留着真实评审批注，Word 看不见，闸门也抓不到

`strip_revisions` 的 `main()` 同时跑 `apply(doc)`（清正文锚点）和 `apply_path()`
（清空 comments.xml）。ACTIONS 声明 `kind=doc`，**只跑了前一半**。

|  | comments.xml | `w:comment` 数 | 正文 commentReference |
|---|---|---:|---:|
| 原件 | 4178 B | 2 | 2 |
| 逐条敲 CLI | 140 B | 0 | 0 |
| spec 引擎（修前） | 4178 B | **2** | 0 |

留下来的是 `作者=dh9304 / 2025-12-08 / 正文是内部核算口径`。**没锚点 = Word 里看不见**，
`bid_residue_scan` 修前修后都报 FAIL 75、一个字不差 —— 没有任何闸门抓得到，随包发出去。

引擎作者在 Q4 注释里写过这个取舍（"清 comments 那半边在 apply_path 里，doc 阶段够不着，
为一个 bool 改 Action 结构不值"）。**那个取舍是错的**：用户在 spec 里写 `strip_revisions`，
意思绝不可能是「清正文但把批注定义留着」。这不是少配一个 bool，是一个动作被拆掉一半
还宣称做完了。

修法：`Action.tail_path` —— 一个 doc 动作可以声明它在 path 段还有「另一半」，
两半绑定，spec 里仍然只有一个开关，**不给用户漏掉的机会**。
反向验证：修后 `批注定义 0 条 · 正文锚点 0 个`，段 1254→1246 / 表 57 / 节 17 / CRC ok。

### B 纯只读的 spec 也把 59MB 交付件整包重写 + 落一个 59MB 的 .bak

`run()` 原来只要 `pre_plan` 非空就无条件 `doc.save()`，不看有没有动过树。
语义上没坏（收口救了 C14N），但 mtime 变、字节变、交付目录里多一份 59MB 垃圾，
而且 `lsof` 门会让「Word 开着时想跑个只读体检」直接 rc=2。

修法：`Action.readonly` + 三处判据（跳过 lsof、跳过备份、跳过存盘）。
反向验证：纯 audit spec → `rc=0 · md5 同 · mtime 同 · bak 数 0`。

### 修 A 时我自己引入的回归（当场抓住）

加了 tail 之后 `--dry-run` **开始真写盘**了 —— `strip_revisions.apply_path` 根本没有
dry_run 守卫，无条件 `shutil.move(tmp, docx)`。

我先前用一条 `sed -n "/def apply_path/,/^def main/p" | grep -c dry_run` 得出「它认 dry_run」，
**那条命令把范围切进了 main()**，数到的 4 次全在 main 里。**用粗糙的命令验证 = 得到错误答案**，
这次是自己踩的。修法是在引擎里挡住（tail 在 dry-run 时不跑并记进报告），而不是去改那个
跑得好好的脚本 —— tail 是本引擎新加的调用路径，这条路径的安全由本引擎负责。

### 收尾自验（全部本轮实跑）

| | 结果 |
|---|---|
| 只读 spec | rc=0 · md5 同 · mtime 同 · bak 0 |
| `--dry-run` | rc=0 · md5 同 |
| 真跑 | rc=0 · 批注 0 · 锚点 0 · 段 1246 / 表 57 / 节 17 · CRC ok |
| fail-closed 三条（不认识的动作 / 不认识的选项 / 空 actions） | 全 rc=1 |
| pytest | 74 passed |
| 收口守卫 | 36/36 |
| `--dump-schema` vs 仓里 schema | 逐字一致（派生关系是真的） |

> 有一次 `rc=2 docx 被占用` 是环境干扰：Spotlight 的 `mdworker_` 在索引刚复制的 59MB 文件，
> `lsof_check` 把只读句柄也算占用。等索引完重跑即 rc=0。这条是 `lsof_check` 的既有行为，
> 不在本轮范围内。

## 十 · 核验报的、本轮**没修**的（如实记）

| | 为什么不修 |
|---|---|
| C 报告失真：`strip_bookmarks` 删了 77 个书签，摘要打「只读/无 changed 计数」 | `_record()` 只认 `changed` 键，而这几个脚本返回 `bookmarks_removed` 等。要么改脚本返回值、要么在 Action 上声明计数键名 —— 都属于「再动一轮 12 个脚本」，本轮不叠加 |
| D 两颗历史地雷：`docx_qa.apply(docx, opsfile, no_backup)` / `md_merge_impl.apply(md_file, docx_file, …)` 签名不是这套契约，函数体里开文件存盘 | **不是本轮引入**，但本轮把「顶层 apply = 统一契约」变成了全仓约定，这两处成了显式反例。`load_step` 只看 `hasattr(mod,"apply")` 会把它们判成 kind=doc 然后 TypeError。修法是改名或给 `load_step` 加签名校验 |
| `report-generic.yaml` 实际启用 20 个动作，几处自述写 19 | 数字口径不一致，不影响行为 |

## 十一 · 边界（这些排版动作现在仍然只能串命令）

`center_images` / `line_spacing`（两个接口都没有）· `fix_styleset`（包内相对 import，
`load_step` 加载不到）· `outline`（三个互斥子动作，与 `renumber_headings` 抢同一件事）·
`docx_text_formatter`（唯一一个选项不是标量的动作，scope 白名单在 spec 里表达不出来）·
`delete_chapter` / `delete_table_rows`（**刻意不收**：破坏性动作混进「调版面」的配置里是反模式）·
`fix_heading_disorder` / `reorder_heading_blocks`（缺 `apply_body_styles` 依赖，引擎已标
`unavailable`，被启用时明确报错）。

**闸门语义整类表达不出来**：实测 `audit_caption_outline` 报 `wrong_style_count=78`、
`audit_table_pairing` 报 `orphan_tbls=1 / mismatch=50`，引擎照样 rc=0。
spec 说的是「改成什么样」，闸门说的是「不长这样就别发」。
