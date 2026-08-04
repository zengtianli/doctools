# doctools 冒烟轴 / CI / 可分发包 —— 交付报告

> 2026-08-02 · 施工 3 段（smoke 93 条 / GitHub Actions 闸门 / wheel 自包含）+ 独立核验 2 路
> 涉及 commit：`07eadfb`（smoke）· `6cafe3b`（文档）· `ee74624`（CI）· `f88218b`（dist）—— **全部未 push**
> 本文一切「rc=」与命令输出，凡标「本轮复核」的是写这份报告时我自己敲出来的；标「施工方报」「核验方报」的是转述，已注明来源。

---

## 一句话结论

**本机这条路是实的，别的机器那条路是空的。** 93 条动词的冒烟轴 + 5 道闸门在本机全绿、
且反向验证过能转红；但 **CI 从来没绿过一次**（第一个装依赖步骤在任何非本机的机器上必红），
**wheel 装上去 49 个顶层动词里 14 个是死的**。所以：**不达标，不宣布达标**。差什么见第五节。

---

## 一、覆盖率的真数字

### 1.1 冒烟轴

| 口径 | 数 | 出处 |
|---|---|---|
| `_verb_specs.py` 表项 | **93** | 本轮复核 `len(SPECS)` = 93 |
| 真跑 | **91** | 本轮复核 `check_smoke_coverage.py` rc=0：「真跑 91 条 / skip 2 条 / 共 93 条（97.8%）」 |
| skip | **2** | 同上 |
| pytest 实测 | **92 passed, 2 skipped, 20.9s** | 施工方报（92 = 91 条动词 + 1 条表自检） |

**三个数都对，报的时候必须带口径**（沿用仓内既有约定）：CLI 节点总数 126 · 叶子路径 101 ·
按 `id(parser)` 去重后 **93** 条可跑动词（另 8 条是别名，共用同一 parser，不单独占表项）。
93 是冒烟轴的分母。探明阶段报的「可跑 101 / 不可跑 3」用的是叶子路径口径，**和 93 不是同一把尺**，
两处数字不要混着引用。

### 1.2 skip 的 2 条 —— 理由分类

两条都**实测能跑**（不是跑不通被绕开），skip 理由写进表且 `-rs` 会打印：

| 动词 | 分类 | 理由（表内原文，本轮复核 `check_smoke_coverage` 输出可见） |
|---|---|---|
| `para render` | 外部二进制 + sandbox 外写盘 | 起 LibreOffice(soffice) 转 PDF，约 6s，缓存落 `~/.cache/doctools/` |
| `scan-sensitive` | 非确定性 + 真实成本 | 逐 md 调 LLM（`lib/llm_client.py` → `claude -p`），每次执行有模型成本 |

### 1.3 探明阶段标为「不可跑」的 3 条

| 条目 | 真实性质 |
|---|---|
| `track compare` | **未实现的存根** —— rc=1，stdout 唯一一行「compare 功能将在 v2 实现。」 |
| `md-to-docx extract` | 跑得通，但无条件往 **cwd** 写 `heading_styles.xml`/`styles_info.txt` 且无 `-o`；runner cwd 在仓根 → 会往仓里落未跟踪文件 |
| `md frontmatter` 真跑模式 | 逐 md 调 LLM 并**原地改写**加 YAML；表里收的是 `--dry-run` 形态 |

### 1.4 覆盖率数字与实测强度不匹配（核验方 finding，未修）

97.8% 这个数只证明「进程被拉起 + rc 对上」，**不区分「跑了活」与「零对象空跑」**。
核验方扫 `note` 字段，**18/93（19%）自认零对象**（`style caption` · `fix style-rename` ·
`renumber h4-figures` · `renumber-fig` · `outline` 三条 · `fix-ref` · `strip orphan-media` ·
`blocks relocate` · `chapter delete-empty-h1` · `revise-rules gen` · `compare-ref ref` ·
`audit bookmarks` · `audit table-pairing` · `audit-styleset style-coherence` ·
`health gate` · `para locate`）。

抽查发现这 18 条里**有的过度悲观**：`audit table-pairing` 实跑仓内 fixture 报出 2 条
`caption-name-content-mismatch`；`compare-ref ref` 实跑报「改为段 2 条」——这两条是真在干活。
所以这 18 条**需逐条实测分档，不能整体当账，也不能整体当绿**。

另一处同类缺口：`track` 在表里只占一行、只跑 `read` 分支，`review` / `compare` 两个分支
不在冒烟轴内（`compare` 是存根，`review` 由既有 190 条单测覆盖）。

---

## 二、核验发现（按 severity；**没修的明说没修**）

两路独立核验**均返回 `ok: false`**。以下按严重度排，最后一列是处置真相。

### critical

| # | 问题 | 证据（核验方在干净 ubuntu:24.04 容器实跑；括号内为我本轮的独立复核） | 修了吗 |
|---|---|---|---|
| C1 | **CI 从来没绿过。** `pyproject.toml` L86 的 force-include 源端 `"../dev/lib/parallel_contract.py"` 相对 project root 解析，GitHub runner 上兄弟目录 `dev/` 不存在；hatchling 对 **editable** 构建同样施加 force-include → `uv sync` 走 `build_editable` 直接 `FileNotFoundError` | `× Failed to build doctools @ file:///w` → `FileNotFoundError: Forced include not found: /dev/lib/parallel_contract.py`，rc=1。（本轮复核：`pyproject.toml:86` 确为该字面量；本机 `../dev` = `~/Dev/tools/dev` 存在，所以本机永远撞不到） | **未修** |
| C2 | **wheel 装上后 49 个顶层动词里 14 个是死的**：`extract`/`read`/`check`/`compare`/`diff`/`snapshot`/`track`/`format`/`template`/`text-fmt`/`md`/`md-to-docx`/`image-caption`/`scan-sensitive`。`_bundled/` 里的代码 import 了 6 个总部 `dev/lib` 模块，pyproject 只捞了 `parallel_contract` 一个 | 全新 venv、仓库外：`doctools extract fixture.docx` → `ModuleNotFoundError: No module named 'file_ops'`，rc=2；49 个动词逐个 `--help`：`ok=35 modulenotfound=14`。（**本轮复核 import 计数**：`file_ops` 10 · `finder` 10 · `display` 9 · `parallel_contract` 4 · `env` 1 · `usage_log` 1） | **未修**（核验方在容器副本里补 5 行 force-include 验证过 `still-broken=0 (was 14)`，**本机仓一个字节未动**） |

**C1 的时间线值得单独记一笔**：CI 是 `ee74624` 加的，破坏它的 force-include 是**下一个** commit
`f88218b` 打进去的。yml 文件头里「实测 `uv sync` rc=0，66 个包装齐」是 `f88218b` **之前**测的，
之后没人重跑。旁证：yml 注释记的 `script_graph` 是「104 个脚本 · 371 条引用」，容器里现在跑出来是
「105 · 373」（`f88218b` 新增了 `check_wheel_selfcontained.py`）——**实测记录停在上一个 commit**，
这条数字差就是它的指纹。

### high

| # | 问题 | 证据 | 修了吗 |
|---|---|---|---|
| H1 | **wheel 自包含门在结构上抓不到 C2，且它自称的最强档是假的。** `check_wheel_selfcontained.py` 的 install 层 checks 写死两条：`--version` 与 `verbs --fn convert`。`verbs` 是纯清单打印，**不加载任何实现模块**，对「实现进没进包」零分辨率。最刺眼的是 `verbs --fn convert` 打出来的正好是 `md-to-docx` 与 `md` ——**它把两个在这个 wheel 里已经死掉的动词的名字打印出来，然后判绿**。`--tier full` 的 docstring 承诺「补完整依赖 + 真 docx 上跑只读动词」，代码里 full 只把 pip 的 `--no-deps` 去掉，**checks 一模一样**，从来没有真 docx | 同一份 wheel 同一时刻：门打 `[RESULT] ✓ wheel 自包含（tier=full）`，而 `doctools md-to-docx --help` → `ModuleNotFoundError: No module named 'display'`。（**本轮复核** `tools/check_wheel_selfcontained.py:170`：`checks = [([exe,"--version"],0), ([exe,"verbs","--fn","convert"],0)]` 是常量；L211 的 full 分支只改 pip 参数不改 checks，与核验方描述完全一致） | **未修** |

### medium

| # | 问题 | 证据 | 修了吗 |
|---|---|---|---|
| M1 | **存在第三道 CI 里跑不了的闸门，而且它是静默的**：`check_wheel_selfcontained.py` 需要 `../dev` 与 `$HOME/Dev/tools/dev/lib`，runner 上两者都没有，`gates.yml` 从头到尾**一个字没提它**。于是「最可能被别人机器消费的产物（wheel）」= 全仓唯一零 CI 覆盖、且这个事实未被声明的产物 | `grep -in "wheel\|dist\|build" .github/workflows/gates.yml` → 无输出。CI 同形布局里跑该门：`[FAIL] 构建失败 rc=2` | **未修** |
| M2 | **`revise-rules gen` 是 100% 空跑，而 fixture 里那段注释宣称「已经避免了空跑」** —— 制造假信心。文件名短横对了，内容形状不对：`parse_md` 的 `SECTION_RE=^## 改动\s*\d+`、`PAIR_BLOCK_RE` 要求 `**原文**` 后**紧跟** `>` 行；`DRAFT_MD` 用的是 `## 意见 1` + 裸 `原文：` | 仓内 fixture：`待处理 MD: 1 份 / 0 条 / 最终 0 条规则`，产物 `out-a.json` = 2 字节 `[]`。核验方造合规草稿重跑 → `2 条 / 最终 2 条规则`，JSON 里是真的 find/replace 对。空跑只是从「0 份文件」下移到了「0 条规则」 | **未修** |
| M3 | **`renumber h4-figures` 声明 `mutates=False`，实为写盘动词**；本 fixture 上重写 0 个部件，整条写盘路径未被测到 | 仓内 fixture：`H4=1 图=0 表=0`，`只重写 0 个部件`，md5 不变 → `mutates=False` 由「什么都没干」满足。核验方造带真 Caption 样式题注的 `rich.docx`：`H4=1 图=4`，`只重写 1 个部件`，md5 `92901b65→8fa83a16`，正文实变 `图1-1→图1.1-1` 等 4 处 | **未修** |
| M4 | **fixture docstring 宣称含「题注」，但 6 条题注全是 `Normal` 样式**；凡按样式名识别题注的动词一律零对象 | 实测样式分布 `{'Normal':24,'Heading 1':3,'Heading 2':2,'Heading 3':2,'Heading 4':1}`，无任何 Caption 样式段。而 `styles.py:1381/1389` 的 `is_fig_caption` 与 `shape_contract.py:131` 都按**样式名**判定。受影响并实测零对象：`renumber h4-figures` · `style caption` · `renumber-fig` 默认模式 | **未修** |
| M5 | **`renumber-fig` 默认模式在空集上打「✓ 连续 1..N」** —— 空集报绿的 fail-open 自检，被 smoke 以 rc=0 **固化为预期** | `图题数: 0 现号顺序: []` → `验证: captions = []` → `✓ 连续 1..N`，rc=0。对照 `--cn-section`：`图题数: 2` → `✓ 每节连续 1..k`。**smoke 只跑默认模式**（`_verb_specs.py:137`），真正有活干的 flag 从未被敲过 | **未修**（这是铁律 #2「拒绝在空集上报绿」的直接违例，且与 CLAUDE.md caption_re 一节记载的历史 bug 同型） |
| M6 | **「两份副本不会漂」的论证漏了真正会漂的那两个**：wheel 里带着**本仓不版本化的总部文件快照**（`lib/llm_client.py` 是 track 进 git 的绝对 symlink，force-include 会解引用成普通文件打进包；`parallel_contract.py` 同理）。没有任何东西钉住总部 revision，而 struct 层的 sha256 是拿包内内容对**构建那台机器当时的总部工作树** → 总部脏工作树/旧 revision 打出来的 wheel 一样判绿。且同一 SSOT 两处两种拼法（pyproject 写 `../dev/lib/…`，闸门写 `$HOME/Dev/tools/dev/lib/…`）本身就是独立漂移点 —— **正是这两种拼法的差异让 C1 炸在 runner 上** | 包内 `doctools/_bundled/lib/llm_client.py` 外部属性 `0o100644`（普通文件，非 symlink），10791 字节 == 总部 `scripts/tools/llm_client.py` 字节数 | **未修** |

### low

| # | 问题 | 处置 |
|---|---|---|
| L1 | md5 断言测的是「语义未变」而非「未写盘」，且这一性质 **load-bearing 在 `docx_safe_save` collar 上**：注入裸 `Document().save()` round-trip 到只读动词 → 仍绿；注入 `add_paragraph()+save` → 红。smoke 从不在 `DOCX_GRAFT_OFF=1`（既有逃生开关）下跑，那条配置下同一段代码会真重写文件 | **未修**（断言本身站得住，缺的是把这层依赖写明） |

### 施工过程中跑出来、写进表备注但**未擅自修**的三个真问题

1. **`chapter delete --prefix` 从写下那天起不可能生效** —— CLI 声明了它，`cli_forward_probe` 确认忠实转发，但 `chapter.py` 只认 `--h1/--h1-text`，用户必得 rc=2 且错误里根本不提 `--prefix`。**三道旧闸门全绿**。修法应是 `Opt("--prefix", forward="--h1-text", ...)`。
2. **`chrome` 与 `chapter delete` 参数不满足时 rc=1 但 stdout/stderr 全空** —— 静默失败。
3. **`compare-ref ref` 与 `revise-rules gen` 默认 glob 不一致**（`*改动草稿.md` vs `*-改动草稿.md`，`biddiff.py:399` vs `:604`）；不匹配时 `revise-rules gen` 静默变成「待处理 MD: 0 份」且 **rc=0**。

---

## 三、CI 与 wheel 到底能不能用

### 3.1 CI —— **不能用。第一个装依赖步骤必红。**

```
$ docker exec -w /w -e PYTHONPATH=/w/.hq-devtools/lib -e UV_PYTHON=3.12 ci bash -c 'uv sync --python 3.12'
  Resolved 68 packages in 30ms
     Building doctools @ file:///w
  × Failed to build `doctools @ file:///w`
  ├─▶ The build backend returned an error
  ╰─▶ Call to `hatchling.build.build_editable` failed (exit status: 1)
      FileNotFoundError: Forced include not found: /dev/lib/parallel_contract.py
  rc=1
```

**除这一条外，CI 的内容是实的。** 核验方把 `../dev` 造出来之后，yml 里其余每一条 run 逐条跑过，
**全部 rc=0**：

```
✓ 总部 lib 14 模块 · symlink 补实 1 条(git ls-files -s mode 120000)
✓ uv sync + 版本自证 + parallel_contract import
✓ cli_surface ×2 逐字节 diff 一致 · 126 节点
✓ cli_forward_probe 67 条 · check_docx_collar 23+16+0 · check_function_axis 93
✓ check_verbs_reachable 49 · check_smoke_coverage 91 真跑 + 2 skip
✓ script_graph 105 脚本 · 373 引用 · 0 孤儿
✓ pytest scripts/document/tests → 190 passed
✓ pytest tests/smoke → 92 passed, 2 skipped (28s)
✓ 3.13 腿单独跑：282 passed, 2 skipped
```

CI 设计里**做对的部分**（值得留）：

- **两道跑不了的闸门是「大声跳过」不是静默跳过**，且那一步会**自证理由仍然成立** ——
  从 `xref.SCAN_ROOTS` 现取清单（不手抄）逐个打印存在性，任何一个根竟然存在就 `exit 1` 并说
  「跑不了的理由已经过期」。反向验证做过：容器里 `mkdir ~/Work` → 当场判红。
- **`cli_surface` 在 CI 里显式命名为「不是等价闸门」**，只主张三件真的：干净依赖集下 parser 树建得起来
  （**这一条正是抓到 `parallel_contract` 硬依赖的那把刀**）· 树非空 · 指纹确定性（同树连 dump 两次逐字节 diff）。
- `UV_PYTHON` 钉死矩阵 —— 差点交付一个假矩阵：仓根 `.python-version`=3.13，只给 `uv sync` 传
  `--python 3.12` 不够，后面第一条 `uv run` 回头读 `.python-version` 把刚装好的 3.12 venv 删了重建成 3.13。
- 禁 `uv sync --locked` —— 仓里 track 的 `uv.lock` 是 workspace 味的旧锁，含 `../../stations/dockit` 的 editable source。

**同时做错的一件事**（第二节 M1）：wheel 自包含门在 CI 里也跑不了，**但它是静默缺席的**，
yml 一个字没提。这正是 yml 自己开头第 3-5 行写的「哑掉的守卫和它要防的 bug 是同一类东西」。

### 3.2 wheel —— **本机装能用，别的机器装 14/49 动词是死的。**

本机（施工方报，全新 venv + 仓库外实跑）：

```
$ /tmp/doctools-dist-test/bin/doctools --version           → doctools 0.1.0        rc=0
$ cd /tmp && … doctools verbs --fn convert                 → md-to-docx / md       rc=0
$ cd /tmp && … doctools audit headings …/sample.docx       → JSON 全字段           rc=0
$ cd /tmp && … doctools extract …/sample.docx                                      rc=0
包内四个根资源全部命中（styles_registry.yaml / schemas/*.json / _bundled/lib/parallel_contract.py）
```

干净容器（核验方，全新 venv + 全量依赖 + cwd 在仓库外）：

```
$ uv pip install --python /opt/clean/venv/bin/python /tmp/doctools-0.1.0-py3-none-any.whl   rc=0
$ cd /root && /opt/clean/venv/bin/doctools extract /root/fixture.docx
  [docx_cli.py] load error docx_tools.py: ModuleNotFoundError: No module named 'file_ops'
  rc=2
$ 49 个顶层动词逐个 --help  →  === ok=35  modulenotfound=14 ===
```

**为什么本机测不出来**：本机 `doctools extract` 走的是共享 `~/Dev/.venv` 与本机总部路径；
CI 走 `PYTHONPATH` 注入总部 lib —— **那条路径 wheel 用户没有**。三条验证路径都恰好绕开了缺口。

wheel 设计里**做对的部分**：不搬 `scripts/`/`lib/` 一个字节，靠 force-include 镜像进 `_bundled/`
且**相对深度与工作树一致** → 全仓 40+ 处 `parents[2]/"lib"` 的层数算术在包内原样成立，实现代码一处没改；
顺手堵掉了 `pdf_cli.py` 原本的**静默降级**（except 里只留 `--workers`，丢掉 `--batch/--phases/--defer/--fanout-evidence`
→ 同一版脚本在两台机器上 flag 面不一样）；两层门都反向验证过能转红。

---

## 四、闸门终态

**以下 rc 全部是本轮我自己在仓根实跑取得的**（不是转述）：

| 闸门 | rc | 输出 |
|---|---|---|
| `tools/check_smoke_coverage.py` | **0** | 真跑 91 / skip 2 / 共 93（97.8%） |
| `tools/check_function_axis.py` | **0** | 93 条：format 41 · content 15 · review 6 · inspect 28 · convert 2 · dispatch 1 |
| `tools/check_verbs_reachable.py` | **0** | 49 个顶层子命令全部进得去 |
| `tools/check_docx_collar.py` | **0** | 23 + 16 + 0 全部挂收口 |
| `tools/cli_forward_probe.py` | **0** | `"bad": []` |
| `tools/cli_surface.py` | 0 | 改前/改后 diff 空（施工方报，三段施工均对拍过） |
| `pytest tests/smoke` | 0 | 92 passed, 2 skipped, 20.9s |
| `pytest scripts/document/tests` | 0 | 190 passed（dist 段复核）· 原有 224 passed 未掉一条（smoke 段） |
| **`tools/check_wheel_selfcontained.py`** | **0，但这个 0 不算数** | 见 H1：checks 写死两条，`verbs` 不加载实现模块，`--tier full` 与 install 档 checks 完全相同 |
| **GitHub Actions `gates.yml`** | **未通过一次** | 见 C1 |

新增闸门的反向验证（施工方报，三个方向，验完全部还原，还原后 rc=0）：

| 注入 | rc | 输出 |
|---|---|---|
| 删掉 `audit fields` 一行 | 1 | `缺契约：audit fields 在 CLI 里可跑，但 _verb_specs 表里没有` |
| 加一条 CLI 不存在的 `audit unicorns` | 1 | `陈旧条目：表里有 audit unicorns，CLI 里没有这条子命令` |
| 把 `_ROWS` 清空 | 2 | `_verb_specs 动词表为空 —— 空表对上空集不算覆盖。` |

核验方另做的两个注入（均转红后还原、md5 校验、`git status --porcelain` 空）：
只读动词 `para scan-ppr` 偷偷写盘 → 红；写盘动词 `strip bookmarks` 静默空跑（stdout 仍打 `wrote … bookmarks`）→ 红。
**`mutates` 双向判定是这套 smoke 最值钱的一条断言**，它问的是**源件**变没变 ——
`template`/`audit images`/`table extract` 都写盘但写的是 `-o`/`--report`/`--out-dir`，源件必须一字节不动。

---

## 五、诚实的总账：离产品标准还差什么

**不达标。** 按「别人拿到这个仓能不能用」这把尺，差四件事，按修复顺序：

| # | 差什么 | 现状 | 为什么它是产品级门槛而不是打磨 |
|---|---|---|---|
| 1 | **CI 能绿一次** | 一次都没绿过（C1） | 一个从未绿过的 workflow 不是保障，是装饰。PR 页面上没有过一次真的 ✅ |
| 2 | **wheel 的 5 个总部模块进包** | 14/49 顶层动词 `ModuleNotFoundError`（C2） | 「可分发」这三个字目前不成立。核验方已在容器副本里验证补 5 行 force-include 即 `still-broken=0`，改动量极小，但**本机仓未动** |
| 3 | **wheel 自包含门去掉零分辨率的 checks，并进 CI（或显式声明它进不了）** | 门写死跑 `--version` + `verbs`，把两个死掉的动词名字打印出来判绿（H1）；CI 里静默缺席（M1） | 这是本次交付里最危险的一处：**它是唯一一道专门防 C2 的门，而它结构上不可能抓到 C2**。哑掉的守卫和它要防的 bug 是同一类 |
| 4 | **冒烟 fixture 长出 Caption 样式段 + 合规草稿，把 18 条零对象拉回有对象** | fixture 6 条题注全 `Normal`（M4）；`revise-rules gen` 100% 空跑且注释宣称已避免（M2）；`renumber h4-figures` 的 `mutates=False` 由「什么都没干」满足（M3）；`renumber-fig` 空集报绿被固化为预期（M5） | 97.8% 这个数现在**买不到对应强度**。M5 尤其：smoke 把一处 fail-open 自检**写成了预期**，等于给它发了合规证 |

**另有两个不影响「能不能用」、但影响「数字可不可信」的**：M6（wheel 里的总部快照没钉 revision，
struct 层拿构建机当时的工作树对账）· L1（md5 断言的语义依赖 collar，逃生开关下不成立）。

**已经确实拿到手的**（不是估计，是本轮或核验方实跑过的）：

- 93 条动词逐条真敲的契约表 + 双向 fail-closed 覆盖闸门（缺契约/陈旧条目/空表三种注入均转红）
- `mutates` 双向判定 —— 已实证能抓「只读动词偷偷写盘」与「写盘动词静默空跑」两类
- 5 道旧闸门在本机全绿并复核过
- 三个用旧闸门抓不到的真 bug 被抓出来并写进表（`chapter delete --prefix` 从写下那天起不可能生效，
  且 `cli_surface`/`cli_forward_probe`/`check_function_axis` 三道全绿）
- 三处「只在这台机器上跑得起来」被修实（`lib/styles.py` 按 `__file__` 定位仓根 · `pyproject` 加 `docxcompose`
  · `pdf_cli` 静默降级堵掉）
- CI 里两道跑不了的闸门是大声跳过 + 自证理由未过期，反向验证过

---

## 附：交付文件（绝对路径）

| 文件 | 是什么 |
|---|---|
| `/Users/tianli/Dev/tools/doctools/tests/smoke/_verb_specs.py` | 93 行契约表（tuple-of-tuples） |
| `/Users/tianli/Dev/tools/doctools/tests/smoke/test_verbs_smoke.py` | 参数化用例（md5 双向判定在 `:96`） |
| `/Users/tianli/Dev/tools/doctools/tests/smoke/conftest.py` | session 造一次 + function 整树拷贝 |
| `/Users/tianli/Dev/tools/doctools/tests/smoke/_fixture.py` | fixture 造件器（M2/M4 的两处宣称在这里） |
| `/Users/tianli/Dev/tools/doctools/tools/check_smoke_coverage.py` | 覆盖闸门（双向 fail-closed） |
| `/Users/tianli/Dev/tools/doctools/tools/check_wheel_selfcontained.py` | wheel 自包含门（H1 在 `:170` / `:211`） |
| `/Users/tianli/Dev/tools/doctools/.github/workflows/gates.yml` | CI 闸门 workflow，312 行 |
| `/Users/tianli/Dev/tools/doctools/pyproject.toml` | C1 在 `:86`，C2 的缺口在 `:78-86` |
| `/Users/tianli/Dev/tools/doctools/CLAUDE.md` · `README.md` | 新增「冒烟轴」一节 + 闸门清单 |

## 附：本报告自身的可信度边界

- 施工方交给我的三段 JSON 中，**smoke 段与 dist 段均在中途被截断**（smoke 段只收到 39 条实测证据、
  dist 段末尾清理小节断开），核验方第二路的 M6 证据块也在末尾断开。**被截断的部分我没有补写、没有推断**，
  本报告里凡引用施工方数字处均已标注来源。
- 第二、三节里带「本轮复核」标记的事实（`pyproject.toml:86` 字面量 · `check_wheel_selfcontained.py:170`
  的常量 checks · 6 个总部模块的 import 计数 · 五道闸门 rc）是我自己敲命令取得的，不是转述。
- **HEAD 在施工期间被并发会话推进过**（`af2c730` → `f88218b`，其中 `0396e65` 是另一个会话提交的
  `_fixture.py`）。全程 pathspec 钉死提交，**四个 commit 均未 push**。
