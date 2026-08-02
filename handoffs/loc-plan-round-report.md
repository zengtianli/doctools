# doctools 本轮交付报告（四件施工 + 四个核验镜头）

- 仓库：`/Users/tianli/Dev/tools/doctools`
- 基线：`63d8560`（本轮开工前）→ 终态 `0cb5569`，**4 个 commit，均未 push**（`git rev-list --count origin/main..HEAD` = 4）
- 工作树：`git status --porcelain` 空
- 报告落盘：`/Users/tianli/Dev/tools/doctools/handoffs/loc-plan-round-report.md`

---

## 一、四件事各自做成了什么

| # | commit | 做的事 | 净行数 | 落在哪些文件 |
|---|---|---|---|---|
| W1 | `6b63885` | 12 个家族模块各自手写的 `main()` 分发器（19 行逐字相同）下沉为 `_cli_common.family_main(subcommands, argv, *, file, usage_args)` | **+98 / −216 = −118**，13 files | `scripts/document/sub/_cli_common.py`（+60，其中 38 行是 docstring）· `sub/{audit,biddiff,blocks,caption,chapter,freeze,image,renumber,split,strip,table,typeset_ops}.py`（净 −178） |
| W2 | `a5cba6a` | 10 份 `lsof_check` 的四种语义（A/B/C/D 派）收敛为一份 canonical；6 处备份路径命名收编进 `_cli_common`；顺带清死 import | `scripts/` **+109 / −183 = −74**；只算非空非注释行 **−98**。另 `lsof_backup_ab.py` +119/−44（判据换轴，非删减）、`CLAUDE.md` +27 | `sub/_cli_common.py` · `sub/{outline,caption,blocks,styles,slim,freeze,add_header_footer,pipeline_lib,fix_styleset}.py` · `pptx_cli.py` · `handoffs/_loc_plan_harness/lsof_backup_ab.py` · `CLAUDE.md` |
| W3 | `2ac7bfd` | 补上归并遗漏的最后一处手写题注正则：`strip.py` 的 `CAPTION_PATTERN` 改由 `lib/caption_re.py` 的 `STRIP_OUTLINELVL_CAPTION` spec 派生（该 spec **就是** `SECTIONED_CAPTION` 同一对象，不是复制参数） | 测试 44→**55** 条；全仓 213→**224** | `lib/caption_re.py`（:397）· `sub/strip.py`（:1153）· `tests/test_caption_re.py` · `CLAUDE.md`（L197 计数 44→55） |
| W4 | `0cb5569` | `doctools` 装成可安装 CLI（与老绝对路径入口并存）+ 版本号 SSOT 收成一份 | 新增包壳 2 文件 | `doctools/__init__.py`（`__version__` 唯一一份）· `doctools/cli.py`（新增）· `pyproject.toml`（改 `dynamic = ["version"]` + hatch version path）· `scripts/document/docx_cli.py` · `CLAUDE.md` · `README.md` |

**净删合计**：`scripts/` 树 W1+W2 = **−192 行**（W3/W4 是净加，加的是测试与包壳）。

W3 纠正了任务书一处事实：受影响的写盘动词是 **`strip outlinelvl`**，不是 `strip empty-captions`——后者走 `is_caption_style()` 样式名判据，压根不碰 `CAPTION_PATTERN`。

W4 的关键手法：`--version` 在 `main()` 里手动拦第 0 位 argv，**不注册成 argparse 参数**，因为 `cli_surface` 是逐 action 比对整棵树的。反向验证做过——临时改成 `add_argument("--version", action="version")`，闸门当场转红并点出多出的 `_VersionAction` 节点。所以 surface 指纹 `md5 = 3bbad2cd98a98acd7c7599956378c151` 改前改后同值，不是"没测出来"，是那条路没走。

---

## 二、核验发现的问题（按 severity 排）

**六条全部没修。** 下面逐条写明没修的原因，不包装成"已知限制"。

### M1 · 4 个顶层组 / 11 条动词是死的（本轮之前就存在）

`scripts/document/docx_cli.py:466` 的 `DISTILLED_GROUPS` 集合里没有 `fix` / `seqdiff` / `compare-ref` / `revise-rules`，但这 4 个组注册进了 subparser、`--help` 里列着、`verbs` 表里也列着。`main()` 因此走"未知顶层 token → 塞进 top_argv"分支，而 `top_p` 是 `add_help=False + parse_known_args` 不报错，随后 `sub_cmd is None → print_help(); return 0`。

后果：这 4 组下的 **11 条动词**（fix 7 + seqdiff 2 + compare-ref 1 + revise-rules 1，占 93 条的 12%）敲下去只打印根 help、rc=0、什么都不做、也不报错。

我本人复跑确认（不是转述）：

```
$ python3 scripts/document/docx_cli.py fix          → rc=0
$ python3 scripts/document/docx_cli.py seqdiff      → rc=0
$ python3 scripts/document/docx_cli.py compare-ref  → rc=0
$ python3 scripts/document/docx_cli.py revise-rules → rc=0
（其余 43 个顶层组 rc=1/2 或正常执行）
```

核验镜头还跑了：`docx_cli.py fix clear-direct-format <docx> --inplace` → 整页根 help、rc=0、目录里没备份没写盘没报错；在 `63d8560` 旧树复跑同样 4 组全死 → **确认非本轮引入**；`git log -S '"fix",' -- scripts/document/docx_cli.py` 无命中 → 从没进过 `DISTILLED_GROUPS`。

**和本轮的关系**：W2 把 `fix_styleset.py` 的 lsof 判据收编进 `_cli_common`，而 fix 的 7 条动词全在这个死组里——那个调用点从文档化的 CLI 根本走不到，本轮对它的任何"实测"都不可能覆盖。更要紧的是本轮依赖的两道闸门此刻 rc 全 0：`cli_forward_probe`（67 条，`grep -c clear-direct-format probe.out` = 0）与 `check_function_axis`（93 条，表里明确列着 `fix clear-direct-format`）。**闸门绿着、11 条动词是死的**，这正是假绿的形状。

**没修的原因**：修它要动 `docx_cli.main()` 的顶层分发，会改 CLI 运行时行为（11 条动词从"打印 help 退 0"变成"真跑"），超出本轮四件施工的授权范围，且需要先确认这 4 组的实现是否还能跑通。属于必须单独立项的破坏性/行为性变更。

### M2 · `blocks reorder` 这条写盘路径没有占用闸门（本轮之前就存在）

`scripts/document/sub/blocks.py` 里 `lsof_check(` 的调用只有一处，在 **1572 行**，位于 `main_relocate()`（1549 起）内；`main_reorder()`（1221 起）完全没有。我本人 grep 复核确认。

核验实测：起 holder 进程 open 住 `t.docx`（先断言 `lsof -- <docx>` rc=0、2 行输出），跑 `blocks reorder` → old/new 都 rc=0、都不出占用字样，stdout 是 `[backup] … / [plan] blocks=1 / [apply] moves=0 → saved`，即**文件被 Word/WPS 打开时照样备份 + 覆写**。同一 harness 下有闸门的 9 处（styles ×3、slim、add_header_footer ×3、outline ×3、strip ×3、freeze ×3、pipeline_lib、blocks relocate ×3、caption number ×2）old/new 退出码逐条相同。

**没修的原因**：W2 收敛的是判据（十份实现变一份），不是覆盖面（哪些写盘路径该挂闸门）。给 `reorder` 加闸门是新增行为，不在授权内。

### M3 · `check_docx_collar` 第三判据在 `lib/` 下方向反了

`tools/check_docx_collar.py:26` 的 `SCAN = ROOT / "scripts"`（判据 1 扫描根不含 `lib/`），而判据 3 在 `:173` 显式 `continue  # 已由第一判据管着，不重复点名`。于是 `lib/` 下「import python-docx + `.save()` + 无收口」这一格：判据 1 扫不到、判据 2 只管 ZipFile、判据 3 主动让开——三条判据同时看不见。

注入证据（核验镜头跑的，注入件已删、`git status` 空）：

| 注入 | collar rc |
|---|---|
| `lib/_zz_libsaver.py`：**带** `from docx import Document` + `doc.save()`，无收口 | **0（绿）** |
| 同一文件，唯一差别是删掉那行 docx import | **1（红，点名该文件）** |
| 对照：同样形状放 `scripts/document/sub/_zz_saver.py` | **1（红）** |

即：一个 `lib/` 下的存盘 helper **加上** docx import 反而从红变绿。`offenders()` 名册里 `lib/` 文件数 = 0。

这正是 W2 这一轮的形状（抽公共函数进共享层），而 `lib/` 是本仓 CLAUDE.md 钦定的公共模块层。**今天 `lib/` 里没有 `.save(`，所以是潜伏洞不是在跑的 bug。**

**没修的原因**：任务是四件施工 + 核验，改闸门本身不在授权内（核验镜头明确"没有改任何闸门本身"，以免把被测对象和测量工具一起动）。

### L1 · `check_docx_collar` 的 stem 碰撞让 3 个文件整体隐身

`tools/check_docx_collar.py:126` 用 `out.setdefault(p.stem, p)` 建 stem→路径表，同名只留先扫到的（`scripts/` 在 `lib/` 前）。实测碰撞 3 处：

```
__init__     : keep=scripts/document/sub/__init__.py   DROPPED=lib/__init__.py
docx_revise  : keep=scripts/document/docx_revise.py    DROPPED=lib/docx_revise.py
styles       : keep=scripts/document/sub/styles.py     DROPPED=lib/styles.py
```

`lib/styles.py`、`lib/docx_revise.py` 这两个 docx 相关模块因此整个从判据 3 视野消失。注入：往 `lib/styles.py` 追加 `def _zz_save(doc, dst): doc.save(str(dst))` → collar **rc=0（绿）**。已 `git checkout` 还原。

**没修的原因**：同 M3，改闸门本身不在授权内。

### L2 · W4 新加的 `doctools/` 包在所有闸门扫描根之外

- `tools/check_docx_collar.py:26` `SCAN = ROOT / "scripts"`；`:215` `scan_roots = [SCAN, ROOT / "lib"] + extra`
- `tools/script_graph.py:44` `SCAN = [ROOT/"scripts", ROOT/"lib", ROOT/"tools"]`
- `check_external_refs.py` 的 `SCAN_ROOTS` 同样没有它

注入 `doctools/_zz_saver.py`（裸 `Document()` + `.save()`）→ collar / script_graph / xref / axis / pytest **五道门全 rc=0**。已删还原。

今天 `doctools/` 只有 `__init__.py` + `cli.py` 两个壳、docstring 写明"业务逻辑一行都不在这里"，**但那条规矩现在没有任何机器层兜着**（铁律 #13：只写文档不算约束）。

**没修的原因**：需要改三个闸门的扫描根，属于改闸门本身。

### L3 · commit `a5cba6a` 的 message 里引用了一个错的数字

message 写「以 `-` 开头的相对文件名被裸 lsof 解析成 `-h`(帮助) → rc=0 且 stdout 空」。实测 rc 是 **1** 不是 0，stderr 打的是 `lsof: illegal option character: .` 而不是帮助：

```
['lsof', '-t.docx']       rc=1  stdout 0 行  stderr='lsof: illegal option character: .\nlsof 4.91 …'
['lsof', '--', '-t.docx'] rc=0  stdout 2 行  （COMMAND/PID 表头 + 持有行）
```

结论方向没错（老 A/C/D 三派确实一致误判为空闲，`--` 确实拦得住），`_cli_common.lsof_check` 的 docstring 措辞本身是准确的（只说"会被解析成选项 → 报 usage 或误判"），**错的只有 commit message**。

**没修的原因**：改它要 `git commit --amend`，会重写已存在的 commit hash，本轮四个 commit 之间有引用关系（报告与 handoff 里已引 `a5cba6a`）。留作事实记录在此。

---

## 三、行为变更清单（实测前后对照）

### 已批准的两处

**① lsof 判据收敛（W2）** — 判据矩阵，真开 fd 持有 + 假 lsof 构造两路，全部实跑过：

| 场景 | 新 canonical | 旧 A | 旧 C(bool) | 旧 D(guard) |
|---|---|---|---|---|
| 真 fd 持有（绝对路径） | OCCUPIED | OCCUPIED | OCCUPIED | OCCUPIED |
| 存在但无人持有 | free | free | free | free |
| 路径不存在 | free | free | free | free |
| **真 fd 持有 + 前导 `-` 的相对文件名** | **OCCUPIED** | free | free | free |
| 假 lsof：rc=1 + 2 行 stdout | free | free | **OCCUPIED** | — |
| 假 lsof：rc=0 + 仅表头 1 行 | free | **OCCUPIED** | **OCCUPIED** | — |

- **唯一的真行为变更是第 4 行**，方向从错到对：老三派会对着 Word 正开着的文件放行写盘。rc 检查救不了这条，只有 `--` 能。
- 第 5/6 行只有假 lsof 造得出（真 lsof 单路径查询找不到时是 rc=1 + stdout 空），属加固不是回归。
- 各调用点的文案与退出码逐条不变，端到端实测：`caption number` 占用时 rc=2 + `[ABORT] … 被 Word/WPS 打开, 请先关闭。`；`outline normalize-arabic` rc=3；`pptx_cli titlecolor` rc=1；空闲时全 rc=0。
- A 派返回值形状变了（raw stdout 带尾换行 → strip 后 join），实测 `old.strip() == new` 为 True，只影响错误信息末尾那个空行。
- 备份命名：7 条写盘动词各跑两遍逼出 N 自增，old/new 产出文件名集合**全 SAME**（`t.bak-1-2026-08-02.docx` / `t.bak-2-…`）。

**② `strip outlinelvl` 题注判据扩大（W3）**：

| 语料 | before 命中 | after 命中 |
|---|---|---|
| W3 fixture `dash-captions.docx`（8 段全带 `w:outlineLvl`） | `captions_processed` = **2** | **7** |
| 核验镜头自建 22 段对抗语料 | **7** | **12** |

新抓到的 5 条（原本漏网、`w:outlineLvl` 留在文档里继续污染 Word 导航窗格）：

| 文本 | 漏网原因 |
|---|---|
| `图3.1‑2` | U+2011 NON-BREAKING HYPHEN |
| `表3.1–3` | U+2013 EN DASH |
| `图3.1—4` | U+2014 EM DASH |
| `表3.1－5` | U+FF0D FULLWIDTH HYPHEN |
| `图3．1-6` | 章号内 U+FF0E 全角句点 |

范围**只**扩在短横/句点字符类。负例两边一致不中、零丢失：`3.1 这是真章节标题`（对照真章节）、`表3-1` 扁平编号、`图书馆3.1-1`、`表面粗糙度`、`附图`、`见图…` 内联引用、`Figure`、中文数字章号 —— 核验镜头 10 条对抗负例 old/new 完全一致。

两个入口（`sub/strip.py` 直敲 与 `docx_cli.py strip outlinelvl`）结果一致。全程 `--dry-run`，未对任何真实文件跑写模式。

### 其余三处：接口新增/形状变化，非语义变更

| 变更 | 实测 |
|---|---|
| W1 家族 `main()` 下沉 | `family_ab` 144 例（12 族 × 12 CASE）rc/stdout/stderr **byte-identical**；核验镜头独立重写的 180 例 harness（加了每子命令 `--help` 与裸敲）只有 1 处 diff = `strip outlinelvl --help` 的 docstring 散文，逐行确认只动了两段说明文字（W3 有意改的），**非 W1 引入** |
| W1 `exec_script` 真调用路径 | `_dispatch.exec_script(f, ['bogus'])` 对 12 个族全跑，12/12 打印 `[<族名>] unknown subcommand: 'bogus'; choices=[...]` 且 rc 全 **2**（裸 import 会在这里炸，走的是 `spec_from_file_location` 真路径） |
| W4 新增 `doctools` 命令 | `.venv/bin/doctools --version` → `doctools 0.1.0`；`doctools verbs \| tail -1` → `共 93 条`；`doctools verbs --fn convert` → 2 条。真活对拍：`audit headings` 跑 `templates/template.docx`，装出来的命令与老绝对路径 stdout **逐字节相同**（`cmp`=0，1693 字节，rc 都 0）。老入口照旧输出 93 条。**surface 指纹未变** |
| W4 版本号 SSOT | 改成 `9.9.9` → 包元数据 / `doctools --version` / 老路径 `--version` **三处一起变**；改成非 PEP440 的 `9.9.9-probe` → hatchling 构建 `ValueError: Invalid version … from source 'regex'`。`grep -cE '^version = ' pyproject.toml` 现在是 0 |

---

## 四、闸门终态

**本报告写作时由我本人在终态 `0cb5569` 上重跑，不是转述 worker 报告**：

| 闸门 | rc |
|---|---|
| `tools/cli_surface.py` | **0** |
| `tools/cli_forward_probe.py` | **0**（67 条，异常 0） |
| `tools/check_docx_collar.py` | **0**（23 / 16 / 0 三判据） |
| `tools/check_function_axis.py` | **0**（93 条） |
| `tools/check_external_refs.py` | **0**（219 处） |
| `python3 -m pytest -q` | **0**，**224 passed, 1 warning** |
| `git status --porcelain` | 空 |

另由施工方/核验方跑过的专用对拍器：`family_ab.py` 144 例 diff 空 · `lsof_backup_ab.py` rc=0（换轴后 4 条判据）· 核验方自写 `fam_ab_mine.py` 180 例（1 处已解释的 docstring diff）。

**注意**：这套 rc 全绿与 §二 M1 的 11 条死动词并存——闸门测的是 argv 转发与职能标签，测不到"顶层分发器认不认这个组"。**"闸门全绿"在本轮不等于"CLI 全活"**，这是本报告最重要的一条限定。

---

## 五、诚实的总账

### 还剩什么没做完

| 项 | 状态 | 为什么 |
|---|---|---|
| M1 · 4 死组 / 11 条动词 | **未修** | 修它改 CLI 运行时行为，超本轮授权；需单独立项 + 先确认这 4 组实现还能跑 |
| M2 · `blocks reorder` 无占用闸门 | **未修** | 是覆盖面问题不是判据问题，W2 授权范围只到判据收敛 |
| M3 · collar 判据 3 在 `lib/` 下方向反 | **未修** | 改闸门本身不在授权内；今天 `lib/` 无 `.save(`，是潜伏洞 |
| L1 · collar stem 碰撞吞掉 3 文件 | **未修** | 同上 |
| L2 · `doctools/` 在所有扫描根外 | **未修** | 需改三个闸门扫描根 |
| L3 · `a5cba6a` message 数字错 | **未修** | 要 amend 重写 hash，且报告已引该 hash |
| 4 个 commit push | **未做** | 按要求不 push |
| `CLAUDE.md` 补 `family_main` 的记述 | **未做** | W1 任务未要求且不在其 pathspec；`sub/ (44)` 文件数未变（没增删文件），只是记述可补一句 |
| `uv sync --all-packages` 会清 venv 漂移包 | **未处理，需你拍板** | `bs4 / playwright / tqdm / pytz / decorator / docx2pdf / appscript / pycryptodome / retry / volcengine` 从来不在 lock 里，是手工 pip 装进共享 venv 的漂移，**任何一次 sync 都会清**（`git diff uv.lock` 只有 3 行 `virtual→editable`，包集合没动，不是本轮引起）。真有消费的只有两处、都是函数内懒 import：`tools/dev/repo_manager.py` 的 playwright、`tools/dev/lib/tools/cloud/volc_usage.py` 的 volcengine。修法是声明进对应 member 的 `pyproject.toml`，属 `tools/dev` 范围 |
| 改版本号后 `uv sync` 不重建 dist-info | **已知并绕开** | `_pkg_version()` 先读源文件、元数据只作兜底（有意为之，让 `--version` 在元数据发霉时仍说真话）。要让元数据跟上须 `uv pip install -e tools/doctools --no-deps --reinstall`，已写进 CLAUDE.md 与 README |

### 哪些断言有命令支撑，哪些只是"看起来对"

**我本人在写报告时跑过的**（第一手）：五道闸门 rc + pytest 224 + `git log/status/rev-list` + 4 个死组逐个裸敲 rc=0 + `grep -n 'lsof_check(' blocks.py` 只有 1572 行且在 `main_relocate` 内 + collar 的 `SCAN`/`setdefault`/`scan_roots` 三处行号。

**两个核验镜头独立跑过、且做了反向注入验证的**（第二手但有证伪）：

| 断言 | 反向注入 |
|---|---|
| 家族折叠 rc/stdout/stderr 不变 | 把 old 树 `strip.py` 的 usage 尾巴改反 → diffs 1→4，harness 真能抓 |
| W3 回归门真会咬 | worktree 里把 `CAPTION_PATTERN` 换回旧字面量 → 224 变 2 failed，CLI 命中从 12 掉回 7 |
| collar 在 `lib/` 下有洞 | 同一文件加/删一行 docx import → rc 在 0 和 1 之间翻转 |
| collar stem 碰撞 | 往 `lib/styles.py` 加裸 `.save()` → 仍 rc=0 |
| `doctools/` 在扫描根外 | 放裸存盘件 → 五道门全绿 |
| W4 `--version` 不进指纹 | 改成 `action="version"` → surface 当场转红 |
| W1 usage 差异真被 family_ab 看见 | 注入后 family_ab 3 例转红，而 surface/probe/axis/pytest **四道门全绿看不见** |
| W3 spec 共用而非复制 | 塞回旧正则 → 两条结构性测试转红 |

核验镜头还自带 UNREACHED 判据（空闲场景 rc 必须为 0 否则不计入结论），正是它揪出了第一版里 `fix_styleset`（漏 `--inplace`）、`blocks`（用错子命令）、`pipeline`（步骤名不存在）三条假绿并重跑。

**只有施工方单方报告、未被独立复跑的**（信但未证伪）：W2 的假 lsof 构造场景（第 5/6 行矩阵）· W4 的 hatchling `ValueError` 探针与 `9.9.9` 三处联动 · W1 的一次性替换脚本 `expected(fam)` 逐字重建 + `count()==1` 断言 12/12。这几条各自都有具体命令输出贴在原返回里，但没有第二人复跑。

**核验镜头复跑过的交叉项**：`grep -rnE '^[[:space:]]*(import styles|from styles)' scripts lib tools doctools` **零命中** —— W1 扩大 `sys.path` 带来的 `lib/styles.py` vs `sub/styles.py` 同名风险当前未触发。这条 W1 声称过、核验独立复核过，一致。

**没有任何断言是纯读代码推理得出的。** 唯一带推理成分的是 M3 的"今天是潜伏洞不是在跑的 bug"——依据是 `lib/` 下当前无 `.save(` 的 grep 结果，属事实而非推测。
