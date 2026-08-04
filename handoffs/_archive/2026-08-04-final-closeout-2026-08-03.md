# doctools 收尾轮 · 2026-08-03

三个真 bug + 冒烟 fixture 空跑面 + hq-devlib 可分发。8 个 agent 施工 + 3 个独立核验镜头，
主会话逐条复核后又自己揪出并修掉 5 条镜头没覆盖的。

## 一、施工项与实测证据

| # | 问题 | 修法 | 反向验证 |
|---|---|---|---|
| 1 | `image-caption --help` **去读 Finder 选中项**，rc=1（选中项是 docx 就会被写盘） | 闸门下沉到模块顶部、`import docx` **之前**（`_exec_script` 是 `exec_module` 后才调 `main()`，只拦 main 挡不住重家伙加载）；未知 flag rc=2 禁 fallthrough | 关掉闸门 → 又打 `/Users/tianli/Work/...` 路径且 rc=1 |
| 2 | `text-fmt --help` 被手搓 parser 当未知参数，rc=2 | 补正经 usage，未知 flag 仍 rc=2 | 去掉 help 分支 → 回到 rc=2 |
| 3 | `track compare` 是「compare 功能将在 v2 实现。」**rc=0** 的空桩 | 段落级 diff → `w:ins`/`w:del` 注入（删除态走 `w:delText`） | 主会话实测：`w:ins`×6 `w:del`×4 `w:delText`×2，`w:del` 内裸 `w:t` = 0；两份相同的文档 **rc=1** 明说无差异 |
| 4 | `renumber-fig` 空集上打「✓ 每节连续」 | `EMPTY_RC=3` + `_empty_set_exit()`，9 个调用点；判据下沉 `lib/caption_re` | 见下 §二.4 |
| 5 | fixture 零对象（题注全 `Normal`、无图、`revise-rules gen` 空跑） | 真题注样式 + 6 张现造内嵌图 + 3 种短横字符 | `_verb_specs` 的 `expect_rc`/`mutates` 按新 fixture 逐条重测 |
| 6 | `hq-devlib` 没发 index → 只拿 doctools wheel 在**依赖解析**就死 | PEP 508 直接引用 `git+https://github.com/zengtianli/devtools.git`（仓根即 hq-devlib，**无** `#subdirectory`）+ `allow-direct-references` | 本机零变化：workspace 根 `[tool.uv.sources]` 覆盖，`uv.lock` 仍 `virtual = "tools/dev"` |

## 二、主会话复核时又抓到的（镜头没覆盖）

### 1. 判据 C 此前**只存在于文档里** — 已补回实现

`pyproject.toml` 与 `CLAUDE.md` 都指着它说「包内副本 == `git ls-files`，逐文件 sha256，
两个方向都查」，而 `grep -n 'sha256\|ls-files' tools/check_wheel_selfcontained.py` **零命中** ——
08-02 重写这道门时删掉了。**被文档背书却不存在的门比没有门更坏**：读的人不会再去查。

### 2. 判据 C 补回后第一次运行就抓到跨机构建阻断

`lib/llm_client.py` 是 git 里 **mode 120000 的 symlink**，指向仓外绝对路径。
整目录 force-include 让**构建本身**依赖那条路径在本机解析得开：

```
$ ln -sf /nonexistent/... <proj>/lib/llm_client.py && uv build --wheel
FileNotFoundError: [Errno 2] No such file or directory: '<proj>/lib/llm_client.py'
error: Call to `hatchling.build.build_wheel` failed (exit status: 1)   ← rc=2
```

**即这个 wheel 在这台机器之外根本构建不出来，而判据 A 一直在打 ✓** —— 与 08-02 那次
`$HOME` 假绿同形。`exclude` 救不了：force-include 不受它约束（实测加了照样在包里）。

修法：`lib` 改**逐文件** force-include（16 行）排除该 symlink，模块改由 `hq-devlib` 出
（总部包 `sources` 加 `scripts/tools` 前缀，`llm_client.py` 装进 site-packages 顶层）。
**两条路都不动任何一行 import**。判据 C 对 symlink 用**反向**断言：进了包就判红。

- 反向验证：`lib` 改回整目录 → 判据 C 红并点名 `仓外 symlink 混入: lib/llm_client.py`
- 修后：悬空 symlink 下构建 **rc=0**（修前 rc=2）

### 3. `script_graph.py` 的动词轴是 fail-open

`verb_map()` 载入失败只 print 一行 ⚠、**rc 照样 0**。实测注入一句 ImportError 后仍输出
「106 脚本 · 374 引用 · 0 孤儿」rc=0，而三视图里整整一视图已空。它在必跑闸门清单里，
rc=0 被当过「动词映射对得上」的证据。改 `exit 2`；反向验证 注入 rc=2 / 正常 rc=0。

### 4. `EMPTY_RC=3` 与加强 fixture 同一轮落地 → 新判据当场被自己那轮的 fixture 绕过

谁把 `_empty_set_exit` 改回 `return`，92 条 smoke 一条都不会红。补
`scripts/document/tests/test_renum_empty_set.py`（7 条，覆盖 `figures`/`tabfig`/`chapter`）。
第 7 条是**非空对照** —— 没有它，把 `_empty_set_exit` 挪到函数开头无条件调用也能让前 6 条全绿。
反向验证：注入 `return` → 6 红，恢复后 7 绿，文件逐字节还原。

> 写这条测试时自己踩了一次：`figures` 的**默认 kind 是英文 `Figure`**，拿中文题注做对照组
> 会得到 rc=3 —— 判据没错，是对照组构造错了。注释已留在测试里。

### 5. `check_external_refs` rc=1

断链在 `~/Work/projects/reclaim/handoffs/sessions-recap.md:65`（今天 01:38 另一路会话写入，
引用 07-31 已折进 `docx_fmt.py` 的旧脚本名）。已改指真实路径并在 reclaim 仓提交。

## 三、终态

```
cli_surface 0 · cli_forward_probe 0 · check_function_axis 0 · check_verbs_reachable 0
check_smoke_coverage 0 · check_docx_collar 0 · check_external_refs 0 · script_graph 0
check_wheel_selfcontained 0  ← A/B/C 三判据全绿，49/49
pytest: 336 passed, 2 skipped
107 脚本 · 378 引用 · 93 动词 · 0 孤儿
```

三仓已推：`devtools b81bc49` · `doctools 7a20879` · `reclaim 948038f`

## 四、唯一未闭合项（属用户专属操作）——已闭合（2026-08-04）

用户已建细粒度 PAT 并 `gh secret set HQ_DEVTOOLS_TOKEN`。第一次 rerun 仍红：checkout 对
`zengtianli/devtools` 返 **404**（不是 401）—— token 合法但看不见私库，根因是新建页
Repository access 停在默认的 **Public repositories**。改成 Only select repositories 圈进
devtools（原地改，token 值不变，secret 不用重配）后 rerun：
**run 30804868071 两个矩阵 job（py3.12/3.13）全绿 + artifacts 上传成功 —— 该 workflow 首次全绿。**

以下为闭合前的记录，留作凭证：

**CI 曾红**，卡在 `Input required and not supplied: token` —— `gh secret list -R zengtianli/doctools`
为空，`HQ_DEVTOOLS_TOKEN` 没配。`zengtianli/devtools` 是 **PRIVATE**，两处需要它：
① `actions/checkout` 取总部件 ② `uv sync` clone git URL 依赖（workflow 里已用同一个 secret
配 `url.insteadOf`，不需要第二个 secret，也不需要 deploy key）。

要建的是**细粒度 PAT**：仅 `zengtianli/devtools`，权限 `Contents: Read`。

```
https://github.com/settings/personal-access-tokens/new
gh secret set HQ_DEVTOOLS_TOKEN -R zengtianli/doctools
```

⚠ **不要图省事用 `gh auth token` 那把** —— 它是 classic PAT，带 `repo` / `delete_repo` /
`workflow` 全量写权限，而 doctools 是 **public 仓**；塞进它的 Actions secret，
泄露面是整个 GitHub 账号。

其余三条路都不可行：devtools 转 public（仓里有 `.claude/settings.local.json`、日志、
文件名带真人姓名的 CSV）· 把 6+1 个模块 vendor 进 doctools（造第二份 SSOT，正是选方案 b
要避免的）· 发 PyPI（用户已否）。
