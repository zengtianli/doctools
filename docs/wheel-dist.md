# wheel 分发全记录（从 CLAUDE.md 外迁 · 2026-08-04）

> 常驻结论与硬约束在 `CLAUDE.md` §「装出去的 wheel 必须自包含」；本页是完整叙事与实证，
> 改 `pyproject.toml` 打包段 / `tools/check_wheel_selfcontained.py` 前先读这里。

## 形状

`scripts/` `lib/` **不在 `src/doctools/` 里**，而 `~/Work` 有 130 处绝对路径钉着它们的
现有位置，所以不能搬。改用 pyproject 的 `force-include` 在**构建时**把
`scripts/ lib/ config/ schemas/` 四个根镜像进 `doctools/_bundled/`，
**工作树一个字节都不动**。`src/doctools/cli.py` 按顺序试两个实现根：
工作树（editable 装）→ `_bundled/`（wheel 装）。

镜像时**相对深度与工作树完全一致**，所以全仓 40+ 处 `parents[2]/"lib"` /
`parents[3]/"lib"` 的层数算术在包内原样成立，一处都不用改。

**wheel 有条件可分发**（2026-08-03 实测 **49/49 全起得来**；同日中途 48/49、08-02 是 47/49、08-02 上午还是 0/49）。

## 判据 A 曾是假绿（2026-08-03 抓到）

⚠ **判据 A 在 08-03 之前是假绿**：它验的是「在没有兄弟 `dev/` 的位置构建成功」，
但 `lib/llm_client.py` 那条指向仓外绝对路径的 symlink 让构建暗中依赖本机路径，
在别的机器上 `hatchling.build.build_wheel` 直接 `FileNotFoundError` 退出。
详见下面判据 C 那节 —— 那个洞是判据 C 补回实现后**第一次运行**抓到的。

⚠ 这个 49 是**按工作树快照**测的，不是按 HEAD。本门 `git archive` 的 ref 写死是 `HEAD`，
所以工作树有未 commit 的打包相关改动时直接跑它必红（rc=2）。
**那是时序不是回归。** 攒证据的正确姿势是把工作树做成悬空提交再让本门按那个 ref 取源：
`GIT_INDEX_FILE=<临时> git read-tree HEAD && git add -A && git write-tree` → `git commit-tree`，
全程不碰分支 / 真 index / 工作树。

## hq-devlib：git 直接引用（2026-08-03 改，用户拍板不发 PyPI）

⚠ **「有条件」三个字不能省，但条件已经从「摆两个 wheel」换成「有私库读权限」**。
原来 `hq-devlib>=0.1` 是普通 specifier，而它没发到任何公开
index，于是**只拿 doctools 的 wheel 会直接解析失败**（`Because hq-devlib was not
found in the package registry …`）—— 不是「装上后几条动词不能用」，是 pip 一步都
进不去；CI 的 `uv sync` 死在同一处。现在改成 **PEP 508 git 直接引用**：

```toml
"hq-devlib @ git+https://github.com/zengtianli/devtools.git",
```

三件必须一起记住的事：

| | |
|---|---|
| **不带 `#subdirectory=`** | `zengtianli/devtools` 这个仓的**根就是** `~/Dev/tools/dev`，pyproject 在仓根。写 `#subdirectory=tools/dev` 会 clone 完当场 `has no subdirectory tools/dev`（实测） |
| **`[tool.hatch.metadata] allow-direct-references = true` 必须跟着加** | 否则 hatchling 把直接引用判成硬错误，连本机 `uv sync` 都过不去（`cannot be a direct reference unless …`，实测）。它默认关着是因为**带直接引用的包 PyPI 拒收** —— 也就是说这条路和「发 index」互斥，哪天要发得两处一起改回去 |
| **`zengtianli/devtools` 是 PRIVATE** | 所以这条 URL 只在**有该仓读权限**的环境里成立：本机（git 常规凭证）✓ · CI（secret `HQ_DEVTOOLS_TOKEN` 配 `url.insteadOf`，见 `.github/workflows/gates.yml`；**细粒度 PAT 仅 devtools · Contents: Read**，2026-08-04 配好后 gates 首次全绿；踩坑 = 新建页 Repository access 默认 Public repositories，对私库返 404 不是 401）✓ · 没凭证的第三方 ✗。要让任意第三方装得动，只有把 devtools 转 public 这一条路 |

**本机行为一个字节没变**：workspace 根 `~/Dev/pyproject.toml` 的
`[tool.uv.sources] hq-devlib = { workspace = true }` 覆盖掉这条 URL。实测
`cd ~/Dev && uv sync --all-packages` rc=0、`~/Dev/uv.lock` 改前改后**逐字节相同**
（仍是 `source = { virtual = "tools/dev" }`），且 `uv sync --all-packages --offline`
照样 rc=0 —— 根本不走网络。

本仓运行时依赖总部 `~/Dev/tools/dev/lib/` 的 6 个**平铺模块**（finder / file_ops /
display / parallel_contract / usage_log / env，共 1156 行）。它们原来不是包、只能靠
sys.path 注入导入，所以「声明成 dependency」这条路当时走不通；现在总部仓把这 6 个打成了
分发名 **`hq-devlib`**（`lib/*.py` 零改动、44 处 sys.path 注入原样继续工作）。
`hq-devlib` 在 workspace 里是 virtual，`uv sync` 不会把它装进共享的 `~/Dev/.venv`
（装了也只是多一条排在 sys.path 注入之后的来源）。

## `text-fmt` / `image-caption` 的 `--help`（2026-08-03 修，凑齐 49/49）

`text-fmt` 的 `--help` 原来 rc=2（打自己那份「未知参数」清单）——
`docx_fmt.py::text_main()` 开头拦 `-h/--help` 打 `TEXT_USAGE` 并 `return 0`。拦截**必须在
`--scope` 取值与未知 flag 判定之前**，否则 `--scope --help` 会被 `next(_it)` 先吃掉；
`-h` 也一并认了 —— 旧判据只看 `startswith("--")`，`-h` 会被当文件名吞进 `get_input_files()`。

`image-caption` 原 rc=1，根因不是分发：它**没有 argparse**，main() 把 argv 直接丢给
`get_input_files()`，而后者见「一个位置参数都没有」就**回落去读 Finder 当前选中项**并按
写模式处理 —— 于是 `--help` 会去动用户此刻选中的文件。这正是「破坏性动作必须自己占一个
动词」一节判过死刑的那条（「`--help` 弹 Finder + 往选中文件写盘，曾写进 ~/Work 在跑的
项目」，判据 = **未知 flag 一律 `sys.exit(2)`，禁 fallthrough**），当时的漏网点。

修法与**位置**都要照抄：闸门在 `sub/docx_apply_image_caption.py` **模块顶部、
`from docx import ...` 之前**，不是在 `main()` 里。因为 docx_cli 的 `_exec_script` 走
`spec_from_file_location` + `exec_module` 再调 `main()`，只在 main() 里拦，重家伙在
模块加载阶段就已经 import 完了。顶层那道用 `_invoked_as_cli()` 把自己关掉，否则
`typeset_apply` 的 `load_step()`（同样是 spec 载入，只为拿 `apply()`，此刻 `sys.argv`
是 typeset 自己的、带 `--dry-run`）一进这步就会 exit(2)。main() 里还留了同一个纯函数
作第二道，兜 `sys.modules` 已缓存、顶层不再执行的第二次转发。

⚠ 连带行为变更：`image-caption <docx> --dry-run` 从「**静默忽略该 flag、照常写盘**」
变成 rc=2。本脚本从来没实现过 dry-run，旧行为是假装接受。

## 验法三戒

⚠ 在本机验「别的机器能不能用」时**必须中和 `$HOME`**：`docx_cli.py` 有一句
`Path.home()/"Dev"/"tools"/"dev"/"lib"` 兜底导入，不中和的话 0/49 的 wheel 也能
测成 49/49 —— 2026-08-02 上午就是这么报出一次假绿的。这门自己中和（`HOME` 指空
目录 + 清 `PYTHONPATH`），别绕过它手验。

⚠ 判据 B **带依赖装**（`pip install <whl>`，不是 `--no-deps`）。`--no-deps` 装法下
这门永远只能测出 0，而且卡点会随修复一路搬家（parallel_contract → yaml → lxml），
把「依赖声明对不对」这个真问题挡在门外。

⚠ 门里原来还有一段「把兄弟目录 `../dev` 构建成 wheel 丢进 `--find-links` 目录当
index 替身」，**2026-08-03 已删**：hq-devlib 改成 git 直接引用之后，直接引用的优先级
高于任何 index / find-links，那个目录**永远不会被看一眼**（实测：给了 find-links，
装出来的仍是 `+ hq-devlib==0.1.0 (from git+…)`）。
留着它只会让这门看起来在测「index 上取得到」，而它实际测的是「git URL clone 得动」。
代价是本门现在**要能上网 + 要有 devtools 私库读权限**，两者缺一即 `SystemExit(2)`
（fail-closed，不退化成假绿；反向验证：把 URL 换成不存在的仓，本门实测 rc=2 并打出
`Repository not found`）。装包那一步走真实环境，跑动词时 `HOME`/`PYTHONPATH` 照样中立。
（`--tier struct` / `--tier full` 是 2026-08-02 写进文档却从未存在过的 flag，
已按实际 `--help` 更正。改 `DECLARED_WORKING` 之前先跑
`python3 tools/check_wheel_selfcontained.py --scan` 看实测，它只打印不判定。）

## 判据 C：包内副本 == HEAD

**「两份副本会不会漂」的答案是「仓库里根本没有第二份」**：`_bundled/` 只存在于构建
产物中，是构建那一刻的逐字节快照。**判据 C** 就是这句话的机器判据：包内文件集 ==
`git ls-tree -r HEAD -- scripts lib config schemas`，且逐文件 sha256 相同，
**两个方向都查**（多一个/少一个都判红）。

⚠ **这条判据 2026-08-02~08-03 之间只存在于文档里**：08-02 重写这道门时它被删掉，
而文档与 `pyproject.toml` 继续指着它说「漂移有人管」（`grep sha256\|ls-files` 当时
零命中）。一道被文档背书、实际不存在的门比没有门更坏 —— 读的人不会再去查。
08-03 已按文档承诺补回实现。

判据 C 补回来的**第一次运行就抓到一个致命问题**：`lib/llm_client.py` 是 git 里
mode 120000 的 **symlink**，指向仓外绝对路径 `~/Dev/tools/dev/scripts/tools/llm_client.py`。
原来 `"lib" = "doctools/_bundled/lib"` 整目录 force-include，于是**构建本身**依赖
那条绝对路径在本机解析得开 —— 实测把目标换成悬空链接后：

```
FileNotFoundError: [Errno 2] No such file or directory: '<proj>/lib/llm_client.py'
error: Call to `hatchling.build.build_wheel` failed (exit status: 1)
```

**即这个 wheel 在这台机器之外根本构建不出来，而判据 A 一直在打 ✓**（它在本机构建，
绝对路径恰好指得到）—— 与 08-02 那次 `$HOME` 假绿是同一形状：验「别的机器行不行」
时把本机的东西喂了进去。`[tool.hatch.build.targets.wheel].exclude` 救不了，
**force-include 不受 exclude 约束**（实测加了照样在包里）。

修法不是把文件拷一份进仓（那会造出第二份 SSOT），而是让它由 **`hq-devlib` 出**：
总部包已把 `scripts/tools/llm_client.py` 装进 site-packages 顶层，wheel 场景
`from llm_client import chat` 照样解析得到，本机场景仍走 `lib/` 下那条 symlink，
**两条路都不动任何一行 import**。doctools 侧 `lib` 改**逐文件** force-include
（16 行）把该 symlink 排除在外。漏加新文件不会静默 —— 判据 C 的「少一个」就是为它准备的。

判据 C 对 symlink 用的是**反向断言**：既不要求它在包里，也**不许**它在包里
（进了包 = 把机器依赖打了进去）。反向验证：把 `lib` 改回整目录写法 → 判据 C 判红并
点名 `仓外 symlink 混入: lib/llm_client.py`。

⚠ `cli_surface` / `cli_forward_probe` **看不见包内副本那条分支** —— 它们在工作树里跑，
命中的永远是第一条。wheel-only 分支唯一的测法就是上面这道门真去装一遍。
2026-08-02 基线实测：加 force-include 之前 wheel 只有 **7 个文件**，
clean venv 装完 `doctools --version` 直接 `FATAL: 找不到实现入口` rc=2。

## parallel_contract 三来源

`parallel_contract`（总部 SSOT，`~/Dev/tools/dev/lib/`）**不在本仓**，而
`docx_cli.py` / `pdf_cli.py` 缺它都是 **fail-closed exit 2 不是降级**
（pdf_cli 2026-08-04 改齐：旧 except 分支静默只留 `--workers`、丢
`--batch/--phases/--defer/--fanout-evidence`，同一版 CLI 在两台机器上 flag 面不一样
是最难查的漂移；反向验证 = 中立 HOME 下 `--help` 从「rc=0 少 flag」变 rc=2 FATAL）。
三条来源按此顺序：

| # | 来源 | 谁在用 |
|---|---|---|
| 1 | `~/Dev/tools/dev/lib`（`insert(0)`，**目录不存在就不塞**） | 本机所有直接敲绝对路径的调用 |
| 2 | `<根>/lib`（**append 不是 insert(0)**） | 留给包内镜像；本机这里**没有** parallel_contract.py |
| 3 | 装出来的 `hq-devlib` 包（site-packages 顶层） | wheel 装到别的机器上时 |

2 用 append 的理由：`lib/` 下有 styles.py / schemas.py / progress.py 等与他处同名的
模块，顶到 sys.path 首位会改变全进程解析优先级（形状对齐 `scripts/data/data.py`
既有先例）。3 排在最后，所以**本机行为一个字节没变** —— 解析照旧落到 1。
