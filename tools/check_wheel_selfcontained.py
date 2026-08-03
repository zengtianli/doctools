#!/usr/bin/env python3
"""wheel 分发能力对账门（2026-08-02 重写 —— 上一版是一道假门）。

    python3 tools/check_wheel_selfcontained.py            # 判红即 exit 1
    python3 tools/check_wheel_selfcontained.py --scan     # 只打印实测能力，不判定

## 上一版错在哪（两处，都让它报了假绿）

1. **install 层的两条 check 对「实现进没进包」零分辨率**：跑的是 `--version` 与
   `verbs --fn convert`，而 `verbs` 是纯清单打印，从头到尾不加载任何实现模块。
   最刺眼的是 `verbs --fn convert` 打出来的正好是 `md-to-docx` 与 `md` ——
   它把两个在那个 wheel 里已经死掉的动词的名字打印出来，然后判绿。

2. **「仓外 clean venv 跑通」是被 `$HOME` 喂出来的**：`docx_cli.py` 有一句
   `_LIB = Path.home()/"Dev"/"tools"/"dev"/"lib"` 兜底导入，本机 `$HOME` 永远
   摸得到总部 lib。实测把 `HOME` 换成空目录后，同一个 wheel 从 49/49 变成 0/49。
   **在本机验「别的机器能不能用」，不中和 `$HOME` 就是白验。**

## 现在的判据（两条，都必须真跑）

  A. **构建可移植**：把 HEAD 导出到一个**没有兄弟目录 `dev/` 的位置**，能构建出 wheel。
     （2026-08-02 之前 pyproject 里有 `"../dev/lib/parallel_contract.py"` 这个跨仓
     force-include，于是 GitHub runner / 任何别人的机器上构建当场
     `FileNotFoundError: Forced include not found` —— 而 hatchling 对 **editable**
     构建同样施加 force-include，所以连 `uv sync` 都红。CI 从加上那天起没绿过一次。）

  B. **运行能力 == 声明值**：wheel 装进 clean venv，**`HOME` 指到空目录**，
     逐个敲顶层动词，实际能跑的条数必须**恰好等于** `DECLARED_WORKING`。

判据 B 的意义不是「必须全能跑」，而是**不许对能力撒谎**：实测多少就写多少，
声明与现实对不上就判红，多了少了两个方向都堵死。

## 判据 B 装包时**带依赖装**（2026-08-02 改，别再加回 `--no-deps`）

原来这里是 `pip install --no-deps`，那测的是「只有 doctools 自己的代码」——
而本仓 49 条动词全都要 `parallel_contract`，一多半还要 `yaml` / `lxml` /
`python-docx`。`--no-deps` 装法下这门永远只能测出「0 条可用」，且卡点会随着
修复一路搬家（parallel_contract → yaml → lxml），把「依赖声明对不对」这个真问题
挡在门外。现在按**真实分发姿势**装：`pip install <doctools.whl>`，依赖照解。

`hq-devlib`（= 总部 `~/Dev/tools/dev/lib` 的 6 个平铺模块的可分发形态）现在由
pyproject 声明成 **git 直接引用**（`hq-devlib @ git+https://github.com/zengtianli/
devtools.git`，2026-08-03 用户拍板：走 git URL，不发 PyPI），所以装包这一步会真的
去 GitHub clone 它。**本门原来在本地把它构建成 wheel 丢进 `--find-links` 目录当
「它在某个 index 上可取」的替身，那一段已经删掉** —— 不是因为嫌慢，是因为它从
声明改成直接引用的那一刻起就**永远不会被用到**：PEP 508 直接引用的优先级高于任何
index / find-links。实测证据（2026-08-03，`uv pip install --find-links <放着
hq_devlib-0.1.0-py3-none-any.whl 的目录> doctools-0.1.0-*.whl`）：

    Updating https://github.com/zengtianli/devtools.git (HEAD)
    + hq-devlib==0.1.0 (from git+https://github.com/zengtianli/devtools.git@6431dfd…)

—— 本地那个 wheel 一眼都没看。留着它只会让这门看起来在测「index 上取得到」，
而它实际测的是「git URL clone 得动」，是一段会骗人的死代码。

⚠ 由此本门多了两个前置条件，红了先看是不是它们：
  · **要能上网**；
  · **要有 zengtianli/devtools 的读权限**（该仓是 PRIVATE，`gh repo view … --json
    visibility` 实测）。本机走 git 的常规凭证（credential helper）即可，装包这一步
    用的是真实环境；**只有跑动词那一步才把 `HOME` 换成空目录**。
两者缺一 → 装包失败 → SystemExit(2)，fail-closed，不会退化成假绿。

**中和 `HOME` 这条没变**：装完之后跑动词时 `HOME` 是空目录、`PYTHONPATH` 被清掉，
`docx_cli.py` 的 `Path.home()/"Dev"/...` 兜底照样摸不到任何东西。要验的正是
「离开这台机器的目录布局之后还跑不跑得动」。

fail-closed：构建失败 / 装不上（含 clone 不动）/ 枚举为空，一律非 0 退出。
"""
from __future__ import annotations

import argparse
import os
import shutil
import subprocess
import sys
import tempfile
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]

# ── 能力声明（改这个数之前先跑 --scan 看实测）────────────────────────────
# 49 = clean venv + 中立 HOME 下**全部**顶层动词都起得来（2026-08-02 实测）。
#
# 变成 49 之前它是 0：本仓运行时依赖总部 `~/Dev/tools/dev/lib/` 的 6 个**平铺模块**
# （finder×10 · file_ops×10 · display×9 · parallel_contract×3 · usage_log×1 · env×1，
# 共 1156 行），它们不是包、只能靠 sys.path 注入导入。走的是上面注释里的出路 (b)：
# 总部仓加 `[build-system]` + `only-include` 那 6 个 + `sources = ["lib"]`，打成
# 分发名 `hq-devlib`（`lib/*.py` 零改动、44 处 sys.path 注入原样继续工作），
# 本仓 pyproject 声明 `hq-devlib>=0.1` 依赖。
#
# ⚠ 这个数只说明「动词起得来」（每条敲 `--help`，import 链全通），**不等于**
# 每条动词的完整功能都在别的机器上验过。别把它当功能覆盖率读。
# 2026-08-03 由 48 → 49：`text-fmt --help` 也修好了（`scripts/document/docx_fmt.py`
# 的 text_main() 开头加 -h/--help 拦截 → 打 TEXT_USAGE 并 return 0；位置必须在
# `--scope` 取值与未知 flag 判定之前，否则 `--scope --help` 会先被 next(_it) 吃掉）。
# 至此 49 条顶层动词的 `--help` 全部 rc=0，不再有「既有非 0」的例外。
#
# 2026-08-03 由 47 → 48：`image-caption --help` 修好了。它原来 rc=1 的根因不是 import，
# 是**没有 argparse 而 flag 会 fallthrough** —— main() 把 argv 直接丢给 get_input_files()，
# 后者见「没有位置参数」就回落去读 Finder 选中项，于是 `--help` 去动用户此刻选中的文件
# （同 CLAUDE.md「破坏性动作必须自己占一个动词」判过的那条死刑）。现在 sub/
# docx_apply_image_caption.py 顶部有闸门，在任何 import / Finder 读之前拦下。
#
# ⚠ 本门的 wheel 是 `git archive HEAD` 构建的，**读的是 HEAD 不是工作树**：改完还没
# commit 时它仍按老代码测出 47，与这里的 48 不符而报红，那是时序不是回归，commit 后即绿。
# 改这个数当时的证据（工作树，中立 HOME + 清 PYTHONPATH，等价于本门判据 B 的环境）：
#   HEAD 版  → rc=1  ModuleNotFoundError: No module named 'file_ops'
#   工作树版 → rc=0  stderr 空，stdout 是用法
# 并且用本门自己跑过：把 archive 的 ref 临时换成「HEAD + 只有这一个文件的改动」构出的
# 树（git commit-tree 的悬空提交，没碰分支/index），本门实测 48/49、判据 B 绿。
#
# 2026-08-03 收口实测（本门自己跑的，不是推的）：把**整棵工作树**做成悬空提交
# （`git commit-tree`，用临时 GIT_INDEX_FILE，没碰分支/真 index/工作树）再让本门
# 按那个 ref 去 `git archive`：
#   判据 A ✓ 构建成功 · 判据 B ✓ 装得上（pyproject 的 git URL 生效）
#   「clean venv + 中立 HOME 下 49/49 个顶层动词可用」→ 与本声明相符
# 同一时刻按字面 HEAD 跑则是 rc=2「hq-devlib was not found in the package registry」——
# 因为 HEAD 的 pyproject 还是裸 `hq-devlib>=0.1`，而 git URL 那版还没 commit。
# **两个红都是时序不是回归**：image-caption / text-fmt / pyproject 三处一起 commit 后即绿。
DECLARED_WORKING = 49


_ENUM_SNIPPET = """
import argparse, json, importlib.util, sys
spec = importlib.util.find_spec("doctools")
import doctools.cli as c
m = c._load_docx_cli()
p = m._build_parser()
names = {n for a in p._actions if isinstance(a, argparse._SubParsersAction) for n in a.choices}
print(json.dumps(sorted(names | set(m.CMD_TABLE))))
"""


def tops_from_wheel(py: Path, env: dict) -> list[str]:
    """动词清单必须**从被测 wheel 里**枚举，不能从工作树。

    原来是 `sys.path.insert(ROOT/"scripts"/"document"); import docx_cli` —— 那是
    **工作树**的 parser，而被测 wheel 是 `git archive HEAD` 构建的，两者可以不一致。
    2026-08-02 反向验证实证：往工作树 CMD_TABLE 注入一条 wheel 里根本没有的
    `ghost-verb-not-in-wheel`，再按本门红字提示把声明值 +1，本门当场报
    「50/50 可用 ✓」——而真敲那条动词是 `rc=2 未知子命令`。
    这正是本门 docstring 自称杀掉的那类假绿，机制只是从 verbs 打印挪到了枚举源。
    """
    r = subprocess.run([str(py), "-c", _ENUM_SNIPPET], capture_output=True, text=True,
                       cwd="/tmp", env=env, timeout=120)
    if r.returncode != 0:
        print(f"⛔ 从 wheel 里枚举动词失败 —— 拒绝退回工作树枚举\n{r.stderr[-400:]}",
              file=sys.stderr)
        raise SystemExit(2)
    import json as _json
    return _json.loads(r.stdout.strip().splitlines()[-1])


def build_portable(workdir: Path) -> Path | None:
    """判据 A：在没有兄弟 `dev/` 的位置构建。返回 wheel 路径，失败返回 None。"""
    proj = workdir / "proj"
    proj.mkdir(parents=True)
    tar = subprocess.run(["git", "-C", str(ROOT), "archive", "HEAD"], capture_output=True)
    if tar.returncode != 0:
        print("⛔ git archive 失败 —— 拒绝在拿不到源码时报绿", file=sys.stderr)
        return None
    subprocess.run(["tar", "-x", "-C", str(proj)], input=tar.stdout, check=True)
    out = workdir / "dist"
    r = subprocess.run([sys.executable, "-m", "pip", "wheel", "--no-deps", "-w", str(out), "."],
                       cwd=str(proj), capture_output=True, text=True)
    if r.returncode != 0:
        tail = "\n".join(r.stdout.splitlines()[-6:] + r.stderr.splitlines()[-6:])
        print(f"⛔ 判据 A 失败：wheel 在没有兄弟 dev/ 的位置构建不出来\n{tail}", file=sys.stderr)
        return None
    whls = sorted(out.glob("doctools-*.whl"))
    if not whls:
        print("⛔ 构建报成功却没产出 wheel", file=sys.stderr)
        return None
    return whls[0]


def probe_capability(whl: Path, workdir: Path) -> tuple[int, int, dict[str, str]]:
    """判据 B：clean venv + 中立 HOME，逐个敲顶层动词。→ (能跑, 总数, 挂掉的)"""
    venv = workdir / "venv"
    subprocess.run([sys.executable, "-m", "venv", str(venv)], check=True, capture_output=True)
    py, exe = venv / "bin" / "python", venv / "bin" / "doctools"
    # **带依赖装**（别加回 --no-deps，理由见模块 docstring）。有 uv 就用 uv：同一套
    # 依赖 pip 要跑几分钟，uv 走本机缓存 ~26s；装出来的 site-packages 内容一致。
    #
    # ⚠ 这里**不许再加 `--find-links <本地 hq-devlib wheelhouse>`**（2026-08-03 删）：
    # hq-devlib 已是 PEP 508 git 直接引用，直接引用优先级高于任何 index/find-links，
    # 那个目录实测一眼都不会被看（见模块 docstring 里的实测输出）。加回去 = 一段
    # 让这门看起来在测别的东西的死代码。本步骤走**真实环境**（真 HOME → 真 git 凭证），
    # 中和 HOME 只发生在下面跑动词那一步。
    uv = shutil.which("uv")
    cmd = ([uv, "pip", "install", "--python", str(py), str(whl)]
           if uv else
           [str(venv / "bin" / "pip"), "install", "-q", str(whl)])
    r = subprocess.run(cmd, capture_output=True, text=True)
    if r.returncode != 0:
        print(f"⛔ 判据 B 失败：wheel 连同依赖装不上（installer={'uv' if uv else 'pip'}）\n"
              f"   若报的是 `could not read Username for 'https://github.com'` / "
              f"`Repository not found`：hq-devlib 是 git 直接引用，而 "
              f"zengtianli/devtools 是 **私库** —— 这台机器要么没网，要么没有它的读权限。\n"
              f"{r.stdout[-500:]}\n{r.stderr[-500:]}", file=sys.stderr)
        raise SystemExit(2)

    fake_home = workdir / "home"      # ← 关键：本机 $HOME 会把总部 lib 喂进去
    fake_home.mkdir()
    env = {k: v for k, v in os.environ.items() if k != "PYTHONPATH"}
    env["HOME"] = str(fake_home)

    tops = tops_from_wheel(py, env)
    if not tops:
        print("⛔ 一个顶层动词都没枚举到 —— 判据坏了，拒绝报绿", file=sys.stderr)
        raise SystemExit(2)
    ok, broken = 0, {}
    for n in tops:
        p = subprocess.run([str(exe), n, "--help"], capture_output=True, text=True,
                           cwd="/tmp", env=env, timeout=120)
        # 判据必须同时看 **returncode** 和 import 错误。原来只 grep stderr 里的
        # "No module named"，于是「未知子命令 rc=2」「argparse 报错 rc=2」这类
        # 启动崩溃一条都拦不住 —— 实测 image-caption rc=1 / text-fmt rc=2 当时被计为 ok。
        if "No module named" in p.stderr:
            broken[n] = f"ImportError: {p.stderr.split('No module named')[-1].strip().splitlines()[0]}"
        elif p.returncode != 0:
            first = (p.stderr.strip().splitlines() or [""])[0][:60]
            broken[n] = f"rc={p.returncode} {first}"
        else:
            ok += 1
    return ok, len(tops), broken


def check_bundled_snapshot(whl: Path) -> bool:
    """判据 C：包内 `doctools/_bundled/` == HEAD 里那四个根，**两个方向都查**。

    `scripts/` `lib/` 不能搬（~/Work 有 130 处绝对路径钉着），所以用 force-include 在
    构建时镜像一份进包。「那两份副本会不会漂」的答案是「仓库里根本没有第二份」——
    `_bundled/` 只存在于构建产物里，是构建那一刻的逐字节快照。**这个函数就是那句话
    的机器判据**，`pyproject.toml:122-125` 与 CLAUDE.md 都指着它。

    ⚠ 2026-08-03 补写：它此前**只存在于文档里**。08-02 重写这道门时被删掉，而两处
    文档继续承诺着它 —— 一道被文档背书、实际不存在的门，比没有门更坏：读的人会以为
    漂移有人管。（`grep -n 'sha256\\|ls-files' tools/check_wheel_selfcontained.py`
    当时零命中，是核验镜头逐字比对文档与实现时发现的。）

    两个方向都判红：包里多一个（force-include 源端扫进了不该进的东西）与少一个
    （某个根漏配、或 VCS ignore 把该进的挡了）都是漂移。逐文件 sha256 再抓内容级差异。
    比较对象是 **HEAD** 不是工作树 —— 因为 wheel 就是 `git archive HEAD` 构建的，
    拿工作树比会在「改了没提交」时报出一片假红。
    """
    import hashlib
    import zipfile

    roots = ("scripts", "lib", "config", "schemas")
    r = subprocess.run(["git", "-C", str(ROOT), "ls-tree", "-r", "HEAD",
                        "--format=%(objectmode) %(path)", "--", *roots],
                       capture_output=True, text=True)
    if r.returncode != 0:
        print(f"⛔ 判据 C：拿不到 HEAD 的文件清单 —— 拒绝在没有基准时报绿\n{r.stderr[-300:]}",
              file=sys.stderr)
        return False
    want, links = set(), set()
    for ln in r.stdout.splitlines():
        if not ln.strip():
            continue
        mode, path = ln.split(" ", 1)
        # mode 120000 = symlink。本仓有一条（lib/llm_client.py → 总部 SSOT 的绝对路径）。
        # 它**必须留在包外**：整目录 force-include 会让构建依赖那条绝对路径解析得开，
        # 实测目标悬空时 hatchling 直接 FileNotFoundError、wheel 在本机之外根本构建不出来。
        # 所以对 symlink 的断言是**反向**的 —— 不在包里才对，进了包就是把机器依赖打了进去。
        (links if mode == "120000" else want).add(path)
    if not want:
        print("⛔ 判据 C：HEAD 里这四个根一个文件都没有 —— 判据坏了，拒绝在空集上报绿",
              file=sys.stderr)
        return False

    with zipfile.ZipFile(whl) as z:
        pre = "doctools/_bundled/"
        got = {n[len(pre):]: n for n in z.namelist()
               if n.startswith(pre) and not n.endswith("/")}
        if not got:
            print("⛔ 判据 C：wheel 里没有 doctools/_bundled/ —— force-include 整个没生效。"
                  "\n   这正是 2026-08-02 的基线态：那时 wheel 只有 7 个文件，"
                  "clean venv 装完 `doctools --version` 直接 FATAL。", file=sys.stderr)
            return False
        missing = sorted(want - set(got))
        extra = sorted(set(got) - want - links)
        # 仓外 symlink 进了包 = 构建的机器依赖被打了进去，单列一类，别混进「多一个」
        smuggled = sorted(links & set(got))
        drift = []
        for rel in sorted(want & set(got)):
            blob = subprocess.run(["git", "-C", str(ROOT), "show", f"HEAD:{rel}"],
                                  capture_output=True).stdout
            if hashlib.sha256(blob).hexdigest() != hashlib.sha256(z.read(got[rel])).hexdigest():
                drift.append(rel)

    if missing or extra or drift or smuggled:
        print(f"⛔ 判据 C 失败：包内副本与 HEAD 不一致（少 {len(missing)} / 多 {len(extra)}"
              f" / 内容不同 {len(drift)} / 仓外 symlink 被打进包 {len(smuggled)}）",
              file=sys.stderr)
        for tag, lst in (("少", missing), ("多", extra), ("内容不同", drift),
                         ("仓外 symlink 混入", smuggled)):
            for rel in lst[:8]:
                print(f"    {tag}: {rel}", file=sys.stderr)
            if len(lst) > 8:
                print(f"    {tag}: …… 另 {len(lst) - 8} 条", file=sys.stderr)
        if missing:
            print("\n   「少」多半是 force-include 漏配（lib/ 是逐文件列的，加了新文件要补一行）。",
                  file=sys.stderr)
        if smuggled:
            print("\n   「仓外 symlink 混入」= 构建的机器依赖被打进了包：那条链接指向仓外绝对\n"
                  "   路径，目标不存在时 hatchling 直接 FileNotFoundError，wheel 在别的机器上\n"
                  "   根本构建不出来。该模块应由依赖（hq-devlib）提供，不是镜像进包。",
                  file=sys.stderr)
        return False
    note = f"，另 {len(links)} 条仓外 symlink 按约定留在包外（由 hq-devlib 提供）" if links else ""
    print(f"✓ 判据 C：包内 _bundled/ 的 {len(want)} 个文件与 HEAD 逐个 sha256 相同"
          f"（多一个/少一个同样判红）{note}")
    return True


def main() -> int:
    scan_only = "--scan" in sys.argv
    with tempfile.TemporaryDirectory(prefix="doctools-dist-") as d:
        work = Path(d)
        whl = build_portable(work)
        if whl is None:
            return 1
        print(f"✓ 判据 A：wheel 在没有兄弟 dev/ 的位置构建成功（{whl.name}）")
        if not check_bundled_snapshot(whl) and not scan_only:
            return 1
        ok, total, broken = probe_capability(whl, work)

    print(f"  判据 B 实测：clean venv + 中立 HOME 下 {ok}/{total} 个顶层动词可用")
    if broken:
        for n, why in sorted(broken.items()):
            print(f"    ✗ {n:20} {why}")
    if scan_only:
        return 0
    if ok != DECLARED_WORKING:
        print(f"⛔ 判据 B 失败：实测 {ok} 条可用，而 DECLARED_WORKING 声明 "
              f"{DECLARED_WORKING} 条。\n"
              f"   本门不要求「必须全能跑」，要求的是**不许对能力撒谎** —— "
              f"能力变了就同步改声明（改之前先跑 --scan）。", file=sys.stderr)
        return 1
    print(f"✓ 判据 B：实测可用动词数 == 声明值 {DECLARED_WORKING}")
    if DECLARED_WORKING < total:
        print(f"ℹ {total - DECLARED_WORKING} 条动词起不来（原因见上面逐条列出的 rc / "
              f"ImportError）。这不必然是分发问题 —— 本机工作树上同样非 0 的属既有问题。")
    return 0


if __name__ == "__main__":
    sys.exit(main())
