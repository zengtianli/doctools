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

`hq-devlib`（= 总部 `~/Dev/tools/dev/lib` 的 6 个平铺模块的可分发形态）尚未发到
任何公开 index，所以这里先在本地把它构建成 wheel 丢进一个 `--find-links` 目录，
等价于「它在某个 index 上可取」。**这不是把本机 `~/Dev` 喂给被测 wheel** ——
两者的区别是：`--find-links` 只影响**装包**那一步，装完之后跑动词时 `HOME` 仍是
空目录、`PYTHONPATH` 仍被清掉，`docx_cli.py` 的 `Path.home()/"Dev"/...` 兜底照样
摸不到任何东西。要验的正是「离开这台机器的目录布局之后还跑不跑得动」。

fail-closed：构建失败 / 装不上 / 枚举为空 / 找不到总部仓，一律非 0 退出。
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
DECLARED_WORKING = 49


def _tops() -> list[str]:
    sys.path.insert(0, str(ROOT / "scripts" / "document"))
    import docx_cli                                        # noqa: E402
    p = docx_cli._build_parser()
    names = {n for a in p._actions
             if isinstance(a, argparse._SubParsersAction) for n in a.choices}
    return sorted(names | set(docx_cli.CMD_TABLE))


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


def build_hq_devlib(workdir: Path) -> Path:
    """把总部 `tools/dev` 打成 wheel 丢进本地 wheelhouse，供判据 B 的装包步解析。

    总部仓 = 本仓的**兄弟目录**（`ROOT.parent / "dev"`），不写 `$HOME` 字面量。
    hq-devlib 还没上任何公开 index，本地 wheelhouse 就是「它在 index 上可取」的替身；
    发上去之后这一步删掉即可，判据不变。找不到总部仓 → 直接非 0 退出，不静默跳过
    （静默跳过会让这门退化回 `--no-deps` 那种「测的是别的东西」）。
    """
    src = ROOT.parent / "dev"
    if not (src / "pyproject.toml").is_file():
        print(f"⛔ 找不到总部仓 {src} —— hq-devlib 造不出来，拒绝在测不了依赖时报绿",
              file=sys.stderr)
        raise SystemExit(2)
    house = workdir / "wheelhouse"
    house.mkdir(parents=True, exist_ok=True)
    r = subprocess.run([sys.executable, "-m", "pip", "wheel", "--no-deps", "-w", str(house), str(src)],
                       capture_output=True, text=True)
    if r.returncode != 0 or not list(house.glob("hq_devlib-*.whl")):
        tail = "\n".join(r.stdout.splitlines()[-6:] + r.stderr.splitlines()[-6:])
        print(f"⛔ hq-devlib 构建失败（源 {src}）\n{tail}", file=sys.stderr)
        raise SystemExit(2)
    return house


def probe_capability(whl: Path, workdir: Path) -> tuple[int, int, dict[str, str]]:
    """判据 B：clean venv + 中立 HOME，逐个敲顶层动词。→ (能跑, 总数, 挂掉的)"""
    house = build_hq_devlib(workdir)
    venv = workdir / "venv"
    subprocess.run([sys.executable, "-m", "venv", str(venv)], check=True, capture_output=True)
    py, exe = venv / "bin" / "python", venv / "bin" / "doctools"
    # **带依赖装**（别加回 --no-deps，理由见模块 docstring）。有 uv 就用 uv：同一套
    # 依赖 pip 要跑几分钟，uv 走本机缓存 ~26s；装出来的 site-packages 内容一致。
    uv = shutil.which("uv")
    cmd = ([uv, "pip", "install", "--python", str(py), "--find-links", str(house), str(whl)]
           if uv else
           [str(venv / "bin" / "pip"), "install", "-q", "--find-links", str(house), str(whl)])
    r = subprocess.run(cmd, capture_output=True, text=True)
    if r.returncode != 0:
        print(f"⛔ 判据 B 失败：wheel 连同依赖装不上（installer={'uv' if uv else 'pip'}）\n"
              f"{r.stdout[-500:]}\n{r.stderr[-500:]}", file=sys.stderr)
        raise SystemExit(2)

    fake_home = workdir / "home"      # ← 关键：本机 $HOME 会把总部 lib 喂进去
    fake_home.mkdir()
    env = {k: v for k, v in os.environ.items() if k != "PYTHONPATH"}
    env["HOME"] = str(fake_home)

    tops = _tops()
    if not tops:
        print("⛔ 一个顶层动词都没枚举到 —— 判据坏了，拒绝报绿", file=sys.stderr)
        raise SystemExit(2)
    ok, broken = 0, {}
    for n in tops:
        p = subprocess.run([str(exe), n, "--help"], capture_output=True, text=True,
                           cwd="/tmp", env=env, timeout=120)
        if "No module named" in p.stderr:
            broken[n] = p.stderr.split("No module named")[-1].strip().splitlines()[0]
        else:
            ok += 1
    return ok, len(tops), broken


def main() -> int:
    scan_only = "--scan" in sys.argv
    with tempfile.TemporaryDirectory(prefix="doctools-dist-") as d:
        work = Path(d)
        whl = build_portable(work)
        if whl is None:
            return 1
        print(f"✓ 判据 A：wheel 在没有兄弟 dev/ 的位置构建成功（{whl.name}）")
        ok, total, broken = probe_capability(whl, work)

    print(f"  判据 B 实测：clean venv + 中立 HOME 下 {ok}/{total} 个顶层动词可用")
    if broken:
        print(f"  缺的模块：{', '.join(sorted(set(broken.values())))}"
              f"（{len(broken)} 个动词受影响）")
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
        print(f"ℹ 当前 wheel **不可分发**：{total - DECLARED_WORKING} 条动词依赖总部 "
              f"~/Dev/tools/dev/lib 的平铺模块。出路见本文件 DECLARED_WORKING 处的注释。")
    return 0


if __name__ == "__main__":
    sys.exit(main())
