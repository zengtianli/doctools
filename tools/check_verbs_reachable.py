#!/usr/bin/env python3
"""check_verbs_reachable — 每个注册过的顶层子命令，敲下去必须真的进得去。

    python3 tools/check_verbs_reachable.py          # 判红即 exit 1
    python3 tools/check_verbs_reachable.py --list

## 为什么要这个守卫（2026-08-02 立）

`docx_cli.main()` 自己做 token 切分（因为 rest 要原样透传给实现脚本），
判断「这个 token 是不是子命令」靠一份**手维护**的名字集合。它和 `_build_parser()`
里真正注册的 49 个顶层名是两份要人肉同步的东西 —— 于是漏了 4 个：
`fix` / `seqdiff` / `compare-ref` / `revise-rules`，共 **11 条动词（93 条的 12%）**。

漏掉的后果不是报错，是**静默无操作**：token 落进「未知」分支 → 顶层 parser 是
`add_help=False` + `parse_known_args` 不报错 → 打印根 help、**rc=0、不写盘、不报错**。
`docx_cli.py fix clear-direct-format X.docx --inplace` 看着像跑完了，连备份都没产生。

而当时 `cli_surface`(126 节点) / `cli_forward_probe`(67 条) / `check_function_axis`(93 条)
**rc 全 0** —— 前两者只看 parser 树与转发 argv，第三者只对账表和 parser，
**没有一个走 `main()` 那条真实入口**。这就是「闸门绿着，1/8 的动词是死的」。

那份集合现在已改为从 parser 派生（结构上不可能再漏同步）。本守卫是第二道：
不管将来实现怎么变，**每个注册过的顶层名敲下去都不许落进「打印根 help + rc=0」**。

判据有**两条**，命中任一即判红：

  A. `rc == 0 且 stdout 里出现根 help 的特征串` —— 静默无操作（漏同步的旧形态）
  B. stderr 里出现「未知子命令」—— 已注册的名字被 main() 当成不认识的（漏同步的新形态）

⚠ 只有 A 是不够的（2026-08-02 反向验证抓到）：同一轮里「未知 token」的出口从
「打 help + rc=0」改成了「报错 + rc=2」，于是把漏同步注入回去之后，
`fix` 变成 rc=2 的「未知子命令」，**A 判据当场失效、守卫报绿**。
一个只认得住旧形态的守卫，和它要防的 bug 是同一类东西。

（组命令缺子动作 → argparse rc=2；扁平命令缺参数 → 各自报错；`verbs` 正常输出清单。
 这三种既不是根 help、也不是「未知子命令」，都算通过。）

fail-closed：枚举为空 / 拿不到顶层名一律非 0 退出。
"""
from __future__ import annotations

import argparse
import subprocess
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
CLI = ROOT / "scripts" / "document" / "docx_cli.py"

# 根 help 的特征串（`_build_parser()` 的 description 首行）
ROOT_HELP_MARK = "doctools 文档处理统一 CLI"
UNKNOWN_MARK = "未知子命令"


def top_names() -> list[str]:
    sys.path.insert(0, str(CLI.parent))
    import docx_cli                                   # noqa: E402
    p = docx_cli._build_parser()
    names = set()
    for a in p._actions:
        if isinstance(a, argparse._SubParsersAction):
            names |= set(a.choices)
    names |= set(docx_cli.CMD_TABLE)
    return sorted(names)


def main() -> int:
    if not CLI.is_file():
        print(f"⛔ 找不到 {CLI} —— 拒绝在空集上报通过", file=sys.stderr)
        return 2
    names = top_names()
    if not names:
        print("⛔ 一个顶层子命令都没枚举到 —— 判据坏了，拒绝报绿", file=sys.stderr)
        return 2

    dead: list[str] = []
    for n in names:
        r = subprocess.run([sys.executable, str(CLI), n],
                           capture_output=True, text=True, timeout=120, cwd=str(ROOT))
        if (r.returncode == 0 and ROOT_HELP_MARK in r.stdout) or UNKNOWN_MARK in r.stderr:
            dead.append(n)
        if "--list" in sys.argv:
            mark = "✗" if n in dead else "✓"
            print(f"{mark} {n:20} rc={r.returncode}")

    if "--list" in sys.argv:
        return 0
    if dead:
        print(f"⛔ {len(dead)}/{len(names)} 个注册过的顶层子命令进不去"
              f"（要么静默打根 help 且 rc=0，要么被 main() 当成未知子命令）：", file=sys.stderr)
        for n in dead:
            print(f"    {n}", file=sys.stderr)
        print("\n多半是 main() 的 token 切分不认识它。那份集合应当从 parser 派生，"
              "\n别手维护 —— 手维护的那版就是这么漏了 fix/seqdiff/compare-ref/revise-rules "
              "\n共 11 条动词的。", file=sys.stderr)
        return 1
    print(f"✓ {len(names)} 个顶层子命令全部进得去（没有一个落进「根 help + rc=0」）")
    return 0


if __name__ == "__main__":
    sys.exit(main())
