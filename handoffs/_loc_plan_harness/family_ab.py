#!/usr/bin/env python3
"""family_ab.py — 12 个 sub/ 家族 main() 分发器的 CLI 行为快照（W1 的唯一验收门）。

    python3 handoffs/_loc_plan_harness/family_ab.py <仓库根绝对路径> <输出 json>

为什么必须有它：`cli_surface` 只管接口形状，`cli_forward_probe` 把 `exec_script`
换成录音机、**根本进不到家族 main()**。W1 动的恰恰是家族 main()，那两道门在这层是瞎的。
改前改后各跑一次，`diff` 必须完全空。

norm() 抹掉绝对路径 / traceback 行号 / 对象地址 —— 这三样每次跑都不同，不抹就没法比。
**别顺手把栈帧也抹了**：W1 会让 traceback 多一层 `_cli_common.py in family_main`，
那正是要人眼确认的东西（只允许出现在本来就崩的既存 case 上，rc 与异常文本必须一字不差）。
"""
import json
import re
import subprocess
import sys
from pathlib import Path

FAMS = ["audit", "biddiff", "blocks", "caption", "chapter", "freeze",
        "image", "renumber", "split", "strip", "table", "typeset_ops"]

# 12 例覆盖：空 argv / 两种 help / 大小写 / 未知子命令 / 带尾参 / 空串 /
# 裸横线 / 裸双横线 / 未知 flag / help 带尾参 / 混合大小写子命令
CASES = [[], ["-h"], ["--help"], ["--Help"], ["bogus"], ["bogus", "x"],
         [""], ["-"], ["--"], ["--nope"], ["-h", "extra"], ["Bogus-Sub"]]


def norm(s: str, root: str) -> str:
    s = s.replace(root, "@ROOT@")
    s = re.sub(r'File "[^"]+", line \d+', 'File "@F@", line @L@', s)
    return re.sub(r"0x[0-9a-f]+", "@ADDR@", s)


def main() -> int:
    if len(sys.argv) != 3:
        print(__doc__.strip().splitlines()[2].strip(), file=sys.stderr)
        return 2
    root = str(Path(sys.argv[1]).resolve())
    out: dict[str, list] = {}
    for fam in FAMS:
        script = f"{root}/scripts/document/sub/{fam}.py"
        if not Path(script).is_file():          # fail-closed：少一个族就别出快照
            print(f"⛔ 不存在: {script} —— 家族名单已漂，拒绝出不完整的快照", file=sys.stderr)
            return 2
        for case in CASES:
            r = subprocess.run([sys.executable, script, *case], capture_output=True,
                               text=True, timeout=120, cwd=root)
            out[f"{fam}|{case!r}"] = [r.returncode, norm(r.stdout, root), norm(r.stderr, root)]
    if len(out) != len(FAMS) * len(CASES):
        print(f"⛔ 期望 {len(FAMS) * len(CASES)} 例，实得 {len(out)}", file=sys.stderr)
        return 2
    Path(sys.argv[2]).write_text(json.dumps(out, ensure_ascii=False, indent=1, sort_keys=True))
    print(f"cases={len(out)} -> {sys.argv[2]}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
