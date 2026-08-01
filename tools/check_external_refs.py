#!/usr/bin/env python3
"""check_external_refs — 全生态引用 doctools 脚本路径的**存在性**闸门。

    python3 tools/check_external_refs.py            # 有死路径判红 exit 1
    python3 tools/check_external_refs.py --list     # 列全部命中（含活的）
    python3 tools/check_external_refs.py --json     # 机读

## 为什么要这个（2026-08-01 立）

doctools 的脚本被**仓外 200 多处**硬编码引用：~/Work 的项目脚本、cc-home 的 hook 与
skill、总部 dev/lib 的 subprocess 常量、~/Apps 的 Swift 与 catalog。2026-07-31 那轮
家族折叠（54 个文件并成 24 个）改完当天，靠三向人工核验才抓出一批断链 —— 也就是说
**这条面上没有任何机器守着**。

`/refactor dir` 的 rewrite-dead / scan-dead 帮不上：`dev/lib/tools/ssot/paths.py` 的
`_iter_scan_files` roots 只有 `~/Dev`，`SCAN_EXCLUDE_DIRS` 里明写 `"Work"`。而按实测，
引用的大头恰恰在 ~/Work 与 ~/Apps。

断链的形状很温和，这正是它危险的地方：
  · `hooks/docx-font-guard.sh` 路径断了 → 只 printf 一行「本次没查」继续放行，
    「所有 docx 禁等线」这条用户钦定硬约束**静默停摆**。
  · `~/Apps/mac/doc-tools/build.sh` 的解码契约自检是 `if [ -f "$BACKEND" ] && …`，
    没有 else 分支 —— 路径一断，自检被静默跳过，构建照样绿。

fail-closed：扫描根不存在、或一处命中都没扫到，一律非 0 退出（判据坏了比漏报更该红）。
"""
from __future__ import annotations

import json
import re
import subprocess
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
HOME = Path.home()

# 扫哪些地方。~/Work 与 ~/Apps 是重点 —— 它们正是 /refactor dir 够不到的地方。
SCAN_ROOTS = [
    HOME / "Work",
    HOME / "Dev/tools/cc-home",
    HOME / "Dev/tools/dev/lib",
    HOME / "Dev/tools/configs",
    HOME / "Dev/wiki",
    HOME / "Dev/stations",
    HOME / "Dev/apps",
    HOME / "Apps",
]

EXTS = ("py", "sh", "md", "yaml", "yml", "json", "swift", "toml")

# 排除两类：
# ① 归档/依赖树/oss 脱敏副本（它走相对路径引自己的 backend/，一律不碰）
# ② **会话历史**——cc-home 的 projects/(transcript+workflow) · jobs/ · plans/ ·
#    plugins/*/cache。这些是「当时说过什么」的记录，不是「现在会被执行的引用」。
#    历史里提到已退役的脚本名是正常的，把它算成断链会让守卫长红、进而被无视。
EXCLUDE = re.compile(
    r"/(_archive|archives?|\.Trash|node_modules|\.venv|site-packages|dist|build|"
    r"transcript-cache|__pycache__|\.git)/"
    r"|/_backup|/_pre-shim|/09-归档/|/Apps/oss/"
    r"|/cc-home/(projects|jobs|plans)/|/cc-home/plugins/[^/]+/(config-)?cache/"
    r"|/\.claude/(projects|jobs|plans|todos|state|statsig|shell-snapshots)/"
)

# 命中形状：doctools 仓内的一条脚本路径。前缀可以是 ~ 或 /Users/tianli 或 $HOME。
REF = re.compile(
    r"(?:~|/Users/tianli|\$HOME|\$\{HOME\})/Dev/tools/doctools/"
    r"([A-Za-z0-9_./-]+\.(?:py|sh|yaml|yml|json))"
)


def scan() -> list[dict]:
    """grep 全部扫描根，回 [{file, line, ref, exists}]。"""
    missing_roots = [r for r in SCAN_ROOTS if not r.is_dir()]
    roots = [r for r in SCAN_ROOTS if r.is_dir()]
    if not roots:
        print("⛔ 一个扫描根都不存在 —— 拒绝在空集上报绿", file=sys.stderr)
        raise SystemExit(2)
    if missing_roots:
        print(f"⚠ 扫描根缺失（跳过）：{[str(r) for r in missing_roots]}", file=sys.stderr)

    cmd = ["grep", "-rn", "--binary-files=without-match"]
    for e in EXTS:
        cmd += [f"--include=*.{e}"]
    cmd += ["-E", r"(~|/Users/tianli|\$\{?HOME\}?)/Dev/tools/doctools/"]
    cmd += [str(r) for r in roots]
    # grep rc: 0=有命中 1=无命中 2=出错。无命中在这里是判据坏了。
    proc = subprocess.run(cmd, capture_output=True, text=True)
    if proc.returncode == 2 and not proc.stdout:
        print(f"⛔ grep 失败：{proc.stderr[:400]}", file=sys.stderr)
        raise SystemExit(2)

    hits = []
    for raw in proc.stdout.splitlines():
        try:
            path, lineno, text = raw.split(":", 2)
        except ValueError:
            continue
        if EXCLUDE.search(path + "/"):
            continue
        for m in REF.finditer(text):
            rel = m.group(1)
            hits.append({
                "file": path,
                "line": int(lineno),
                "ref": rel,
                "exists": (ROOT / rel).exists(),
                "text": text.strip()[:160],
            })
    if not hits:
        print("⛔ 一处 doctools 路径引用都没扫到 —— 判据肯定坏了，拒绝报绿", file=sys.stderr)
        raise SystemExit(2)
    return hits


def main() -> int:
    hits = scan()
    dead = [h for h in hits if not h["exists"]]
    refs = {h["ref"] for h in hits}

    if "--json" in sys.argv:
        print(json.dumps({"hits": hits, "dead": dead}, ensure_ascii=False, indent=1))
        return 1 if dead else 0

    if "--list" in sys.argv:
        for h in sorted(hits, key=lambda x: (not x["exists"], x["file"])):
            mark = "✓" if h["exists"] else "✗"
            print(f"{mark} {h['ref']:52} {h['file'].replace(str(HOME), '~')}:{h['line']}")
        return 1 if dead else 0

    if dead:
        by_ref: dict[str, list[dict]] = {}
        for h in dead:
            by_ref.setdefault(h["ref"], []).append(h)
        print(f"⛔ {len(dead)} 处引用指向**不存在**的 doctools 脚本"
              f"（共 {len(hits)} 处引用 / {len(refs)} 条不同路径）：", file=sys.stderr)
        for ref, hs in sorted(by_ref.items()):
            print(f"\n  ✗ {ref}", file=sys.stderr)
            for h in hs:
                print(f"      {h['file'].replace(str(HOME), '~')}:{h['line']}", file=sys.stderr)
                print(f"        {h['text']}", file=sys.stderr)
        print("\n修法：把引用改指现在的真实路径（家族折叠后旧脚本名多半变成了子命令，"
              "对照 README.md 的入口表与 CLAUDE.md 的独立入口表）。"
              "\n注意断链是静默的：hook 与 build.sh 的存在性判断多是 fail-open，"
              "路径断了它们不报错、只是不干活。", file=sys.stderr)
        return 1

    print(f"✓ {len(hits)} 处仓外引用 / {len(refs)} 条不同路径，全部指向存在的文件")
    return 0


if __name__ == "__main__":
    sys.exit(main())
