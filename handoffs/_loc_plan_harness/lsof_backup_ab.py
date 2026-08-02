#!/usr/bin/env python3
"""lsof_backup_ab.py — 收敛后的一致性门：10 处 lsof + 备份路径是否真是同一份实现。

    python3 handoffs/_loc_plan_harness/lsof_backup_ab.py     # rc=0 即四项全过

## 判据换过一次（2026-08-02），别按旧的读

**改写前**判的是「A 派内部一致 + B 派内部一致 + A.strip() == B」——那是**共存期**的
判据：当时全仓有 4 种语义（A 裸 lsof / B 带 `--` 且要求行数>1 / C 返 bool 不看
returncode / D 无 timeout），门只能证明「同派内没有再分叉」，不能证明它们该合并。

**改写后**（用户拍板取最健壮的 B 语义收敛）判的是**收敛本身**：

  1. **10 处 lsof 全部返回逐字相同的结果** —— 8 处 `str|None` 直接串比对，
     2 处 bool 外壳（`caption.lsof_check_bool` 返 True=占用、
     `pptx_cli._lsof_guard` 返 True=**空闲**，方向相反）按各自外壳折算后比对。
     任何一处偷偷保留本地实现 → 当场红。
  2. **`--` 真的生效** —— 造一个以 `-` 开头的文件名并真开一个 fd 持有它，
     确认 `lsof -- <name>` 认出占用，而裸 `lsof <name>` 把它当 flag 吃掉
     （实测 `lsof -hold.docx` = `-h` 帮助：**rc=0、stdout 空**，于是老的 A/C/D
     三派一致判「空闲」→ 对着 Word 正开着的文件放行写盘）。
     这一条是这次收敛唯一的**真行为变更**，所以必须有断言钉住它，
     否则哪天有人把 `--` 顺手删了，上面第 1 条照样全绿。
  3. **备份路径 6 处委派同解** —— 7 种文件名 × 3 轮 N 自增，0 mismatch。
  4. **两处内联备份循环已拆掉** —— `add_header_footer.py` / `blocks.py` 里那三段
     手写 `while True: .bak-{n}-{today}` 必须消失且改成 `_cc.find_next_backup`。
     （源码级断言：内联循环跟 1./3. 的输出恰好一样，行为门看不见它还在。）

⚠ 必须 import 生产实现，禁自己按「我以为它是这么实现的」重写一个替身来测
（铁律 #2）。第 2 条里那个「裸 lsof」对照组是**故意**手写的 —— 它扮演的是
**已被删掉的旧实现**，不是被测对象。
"""
from __future__ import annotations

import importlib
import os
import subprocess
import sys
import tempfile
from pathlib import Path

ROOT = Path(__file__).resolve().parents[2]
SUB = ROOT / "scripts" / "document" / "sub"

# str|None 外壳：8 处
STR_SITES = [("sub.outline", "lsof_check"), ("sub.styles", "lsof_check"),
             ("sub.fix_styleset", "_lsof_check"), ("sub.slim", "_lsof_check"),
             ("sub.pipeline_lib", "lsof_check"),
             ("sub.add_header_footer", "lsof_check"), ("sub.caption", "lsof_check"),
             ("sub.blocks", "lsof_check")]
# bool 外壳：2 处。fold = 把 canonical 的 str|None 折算成该外壳的期望返回值
BOOL_SITES = [("sub.caption", "lsof_check_bool", bool),
              ("pptx_cli", "_lsof_guard", lambda occ: not occ)]

BACKUP_SITES = [("sub.outline", "pick_backup_path"), ("sub.styles", "pick_backup_path"),
                ("sub.fix_styleset", "_pick_backup_path"), ("sub.slim", "_find_next_backup"),
                ("sub.pipeline_lib", "make_backup_path"), ("sub.caption", "pick_backup_path")]
NAMES = ["a.docx", "b.c.docx", "报告 v2.final.docx", "x.DOCX",
         "no-ext", ".hidden.docx", "x.tar.gz"]

# 曾经手写内联备份循环的两个文件（2026-08-02 拆掉）
INLINE_FREE = ["add_header_footer.py", "blocks.py"]


def load(dotted: str):
    """按包路径导入（`sub.fix_styleset` 等用相对导入，脱离包加载会 ImportError）。"""
    return importlib.import_module(dotted)


def naked_lsof(p) -> str | None:
    """**已删掉的旧 A 派实现的复制品**，只作第 2 条的对照组，不是被测对象。"""
    try:
        r = subprocess.run(["lsof", str(p)], capture_output=True, text=True, timeout=5)
        if r.returncode == 0 and r.stdout.strip():
            return r.stdout
    except (FileNotFoundError, subprocess.TimeoutExpired):
        pass
    return None


def main() -> int:
    sys.path.insert(0, str(ROOT / "scripts" / "document"))   # 让 `sub` 成为可导入的包
    sys.path.insert(0, str(SUB))
    sys.path.insert(0, str(ROOT / "lib"))
    import _cli_common as cc                                 # noqa: E402

    bad = 0

    # ── 3) 备份路径：6 处委派同解 ─────────────────────────────────────
    mism = 0
    with tempfile.TemporaryDirectory() as d:
        for n in NAMES:
            p = Path(d) / n
            for _ in range(3):
                ref = cc.find_next_backup(p)
                for mod, fn in BACKUP_SITES:
                    got = getattr(load(mod), fn)(p)
                    if str(got) != str(ref):
                        mism += 1
                        print(f"MISMATCH backup {mod}.{fn} {n}: {got} != {ref}")
                ref.write_bytes(b"x")
    print(f"[3] backup mismatches = {mism}   (必须 0)")
    bad += bool(mism)

    # ── 4) 两处内联备份循环确已拆掉（源码级） ──────────────────────────
    inline = 0
    for fname in INLINE_FREE:
        txt = (SUB / fname).read_text(encoding="utf-8")
        for marker in (".bak-{n}-{today}", ".bak-{m}-{today}"):
            if marker in txt:
                inline += 1
                print(f"INLINE-LOOP 残留 {fname}: {marker}")
        if "_cc.find_next_backup" not in txt:
            inline += 1
            print(f"INLINE-LOOP {fname} 没有委派 _cc.find_next_backup")
    print(f"[4] 内联备份循环残留 = {inline}   (必须 0)")
    bad += bool(inline)

    # ── 1) 10 处 lsof 逐字同解 ────────────────────────────────────────
    target = Path("/etc/hosts")
    with open(target):
        ref = cc.lsof_check(target)
        assert ref, "前置条件不成立：自己开着 fd，canonical 却判空闲"
        diverged = []
        for mod, fn in STR_SITES:
            got = getattr(load(mod), fn)(target)
            if got != ref:
                diverged.append(f"{mod}.{fn} -> {got!r}")
        for mod, fn, fold in BOOL_SITES:
            got = getattr(load(mod), fn)(target)
            want = fold(ref)
            if got != want:
                diverged.append(f"{mod}.{fn} -> {got!r} (want {want!r})")
    for line in diverged:
        print(f"DIVERGED lsof {line}")
    print(f"[1] lsof 分歧点 = {len(diverged)} / {len(STR_SITES) + len(BOOL_SITES)} 处   (必须 0)")
    bad += bool(diverged)

    # ── 2) `--` 真的生效（这次收敛唯一的真行为变更，必须钉住） ──────────
    cwd = os.getcwd()
    with tempfile.TemporaryDirectory() as d:
        dashed = Path(d) / "-hold.docx"
        dashed.write_bytes(b"x")
        rel = Path("-hold.docx")            # 相对路径才吃得到 `-` 开头
        with open(dashed):
            os.chdir(d)
            try:
                canon = cc.lsof_check(rel)
                naked = naked_lsof(rel)
            finally:
                os.chdir(cwd)
    ok_dash = bool(canon) and not naked
    print(f"[2] 前导 `-` 文件名: canonical={'OCCUPIED' if canon else 'free'} "
          f"vs 裸 lsof(旧实现)={'OCCUPIED' if naked else 'free'}   "
          f"(必须 OCCUPIED vs free — 证明 `--` 生效且旧实现真会漏判)")
    if not ok_dash:
        print("DASH-GUARD 失效：`--` 被删了，或本机 lsof 行为与实测不符")
    bad += (not ok_dash)

    print("✓ 全过" if not bad else f"⛔ {bad} 项不过")
    return 1 if bad else 0


if __name__ == "__main__":
    sys.exit(main())
