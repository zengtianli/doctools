#!/usr/bin/env python3
"""typeset_pipeline — 排版轴一条龙 driver（/typeset all 的后端）。

把高价值排版步串成 checklist pipeline，每步 snapshot→run→完整性自检→
保留/回滚→记录，输出对账卡。**韧性编排**：任一步报错或把 docx 跑坏（unzip -t
失败）自动回滚该步、标 ⚠ 跳过，不连累全局。

步序（从 0624 vs 0625 真实 diff 推出，高价值优先）：
  ① styleset  样式池清理（去 TOC10 等重复后缀）      docx_cli styleset restore
  ② spacing   正文固定行距（对参照, best-effort）     line_spacing --fix --ref
  ③ figs      图号节内重排+补号+居中                  renumber-fig --cn-section --fix-center
  ④ chrome    逐章页眉/水印/分节/横表（对参照复刻）    docx_cli chrome --raw --template
  ⑤ gate      交付不变量自检（read-only 收尾）         docx_cli health gate

surgical 安全网：每步后 unzip -t 验完整性，坏了回滚 → 全程不产出损坏 docx。

Usage:
  typeset_pipeline.py <docx> --ref <golden>.docx [--county 天台县] [--out OUT] [--inplace]
"""
from __future__ import annotations

import argparse
import shutil
import subprocess
import sys
import zipfile
from datetime import datetime
from pathlib import Path

HERE = Path(__file__).resolve().parent
DC = HERE / "docx_cli.py"
LINE_SPACING = HERE / "sub" / "line_spacing.py"


def _intact(p: Path) -> bool:
    try:
        with zipfile.ZipFile(p) as z:
            return z.testzip() is None
    except Exception:
        return False


def _run(cmd, timeout=300):
    try:
        r = subprocess.run(cmd, capture_output=True, text=True, timeout=timeout)
        out = (r.stdout or "") + (r.stderr or "")
        out = "\n".join(l for l in out.splitlines() if "pkg_resources" not in l)
        return r.returncode, out
    except subprocess.TimeoutExpired:
        return 124, "TIMEOUT"
    except Exception as e:  # noqa
        return 1, f"EXC {e}"


def main(argv=None):
    ap = argparse.ArgumentParser(description="排版 pipeline driver（/typeset all）")
    ap.add_argument("docx", type=Path)
    ap.add_argument("--ref", type=Path, help="golden 参照件（chrome/spacing 对照）")
    ap.add_argument("--county", default=None, help="县名（chrome 页眉 swap）")
    ap.add_argument("--out", type=Path, default=None)
    ap.add_argument("--inplace", action="store_true")
    a = ap.parse_args(argv)

    if not a.docx.exists():
        print(f"找不到: {a.docx}", file=sys.stderr)
        return 1

    # 工作副本（绝不动原件，除非 --inplace）
    if a.inplace:
        work = a.docx
        bak = a.docx.with_suffix(a.docx.suffix + f".bak-{datetime.now():%Y%m%d-%H%M%S}")
        shutil.copy2(a.docx, bak)
        print(f"备份原件 → {bak.name}")
    else:
        work = a.out or a.docx.with_name(a.docx.stem + "_typeset.docx")
        shutil.copy2(a.docx, work)

    steps = []
    steps.append(("① styleset 样式池清理",
                  [sys.executable, str(DC), "styleset", "restore", str(work)]))
    if a.ref:
        steps.append(("② spacing 固定行距(对参照)",
                      [sys.executable, str(LINE_SPACING), str(work),
                       "--fix", "--ref", str(a.ref), "--no-backup"]))
    steps.append(("③ figs 图号重排+居中",
                  [sys.executable, str(DC), "renumber-fig", "--cn-section",
                   "--fix-center", "--inplace", str(work)]))

    report = []
    for name, cmd in steps:
        snap = work.with_suffix(work.suffix + ".snap")
        shutil.copy2(work, snap)
        rc, out = _run(cmd)
        if rc != 0 or not _intact(work):
            shutil.copy2(snap, work)  # 回滚
            report.append((name, "⚠ 跳过(回滚)", out.strip().splitlines()[-1] if out.strip() else f"rc={rc}"))
        else:
            tail = next((l for l in reversed(out.splitlines()) if l.strip()), "")
            report.append((name, "✅ 完成", tail[:60]))
        snap.unlink(missing_ok=True)

    # ④ chrome（产出 _chrome 新文件，单独处理）
    if a.ref:
        snap = work.with_suffix(work.suffix + ".snap")
        shutil.copy2(work, snap)
        ccmd = [sys.executable, str(DC), "chrome", "--raw", str(work),
                "--template", str(a.ref)]
        if a.county:
            ccmd += ["--county", a.county]
        rc, out = _run(ccmd, timeout=300)
        chrome_out = work.with_name(work.stem + "_chrome.docx")
        if rc == 0 and chrome_out.exists() and _intact(chrome_out):
            chrome_out.replace(work)
            tail = next((l for l in out.splitlines() if "节数" in l), "")
            report.append(("④ chrome 逐章页眉装帧", "✅ 完成", tail.strip()[:60]))
        else:
            shutil.copy2(snap, work)
            chrome_out.unlink(missing_ok=True)
            report.append(("④ chrome 逐章页眉装帧", "⚠ 跳过(回滚)",
                           out.strip().splitlines()[-1] if out.strip() else f"rc={rc}"))
        snap.unlink(missing_ok=True)

    # ⑤ gate（read-only 收尾自检）
    rc, out = _run([sys.executable, str(DC), "health", "gate", str(work)])
    gate_ok = rc == 0
    report.append(("⑤ gate 交付不变量自检", "✅ exit0" if gate_ok else f"❌ found(rc={rc})",
                   next((l for l in out.splitlines() if l.strip()), "")[:60]))

    # 对账卡
    print("\n" + "=" * 64)
    print(f"  排版 pipeline 对账卡 · {work.name}")
    print("=" * 64)
    for name, status, note in report:
        print(f"  {status:14} {name}")
        if note:
            print(f"                 └ {note}")
    print("=" * 64)
    print(f"  产出: {work}")
    print(f"  完整性: {'✅ unzip -t OK' if _intact(work) else '❌ 损坏'}")
    return 0 if gate_ok else 2


if __name__ == "__main__":
    sys.exit(main())
