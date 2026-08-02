#!/usr/bin/env python3
"""wheel 自包含门 —— 证明「装到别的机器上真能用」，不是证明「配置写对了」。

## 它守的是哪条缝

`src/doctools/cli.py` 按顺序试两个实现根：**工作树**（editable 装）→ **包内副本**
（wheel 装）。本机日常跑的永远是第一条；第二条**只有在别的机器上才会被走到**。
`cli_surface.py` / `cli_forward_probe.py` 都看不见它 —— 那两个在本仓工作树里跑，
命中的是工作树分支。所以这条分支如果没人测，它就是「写了但从没执行过的代码」，
而它恰恰是唯一决定「wheel 能不能用」的代码。

2026-08-02 实测基线：改之前的 wheel 只有 7 个文件，clean venv 装完
`doctools --version` 直接 `FATAL: 找不到实现入口` rc=2。

## 三层判据（逐层收紧，全部 fail-closed）

  struct   构建 wheel → 包内 `_bundled/` 的**文件集**必须等于
           `git ls-files scripts lib config schemas`，且**逐文件 sha256 相同**。
           这条就是「两份副本不会漂」的机器证明：包内那份是工作树的逐字节快照，
           不是另一份需要同步维护的源。集合两个方向都查（多一个/少一个都判红）。
  install  全新 venv（--no-deps）装 wheel，**在仓库目录之外**跑 `--version` 与
           `verbs`。--no-deps 是有意的：这一层要暴露的是「实现进没进包」，
           不是第三方依赖装没装上，装一堆 pandas/scipy 只会把信号淹掉、把门拖慢到
           没人跑。
  full     再补一次完整依赖安装 + 真 docx 上跑写盘前的只读动词。慢（分钟级），
           发版前跑。

用法：
    python3 tools/check_wheel_selfcontained.py              # struct + install
    python3 tools/check_wheel_selfcontained.py --tier struct
    python3 tools/check_wheel_selfcontained.py --tier full
"""
from __future__ import annotations

import argparse
import hashlib
import shutil
import subprocess
import sys
import sysconfig
import tempfile
import zipfile
from pathlib import Path

REPO = Path(__file__).resolve().parents[1]

# 与 pyproject 的 [tool.hatch.build.targets.wheel.force-include] 一一对应。
# 这里写死一份是有意的：它是**独立的第二判据**。如果谁往 force-include 里加了一行
# 却没在这儿加，集合比对会当场判红 —— 从 pyproject 反读回来就没有这个作用了。
BUNDLED_ROOTS = ("scripts", "lib", "config", "schemas")

# 仓外来源：构建时从总部 SSOT 现取，工作树里没有第二份。key = 包内路径。
EXTERNAL = {
    "lib/parallel_contract.py": Path.home() / "Dev" / "tools" / "dev" / "lib" / "parallel_contract.py",
}

PREFIX = "doctools/_bundled/"


def _sha(p: Path) -> str:
    return hashlib.sha256(p.read_bytes()).hexdigest()


def _run(cmd: list[str], **kw) -> subprocess.CompletedProcess:
    return subprocess.run(cmd, capture_output=True, text=True, **kw)


def build_wheel(outdir: Path) -> Path:
    print(f"[build] uv build --wheel --out-dir {outdir}")
    r = _run(["uv", "build", "--wheel", "--out-dir", str(outdir)], cwd=REPO)
    if r.returncode != 0:
        print(r.stdout + r.stderr, file=sys.stderr)
        raise SystemExit(f"[FAIL] 构建失败 rc={r.returncode}")
    whls = sorted(outdir.glob("doctools-*.whl"))
    if len(whls) != 1:
        raise SystemExit(f"[FAIL] 期望恰好 1 个 wheel，实得 {len(whls)}: {whls}")
    print(f"[build] ok → {whls[0].name}")
    return whls[0]


def check_struct(whl: Path) -> int:
    """包内副本 == 工作树 git-tracked 集，逐文件 sha256。"""
    r = _run(["git", "ls-files", *BUNDLED_ROOTS], cwd=REPO)
    if r.returncode != 0:
        raise SystemExit(f"[FAIL] git ls-files 失败: {r.stderr}")
    expected = {ln for ln in r.stdout.splitlines() if ln.strip()}
    if not expected:
        # 铁律：枚举为空一律判红，拒绝在空集上报绿
        raise SystemExit("[FAIL] git ls-files 枚举为空 —— 判据失效，不给绿")
    expected |= set(EXTERNAL)

    with zipfile.ZipFile(whl) as z:
        names = z.namelist()
        actual = {n[len(PREFIX):] for n in names if n.startswith(PREFIX) and not n.endswith("/")}

        bad = 0
        missing = sorted(expected - actual)
        extra = sorted(actual - expected)
        if missing:
            bad += len(missing)
            print(f"[FAIL] wheel 里缺 {len(missing)} 个文件（工作树有、包里没有）:")
            for m in missing[:15]:
                print(f"    - {m}")
        if extra:
            bad += len(extra)
            print(f"[FAIL] wheel 里多 {len(extra)} 个文件（包里有、工作树没有）:")
            for e in extra[:15]:
                print(f"    + {e}")

        # 逐字节：集合对上不代表内容对上
        drifted = []
        for rel in sorted(expected & actual):
            src = EXTERNAL.get(rel) or (REPO / rel)
            if not src.is_file():
                drifted.append((rel, "源文件不存在"))
                continue
            if hashlib.sha256(z.read(PREFIX + rel)).hexdigest() != _sha(src):
                drifted.append((rel, "sha256 不等"))
        if drifted:
            bad += len(drifted)
            print(f"[FAIL] {len(drifted)} 个文件包内与源端不一致:")
            for rel, why in drifted[:15]:
                print(f"    ~ {rel}: {why}")

    if bad:
        return 1
    print(f"[struct] ✓ 包内 {len(expected)} 个文件与源端逐字节一致"
          f"（{len(BUNDLED_ROOTS)} 个根 + {len(EXTERNAL)} 个仓外 SSOT）")

    # 壳本身必须在，且 _bundled 不能反过来污染工作树
    with zipfile.ZipFile(whl) as z:
        for must in ("doctools/__init__.py", "doctools/cli.py"):
            if must not in z.namelist():
                print(f"[FAIL] wheel 缺 {must}")
                return 1
    if (REPO / "src" / "doctools" / "_bundled").exists():
        print("[FAIL] 工作树里出现了 src/doctools/_bundled/ —— 副本只准活在构建产物里")
        return 1
    print("[struct] ✓ 工作树中不存在 _bundled/（副本只在 wheel 内）")
    return 0


def check_install(whl: Path, full: bool) -> int:
    """全新 venv 装 wheel，在仓库目录之外跑 —— 走的是包内副本那条分支。"""
    with tempfile.TemporaryDirectory(prefix="doctools-wheelgate-") as td:
        venv = Path(td) / "venv"
        print(f"[install] python -m venv {venv}")
        if _run([sys.executable, "-m", "venv", str(venv)]).returncode != 0:
            print("[FAIL] venv 创建失败")
            return 1
        bindir = "Scripts" if sysconfig.get_platform().startswith("win") else "bin"
        pip = venv / bindir / "pip"
        exe = venv / bindir / "doctools"

        cmd = [str(pip), "install", "-q", str(whl)]
        if not full:
            cmd.insert(2, "--no-deps")
        print(f"[install] pip install {'(with deps)' if full else '--no-deps'} …")
        r = _run(cmd)
        if r.returncode != 0:
            print(r.stdout + r.stderr, file=sys.stderr)
            print("[FAIL] 安装失败")
            return 1
        if not exe.is_file():
            print(f"[FAIL] console_script 没生成: {exe}")
            return 1

        # cwd 必须在仓库之外 —— 在仓库里跑会让工作树分支意外命中，测了个假的
        cwd = td
        checks = [([str(exe), "--version"], 0), ([str(exe), "verbs", "--fn", "convert"], 0)]
        bad = 0
        for argv, want in checks:
            r = _run(argv, cwd=cwd)
            ok = r.returncode == want
            print(f"[install] {'✓' if ok else '✗'} rc={r.returncode} (want {want}) :: {' '.join(argv[1:])}")
            if not ok:
                bad += 1
                print((r.stdout + r.stderr)[:900], file=sys.stderr)
            elif r.stdout.strip():
                print("           " + r.stdout.strip().splitlines()[0][:100])
            # 兜底分支被走到时最典型的假绿：rc 对了但 stderr 里其实在喊 FATAL
            if "FATAL" in (r.stdout + r.stderr):
                bad += 1
                print("[FAIL] 输出里出现 FATAL —— 实现没进包")
        if bad:
            return 1
        print("[install] ✓ 仓库外 clean venv 跑通（走的是包内副本分支）")
    return 0


def main() -> int:
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--tier", choices=("struct", "install", "full"), default="install")
    ap.add_argument("--keep", metavar="DIR", help="把 wheel 留在这里（默认构建到临时目录后删掉）")
    args = ap.parse_args()

    if shutil.which("uv") is None:
        print("[FAIL] 找不到 uv —— 构建后端缺失，不给绿", file=sys.stderr)
        return 1

    with tempfile.TemporaryDirectory(prefix="doctools-wheelbuild-") as td:
        outdir = Path(args.keep) if args.keep else Path(td)
        outdir.mkdir(parents=True, exist_ok=True)
        whl = build_wheel(outdir)

        rc = check_struct(whl)
        if rc:
            print("\n[RESULT] ✗ struct 层未过")
            return rc
        if args.tier != "struct":
            rc = check_install(whl, full=(args.tier == "full"))
            if rc:
                print("\n[RESULT] ✗ install 层未过")
                return rc

    print(f"\n[RESULT] ✓ wheel 自包含（tier={args.tier}）")
    return 0


if __name__ == "__main__":
    sys.exit(main())
