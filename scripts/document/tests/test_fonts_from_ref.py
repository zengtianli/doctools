#!/usr/bin/env python3
"""反向验证 docx_fmt.py fonts 的改动：把原 bug 放回去，确认新行为不再复现。

原 bug（2026-08-06 实证）：`docx_fmt.py fonts <金华golden> --apply` 因为 --font 默认值是
仿宋_GB2312，把金华的宋体外壳 theme/docDefaults/fontTable 三层整层改成仿宋。
"""
import shutil, subprocess, sys, tempfile, os
from pathlib import Path

# 本文件是**脚本形态**的验证器：整个文件在模块层直接执行，末尾裸 sys.exit()。
# pytest 收集时会 import 它 → SystemExit 打到 collection 阶段 → INTERNALERROR，
# **整仓 pytest 一条测试都跑不起来**（2026-08-14 实测 rc=3）。CLAUDE.md 把
# `python3 -m pytest -q` 列为闸门，而这个闸门当时是哑的。
# 直接跑（python3 本文件）行为不变；被 import 时干净跳过，跳过理由可见不静默。
if __name__ != "__main__":
    import pytest as _pytest

    _pytest.skip(
        "脚本形态验证器：模块层直接执行 + 裸 sys.exit()，且依赖 ~/Work 下的真实"
        "docx（金华 golden / 吴兴 broken），不是确定性单测。直接跑："
        "python3 scripts/document/tests/test_fonts_from_ref.py",
        allow_module_level=True,
    )

FMT = "/Users/tianli/Dev/tools/doctools/scripts/document/docx_fmt.py"
GOLDEN = ("/Users/tianli/Work/projects/reclaim/01-源/参考资料/"
          "金华江流域生态流量分类管控保障项目-水利厅验收0805/"
          "1项目任务完成情况/2项目绩效自评报告/项目绩效自评报告.docx")
BROKEN = ("/Users/tianli/Work/projects/reclaim/02-工作区/试点验收材料/"
          "00吴兴区工业集聚区水资源循环梯级利用试点项目-2025年自评说明.docx")

sys.path.insert(0, os.path.dirname(FMT))
os.environ.setdefault("PYTHONPATH", "/Users/tianli/Dev/tools/doctools/lib:/Users/tianli/Dev/tools/dev/lib")


def run(args):
    r = subprocess.run([sys.executable, FMT, "fonts"] + args, capture_output=True, text=True,
                       env={**os.environ,
                            "PYTHONPATH": "/Users/tianli/Dev/tools/doctools/lib:/Users/tianli/Dev/tools/dev/lib"})
    out = "\n".join(l for l in (r.stdout + r.stderr).splitlines() if "pkg_resources" not in l)
    return r.returncode, out


_DFMT = None


def ea_mode(p):
    """必须调 docx_fmt 自己的 effective_eastasia，禁自己重写一遍去"实测"（铁律 #2）。"""
    global _DFMT
    if _DFMT is None:
        for d in ("/Users/tianli/Dev/tools/doctools/lib", "/Users/tianli/Dev/tools/dev/lib",
                  os.path.dirname(FMT)):
            if d not in sys.path:
                sys.path.insert(0, d)
        import importlib
        _DFMT = importlib.import_module("docx_fmt")
    return _DFMT.effective_eastasia(Path(p))


tmp = tempfile.mkdtemp(prefix="fontfix-")
ok = True

print("=== 0. 解析器自检：金华 golden 应解析出宋体，吴兴事故遗留应解析出仿宋 ===")
g = ea_mode(GOLDEN)
print(f"  金华 golden 有效中文字体: {g[:3]}")
ok &= bool(g) and g[0][0] == "宋体"
print(f"  {'✅' if ok else '❌'} 金华众数 = 宋体")

print("\n=== 1. 原 bug 重放：对金华 golden 跑 --apply 且不给 --font ===")
shell = os.path.join(tmp, "golden.docx")
shutil.copy2(GOLDEN, shell)
rc, out = run([shell, "--apply"])
print("  " + "\n  ".join(out.splitlines()[:8]))
after = ea_mode(shell)
print(f"  改后有效中文字体: {after[:3]}")
bad = any(f == "仿宋_GB2312" for f, _ in after)
print(f"  {'❌ 仍被改成仿宋（未修复）' if bad else '✅ 没有被改成仿宋'}")
ok &= not bad

print("\n=== 2. 复刻 golden 的正路：--from-ref 应把目标定成宋体 ===")
t2 = os.path.join(tmp, "derived.docx")
shutil.copy2(BROKEN, t2)
rc, out = run([t2, "--apply", "--from-ref", GOLDEN])
has = "有效中文字体众数 = 宋体" in out
print("  " + "\n  ".join(out.splitlines()[:6]))
print(f"  {'✅' if has else '❌'} --from-ref 正确推出宋体")
ok &= has

print("\n=== 3. fail-closed：判不出字体时必须拒跑，不许拿默认值覆盖 ===")
import zipfile
from xml.etree import ElementTree as ET
NS = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
t3 = os.path.join(tmp, "nofont.docx")
shutil.copy2(GOLDEN, t3)
# 造一个"中文字体全是等线"的样例：把所有 eastAsia 值换成等线
with zipfile.ZipFile(t3) as z:
    blobs = {n: z.read(n) for n in z.namelist()}
for n in ("word/document.xml", "word/styles.xml"):
    if n in blobs:
        s = blobs[n].decode("utf8")
        s = s.replace('w:eastAsia="宋体"', 'w:eastAsia="等线"').replace('w:eastAsia="黑体"', 'w:eastAsia="等线"')
        blobs[n] = s.encode("utf8")
with zipfile.ZipFile(t3, "w", zipfile.ZIP_DEFLATED) as z:
    for n, b in blobs.items():
        z.writestr(n, b)
rc, out = run([t3, "--apply"])
refused = "判不出目标中文字体" in out and rc != 0
print("  " + "\n  ".join(out.splitlines()[-3:]))
print(f"  rc={rc}  {'✅ 拒跑' if refused else '❌ 没拒跑（fail-open）'}")
ok &= refused

print("\n=== 4. 主流用法不回归：清等线时沿用本文档自身字体 ===")
t4 = os.path.join(tmp, "own.docx")
shutil.copy2(BROKEN, t4)
rc, out = run([t4, "--apply"])
kept = "本文档自身众数" in out or "无等线风险" in out
print("  " + "\n  ".join(out.splitlines()[:5]))
print(f"  {'✅ 沿用自身字体（不需要用户显式给）' if kept else '❌ 行为异常'}")
ok &= kept

shutil.rmtree(tmp, ignore_errors=True)
print("\n" + ("✅ 四项全过" if ok else "❌ 有未过项"))
sys.exit(0 if ok else 1)
