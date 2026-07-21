#!/usr/bin/env python3
"""docx_font_normalize.py — 去「等线」：把 docx 的中文字体归一到 宋体/仿宋 体系。

用户钦定（2026-07-20 /govern）：**我的所有 docx 文档都不用等线字体，要么宋体 要么仿宋**。

等线是怎么混进来的（三条路，缺一不可全堵）：
  ① docDefaults 用 `w:eastAsiaTheme="minorEastAsia"`，而 theme1.xml 的 `<a:ea typeface=""/>` 为空
     → Word 回退到 UI 默认字体 **等线**。母版看着没问题，脚本新加的 run 全中招。
  ② run/style 里显式写了 等线 / DengXian / Deng Xian。
  ③ 脚本（python-docx 等）新建 run 不设 rFonts → 继承 ①。

用法:
    python3 docx_font_normalize.py <docx...> --check      # 只报告, 退出码 1 = 有等线风险
    python3 docx_font_normalize.py <docx...> --apply      # 原地修复 + .bak-时间戳
    --font 仿宋_GB2312 (默认) | 宋体 | 仿宋       中文正文字体
    --ascii "Times New Roman" (默认)              西文字体

修复动作(--apply):
  · docDefaults: eastAsiaTheme → 显式 w:eastAsia="<font>"; asciiTheme/hAnsiTheme → 显式西文
  · theme1.xml : major/minor 的 <a:ea typeface=""> 补成 <font>（兜底, 防别处再回退）
  · 全 xml part: 显式 等线/DengXian/Deng Xian → <font>
  · 无 rFonts 的 run 不动（此时已由 docDefaults 兜到 <font>, 不必逐 run 塞）
"""
from __future__ import annotations

import argparse
import datetime as _dt
import re
import shutil
import sys
import zipfile
from pathlib import Path

DENGXIAN = re.compile(r"等线|DengXian|Deng Xian")
THEME_EA_RE = re.compile(r'(<a:ea\s+typeface=")([^"]*)(")')
DOCDEFAULT_RFONTS = re.compile(r"(<w:docDefaults>.*?<w:rPrDefault>.*?<w:rPr>.*?)(<w:rFonts[^/>]*/>)", re.S)


def _parts(zf: zipfile.ZipFile) -> list[str]:
    return [n for n in zf.namelist() if n.endswith(".xml")]


def scan(path: Path) -> list[str]:
    """返回问题列表, 空 = 干净。"""
    issues: list[str] = []
    with zipfile.ZipFile(path) as z:
        names = _parts(z)
        for n in names:
            text = z.read(n).decode("utf8", errors="ignore")
            if DENGXIAN.search(text):
                issues.append(f"{n}: 显式等线/DengXian")
        try:
            styles = z.read("word/styles.xml").decode("utf8", errors="ignore")
        except KeyError:
            styles = ""
        theme_ea = []
        for n in names:
            if n.startswith("word/theme/"):
                theme_ea += THEME_EA_RE.findall(z.read(n).decode("utf8", errors="ignore"))
        if "eastAsiaTheme" in styles and any(t[1].strip() == "" for t in theme_ea):
            issues.append("word/styles.xml: docDefaults 走 minorEastAsia 主题，而 theme 的 <a:ea> 为空 → Word 回退等线")
    return issues


def fix(path: Path, font: str, ascii_font: str) -> list[str]:
    changed: list[str] = []
    ts = _dt.datetime.now().strftime("%Y%m%d-%H%M%S")
    shutil.copy(path, f"{path}.bak-{ts}")
    with zipfile.ZipFile(path) as zin:
        items = zin.infolist()
        data = {i.filename: zin.read(i.filename) for i in items}

    for name in list(data):
        if not name.endswith(".xml"):
            continue
        text = data[name].decode("utf8", errors="ignore")
        orig = text

        if DENGXIAN.search(text):
            text = DENGXIAN.sub(font, text)
            changed.append(f"{name}: 显式等线 → {font}")

        if name.startswith("word/theme/"):
            def _ea(m: re.Match[str]) -> str:
                return m.group(1) + (m.group(2) or font if m.group(2).strip() else font) + m.group(3)
            new = THEME_EA_RE.sub(_ea, text)
            if new != text:
                text, _ = new, changed.append(f"{name}: theme <a:ea> 空 → {font}")

        if name == "word/styles.xml":
            def _rf(m: re.Match[str]) -> str:
                return (
                    m.group(1)
                    + f'<w:rFonts w:ascii="{ascii_font}" w:eastAsia="{font}" '
                      f'w:hAnsi="{ascii_font}" w:cs="{ascii_font}"/>'
                )
            new = DOCDEFAULT_RFONTS.sub(_rf, text, count=1)
            if new != text:
                text = new
                changed.append(f"{name}: docDefaults 主题字体 → 显式 {font} / {ascii_font}")

        if text != orig:
            data[name] = text.encode("utf8")

    if changed:
        with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as zout:
            for i in items:
                zout.writestr(i, data[i.filename])
    else:
        Path(f"{path}.bak-{ts}").unlink(missing_ok=True)
    return changed


def main() -> int:
    ap = argparse.ArgumentParser(description="docx 去等线（中文字体归一到 宋体/仿宋）")
    ap.add_argument("docx", nargs="+")
    ap.add_argument("--check", action="store_true", help="只报告不改（默认）")
    ap.add_argument("--apply", action="store_true", help="原地修复 + .bak-时间戳")
    ap.add_argument("--font", default="仿宋_GB2312", help="中文字体（默认 仿宋_GB2312）")
    ap.add_argument("--ascii", dest="ascii_font", default="Times New Roman")
    args = ap.parse_args()

    rc = 0
    for p in args.docx:
        path = Path(p)
        if path.name.startswith("~$"):
            continue
        if not path.exists():
            print(f"[跳过] 不存在: {path}", file=sys.stderr)
            continue
        issues = scan(path)
        if not issues:
            print(f"✅ {path.name}: 无等线风险")
            continue
        rc = 1
        print(f"⚠️  {path.name}")
        for i in issues:
            print(f"    · {i}")
        if args.apply:
            for c in fix(path, args.font, args.ascii_font):
                print(f"    ✔ {c}")
            left = scan(path)
            print("    ✅ 复检干净" if not left else f"    ❌ 仍有: {left}")
            rc = 0 if not left else 1
    return rc


if __name__ == "__main__":
    sys.exit(main())
