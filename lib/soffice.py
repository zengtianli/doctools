#!/usr/bin/env python3
"""soffice(LibreOffice) 可执行路径解析 —— 全生态 SSOT。

Why(2026-07-27 天台会话踩坑 · /govern 固化)：
  `/Applications/LibreOffice.app` 被 brew cask **升级中断**留成了空目录(0 内容)，
  真身停在 `/opt/homebrew/Caskroom/libreoffice/26.2.4.upgrading/LibreOffice.app`。
  当时 doctools 有 6 处各写各的路径(硬编码常量 / which-or-硬编码 / 候选表缺 Caskroom /
  独立 _soffice())，**全部指向那条死路径** → `/typeset`、`/docx para` 的渲染截屏门、
  老 doc 转换、图像转 SVG 全线哑火，且每个会话都要重新 mdfind 一遍才能干活。
  根因不是「LibreOffice 坏了」，是**路径解析没有 SSOT + 判据错**(见下)。

判据必须是「文件存在且可执行」，不是「.app 目录存在」：
  空壳 `.app` 目录照样 `Path(...).exists()` 为真 —— 用目录判会得到「装了」的假绿，
  然后在 subprocess 那步才炸。所以一律 isfile + os.access(X_OK) 判到 MacOS/soffice 本身。

用法：
    from soffice import find_soffice, require_soffice
    exe = find_soffice()            # 找不到返回 None
    exe = require_soffice()         # 找不到抛 SofficeNotFound(消息自带修复命令)

覆盖：环境变量 `SOFFICE_BIN=/path/to/soffice` 优先级最高(CI / 非标准安装位)。
"""
from __future__ import annotations

import os
import shutil
from glob import glob
from pathlib import Path

__all__ = ["find_soffice", "require_soffice", "SofficeNotFound", "CANDIDATES",
           "office_font_dirs", "prepare_profile", "docx_to_pdf"]

# 按优先级；带 * 的走 glob。Caskroom 一条覆盖 brew 升级中断/未 link 的情形(含 *.upgrading)。
CANDIDATES = [
    "/Applications/LibreOffice.app/Contents/MacOS/soffice",
    "~/Applications/LibreOffice.app/Contents/MacOS/soffice",
    "/opt/homebrew/bin/soffice",
    "/usr/local/bin/soffice",
    "/opt/homebrew/Caskroom/libreoffice/*/LibreOffice.app/Contents/MacOS/soffice",
    "/Applications/LibreOffice*.app/Contents/MacOS/soffice",
    "/usr/bin/soffice",
    "/usr/bin/libreoffice",
]


class SofficeNotFound(RuntimeError):
    """找不到可执行的 soffice；消息自带修复命令(fail-closed, 不静默降级)。"""


def _ok(p: str | Path) -> bool:
    p = Path(p).expanduser()
    return p.is_file() and os.access(p, os.X_OK)


def find_soffice() -> str | None:
    """返回可执行 soffice 的绝对路径；找不到返回 None。"""
    env = os.environ.get("SOFFICE_BIN")
    if env and _ok(env):
        return str(Path(env).expanduser())
    for name in ("soffice", "libreoffice"):
        w = shutil.which(name)
        if w and _ok(w):
            return w
    for pat in CANDIDATES:
        pat = str(Path(pat).expanduser())
        for hit in sorted(glob(pat), reverse=True):   # 多版本取新
            if _ok(hit):
                return hit
    return None


def require_soffice() -> str:
    """同 find_soffice，找不到直接抛错(消息含修复命令)，禁静默跳过。"""
    exe = find_soffice()
    if exe:
        return exe
    raise SofficeNotFound(
        "找不到可执行的 soffice(LibreOffice)。\n"
        "  检查过: $SOFFICE_BIN / PATH / " + " / ".join(CANDIDATES) + "\n"
        "  修复: brew reinstall --cask libreoffice\n"
        "  (若 /Applications/LibreOffice.app 是空目录 = brew 升级中断，\n"
        "   真身通常在 /opt/homebrew/Caskroom/libreoffice/*/LibreOffice.app)"
    )


# ═══════════════════════════════════════════════════════════════════════════
# Word 转 PDF（2026-10-14 立）
#
# 无界面的 soffice 看不到 Word/WPS 包里自带的宋体、黑体，中文会被换成别的字体甚至不显示。
# 办法：给 soffice 一个独立的用户配置目录，把这些字体放进它的 user/fonts（soffice 各平台
# 都会加载这个目录）。不往系统字体目录装任何东西，也不碰本人日常用的 LibreOffice 配置。
# 分页由 LibreOffice 排，可能与 Word 差一两行；要与 Word 逐页一致仍以 Word 另存为准。
# ═══════════════════════════════════════════════════════════════════════════

#: 自带中文字体的办公软件目录。存在哪个用哪个。
OFFICE_FONT_DIRS = [
    "/Applications/Microsoft Word.app/Contents/Resources/DFonts",
    "/Applications/wpsoffice.app/Contents/Resources/office6/fonts",
    "/Applications/wpsoffice.app/Contents/Frameworks/office6/fonts",
]
#: 报告用得到的字体文件名（小写比较）。只取这些，不把几百兆字体全搬进去。
_REPORT_FONT_FILES = ("simsun", "simhei", "kaiti", "fangsong", "times", "arial", "dengxian", "deng", "msyh",
                      "方正小标宋", "fzxiaobiaosong")

_FONT_ALIAS = {"宋体": "simsun", "新宋体": "nsimsun", "黑体": "simhei", "楷体": "kaiti", "仿宋": "fangsong",
               "等线": "dengxian", "微软雅黑": "microsoftyahei", "timesnewroman": "timesnewroman",
               "仿宋gb2312": "fangsong", "楷体gb2312": "kaiti",
               "方正小标宋简体": "fzxbsjwgb10", "fzxiaobiaosongb05s": "fzxbsjwgb10"}
_LEGACY_FONT_PAIRS = {"仿宋_GB2312": "仿宋", "楷体_GB2312": "楷体"}


def office_font_dirs() -> list[str]:
    # Isolated conversions need the user's installed report fonts too.
    locations = [*OFFICE_FONT_DIRS, str(Path.home() / 'Library/Fonts'), '/Library/Fonts']
    return [d for d in dict.fromkeys(locations) if Path(d).is_dir()]


def prepare_profile(profile: str | Path | None = None) -> Path:
    """建好带办公字体的 soffice 独立配置目录并返回。重复调用只补缺的字体。"""
    profile = Path(profile or Path.home() / "Library/Caches/doctools/soffice-profile").expanduser()
    fonts = profile / "user" / "fonts"
    fonts.mkdir(parents=True, exist_ok=True)
    for d in office_font_dirs():
        for f in sorted(Path(d).iterdir()):
            if f.suffix.lower() not in (".ttf", ".ttc", ".otf"):
                continue
            if not f.name.lower().startswith(_REPORT_FONT_FILES):
                continue
            dst = fonts / f.name
            if not dst.exists() or dst.stat().st_size != f.stat().st_size:
                # Font bytes matter here; macOS protected file flags must not
                # be transplanted onto the disposable conversion profile.
                shutil.copyfile(f, dst)
    # Official LibreOffice Common/Font/Substitution schema; changes only this
    # conversion profile. Historical Word GB2312 names map to installed faces.
    import xml.etree.ElementTree as ET
    oor = "http://openoffice.org/2001/registry"
    settings = profile / "user/registrymodifications.xcu"
    root = ET.parse(settings).getroot() if settings.exists() else ET.Element(
        "{" + oor + "}items", {"xmlns:xs":"http://www.w3.org/2001/XMLSchema",
                                "xmlns:xsi":"http://www.w3.org/2001/XMLSchema-instance"})
    path = "/org.openoffice.Office.Common/Font/Substitution"
    # Keep all unrelated preferences; replace only our two known compatibility entries.
    for item in list(root):
        item_path = item.get("{" + oor + "}path")
        if item_path == path:
            for prop in list(item):
                if prop.get("{" + oor + "}name") == "Replacement": item.remove(prop)
        elif item_path == path + "/FontPairs":
            for node in list(item):
                if str(node.get("{" + oor + "}name", "")).startswith("doctools-legacy-"): item.remove(node)
        if item_path in {path,path + "/FontPairs"} and not len(item): root.remove(item)
    item = ET.SubElement(root, "item", {"{" + oor + "}path":path})
    prop = ET.SubElement(item, "prop", {"{" + oor + "}name":"Replacement", "{" + oor + "}op":"fuse"})
    ET.SubElement(prop, "value").text = "true"
    item = ET.SubElement(root, "item", {"{" + oor + "}path":path + "/FontPairs"})
    for i, (requested, installed) in enumerate(_LEGACY_FONT_PAIRS.items()):
        node = ET.SubElement(item, "node", {"{" + oor + "}name":"doctools-legacy-" + str(i), "{" + oor + "}op":"replace"})
        for key, value in {"ReplaceFont":requested,"SubstituteFont":installed,"Always":"true","OnScreenOnly":"false"}.items():
            prop = ET.SubElement(node,"prop",{"{" + oor + "}name":key})
            ET.SubElement(prop,"value").text=value
    ET.register_namespace("oor", oor)
    ET.ElementTree(root).write(settings,encoding="utf-8",xml_declaration=True)
    return profile


def _norm_font(name: str) -> str:
    import re
    name = re.sub(r"^[A-Z]{6}\+", "", name or "")
    low = re.sub(r"[\s\-_]", "", name).casefold()
    low = _FONT_ALIAS.get(low, low)
    return re.sub(r"(psmt|ps|mt|bold|italic|regular|bolditalic)+$", "", low) or low


def _docx_declared_fonts(docx: Path) -> set[str]:
    import re
    import zipfile
    names: set[str] = set()
    with zipfile.ZipFile(docx) as z:
        for part in z.namelist():
            if part.startswith("word/") and part.endswith(".xml"):
                xml = z.read(part).decode("utf-8", "ignore")
                names.update(re.findall(r'w:(?:ascii|hAnsi|eastAsia|cs)="([^"]+)"', xml))
    return {_norm_font(n) for n in names}


def docx_to_pdf(src: str | Path, out_pdf: str | Path, *, profile: str | Path | None = None,
                timeout: int = 300) -> dict:
    """Word 转 PDF，字体取办公软件自带的宋体、黑体等。返回 {pdf, pages, fonts, unexpected_fonts}。

    unexpected_fonts 非空 = PDF 里出现了 Word 稿没有声明的字体，也就是被替换了；调用方据此
    决定要不要把这份 PDF 交出去。目标已存在时报错，不覆盖。
    """
    import subprocess
    import tempfile
    src, out_pdf = Path(src).expanduser().resolve(), Path(out_pdf).expanduser().resolve()
    if not src.is_file():
        raise FileNotFoundError(src)
    if out_pdf.exists():
        raise FileExistsError(f"目标已存在，不覆盖：{out_pdf}")
    exe = require_soffice()
    prof = prepare_profile(profile)
    with tempfile.TemporaryDirectory() as tmp:
        work = Path(tmp) / "in.docx"            # 中文长文件名在 soffice 命令行上偶有问题，统一用短名
        shutil.copy2(src, work)
        run = subprocess.run(
            [exe, f"-env:UserInstallation={prof.as_uri()}", "--headless", "--convert-to", "pdf",
             "--outdir", tmp, str(work)], capture_output=True, text=True, timeout=timeout)
        made = Path(tmp) / "in.pdf"
        if run.returncode != 0 or not made.is_file():
            raise RuntimeError(f"soffice 转 PDF 失败（退出码 {run.returncode}）：{(run.stderr or run.stdout)[-400:]}")
        out_pdf.parent.mkdir(parents=True, exist_ok=True)
        shutil.move(str(made), out_pdf)
    info: dict = {"pdf": str(out_pdf), "pages": None, "fonts": [], "unexpected_fonts": [],
                  "font_dirs": office_font_dirs(), "font_aliases": dict(_LEGACY_FONT_PAIRS), "engine": "LibreOffice"}
    if shutil.which("pdffonts"):
        rows = subprocess.run(["pdffonts", str(out_pdf)], capture_output=True, text=True).stdout.splitlines()[2:]
        fonts = sorted({r.split()[0] for r in rows if r.strip()})
        declared = _docx_declared_fonts(src)
        info["fonts"] = fonts
        info["unexpected_fonts"] = [f for f in fonts if _norm_font(f) not in declared]
    else:
        info["fonts_unchecked"] = "本机没有 pdffonts，未核字体"
    if shutil.which("pdfinfo"):
        for line in subprocess.run(["pdfinfo", str(out_pdf)], capture_output=True, text=True).stdout.splitlines():
            if line.startswith("Pages:"):
                info["pages"] = int(line.split()[1])
    return info


if __name__ == "__main__":
    import sys
    if len(sys.argv) >= 3 and sys.argv[1] == "to-pdf":
        import json
        out = sys.argv[sys.argv.index("-o") + 1] if "-o" in sys.argv else str(Path(sys.argv[2]).with_suffix(".pdf"))
        result = docx_to_pdf(sys.argv[2], out)
        print(json.dumps(result, ensure_ascii=False, indent=2))
        raise SystemExit(1 if result["unexpected_fonts"] else 0)
    exe = find_soffice()
    print(exe or "NOT FOUND")
    raise SystemExit(0 if exe else 1)
