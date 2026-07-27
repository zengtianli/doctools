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

__all__ = ["find_soffice", "require_soffice", "SofficeNotFound", "CANDIDATES"]

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


if __name__ == "__main__":
    exe = find_soffice()
    print(exe or "NOT FOUND")
    raise SystemExit(0 if exe else 1)
