#!/usr/bin/env python3
"""
剪贴板操作模块
"""

import shutil
import subprocess
from pathlib import Path


def copy_to_clipboard(text: str):
    """复制文本到剪贴板"""
    process = subprocess.Popen(['pbcopy'], stdin=subprocess.PIPE)
    process.communicate(text.encode('utf-8'))


def get_from_clipboard() -> str:
    """从剪贴板获取文本"""
    result = subprocess.run(['pbpaste'], capture_output=True, text=True)
    return result.stdout


def get_clipboard_files() -> list[Path]:
    """获取剪贴板中的文件路径"""
    script = '''
    tell application "System Events"
        try
            set theFiles to the clipboard as «class furl»
            set output to ""
            repeat with f in theFiles
                set output to output & (POSIX path of (f as text)) & linefeed
            end repeat
            return output
        on error
            try
                return POSIX path of (the clipboard as «class furl»)
            on error
                return ""
            end try
        end try
    end tell
    '''
    result = subprocess.run(['osascript', '-e', script], capture_output=True, text=True)
    return [Path(p) for p in result.stdout.strip().split('\n') if p and Path(p).exists()]


def paste_files(target_dir: str | Path) -> list[Path]:
    """粘贴剪贴板文件到目标目录"""
    files = get_clipboard_files()
    pasted = []
    target = Path(target_dir)

    for src in files:
        dst = target / src.name
        if dst.exists():
            base, ext = dst.stem, dst.suffix
            counter = 1
            while dst.exists():
                dst = target / f"{base}_{counter}{ext}"
                counter += 1
        try:
            if src.is_dir():
                shutil.copytree(src, dst)
            else:
                shutil.copy2(src, dst)
            pasted.append(dst)
        except Exception:
            pass
    return pasted
