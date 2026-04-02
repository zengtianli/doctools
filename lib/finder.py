#!/usr/bin/env python3
"""
Finder 交互模块
提供与 macOS Finder 交互的函数
"""

import os
import subprocess
from pathlib import Path

from display import show_error, show_warning


def get_finder_selection() -> list[str]:
    """
    获取 Finder 选中的文件列表

    Returns:
        文件路径列表，如果没有选中返回空列表
    """
    script = '''
    tell application "Finder"
        set selectedItems to selection as list
        if (count of selectedItems) = 0 then
            return ""
        end if
        set posixPaths to {}
        repeat with i from 1 to count of selectedItems
            set end of posixPaths to POSIX path of (item i of selectedItems as alias)
        end repeat
        set AppleScript's text item delimiters to "\\n"
        set pathsText to posixPaths as text
        set AppleScript's text item delimiters to ""
        return pathsText
    end tell
    '''
    try:
        result = subprocess.run(
            ['osascript', '-e', script],
            capture_output=True,
            text=True,
            timeout=5
        )
        paths = result.stdout.strip()
        if not paths:
            return []
        return [p for p in paths.split('\n') if p]
    except Exception:
        return []


def get_finder_selection_single() -> str | None:
    """
    获取 Finder 选中的单个文件

    Returns:
        文件路径，如果没有选中或选中多个返回 None
    """
    files = get_finder_selection()
    if len(files) == 1:
        return files[0]
    return None


def get_finder_current_dir() -> str | None:
    """
    获取 Finder 当前目录

    Returns:
        目录路径
    """
    script = '''
    tell application "Finder"
        if (count of (selection as list)) > 0 then
            set firstItem to item 1 of (selection as list)
            if class of firstItem is folder then
                POSIX path of (firstItem as alias)
            else
                POSIX path of (container of firstItem as alias)
            end if
        else
            POSIX path of (insertion location as alias)
        end if
    end tell
    '''
    try:
        result = subprocess.run(
            ['osascript', '-e', script],
            capture_output=True,
            text=True,
            timeout=5
        )
        return result.stdout.strip() or None
    except Exception:
        return None


def get_input_files(
    args: list[str],
    expected_ext = None,
    allow_multiple: bool = True
) -> list[str]:
    """
    获取输入文件列表
    优先使用命令行参数，没有参数时从 Finder 获取

    Args:
        args: 命令行参数（sys.argv[1:]）
        expected_ext: 期望的扩展名（字符串或列表，可选）
        allow_multiple: 是否允许多个文件

    Returns:
        有效的文件路径列表
    """
    files = []

    # 优先使用命令行参数（过滤空白参数，如 Raycast optional 参数未填时传入的空字符串）
    clean_args = [a for a in args if a.strip() and not a.startswith('-')]
    if clean_args:
        files = clean_args
    else:
        # 从 Finder 获取
        files = get_finder_selection()
        if not files:
            show_error("请在 Finder 中选择文件，或通过命令行传入文件路径")
            return []

    # 标准化 expected_ext
    allowed_exts = None
    if expected_ext:
        if isinstance(expected_ext, str):
            allowed_exts = [expected_ext.lower().lstrip('.')]
        else:
            allowed_exts = [e.lower().lstrip('.') for e in expected_ext]

    # 过滤有效文件
    valid_files = []
    for f in files:
        if not os.path.exists(f):
            show_warning(f"文件不存在: {f}")
            continue

        if allowed_exts:
            ext = Path(f).suffix.lower().lstrip('.')
            if ext not in allowed_exts:
                show_warning(f"跳过非 .{'/'.join(allowed_exts)} 文件: {f}")
                continue

        valid_files.append(f)

    # 检查数量
    if not allow_multiple and len(valid_files) > 1:
        show_warning("只支持处理单个文件，将处理第一个")
        valid_files = valid_files[:1]

    return valid_files


def require_single_file(args: list[str], expected_ext: str | None = None) -> str | None:
    """
    获取单个输入文件

    Args:
        args: 命令行参数（sys.argv[1:]）
        expected_ext: 期望的扩展名

    Returns:
        文件路径，失败返回 None
    """
    files = get_input_files(args, expected_ext, allow_multiple=False)
    return files[0] if files else None

