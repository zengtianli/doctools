#!/usr/bin/env python3
"""
通用函数库 - 统一导出入口

使用方式（sys.path 已指向 lib/）：
    from display import show_success
    from finder import get_input_files

或通过包导入：
    from lib import show_success, get_input_files
"""

__version__ = "1.0.0"

# === clipboard ===
from clipboard import (  # noqa: F401
    copy_to_clipboard,
    get_clipboard_files,
    get_from_clipboard,
    paste_files,
)

# === display ===
from display import (  # noqa: F401
    show_error,
    show_info,
    show_processing,
    show_progress,
    show_success,
    show_warning,
)

# === env ===
from env import PROJECT_ROOT, PYTHON_PATH, USAGE_LOG  # noqa: F401

# === file_ops ===
from file_ops import (  # noqa: F401
    add_prefix,
    check_command_exists,
    check_file_exists,
    check_file_extension,
    check_python_packages,
    create_folder,
    ensure_directory,
    fatal_error,
    find_files_by_extension,
    flatten_dir,
    get_file_basename,
    move_up,
    organize_by_type,
    show_help_footer,
    show_help_header,
    show_version_info,
    validate_input_file,
)

# === finder ===
from finder import (  # noqa: F401
    get_finder_current_dir,
    get_finder_selection,
    get_finder_selection_single,
    get_input_files,
    require_single_file,
)

# === progress ===
from progress import ProgressTracker  # noqa: F401

# === usage_log ===
from usage_log import log_usage  # noqa: F401
