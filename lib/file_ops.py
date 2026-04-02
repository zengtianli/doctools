#!/usr/bin/env python3
"""
文件操作工具模块
"""

import json
import os
import shutil
import sys
from datetime import datetime
from pathlib import Path

from display import show_error, show_info


def check_file_extension(filepath: str, expected_ext: str) -> bool:
    """检查文件扩展名"""
    actual_ext = Path(filepath).suffix.lower().lstrip('.')
    return actual_ext == expected_ext.lower().lstrip('.')


def check_file_exists(filepath: str) -> bool:
    """检查文件是否存在"""
    return os.path.isfile(filepath)


def validate_input_file(filepath: str, expected_ext: str | None = None) -> bool:
    """验证输入文件"""
    if not check_file_exists(filepath):
        show_error(f"文件不存在: {filepath}")
        return False
    if expected_ext and not check_file_extension(filepath, expected_ext):
        show_error(f"请选择 .{expected_ext} 文件")
        return False
    return True


def find_files_by_extension(
    paths: str | Path | list[str | Path],
    extensions: str | list[str],
    recursive: bool = False
) -> list[Path]:
    """根据扩展名查找文件"""
    if isinstance(paths, (str, Path)):
        paths = [paths]
    if isinstance(extensions, str):
        extensions = [extensions]

    all_files = []
    for path in paths:
        path = Path(path)
        if path.is_file():
            ext = path.suffix.lower().lstrip('.')
            if ext in [e.lower().lstrip('.') for e in extensions]:
                all_files.append(path)
            continue
        if path.is_dir():
            for extension in extensions:
                extension = extension.lower().lstrip('.')
                pattern = f"**/*.{extension}" if recursive else f"*.{extension}"
                all_files.extend(path.glob(pattern))
    return sorted(set(all_files))


def ensure_directory(dir_path: str | Path) -> bool:
    """确保目录存在"""
    try:
        Path(dir_path).mkdir(parents=True, exist_ok=True)
        return True
    except Exception as e:
        show_error(f"无法创建目录: {dir_path} - {e}")
        return False


def check_command_exists(command: str) -> bool:
    """检查命令是否存在"""
    return shutil.which(command) is not None


def fatal_error(message: str):
    """致命错误 - 立即退出"""
    show_error(message)
    sys.exit(1)


def get_file_basename(filepath: str) -> str:
    """获取文件名（不含扩展名）"""
    return Path(filepath).stem


def check_python_packages(*packages: str) -> bool:
    """检查 Python 包是否已安装"""
    missing = []
    for pkg in packages:
        try:
            __import__(pkg)
        except ImportError:
            missing.append(pkg)
    if missing:
        show_error(f"缺少依赖包: {', '.join(missing)}")
        show_info(f"请运行: pip install {' '.join(missing)}")
        return False
    return True


def show_version_info(name: str, version: str, author: str = "", updated: str = ""):
    """显示版本信息"""
    print(f"📦 {name} v{version}")
    if author:
        print(f"👤 作者: {author}")
    if updated:
        print(f"📅 更新: {updated}")


def show_help_header(title: str, description: str = ""):
    """显示帮助信息头"""
    print(f"\n{'='*60}")
    print(f"📖 {title}")
    if description:
        print(f"   {description}")
    print(f"{'='*60}\n")


def show_help_footer():
    """显示帮助信息尾"""
    print(f"\n{'='*60}\n")


# ===== 文件批量操作 =====

def add_prefix(files: list[Path], prefix: str) -> list[tuple]:
    """批量添加文件名前缀"""
    results = []
    for f in files:
        if not f.is_file():
            continue
        new_name = f.parent / f"{prefix}{f.name}"
        try:
            f.rename(new_name)
            results.append((f.name, new_name.name))
        except Exception:
            pass
    return results


def move_up(files: list[Path]) -> list[Path]:
    """将文件移动到上级目录"""
    moved = []
    for f in files:
        if not f.exists():
            continue
        parent = f.parent.parent
        if not parent.exists():
            continue
        dst = parent / f.name
        if dst.exists():
            continue
        try:
            shutil.move(str(f), str(dst))
            moved.append(dst)
        except Exception:
            pass
    return moved


def flatten_dir(dir_path: str | Path) -> list[Path]:
    """将子目录中的文件扁平化到当前目录"""
    root = Path(dir_path)
    moved = []
    for f in root.rglob('*'):
        if f.is_file() and f.parent != root:
            dst = root / f.name
            if dst.exists():
                base, ext = dst.stem, dst.suffix
                counter = 1
                while dst.exists():
                    dst = root / f"{base}_{counter}{ext}"
                    counter += 1
            try:
                shutil.move(str(f), str(dst))
                moved.append(dst)
            except Exception:
                pass
    for d in sorted(root.rglob('*'), reverse=True):
        if d.is_dir() and not any(d.iterdir()):
            d.rmdir()
    return moved


def organize_by_type(files: list[Path], target_dir: str | Path) -> dict:
    """按文件类型整理到子目录"""
    TYPE_MAP = {
        'images': ['jpg', 'jpeg', 'png', 'gif', 'webp', 'svg', 'ico'],
        'documents': ['pdf', 'doc', 'docx', 'xls', 'xlsx', 'ppt', 'pptx', 'txt', 'md'],
        'videos': ['mp4', 'mov', 'avi', 'mkv', 'wmv'],
        'audio': ['mp3', 'wav', 'flac', 'aac', 'm4a'],
        'archives': ['zip', 'rar', '7z', 'tar', 'gz'],
        'code': ['py', 'js', 'ts', 'html', 'css', 'json', 'yaml', 'sh'],
    }
    results = {}
    target = Path(target_dir)
    for f in files:
        if not f.is_file():
            continue
        ext = f.suffix.lower().lstrip('.')
        folder = 'others'
        for cat, exts in TYPE_MAP.items():
            if ext in exts:
                folder = cat
                break
        dst_dir = target / folder
        dst_dir.mkdir(exist_ok=True)
        dst = dst_dir / f.name
        try:
            shutil.move(str(f), str(dst))
            results.setdefault(folder, []).append(f.name)
        except Exception:
            pass
    return results


def create_folder(name: str, target_dir: str | Path) -> Path | None:
    """在目标目录创建文件夹"""
    folder = Path(target_dir) / name
    try:
        folder.mkdir(parents=True, exist_ok=True)
        return folder
    except Exception:
        return None


# ===== 数据文件操作（合并自 hydraulic/_lib） =====

def read_geojson(geojson_path: str) -> dict:
    """
    读取 GeoJSON 文件

    Args:
        geojson_path: GeoJSON 文件路径

    Returns:
        dict: GeoJSON 数据字典

    Raises:
        SystemExit: 文件不存在或 JSON 解析失败时退出
    """
    try:
        with open(geojson_path, encoding='utf-8') as f:
            data = json.load(f)
        show_info(f"成功读取 GeoJSON 文件: {geojson_path}")
        show_info(f"  包含 {len(data.get('features', []))} 条记录")
        return data
    except FileNotFoundError:
        show_error(f"找不到文件 {geojson_path}")
        sys.exit(1)
    except json.JSONDecodeError as e:
        show_error(f"JSON 解析失败 - {e}")
        sys.exit(1)


def create_backup(file_path: str, suffix: str = "backup") -> str:
    """
    创建文件备份（带时间戳）

    Args:
        file_path: 要备份的文件路径
        suffix: 备份后缀名

    Returns:
        str: 备份文件路径

    Raises:
        SystemExit: 文件不存在时退出
    """
    if not os.path.exists(file_path):
        show_error(f"文件不存在 {file_path}")
        sys.exit(1)

    base_dir = os.path.dirname(file_path)
    base_name = os.path.splitext(os.path.basename(file_path))[0]
    ext = os.path.splitext(file_path)[1]
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    backup_path = os.path.join(base_dir, f"{base_name}_{suffix}_{timestamp}{ext}")

    shutil.copy2(file_path, backup_path)
    show_info(f"已创建备份文件: {backup_path}")

    return backup_path


def save_report(content: str, report_path: str):
    """
    保存文本报告到文件

    Args:
        content: 报告内容
        report_path: 报告保存路径
    """
    with open(report_path, 'w', encoding='utf-8') as f:
        f.write(content)
    show_info(f"报告已保存至: {report_path}")
