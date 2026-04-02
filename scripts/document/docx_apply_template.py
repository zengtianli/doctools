#!/usr/bin/env python3
"""
docx_apply_template.py - Word 文档样式工具（套模板 + 清理）

子命令：
  apply   - 给普通 docx 套上模板样式（默认）
  cleanup - 样式清理：重命名、合并、删除未使用样式

用法：
  # 套模板（默认子命令，兼容旧调用方式）
  python docx_apply_template.py input.docx
  python docx_apply_template.py apply input.docx -t template.docx -o output.docx

  # 样式清理
  python docx_apply_template.py cleanup input.docx -o output.docx --config rules.json
  python docx_apply_template.py cleanup input.docx --preview
"""

import argparse
import json
import os
import subprocess
import sys
import tempfile
import zipfile
from copy import deepcopy
from io import BytesIO
from pathlib import Path

sys.path.insert(0, str(Path(__file__).parent.parent.parent / "lib"))

try:
    from lxml import etree
except ImportError:
    print("❌ 需要安装 lxml: pip install lxml")
    sys.exit(1)

from docx_xml import NSMAP

# 复用 md_docx_template 的样式提取
from md_docx_template import DEFAULT_TEMPLATE, extract_styles_xml

# 放在 document/ 同级，但导入需要处理路径
sys.path.insert(0, str(Path(__file__).parent))

W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"


# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
#  apply 子命令：套模板样式
# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━


def apply_styles_to_docx(docx_path, styles_xml_path, output_path):
    """把模板样式注入到已有 docx 中"""

    # 读取提取的样式
    with open(styles_xml_path, "rb") as f:
        new_styles = etree.parse(f).getroot()

    with tempfile.TemporaryDirectory() as tmpdir:
        # 解压目标 docx
        with zipfile.ZipFile(docx_path, "r") as zf:
            zf.extractall(tmpdir)

        # 读取原 styles.xml
        orig_styles_path = os.path.join(tmpdir, "word", "styles.xml")
        with open(orig_styles_path, "rb") as f:
            orig_root = etree.parse(f).getroot()

        # 收集新样式的 ID
        new_style_ids = set()
        for style in new_styles.findall(".//w:style", NSMAP):
            style_id = style.get(f"{{{NSMAP['w']}}}styleId")
            new_style_ids.add(style_id)

        # 删除原文档中同名样式
        for style in orig_root.findall(".//w:style", NSMAP):
            style_id = style.get(f"{{{NSMAP['w']}}}styleId")
            if style_id in new_style_ids:
                orig_root.remove(style)

        # 添加新样式
        for style in new_styles.findall(".//w:style", NSMAP):
            orig_root.append(deepcopy(style))

        # 更新 docDefaults
        new_defaults = new_styles.find(".//w:docDefaults", NSMAP)
        if new_defaults is not None:
            old_defaults = orig_root.find(".//w:docDefaults", NSMAP)
            if old_defaults is not None:
                orig_root.remove(old_defaults)
            orig_root.insert(0, deepcopy(new_defaults))

        # 写回 styles.xml
        tree = etree.ElementTree(orig_root)
        tree.write(orig_styles_path, encoding="UTF-8", xml_declaration=True)

        # 重新打包
        with zipfile.ZipFile(output_path, "w", zipfile.ZIP_DEFLATED) as zf:
            for root_dir, _dirs, files in os.walk(tmpdir):
                for file in files:
                    file_path = os.path.join(root_dir, file)
                    arcname = os.path.relpath(file_path, tmpdir)
                    zf.write(file_path, arcname)

    print(f"✅ 输出: {output_path}")
    return output_path


def get_finder_selection():
    """获取 Finder 选中的文件"""
    script = """
    tell application "Finder"
        set sel to selection
        if (count of sel) > 0 then
            return POSIX path of (item 1 of sel as alias)
        end if
    end tell
    """
    result = subprocess.run(["osascript", "-e", script], capture_output=True, text=True)
    return result.stdout.strip()


def cmd_apply(args):
    """apply 子命令入口"""
    # 无参数时从 Finder 获取选中的 .docx 文件
    if not args.input:
        finder_file = get_finder_selection()
        if finder_file and finder_file.endswith(".docx"):
            args.input = finder_file
            print(f"📄 从 Finder 获取: {os.path.basename(finder_file)}")
        else:
            print("❌ 请在 Finder 中选择一个 .docx 文件")
            sys.exit(1)

    if not os.path.exists(args.input):
        print(f"❌ 文件不存在: {args.input}")
        sys.exit(1)

    # 模板
    template = args.template or DEFAULT_TEMPLATE
    if not os.path.exists(template):
        print(f"❌ 模板不存在: {template}")
        print("   请指定模板: -t template.docx")
        sys.exit(1)

    # 输出路径
    if args.output:
        output_path = args.output
    else:
        base, ext = os.path.splitext(args.input)
        output_path = f"{base}_styled{ext}"

    # 提取模板样式 → 注入目标 docx
    with tempfile.TemporaryDirectory() as tmpdir:
        print(f"📋 模板: {os.path.basename(template)}")
        styles_path, config = extract_styles_xml(template, tmpdir)
        apply_styles_to_docx(args.input, styles_path, output_path)


# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
#  cleanup 子命令：样式清理
# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━


def load_docx(path: str):
    """加载 docx，返回 (zip_bytes_dict, styles_tree, doc_tree)"""
    files = {}
    with zipfile.ZipFile(path, "r") as zf:
        for info in zf.infolist():
            files[info.filename] = zf.read(info.filename)

    styles_tree = etree.fromstring(files["word/styles.xml"])
    doc_tree = etree.fromstring(files["word/document.xml"])
    return files, styles_tree, doc_tree


def save_docx(files: dict, styles_tree, doc_tree, output_path: str):
    """保存修改后的 docx"""
    files["word/styles.xml"] = etree.tostring(styles_tree, xml_declaration=True, encoding="UTF-8", standalone=True)
    files["word/document.xml"] = etree.tostring(doc_tree, xml_declaration=True, encoding="UTF-8", standalone=True)

    buf = BytesIO()
    with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as zf:
        for name, data in files.items():
            zf.writestr(name, data)
    with open(output_path, "wb") as f:
        f.write(buf.getvalue())


def get_style_map(styles_tree) -> dict:
    """从 styles.xml 提取 styleId → {name, type, basedOn}"""
    result = {}
    for s in styles_tree.findall(f".//{{{W}}}style"):
        sid = s.get(f"{{{W}}}styleId", "")
        stype = s.get(f"{{{W}}}type", "")
        ne = s.find(f"{{{W}}}name")
        name = ne.get(f"{{{W}}}val", sid) if ne is not None else sid
        based = s.find(f"{{{W}}}basedOn")
        base_id = based.get(f"{{{W}}}val", "") if based is not None else ""
        result[sid] = {"name": name, "type": stype, "basedOn": base_id, "elem": s}
    return result


def get_used_style_ids(doc_tree) -> set:
    """扫描 document.xml，找出所有被使用的样式 ID"""
    used = set()
    # 段落样式
    for ps in doc_tree.iter(f"{{{W}}}pStyle"):
        used.add(ps.get(f"{{{W}}}val", ""))
    # 字符样式
    for rs in doc_tree.iter(f"{{{W}}}rStyle"):
        used.add(rs.get(f"{{{W}}}val", ""))
    # 表格样式
    for ts in doc_tree.iter(f"{{{W}}}tblStyle"):
        used.add(ts.get(f"{{{W}}}val", ""))
    used.discard("")
    return used


def get_basedOn_ids(style_map: dict, used_ids: set) -> set:
    """递归找出所有被使用样式依赖的 basedOn 样式"""
    needed = set(used_ids)
    queue = list(used_ids)
    while queue:
        sid = queue.pop()
        info = style_map.get(sid)
        if info and info["basedOn"] and info["basedOn"] not in needed:
            needed.add(info["basedOn"])
            queue.append(info["basedOn"])
    return needed


# Word 内置样式 ID，不能删
BUILTIN_KEEP = {
    "a",  # Normal
    "a0",  # Default Paragraph Font
}


def cleanup(
    files: dict,
    styles_tree,
    doc_tree,
    renames: dict = None,
    merges: dict = None,
    delete_unused: bool = True,
    preview: bool = False,
) -> list[str]:
    """
    执行样式清理。

    Args:
        renames: {old_display_name: new_display_name} 重命名映射
        merges: {from_style_id: to_style_id} 合并映射
        delete_unused: 是否删除未使用的样式
        preview: 只生成报告不修改

    Returns:
        操作日志列表
    """
    log = []
    style_map = get_style_map(styles_tree)

    # ── 1. 合并样式 ──
    if merges:
        for from_id, to_id in merges.items():
            if from_id not in style_map:
                log.append(f"⚠️  合并跳过：源样式 ID '{from_id}' 不存在")
                continue
            if to_id not in style_map:
                log.append(f"⚠️  合并跳过：目标样式 ID '{to_id}' 不存在")
                continue

            count = 0
            if not preview:
                for ps in doc_tree.iter(f"{{{W}}}pStyle"):
                    if ps.get(f"{{{W}}}val") == from_id:
                        ps.set(f"{{{W}}}val", to_id)
                        count += 1
            else:
                for ps in doc_tree.iter(f"{{{W}}}pStyle"):
                    if ps.get(f"{{{W}}}val") == from_id:
                        count += 1

            from_name = style_map[from_id]["name"]
            to_name = style_map[to_id]["name"]
            log.append(f"🔀 合并：{from_name} ({from_id}) → {to_name} ({to_id})，{count} 个段落")

    # ── 2. 重命名样式 ──
    if renames:
        # 建立 name → styleId 映射
        name_to_id = {}
        for sid, info in style_map.items():
            name_to_id[info["name"]] = sid

        for old_name, new_name in renames.items():
            sid = name_to_id.get(old_name)
            if not sid:
                log.append(f"⚠️  重命名跳过：找不到样式 '{old_name}'")
                continue
            if not preview:
                ne = style_map[sid]["elem"].find(f"{{{W}}}name")
                if ne is not None:
                    ne.set(f"{{{W}}}val", new_name)
            log.append(f"✏️  重命名：{old_name} → {new_name}")

    # ── 3. 删除未使用样式 ──
    if delete_unused:
        # 重新扫描（合并后使用情况可能变了）
        used_ids = get_used_style_ids(doc_tree)
        needed_ids = get_basedOn_ids(style_map, used_ids)
        needed_ids |= BUILTIN_KEEP

        # 内置 heading/toc 样式保留（即使没用到，删了可能有问题）
        for sid in style_map:
            if sid.startswith("TOC") or sid in ("a", "a0"):
                needed_ids.add(sid)

        to_delete = []
        for sid, info in style_map.items():
            if sid not in needed_ids:
                to_delete.append((sid, info["name"]))

        if to_delete:
            log.append(f"\n🗑️  删除 {len(to_delete)} 个未使用样式：")
            for sid, name in sorted(to_delete, key=lambda x: x[1]):
                log.append(f"   - {name} ({sid})")
                if not preview:
                    for s in styles_tree.findall(f".//{{{W}}}style"):
                        if s.get(f"{{{W}}}styleId") == sid:
                            s.getparent().remove(s)
                            break

    return log


def cmd_cleanup(args):
    """cleanup 子命令入口"""
    if not args.output and not args.preview:
        print("请指定 -o 输出文件或使用 --preview 预览模式")
        sys.exit(1)

    # 加载配置
    if args.config:
        with open(args.config, encoding="utf-8") as f:
            config = json.load(f)
    else:
        config = {}

    files, styles_tree, doc_tree = load_docx(args.input)

    log = cleanup(
        files,
        styles_tree,
        doc_tree,
        renames=config.get("renames"),
        merges=config.get("merges"),
        delete_unused=config.get("delete_unused", True),
        preview=args.preview,
    )

    for line in log:
        print(line)

    if not args.preview and args.output:
        save_docx(files, styles_tree, doc_tree, args.output)
        print(f"\n📄 已保存到 {args.output}")
    elif args.preview:
        print("\n（预览模式，未修改文件）")


# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
#  主入口：argparse 子命令
# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━


def main():
    parser = argparse.ArgumentParser(
        description="Word 文档样式工具（套模板 + 清理）",
        formatter_class=argparse.RawDescriptionHelpFormatter,
    )
    subparsers = parser.add_subparsers(dest="command")

    # ── apply 子命令 ──
    apply_parser = subparsers.add_parser("apply", help="给普通 docx 套上模板样式")
    apply_parser.add_argument("input", nargs="?", help="目标 docx 文件")
    apply_parser.add_argument("-t", "--template", help="模板 docx 文件")
    apply_parser.add_argument("-o", "--output", help="输出文件")

    # ── cleanup 子命令 ──
    cleanup_parser = subparsers.add_parser("cleanup", help="样式清理：重命名、合并、删除未使用")
    cleanup_parser.add_argument("input", help="输入 .docx 文件")
    cleanup_parser.add_argument("-o", "--output", help="输出 .docx 文件")
    cleanup_parser.add_argument("--config", help="清理规则 JSON 文件")
    cleanup_parser.add_argument("--preview", action="store_true", help="只预览，不实际修改")

    args = parser.parse_args()

    # 默认行为：无子命令时当作 apply（兼容旧调用方式）
    if args.command is None:
        # 重新解析，把所有参数当 apply 处理
        apply_parser_compat = argparse.ArgumentParser(description="给普通 docx 套上模板样式")
        apply_parser_compat.add_argument("input", nargs="?", help="目标 docx 文件")
        apply_parser_compat.add_argument("-t", "--template", help="模板 docx 文件")
        apply_parser_compat.add_argument("-o", "--output", help="输出文件")
        args = apply_parser_compat.parse_args()
        cmd_apply(args)
    elif args.command == "apply":
        cmd_apply(args)
    elif args.command == "cleanup":
        cmd_cleanup(args)


if __name__ == "__main__":
    main()
