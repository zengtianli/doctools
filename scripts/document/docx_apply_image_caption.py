#!/usr/bin/env python3
"""
应用图片和图名样式工具
将文档中的图片段落和图名（题注）统一应用"ZDWP图名"样式

功能：
1. 找到包含图片的段落，应用"ZDWP图名"样式（图片居中）
2. 图片下一行（图名/题注）也应用"ZDWP图名"样式（文字居中）
3. 图片题注后面添加空行
4. 统计处理的图片数量

用法:
    python3 apply_image_caption_style.py <input.docx> [样式名称]

示例:
    python3 apply_image_caption_style.py document.docx "ZDWP图名"
    python3 apply_image_caption_style.py document.docx  # 默认使用"ZDWP图名"
"""

import sys
from pathlib import Path

from docx import Document
from docx.oxml.ns import qn

sys.path.insert(0, str(Path(__file__).parent.parent.parent / "lib"))
from finder import get_input_files
from progress import ProgressTracker


def find_style_fuzzy(doc, style_name):
    """
    模糊查找文档中的段落样式
    支持：
    1. 精确匹配
    2. 忽略空格差异（"ZDWP图名" 能匹配 "ZDWP 图名"）
    3. 包含关键词（如果精确+空格都不匹配，则查找包含关键词的）

    Args:
        doc: Document对象
        style_name: 样式名称

    Returns:
        str或None: 找到的实际样式名称，未找到返回None
    """
    try:
        styles = doc.styles

        # 准备搜索用的名称（去除所有空格）
        search_name_normalized = style_name.replace(" ", "").replace("\u3000", "")

        exact_match = None
        space_match = None
        partial_matches = []

        for style in styles:
            if not style.name or style.type != 1:  # 只查找段落样式
                continue

            # 1. 精确匹配
            if style.name == style_name:
                exact_match = style.name
                break

            # 2. 忽略空格的匹配
            style_normalized = style.name.replace(" ", "").replace("\u3000", "")
            if style_normalized == search_name_normalized:
                space_match = style.name

            # 3. 包含关键词的匹配（两个方向都检查）
            if search_name_normalized in style_normalized or style_normalized in search_name_normalized:  # noqa: SIM102
                # 只记录相关性高的（长度差距不要太大）
                if abs(len(style_normalized) - len(search_name_normalized)) <= 5:
                    partial_matches.append(style.name)

        # 返回优先级：精确匹配 > 空格匹配 > 部分匹配
        if exact_match:
            return exact_match
        if space_match:
            return space_match
        if partial_matches:
            return partial_matches[0]  # 返回第一个部分匹配

        return None

    except Exception as e:
        print(f"⚠️ 样式查找出错: {e}")
        return None


def check_style_exists(doc, style_name):
    """
    检查文档中是否存在指定的段落样式
    使用模糊匹配

    Args:
        doc: Document对象
        style_name: 样式名称

    Returns:
        bool: 样式是否存在
    """
    found_name = find_style_fuzzy(doc, style_name)
    return found_name is not None


def has_image(paragraph):
    """
    判断段落是否包含图片

    Args:
        paragraph: Paragraph对象

    Returns:
        bool: 是否包含图片
    """
    # 检查段落中的所有run
    for run in paragraph.runs:
        # 检查run中是否有图片元素
        for child in run._element:
            # w:drawing 表示图片/图形
            if child.tag == qn("w:drawing"):
                return True
            # w:pict 表示旧版图片格式
            if child.tag == qn("w:pict"):
                return True

    return False


def is_in_table(paragraph):
    """
    判断段落是否在表格内

    Args:
        paragraph: Paragraph对象

    Returns:
        bool: 是否在表格内
    """
    parent = paragraph._element.getparent()

    while parent is not None:
        if parent.tag == qn("w:tc"):  # w:tc = table cell
            return True
        parent = parent.getparent()

    return False


def add_blank_line_after_paragraph(doc, paragraph):
    """
    在段落后面添加一个空行

    Args:
        doc: Document对象
        paragraph: Paragraph对象
    """
    from docx.oxml import OxmlElement

    # 获取文档的body元素
    body = doc.element.body

    # 找到段落元素的位置
    para_element = paragraph._element
    para_index = list(body).index(para_element)

    # 检查段落后面是否已经有空段落
    if para_index + 1 < len(body):
        next_element = body[para_index + 1]
        # 如果后面是段落且为空，不再添加
        if next_element.tag == qn("w:p"):
            from docx.text.paragraph import Paragraph as Para

            next_para = Para(next_element, doc)
            if not next_para.text.strip():
                return  # 已经有空行了

    # 创建新的空段落元素
    new_para = OxmlElement("w:p")

    # 在段落后面插入空段落
    body.insert(para_index + 1, new_para)


def apply_image_caption_style(input_file, style_name="ZDWP图名"):
    """
    应用图片和图名样式

    Args:
        input_file: 输入文件路径
        style_name: 样式名称（默认"ZDWP图名"）
    """
    input_path = Path(input_file)

    # 检查文件是否存在
    if not input_path.exists():
        print(f"❌ 错误: 文件不存在: {input_file}")
        sys.exit(1)

    if input_path.suffix.lower() != ".docx":
        print("❌ 错误: 只支持 .docx 文件")
        sys.exit(1)

    print(f"🔄 正在处理文件: {input_path.name}")

    # 加载文档
    try:
        doc = Document(str(input_path))
    except Exception as e:
        print(f"❌ 错误: 无法打开文件: {e}")
        sys.exit(1)

    # 检查样式是否存在（使用模糊匹配）
    actual_style_name = find_style_fuzzy(doc, style_name)

    if not actual_style_name:
        print(f"❌ 错误: 文档中不存在段落样式 '{style_name}' 或类似样式")
        print("💡 请在Word中先添加该样式，或检查样式名称是否正确")
        sys.exit(1)

    if actual_style_name != style_name:
        print(f"ℹ️ 使用样式: {actual_style_name} (匹配: {style_name})")
    else:
        print(f"✅ 找到样式: {actual_style_name}")

    # 使用实际找到的样式名称
    style_name = actual_style_name

    # 备份原文件
    backup_path = input_path.with_suffix(".docx.backup")
    try:
        import shutil

        shutil.copy2(str(input_path), str(backup_path))
        print(f"ℹ️ 已备份原文件: {backup_path.name}")
    except Exception as e:
        print(f"⚠️ 备份失败: {e}")

    # 统计信息
    image_count = 0
    caption_count = 0
    error_count = 0

    print("🔄 正在应用图片和图名样式...")

    paragraphs = doc.paragraphs

    for i, paragraph in enumerate(paragraphs):
        try:
            # 跳过表格内的段落
            if is_in_table(paragraph):
                continue

            # 检查是否包含图片
            if has_image(paragraph):
                # 应用样式到图片段落
                paragraph.style = style_name
                image_count += 1

                # 检查下一行是否存在
                if i + 1 < len(paragraphs):
                    next_paragraph = paragraphs[i + 1]

                    # 如果下一行不在表格内且有内容，应用样式
                    if not is_in_table(next_paragraph) and next_paragraph.text.strip():
                        next_paragraph.style = style_name
                        caption_count += 1

                        # 在图片题注后面添加空行
                        add_blank_line_after_paragraph(doc, next_paragraph)

        except Exception as e:
            print(f"⚠️ 段落 {i + 1} 处理失败: {e}")
            error_count += 1

    # 保存文档
    try:
        doc.save(str(input_path))
        print("✅ 样式应用完成!")
        print(f"   - 图片段落: {image_count} 个")
        print(f"   - 图名段落: {caption_count} 个")
        if error_count > 0:
            print(f"   - 失败: {error_count} 个")
        print(f"   - 已保存: {input_path.name}")
        if backup_path.exists():
            print(f"ℹ️ 如需恢复，请使用备份文件: {backup_path.name}")
    except Exception as e:
        print(f"❌ 保存失败: {e}")
        sys.exit(1)


def main():
    # 获取输入文件（优先命令行参数，否则从 Finder 获取）
    files = get_input_files(sys.argv[1:], expected_ext="docx")

    if not files:
        print("用法: python3 apply_image_caption_style.py <input.docx> [样式名称]")
        print("      或在 Finder 中选择 .docx 文件后运行")
        print("示例: python3 apply_image_caption_style.py document.docx")
        print('      python3 apply_image_caption_style.py document.docx "ZDWP图名"')
        sys.exit(1)

    # 样式名称从第一个非文件参数获取
    style_name = "ZDWP图名"
    for arg in sys.argv[1:]:
        if not Path(arg).exists() and not arg.startswith("-"):
            style_name = arg
            break

    tracker = ProgressTracker()

    for file_path in files:
        print(f"\n{'=' * 50}")
        print(f"处理文件: {Path(file_path).name}")
        print("=" * 50)
        try:
            apply_image_caption_style(str(file_path), style_name)
            tracker.add_success()
        except Exception as e:
            print(f"❌ 处理失败: {e}")
            tracker.add_error()

    print(f"\n{'=' * 50}")
    tracker.show_summary("文件处理")


if __name__ == "__main__":
    main()
