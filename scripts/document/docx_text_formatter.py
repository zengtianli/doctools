#!/usr/bin/env python3
"""
文本格式自动修复工具 - DOCX版本
支持处理Word文档中的格式问题

功能列表：
1. 双引号格式修复
   将所有双引号统一为中文标准引号："和"
   奇数个引号 → " (中文左双引号 U+201C)
   偶数个引号 → " (中文右双引号 U+201D)

2. 英文标点符号转中文
   , → ，  (逗号)
   : → ：  (冒号)
   ; → ；  (分号)
   ! → ！  (感叹号)
   ? → ？  (问号)
   ( → （  (左括号)
   ) → ）  (右括号)

3. 中文单位转标准符号
   面积：平方米 → m²、平方公里 → km²
   体积：立方米 → m³、立方厘米 → cm³
   长度：公里 → km、厘米 → cm、毫米 → mm
   质量：公斤 → kg、毫克 → mg
   容量：毫升 → mL、微升 → μL
   时间：小时 → h、分钟 → min、秒钟 → s
   温度：摄氏度 → ℃、华氏度 → ℉

4. 单位上标格式化
   m2 → m²、m3 → m³、km2 → km²、km3 → km³

覆盖范围（2026-07-26 扩齐）：
   正文 + 表格（含嵌套表）+ 文本框 + 超链接 + **审阅修订（w:ins 插入态 / w:del 删除态）**
   + **批注 comments.xml** + **脚注 / 尾注** + 页眉页脚（SKIP_FOOTER=False 时）。
   原先走 python-docx 的 `Document.paragraphs` / `Paragraph.runs`，那两个 API 只认
   直接子节点，带修订的稿子里 w:ins/w:del 段一个字都改不到（实测某审阅稿 24 个 run /
   3101 字全漏），批注与脚注更是整个 part 没碰。现统一走 docx_xml 的元素级遍历。
   删除态文本改的是 w:delText（不退化成 w:t，修订结构原样保留）。

使用方法：
    python3 text_formatter_docx.py 文件名.docx
"""

import copy
import sys
from pathlib import Path

from docx import Document
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

sys.path.insert(0, str(Path(__file__).parent.parent.parent / "lib"))
sys.path.insert(0, str(Path.home() / "Dev" / "tools" / "dev" / "lib"))  # canonical 5 modules
from docx_xml import (
    in_deleted,
    in_inserted,
    iter_paragraphs,
    iter_text_roots,
    para_own_runs,
    run_text,
    set_run_text,
    text_tag_for,
)
from file_ops import clear_quarantine
from finder import get_input_files
from progress import ProgressTracker
from text_fixes import fix_punctuation, fix_quotes, fix_units

sys.path.insert(0, str(Path(__file__).resolve().parent))
from docx_write_gate import WriteGate  # 原地写回并发门（同目录 SSOT）

# ===== 配置选项 =====
SKIP_FOOTER = True

# 选择性修复开关（默认全开=向后兼容；命令行 --quotes/--punct/--units 任一指定则只做指定项）
# 用途：中文期刊投稿(GB/T 7714 参考文献区标点须半角)只想修引号→ --quotes，不碰标点/单位
DO_QUOTES = True
DO_PUNCT = True
DO_UNITS = True
IN_PLACE = False   # True 时原地写回源文件 + 备份 .bak-时间戳（Work §1.5 协议），不另存 _fixed


QUOTE_CHARS = {"\u201c", "\u201d"}
QUOTE_FONT = "\u5b8b\u4f53"  # 宋体


def _set_run_font_songti(run_element):
    """为 run 的 rPr 设置宋体字体（ascii + hAnsi + eastAsia）"""
    rPr = run_element.find(qn("w:rPr"))
    if rPr is None:
        rPr = OxmlElement("w:rPr")
        run_element.insert(0, rPr)
    rFonts = rPr.find(qn("w:rFonts"))
    if rFonts is None:
        rFonts = OxmlElement("w:rFonts")
        rPr.insert(0, rFonts)
    rFonts.set(qn("w:ascii"), QUOTE_FONT)
    rFonts.set(qn("w:hAnsi"), QUOTE_FONT)
    rFonts.set(qn("w:eastAsia"), QUOTE_FONT)
    rFonts.set(qn("w:hint"), "eastAsia")


def _split_text_at_quotes(text):
    """
    将文本在引号位置切片，返回 [(text, is_quote), ...] 片段列表。
    如果没有引号，返回 None。
    """
    if not text or not any(c in QUOTE_CHARS for c in text):
        return None

    segments = []
    buf = []
    for c in text:
        if c in QUOTE_CHARS:
            if buf:
                segments.append(("".join(buf), False))
                buf = []
            segments.append((c, True))
        else:
            buf.append(c)
    if buf:
        segments.append(("".join(buf), False))
    return segments


def _apply_quote_split(r, segments):
    """
    根据 segments 拆分 run 元素 r：第一段复用原 run，后续段插入新 run。
    引号段设置宋体。删除态 run（w:del 内）新建的文本元素仍是 w:delText。
    """
    parent = r.getparent()
    text_tag = text_tag_for(r)          # w:del / w:moveFrom 内必须继续用 w:delText

    # 第一段复用原 run
    first_text, first_is_quote = segments[0]
    set_run_text(r, first_text)
    if first_is_quote:
        _set_run_font_songti(r)

    # 后续段：复制原 run 的格式，插入到原 run 之后
    insert_after = r
    for seg_text, is_quote in segments[1:]:
        new_r = copy.deepcopy(r)
        # 只留 rPr：deepcopy 带来的旧文本、以及 drawing/footnoteReference 等
        # 一次性载荷都必须清掉，否则拆一次 run 就把图片/脚注引用复制一份出来
        for child in list(new_r):
            if child.tag != qn("w:rPr"):
                new_r.remove(child)
        t_elem = OxmlElement(text_tag)
        t_elem.text = seg_text
        # 保留空格
        t_elem.set(qn("xml:space"), "preserve")
        new_r.append(t_elem)

        if is_quote:
            _set_run_font_songti(new_r)
        else:
            # 非引号段恢复原 run 的字体（去掉宋体覆盖）
            rPr = new_r.find(qn("w:rPr"))
            if rPr is not None:
                rFonts = rPr.find(qn("w:rFonts"))
                orig_rPr = r.find(qn("w:rPr"))
                orig_rFonts = orig_rPr.find(qn("w:rFonts")) if orig_rPr is not None else None
                if rFonts is not None and orig_rFonts is not None:
                    # 用原始字体信息覆盖
                    rPr.replace(rFonts, copy.deepcopy(orig_rFonts))
                elif rFonts is not None and orig_rFonts is None:
                    rPr.remove(rFonts)

        parent.insert(list(parent).index(insert_after) + 1, new_r)
        insert_after = new_r


def _tally(stats, scope, r, n):
    """按域记账：修订态单独计，好让用户看见「审阅段确实改到了」。"""
    if not n:
        return
    key = scope
    if scope == "正文":
        if in_deleted(r):
            key = "修订·删除态"
        elif in_inserted(r):
            key = "修订·插入态"
    stats["scopes"][key] = stats["scopes"].get(key, 0) + n


def process_paragraph_element(p, stats, scope="正文"):
    """
    处理一个 w:p 元素内的文本，逐 run 处理以保持每个 run 的字体格式。
    引号拆分为独立 run 并设置宋体。

    走 para_own_runs 而非 python-docx 的 `Paragraph.runs`：后者只认 `./w:r`，
    审阅修订（w:ins/w:del）、超链接、smartTag 里的 run 会被静默跳过。
    """
    # 先收集需要处理的 runs（遍历中会插入新 run，所以先快照）
    original_runs = para_own_runs(p)
    if not original_runs:
        return

    # 引号计数器跨run维护，保证配对正确
    quote_counter = 0

    for r in original_runs:
        original_text = run_text(r)
        if not original_text:
            continue

        # 按 toggle 选择性应用转换（默认全开）
        fixed_text = original_text
        quote_count = punct_count = unit_count = 0
        if DO_QUOTES:
            fixed_text, quote_count, quote_counter = fix_quotes(fixed_text, quote_counter)
        if DO_PUNCT:
            fixed_text, punct_count = fix_punctuation(fixed_text)
        if DO_UNITS:
            fixed_text, unit_count = fix_units(fixed_text)

        # 更新统计
        stats["quotes"] += quote_count
        stats["punctuation"] += punct_count
        stats["units"] += unit_count
        _tally(stats, scope, r, quote_count + punct_count + unit_count)

        if fixed_text != original_text:
            set_run_text(r, fixed_text)

        # 拆分引号为独立 run 并设置宋体（仅在修引号时）
        if DO_QUOTES:
            segments = _split_text_at_quotes(run_text(r))
            if segments:
                _apply_quote_split(r, segments)


def process_root(root, stats, scope="正文"):
    """处理一个 XML 根（正文 body / 批注 / 脚注 / 页眉…）下的全部段落。

    `iter_paragraphs` 递归到嵌套表格、文本框、修订块里的 w:p；`para_own_runs`
    在嵌套 w:p 处剪枝，所以每个 run 恰好被处理一次。
    """
    for p in list(iter_paragraphs(root)):
        process_paragraph_element(p, stats, scope)


def process_docx(input_file):
    """
    处理DOCX文件
    """
    input_path = Path(input_file)

    if not input_path.exists():
        print(f"❌ 错误：文件不存在 - {input_file}")
        return False

    if input_path.suffix.lower() != ".docx":
        print("❌ 错误：文件必须是.docx格式")
        return False

    # 生成输出文件名（--in-place 时原地写回 + 备份；否则另存 _fixed）
    write_gate = WriteGate(input_path) if IN_PLACE else None  # 读入前 capture 基线
    if IN_PLACE:
        import shutil as _shutil
        from datetime import datetime as _dt
        bak = input_path.with_name(
            input_path.name + ".bak-" + _dt.now().strftime("%Y%m%d-%H%M%S"))
        _shutil.copy2(input_path, bak)
        print(f"🗄  已备份: {bak.name}")
        output_path = input_path
    else:
        output_path = input_path.parent / f"{input_path.stem}_fixed{input_path.suffix}"

    try:
        # 读取文档
        print(f"📖 正在读取文件: {input_path.name}")
        doc = Document(input_path)

        # 统计信息（scopes 分域记账：正文 / 修订·插入态 / 修订·删除态 / 批注 / 脚注…）
        stats = {"quotes": 0, "punctuation": 0, "units": 0, "scopes": {}}

        # 处理所有承载文本的 part：正文（含表格/嵌套表/文本框/超链接/审阅修订）
        # + 批注 comments.xml + 脚注 + 尾注（+ 页眉页脚，仅 SKIP_FOOTER=False 时）
        print("🔄 正在处理正文 / 表格 / 修订 / 批注 / 脚注...")
        flushes = []
        for label, root, flush in iter_text_roots(doc, include_headers=not SKIP_FOOTER):
            process_root(root, stats, label)
            if flush is not None:
                flushes.append(flush)
        for flush in flushes:          # 通用 Part（脚注/尾注）改完必须回写 blob
            flush()

        # 处理页眉页脚
        if SKIP_FOOTER:
            # 彻底删除页眉页脚
            print("🗑️  正在删除页眉页脚...")
            from docx.oxml.ns import qn

            for section in doc.sections:
                sectPr = section._sectPr
                # 删除所有页眉引用
                for headerRef in sectPr.findall(qn("w:headerReference")):
                    sectPr.remove(headerRef)
                # 删除所有页脚引用
                for footerRef in sectPr.findall(qn("w:footerReference")):
                    sectPr.remove(footerRef)

        # 保存文件
        print("💾 正在保存文件...")
        if write_gate is not None:
            write_gate.assert_unchanged()  # 源文件被 WPS/其他会话改过 → 拒写(逃生 DOCX_GATE_OK=1)
        try:
            doc.save(output_path)
            clear_quarantine(output_path)
        except (PermissionError, OSError) as e:
            fallback = Path.home() / "Downloads" / Path(output_path).name
            print(f"⚠️ 源目录不可写: {e}")
            print(f"   降级到: {fallback}")
            doc.save(fallback)
            clear_quarantine(fallback)
            output_path = fallback

        print("✅ 处理完成！")
        print(f"   - 共替换了 {stats['quotes']} 个引号")
        print(f"   - 共替换了 {stats['punctuation']} 个标点符号")
        print(f"   - 共转换了 {stats['units']} 个单位")
        extra = {k: v for k, v in stats["scopes"].items() if k != "正文"}
        if extra:
            print("   - 非正文域命中: " + "  ".join(f"{k} {v} 处" for k, v in extra.items()))
        print(f"   - 输出文件: {output_path.name}")

        return True

    except Exception as e:
        print(f"❌ 处理失败：{e}")
        import traceback

        traceback.print_exc()
        return False


if __name__ == "__main__":
    # 摘选择性 flag（在 get_input_files 前），剩余为文件参数
    _argv = sys.argv[1:]
    _flags = {"--quotes", "--punct", "--units", "--in-place", "--keep-headers"}
    _sel = {a for a in _argv if a in _flags}
    _unknown = [a for a in _argv if a.startswith("--") and a not in _flags]
    if _unknown:   # 未知 flag 不得被当成文件名静默吞掉→曾致 --help 直接跑全量改写
        print(f"❌ 未知参数：{' '.join(_unknown)}")
        print(f"   可用：{' '.join(sorted(_flags))}")
        sys.exit(2)
    _argv = [a for a in _argv if a not in _flags]
    if _sel & {"--quotes", "--punct", "--units"}:
        DO_QUOTES = "--quotes" in _sel
        DO_PUNCT = "--punct" in _sel
        DO_UNITS = "--units" in _sel
    IN_PLACE = "--in-place" in _sel
    if "--keep-headers" in _sel:
        SKIP_FOOTER = False    # 保留页眉页脚引用（默认 True=删，供 md2word 新建链用）

    # 获取输入文件（优先命令行参数，否则从 Finder 获取）
    files = get_input_files(_argv, expected_ext="docx")

    if not files:
        print("❌ 错误：缺少文件名参数")
        print("\n使用方法：")
        print("  1. 在 Finder 中选中 .docx 文件，然后运行此脚本")
        print("  2. 或在命令行中提供文件路径")
        print("\n示例：")
        print("    python3 docx_text_formatter.py 文件名.docx")
        print("    python3 docx_text_formatter.py file1.docx file2.docx")
        sys.exit(1)

    print("=" * 50)
    print("文本格式自动修复工具 - DOCX版本")
    print("=" * 50)

    tracker = ProgressTracker()

    for file_path in files:
        print(f"\n处理文件: {Path(file_path).name}")
        success = process_docx(str(file_path))
        if success:
            tracker.add_success()
        else:
            tracker.add_error()

    print("\n" + "=" * 50)
    tracker.show_summary("文件处理")
