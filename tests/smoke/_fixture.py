#!/usr/bin/env python3
"""_fixture.py — smoke 用的 fixture 造件器（纯 stdlib + python-docx，不读仓外真实文件）。

为什么造而不是拿现成的：`templates/template.docx` 是指向 `~/Work` 真实交付件的
symlink（15MB 真业务数据）。smoke 一旦指向它，① 测试依赖一个不在版本控制里的文件
② 写模式动词有机会写进用户在跑的项目。所以 fixture 全部现造，且只落 `tmp_path`。

造出来的东西要足够"真"，否则大半动词是**空跑**——rc=0 但一件事没干，
而空跑的 rc=0 和真干活的 rc=0 长得一模一样。所以 fixture 里塞齐：

    多级标题(H1~H4) · 正文 · 2 张表 · 3 张图 · 题注(图1-1/表1-1) · 交叉引用
    两个分节 · 页眉页脚 · 书签(_Toc_*) · 域(REF/TOC/PAGE) · 重复图片(供去重)
    直接格式脏段(供 clear-direct-format / scan-ppr) · 等线 run(供 fonts normalize)

⚠ **fixture 内容与若干动词的 expect_rc 是绑定的**（`audit-styleset *` 靠样式集判
severity、`para scan-ppr` 靠脏段数、`health diagnose` 靠病种命中）。改这个文件
= 可能改掉 `_verb_specs` 里那几条的 expect_rc，改完必须重跑 smoke。
"""
from __future__ import annotations

import json
import struct
import sys
import zlib
from pathlib import Path

REPO = Path(__file__).resolve().parents[2]
sys.path.append(str(REPO / "lib"))
import docx_safe_save  # noqa: E402,F401  （存盘收口，见 CLAUDE.md）

from docx import Document  # noqa: E402
from docx.enum.section import WD_SECTION  # noqa: E402
from docx.enum.text import WD_ALIGN_PARAGRAPH  # noqa: E402
from docx.oxml import OxmlElement  # noqa: E402
from docx.oxml.ns import qn  # noqa: E402
from docx.shared import Cm, Pt, RGBColor  # noqa: E402


# ─────────────────────────── 零件 ───────────────────────────

def make_png(path: Path, w: int = 160, h: int = 120, rgb=(66, 133, 244)) -> Path:
    """纯 stdlib 造 PNG（不引 PIL —— smoke 不该因为缺一个绘图库就跑不起来）。"""
    def chunk(typ: bytes, data: bytes) -> bytes:
        c = typ + data
        return struct.pack(">I", len(data)) + c + struct.pack(">I", zlib.crc32(c) & 0xFFFFFFFF)

    raw = b"".join(b"\x00" + bytes(rgb) * w for _ in range(h))
    png = (b"\x89PNG\r\n\x1a\n"
           + chunk(b"IHDR", struct.pack(">IIBBBBB", w, h, 8, 2, 0, 0, 0))
           + chunk(b"IDAT", zlib.compress(raw, 9))
           + chunk(b"IEND", b""))
    path.write_bytes(png)
    return path


def _add_bookmark(paragraph, name: str, bid: int) -> None:
    start = OxmlElement("w:bookmarkStart")
    start.set(qn("w:id"), str(bid))
    start.set(qn("w:name"), name)
    end = OxmlElement("w:bookmarkEnd")
    end.set(qn("w:id"), str(bid))
    paragraph._p.insert(0, start)
    paragraph._p.append(end)


def _add_field(paragraph, instr: str = "PAGE") -> None:
    """塞一个完整的 w:fldChar begin/instrText/separate/结果/end 域。"""
    r1 = OxmlElement("w:r")
    fc1 = OxmlElement("w:fldChar")
    fc1.set(qn("w:fldCharType"), "begin")
    r1.append(fc1)

    r2 = OxmlElement("w:r")
    it = OxmlElement("w:instrText")
    it.set(qn("xml:space"), "preserve")
    it.text = f" {instr} "
    r2.append(it)

    r3 = OxmlElement("w:r")
    fc3 = OxmlElement("w:fldChar")
    fc3.set(qn("w:fldCharType"), "separate")
    r3.append(fc3)

    r4 = OxmlElement("w:r")
    t = OxmlElement("w:t")
    t.text = "1"
    r4.append(t)

    r5 = OxmlElement("w:r")
    fc5 = OxmlElement("w:fldChar")
    fc5.set(qn("w:fldCharType"), "end")
    r5.append(fc5)

    for r in (r1, r2, r3, r4, r5):
        paragraph._p.append(r)


# ─────────────────────────── 主 fixture docx ───────────────────────────

def build_docx(dst: Path, img_dir: Path | None = None) -> Path:
    """造主 fixture docx。内容固定 —— 固定是特性，rc 才可断言。"""
    img_dir = img_dir or dst.parent
    img_dir.mkdir(parents=True, exist_ok=True)
    img1 = make_png(img_dir / "fig1.png", 160, 120, (66, 133, 244))
    img2 = make_png(img_dir / "fig2.png", 160, 120, (219, 68, 55))
    img3 = img_dir / "fig3.png"
    img3.write_bytes(img1.read_bytes())      # 逐字节重复 → 供 seqdiff image 去重

    doc = Document()

    p = doc.add_paragraph("某某水利工程可行性研究报告")
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER

    h1 = doc.add_heading("第1章 概述", level=1)
    _add_bookmark(h1, "_Toc_chapter1", 1)
    doc.add_paragraph(
        "本报告依据《水利水电工程可行性研究报告编制规程》(SL/T 618-2021)编制。"
        "工程位于某省某市，总投资约 12000 万元，工期 24 个月。参见\"设计标准\"一节。")
    doc.add_heading("1.1 工程概况", level=2)
    doc.add_paragraph(
        "工程规模为中型，设计洪水标准 50 年一遇，校核 500 年一遇。渠道全长 12.5km，"
        "设计流量 8.5m3/s。温度控制在 25 摄氏度以内，间距 100 毫米。")
    doc.add_heading("1.1.1 地理位置", level=3)
    doc.add_paragraph("工程区地处东经 118°21′，北纬 32°05′，属亚热带季风气候区。")

    doc.add_picture(str(img1), width=Cm(8))
    doc.paragraphs[-1].alignment = WD_ALIGN_PARAGRAPH.CENTER
    cap = doc.add_paragraph("图1-1 工程地理位置示意图")
    cap.alignment = WD_ALIGN_PARAGRAPH.CENTER

    doc.add_paragraph("如图1-1所示，工程区交通便利。")
    fp_ref = doc.add_paragraph("详见第 ")
    _add_field(fp_ref, r"REF _Toc_chapter2 \h")
    fp_ref.add_run(" 章。")
    fp_toc = doc.add_paragraph()
    _add_field(fp_toc, r"TOC \o \"1-3\" \h \z \u")

    tcap = doc.add_paragraph("表1-1 工程特性表")
    tcap.alignment = WD_ALIGN_PARAGRAPH.CENTER
    tbl = doc.add_table(rows=4, cols=3)
    tbl.style = "Table Grid"
    for i, row in enumerate([["序号", "项目", "数值"],
                             ["1", "设计流量", "8.5 m3/s"],
                             ["2", "渠道长度", "12.5 km"],
                             ["3", "总投资", "12000 万元"]]):
        for j, v in enumerate(row):
            tbl.cell(i, j).text = v
    doc.add_paragraph("表1-1 列出了主要工程特性指标。")

    doc.add_section(WD_SECTION.NEW_PAGE)

    h2 = doc.add_heading("第2章 水文分析", level=1)
    _add_bookmark(h2, "_Toc_chapter2", 2)
    doc.add_paragraph("采用 1980-2020 年共 41 年实测资料进行频率分析。数据来源于某某水文站。")
    doc.add_heading("2.1 设计标准", level=2)
    doc.add_paragraph("按照规范要求，本工程防洪标准取 50 年一遇。")
    doc.add_heading("2.1.1 径流计算", level=3)
    doc.add_paragraph("多年平均径流量 3.2 亿 m3，径流系数 0.42。")
    doc.add_heading("2.1.1.1 参数率定", level=4)
    doc.add_paragraph("率定期 NSE=0.85，验证期 NSE=0.81。")

    doc.add_picture(str(img2), width=Cm(8))
    doc.paragraphs[-1].alignment = WD_ALIGN_PARAGRAPH.CENTER
    cap2 = doc.add_paragraph("图2-1 年径流过程线")
    cap2.alignment = WD_ALIGN_PARAGRAPH.CENTER

    tcap2 = doc.add_paragraph("表2-1 设计洪水成果表")
    tcap2.alignment = WD_ALIGN_PARAGRAPH.CENTER
    tbl2 = doc.add_table(rows=3, cols=4)
    tbl2.style = "Table Grid"
    for i, row in enumerate([["频率", "洪峰(m3/s)", "洪量(万m3)", "备注"],
                             ["2%", "520", "3200", "设计"],
                             ["0.2%", "780", "4800", "校核"]]):
        for j, v in enumerate(row):
            tbl2.cell(i, j).text = v

    doc.add_picture(str(img3), width=Cm(8))
    cap3 = doc.add_paragraph("图2-2 工程地理位置示意图（重复）")
    cap3.alignment = WD_ALIGN_PARAGRAPH.CENTER

    doc.add_heading("第3章 结论与建议", level=1)
    doc.add_paragraph("本工程技术可行、经济合理，建议尽快开展初步设计。"
                      "联系人：张三，电话 13800138000。")
    doc.add_paragraph("“引号测试”与'单引号'，以及 100mm、25℃ 等单位。")

    # 直接格式脏段：供 fix clear-direct-format / para scan-ppr / fonts normalize
    dirty = doc.add_paragraph()
    r = dirty.add_run("这一段带直接格式：加粗、变色、等线。")
    r.bold = True
    r.font.size = Pt(14)
    r.font.color.rgb = RGBColor(0xC0, 0x00, 0x00)
    r.font.name = "等线"
    r._element.rPr.rFonts.set(qn("w:eastAsia"), "等线")

    for sec in doc.sections:
        sec.header.is_linked_to_previous = False
        sec.footer.is_linked_to_previous = False
        hp = sec.header.paragraphs[0] if sec.header.paragraphs else sec.header.add_paragraph()
        hp.text = "某某水利工程可行性研究报告"
        hp.alignment = WD_ALIGN_PARAGRAPH.CENTER
        fp = sec.footer.paragraphs[0] if sec.footer.paragraphs else sec.footer.add_paragraph()
        fp.text = ""
        fp.alignment = WD_ALIGN_PARAGRAPH.CENTER
        _add_field(fp, "PAGE")

    dst.parent.mkdir(parents=True, exist_ok=True)
    doc.save(str(dst))
    return dst


def build_docx_v2(dst: Path, img_dir: Path | None = None) -> Path:
    """主 fixture 的"下一版"：改几句话 + 少一张图，供 compare / seqdiff / track 用。

    完全相同的两份文件会让 diff 类动词跑出「无差异」——那也是 rc=0，
    但它证明不了 diff 真的在比。所以这里必须真的不一样。
    """
    build_docx(dst, img_dir=img_dir)
    doc = Document(str(dst))
    for p in doc.paragraphs:
        if p.text.startswith("按照规范要求"):
            for r in p.runs:
                r.text = ""
            if p.runs:
                p.runs[0].text = "按照最新规范要求，本工程防洪标准调整为 100 年一遇。"
        elif p.text.startswith("率定期"):
            for r in p.runs:
                r.text = ""
            if p.runs:
                p.runs[0].text = "率定期 NSE=0.88，验证期 NSE=0.83（重新率定）。"
    doc.add_paragraph("本段是 v2 新增的补充说明，用于 sequence diff 验证新增段被识别。")
    doc.save(str(dst))
    return dst


def build_template_docx(dst: Path) -> Path:
    """一份"模板" docx：只有样式与页眉页脚，没有正文。

    `template` / `format apply` / `chrome` 这类动词默认会去读
    `templates/template.docx`（→ ~/Work 的 symlink）。smoke 必须显式传这一份把
    那条依赖断掉，否则测试既碰真实业务文件、又在别的机器上必然失败。
    """
    doc = Document()
    doc.add_heading("模板标题样式", level=1)
    doc.add_heading("模板二级样式", level=2)
    doc.add_paragraph("模板正文样式示例。")
    sec = doc.sections[0]
    sec.header.is_linked_to_previous = False
    sec.footer.is_linked_to_previous = False
    sec.header.paragraphs[0].text = "模板页眉"
    fp = sec.footer.paragraphs[0]
    fp.alignment = WD_ALIGN_PARAGRAPH.CENTER
    _add_field(fp, "PAGE")
    dst.parent.mkdir(parents=True, exist_ok=True)
    doc.save(str(dst))
    return dst


# ─────────────────────────── 配套的非 docx 造件 ───────────────────────────

FIXTURE_MD = """\
# 某某水利工程可行性研究报告

## 第1章 概述

本报告依据《水利水电工程可行性研究报告编制规程》编制。工程位于某省某市。

![图1-1 工程地理位置示意图](fig1.png)

| 序号 | 项目 | 数值 |
|---|---|---|
| 1 | 设计流量 | 8.5 m3/s |
| 2 | 渠道长度 | 12.5 km |

## 第2章 水文分析

采用 1980-2020 年共 41 年实测资料进行频率分析。

### 2.1 设计标准

按照规范要求，本工程防洪标准取 50 年一遇。
"""

FIXTURE_MD2 = """\
# 附录

## 第3章 结论与建议

本工程技术可行、经济合理，建议尽快开展初步设计。
"""

# `compare-ref ref` / `revise-rules gen` 认的是这个形状：
# `**改为**：` 后面紧跟一个 `>` 引用块（正则见 CLAUDE.md 那节）。
DRAFT_MD = """\
# 第1章改动草稿

## 意见 1

原文：按照规范要求，本工程防洪标准取 50 年一遇。

**改为**：

> 按照最新规范要求，本工程防洪标准调整为 100 年一遇，并补充校核工况说明。

## 意见 2

**改为**：

> 多年平均径流量 3.2 亿 m3，径流系数 0.42，另需补充蒸发折算系数。
"""

# `scan-sensitive` 的输入是**目录**，且目录里必须有 .md
BID_MD = """\
# 投标承诺

我方为该产品的唯一指定供应商，绝对保证工期最短，产品为国家级免检产品。
"""


def build_side_files(root: Path) -> dict:
    """造齐非 docx 的配套件，返回路径字典。"""
    root.mkdir(parents=True, exist_ok=True)

    md_dir = root / "md"
    md_dir.mkdir(exist_ok=True)
    (md_dir / "01-第1章.md").write_text(FIXTURE_MD, encoding="utf-8")
    (md_dir / "02-第3章.md").write_text(FIXTURE_MD2, encoding="utf-8")
    make_png(md_dir / "fig1.png")

    drafts = root / "drafts"
    drafts.mkdir(exist_ok=True)
    # ⚠ 文件名里那个短横不是随手加的：`compare-ref ref` 的默认 glob 是
    # `*改动草稿.md`，`revise-rules gen` 的默认 glob 是 `*-改动草稿.md`
    # （biddiff.py:399 vs :604 —— 两条默认值本身就不一致）。
    # `第1章-改动草稿.md` 是同时满足两者的唯一写法；去掉短横，revise-rules
    # 就变成「待处理 MD: 0 份」的空跑，而它照样 rc=0，测试看不出来。
    (drafts / "第1章-改动草稿.md").write_text(DRAFT_MD, encoding="utf-8")

    bid = root / "bid"
    bid.mkdir(exist_ok=True)
    (bid / "bid.md").write_text(BID_MD, encoding="utf-8")

    out_dir = root / "out"
    out_dir.mkdir(exist_ok=True)

    # `caption pair` 的 --decision 是 required，且过 schema 校验（v1）
    decision = root / "decision.json"
    decision.write_text(json.dumps({
        "version": "1",
        "source_docx": "fixture.docx",
        "ops": [{"op": "renumber-all-tables"}],
        "operations": [{"op": "renumber-all-tables"}],
    }, ensure_ascii=False, indent=1), encoding="utf-8")

    # `blocks relocate` 的 --plan 是 required；source_heading_text 必须在文档里找得到，
    # 否则该 move 被记成 skip（仍 rc=0，所以这里得用真实存在的标题文本）
    plan = root / "relocate-plan.json"
    plan.write_text(json.dumps({
        "version": "1",
        "source_docx": "fixture.docx",
        "moves": [{
            "source_idx": 19,
            "target_idx": 5,
            "source_heading_idx": 19,
            "source_block_end_idx": 19,
            "target_insert_after_idx": 5,
            "source_heading_text": "2.1.1 径流计算",
            "target_context_text": "1.1.1 地理位置",
            "reason": "smoke: 把 2.1.1 挪到 1.1.1 后面",
        }],
    }, ensure_ascii=False, indent=1), encoding="utf-8")

    return {
        "md_dir": md_dir,
        "md": md_dir / "01-第1章.md",
        "md2": md_dir / "02-第3章.md",
        "drafts": drafts,
        "bid_dir": bid,
        "out_dir": out_dir,
        "decision": decision,
        "relocate_plan": plan,
    }


if __name__ == "__main__":  # 手工造一份看看
    target = Path(sys.argv[1] if len(sys.argv) > 1 else "/tmp/doctools-smoke/probe")
    target.mkdir(parents=True, exist_ok=True)
    build_docx(target / "fixture.docx")
    build_docx_v2(target / "fixture-v2.docx")
    build_template_docx(target / "template.docx")
    build_side_files(target)
    print("OK ->", target)
