"""health.check_numbering_depth_uniformity 回归 —— 2026-09-18 海宁标第 6 章实证。

ztl 章节模板的标题样式 ID 是 1/21/31/4（样式名 heading 1/2/3/4）。原实现只认固定 ID，
21/31 两级被当成非标题，四级齐全的文档被判「缺中间级 [2, 3]」，health gate 误报 FAIL。

判据：
  1. 样式 ID 为 21/31、样式名为 heading 2/3 的四级齐全文档不报跳级
  2. 真正缺中间级（1 级后直接 1.1.1.1）仍然报出
"""
from __future__ import annotations

import sys
import zipfile
from pathlib import Path

HERE = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(HERE))
from sub import health as H  # noqa: E402

W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'
STYLES = (f'<w:styles {W}>'
          + ''.join(f'<w:style w:type="paragraph" w:styleId="{sid}"><w:name w:val="heading {lvl}"/></w:style>'
                    for sid, lvl in (("1", 1), ("21", 2), ("31", 3), ("4", 4)))
          + '</w:styles>')


def make(tmp_path, heads):
    body = ''.join(f'<w:p><w:pPr><w:pStyle w:val="{sid}"/></w:pPr><w:r><w:t>{t}</w:t></w:r></w:p>'
                   for sid, t in heads)
    p = tmp_path / "x.docx"
    with zipfile.ZipFile(p, "w") as z:
        z.writestr("word/document.xml", f'<w:document {W}><w:body>{body}</w:body></w:document>')
        z.writestr("word/styles.xml", STYLES)
    return p


def test_named_style_ids_resolve(tmp_path):
    p = make(tmp_path, [("1", "6 项目理解"), ("21", "6.1 概况"), ("31", "6.1.1 区位"),
                        ("4", "6.1.1.1 地形"), ("4", "6.1.1.2 气候")])
    r = H.check_numbering_depth_uniformity(p)
    assert r["found"] is False and r["missing_mid_depths"] == []


def test_real_skip_still_flagged(tmp_path):
    p = make(tmp_path, [("1", "6 项目理解"), ("4", "6.1.1.1 地形"), ("4", "6.1.1.2 气候")])
    r = H.check_numbering_depth_uniformity(p)
    assert r["found"] is True and r["missing_mid_depths"] == [2, 3]
