"""health：扁平表号「表 1 ××」紧贴表上方算配好的表题，章节式「表3-1」不归这条管。

实例（2026-10-13 庆元托溪申报方案）：4 张表都是「表 N ××」，gate 把它们全报成孤儿表、
编号集 0 ≠ 表题段 4。
"""
import sys
import zipfile
from pathlib import Path

HERE = Path(__file__).resolve().parent
sys.path.insert(0, str(HERE.parents[1]))
sys.path.insert(0, str(HERE.parents[3] / "lib"))

from sub import health  # noqa: E402

W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"


def _docx(path, blocks):
    body = "".join(
        "<w:tbl><w:tr><w:tc><w:p><w:r><w:t>格</w:t></w:r></w:p></w:tc></w:tr></w:tbl>" if b is None
        else f"<w:p><w:r><w:t>{b}</w:t></w:r></w:p>" for b in blocks)
    with zipfile.ZipFile(path, "w") as zf:
        zf.writestr("word/document.xml", f'<?xml version="1.0"?><w:document xmlns:w="{W}"><w:body>{body}</w:body></w:document>')


def test_flat_captions_found_above_tables(tmp_path):
    f = tmp_path / "a.docx"
    _docx(f, ["引出段，见表 1。", "表 1 管理断面", "", None, "见表 2。", "表 2 实施计划", None, None])
    assert health._flat_table_captions(f) == {1: 1, 2: 2}


def test_sectioned_and_inline_are_not_flat(tmp_path):
    f = tmp_path / "b.docx"
    _docx(f, ["表3-1 章节式", None, "表 3.1-2 章节式", None, "表 4 的数据见下", "正文", None])
    assert health._flat_table_captions(f) == {}


def test_flat_regex_boundaries():
    m = health._FLAT_TBL_CAPTION.match
    assert m("表 1 托溪流域庆元县生态流量管理断面").group(1) == "1"
    assert m("表12　投资估算").group(1) == "12"
    assert not m("表1-2 分项")
    assert not m("表1.2 分项")
    assert not m("附表 1 清单")
