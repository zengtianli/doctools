"""template 套模板：源稿样式与模板同名不同号时按名称对齐，不能让别的样式顶号。

实例（2026-10-13 庆元托溪申报方案）：源稿 ZDWP0=表名、ZDWP1=表格内容；模板 ZDWP1=表名、
ZDWP3=表格内容。只按编号替换定义，98 个表格单元格段全部变成表名。
"""
import sys
import zipfile
from pathlib import Path

from lxml import etree

HERE = Path(__file__).resolve().parent
sys.path.insert(0, str(HERE.parent))
sys.path.insert(0, str(HERE.parents[2] / "lib"))

import docx_fmt  # noqa: E402

W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
NS = f'xmlns:w="{W}"'


def _style(sid, name, extra=""):
    return f'<w:style w:type="paragraph" w:styleId="{sid}"><w:name w:val="{name}"/>{extra}</w:style>'


def _docx(path, styles, paras):
    body = "".join(f'<w:p><w:pPr><w:pStyle w:val="{sid}"/></w:pPr><w:r><w:t>{text}</w:t></w:r></w:p>' for sid, text in paras)
    parts = {
        "[Content_Types].xml": '<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
        '<Default Extension="xml" ContentType="application/xml"/></Types>',
        "word/document.xml": f'<?xml version="1.0"?><w:document {NS}><w:body>{body}</w:body></w:document>',
        "word/styles.xml": f'<?xml version="1.0"?><w:styles {NS}>{"".join(styles)}</w:styles>',
    }
    with zipfile.ZipFile(path, "w") as zf:
        for name, data in parts.items():
            zf.writestr(name, data)


def _read(path):
    with zipfile.ZipFile(path) as zf:
        doc = etree.fromstring(zf.read("word/document.xml"))
        styles = etree.fromstring(zf.read("word/styles.xml"))
    names = {s.get(f"{{{W}}}styleId"): s.find(f"{{{W}}}name").get(f"{{{W}}}val") for s in styles.iter(f"{{{W}}}style")}
    paras = [(names[p.find(f".//{{{W}}}pStyle").get(f"{{{W}}}val")], "".join(t.text for t in p.iter(f"{{{W}}}t")))
             for p in doc.iter(f"{{{W}}}p")]
    return names, paras


def test_same_name_different_id_follows_template(tmp_path):
    src, tmpl, out = tmp_path / "src.docx", tmp_path / "tmpl.xml", tmp_path / "out.docx"
    _docx(src,
          [_style("ZDWP", "ZDWP正文"), _style("ZDWP0", "ZDWP 表名"),
           _style("ZDWP1", "ZDWP表格内容", '<w:basedOn w:val="ZDWP0"/>'), _style("Mine", "我的样式")],
          [("ZDWP", "正文"), ("ZDWP0", "表 1"), ("ZDWP1", "单元格"), ("Mine", "别的")])
    tmpl.write_text(f'<?xml version="1.0"?><w:styles {NS}>'
                    + _style("ZDWP", "ZDWP正文") + _style("ZDWP1", "ZDWP 表名") + _style("ZDWP3", "ZDWP表格内容")
                    + "</w:styles>", encoding="utf-8")
    docx_fmt.apply_styles_to_docx(str(src), str(tmpl), str(out))
    names, paras = _read(out)
    assert paras == [("ZDWP正文", "正文"), ("ZDWP 表名", "表 1"), ("ZDWP表格内容", "单元格"), ("我的样式", "别的")]
    assert sorted(names.values()) == sorted(set(names.values())), "不能留下重名样式"
    assert names["ZDWP1"] == "ZDWP 表名" and names["ZDWP3"] == "ZDWP表格内容"


def test_source_style_squatting_template_id_is_kept(tmp_path):
    src, tmpl, out = tmp_path / "src.docx", tmp_path / "tmpl.xml", tmp_path / "out.docx"
    _docx(src, [_style("ZDWP1", "业主自定义样式")], [("ZDWP1", "保留")])
    tmpl.write_text(f'<?xml version="1.0"?><w:styles {NS}>' + _style("ZDWP1", "ZDWP 表名") + "</w:styles>", encoding="utf-8")
    docx_fmt.apply_styles_to_docx(str(src), str(tmpl), str(out))
    _, paras = _read(out)
    assert paras == [("业主自定义样式", "保留")]


def test_matching_ids_untouched(tmp_path):
    src, tmpl, out = tmp_path / "src.docx", tmp_path / "tmpl.xml", tmp_path / "out.docx"
    _docx(src, [_style("ZDWP", "ZDWP正文")], [("ZDWP", "正文")])
    tmpl.write_text(f'<?xml version="1.0"?><w:styles {NS}>' + _style("ZDWP", "ZDWP正文") + "</w:styles>", encoding="utf-8")
    docx_fmt.apply_styles_to_docx(str(src), str(tmpl), str(out))
    with zipfile.ZipFile(src) as a, zipfile.ZipFile(out) as b:
        assert a.read("word/document.xml") == b.read("word/document.xml")
