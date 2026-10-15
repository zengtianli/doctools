"""soffice.docx_to_pdf：字体名归一、被替换字体的判定；有 soffice 和办公字体时实转一份。"""
import shutil
import sys
import zipfile
from pathlib import Path

import pytest

sys.path.insert(0, str(Path(__file__).resolve().parents[3] / "lib"))

import soffice  # noqa: E402

W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"


def _docx(path, font="宋体", text="生态流量 0.14 m³/s"):
    body = (f'<w:p><w:r><w:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman" w:eastAsia="{font}"/></w:rPr>'
            f"<w:t>{text}</w:t></w:r></w:p>")
    with zipfile.ZipFile(path, "w") as zf:
        zf.writestr("[Content_Types].xml",
                    '<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
                    '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
                    '<Default Extension="xml" ContentType="application/xml"/>'
                    '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>')
        zf.writestr("_rels/.rels",
                    '<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
                    '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>')
        zf.writestr("word/document.xml", f'<?xml version="1.0"?><w:document xmlns:w="{W}"><w:body>{body}</w:body></w:document>')


def test_font_names_normalize_to_same_key():
    n = soffice._norm_font
    assert n("BAAAAA+SimSun") == n("宋体") == n("SimSun")
    assert n("HAAAAA+TimesNewRomanPS-BoldMT") == n("Times New Roman")
    assert n("CAAAAA+SimHei") == n("黑体")
    assert n("PingFangSC-Regular") != n("宋体")


def test_declared_fonts_read_from_docx(tmp_path):
    f = tmp_path / "a.docx"
    _docx(f)
    assert {soffice._norm_font("宋体"), soffice._norm_font("Times New Roman")} <= soffice._docx_declared_fonts(f)


def test_existing_target_is_not_overwritten(tmp_path):
    f, out = tmp_path / "a.docx", tmp_path / "a.pdf"
    _docx(f)
    out.write_bytes(b"keep")
    with pytest.raises(FileExistsError):
        soffice.docx_to_pdf(f, out)
    assert out.read_bytes() == b"keep"


@pytest.mark.skipif(not (soffice.find_soffice() and soffice.office_font_dirs() and shutil.which("pdffonts")),
                    reason="需要 soffice、Word/WPS 自带字体和 pdffonts")
def test_real_conversion_keeps_fonts(tmp_path):
    f, out = tmp_path / "a.docx", tmp_path / "out" / "a.pdf"
    _docx(f)
    info = soffice.docx_to_pdf(f, out, profile=tmp_path / "profile")
    assert out.is_file() and info["pages"] == 1
    assert info["unexpected_fonts"] == [], info
    assert any("SimSun" in x for x in info["fonts"])
