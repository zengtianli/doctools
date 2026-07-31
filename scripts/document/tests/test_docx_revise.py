"""docx_revise 引擎回归测试 —— 把 2026-07-31 对真件跑过的反向验证固化下来。

覆盖：四种 action、锚点唯一性 fail-closed（含 within 消歧）、目录段跳过、
中段替换 head/tail 保留、批注部件注册、replace_para 两种形态。
夹具是最小合法 docx（引擎只读 document.xml / [Content_Types].xml / rels）。
"""
import sys
import zipfile
from pathlib import Path

import pytest
from lxml import etree

sys.path.append(str(Path(__file__).resolve().parents[3] / 'lib'))
from docx_revise import ReviseError, apply_ops  # noqa: E402

W = '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'
WNS = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'

META = {'author': '测试者', 'initials': 'T', 'date': '2026-07-31T00:00:00Z'}


def _p(text, style=None):
    st = f'<w:pPr><w:pStyle w:val="{style}"/></w:pPr>' if style else ''
    return (f'<w:p>{st}<w:r><w:rPr><w:sz w:val="24"/></w:rPr>'
            f'<w:t xml:space="preserve">{text}</w:t></w:r></w:p>')


def _mini_docx(path: Path):
    doc = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
           f'<w:document xmlns:w="{WNS}"><w:body>'
           + _p('4.2 方案')                       # 章标题(前)
           + _p('在现状的基础上推进改造')           # 重复文本 · 第 1 处(§4.2)
           + _p('5.1 结论', style='2')
           + _p('5.1 结论……页码', style='TOC2')   # 伪目录项：必须被跳过
           + _p('在现状的基础上推进改造')           # 重复文本 · 第 2 处(§5.1)
           + _p('规范编号 SL/T 835—2024 施行。')
           + _p('〔待补：占位段〕')
           + _p('5.2 展望', style='2')
           + '</w:body></w:document>')
    ct = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
          '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
          '<Default Extension="rels" ContentType="application/vnd.openxmlformats-'
          'package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/>'
          '<Override PartName="/word/document.xml" ContentType="application/vnd.openxml'
          'formats-officedocument.wordprocessingml.document.main+xml"/></Types>')
    rels = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/'
            'relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats'
            '.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>'
            '</Relationships>')
    with zipfile.ZipFile(path, 'w') as z:
        z.writestr('[Content_Types].xml', ct)
        z.writestr('_rels/.rels', '<?xml version="1.0"?><Relationships xmlns='
                   '"http://schemas.openxmlformats.org/package/2006/relationships"/>')
        z.writestr('word/document.xml', doc)
        z.writestr('word/_rels/document.xml.rels', rels)
        z.writestr('word/styles.xml', f'<w:styles xmlns:w="{WNS}"/>')


@pytest.fixture()
def src(tmp_path):
    p = tmp_path / 'src.docx'
    _mini_docx(p)
    return p


def _body_text(path):
    with zipfile.ZipFile(path) as z:
        d = etree.fromstring(z.read('word/document.xml'))
    return ''.join(t.text or '' for t in d.iter(W + 't'))


def _cfg(*ops):
    return {**META, 'ops': list(ops)}


def test_ambiguous_anchor_fails_and_within_disambiguates(src, tmp_path):
    out = tmp_path / 'o.docx'
    with pytest.raises(ReviseError):
        apply_ops(src, out, _cfg({'id': 'x', 'find': '在现状的基础上',
                                  'action': 'comment', 'comment': 'c'}))
    stats = apply_ops(src, out, _cfg(
        {'id': 'x', 'find': '在现状的基础上', 'within': '5.1',
         'action': 'insert_after', 'paras': ['新段落甲'],
         'comment': '批注甲', 'assert_between': ['5.1', '5.2']}))
    assert stats['ins_para'] == 1 and stats['comments'] == 1
    with zipfile.ZipFile(out) as z:
        assert 'word/comments.xml' in z.namelist()
        assert 'comments+xml' in z.read('[Content_Types].xml').decode()
        assert 'comments.xml' in z.read('word/_rels/document.xml.rels').decode()


def test_within_skips_toc(src, tmp_path):
    # within '5.1 结论' 有正文标题与 TOC 伪目录项两个候选——必须只认正文那个，
    # 且检索起点落在正文标题后(命中第 2 处重复文本而不是报歧义)
    stats = apply_ops(src, tmp_path / 'o.docx', _cfg(
        {'id': 'x', 'find': '在现状的基础上', 'within': '5.1 结论',
         'action': 'comment', 'comment': 'c'}))
    assert stats['comments'] == 1


def test_mid_replace_keeps_head_and_tail(src, tmp_path):
    out = tmp_path / 'o.docx'
    apply_ops(src, out, _cfg({'id': 'x', 'find': '规范编号', 'action': 'replace',
                              'old': '835', 'new': '835X'}))
    # 接受修订后的可见文本：head/tail 原样，仅目标片段换掉
    assert '规范编号 SL/T 835X—2024 施行。' in _body_text(out)
    with zipfile.ZipFile(out) as z:
        d = z.read('word/document.xml').decode()
    assert '<w:delText xml:space="preserve">835</w:delText>' in d


def test_replace_missing_old_fails(src, tmp_path):
    with pytest.raises(ReviseError):
        apply_ops(src, tmp_path / 'o.docx',
                  _cfg({'id': 'x', 'find': '规范编号', 'action': 'replace',
                        'old': '不存在', 'new': 'y'}))


def test_replace_para_tracked_and_untracked(src, tmp_path):
    out1 = tmp_path / 'o1.docx'
    apply_ops(src, out1, _cfg({'id': 'x', 'find': '〔待补：占位段〕',
                               'action': 'replace_para', 'paras': ['真内容']}))
    with zipfile.ZipFile(out1) as z:
        d = z.read('word/document.xml').decode()
    assert '真内容' in d and '<w:delText' in d          # tracked：旧段留 w:del 痕

    out2 = tmp_path / 'o2.docx'
    apply_ops(src, out2, _cfg({'id': 'x', 'find': '〔待补：占位段〕',
                               'action': 'replace_para', 'tracked': False,
                               'paras': ['真内容']}))
    assert '占位段' not in _body_text(out2)             # untracked：占位段直接消失


def test_dry_run_writes_nothing(src, tmp_path):
    out = tmp_path / 'o.docx'
    stats = apply_ops(src, out, _cfg({'id': 'x', 'find': '规范编号',
                                      'action': 'comment', 'comment': 'c'}), dry=True)
    assert stats.get('dry_run') and not out.exists()
