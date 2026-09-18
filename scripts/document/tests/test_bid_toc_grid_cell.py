"""pandoc 网格表里评分项名跨行时，tender_scoring_items 要按同列拼回完整名称。"""
from pathlib import Path
import sys

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
import bid_toc_gate  # noqa: E402

SEP = '+------+--------------------+------------------+-------+'
ROWS = [
    ('1、企业综合实力（4分）', '说明', '0-4分'),
    ('2、投标人类似业绩（2分）', '说明', '0-2分'),
    ('3、项目组人员配备（13分）', '说明', '0-13分'),
    None,  # 第 4 项跨行
    ('5、项目实施方案（20分）', '说明', '0-20分'),
    ('6、合理化建议', '只有首行、同列续行是长段说明', '0-4分'),
]


def _table():
    out = [SEP]
    for r in ROWS:
        if r is None:
            out += [
                '|      | 4、项目组成员      | 团队实力酌情评分 | 0-4分 |',
                '|      |                    |                  |       |',
                '|      | 组织方案（4分）    | 团队成员配备专业 |       |',
            ]
        elif r[0].startswith('6、'):
            out += [
                '|      | 6、合理化建议      | 酌情评分         | 0-4分 |',
                '|      | 建议可操作性强的得 | 满分             |       |',
            ]
        else:
            out.append('|      | %s | %s | %s |' % r)
        out.append(SEP)
    return '\n'.join(out) + '\n'


def test_multiline_cell_name_and_score(tmp_path):
    (tmp_path / '招标文件').mkdir()
    (tmp_path / '招标文件' / 't.md').write_text(_table(), encoding='utf-8')
    items, src = bid_toc_gate.tender_scoring_items(tmp_path)
    assert src == 't.md'
    got = {no: (name, score) for no, name, score in items}
    assert got[4] == ('项目组成员组织方案', 4)
    assert got[1] == ('企业综合实力', 4)
    assert got[5] == ('项目实施方案', 20)
    # 续行凑不出「（N分）」时保留首行名，不拼进同列说明
    assert got[6] == ('合理化建议', None)


def test_heads_from_docx_resolves_h1_by_style_name(tmp_path):
    """一级标题 styleId 不是固定写法（WPS 常见 "10"），按样式名 heading 1 认出。"""
    import zipfile
    W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'
    doc = (f'<w:document {W}><w:body><w:p><w:pPr><w:pStyle w:val="10"/></w:pPr>'
           '<w:r><w:t>7 售后服务方案</w:t></w:r></w:p></w:body></w:document>')
    sty = f'<w:styles {W}><w:style w:type="paragraph" w:styleId="10"><w:name w:val="heading 1"/></w:style></w:styles>'
    p = tmp_path / "x.docx"
    with zipfile.ZipFile(p, "w") as z:
        z.writestr("word/document.xml", doc)
        z.writestr("word/styles.xml", sty)
    assert [h[:2] for h in bid_toc_gate.heads_from_docx(p)] == [("7", "售后服务方案")]
