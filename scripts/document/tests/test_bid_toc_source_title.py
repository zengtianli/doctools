"""Expanded body titles must retain an independently checked tender item name."""
import json
from pathlib import Path
import subprocess
import sys

import pytest
import yaml


GATE = Path(__file__).resolve().parents[1] / 'bid_toc_gate.py'
EXPANDED = '售后服务方案、售后服务承诺及服务承诺落实的保障措施'


@pytest.mark.parametrize('source_title,body_title,score,want', [
    ('售后服务方案', EXPANDED, 6, 0),
    (None, EXPANDED, 6, 2),
    ('错误的原评分项', EXPANDED, 6, 2),
    ('售后服务方案', '未经登记的标题', 6, 2),
    ('售后服务方案', EXPANDED, 5, 2),
])
def test_source_and_body_titles_are_checked_separately(tmp_path, source_title, body_title, score, want):
    (tmp_path / '_project.yaml').write_text('{}\n')
    tender = tmp_path / '招标文件'
    tender.mkdir()
    (tender / 'scoring.json').write_text(json.dumps({'tech_business': {
        'score': 6, 'items': [{'no': 7, 'name': '售后服务方案', 'max': 6}]}}))
    chapter = {'item_no': 7, 'title': EXPANDED, 'score': score, 'file': 'chapter.md'}
    if source_title is not None:
        chapter['source_title'] = source_title
    (tmp_path / 'chapters.yaml').write_text(yaml.safe_dump({'chapters': [chapter]}, allow_unicode=True))
    (tmp_path / 'chapter.md').write_text('# 第七章 ' + body_title + '\n\n服务内容。\n')
    result = subprocess.run([sys.executable, str(GATE), str(tmp_path)], capture_output=True, text=True)
    assert result.returncode == want, result.stdout + result.stderr
    if want == 2:
        assert '≠' in result.stdout, result.stdout
