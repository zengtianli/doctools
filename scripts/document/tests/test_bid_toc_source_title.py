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
    else:
        assert '原评分项自动比对：scoring.json' in result.stdout
        assert '本次未自动核对招标原文' in result.stdout
        assert '[PASS] 招标表 ≡' not in result.stdout


@pytest.mark.parametrize('mismatched', [False, True])
def test_transcribed_source_requires_full_name_match(tmp_path, mismatched):
    (tmp_path / '_project.yaml').write_text('{}\n')
    tender = tmp_path / '招标文件'
    tender.mkdir()
    names = ['企业证书', '项目业绩', '项目负责人', '人员配置', '售后服务方案']
    (tender / '评分表.md').write_text('\n'.join(f'| {i}、{name} |' for i, name in enumerate(names, 1)))
    chapters = [{'item_no': i, 'title': name, 'score': 6, 'file': f'ch{i}.md'}
                for i, name in enumerate(names, 1)]
    chapters[-1].update(title=EXPANDED, source_title='售后服务其他要求' if mismatched else names[-1])
    (tmp_path / 'chapters.yaml').write_text(yaml.safe_dump({'chapters': chapters}, allow_unicode=True))
    for c in chapters:
        (tmp_path / c['file']).write_text(f"# {c['item_no']} {c['title']}\n")
    result = subprocess.run([sys.executable, str(GATE), str(tmp_path)], capture_output=True, text=True)
    assert result.returncode == (2 if mismatched else 0), result.stdout + result.stderr
    if mismatched:
        assert '≠ 评分表' in result.stdout
    else:
        assert '原评分项自动比对：招标转录评分表' in result.stdout
        assert '本次未自动核对招标原文' not in result.stdout


def test_empty_source_title_is_a_configuration_error(tmp_path):
    (tmp_path / '_project.yaml').write_text('{}\n')
    (tmp_path / 'chapters.yaml').write_text(yaml.safe_dump({'chapters': [
        {'item_no': 7, 'title': EXPANDED, 'source_title': '', 'score': 6}]}))
    result = subprocess.run([sys.executable, str(GATE), str(tmp_path)], capture_output=True, text=True)
    assert result.returncode == 2
    assert '标题配置不合法' in result.stdout
    assert 'Traceback' not in result.stderr
