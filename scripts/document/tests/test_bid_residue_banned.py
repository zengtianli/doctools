"""bid_residue_lib 项目禁词 banned_terms —— 2026-09-18 海宁标实证。

领导退回第 6 章：正文 42 处“招标文件/采购人/评分”等词。原有扫描只在 pei 模式查身份词，
main 实名稿没有项目级禁词入口，退稿版扫描全零。

判据：
  1. 规则给了 banned_terms，main 模式也报出命中段（类别 3）
  2. 不给规则时行为不变（不凭空禁词，pei 稿自称“投标人”仍合法）
"""
from __future__ import annotations

import sys
from pathlib import Path

HERE = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(HERE))
import bid_residue_lib as L  # noqa: E402

W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'


def parts(*texts):
    body = "".join(f"<w:p><w:r><w:t>{t}</w:t></w:r></w:p>" for t in texts)
    return {"word/document.xml": f"<w:document {W}><w:body>{body}</w:body></w:document>".encode()}


def test_banned_terms_hit_in_main():
    rules = L.load_rules(None)
    rules["banned_terms"] = ["招标文件", "采购人"]
    f = L.scan_parts(parts("按招标文件要求开展调查。", "本项目服务期为一年。"), mode="main", rules=rules)
    assert [(x["cat"], x["para"]) for x in f] == [(3, 1)]
    assert "禁词:招标文件" in f[0]["marks"]


def test_no_rules_no_ban():
    f = L.scan_parts(parts("投标人按招标文件要求开展调查。"), mode="main")
    assert f == []
