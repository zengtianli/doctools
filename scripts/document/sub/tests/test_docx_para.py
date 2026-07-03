"""tests for docx_para.py — 段落级 查-改-验 工作台。

手搓最小 docx(zipfile 直写 XML)测:
  · locate 归一化命中(防 run-split 假阴性) / 多命中 / 无命中
  · edit 跨 run 替换 / 跨结构边界拒动 / --expect 防呆
  · fix-ppr 剥 sectPr
  · surgical_rewrite verbatim(媒体 CRC 逐项相等)
  · scan-ppr 排除表格单元格段
  · word_lock_file 锁检测
"""

from __future__ import annotations

import importlib.util
import sys
import zipfile
from pathlib import Path

import pytest

_HERE = Path(__file__).resolve().parent
_SUB = _HERE.parent

# 直接按路径载入被测模块(不需 package 上下文);载入时其顶部会把 lib 挂上 sys.path 并 import docx_surgical
spec = importlib.util.spec_from_file_location("docx_para_uut", str(_SUB / "docx_para.py"))
dp = importlib.util.module_from_spec(spec)
sys.modules["docx_para_uut"] = dp
assert spec.loader is not None
spec.loader.exec_module(dp)
ds = dp.ds  # docx_surgical

W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"

_CT = (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
    '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
    '<Default Extension="xml" ContentType="application/xml"/>'
    '<Default Extension="png" ContentType="image/png"/>'
    '<Override PartName="/word/document.xml" '
    'ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>'
    "</Types>"
)
_RELS = (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
    '<Relationship Id="rId1" '
    'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" '
    'Target="word/document.xml"/></Relationships>'
)
_NS = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'


def _doc(body_inner: str) -> bytes:
    return (
        f"<?xml version='1.0' encoding='UTF-8' standalone='yes'?>"
        f"<w:document {_NS}><w:body>{body_inner}</w:body></w:document>"
    ).encode()


def _r(text: str, rpr: str = "") -> str:
    return f'<w:r>{rpr}<w:t xml:space="preserve">{text}</w:t></w:r>'


def _p(runs: str, ppr: str = "") -> str:
    return f"<w:p>{ppr}{runs}</w:p>"


def _ppr_style(style: str, ind: str = "") -> str:
    return f'<w:pPr><w:pStyle w:val="{style}"/>{ind}</w:pPr>'


def _make_docx(path: Path, document_xml: bytes, media: dict | None = None) -> None:
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as z:
        z.writestr("[Content_Types].xml", _CT)
        z.writestr("_rels/.rels", _RELS)
        z.writestr("word/document.xml", document_xml)
        for name, data in (media or {}).items():
            z.writestr(name, data)


def _reopen_paras(path: Path):
    root = ds.parse_document(path)
    return root, ds.iter_paras(root)


# ─── locate ──────────────────────────────────────────────────────────────
def test_locate_run_split_normalization(tmp_path, capsys):
    """查询串跨 run(生态流量|核定 分两 run)仍命中 — 归一化拼整段再匹配。"""
    d = tmp_path / "x.docx"
    _make_docx(
        d,
        _doc(
            _p(_r("无关段"), _ppr_style("Body"))
            + _p(_r("生态流量") + _r("核定"), _ppr_style("Body"))
            + _p(_r("尾段"), _ppr_style("Body"))
        ),
    )
    rc = dp.main(["locate", str(d), "生态流量核定"])
    out = capsys.readouterr().out
    assert rc == 0, out
    assert "★" in out and "PARA 1" in out


def test_locate_ambiguous(tmp_path, capsys):
    d = tmp_path / "x.docx"
    _make_docx(d, _doc(_p(_r("重复文本"), _ppr_style("Body")) + _p(_r("重复文本"), _ppr_style("Body"))))
    rc = dp.main(["locate", str(d), "重复文本"])
    assert rc == 3
    assert "AMBIGUOUS" in capsys.readouterr().out


def test_locate_miss(tmp_path, capsys):
    d = tmp_path / "x.docx"
    _make_docx(d, _doc(_p(_r("只有这段"), _ppr_style("Body"))))
    rc = dp.main(["locate", str(d), "根本不存在的文本"])
    assert rc == 1
    assert "未命中" in capsys.readouterr().out


# ─── edit ────────────────────────────────────────────────────────────────
def test_edit_cross_run_replace(tmp_path, capsys):
    """替换 'cde' 跨 abc|def 两 run — new 全落首命中 t, 后续片段清除。"""
    d = tmp_path / "x.docx"
    _make_docx(d, _doc(_p(_r("abc") + _r("def"), _ppr_style("Body"))))
    rc = dp.main(["edit", str(d), "--para", "0", "--replace", "cde", "XYZ", "--no-gate", "--no-backup"])
    assert rc == 0, capsys.readouterr().out
    _, paras = _reopen_paras(d)
    assert ds.para_text(paras[0]) == "abXYZf"


def test_edit_structural_boundary_refused(tmp_path, capsys):
    """命中范围跨 w:br → exit 2, 文件不动。"""
    d = tmp_path / "x.docx"
    body = _p(_r("abc") + "<w:r><w:br/></w:r>" + _r("def"), _ppr_style("Body"))
    _make_docx(d, _doc(body))
    before = d.read_bytes()
    rc = dp.main(["edit", str(d), "--para", "0", "--replace", "cd", "ZZ", "--no-gate"])
    assert rc == 2
    assert d.read_bytes() == before  # 未动


def test_edit_expect_mismatch(tmp_path):
    """--expect 不符 → exit 2, 文件不动。"""
    d = tmp_path / "x.docx"
    _make_docx(d, _doc(_p(_r("真实内容"), _ppr_style("Body"))))
    before = d.read_bytes()
    rc = dp.main(["edit", str(d), "--para", "0", "--replace", "真实", "改后", "--expect", "不存在的锚", "--no-gate"])
    assert rc == 2
    assert d.read_bytes() == before


def test_edit_index_out_of_range(tmp_path):
    d = tmp_path / "x.docx"
    _make_docx(d, _doc(_p(_r("只有一段"), _ppr_style("Body"))))
    rc = dp.main(["edit", str(d), "--para", "99", "--replace", "a", "b", "--no-gate"])
    assert rc == 2


# ─── fix-ppr ─────────────────────────────────────────────────────────────
def test_fix_ppr_strips_sectpr(tmp_path, capsys):
    """克隆含 sectPr 的源段 pPr → 目标段 pPr 不得含 sectPr(防节断悬空 rId)。"""
    d = tmp_path / "x.docx"
    src_ppr = (
        '<w:pPr><w:pStyle w:val="Body"/>'
        '<w:ind w:firstLine="480"/>'
        '<w:sectPr><w:pgSz w:w="11906" w:h="16838"/></w:sectPr></w:pPr>'
    )
    body = _p(_r("源段"), src_ppr) + _p(_r("目标段带脏格式"), _ppr_style("Body", '<w:ind w:firstLine="0"/>'))
    _make_docx(d, _doc(body))
    rc = dp.main(["fix-ppr", str(d), "--para", "1", "--clone-from", "0", "--no-gate", "--no-backup"])
    assert rc == 0, capsys.readouterr().out
    _, paras = _reopen_paras(d)
    tgt_ppr = paras[1].find(W + "pPr")
    assert tgt_ppr is not None
    assert tgt_ppr.find(W + "sectPr") is None, "sectPr 必须被剥掉"
    # 克隆到了源段的 ind(480), 覆盖了目标原来的脏 ind(0)
    ind = tgt_ppr.find(W + "ind")
    assert ind is not None and ind.get(W + "firstLine") == "480"


def test_fix_ppr_style_mismatch_note(tmp_path, capsys):
    """克隆源样式 ≠ 目标样式 → 打 NOTE 提醒会改样式。"""
    d = tmp_path / "x.docx"
    body = _p(_r("标题段"), _ppr_style("Heading")) + _p(_r("正文段"), _ppr_style("Body"))
    _make_docx(d, _doc(body))
    rc = dp.main(["fix-ppr", str(d), "--para", "1", "--clone-from", "0", "--no-gate", "--no-backup"])
    out = capsys.readouterr().out
    assert rc == 0
    assert "NOTE" in out and "≠" in out


# ─── surgical_rewrite verbatim ───────────────────────────────────────────
def test_surgical_rewrite_verbatim_media_crc(tmp_path):
    """只改 document.xml,媒体等其余 zip 项 (filename, CRC, size) 逐项相等。"""
    d = tmp_path / "x.docx"
    media = {"word/media/img1.png": b"\x89PNG\r\n\x1a\n" + b"FAKEIMAGEDATA" * 500}
    _make_docx(d, _doc(_p(_r("原文"), _ppr_style("Body"))), media=media)

    def infomap(z):
        return {i.filename: (i.CRC, i.file_size) for i in zipfile.ZipFile(z).infolist()}

    before = infomap(d)
    root = ds.parse_document(d)
    paras = ds.iter_paras(root)
    next(paras[0].iter(W + "t")).text = "改后"  # 改 document.xml (w:t 嵌在 w:r 内)
    ds.surgical_rewrite(d, ds.serialize(root), backup=False)
    after = infomap(d)
    assert set(before) == set(after)
    diff = [fn for fn in before if before[fn] != after[fn]]
    assert diff == ["word/document.xml"], f"只应 document.xml 变, 实际 {diff}"
    # 重开自检(verify_repacked 已在 surgical_rewrite 内跑过, 这里再确认可 parse)
    _, paras2 = _reopen_paras(d)
    assert ds.para_text(paras2[0]) == "改后"


def test_surgical_rewrite_rollback_on_corrupt(tmp_path):
    """写入坏 XML → verify 失败 → 从 bak 回滚 + raise RepackError。"""
    d = tmp_path / "x.docx"
    _make_docx(d, _doc(_p(_r("原文"), _ppr_style("Body"))))
    before = d.read_bytes()
    with pytest.raises(ds.RepackError):
        ds.surgical_rewrite(d, b"<w:document>NOT CLOSED", backup=True)
    assert d.read_bytes() == before, "回滚后应与改前逐字节相等"


# ─── scan-ppr ────────────────────────────────────────────────────────────
def test_scan_ppr_excludes_table_cells(tmp_path, capsys):
    """表格单元格段的直接 ind 不算正文流脏格式;正文流的脏 ind 才报。"""
    d = tmp_path / "x.docx"
    clean = "".join(_p(_r(f"正文第{i}段"), _ppr_style("Body")) for i in range(6))
    dirty_body = _p(_r("正文脏段"), _ppr_style("Body", '<w:ind w:firstLine="0"/>'))
    tbl = (
        "<w:tbl><w:tr><w:tc>"
        + _p(_r("表内单元格"), _ppr_style("Body", '<w:ind w:firstLine="0"/>'))
        + "</w:tc></w:tr></w:tbl>"
    )
    _make_docx(d, _doc(clean + dirty_body + tbl))
    rc = dp.main(["scan-ppr", str(d)])
    out = capsys.readouterr().out
    assert rc == 3, out
    assert "正文脏段" in out
    assert "表内单元格" not in out, "表格单元格段不应被 scan-ppr 报出"


# ─── word lock ───────────────────────────────────────────────────────────
def test_word_lock_detected_blocks_mutation(tmp_path):
    """存在 ~$ owner 文件 → mutating 命令 exit 2 且不动文件(无 --force)。"""
    d = tmp_path / "报告.docx"
    _make_docx(d, _doc(_p(_r("内容"), _ppr_style("Body"))))
    lock = tmp_path / ("~$" + d.name)
    lock.write_bytes(b"owner")
    assert ds.word_lock_file(d) == lock
    before = d.read_bytes()
    rc = dp.main(["fix-ppr", str(d), "--para", "0", "--clone-from", "0", "--no-gate"])
    assert rc == 2
    assert d.read_bytes() == before


def test_word_lock_force_overrides(tmp_path):
    """--force 越过锁 → 正常写(clone-from 自身即无变化, 但流程通过 exit 0)。"""
    d = tmp_path / "报告.docx"
    _make_docx(
        d, _doc(_p(_r("内容"), _ppr_style("Body", '<w:ind w:firstLine="0"/>')) + _p(_r("干净段"), _ppr_style("Body")))
    )
    (tmp_path / ("~$" + d.name)).write_bytes(b"owner")
    rc = dp.main(["fix-ppr", str(d), "--para", "0", "--clone-from", "1", "--force", "--no-gate", "--no-backup"])
    assert rc == 0


# ─── missing file ────────────────────────────────────────────────────────
def test_missing_file(tmp_path):
    rc = dp.main(["locate", str(tmp_path / "nope.docx"), "x"])
    assert rc == 2


# ─── 对抗审查回归 (2026-07-03 review) ──────────────────────────────────────
_XMLSPACE = "{http://www.w3.org/XML/1998/namespace}space"


def test_edit_preserve_on_all_modified_nodes(tmp_path):
    """跨 run 替换后, 被清片段的后续 node 若残留边缘空白, 必须补 xml:space=preserve
    (否则 Word 吞空格) —— 不止首 node。"""
    d = tmp_path / "x.docx"
    # 两 run: "Hello" | "World Foo"; replace "HelloWorld"→"X" → node2 残 " Foo"(前导空格)
    _make_docx(d, _doc(_p(_r("Hello") + _r("World Foo"), _ppr_style("Body"))))
    rc = dp.main(["edit", str(d), "--para", "0", "--replace", "HelloWorld", "X", "--no-gate", "--no-backup"])
    assert rc == 0
    _, paras = _reopen_paras(d)
    assert ds.para_text(paras[0]) == "X Foo"
    ts = list(paras[0].iter(W + "t"))
    tail = [t for t in ts if (t.text or "").startswith(" ")]
    assert tail, "应有残留前导空格的 w:t"
    assert tail[0].get(_XMLSPACE) == "preserve", "残留边缘空白 node 必须带 xml:space=preserve"


def test_fix_ppr_preserves_target_own_sectpr(tmp_path):
    """目标段自带 sectPr(分节标记)必须在 pPr 克隆替换后保留 —— 否则节边界丢失(HIGH)。"""
    d = tmp_path / "x.docx"
    src = _p(_r("源段"), _ppr_style("Body"))
    tgt_ppr = (
        '<w:pPr><w:pStyle w:val="Body"/><w:ind w:firstLine="0"/>'
        '<w:sectPr><w:pgSz w:w="16838" w:h="11906" w:orient="landscape"/></w:sectPr></w:pPr>'
    )
    tgt = _p(_r("横向节末段"), tgt_ppr)
    _make_docx(d, _doc(src + tgt))
    rc = dp.main(["fix-ppr", str(d), "--para", "1", "--clone-from", "0", "--no-gate", "--no-backup"])
    assert rc == 0
    _, paras = _reopen_paras(d)
    ppr = paras[1].find(W + "pPr")
    assert ppr is not None
    sect = ppr.find(W + "sectPr")
    assert sect is not None, "目标段自身 sectPr 必须保留(节边界不丢)"
    assert sect.find(W + "pgSz").get(W + "orient") == "landscape", "sectPr 内容(横向)完整保留"
    # 脏 ind(firstLine=0)被清(源无 ind)
    assert ppr.find(W + "ind") is None


def test_inspect_hint_expect_is_real_substring(tmp_path, capsys):
    """inspect 末行 HINT 的 --expect 值必须是段文本真子串(禁 … 省略号, 否则照抄必被拦)。"""
    import re as _re

    d = tmp_path / "x.docx"
    long_text = "这是一段需要超过十二个字符的正文内容用于测试"
    _make_docx(
        d,
        _doc(
            _p(_r("邻段甲"), _ppr_style("Body"))
            + _p(_r(long_text), _ppr_style("Body", '<w:ind w:firstLine="0"/>'))
            + _p(_r("邻段乙"), _ppr_style("Body"))
        ),
    )
    dp.main(["inspect", str(d), "--para", "1"])
    out = capsys.readouterr().out
    m = _re.search(r'--expect "([^"]*)"', out)
    assert m, out
    expect = m.group(1)
    assert "…" not in expect
    _, paras = _reopen_paras(d)
    assert expect in ds.para_text(paras[1]), f"--expect {expect!r} 必须是段文本真子串"


def test_prune_cache_no_prefix_crossmatch(tmp_path, monkeypatch):
    """prune 严格 <stem>-<12hex> 匹配, 不得误删 stem 为前缀的别文档缓存。"""
    import os

    cache_root = tmp_path / "cache"
    cache_root.mkdir()
    monkeypatch.setattr(dp, "CACHE_ROOT", cache_root)
    victim = cache_root / "报告-终稿-bbbbbbbbbbbb"  # 别文档, 不该被 prune "报告" 波及
    victim.mkdir()
    os.utime(victim, (1, 1))  # 设最旧 → 若跨匹配必首先被删
    for k, h in enumerate(["a" * 12, "c" * 12, "d" * 12, "e" * 12, "f" * 12]):
        p = cache_root / f"报告-{h}"
        p.mkdir()
        os.utime(p, (100 + k, 100 + k))
    dp._prune_cache(tmp_path / "报告.docx", keep=3)
    assert victim.exists(), "报告-终稿 缓存不得被 prune 报告 误删"
    remaining = sorted(x.name for x in cache_root.iterdir() if x.name.startswith("报告-") and "终稿" not in x.name)
    assert len(remaining) == 3, f"报告-<hex> 应保留 keep=3, 实剩 {remaining}"


def _seed_cache(docx: Path, pages_lines: list[str]) -> None:
    import json as _json

    cd = dp._cache_dir(docx)
    cd.mkdir(parents=True, exist_ok=True)
    (cd / "full.pdf").write_bytes(b"%PDF-1.4\n")
    (cd / "pages.txt").write_text("\n".join(pages_lines), encoding="utf-8")
    (cd / "meta.json").write_text(_json.dumps({"docx": str(docx), "pages": len(pages_lines)}), encoding="utf-8")


def test_render_no_para_no_page_clean_exit(tmp_path, monkeypatch):
    """render 无 --para/--page/--warm(缓存已就绪)→ exit 2 干净报错, 不 TypeError 崩。"""
    monkeypatch.setattr(dp, "CACHE_ROOT", tmp_path / "cache")
    d = tmp_path / "x.docx"
    _make_docx(d, _doc(_p(_r("正文"), _ppr_style("Body"))))
    _seed_cache(d, ["正文各页不同内容一", "正文各页不同内容二"])
    rc = dp.main(["render", str(d)])  # 无 --para/--page
    assert rc == 2


def test_render_header_footer_noise_not_misleading(tmp_path, monkeypatch):
    """短段文本与每页页眉/页脚重复 → 命中大半页 → exit 3 明确报不可定位, 不渲误导首页。"""
    monkeypatch.setattr(dp, "CACHE_ROOT", tmp_path / "cache")
    d = tmp_path / "x.docx"
    _make_docx(d, _doc(_p(_r("页眉文字"), _ppr_style("Body"))))
    _seed_cache(d, ["页眉文字第%d页正文" % k for k in range(10)])  # 每页都含"页眉文字"
    rc = dp.main(["render", str(d), "--para", "0"])
    assert rc == 3
