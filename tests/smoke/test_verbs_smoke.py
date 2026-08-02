#!/usr/bin/env python3
"""test_verbs_smoke.py — 逐条真敲 docx_cli 的每一个动词。

一条动词一个用例，契约来自 `_verb_specs.py`（那是 SSOT，别在这里再抄一遍 argv）。

断言两条：

    1. rc == expect_rc
    2. mutates=False → 源 fixture 的 md5 必须没变
       mutates=True  → 源 fixture 的 md5 必须变了

第 2 条是这套 smoke 里最值钱的部分。它是双向的，两个方向抓的是两类相反的 bug：

    只读动词偷偷写盘   —— `audit` / `snapshot` / `extract` 跑完源件被改了
    写盘动词其实没写   —— 静默空跑，rc 照样 0，看输出也像干了活

跑：

    python3 -m pytest tests/smoke -q
    python3 -m pytest tests/smoke -q -k "audit"        # 只跑某一族
    python3 -m pytest tests/smoke -q -rs               # 连 skip 理由一起打出来

覆盖面由 `python3 tools/check_smoke_coverage.py` 单独把关（CLI 有表里没有、
表里有 CLI 没有，两个方向都判红）—— **本文件不负责证明"没漏"**，
因为一个只跑自己列出来的用例的测试，永远证明不了自己列全了。
"""
from __future__ import annotations

import hashlib
import subprocess
import sys
from pathlib import Path

import pytest

sys.path.insert(0, str(Path(__file__).resolve().parent))
from _verb_specs import SPECS, VERBS  # noqa: E402

REPO = Path(__file__).resolve().parents[2]
DOCX_CLI = REPO / "scripts" / "document" / "docx_cli.py"
TIMEOUT = 300


def _md5(p: Path) -> str:
    return hashlib.md5(p.read_bytes()).hexdigest()


def _resolve(argv: tuple[str, ...], subst: dict) -> list[str]:
    return [subst.get(tok, tok) for tok in argv]


def _pytest_params():
    out = []
    for verb in VERBS:
        spec = SPECS[verb]
        marks = []
        if spec["skip_reason"]:
            # 理由必须打得出来：`-rs` 会把它原样列出，不然 skip 就是个黑盒，
            # 时间一长没人知道当初为什么跳过、还该不该继续跳。
            marks.append(pytest.mark.skip(reason=f"{verb}: {spec['skip_reason']}"))
        out.append(pytest.param(verb, marks=marks, id=verb.replace(" ", "-")))
    return out


@pytest.mark.parametrize("verb", _pytest_params())
def test_verb_smoke(verb: str, workspace: dict) -> None:
    spec = SPECS[verb]
    argv = _resolve(spec["argv"], workspace["subst"])
    docx = workspace["docx"]

    before = _md5(docx)
    proc = subprocess.run(
        [sys.executable, str(DOCX_CLI), *argv],
        cwd=str(REPO), capture_output=True, text=True,
        timeout=TIMEOUT, stdin=subprocess.DEVNULL,
    )
    after = _md5(docx) if docx.exists() else "<源 docx 被删掉了>"

    cmd = "docx_cli.py " + " ".join(argv)
    detail = (
        f"\n  命令: {cmd}"
        f"\n  rc  : {proc.returncode}（预期 {spec['expect_rc']}）"
        f"\n  备注: {spec['note'] or '—'}"
        f"\n  stdout(尾): {proc.stdout.strip()[-400:] or '（空）'}"
        f"\n  stderr(尾): {proc.stderr.strip()[-400:] or '（空）'}"
    )

    assert proc.returncode == spec["expect_rc"], f"[{verb}] 退出码不符{detail}"

    if spec["mutates"]:
        assert after != before, (
            f"[{verb}] 声明 mutates=True，但源 docx 一个字节都没变 —— "
            f"这条八成是静默空跑（rc 对、活没干）{detail}"
        )
    else:
        assert after == before, (
            f"[{verb}] 声明 mutates=False，源 docx 却被改了 —— "
            f"只读动词偷偷写盘，或该动词的输出该走 -o/--report 而没走{detail}"
        )


def test_specs_table_is_wellformed() -> None:
    """表本身的自检：不许有重复动词、字段类型不许飘。

    `_verb_specs._build()` 已经拦了重复键（tuple 表不像 dict 会静默覆盖，
    但也不会自动报错，得自己拦）。这里再钉一遍字段形状，免得哪天有人
    往 argv 里塞了个 list、跑起来才发现替换不掉占位符。
    """
    assert VERBS, "动词表为空 —— 空集上的一切断言都成立，这种绿不算通过"
    for verb in VERBS:
        spec = SPECS[verb]
        assert isinstance(spec["argv"], tuple) and spec["argv"], f"{verb}: argv 必须是非空 tuple"
        assert all(isinstance(t, str) for t in spec["argv"]), f"{verb}: argv 元素必须都是 str"
        assert isinstance(spec["expect_rc"], int), f"{verb}: expect_rc 必须是 int"
        assert isinstance(spec["mutates"], bool), f"{verb}: mutates 必须是 bool"
