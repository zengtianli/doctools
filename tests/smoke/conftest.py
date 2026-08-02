#!/usr/bin/env python3
"""conftest.py — smoke fixture 造件与工作区隔离。

两层：

    _master_workspace   session 级，造一次（python-docx 存盘不便宜，93 条各造一遍
                        要多花一分多钟）
    workspace           function 级，从 master **整树拷贝**一份

拷贝而不是共用，是因为一半动词原地写盘。共用一份的话，`fix style-rebrand` 改完
样式，后面 `audit-styleset role-coverage` 读到的就不是它期望的那份文档 —— 而且
这种串味是按测试执行顺序变的，表现为「单跑绿、全跑红」，最难查的一类。

所有造件只落 `tmp_path`。仓内 `templates/template.docx` 是指向 ~/Work 真实交付件
的 symlink，smoke 一个字节都不碰它。
"""
from __future__ import annotations

import shutil
import sys
from pathlib import Path

import pytest

sys.path.insert(0, str(Path(__file__).resolve().parent))
import _fixture  # noqa: E402

REPO = Path(__file__).resolve().parents[2]
DOCX_CLI = REPO / "scripts" / "document" / "docx_cli.py"


@pytest.fixture(scope="session")
def _master_workspace(tmp_path_factory) -> Path:
    root = tmp_path_factory.mktemp("smoke-master")
    _fixture.build_docx(root / "fixture.docx")
    _fixture.build_docx_v2(root / "fixture-v2.docx")
    _fixture.build_template_docx(root / "template.docx")
    _fixture.build_side_files(root)
    return root


@pytest.fixture
def workspace(_master_workspace: Path, tmp_path: Path) -> dict:
    """每条动词一份新鲜副本 + 一张占位符替换表。"""
    work = tmp_path / "w"
    shutil.copytree(_master_workspace, work)

    docx = work / "fixture.docx"
    out_dir = work / "out"
    return {
        "root": work,
        "docx": docx,
        "subst": {
            "{DOCX}": str(docx),
            "{DOCX2}": str(work / "fixture-v2.docx"),
            "{TEMPLATE}": str(work / "template.docx"),
            "{DIR}": str(out_dir),
            "{OUT}": str(out_dir / "out.txt"),
            "{OUTDOCX}": str(out_dir / "out.docx"),
            "{OUTJSON}": str(out_dir / "out.json"),
            "{MD}": str(work / "md" / "01-第1章.md"),
            "{MD2}": str(work / "md" / "02-第3章.md"),
            "{MDDIR}": str(work / "md"),
            "{DRAFTS}": str(work / "drafts"),
            "{BIDDIR}": str(work / "bid"),
            "{DECISION}": str(work / "decision.json"),
            "{PLAN}": str(work / "relocate-plan.json"),
        },
    }
