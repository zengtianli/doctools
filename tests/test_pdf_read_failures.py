"""PDF read部分失败不能以其余页/metadata成功报绿。"""
import importlib.util
import subprocess
import sys
from pathlib import Path
from types import SimpleNamespace

import pytest

PATH = Path(__file__).resolve().parents[1] / "scripts/document/pdf_cli.py"
spec = importlib.util.spec_from_file_location("pdf_read_failure_tests", PATH)
pdf_cli = importlib.util.module_from_spec(spec)
sys.modules[spec.name] = pdf_cli
spec.loader.exec_module(pdf_cli)


@pytest.fixture
def args(tmp_path):
    file = tmp_path / "fixture.pdf"
    file.write_bytes(b"%PDF-unit-test-input")
    return SimpleNamespace(pdf=str(file), list=False, page=None, pages=None)


def test_one_failed_page_retains_other_text_and_returns_failure(args, monkeypatch, capsys):
    class Page:
        def __init__(self, fail=False):
            self.fail = fail

        def extract_text(self):
            if self.fail:
                raise ValueError("bad page")
            return "valid second page"

    monkeypatch.setattr(pdf_cli, "pdfplumber", SimpleNamespace(open=lambda _: SimpleNamespace(
        pages=[Page(True), Page()], close=lambda: None)))
    assert pdf_cli._cmd_read(args) == 1
    output = capsys.readouterr()
    assert "valid second page" in output.out
    assert "ERROR page 1" in output.err
    assert Path(args.pdf).read_bytes() == b"%PDF-unit-test-input"


def test_read_success_is_still_zero(args, monkeypatch):
    monkeypatch.setattr(pdf_cli, "pdfplumber", SimpleNamespace(open=lambda _: SimpleNamespace(
        pages=[SimpleNamespace(extract_text=lambda: "complete page")], close=lambda: None)))
    assert pdf_cli._cmd_read(args) == 0


def test_outline_dependency_is_required_even_when_pdfinfo_succeeds(args, monkeypatch):
    args.list = True
    monkeypatch.setattr(pdf_cli.subprocess, "run", lambda *a, **kw: subprocess.CompletedProcess(a, 0, "Pages: 1", ""))
    monkeypatch.setattr(pdf_cli, "PdfReader", None)
    assert pdf_cli._cmd_read(args) == 2


def test_outline_parse_failure_is_not_metadata_success(args, monkeypatch):
    args.list = True
    monkeypatch.setattr(pdf_cli.subprocess, "run", lambda *a, **kw: subprocess.CompletedProcess(a, 0, "Pages: 1", ""))

    def broken(_):
        raise ValueError("bad outline")

    monkeypatch.setattr(pdf_cli, "PdfReader", broken)
    assert pdf_cli._cmd_read(args) == 1


def test_failed_metadata_is_not_outline_success(args, monkeypatch):
    args.list = True
    monkeypatch.setattr(pdf_cli.subprocess, "run", lambda *a, **kw: subprocess.CompletedProcess(a, 1, "", "bad metadata"))
    monkeypatch.setattr(pdf_cli, "PdfReader", lambda _: SimpleNamespace(outline=[SimpleNamespace(title="fixture")],
        get_destination_page_number=lambda _: 0))
    assert pdf_cli._cmd_read(args) == 1
