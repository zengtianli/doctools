"""Shared, isolated acceptance helpers; Chapter owns delivery-evidence.json."""
from __future__ import annotations

import hashlib
import json
import os
from pathlib import Path
import subprocess
import sys
import tempfile

MAC_ROOT = Path(__file__).resolve().parents[2]
REPO_ROOT = MAC_ROOT.parent
BACKEND = REPO_ROOT / "scripts/document/doc_gui_backend.py"
PYTHON = Path.home() / "Dev/.venv/bin/python"
DOCKIT = MAC_ROOT / "bin/dockit"


def workspace(name: str):
    return tempfile.TemporaryDirectory(prefix=f"dockit-accept-{name}-")


def backend(*args: str, timeout: float = 60) -> dict:
    result = subprocess.run(
        [str(PYTHON), str(BACKEND), *map(str, args)],
        cwd=REPO_ROOT, capture_output=True, text=True, timeout=timeout,
    )
    assert result.returncode == 0, f"backend exit {result.returncode}: {result.stderr}"
    value = json.loads(result.stdout)
    assert isinstance(value, dict) and isinstance(value.get("ok"), bool), value
    return value


def cli(*args: str, timeout: float = 120) -> tuple[int, dict]:
    """Run the agent CLI wrapper that build.sh copies into Contents/Resources/bin; always asks for --json."""
    env = {**os.environ, "BROWSER": "/usr/bin/true"}
    result = subprocess.run(
        [str(DOCKIT), *map(str, args), "--json"],
        cwd=REPO_ROOT, capture_output=True, text=True, timeout=timeout, env=env,
    )
    value = json.loads(result.stdout)
    assert isinstance(value, dict) and isinstance(value.get("ok"), bool), value
    return result.returncode, value


def sha256(path: Path) -> str:
    return hashlib.sha256(path.read_bytes()).hexdigest()


def finish(name: str, summary: str, **details) -> None:
    value = {"summary": summary, **details}
    out = os.environ.get("SOP_OUT_DIR")
    if out:
        target = Path(out) / f"{name}.detail.json"
        target.write_text(json.dumps(value, ensure_ascii=False, indent=2) + "\n")
    print(json.dumps(value, ensure_ascii=False))
