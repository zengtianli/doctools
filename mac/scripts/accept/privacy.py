#!/usr/bin/env python3
"""Fail-closed cloud consent and actual local processing under OS network denial."""
from __future__ import annotations

import json
from pathlib import Path
import subprocess

from _common import BACKEND, PYTHON, REPO_ROOT, backend, finish, sha256, workspace


NETWORK_PROFILE = "(version 1) (allow default) (deny network*)"
# A regression must not be able to launch Claude while testing rejected consent.
SCAN_PROFILE = NETWORK_PROFILE + ' (deny process-exec) (allow process-exec (regex ".*/python[0-9.]*$"))'


def isolated(*args: str, profile: str = NETWORK_PROFILE) -> dict:
    result = subprocess.run(
        ["/usr/bin/sandbox-exec", "-p", profile, str(PYTHON), str(BACKEND), *args],
        cwd=REPO_ROOT, capture_output=True, text=True, timeout=30,
    )
    assert result.returncode == 0, f"Sandbox/backend failed: {result.stderr}"
    payload = json.loads(result.stdout)
    assert isinstance(payload.get("ok"), bool), payload
    return payload


def main() -> None:
    assert Path("/usr/bin/sandbox-exec").exists(), "OS network sandbox unavailable; privacy acceptance cannot be claimed"
    # Prove enforcement rather than interpreting a refused connection as isolation.
    probe = subprocess.run(
        ["/usr/bin/sandbox-exec", "-p", NETWORK_PROFILE, str(PYTHON), "-c",
         "import socket\ntry:\n socket.socket().connect(('127.0.0.1',9))\n"
         "except PermissionError:\n print('NETWORK_DENIED')\nelse:\n raise SystemExit(1)"],
        capture_output=True, text=True, timeout=10,
    )
    assert probe.returncode == 0 and probe.stdout.strip() == "NETWORK_DENIED", probe.stderr
    launch_probe = subprocess.run(
        ["/usr/bin/sandbox-exec", "-p", SCAN_PROFILE, str(PYTHON), "-c",
         "import subprocess\ntry:\n subprocess.run(['/usr/bin/true'],check=True)\n"
         "except PermissionError:\n print('CHILD_DENIED')\nelse:\n raise SystemExit(1)"],
        capture_output=True, text=True, timeout=10,
    )
    assert launch_probe.returncode == 0 and launch_probe.stdout.strip() == "CHILD_DENIED", launch_probe.stderr

    catalog = backend("gui-ops")
    scan = next(op for op in catalog["ops"] if op["id"] == "scan")
    consent = next(opt for opt in scan["options"] if opt["id"] == "privacy.cloud_consent")
    assert consent["type"] == "bool" and consent["default"] is False
    disclosure = scan["subtitle"] + consent["title"]
    assert all(term in disclosure for term in ("文件名", "内容", "Claude", "发送"))

    with workspace("privacy") as temp:
        root = Path(temp)
        canary = "DOCKIT_SYNTHETIC_PRIVATE_7192c47a"
        source = root / "private-fixture.md"
        source.write_text(f'正文:"本机",面积10平方米。\n{canary}\n', encoding="utf-8")
        original = sha256(source)
        before = sorted(path.name for path in root.iterdir())
        for value in (None, "0", "false", "invalid-consent"):
            options = [] if value is None else ["--opt", f"privacy.cloud_consent={value}"]
            payload = isolated("gui-run", "--op", "scan", *options, "--files", str(root), profile=SCAN_PROFILE)
            assert payload["ok"] is False and payload["error"], payload
            expected = "不是布尔" if value == "invalid-consent" else "外发授权"
            assert expected in payload["error"], payload
            assert canary not in json.dumps(payload)
            assert sha256(source) == original and sorted(path.name for path in root.iterdir()) == before

        # Accepted syntax reaches directory validation, never an actual cloud invocation.
        payload = isolated("gui-run", "--op", "scan", "--opt", "privacy.cloud_consent=1",
                           "--files", str(root / "absent"), profile=SCAN_PROFILE)
        assert payload["ok"] is False and "不是目录" in payload["error"], payload

        for operation, expected in (("clean", "正文：“本机”，面积10m²。"),
                                    ("quotes", "正文:“本机”,面积10平方米。")):
            item = root / f"{operation}.md"
            item.write_bytes(source.read_bytes())
            payload = isolated("gui-run", "--op", operation, "--files", str(item))
            assert payload["ok"] is True and payload["succeeded"] == 1, payload
            assert canary not in json.dumps(payload), "Document text leaked into GUI status/log envelope"
            output = Path(payload["results"][0]["outputs"][0])
            assert output.resolve().is_relative_to(root.resolve())
            content = output.read_text(encoding="utf-8")
            assert expected in content and canary in content
            assert sha256(item) == original and sha256(source) == original

    finish("privacy", "云端扫描需明确同意；本机处理通过操作系统禁网验收", cloud_consent_default=False,
           rejected_consent_cases=4, network_sandbox="macOS sandbox-exec, deny network*, negative control passed",
           local_operations=["clean Markdown", "quotes Markdown"],
           canary="Synthetic document text preserved only in output, absent from GUI envelope/log",
           scope="No real cloud request or account data used; this checks the GUI scan consent boundary and two local paths, not all engines or third-party provider retention")


if __name__ == "__main__":
    main()
