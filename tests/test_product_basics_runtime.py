"""Product acceptance permits real external hot reads without weakening provenance."""
import hashlib
import importlib.util
import json
from pathlib import Path
import plistlib
import shutil
import subprocess
import sys
import tempfile
import unittest
from unittest.mock import patch

SCRIPT = Path(__file__).resolve().parents[1] / "mac/scripts/accept/product_basics.py"
spec = importlib.util.spec_from_file_location("dockit_product_basics", SCRIPT)
basics = None
MAC_SIGNING = sys.platform == "darwin" and (Path.home() / "Apps/chapter/engine/app_sop.py").is_file()
if MAC_SIGNING:
    basics = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(basics)


@unittest.skipUnless(MAC_SIGNING, "requires macOS signing and the local Chapter engine")
class RuntimeProvenanceTests(unittest.TestCase):
    def setUp(self):
        self.folder = tempfile.TemporaryDirectory(prefix="dockit-provenance-")
        self.addCleanup(self.folder.cleanup)
        root = Path(self.folder.name).resolve()
        self.repo = root / "repo"
        (self.repo / "Sources").mkdir(parents=True)
        self.compiled = self.repo / "Sources/main.swift"
        self.compiled.write_text("compiled")
        self.runtime, self.other = root / "backend.py", root / "other.py"
        self.runtime.write_text("original backend")
        self.other.write_text("other build input")
        self.app = root / "Fixture.app"
        binary = self.app / "Contents/MacOS/Fixture"
        binary.parent.mkdir(parents=True)
        shutil.copyfile("/usr/bin/true", binary)
        binary.chmod(0o755)
        self.resource = self.app / "Contents/Resources/data.txt"
        self.resource.parent.mkdir()
        self.resource.write_text("signed package data")
        (self.app / "Contents/Info.plist").write_bytes(plistlib.dumps({
            "CFBundleIdentifier": "test.tianli.dockit.provenance", "CFBundleExecutable": "Fixture",
            "CFBundlePackageType": "APPL", "CFBundleShortVersionString": "1", "CFBundleVersion": "1"}))
        subprocess.run(["/usr/bin/codesign", "--force", "--sign", "-", "--timestamp=none", str(self.app)],
                       check=True, capture_output=True)
        self.component = {"id": "doc-tools", "repo": self.repo, "sop": {
            "source": ["Sources/**"], "source_extra_files": [str(self.runtime), str(self.other)],
            "source_runtime_files": [str(self.runtime)]}}
        self.receipt = self.repo / "perf/build-receipt.json"
        self.receipt.parent.mkdir()
        self.original = basics.app_sop.app_source_snapshot(self.component, ["Sources/**"], True)
        self.receipt.write_text(json.dumps({"schema_version": 1, "built_at": "2026-10-16T00:00:00Z",
            "build_command_sha256": "a" * 64, "artifact": basics.app_sop.artifact_snapshot(self.app),
            "source": self.original}))
        self.receipt_bytes = self.receipt.read_bytes()
        self.runtime.write_text("changed live backend")
        self.root_patch = patch.object(basics, "ROOT", self.repo)
        self.load_patch = patch.object(basics.app_sop, "load_apps", return_value=[self.component])
        self.root_patch.start()
        self.load_patch.start()
        self.addCleanup(self.root_patch.stop)
        self.addCleanup(self.load_patch.stop)

    def verify(self):
        return basics.verify_original(self.app, self.receipt)

    def test_live_backend_passes_and_current_map_differs_from_immutable_build_receipt(self):
        _, receipt, current = self.verify()
        self.assertNotEqual(current["sha256"], receipt["source"]["sha256"])
        self.assertEqual(current["files"]["@external:" + str(self.runtime)],
                         hashlib.sha256(self.runtime.read_bytes()).hexdigest())
        self.assertEqual(self.receipt.read_bytes(), self.receipt_bytes)

    def test_compiled_undeclared_and_signed_package_changes_are_rejected(self):
        for path in (self.compiled, self.other, self.resource):
            with self.subTest(path=path):
                before = path.read_bytes()
                path.write_bytes(before + b" changed")
                with self.assertRaises(AssertionError):
                    self.verify()
                path.write_bytes(before)
        self.component["sop"].pop("source_runtime_files")
        with self.assertRaises(AssertionError):
            self.verify()

    def test_corrupt_build_map_is_rejected(self):
        receipt = json.loads(self.receipt_bytes)
        receipt["source"]["file_count"] = 1
        self.receipt.write_text(json.dumps(receipt))
        with self.assertRaisesRegex(AssertionError, "Incomplete build input map"):
            self.verify()


if __name__ == "__main__":
    unittest.main()
