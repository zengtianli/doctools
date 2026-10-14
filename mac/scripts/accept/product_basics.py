#!/usr/bin/env python3
"""Real DocKit originals, isolated preferences/cloud, hidden native controls; never install the product.

Default: installed DocKit and its existing receipt. --app/--receipt (or the corresponding
SOP_PRODUCT_BASICS_APP/RECEIPT environment) explicitly select a current-source candidate.
"""
from __future__ import annotations

import argparse
import hashlib
import json
import os
from pathlib import Path
import plistlib
import shutil
import subprocess
import sys
import time
import uuid

ROOT = Path(__file__).resolve().parents[2]
PRODUCT = "cyou.tianli.DocTools"
FEATURES = ("settings", "shortcuts", "icloud", "updates")


def sha(path):
    return hashlib.sha256(Path(path).read_bytes()).hexdigest()


def write(path, value):
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(json.dumps(value, ensure_ascii=False, indent=2) + "\n")


def defaults(domain):
    done = subprocess.run(["/usr/bin/defaults", "export", domain, "-"], capture_output=True, timeout=15)
    return plistlib.loads(done.stdout) if done.returncode == 0 else {}


def owner_snapshot():
    values = {"preferences": hashlib.sha256(plistlib.dumps(defaults(PRODUCT), sort_keys=True)).hexdigest()}
    for path in [Path("/Applications/DocKit.app/Contents/MacOS/DocTools"),
                 Path("/Applications/DocKit.app/Contents/Info.plist"),
                 Path.home() / "Library/Mobile Documents/com~apple~CloudDocs/TianliApps/Settings" / (PRODUCT + ".json")]:
        values[str(path)] = sha(path) if path.is_file() else None
    return values


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--app", type=Path, default=os.environ.get("SOP_PRODUCT_BASICS_APP", "/Applications/DocKit.app"))
    parser.add_argument("--receipt", type=Path, default=os.environ.get("SOP_PRODUCT_BASICS_RECEIPT", str(ROOT / "perf/build-receipt.json")))
    args = parser.parse_args()
    spec = json.loads(os.environ["SOP_PRODUCT_BASICS_SPEC"])
    app_id = os.environ["SOP_APP_ID"]
    assert app_id == spec["app_id"] == "doc-tools"
    out = Path(os.environ["SOP_OUT_DIR"]).resolve()
    assert out.is_relative_to(ROOT / "perf/acceptance")
    run = out / ("product-basics-" + uuid.uuid4().hex)
    run.mkdir(parents=True)
    checks = {feature: {} for feature in FEATURES}
    cases = []
    binding = {}
    error = None
    identity = "test.tianli.dockit.basics." + uuid.uuid4().hex
    suite = identity + ".preferences"
    before = owner_snapshot()
    write(run / "owner-before.json", before)

    def check(feature, name, condition):
        checks[feature][feature + "." + name] = bool(condition)

    try:
        app, receipt_path = args.app.resolve(), args.receipt.resolve()
        receipt = json.loads(receipt_path.read_text())
        artifact = receipt["artifact"]
        info = plistlib.loads((app / "Contents/Info.plist").read_bytes())
        executable = app / artifact["executable"]
        assert info["CFBundleIdentifier"] == artifact["bundle_id"] == PRODUCT
        assert info["CFBundleShortVersionString"] == artifact["version"] and info["CFBundleVersion"] == artifact["build"]
        assert sha(executable) == artifact["sha256"], "Executable differs from build receipt"
        assert receipt.get("built_at") and receipt.get("build_command_sha256")
        files = receipt["source"]["files"]
        assert hashlib.sha256(json.dumps(files, sort_keys=True, ensure_ascii=False).encode()).hexdigest() == receipt["source"]["sha256"]
        for path, digest in files.items():
            source = Path(path.removeprefix("@external:")) if path.startswith("@external:") else ROOT / path
            assert source.is_file() and sha(source) == digest, f"Build input changed: {path}"
        signature = subprocess.run(["codesign", "--verify", "--deep", "--strict", str(app)], capture_output=True, text=True, timeout=30)
        (run / "signature.log").write_text(signature.stdout + signature.stderr)
        assert signature.returncode == 0, "Invalid original signature"
        binding = {"mode": "installed" if app == Path("/Applications/DocKit.app") else "candidate",
                   "original_app": str(app), "receipt": str(receipt_path), "receipt_sha256": sha(receipt_path),
                   "artifact": artifact, "source_sha256": receipt["source"]["sha256"], "acceptor_sha256": sha(__file__)}
        clone = run / "fixture/DocKit.app"
        clone.parent.mkdir()
        shutil.copytree(app, clone, symlinks=True)
        assert sha(clone / artifact["executable"]) == artifact["sha256"]
        info.update(CFBundleIdentifier=identity, LSUIElement=True)
        (clone / "Contents/Info.plist").write_bytes(plistlib.dumps(info))
        subprocess.run(["codesign", "--force", "--deep", "--sign", "-", str(clone)], capture_output=True, check=True, timeout=30)
        env = {k: v for k, v in os.environ.items() if not k.startswith(("APP_LIFECYCLE_", "DOCKIT_", "SOP_PRODUCT_BASICS_"))}
        env.update(APP_LIFECYCLE_SUPPORT_DIR=str(run / "support"), APP_LIFECYCLE_CLOUD_DIR=str(run / "cloud"),
                   APP_LIFECYCLE_NO_RELAUNCH="1", APP_LIFECYCLE_FOLLOW_CHANNEL=identity,
                   DOCKIT_LIFECYCLE_SUITE=suite, DOCKIT_DEFAULTS_DOMAIN=suite, DOCKIT_CACHE_DIR=str(run / "cache"),
                   PYTHONDONTWRITEBYTECODE="1")
        cli = clone / "Contents/Resources/bin/dockit"
        binary = clone / artifact["executable"]
        ui_dir = run / "ui"
        ui_dir.mkdir()
        ui_done = subprocess.run([str(binary), "--product-basics-self-test"], env=dict(env, SOP_OUT_DIR=str(ui_dir)),
                                 capture_output=True, text=True, timeout=60)
        (run / "ui.log").write_text(ui_done.stdout + ui_done.stderr)
        ui = json.loads((ui_dir / "ui.json").read_text())["checks"]
        check("settings", "interaction", ui_done.returncode == 0 and all(ui.values()) and ui.get("settings.switch_action") is True)
        check("shortcuts", "action_path", ui.get("shortcuts.actual_refresh_action") is True)
        check("shortcuts", "disable_restore", ui.get("shortcuts.disabled_equivalent") is True and ui.get("shortcuts.restored_equivalent") is True)
        check("shortcuts", "conflict_or_reserved_keys", ui.get("shortcuts.reserved_refused") is True and ui.get("shortcuts.conflict_refused") is True)
        binding["ui"] = {"path": str((ui_dir / "ui.json").relative_to(ROOT)), "sha256": sha(ui_dir / "ui.json")}

        def call(label, words, predicate, code=0, override=None, direct=False):
            done = subprocess.run([str(binary if direct else cli), *words, "--json"], env=override or env,
                                  capture_output=True, text=True, timeout=40, stdin=subprocess.DEVNULL)
            body = json.loads(done.stdout)
            passed = done.returncode == code and bool(predicate(body))
            cases.append({"case": label, "argv": words, "exit": done.returncode, "passed": passed, "response": body, "stderr": done.stderr})
            assert passed, f"{label}: {done.returncode}: {body} {done.stderr}"
            return body

        state = defaults(suite)
        call("settings.fresh_process", ["settings"], lambda b: b["settings"]["last_operation"] == state["dockit.lastOperation"])
        call("shortcuts.fresh_process", ["shortcuts", "status"], lambda b: b["enabled"] is True and b["binding"]["key"] == "r", direct=True)
        check("settings", "persisted_readback", ui.get("settings.saved_operation") is True and state.get("appShortcuts.refresh.enabled") is True)
        exported = run / "export.json"
        call("settings.export", ["config", "export", "-o", str(exported)], lambda b: b["bytes"] > 0)
        envelope = json.loads(exported.read_text())
        incoming = run / "incoming.json"
        bad = dict(envelope, values={**envelope["values"], "defaults.dockit.lastOperation": ["invalid"]})
        write(incoming, bad)
        call("settings.invalid_import", ["config", "import", str(incoming), "--yes"], lambda b: b["error"]["code"] == "import_rejected", code=1)
        check("settings", "error_preserves_data", defaults(suite) == state)
        ops = call("operations", ["ops"], lambda b: len(b["ops"]) > 1)["ops"]
        operations = [row["id"] for row in ops]
        another = next(value for value in operations if value != state["dockit.lastOperation"])
        valid = dict(envelope, values={**envelope["values"], "defaults.dockit.lastOperation": another})
        write(incoming, valid)
        call("settings.valid_import", ["config", "import", str(incoming), "--yes"], lambda b: b["imported"] is True)
        call("settings.import_readback", ["settings"], lambda b: b["settings"]["last_operation"] == another)

        cloud = run / "cloud" / (PRODUCT + ".json")
        call("icloud.push", ["config", "sync", "on", "--yes"], lambda b: b["sync_enabled"] is True)
        pushed = json.loads(cloud.read_text())
        assert pushed["values"]["defaults.dockit.lastOperation"] == another
        call("icloud.off_for_remote", ["config", "sync", "off", "--yes"], lambda b: b["sync_enabled"] is False)
        remote = dict(pushed, values={**pushed["values"], "defaults.dockit.lastOperation": operations[0]})
        write(cloud, remote)
        call("icloud.pull", ["config", "sync", "on", "--yes"], lambda b: b["sync_enabled"] is True)
        call("icloud.remote_applied", ["settings"], lambda b: b["settings"]["last_operation"] == operations[0])
        check("icloud", "push_pull", True)  # all prior push/file/pull/fresh-process assertions completed
        call("icloud.off_for_conflict", ["config", "sync", "off", "--yes"], lambda b: b["sync_enabled"] is False)
        local_value = operations[-1]
        call("icloud.local_edit", ["settings", "set", "last_operation", local_value], lambda b: b["ok"] is True)
        remote = json.loads(cloud.read_text())
        remote["values"].update({"defaults.dockit.lastOperation": operations[1], "defaults.appShortcuts.refresh.enabled": False})
        write(cloud, remote)
        call("icloud.merge", ["config", "sync", "on", "--yes"], lambda b: b["sync_enabled"] is True)
        merged = defaults(suite)
        assert merged["dockit.lastOperation"] == local_value and merged["appShortcuts.refresh.enabled"] is False
        call("icloud.off_for_corrupt", ["config", "sync", "off", "--yes"], lambda b: b["sync_enabled"] is False)
        portable = {key: merged[key] for key in state if key != "appLifecycle.configuration.enabled"}
        good_cloud = cloud.read_bytes()
        cloud.write_text("{broken")
        call("icloud.bad_remote", ["config", "sync", "on", "--yes"], lambda b: b["error"]["code"] == "sync_incomplete", code=1)
        assert all(defaults(suite).get(key) == value for key, value in portable.items())
        cloud.write_bytes(good_cloud)
        call("icloud.off_for_recover", ["config", "sync", "off", "--yes"], lambda b: b["sync_enabled"] is False)
        call("icloud.recover", ["config", "sync", "on", "--yes"], lambda b: b["sync_enabled"] is True)
        check("icloud", "conflict_or_corrupt_recovery", True)
        call("icloud.off_for_unavailable", ["config", "sync", "off", "--yes"], lambda b: b["sync_enabled"] is False)
        unavailable = run / "unavailable-cloud"
        unavailable.write_text("not a directory")
        call("icloud.unavailable", ["config", "sync", "on", "--yes"], lambda b: b["ok"] is False, code=1,
             override=dict(env, APP_LIFECYCLE_CLOUD_DIR=str(unavailable)))
        call("icloud.local_fallback", ["settings"], lambda b: b["settings"]["last_operation"] == local_value)
        check("icloud", "unavailable_local_fallback", all(defaults(suite).get(key) == value for key, value in portable.items()))
        call("icloud.disable", ["config", "sync", "off", "--yes"], lambda b: b["sync_enabled"] is False)
        check("icloud", "disable_preserves_data", cloud.read_bytes() == good_cloud and all(defaults(suite).get(key) == value for key, value in portable.items()))

        call("updates.default_options", ["update", "options"], lambda b: not b["options"]["automaticCheck"])
        call("updates.save_options", ["update", "options", "--interval", "6", "--auto-download", "on", "--yes"], lambda b: b["options"]["automaticDownload"] is True)
        call("updates.fresh_options", ["update", "options"], lambda b: b["options"]["intervalHours"] == 6 and b["options"]["automaticDownload"] is True)
        options_before = defaults(identity)
        call("updates.invalid_options", ["update", "options", "--interval", "0", "--yes"], lambda b: b["error"]["code"] == "usage", code=2)
        check("updates", "options_or_store_entry", defaults(identity) == options_before)
        feed = run / "cloud/TianliApps/Updates" / identity / "private/release.json"
        call("updates.missing_source", ["update", "check"], lambda b: b["error"]["code"] == "check_incomplete", code=1)
        feed.parent.mkdir(parents=True, exist_ok=True)
        feed.write_text("{broken")
        call("updates.corrupt_source", ["update", "check"], lambda b: b["error"]["code"] == "check_incomplete", code=1)
        asset = feed.parent / "fixture.txt"
        asset.write_text("Check-only fixture; no real release is published.\n")
        release = {"version": artifact["version"], "build": str(int(artifact["build"]) + 1), "bundle_id": identity,
                   "channel": "private", "filename": asset.name, "sha256": sha(asset), "size_bytes": asset.stat().st_size}
        write(feed, release)
        call("updates.recovered_check", ["update", "check"], lambda b: b["state"] == "update_available" and b["latest"]["build"] == release["build"] and b["upgrade"]["command"] == "dockit update install --yes")
        check("updates", "check_result", all(ui.get(name + ".upgrade_actual_button_target") is True for name in ("updates-light", "updates-dark")))
        check("updates", "failure_recovery", not defaults(identity)["appUpdateState"]["last_error"])
        call("updates.preview", ["update", "install", "--dry-run"], lambda b: b["dry_run"] is True and b["installed"] is False)
        # Existing real transaction replaces only a temporary test bundle and asserts remembered settings survive.
        regression_env = dict(env, DOCKIT_NATIVE=str(binary))
        regression = subprocess.run([sys.executable, str(ROOT / "tests/test_lifecycle_cli.py"),
                                     "LifecycleCommandTests.test_update_install_replaces_a_throwaway_bundle_and_keeps_the_settings", "-v"],
                                    env=regression_env, cwd=ROOT, capture_output=True, text=True, timeout=120)
        (run / "upgrade.log").write_text(regression.stdout + regression.stderr)
        check("updates", "data_preservation", regression.returncode == 0 and "OK" in regression.stderr and "skipped" not in regression.stderr
              and all(defaults(suite).get(key) == value for key, value in portable.items()) and sha(executable) == artifact["sha256"])
        binding["upgrade_log"] = {"path": str((run / "upgrade.log").relative_to(ROOT)), "sha256": sha(run / "upgrade.log")}
    except Exception as exc:
        error = str(exc)
        for feature in FEATURES:
            check(feature, "execution_completed", False)
    finally:
        for domain in (suite, identity):
            subprocess.run(["/usr/bin/defaults", "delete", domain], capture_output=True, timeout=15)
        after = owner_snapshot()
        write(run / "owner-after.json", after)
        for feature in FEATURES:
            check(feature, "owner_data_unchanged", after == before)

    scope = ("Current-source candidate" if binding.get("mode") == "candidate" else "Installed original") + (
        "; executable SHA and build inputs verified; isolated re-identified copy, native controls offscreen, real CLI and temporary upgrade transaction. "
        "iCloud uses a private local directory transport; Apple server/cross-device delivery was not tested. No production install, release or configuration changes.")
    features = {}
    for feature in FEATURES:
        raw = {"schema_version": 1, "app_id": app_id, "spec_digest": spec["digest"], "checks": checks[feature],
               "cases": cases, "binding": binding, "scope": scope, "checked_at": time.time(), "error": error}
        path = run / (feature + ".json")
        write(path, raw)
        features[feature] = {"result": "passed" if checks[feature] and all(checks[feature].values()) else "failed",
                             "method": "ui_automation" if feature == "settings" else "real_product",
                             "evidence": str(path.relative_to(ROOT)), "evidence_sha256": sha(path), "checks": list(checks[feature])}
    passed = error is None and all(row["result"] == "passed" for row in features.values())
    detail = {"schema_version": 1, "app_id": app_id, "spec_digest": spec["digest"], "features": features, "scope": scope, "binding": binding,
              "summary": f'DocKit {binding.get("mode", "original")} 基础功能{"通过" if passed else "失败"}；14 项必测行为；iCloud 为隔离本机链路，未验苹果跨设备传输'}
    if error:
        detail["error"] = error
    write(out / "product_basics.detail.json", detail)
    print(json.dumps(detail, ensure_ascii=False))
    return 0 if passed else 1


if __name__ == "__main__":
    raise SystemExit(main())
