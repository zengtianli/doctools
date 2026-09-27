# DocKit

[中文](README.md)

SwiftUI client inside [doctools](../README.md), maintained in the parent repository with its existing document backend. Operations and options come from `gui-ops`; the local `~/Dev/.venv` environment is required.

Run `./build.sh --check` for the real backend decoding gate, or `./build.sh` for a signed Release build without installation. Only `./build.sh --install` replaces `/Applications/DocKit.app`.

The catalog owns the display name; bundle ID remains `cyou.tianli.DocTools`. Builds use the shared Xcode selector.

Fixed acceptance scripts live in `scripts/accept/`. Chapter runs them and records the evidence:

```sh
~/Dev/.venv/bin/python ~/Apps/chapter/engine/app_sop.py accept --app doc-tools-doctools --check functionality --check recovery --check privacy --check native_ui --json
```

Functionality and recovery checks use temporary documents. Native UI acceptance builds current source and calls the app's `--ui-self-test` to render real views offscreen without installation, focus changes or clipboard access. This verifies the local build; the installed icon still requires confirmation in Chapter.

Normalization and quotation operations use local document engines. Sensitive-word scanning sends document names and content to the external Claude service and requires explicit consent whenever that operation is selected. Acceptance checks exercise refusal and local processing without sending documents; they do not verify external service availability.

<!-- lightweight:start -->
## Resource use

| Installed | Idle memory | Idle CPU | Read operation list (production backend, including Python process start) |
|---|---|---|---|
| **2.3 MB** | **46 MB** | **0.05%** | **29 ms** |

Native SwiftUI client reusing the local shared Python document engine; backend processes exit when idle.

<sub>v1.0 (298) · Mac16,12 / Apple M4 / macOS 27.2 · Installed personal app with 14 real operations and dynamic options loaded; no user document opened or modified. · measured 2026-09-26. Measured on the listed device; re-measured for each version. Memory uses phys_footprint; CPU is CPU time ÷ wall time over a 60-second sampling window; sizes in decimal MB. Raw data: [perf/lightweight.json](perf/lightweight.json).</sub>
<!-- lightweight:end -->
