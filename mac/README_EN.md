# DocKit

[中文](README.md)

SwiftUI client inside [doctools](../README.md), maintained in the parent repository with its existing document backend. Operations and options come from `gui-ops`; the local `~/Dev/.venv` environment is required.

Run `./build.sh --check` for the real backend decoding gate, or `./build.sh` for a signed Release build without installation. Only `./build.sh --install` replaces `/Applications/DocKit.app`.

The catalog owns the display name; bundle ID remains `cyou.tianli.DocTools`. Builds use the shared Xcode selector.

<!-- lightweight:start -->
## Lightweight (measured)

| Installed | Idle memory | Idle CPU | Read operation list (production backend, including Python process start) |
|---|---|---|---|
| **2.3 MB** | **46 MB** | **0.05%** | **29 ms** |

Native SwiftUI client reusing the local shared Python document engine; backend processes exit when idle.

<sub>v1.0 (298) · Mac16,12 / Apple M4 / macOS 27.2 · Installed personal app with 14 real operations and dynamic options loaded; no user document opened or modified. · measured 2026-09-26. Memory is phys_footprint (the Memory column in Activity Monitor); CPU is CPU time ÷ wall time over 60 idle seconds; sizes in decimal MB. Raw data: [perf/lightweight.json](perf/lightweight.json).</sub>
<!-- lightweight:end -->
