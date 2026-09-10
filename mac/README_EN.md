# DocKit

[中文](README.md)

SwiftUI client inside [doctools](../README.md), maintained in the parent repository with its existing document backend. Operations and options come from `gui-ops`; the local `~/Dev/.venv` environment is required.

Run `./build.sh --check` for the real backend decoding gate, or `./build.sh` for a signed Release build without installation. Only `./build.sh --install` replaces `/Applications/DocKit.app`.

The catalog owns the display name; bundle ID remains `cyou.tianli.DocTools`. Builds use the shared Xcode selector.
