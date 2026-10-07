# DocKit

[中文](README.md)

The app menu adds configuration import/export, optional iCloud sync (off by default), and update checks. It remembers the last operation and target formats. Rule toggles, external-service consent, template paths, documents and task history remain local to each task. This edition reads updates from your private iCloud release folder.

SwiftUI client inside [doctools](../README.md), maintained in the parent repository with its existing document backend. Operations and options come from `gui-ops`; the local `~/Dev/.venv` environment is required.

Run `./build.sh --check` for the real backend decoding gate, or `./build.sh` for a signed Release build without installation. Only `./build.sh --install` replaces `/Applications/DocKit.app`.

The catalog owns the display name; bundle ID remains `cyou.tianli.DocTools`. Builds use the shared Xcode selector.

Fixed acceptance scripts live in `scripts/accept/`. Chapter runs them and records the evidence:

```sh
~/Dev/.venv/bin/python ~/Apps/chapter/engine/app_sop.py accept --app doc-tools-doctools --check functionality --check recovery --check privacy --check native_ui --json
```

Functionality and recovery checks use temporary documents. Native UI acceptance builds current source and calls the app's `--ui-self-test` to render real views offscreen without installation, focus changes or clipboard access. This verifies the local build; the installed icon still requires confirmation in Chapter.

Normalization and quotation operations use local document engines. Sensitive-word scanning sends the names and content of `.md` files in the chosen folder (DOCX is not read) to the external Claude service and requires explicit consent whenever that operation is selected. Acceptance checks exercise refusal and local processing without sending documents; they do not verify external service availability.

## Command line (for agents)

The GUI is for people; `dockit` is for agents. Both call `gui_run` in the same backend, `doc_gui_backend.py`, so the operation catalog, option validation, cloud-consent gate, overwrite gate and output attribution have a single implementation. `dockit` is a POSIX script inside the app bundle (`DocKit.app/Contents/Resources/bin/dockit`) that runs the working-tree backend with the shared `~/Dev/.venv`, so backend changes take effect without rebuilding the app. `./build.sh --install` links it to `~/.local/bin/dockit`; if the venv or backend is missing it exits 69.

```sh
dockit ops                                          # 14 operations, flagged destructive / cloud / diagnostic / overwrite needs --yes
dockit ops clean                                    # input formats, target, options, defaults and an example
dockit ops --json                                   # same as the GUI's gui-ops: {"ok":true,"ops":[...]}
dockit run clean --opt scope.table=0 --json /abs/report.docx
dockit run convert --to md /abs/a.docx /abs/b.pdf
dockit run convert --to word --dry-run /abs/b.md     # check whether an existing b.docx would be overwritten
dockit run formatclone --opt ref=/abs/reference.docx /abs/draft.md
dockit run renum --to tabfig /abs/report.docx
dockit run bidfinal --to pei --json /abs/bid.docx    # results[].verdict: pass / red / error
dockit run view /abs/notes.md                        # reports the HTML path; --open opens a browser
dockit run lowercase --dry-run /abs/data.xlsx        # preview destructive operations, then add --yes
dockit run typeset --background --json /abs/report.md   # long task: returns a job id at once
dockit status <job> --json                          # running / done / interrupted / lost, full result when finished
dockit status --json                                # read back current state: installed version and build, remembered settings, recent background jobs
dockit settings --json                              # the last operation and per-operation target format the app remembers
dockit settings set target_formats.convert md       # change one of them; the other key is last_operation <op>
dockit doctor                                       # venv, uv, soffice, markitdown, pdftotext, claude and more
dockit config status --json                         # the "remember settings in iCloud" switch, portable settings, whether the app is running
dockit config export -o /abs/dockit-config.json     # export settings; config import <file> --yes imports, config sync on|off --yes flips the switch
dockit update check --json                          # current version, latest on this channel, whether there is an update, how to upgrade
```

Exit codes: 0 success; 1 business failure (an input failed or was skipped, a bid gate is red, or doctor found a missing dependency); 2 usage or validation error (unknown operation, option or target, missing required option, missing cloud consent, destructive operation without `--yes`, an existing file would be overwritten without `--yes`, or no valid files); 75 another DocKit operation is running in the same folder tree (the folder, a parent or a subfolder), so retry later; 128+N interrupted by signal N (an agent timeout's SIGTERM gives 143), and the running engine process group is stopped too; 69 the venv or backend is missing.

`--json` prints the same result envelope as the GUI: `ok`, `all_ok`, `results[].outputs` (absolute paths), `results[].modified_in_place`, `succeeded`/`total`, `skipped_missing` and `log`, plus `error` and `error_code` on failure. `ok` is request-level: it is true once the request was processed, even if some inputs failed. To know whether everything succeeded, read `all_ok`, each `results[].ok`, or the exit code. Usage errors are also JSON when `--json` is given.

Long tasks: `run` is synchronous by default, and template typesetting or a cold soffice start can take minutes, so allow a 10-minute timeout. `--background` runs all validation in the foreground first (validation errors still exit 2), then runs the operation in the background and returns `job.id` at once. `dockit status <id>` reports progress and, when the job has finished, the full result envelope with the same exit code a synchronous run would have. Job records live in `~/Library/Caches/cyou.tianli.DocTools/jobs` for 7 days.

Cloud, overwrite and destructive operations:

- `scan` sends the names and content of `.md` files in the folder to Claude through `claude -p`. Without `--opt privacy.cloud_consent=1` it exits 2 before listing the folder. Add the option only after the owner explicitly agrees; it starts `claude -p`, so do not run it from inside a Claude Code session (llm_client nesting limit). Findings are returned in `results[0].findings`. A folder with no `.md` files starts no engine and sends nothing; `findings` is `[]` and `md_files` is 0.
- Overwrite gate: some outputs of `convert`, `typeset`, `merge` and `split` can share a name with an existing file (`b.md` converted to `b.docx`, `merged.md`, `merged.csv`, per-sheet names such as `b_Sheet1.xlsx`), and the engines overwrite it without a backup. When such a target already exists these operations refuse by default (`error_code: would_overwrite`, with the paths in `overwrites`). Add `--yes` to confirm, or tick "允许覆盖已有同名文件" (allow overwriting existing files) in the GUI. `--dry-run` reports the same list. Outputs with a DocKit suffix (`_fixed`, `_styled`, `_序号修正`, `_lower`, `_split` folders) replace the previous run's output when you rerun; `formatclone`'s `_成品` output never overwrites.
- `fontunify`, and `clean` on a pptx, rewrite the source file in place, and the result marks it in `modified_in_place`. The first backup is `<name>.pptx.backup`; if a `.backup` already exists, a new `<name>.bak-N-<date>.pptx` is written instead, so a rerun never replaces the original with an already-modified copy. `fontunify`, `lowercase` and `stripchrome` require `--yes`; `--dry-run` only lists the files that would be processed and writes nothing.
- All other outputs are written next to the source and sources stay unchanged. Only one DocKit write operation may run in a folder tree (the folder, its parents and subfolders) at a time, so outputs are never attributed to another DocKit run. Files that other programs write into the folder at the same time can still be counted as outputs.

Coverage:

| GUI capability | Command line |
|---|---|
| Sidebar operation list, descriptions, supported formats, destructive flags | `dockit ops [--json]` |
| Target picker, option groups, defaults, required options, "allow overwriting existing files" | `dockit ops <op>`; `run --to T`, `--opt K=V`, `--yes` |
| Run button and results (per-file status, outputs, succeeded N/M, skipped files, log) | `dockit run … [--json] [-v]` |
| Waiting while an operation runs | `run --background` and `dockit status <id>` |
| "Ready / backend unreachable" | `dockit doctor [--json]` |
| Cloud consent for sensitive-word scanning | `--opt privacy.cloud_consent=1` (the same gate) |
| Remembered last operation and per-operation target format | `dockit settings [--json]`; `dockit settings set <key> <value>` |
| Version and build shown in the "Settings and Updates" window | `app` in `dockit status [--json]` |
| "Settings and Updates" window: remember settings in iCloud, export settings, import settings | `dockit config sync on\|off --yes [--dry-run]`, `config export -o <file> [--force]`, `config import <file> --yes`; read back with `config status` |
| "Settings and Updates" window: check for updates | `dockit update check [--json]` |

The read commands are `ops`, `status`, `doctor`, `settings`, `run --dry-run`, `config status` and `update check`; they write no files, preferences or job records. The write commands are `run`, `settings set`, `config export` (changes no setting, writes only the file you name), `config import` and `config sync`. `dockit --help` lists both groups, the `--json` shape of every command, the exit code table and the window-only actions.

`settings` reads and writes the app's own two preference keys (`dockit.lastOperation`, `dockit.targetFormats`) rather than keeping a second copy. Values are validated against the operation catalog first; an unknown operation or target exits 2 and writes nothing. DocKit selects the new values the next time it starts. Other content in the preference domain, such as window positions and recent folders, is neither printed by the read command nor changed by the write command.

Window only: dropping files or folders, clearing the pending files, removing a single pending file, restoring default options, clearing a chosen reference file, revealing outputs in Finder, dismissing the banner, the ⌘K search palette, opening the "Settings and Updates" window, and the `--ui-self-test` offscreen check. The command line uses absolute path arguments, the defaults listed by `ops`, and the output paths in its JSON instead.

`config` and `update` are the items of the "Settings and Updates" window. They belong to the app itself (its preference domain, version and release channel), so the `DocKit.app` executable runs them through the command layer shared across products, without creating a window or a Dock icon; `dockit` forwards the words unchanged and returns the output and exit code unchanged, and a running DocKit follows the change on its own. Their `--json` output has the shared layer's shape, with failures as `{"ok":false,"command":…,"error":{"code","message"}}` rather than the flat `error` / `error_code` of the other commands; exit codes are 0, 1 and 2 only, and `dockit config --help` lists every form and `error.code`. When the words cannot be handed to the app executable the exit code is 1: `app_missing` (no `DocKit.app`), `app_outdated` (the installed app predates these commands and is not started), `app_failed`, `timeout`. `config import` and `config sync` need `--yes`; `update check` reads the release record in your own iCloud Drive.

No command yet: upgrading (no silent install: `update check` reports the new version, the button name, the package location and the steps, and replacing and restarting the app is still confirmed in the window); the live iCloud sync status sentence in the window (held by the running app; a command reports only the sync pass it ran itself). The item-by-item mapping is registered under `sop.agent_cli` in `project.yaml`.

<!-- lightweight:start -->
## Resource use

| Installed | Idle memory | Idle CPU | Read operation list (production backend, including Python process start) |
|---|---|---|---|
| **2.3 MB** | **46 MB** | **0.05%** | **29 ms** |

Native SwiftUI client reusing the local shared Python document engine; backend processes exit when idle.

<sub>v1.0 (298) · Mac16,12 / Apple M4 / macOS 27.2 · Installed personal app with 14 real operations and dynamic options loaded; no user document opened or modified. · measured 2026-09-26. Measured on the listed device; re-measured for each version. Memory uses phys_footprint; CPU is CPU time ÷ wall time over a 60-second sampling window; sizes in decimal MB. Raw data: [perf/lightweight.json](perf/lightweight.json).</sub>
<!-- lightweight:end -->
