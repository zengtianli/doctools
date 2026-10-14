#!/usr/bin/env python3
"""Install one source-bound Developer ID original through build.sh; never launch it."""
import argparse
import hashlib
import json
import os
from pathlib import Path
import plistlib
import re
import shutil
import subprocess
import sys
import tempfile
import time

sys.dont_write_bytecode = True
sys.path.insert(0, str(Path.home() / 'Apps/chapter/engine'))
import app_sop


def run(*args):
    return subprocess.run(list(map(str, args)), capture_output=True, text=True, check=True, timeout=60)


def signature(app):
    run('/usr/bin/codesign', '--verify', '--deep', '--strict', app)
    details = run('/usr/bin/codesign', '-d', '--verbose=4', app).stderr
    team = re.search(r'^TeamIdentifier=(.+)$', details, re.M)
    if not team or 'Authority=Developer ID Application:' not in details or '(runtime)' not in details or 'Timestamp=' not in details:
        raise RuntimeError('Developer ID, hardened runtime and timestamp required; original installation preserved')
    return team[1]


def members(app):
    return {str(p.relative_to(app)): ('link:' + os.readlink(p) if p.is_symlink() else hashlib.sha256(p.read_bytes()).hexdigest())
            for p in app.rglob('*') if p.is_file() or p.is_symlink()}


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('app', type=Path)
    parser.add_argument('receipt', type=Path)
    parser.add_argument('--check-only', action='store_true')
    args = parser.parse_args()
    candidate, receipt = args.app.resolve(), args.receipt.resolve()
    root = Path(__file__).resolve().parents[1]
    component = app_sop.load_apps('doc-tools')[0]
    if component['repo'].resolve() != root:
        raise RuntimeError('Component source mismatch')
    valid, reason = app_sop.verify_build_receipt(component, candidate, receipt)
    if not valid:
        raise RuntimeError(reason)
    info = plistlib.loads((candidate / 'Contents/Info.plist').read_bytes())
    if info['CFBundleIdentifier'] != 'cyou.tianli.DocTools':
        raise RuntimeError('Wrong bundle identity')
    run(sys.executable, Path.home() / 'Dev/tools/dev/lib/tools/macapp/distribution_grade.py', 'check', candidate, '--stage', 'signed', '--json')
    team = signature(candidate)
    target = Path('/Applications/DocKit.app')
    if candidate == target:
        raise RuntimeError('Candidate must be separate from installed original')
    if target.exists():
        previous = plistlib.loads((target / 'Contents/Info.plist').read_bytes())
        if previous['CFBundleIdentifier'] != info['CFBundleIdentifier'] or signature(target) != team:
            raise RuntimeError('Bundle or signing team changed; original installation preserved')
        # codesign writes designated requirements to stdout on this OS; accept either stream explicitly.
        response = run('/usr/bin/codesign', '-d', '-r-', target)
        requirements = re.findall(r'^designated => (.+)$', response.stdout + '\n' + response.stderr, re.M)
        if len(requirements) != 1:
            raise RuntimeError('Cannot read one designated requirement; original installation preserved')
        run('/usr/bin/codesign', '--verify', '-R', '=' + requirements[0], candidate)
    live = subprocess.run(['/usr/bin/pgrep', '-f', r'^/Applications/DocKit[.]app/Contents/MacOS/DocTools([[:space:]]|$)'], capture_output=True)
    if live.returncode == 0:
        print('Running DocKit preserved; resume when it exits', file=sys.stderr)
        return 75
    if live.returncode != 1:
        raise RuntimeError('Cannot verify running App state')
    link = Path.home() / '.local/bin/dockit'
    expected = str(target / 'Contents/Resources/bin/dockit')
    if (link.exists() and not link.is_symlink()) or (link.is_symlink() and os.readlink(link) != expected):
        raise RuntimeError('Existing CLI entry belongs to another target; preserved')
    result = {'installed': False, 'target': str(target), 'version': info['CFBundleShortVersionString'],
              'build': info['CFBundleVersion'], 'team': team, 'candidate_receipt': str(receipt)}
    if args.check_only:
        print(json.dumps(result, ensure_ascii=False))
        return 0
    scratch = Path(tempfile.mkdtemp(prefix='.dockit-install-', dir='/Applications'))
    staged = scratch / 'DocKit.app'
    backup = None
    installed = False
    try:
        run('/usr/bin/ditto', candidate, staged)
        if members(staged) != members(candidate):
            raise RuntimeError('Staged package differs')
        signature(staged)
        if target.exists():
            retired = Path.home() / '.Trash' / ('app-rebuild-' + time.strftime('%Y%m%d-%H%M%S') + '-' + str(os.getpid()))
            retired.mkdir(parents=True)
            backup = retired / 'DocKit.app'
            target.rename(backup)
        staged.rename(target)
        installed = True
        if members(target) != members(candidate):
            raise RuntimeError('Installed package differs')
        signature(target)
        link.parent.mkdir(parents=True, exist_ok=True)
        if not link.is_symlink():
            link.symlink_to(expected)
    except BaseException:
        if installed:
            target.rename(scratch / 'failed.app')
        if backup is not None and backup.exists():
            backup.rename(target)
        raise
    finally:
        shutil.rmtree(scratch, ignore_errors=True)
    result.update(installed=True, previous_kept_at=str(backup) if backup else None)
    print(json.dumps(result, ensure_ascii=False))
    return 0


if __name__ == '__main__':
    raise SystemExit(main())
