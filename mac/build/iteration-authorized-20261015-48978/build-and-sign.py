"""One authorized build/sign run; reuse the original private-tap signer, no notary/publish/install."""
import hashlib
import importlib.util
import json
from pathlib import Path
import plistlib
import subprocess
import sys
import time

sys.dont_write_bytecode = True
ROOT = Path('/Users/tianli/Dev/tools/doctools/mac')
TASK = Path(sys.argv[1]).resolve()
PYTHON = '/Users/tianli/Dev/.venv/bin/python'
SESSION = 'chapter-task:393cf037ba19560db001948ba5282bb0'
LABEL = 'cyou.tianli.DocKit.sign-once.48978.20261015'
SIGNER = ROOT / 'build/iteration-install-20261015/sign-candidate.py'
ENGINE = Path('/opt/homebrew/Library/Taps/zengtianli/homebrew-private/scripts/publish.py')


def run(*args):
    return subprocess.run(list(map(str, args)), check=True, capture_output=True, text=True, timeout=60)


def sha(path):
    return hashlib.sha256(path.read_bytes()).hexdigest()


run(PYTHON, '/Users/tianli/Dev/tools/cc-home/tools/harness/claims.py', '--pid', '48978', '--session', SESSION, 'check', TASK)
module = '/Users/tianli/Dev/tools/cc-home/tools/harness/claims.py'
spec = importlib.util.spec_from_file_location('claims', module)
claims = importlib.util.module_from_spec(spec); spec.loader.exec_module(claims)
rows, errors = claims.read_claims(); assert not errors
record, = [r for r in rows if r.get('pid') == 48978 and r.get('session') == SESSION and r.get('paired', {}).get('token')]
context = claims.paired_record_context(record, 48978, SESSION)
assert claims.paired_rpc(context, 'check', token=record['paired']['token'])['ok']
TASK.mkdir(parents=True)
with (TASK / 'build.log').open('w') as output:
    subprocess.run(['./build.sh', '--build'], cwd=ROOT, stdout=output, stderr=subprocess.STDOUT, check=True)
print('Original build and necessary gates passed.', flush=True)
unsigned = TASK / 'unsigned-r2/DocKit.app'
unsigned.parent.mkdir()
run('/usr/bin/ditto', ROOT / 'build/DerivedData/Build/Products/Release/DocTools.app', unsigned)
run('/usr/bin/ditto', '-c', '-k', '--keepParent', unsigned, TASK / 'unsigned-r2.zip')
uid = run('ssh', '-o', 'BatchMode=yes', 'macbook', 'id -u').stdout.strip()
assert uid == '502'
assert run('ssh', '-o', 'BatchMode=yes', 'macbook', 'shasum -a 256 ' + str(SIGNER)).stdout.split()[0] == sha(SIGNER)
assert run('ssh', '-o', 'BatchMode=yes', 'macbook', 'shasum -a 256 ' + str(ENGINE)).stdout.split()[0] == sha(ENGINE)
remote = TASK / 'signing-r2'
run('ssh', '-o', 'BatchMode=yes', 'macbook', f'test ! -e {remote} && mkdir -p {remote}')
run('rsync', '-a', TASK / 'unsigned-r2.zip', 'macbook:' + str(remote / 'unsigned.zip'))
job = {'Label': LABEL, 'ProgramArguments': [PYTHON, str(SIGNER), str(remote), sha(TASK / 'unsigned-r2.zip'), sha(ENGINE)],
       'RunAtLoad': True, 'ProcessType': 'Background', 'StandardOutPath': str(remote / 'signer.out'),
       'StandardErrorPath': str(remote / 'signer.err')}
job_path = TASK / 'signing-once.plist'
job_path.write_bytes(plistlib.dumps(job))
run('rsync', '-a', job_path, 'macbook:' + str(TASK) + '/')
loaded = False
try:
    run('ssh', '-o', 'BatchMode=yes', 'macbook', f'/bin/launchctl bootstrap gui/{uid} {job_path}')
    loaded = True
    deadline = time.monotonic() + 150
    while time.monotonic() < deadline:
        done = subprocess.run(['ssh', '-o', 'BatchMode=yes', 'macbook', f'test -f {remote}/signed.zip'], capture_output=True, timeout=20)
        if done.returncode == 0:
            break
        status = run('ssh', '-o', 'BatchMode=yes', 'macbook', f'/bin/launchctl print gui/{uid}/{LABEL}').stdout
        if 'state = not running' in status and 'last exit code = 0' not in status:
            raise RuntimeError('One-shot signer stopped without a package: ' + status[-1000:])
        time.sleep(2)
    else:
        raise RuntimeError('Signing deadline exceeded; no installation attempted')
    run('rsync', '-a', 'macbook:' + str(remote / 'signed.zip'), 'macbook:' + str(remote / 'signing-receipt.json'),
        'macbook:' + str(remote / 'signer.out'), 'macbook:' + str(remote / 'signer.err'), str(TASK) + '/')
finally:
    if loaded:
        run('ssh', '-o', 'BatchMode=yes', 'macbook', f'/bin/launchctl bootout gui/{uid}/{LABEL}')
result = json.loads((TASK / 'signer.out').read_text())
assert sha(TASK / 'signed.zip') == result['signed_zip_sha256']
run('/usr/bin/ditto', '-x', '-k', TASK / 'signed.zip', TASK / 'signed')
receipt = json.loads((TASK / 'signing-receipt.json').read_text())
for name, digest in receipt['files'].items():
    assert sha(TASK / 'signed/DocKit.app' / name) == digest
run('/usr/bin/codesign', '--verify', '--deep', '--strict', TASK / 'signed/DocKit.app')
print('Stable Developer ID signature and all transported member hashes verified; one-shot job removed.', flush=True)
