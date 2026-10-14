"""One-run adapter around the private tap's existing signing/byte-comparison functions; no publishing/notary calls."""
import hashlib, importlib.util, json, plistlib, sys
from pathlib import Path

sys.dont_write_bytecode = True
root, expected_zip, expected_engine = Path(sys.argv[1]).resolve(), sys.argv[2], sys.argv[3]
engine = Path('/opt/homebrew/Library/Taps/zengtianli/homebrew-private/scripts/publish.py')
assert hashlib.sha256(engine.read_bytes()).hexdigest() == expected_engine, 'Original signing engine differs'
archive = root/'unsigned.zip'
assert hashlib.sha256(archive.read_bytes()).hexdigest() == expected_zip, 'Unsigned input changed in transport'
spec = importlib.util.spec_from_file_location('private_tap_signing',engine)
pub = importlib.util.module_from_spec(spec);spec.loader.exec_module(pub)
out = root/'signed';out.mkdir(exist_ok=True)
assert not (out/'DocKit.app').exists(), 'Do not replay signing over an unknown result'
pub.run('/usr/bin/ditto','-x','-k',archive,out)
app = out/'DocKit.app'
info = plistlib.loads((app/'Contents/Info.plist').read_bytes())
assert info['CFBundleIdentifier'] == 'cyou.tianli.DocTools'
before = pub.snapshot(app)
identity = pub.identity()
assert identity.endswith('(B9LJH93LA4)'), 'Preserve the installed Developer ID team'
exe = app/'Contents/MacOS'/info['CFBundleExecutable']
code_before = pub.signed_code_digest(exe,out)
identifier = pub.code_identifier(app)
pub.run('/usr/bin/codesign','--force','--sign',identity,'--options','runtime','--timestamp',
        '--preserve-metadata=identifier,entitlements',exe)
pub.run('/usr/bin/codesign','--force','--sign',identity,'--options','runtime','--timestamp',
        '--preserve-metadata=identifier,entitlements',app)
pub.run('/usr/bin/codesign','--verify','--deep','--strict',app)
assert pub.code_identifier(app) == identifier == info['CFBundleIdentifier']
assert pub.signed_code_digest(exe,out) == code_before, 'Signing changed executable content'
after = pub.snapshot(app)
allowed = {'Contents/MacOS/'+info['CFBundleExecutable'],'Contents/CodeResources','Contents/_CodeSignature/CodeResources'}
assert {p:h for p,h in before.items() if p not in allowed} == {p:h for p,h in after.items() if p not in allowed}
grade = json.loads(pub.run(sys.executable,Path.home()/'Dev/tools/dev/lib/tools/macapp/distribution_grade.py',
                          'check',app,'--stage','signed','--json'))
value = {'schema_version':1,'source_zip_sha256':expected_zip,'signing_engine':str(engine),'signing_engine_sha256':expected_engine,
         'bundle_id':info['CFBundleIdentifier'],'version':info['CFBundleShortVersionString'],'build':info['CFBundleVersion'],
         'identity':identity,'code_content_sha256':code_before,'files':after,'signature_grade':grade,
         'notarized':False,'installed':False,'published':False}
(root/'signing-receipt.json').write_text(json.dumps(value,ensure_ascii=False,indent=2)+'\n')
pub.run('/usr/bin/ditto','-c','-k','--keepParent',app,root/'signed.zip')
print(json.dumps({'signed_zip_sha256':pub.sha(root/'signed.zip'),'bundle_id':info['CFBundleIdentifier'],'build':info['CFBundleVersion']}))
