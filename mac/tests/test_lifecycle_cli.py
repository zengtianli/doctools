"""`dockit config …` 与 `dockit update …`:整条链,真实进程,不上屏。

dockit 是 POSIX sh 薄壳加 Python 后端;「配置与更新…」窗口里的几项属于 App 包本身(偏好域、版本、发行渠道)。
所以后端把这两个动词整段交给 App 可执行文件,由它里面的共用命令层(Sources/AppLifecycleCLI.swift)在窗口自己的
配置工厂(Sources/ProductLifecycle.swift)上执行,并在创建 NSApplication 之前返回。本测试把这条链当真实进程跑:

1. bin/dockit → doc_gui_backend.py → App 程序:参数原样过去,stdout、stderr、退出码原样回来;
2. 命令读写的就是窗口那份设置(dockit settings 读写的同两个偏好键);
3. `update check` 只读隔离的发行记录;
4. 运行中的 App 跟随命令拨的开关,而且不把旧值写回去:同一个程序再起一份充当运行中的 App
   (`--lifecycle-follow-probe`:生产的 `Lifecycle.installApp`、真实主视图挂在真实 AppViewModel 上、共用窗口
   按菜单项的构造方式建出,激活策略 prohibited,都不显示),把它手里的值定时写进状态文件;
5. 同步开着、App 在运行时导入(含两次导入背靠背):从命令返回起,另起进程读到的偏好和云端那份一直是导入值。

全程隔离:程序拷进临时目录里一个换了 bundle id、带 LSUIElement 的 .app,一次性的具名偏好域,临时支持目录与
"云"目录,私有通知频道。不出窗口、不进 Dock、不联网,不向已装的 DocKit 发信号,不打开本人的设置。

    DOCKIT_NATIVE=<编好的程序> python tests/test_lifecycle_cli.py            build.sh --check 这样跑
    DOCKIT_APP=<组装好的 .app> python tests/test_lifecycle_cli.py AssembledBundleTests   build.sh 组装后这样跑(只读命令)
    DOCKIT_APP_IN_PLACE=<签过名的 .app> python tests/test_lifecycle_cli.py LifecycleCommandTests
        签名(尤其是公证版)把程序和它的 Info.plist 绑在一起,拷进别的包会被系统直接终止。这个变量让整套用例
        就地跑那个包本身:bundle id 是真的,其余隔离照旧(一次性偏好域、临时目录、私有频道),不读写本人的设置。
"""
import json
import os
from pathlib import Path
import plistlib
import shutil
import subprocess
import sys
import tempfile
import time
import unittest
import uuid

MAC = Path(__file__).resolve().parents[1]
WRAPPER = MAC / "bin" / "dockit"
NATIVE = Path(os.environ.get("DOCKIT_NATIVE") or MAC / "build" / "native" / "DocTools")
IN_PLACE = os.environ.get("DOCKIT_APP_IN_PLACE")
BUNDLE = "test.tianli.dockit.lifecycle"
PRODUCT = "cyou.tianli.DocTools"
CHANNEL = "private"
HEADQUARTERS = Path.home() / "Dev/tools/dev/lib/tools/macapp/swift-shared/AppLifecycleCLI.swift"
KEYS = ["defaults.dockit.lastOperation", "defaults.dockit.targetFormats"]
OFF = "iCloud 配置同步已关闭"


def sweep(prefix):
    """具名测试域清空后,偏好守护进程可能在进程退出后补写一个空壳文件;按本测试自己的前缀删。"""
    for _ in range(3):
        for left in (Path.home() / "Library/Preferences").glob(prefix + "*.plist"):
            left.unlink(missing_ok=True)
        time.sleep(0.3)


def bundle(root, binary, name="DocKit.app", verbs=("config", "update")):
    app = root / name
    (app / "Contents/MacOS").mkdir(parents=True)
    shutil.copy2(binary, app / "Contents/MacOS/DocTools")
    info = {"CFBundleIdentifier": BUNDLE, "CFBundleExecutable": "DocTools", "CFBundlePackageType": "APPL",
            "CFBundleShortVersionString": "1.2", "CFBundleVersion": "7", "LSUIElement": True}
    if verbs is not None:
        info["DocKitCommandVerbs"] = list(verbs)
    (app / "Contents/Info.plist").write_bytes(plistlib.dumps(info))
    return app


@unittest.skipUnless(IN_PLACE or NATIVE.is_file(), "compiled app binary not found; build.sh sets DOCKIT_NATIVE")
class LifecycleCommandTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.folder = tempfile.TemporaryDirectory(prefix="dockit-lifecycle-")
        cls.root = Path(cls.folder.name).resolve()
        cls.app = Path(IN_PLACE).resolve() if IN_PLACE else bundle(cls.root, NATIVE)
        info = plistlib.loads((cls.app / "Contents/Info.plist").read_bytes())
        cls.binary = cls.app / "Contents/MacOS" / info["CFBundleExecutable"]
        cls.bundle_id, cls.version, cls.build = info["CFBundleIdentifier"], info["CFBundleShortVersionString"], info["CFBundleVersion"]
        cls.suite = BUNDLE + "." + uuid.uuid4().hex
        cls.support, cls.cloud = cls.root / "support", cls.root / "cloud"
        cls.env = dict(os.environ, APP_LIFECYCLE_SUPPORT_DIR=str(cls.support), APP_LIFECYCLE_CLOUD_DIR=str(cls.cloud),
                       DOCKIT_LIFECYCLE_SUITE=cls.suite, DOCKIT_APP_BUNDLE=str(cls.app),
                       DOCKIT_DEFAULTS_DOMAIN=cls.suite, DOCKIT_CACHE_DIR=str(cls.root / "cache"))
        cls.env.pop("APP_LIFECYCLE_FOLLOW_CHANNEL", None)
        cls.followers = []
        # 操作与目标取自真实操作目录,不写死 id:一个带多个目标的操作,另挑两个别的操作。
        ops = json.loads(subprocess.run([str(WRAPPER), "ops", "--json"], env=cls.env, capture_output=True, text=True,
                                        stdin=subprocess.DEVNULL, timeout=120, check=True).stdout)["ops"]
        choosy = next(op for op in ops if len(op.get("targets") or []) >= 2)
        cls.op, cls.targets = choosy["id"], [target["id"] for target in choosy["targets"]]
        cls.others = [op["id"] for op in ops if op["id"] != cls.op][:3]

    @classmethod
    def tearDownClass(cls):
        for process in cls.followers:
            process.kill()
            process.wait(timeout=10)
        subprocess.run(["/usr/bin/defaults", "delete", cls.suite], capture_output=True, timeout=30)
        sweep(BUNDLE + ".")
        cls.folder.cleanup()

    def setUp(self):
        """每条用例从同一份设置起:同步关着,记住的操作与一个目标格式是种下的值。"""
        subprocess.run(["/usr/bin/defaults", "delete", self.suite], capture_output=True, timeout=30)
        self.seed(self.others[0], {self.op: self.targets[0]})
        for folder in (self.support, self.cloud):
            shutil.rmtree(folder, ignore_errors=True)

    def seed(self, operation, targets):
        run = dict(check=True, capture_output=True, timeout=30)
        subprocess.run(["/usr/bin/defaults", "write", self.suite, "dockit.lastOperation", "-string", operation], **run)
        pairs = [word for key, value in targets.items() for word in (key, "-string", value)]
        subprocess.run(["/usr/bin/defaults", "write", self.suite, "dockit.targetFormats", "-dict", *pairs], **run)

    # 命令照 agent 敲的样子:sh 薄壳 → Python 后端 → DOCKIT_APP_BUNDLE 指的那个 App 程序。
    def dockit(self, *words, env=None, cwd=None):
        return subprocess.run([str(WRAPPER), *words], env=env or self.env, cwd=cwd, capture_output=True, text=True,
                              stdin=subprocess.DEVNULL, timeout=90)

    def direct(self, *words, env=None):
        return subprocess.run([str(self.binary), *words], env=env or self.env, capture_output=True, text=True,
                              stdin=subprocess.DEVNULL, timeout=90)

    def call(self, *words, expect=0, env=None):
        done = self.dockit(*words, "--json", env=env)
        self.assertEqual(done.returncode, expect, (words, done.stdout, done.stderr))
        body = json.loads(done.stdout)
        self.assertIs(body["ok"], expect == 0, body)
        if expect:
            self.assertTrue(body["error"]["code"] and body["error"]["message"], body)
        return body

    def stored(self):
        """另起进程读存下来的值:(上次操作, 目标格式, 开关)。"""
        done = subprocess.run(["/usr/bin/defaults", "export", self.suite, "-"], capture_output=True, timeout=30)
        values = plistlib.loads(done.stdout) if done.returncode == 0 and done.stdout.strip() else {}
        return (values.get("dockit.lastOperation"), values.get("dockit.targetFormats") or {},
                bool(values.get("appLifecycle.configuration.enabled", False)))

    def envelope(self, name, operation, targets):
        """当前配置导出一份,改成给定的操作与目标格式,存成待导入的文件。"""
        path = self.root / name
        path.unlink(missing_ok=True)
        self.call("config", "export", "-o", str(path))
        body = json.loads(path.read_text())
        body["values"].update({KEYS[0]: operation, KEYS[1]: targets})
        path.write_text(json.dumps(body))
        return path

    def test_help_lines_are_the_shared_layers_own_and_reach_the_top_level_help(self):
        shared = self.direct("config", "--help")
        self.assertEqual(shared.returncode, 0)
        lines = shared.stdout.splitlines()
        reads = lines[lines.index("读（不写任何文件或状态）:") + 1:lines.index("写:")]
        writes = lines[lines.index("写:") + 1:next(i for i, line in enumerate(lines) if line.startswith("--json："))]
        self.assertEqual((len(reads), len(writes)), (2, 3))
        top = self.dockit("--help").stdout
        for line in reads + writes:
            self.assertIn("\n" + line + "\n", top)      # 顶层帮助里行首列出,文字就是共用层自己的那一行
        read_section = top.split("读命令")[1].split("\n\n")[0]
        write_section = top.split("写命令")[1].split("\n\n")[0]
        self.assertTrue(all(line in read_section for line in reads) and all(line in write_section for line in writes))
        self.assertNotIn("config export", read_section)   # 导出会写出你指定的文件,不在「不写任何文件」那一节
        no_command = next(line for line in lines if line.startswith("暂无命令：")).split("：", 1)[1]
        self.assertIn(no_command, top.split("暂无命令:")[1])
        self.assertNotIn("升级到新版", top.split("仅在窗口中")[1].split("暂无命令:")[0])
        forwarded = self.dockit("config", "--help")
        self.assertEqual((forwarded.returncode, forwarded.stdout), (0, shared.stdout))
        # 没装 App 时后端自己给出的那份帮助,与编好的程序逐字相同。
        alone = self.dockit("config", "--help", env=dict(self.env, DOCKIT_APP_BUNDLE=str(self.root / "absent.app")))
        self.assertEqual((alone.returncode, alone.stdout), (0, shared.stdout))
        self.assertIn("退出码", shared.stdout)

    def test_words_output_and_exit_code_pass_through_unchanged(self):
        cases = ((("config", "status", "--json"), 0), (("config", "status"), 0), (("config", "--json"), 0),
                 (("config", "bogus", "--json"), 2), (("config", "status", "--no-such", "--json"), 2), (("config", "bogus"), 2),
                 (("update", "--json"), 2), (("config", "export", "--json"), 2), (("config", "sync", "maybe", "--json"), 2),
                 (("config", "sync", "on", "--json"), 2), (("config", "import", str(self.root / "absent.json"), "--yes", "--json"), 1),
                 (("update", "check", "--json"), 1), (("update", "check"), 1))
        for words, code in cases:
            ours, theirs = self.dockit(*words), self.direct(*words)
            self.assertEqual((ours.returncode, ours.stdout, ours.stderr), (theirs.returncode, theirs.stdout, theirs.stderr), words)
            self.assertEqual(ours.returncode, code, (words, ours.stdout, ours.stderr))
            if "--json" in words:
                body = json.loads(ours.stdout)
                self.assertIs(body["ok"], code == 0)
                self.assertEqual(body["command"].split()[0], words[0])
                if code:
                    self.assertTrue(body["error"]["code"] and body["error"]["message"])
            elif code:
                self.assertEqual(ours.stdout, "")  # 文本模式的错误走 stderr
                self.assertTrue(ours.stderr.strip())
        self.assertEqual(self.call("config", "bogus", expect=2)["error"]["code"], "usage")
        self.assertEqual(self.call("config", "status", "--no-such", expect=2)["error"]["code"], "usage")
        self.assertEqual(self.call("config", "sync", "on", expect=2)["error"]["code"], "confirmation_required")
        self.assertEqual(self.call("update", "check", expect=1)["error"]["code"], "check_incomplete")
        # 相对路径按敲命令的目录解析,不是 App 程序所在的目录。
        done = self.dockit("config", "export", "-o", "here.json", "--json", cwd=self.root)
        self.assertEqual((done.returncode, json.loads(done.stdout)["path"]), (0, str(self.root / "here.json")))
        # 原有命令照旧:仍由 Python 后端自己回答。
        self.assertEqual(json.loads(self.dockit("settings", "--json").stdout)["domain"], self.suite)

    def test_the_command_reads_and_writes_the_windows_own_settings(self):
        status = self.call("config", "status")
        self.assertEqual((status["command"], status["has_settings"], status["sync_enabled"], status["problem"]),
                         ("config status", True, False, None))
        self.assertEqual(status["keys"], KEYS)
        self.assertFalse(self.support.exists() or self.cloud.exists())  # 读命令什么都不写
        exported = self.root / "out.json"
        exported.unlink(missing_ok=True)
        first = self.call("config", "export", "-o", str(exported))
        envelope = json.loads(exported.read_text())
        self.assertEqual((envelope["product"], first["bytes"], first["keys"]), (PRODUCT, exported.stat().st_size, KEYS))
        self.assertEqual((envelope["values"][KEYS[0]], envelope["values"][KEYS[1]]), (self.others[0], {self.op: self.targets[0]}))
        self.assertEqual(self.call("config", "export", "-o", str(exported), expect=2)["error"]["code"], "file_exists")
        self.assertEqual(json.loads(self.dockit("config", "export", "-o", "-").stdout)["values"], envelope["values"])
        # dockit settings 与 config 是同一份:settings 读到的就是导出里的两个键。
        settings = json.loads(self.dockit("settings", "--json").stdout)["settings"]
        self.assertEqual(settings, {"last_operation": self.others[0], "target_formats": {self.op: self.targets[0]}})

        incoming = self.envelope("in.json", self.others[1], {self.op: self.targets[1]})
        self.assertEqual(self.call("config", "import", str(incoming), expect=2)["error"]["code"], "confirmation_required")
        self.assertEqual(self.stored(), (self.others[0], {self.op: self.targets[0]}, False))
        foreign = self.root / "foreign.json"
        foreign.write_text(json.dumps(dict(json.loads(incoming.read_text()), product="someone.else")))
        self.assertEqual(self.call("config", "import", str(foreign), "--yes", expect=1)["error"]["code"], "import_rejected")
        self.assertEqual(self.stored(), (self.others[0], {self.op: self.targets[0]}, False))

        imported = self.call("config", "import", str(incoming), "--yes")
        self.assertTrue(imported["imported"] and "sync" not in imported)
        self.assertEqual(self.stored(), (self.others[1], {self.op: self.targets[1]}, False))
        settings = json.loads(self.dockit("settings", "--json").stdout)["settings"]
        self.assertEqual(settings, {"last_operation": self.others[1], "target_formats": {self.op: self.targets[1]}})
        self.assertEqual(len(list((self.support / PRODUCT / "Backups").iterdir())), 1)
        self.assertFalse(self.cloud.exists())  # 同步关着:什么都不出本机
        # 反过来:dockit settings set 写的值,config 导出里看得到。
        done = self.dockit("settings", "set", "target_formats." + self.op, self.targets[0], "--json")
        self.assertEqual(done.returncode, 0, done.stdout + done.stderr)
        self.assertEqual(json.loads(self.dockit("config", "export", "-o", "-").stdout)["values"][KEYS[1]], {self.op: self.targets[0]})

        dry = self.call("config", "sync", "on", "--dry-run")
        self.assertTrue(dry["dry_run"] and dry["would_change"] and not self.call("config", "status")["sync_enabled"])
        on = self.call("config", "sync", "on", "--yes")
        mirrored = json.loads((self.cloud / (PRODUCT + ".json")).read_text())
        self.assertEqual((on["changed"], on["sync_enabled"], on["check_with"], mirrored["values"][KEYS[0]]),
                         (True, True, "dockit config status", self.others[1]))
        self.assertTrue(self.call("config", "status")["sync_enabled"])
        self.assertIs(self.call("config", "sync", "on", "--yes")["changed"], False)
        self.assertIs(self.call("config", "sync", "off", "--yes")["sync_enabled"], False)

    def test_update_check_reads_only_the_isolated_release_record(self):
        missing = self.call("update", "check", expect=1)
        self.assertEqual((missing["error"]["code"], missing["current"], missing["source"]),
                         ("check_incomplete", {"version": self.version, "build": self.build}, {"kind": "private_cloud", "channel": CHANNEL}))
        feed = self.cloud / "TianliApps/Updates" / self.bundle_id / CHANNEL
        feed.mkdir(parents=True)

        def publish(version, build):
            (feed / "release.json").write_text(json.dumps({"version": version, "build": build, "bundle_id": self.bundle_id, "channel": CHANNEL,
                                                           "filename": f"DocKit-{version}.zip", "sha256": "a" * 64, "size_bytes": 10}))
            return self.call("update", "check")

        newer = publish("99.0", "9")
        self.assertEqual((newer["state"], newer["update_available"], newer["latest"]["version"]), ("update_available", True, "99.0"))
        self.assertIn("配置与更新…", newer["upgrade"]["how"])
        self.assertIn("有新版 99.0 (9)", self.dockit("update", "check").stdout)
        same = publish(self.version, self.build)
        self.assertEqual((same["state"], same["update_available"], same["upgrade"]["button"]), ("up_to_date", False, None))
        self.assertEqual(publish("0.1", "1")["state"], "ahead_of_channel")
        self.assertEqual(sorted(path.name for path in feed.iterdir()), ["release.json"])  # 没有下载,没有安装

    def start_app(self, live):
        state = self.root / f"app-{uuid.uuid4().hex}.json"
        complaints = state.with_suffix(".err")
        with complaints.open("wb") as errors:
            process = subprocess.Popen([str(self.binary), "--lifecycle-follow-probe", str(state)], env=live,
                                       stdin=subprocess.DEVNULL, stdout=subprocess.DEVNULL, stderr=errors)
        self.followers.append(process)
        deadline = time.monotonic() + 60   # 第一次要经 uv 启动后端读操作目录
        while not state.exists() and time.monotonic() < deadline and process.poll() is None:
            time.sleep(0.05)
        self.assertTrue(state.exists(), complaints.read_text() or "the app did not report")

        def seen():
            for _ in range(40):
                try:
                    return json.loads(state.read_text())
                except (OSError, ValueError):
                    time.sleep(0.02)
            raise AssertionError("app state unreadable")

        def reaches(test, seconds=6.0):
            deadline = time.monotonic() + seconds
            while time.monotonic() < deadline:
                if test(seen()):
                    return True
                time.sleep(0.05)
            return False

        return process, seen, reaches

    def stop_app(self, process):
        process.terminate()
        process.wait(timeout=10)
        self.followers.remove(process)

    def test_a_running_app_follows_the_command_and_never_writes_the_old_value_back(self):
        live = dict(self.env, APP_LIFECYCLE_FOLLOW_CHANNEL="test." + uuid.uuid4().hex)
        process, seen, reaches = self.start_app(live)

        def holds(want, seconds=1.0):
            """App 手里的开关、窗口里那个开关控件、存下来的开关(另起进程读)都停在这个值上。"""
            deadline = time.monotonic() + seconds
            while time.monotonic() < deadline:
                now = seen()
                if now["enabled"] is not want or now["window_switch"] is not want or self.call("config", "status")["sync_enabled"] is not want:
                    return False
                time.sleep(0.05)
            return True

        first = seen()
        # 它是作为 App 在运行,而且屏幕上什么都没有。
        self.assertEqual((first["policy_prohibited"], first["windows_on_screen"]), (True, 0))
        self.assertTrue(all(first["window_built"].values()), first["window_built"])  # 共用窗口,按菜单项的构造方式建出
        self.assertEqual((first["enabled"], first["window_switch"], first["status"]), (False, False, OFF))
        self.assertEqual(first["selected_operation"], self.others[0])  # 主视图按记住的操作选中
        for attempt in range(3):
            on = self.call("config", "sync", "on", "--yes", env=live)
            self.assertTrue(on["changed"] and on["app_running"], on)  # 命令看到了运行中的 App
            self.assertTrue(reaches(lambda s: s["enabled"] is True and s["window_switch"] is True and s["status"] != OFF), f"follows sync on ({attempt + 1})")
            self.assertTrue(holds(True), f"sync on is not written back ({attempt + 1})")
            self.call("config", "sync", "off", "--yes", env=live)
            self.assertTrue(reaches(lambda s: s["enabled"] is False and s["window_switch"] is False and s["status"] == OFF), f"follows sync off ({attempt + 1})")
            self.assertTrue(holds(False), f"sync off is not written back ({attempt + 1})")
        # 两条命令背靠背:App 对第一条做了什么,都不能在第二条返回之后把它撤销。
        for attempt in range(3):
            self.call("config", "sync", "on", "--yes", env=live)
            self.call("config", "sync", "off", "--yes", env=live)
            self.assertTrue(reaches(lambda s: s["enabled"] is False and s["window_switch"] is False and s["status"] == OFF), f"settles off after on, off ({attempt + 1})")
            self.assertTrue(holds(False, 1.5), f"on, off back to back stays off ({attempt + 1})")
        self.call("config", "sync", "on", "--yes", env=live)
        self.call("config", "sync", "off", "--yes", env=live)
        self.call("config", "sync", "on", "--yes", env=live)
        self.assertTrue(reaches(lambda s: s["enabled"] is True and s["window_switch"] is True and s["status"] != OFF), "settles on after on, off, on")
        self.assertTrue(holds(True, 1.5), "on, off, on back to back stays on")
        self.call("config", "sync", "off", "--yes", env=live)
        self.assertTrue(reaches(lambda s: s["enabled"] is False and s["window_switch"] is False and s["status"] == OFF), "back to off")
        # 导入:App 重读设置(主视图换到导入的操作与目标),开关不动,导入值不被 App 写回旧值。
        before = seen()
        incoming = self.envelope("follow.json", self.op, {self.op: self.targets[1]})
        self.call("config", "import", str(incoming), "--yes", env=live)
        want = (self.op, {self.op: self.targets[1]}, False)
        self.assertEqual(self.stored(), want)
        self.assertTrue(reaches(lambda s: s["changes"] > before["changes"] and s["selected_operation"] == self.op
                                and s["selected_target"] == self.targets[1]), f"the main view re-reads imported settings: {seen()}")
        deadline = time.monotonic() + 1.5
        while time.monotonic() < deadline:
            self.assertEqual(self.stored(), want, "the running app never stores the old settings back over an import")
            time.sleep(0.05)
        self.assertTrue(holds(False, 0.6), "an import leaves the switch alone")
        last = seen()
        self.assertEqual((last["policy_prohibited"], last["windows_on_screen"]), (True, 0))
        self.assertGreater(last["tick"], first["tick"])
        self.stop_app(process)
        if not IN_PLACE:   # 就地跑时 bundle id 是真的,本人自己开着的 DocKit 也算"在运行"
            self.assertIs(self.call("config", "status")["app_running"], False)

    def test_imports_made_while_sync_is_on_and_the_app_is_running_stay(self):
        live = dict(self.env, APP_LIFECYCLE_FOLLOW_CHANNEL="test." + uuid.uuid4().hex)
        process, seen, reaches = self.start_app(live)
        self.call("config", "sync", "on", "--yes", env=live)
        self.assertTrue(reaches(lambda s: s["enabled"] is True and s["window_switch"] is True), "the app follows sync on")
        cloud = self.cloud / (PRODUCT + ".json")

        def stays(operation, targets, seconds=1.5):
            """从命令返回起:另起进程读到的偏好、云端那份,一直是导入值;开关一直开着。"""
            deadline = time.monotonic() + seconds
            while time.monotonic() < deadline:
                self.assertEqual(self.stored(), (operation, targets, True))
                mirrored = json.loads(cloud.read_text())["values"]
                self.assertEqual((mirrored[KEYS[0]], mirrored[KEYS[1]]), (operation, targets))
                time.sleep(0.05)

        steps = ((self.op, {self.op: self.targets[1]}), (self.others[1], {self.op: self.targets[1]}),
                 (self.others[1], {self.op: self.targets[0]}))
        for number, (operation, targets) in enumerate(steps, 1):
            done = self.call("config", "import", str(self.envelope(f"sync-{number}.json", operation, targets)), "--yes", env=live)
            self.assertTrue(done["sync"]["completed"], done)
            stays(operation, targets)
            self.assertTrue(reaches(lambda s: s["selected_operation"] == operation and s["app_reads_targets"] == targets),
                            f"the app re-reads import {number}: {seen()}")
        # 两次导入背靠背:后一次返回之后,存下来的一直是后一次的值。
        one = self.envelope("pair-1.json", self.others[2], {self.op: self.targets[1]})
        two = self.envelope("pair-2.json", self.op, {self.op: self.targets[0]})
        self.call("config", "import", str(one), "--yes", env=live)
        self.call("config", "import", str(two), "--yes", env=live)
        stays(self.op, {self.op: self.targets[0]})
        self.assertTrue(reaches(lambda s: s["selected_operation"] == self.op and s["selected_target"] == self.targets[0]
                                and s["enabled"] is True and s["window_switch"] is True), f"settles on the last import: {seen()}")
        stays(self.op, {self.op: self.targets[0]}, 1.0)
        # 关掉同步之后,导入值仍在。
        self.call("config", "sync", "off", "--yes", env=live)
        self.assertTrue(reaches(lambda s: s["enabled"] is False and s["window_switch"] is False), "the app follows sync off")
        self.assertEqual(self.stored(), (self.op, {self.op: self.targets[0]}, False))
        self.assertEqual(seen()["windows_on_screen"], 0)
        self.stop_app(process)

    def test_refusals(self):
        # 隔离运行却没说产品自己的设置在哪:在读任何东西之前就拒绝。
        partial = {key: value for key, value in self.env.items() if key != "DOCKIT_LIFECYCLE_SUITE"}
        self.assertEqual(self.call("config", "status", expect=1, env=partial)["error"]["code"], "isolation_incomplete")
        for suite in (PRODUCT, self.bundle_id, str(self.root / "path-based-domain")):
            self.assertEqual(self.call("config", "status", expect=1, env=dict(self.env, DOCKIT_LIFECYCLE_SUITE=suite))["error"]["code"],
                             "isolation_incomplete", suite)
        # 探针只为本测试存在:隔离运行之外,程序在创建 NSApplication 之前就退出。
        outside = {key: value for key, value in self.env.items() if not key.startswith(("APP_LIFECYCLE_", "DOCKIT_LIFECYCLE_"))}
        probe = subprocess.run([str(self.binary), "--lifecycle-follow-probe", str(self.root / "never.json")], env=outside,
                               capture_output=True, text=True, stdin=subprocess.DEVNULL, timeout=30)
        self.assertEqual((probe.returncode, (self.root / "never.json").exists()), (64, False))
        # 没有 App 可交:同一个形状说明原因。
        alone = dict(self.env, DOCKIT_APP_BUNDLE=str(self.root / "absent.app"))
        self.assertEqual(self.call("config", "status", expect=1, env=alone)["error"]["code"], "app_missing")
        text = self.dockit("update", "check", env=alone)
        self.assertEqual((text.returncode, text.stdout), (1, ""))
        self.assertIn("App 可执行文件", text.stderr)

    def test_an_app_without_the_command_layer_is_never_started(self):
        """旧版程序不认 config / update,会把它当普通启动打开窗口:Info.plist 没声明这组命令字的包,一律不转调。"""
        marker = self.root / "started"
        marker.unlink(missing_ok=True)
        stub = self.root / "stub"
        stub.write_text(f"#!/bin/sh\ntouch '{marker}'\n")
        stub.chmod(0o755)
        for name, verbs in (("Old.app", None), ("Partial.app", ("update",))):
            old = dict(self.env, DOCKIT_APP_BUNDLE=str(bundle(self.root, stub, name, verbs)))
            refused = self.call("config", "status", expect=1, env=old)
            self.assertEqual(refused["error"]["code"], "app_outdated")
            text = self.dockit("config", "sync", "on", "--yes", env=old)
            self.assertEqual((text.returncode, text.stdout), (1, ""))
            self.assertFalse(marker.exists(), name)
        self.assertEqual(self.dockit("update", "check", "--json", env=old).returncode, 0)  # 声明了的词照常转调
        self.assertTrue(marker.exists())

    @unittest.skipUnless(HEADQUARTERS.is_file(), "shared source not on this machine")
    def test_the_command_layer_is_the_shared_source_byte_for_byte(self):
        self.assertTrue((MAC / "Sources/AppLifecycleCLI.swift").read_bytes() == HEADQUARTERS.read_bytes(),
                        f"Sources/AppLifecycleCLI.swift differs from the shared source; copy {HEADQUARTERS} over it and build again")


ASSEMBLED = os.environ.get("DOCKIT_APP")


@unittest.skipUnless(ASSEMBLED, "set DOCKIT_APP to an assembled DocKit.app (build.sh does after assembly)")
class AssembledBundleTests(unittest.TestCase):
    """装进包的那条链:包内的 bin/dockit 把 config / update 交给它自己所在包的 App 程序。
    只跑读命令,用一次性偏好域与临时目录;包自己真实的设置不打开。"""

    def setUp(self):
        self.folder = tempfile.TemporaryDirectory(prefix="dockit-lifecycle-app-")
        self.addCleanup(self.folder.cleanup)
        self.root = Path(self.folder.name).resolve()
        self.app = Path(ASSEMBLED).resolve()
        self.suite = BUNDLE + "." + uuid.uuid4().hex
        self.env = dict(os.environ, APP_LIFECYCLE_SUPPORT_DIR=str(self.root / "support"), APP_LIFECYCLE_CLOUD_DIR=str(self.root / "cloud"),
                        DOCKIT_LIFECYCLE_SUITE=self.suite, DOCKIT_CACHE_DIR=str(self.root / "cache"))
        for name in ("APP_LIFECYCLE_FOLLOW_CHANNEL", "DOCKIT_APP_BUNDLE", "DOCKIT_ENTRY_APP"):
            self.env.pop(name, None)   # 隔离且没有频道:不向任何运行中的 App 发信号;App 由入口自己所在的包决定
        self.addCleanup(lambda: sweep(BUNDLE + "."))
        self.addCleanup(lambda: subprocess.run(["/usr/bin/defaults", "delete", self.suite], capture_output=True, timeout=30))

    def entry(self, *words):
        return subprocess.run([str(self.app / "Contents/Resources/bin/dockit"), *words], env=self.env,
                              capture_output=True, text=True, stdin=subprocess.DEVNULL, timeout=90)

    def test_the_packaged_entry_forwards_to_its_own_bundle(self):
        info = plistlib.loads((self.app / "Contents/Info.plist").read_bytes())
        self.assertEqual(info.get("DocKitCommandVerbs"), ["config", "update"])
        binary = self.app / "Contents/MacOS" / info["CFBundleExecutable"]
        status = self.entry("config", "status", "--json")
        body = json.loads(status.stdout)
        self.assertEqual((status.returncode, body["ok"], body["command"], body["sync_enabled"], body["keys"]),
                         (0, True, "config status", False, []))
        direct = subprocess.run([str(binary), "config", "status", "--json"], env=self.env,
                                capture_output=True, text=True, stdin=subprocess.DEVNULL, timeout=90)
        self.assertEqual((status.returncode, status.stdout, status.stderr), (direct.returncode, direct.stdout, direct.stderr))
        # 经软链启动(装机后 ~/.local/bin/dockit 就是这样)也认得出自己所在的包。
        link = self.root / "dockit"
        link.symlink_to(self.app / "Contents/Resources/bin/dockit")
        linked = subprocess.run([str(link), "config", "status", "--json"], env=self.env, capture_output=True, text=True,
                                stdin=subprocess.DEVNULL, timeout=90)
        self.assertEqual((linked.returncode, linked.stdout), (status.returncode, status.stdout))
        wrong = self.entry("config", "status", "--no-such", "--json")
        self.assertEqual((wrong.returncode, json.loads(wrong.stdout)["error"]["code"]), (2, "usage"))
        # 窗口显示的版本、构建号与 bundle id 来自命令所在的包。
        feed = self.root / "cloud/TianliApps/Updates" / info["CFBundleIdentifier"] / CHANNEL
        feed.mkdir(parents=True)
        (feed / "release.json").write_text(json.dumps({"version": info["CFBundleShortVersionString"], "build": info["CFBundleVersion"],
                                                       "bundle_id": info["CFBundleIdentifier"], "channel": CHANNEL,
                                                       "filename": "DocKit.zip", "sha256": "a" * 64, "size_bytes": 10}))
        check = self.entry("update", "check", "--json")
        body = json.loads(check.stdout)
        self.assertEqual((check.returncode, body["state"], body["current"]),
                         (0, "up_to_date", {"version": info["CFBundleShortVersionString"], "build": info["CFBundleVersion"]}))
        top = self.entry("--help").stdout
        self.assertIn("\n  config status ", top)
        self.assertIn("\n  update check ", top)


if __name__ == "__main__":
    unittest.main()
