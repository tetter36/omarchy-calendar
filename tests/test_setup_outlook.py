"""Tests for `sync/setup --outlook`, run end to end against a fake sign-in."""

import json
import os
import pathlib
import shutil
import stat
import subprocess
import tempfile
import unittest

REPO = pathlib.Path(__file__).resolve().parent.parent
CLIENT = "11111111-2222-3333-4444-555555555555"

# Stands in for the device code sign-in: records what it was asked.
FAKE_GRAPH = """import json, os, sys
with open(os.environ["FAKE_LOGIN_LOG"], "w") as stream:
    json.dump(sys.argv[1:], stream)
sys.exit(int(os.environ.get("FAKE_LOGIN_FAIL", "0")))
"""

FAKE_SYNC = """#!/bin/sh
exit "${FAKE_FAIL:-0}"
"""


def _executable(path, text):
    path.write_text(text)
    path.chmod(path.stat().st_mode | stat.S_IEXEC)


class SetupOutlookTest(unittest.TestCase):
    def setUp(self):
        self.root = pathlib.Path(tempfile.mkdtemp())
        self.addCleanup(shutil.rmtree, self.root)
        sync = self.root / "sync"
        (sync / "systemd").mkdir(parents=True)
        shutil.copy(REPO / "sync" / "setup", sync / "setup")
        package = sync / "omarchy_calendar_sync"
        package.mkdir()
        (package / "__init__.py").touch()
        (package / "graph.py").write_text(FAKE_GRAPH)
        for unit in ("service", "timer"):
            (sync / "systemd" / ("omarchy-calendar-sync." + unit)).touch()
        _executable(sync / "omarchy-calendar-sync", FAKE_SYNC)
        bin_dir = self.root / "bin"
        bin_dir.mkdir()
        _executable(bin_dir / "secret-tool", "#!/bin/sh\nexit 0\n")
        _executable(bin_dir / "systemctl", "#!/bin/sh\nexit 0\n")
        self.log = self.root / "login.json"
        self.config = self.root / "home" / ".config" / "omarchy" / "calendar-sync.json"
        self.env = dict(os.environ, HOME=str(self.root / "home"), FAKE_LOGIN_LOG=str(self.log),
                        PATH="%s:%s" % (bin_dir, os.environ["PATH"]))

    def run_setup(self, answers, *flags, **env):
        return subprocess.run(
            ["bash", str(self.root / "sync" / "setup"), "--outlook", *flags], input=answers,
            env={**self.env, **env}, capture_output=True, text=True, timeout=60,
            # python3 -m looks in the current folder first, and the real
            # package must not stand in for the fake sign-in.
            cwd=self.root)

    def write_config(self, config):
        self.config.parent.mkdir(parents=True, exist_ok=True)
        self.config.write_text(json.dumps(config))

    def read_config(self):
        return json.loads(self.config.read_text())

    def test_a_first_setup_reads_and_writes_outlook_alone(self):
        result = self.run_setup(CLIENT + "\ncontoso.example\n")
        self.assertEqual(result.returncode, 0, result.stdout + result.stderr)
        self.assertEqual(self.read_config(), {
            "backend": "graph", "graph": {"clientId": CLIENT, "tenant": "contoso.example"}, "write": True})
        self.assertEqual(stat.S_IMODE(self.config.stat().st_mode), 0o600)
        self.assertEqual(json.loads(self.log.read_text()),
                         ["login", "--client-id", CLIENT, "--tenant", "contoso.example", "--write"])

    def test_outlook_goes_beside_google_and_leaves_writing_as_it_was(self):
        self.write_config({"backend": "gws", "gwsPath": "/usr/bin/gws", "write": False,
                           "window": {"pastDays": 3}})
        result = self.run_setup(CLIENT + "\ncontoso.example\n")
        self.assertEqual(result.returncode, 0, result.stdout + result.stderr)
        config = self.read_config()
        self.assertEqual(config["backend"], ["gws", "graph"])
        self.assertEqual(config["graph"], {"clientId": CLIENT, "tenant": "contoso.example"})
        self.assertFalse(config["write"])
        self.assertEqual(config["window"], {"pastDays": 3})
        self.assertNotIn("--write", json.loads(self.log.read_text()))

    def test_a_config_from_the_google_setup_with_no_backend_named_still_keeps_google(self):
        self.write_config({"gwsPath": "/usr/bin/gws", "write": True})
        self.run_setup(CLIENT + "\ncontoso.example\n")
        self.assertEqual(self.read_config()["backend"], ["gws", "graph"])

    def test_a_second_run_offers_what_the_config_holds_and_adds_nothing_twice(self):
        self.write_config({"backend": ["caldav", "graph"], "write": True,
                           "graph": {"clientId": CLIENT, "tenant": "contoso.example"}})
        result = self.run_setup("\n\n")
        self.assertEqual(result.returncode, 0, result.stdout + result.stderr)
        config = self.read_config()
        self.assertEqual(config["backend"], ["caldav", "graph"])
        self.assertEqual(config["graph"], {"clientId": CLIENT, "tenant": "contoso.example"})

    def test_the_write_flag_turns_writing_on(self):
        self.write_config({"backend": "gws", "gwsPath": "/usr/bin/gws", "write": False})
        self.run_setup(CLIENT + "\ncontoso.example\n", "--write")
        self.assertTrue(self.read_config()["write"])
        self.assertIn("--write", json.loads(self.log.read_text()))

    def test_a_failed_sign_in_changes_nothing(self):
        before = {"backend": "gws", "gwsPath": "/usr/bin/gws", "write": True}
        self.write_config(before)
        result = self.run_setup(CLIENT + "\ncontoso.example\n", FAKE_LOGIN_FAIL="1")
        self.assertEqual(result.returncode, 1)
        self.assertEqual(self.read_config(), before)

    def test_a_failed_test_sync_changes_nothing(self):
        before = {"backend": "gws", "gwsPath": "/usr/bin/gws", "write": True}
        self.write_config(before)
        result = self.run_setup(CLIENT + "\ncontoso.example\n", FAKE_FAIL="1")
        self.assertEqual(result.returncode, 1)
        self.assertEqual(self.read_config(), before)
        self.assertEqual(sorted(path.name for path in self.config.parent.iterdir()), ["calendar-sync.json"])

    def test_no_client_id_or_no_tenant_changes_nothing(self):
        for answers in ("\n", CLIENT + "\n\n"):
            result = self.run_setup(answers)
            self.assertEqual(result.returncode, 1)
            self.assertFalse(self.config.exists())
            self.assertFalse(self.log.exists())


if __name__ == "__main__":
    unittest.main()
