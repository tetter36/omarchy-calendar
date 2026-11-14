"""Tests for `sync/setup --add-google`, run end to end against a fake gws."""

import json
import os
import pathlib
import shutil
import stat
import subprocess
import tempfile
import unittest

REPO = pathlib.Path(__file__).resolve().parent.parent
READ = "https://www.googleapis.com/auth/calendar.readonly"
WRITE = "https://www.googleapis.com/auth/calendar.events"

# Stands in for gws: records each call with the folder it was pointed at.
FAKE_GWS = """#!/bin/sh
echo "$GOOGLE_WORKSPACE_CLI_CONFIG_DIR|$*" >> "$FAKE_GWS_LOG"
case "$*" in
  --help) echo "Google Workspace CLI" ;;
  "auth login"*) exit "${FAKE_LOGIN_FAIL:-0}" ;;
  "auth status") echo "{\\"scopes\\": [$FAKE_GRANTED]}" ;;
esac
exit 0
"""

FAKE_SYNC = """#!/bin/sh
exit "${FAKE_FAIL:-0}"
"""


def _executable(path, text):
    path.write_text(text)
    path.chmod(path.stat().st_mode | stat.S_IEXEC)


class SetupAddGoogleTest(unittest.TestCase):
    def setUp(self):
        self.root = pathlib.Path(tempfile.mkdtemp())
        self.addCleanup(shutil.rmtree, self.root)
        sync = self.root / "sync"
        (sync / "systemd").mkdir(parents=True)
        shutil.copy(REPO / "sync" / "setup", sync / "setup")
        for unit in ("service", "timer"):
            (sync / "systemd" / ("omarchy-calendar-sync." + unit)).touch()
        _executable(sync / "omarchy-calendar-sync", FAKE_SYNC)
        bin_dir = self.root / "bin"
        bin_dir.mkdir()
        _executable(bin_dir / "gws", FAKE_GWS)
        _executable(bin_dir / "systemctl", "#!/bin/sh\nexit 0\n")
        home = self.root / "home"
        self.profile = home / ".config" / "gws-omarchy-calendar"
        self.config = home / ".config" / "omarchy" / "calendar-sync.json"
        self.gws_log = self.root / "gws.log"
        self.env = dict(os.environ, HOME=str(home), FAKE_GWS_LOG=str(self.gws_log),
                        FAKE_GRANTED='"%s", "%s"' % (READ, WRITE),
                        PATH="%s:%s" % (bin_dir, os.environ["PATH"]))
        self.env.pop("PROFILE", None)
        self.env.pop("GWS", None)

    def first_account(self, **config):
        self.profile.mkdir(parents=True)
        (self.profile / "client_secret.json").write_text('{"installed": {}}')
        self.config.parent.mkdir(parents=True, exist_ok=True)
        self.before = {"backend": "gws", "gwsPath": "/usr/bin/gws", "profile": str(self.profile),
                       "write": True, **config}
        self.config.write_text(json.dumps(self.before))

    def run_setup(self, *flags, **env):
        return subprocess.run(
            ["bash", str(self.root / "sync" / "setup"), "--add-google", *flags], input="",
            env={**self.env, **env}, capture_output=True, text=True, timeout=60, cwd=self.root)

    def read_config(self):
        return json.loads(self.config.read_text())

    def folders(self):
        return sorted(path.name for path in self.profile.parent.iterdir() if path.name.startswith("gws-"))

    def test_a_second_account_signs_in_to_a_folder_of_its_own_with_the_same_client(self):
        self.first_account()
        result = self.run_setup()
        self.assertEqual(result.returncode, 0, result.stdout + result.stderr)
        added = pathlib.Path(str(self.profile) + "-2")
        self.assertEqual(self.read_config(), {**self.before, "profiles": [str(added)]})
        self.assertEqual((added / "client_secret.json").read_text(), '{"installed": {}}')
        self.assertEqual(stat.S_IMODE((added / "client_secret.json").stat().st_mode), 0o600)
        self.assertEqual(stat.S_IMODE(added.stat().st_mode), 0o700)
        self.assertEqual(stat.S_IMODE(self.config.stat().st_mode), 0o600)
        self.assertIn("%s|auth login --scopes %s,%s" % (added, READ, WRITE),
                      self.gws_log.read_text().splitlines())

    def test_a_third_account_goes_beside_the_second_and_outlook_stays(self):
        self.first_account(backend=["gws", "graph"], graph={"clientId": "c", "tenant": "t"})
        self.run_setup()
        result = self.run_setup()
        self.assertEqual(result.returncode, 0, result.stdout + result.stderr)
        config = self.read_config()
        self.assertEqual(config["profiles"], [str(self.profile) + "-2", str(self.profile) + "-3"])
        self.assertEqual(config["backend"], ["gws", "graph"])

    def test_reading_only_stays_reading_only(self):
        self.first_account(write=False)
        self.run_setup()
        self.assertFalse(self.read_config()["write"])
        self.assertIn("auth login --scopes %s\n" % READ, self.gws_log.read_text())

    def test_before_the_first_account_it_changes_nothing(self):
        result = self.run_setup()
        self.assertEqual(result.returncode, 1)
        self.assertIn("without --add-google first", result.stdout)
        self.assertFalse(self.config.exists())

    def test_outlook_alone_is_not_a_first_google_account(self):
        self.config.parent.mkdir(parents=True)
        self.config.write_text(json.dumps({"backend": "graph", "write": True}))
        result = self.run_setup()
        self.assertEqual(result.returncode, 1)
        self.assertEqual(self.read_config(), {"backend": "graph", "write": True})

    def test_a_failed_sign_in_an_unticked_box_or_a_failed_sync_leaves_no_trace(self):
        self.first_account()
        for env in ({"FAKE_LOGIN_FAIL": "1"}, {"FAKE_GRANTED": '"%s"' % READ}, {"FAKE_FAIL": "1"}):
            result = self.run_setup(**env)
            self.assertEqual(result.returncode, 1, env)
            self.assertEqual(self.read_config(), self.before)
            self.assertEqual(self.folders(), ["gws-omarchy-calendar"])
            self.assertEqual([path.name for path in self.config.parent.iterdir()], ["calendar-sync.json"])


if __name__ == "__main__":
    unittest.main()
