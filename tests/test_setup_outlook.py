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

TENANT = "99999999-8888-7777-6666-555555555555"

# Stands in for the Azure CLI: records each call, answers the queries.
FAKE_AZ = """#!/bin/sh
echo "$*" >> "$FAKE_AZ_LOG"
case "$*" in
  "account show --query user.name"*)
    [ -n "$FAKE_AZ_SIGNED_OUT" ] && [ ! -e "$FAKE_AZ_LOG.in" ] && exit 1
    echo admin@contoso.example ;;
  "account show --query tenantId"*) echo "%s" ;;
  login*) touch "$FAKE_AZ_LOG.in" ;;
  "ad app list"*) echo "$FAKE_AZ_EXISTING" ;;
  "ad app create"*) [ -n "$FAKE_AZ_CREATE_FAIL" ] && exit 1; echo "%s" ;;
  "ad sp show"*) echo "scope-id-from-microsoft" ;;
  "ad app permission admin-consent"*) exit "${FAKE_AZ_CONSENT_FAIL:-0}" ;;
esac
exit 0
""" % (TENANT, CLIENT)

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
        self.bin_dir = bin_dir
        self.az_log = self.root / "az.log"
        self.log = self.root / "login.json"
        self.config = self.root / "home" / ".config" / "omarchy" / "calendar-sync.json"
        self.env = dict(os.environ, HOME=str(self.root / "home"), FAKE_LOGIN_LOG=str(self.log),
                        FAKE_AZ_LOG=str(self.az_log), OUTLOOK_CONSENT_WAIT="0",
                        PATH="%s:%s" % (bin_dir, os.environ["PATH"]))

    def run_setup(self, answers, *flags, **env):
        # By hand, whether or not this machine has the Azure CLI.
        return self.run_with(answers, "--by-hand", *flags, **env)

    def run_with_az(self, answers, *flags, **env):
        _executable(self.bin_dir / "az", FAKE_AZ)
        return self.run_with(answers, *flags, **env)

    def az_calls(self):
        return self.az_log.read_text().splitlines() if self.az_log.exists() else []

    def run_with(self, answers, *flags, **env):
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

    def test_the_azure_cli_registers_the_app_and_nothing_is_typed(self):
        result = self.run_with_az("")
        self.assertEqual(result.returncode, 0, result.stdout + result.stderr)
        self.assertEqual(self.read_config()["graph"], {"clientId": CLIENT, "tenant": TENANT})
        self.assertEqual(json.loads(self.log.read_text())[:5],
                         ["login", "--client-id", CLIENT, "--tenant", TENANT])
        calls = "\n".join(self.az_calls())
        self.assertIn("ad app create --display-name Omarchy Calendar --sign-in-audience AzureADMyOrg "
                      "--is-fallback-public-client true --public-client-redirect-uris http://localhost", calls)
        self.assertIn("--api-permissions scope-id-from-microsoft=Scope", calls)
        self.assertIn("ad app permission admin-consent --id " + CLIENT, calls)
        self.assertNotIn("\nlogin", calls)
        self.assertNotIn("Application (client) ID", result.stdout)

    def test_the_azure_cli_signs_in_first_when_it_is_signed_out(self):
        result = self.run_with_az("", FAKE_AZ_SIGNED_OUT="1")
        self.assertEqual(result.returncode, 0, result.stdout + result.stderr)
        self.assertIn("login --allow-no-subscriptions --only-show-errors", self.az_calls())

    def test_an_existing_registration_is_brought_up_to_date_not_made_twice(self):
        result = self.run_with_az("", FAKE_AZ_EXISTING=CLIENT)
        self.assertEqual(result.returncode, 0, result.stdout + result.stderr)
        calls = "\n".join(self.az_calls())
        self.assertNotIn("ad app create", calls)
        self.assertIn("ad app update --id %s --is-fallback-public-client true" % CLIENT, calls)

    def test_consent_that_cannot_be_granted_leaves_it_to_the_sign_in(self):
        result = self.run_with_az("", FAKE_AZ_CONSENT_FAIL="1")
        self.assertEqual(result.returncode, 0, result.stdout + result.stderr)
        self.assertIn("the sign-in will ask for it", result.stdout)
        self.assertEqual(self.read_config()["graph"]["clientId"], CLIENT)

    def test_a_registration_that_fails_falls_back_to_the_steps_by_hand(self):
        result = self.run_with_az(CLIENT + "\ncontoso.example\n", FAKE_AZ_CREATE_FAIL="1")
        self.assertEqual(result.returncode, 0, result.stdout + result.stderr)
        self.assertIn("App registrations > New registration", result.stdout)
        self.assertEqual(self.read_config()["graph"], {"clientId": CLIENT, "tenant": "contoso.example"})

    def test_a_second_run_does_not_register_again(self):
        self.write_config({"backend": "graph", "write": True,
                           "graph": {"clientId": CLIENT, "tenant": TENANT}})
        result = self.run_with_az("\n\n")
        self.assertEqual(result.returncode, 0, result.stdout + result.stderr)
        self.assertEqual(self.az_calls(), [])

    def test_the_device_code_flag_reaches_the_sign_in(self):
        self.run_setup(CLIENT + "\ncontoso.example\n", "--device-code")
        self.assertIn("--device-code", json.loads(self.log.read_text()))


if __name__ == "__main__":
    unittest.main()
