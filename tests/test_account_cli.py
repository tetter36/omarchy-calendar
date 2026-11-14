import json
import tempfile
import unittest
from pathlib import Path

from omarchy_calendar_sync import account_cli
from omarchy_calendar_sync.errors import SyncError

GRAPH = {"clientId": "11111111-2222-3333-4444-555555555555", "tenant": "contoso.example"}


class FakeClient:
    def __init__(self, calendars=None, error=None):
        self._calendars, self._error = calendars or [], error

    def calendars(self):
        if self._error:
            raise self._error
        return self._calendars


class TestConfigured(unittest.TestCase):
    def test_no_config_is_no_accounts(self):
        self.assertEqual(account_cli.configured({}), [])
        self.assertEqual(account_cli._raw("/nonexistent/calendar-sync.json"), {})

    def test_every_google_profile_and_outlook_entry_is_an_account(self):
        raw = {"backend": ["gws", "graph"], "profile": "/p/one", "profiles": ["/p/two", "/p/one"],
               "graph": [GRAPH, {**GRAPH, "account": "2"}]}
        self.assertEqual(
            [(one["type"], one["number"]) for one in account_cli.configured(raw)],
            [("google", 1), ("google", 2), ("outlook", 1), ("outlook", 2)])

    def test_an_older_google_config_that_names_no_backend_counts(self):
        self.assertEqual(len(account_cli.configured({"gwsPath": "/usr/bin/gws"})), 1)

    def test_a_backend_that_is_not_an_account_here_is_left_out(self):
        self.assertEqual(account_cli.configured({"backend": "ics", "graph": GRAPH}), [])


class TestList(unittest.TestCase):
    RAW = {"backend": ["gws", "graph"], "profile": "/p/one", "graph": GRAPH}

    def test_each_account_shows_its_address_and_calendar_count(self):
        clients = {
            "google": FakeClient([{"id": "holidays#x", "name": "Holidays"},
                                  {"id": "me@gmail.example", "name": "Me", "primary": True}]),
            "outlook": FakeClient([{"id": "graph:CAL1", "name": "me@contoso.example", "primary": True}]),
        }
        code, reply = account_cli.perform({"action": "list"}, self.RAW,
                                          client_for=lambda account, _raw: clients[account["type"]])
        self.assertEqual(code, 0)
        self.assertEqual([(one["type"], one["address"], one["calendars"], one["ok"]) for one in reply["accounts"]],
                         [("google", "me@gmail.example", 2, True), ("outlook", "me@contoso.example", 1, True)])

    def test_a_sign_in_that_no_longer_works_says_so_and_the_rest_still_list(self):
        clients = {"google": FakeClient(error=SyncError("Your Google sign-in expired.")),
                   "outlook": FakeClient([])}
        _code, reply = account_cli.perform({"action": "list"}, self.RAW,
                                           client_for=lambda account, _raw: clients[account["type"]])
        google, outlook = reply["accounts"]
        self.assertEqual((google["ok"], google["error"]), (False, "Your Google sign-in expired."))
        self.assertTrue(outlook["ok"])

    def test_the_reply_holds_no_settings(self):
        _code, reply = account_cli.perform({"action": "list"}, self.RAW,
                                           client_for=lambda *_: FakeClient([]))
        self.assertNotIn(GRAPH["clientId"], json.dumps(reply))
        self.assertNotIn("/p/one", json.dumps(reply))


class TestConnect(unittest.TestCase):
    def options(self, kind, raw):
        opened = []
        code, reply = account_cli.perform({"action": "connect", "type": kind}, raw, terminal=opened.append)
        return code, reply, opened

    def test_the_first_account_of_a_kind_runs_its_full_setup(self):
        self.assertEqual(self.options("google", {})[2], ["--write"])
        self.assertEqual(self.options("outlook", {"backend": "gws", "gwsPath": "gws"})[2], ["--outlook"])

    def test_a_further_account_only_signs_in(self):
        with tempfile.TemporaryDirectory() as profile:
            raw = {"backend": ["gws", "graph"], "profile": profile, "graph": GRAPH}
            # Without the first setup's OAuth client there is nothing to reuse.
            self.assertEqual(self.options("google", raw)[2], ["--write"])
            (Path(profile) / "client_secret.json").write_text("{}")
            self.assertEqual(self.options("google", raw)[2], ["--add-google"])
            self.assertEqual(self.options("outlook", raw)[2], ["--add-outlook"])

    def test_only_the_known_kinds_and_actions_are_taken(self):
        for request in ({"action": "connect", "type": "--write; rm -rf ~"}, {"action": "connect"},
                        {"action": "remove"}, "connect", None):
            opened = []
            code, reply = account_cli.perform(request, {}, terminal=opened.append)
            self.assertEqual((code, reply["ok"], opened), (1, False, []))

    def test_the_terminal_runs_the_setup_beside_this_module_with_that_one_option(self):
        spawned = []
        account_cli.open_terminal("--add-google", setup="/plug in/sync/setup",
                                  which=lambda name: "/usr/bin/" + name,
                                  spawn=lambda argv, **kwargs: spawned.append((argv, kwargs)))
        argv, kwargs = spawned[0]
        self.assertEqual(argv[:4], ["uwsm-app", "--", "xdg-terminal-exec", "--app-id=org.omarchy.terminal"])
        self.assertEqual(argv[-3:-1], ["bash", "-c"])
        self.assertTrue(argv[-1].startswith("'/plug in/sync/setup' --add-google; "))
        self.assertTrue(kwargs["start_new_session"])

    def test_no_terminal_hands_over_the_command_to_run(self):
        code, reply = account_cli.perform(
            {"action": "connect", "type": "outlook"}, {},
            terminal=lambda option: account_cli.open_terminal(option, which=lambda _name: None))
        self.assertEqual(code, 1)
        self.assertIn("sync/setup --outlook", reply["error"])


if __name__ == "__main__":
    unittest.main()
