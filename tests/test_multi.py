import unittest

from omarchy_calendar_sync import cli, config, multi
from omarchy_calendar_sync.errors import SyncError


class Fake:
    can_write = True

    def __init__(self, name, calendars, fail_check=None, fail_calendars=None):
        self.SOURCE_NAME = name
        self._calendars = calendars
        self.fail_check = fail_check
        self.fail_calendars = fail_calendars
        self.calls = []

    def version(self):
        return (1,)

    def check(self):
        if self.fail_check:
            raise self.fail_check

    def calendars(self):
        self.calls.append(("calendars",))
        if self.fail_calendars:
            raise self.fail_calendars
        return [dict(calendar) for calendar in self._calendars]

    def events(self, calendar_id, time_min, time_max):
        self.calls.append(("events", calendar_id))
        return [{"id": self.SOURCE_NAME + "-event"}]

    def get(self, calendar_id, event_id):
        self.calls.append(("get", calendar_id, event_id))
        return {"id": event_id}

    def create(self, calendar_id, body, send_updates="none"):
        self.calls.append(("create", calendar_id, send_updates))
        return {"id": "new"}

    def update(self, calendar_id, event_id, body, send_updates="none"):
        self.calls.append(("update", calendar_id, event_id))
        return {"id": event_id}

    def replace(self, calendar_id, event_id, resource, send_updates="none"):
        self.calls.append(("replace", calendar_id, event_id))
        return {"id": event_id}

    def delete(self, calendar_id, event_id, send_updates="none"):
        self.calls.append(("delete", calendar_id, event_id))

    def auth_hint(self, _cfg):
        return "sign in to " + self.SOURCE_NAME


class Prefixed(Fake):
    def owns(self, calendar_id):
        return calendar_id.startswith(self.SOURCE_NAME + ":")


def calendar(calendar_id, writable=True):
    return {"id": calendar_id, "name": calendar_id, "color": "#112233", "writable": writable}


def pair(**outlook):
    google = Fake("gws", [calendar("me@example.com")])
    graph = Prefixed("graph", [calendar("graph:CAL1")], **outlook)
    return google, graph, multi.Multi([google, graph])


class TestMulti(unittest.TestCase):
    def test_lists_every_backends_calendars_under_one_source(self):
        _google, _graph, both = pair()
        both.check()
        self.assertEqual([c["id"] for c in both.calendars()], ["me@example.com", "graph:CAL1"])
        self.assertEqual(both.SOURCE_NAME, "gws+graph")
        self.assertTrue(both.can_write)

    def test_events_come_from_the_backend_that_listed_the_calendar(self):
        google, graph, both = pair()
        both.calendars()
        self.assertEqual(both.events("graph:CAL1", "a", "b"), [{"id": "graph-event"}])
        self.assertEqual(both.events("me@example.com", "a", "b"), [{"id": "gws-event"}])
        self.assertNotIn(("events", "graph:CAL1"), google.calls)

    def test_a_write_without_a_sync_first_is_routed_by_the_id_alone(self):
        google, graph, both = pair()
        both.create("graph:CAL1", {}, send_updates="all")
        both.update("me@example.com", "E1", {})
        both.delete("graph:CAL1", "E2")
        self.assertEqual(graph.calls, [("create", "graph:CAL1", "all"), ("delete", "graph:CAL1", "E2")])
        self.assertEqual(google.calls, [("update", "me@example.com", "E1")])

    def test_two_backends_that_cannot_tell_their_ids_apart_are_listed_once(self):
        first = Fake("caldav", [calendar("https://a.example/cal/")])
        second = Fake("gws", [calendar("me@example.com")])
        both = multi.Multi([first, second])
        both.get("me@example.com", "E1")
        self.assertEqual(second.calls[-1], ("get", "me@example.com", "E1"))
        self.assertEqual(first.calls, [("calendars",)])

    def test_a_calendar_two_accounts_see_shows_once_under_one_that_may_write_it(self):
        reader = Fake("gws", [calendar("team@example.com", writable=False), calendar("a@example.com")])
        owner = Fake("gws", [calendar("team@example.com"), calendar("a@example.com")])
        both = multi.Multi([reader, owner])
        listed = both.calendars()
        self.assertEqual([(c["id"], c["writable"]) for c in listed],
                         [("team@example.com", True), ("a@example.com", True)])
        both.update("team@example.com", "E1", {})
        both.update("a@example.com", "E2", {})
        self.assertEqual(owner.calls[-1], ("update", "team@example.com", "E1"))
        self.assertEqual(reader.calls[-1], ("update", "a@example.com", "E2"))

    def test_a_calendar_no_backend_lists_is_refused(self):
        first = Fake("caldav", [])
        second = Fake("gws", [])
        with self.assertRaises(SyncError):
            multi.Multi([first, second]).get("nobody@example.com", "E1")

    def test_one_failed_sign_in_costs_only_its_own_calendars(self):
        _google, _graph, both = pair(fail_check=SyncError("expired"))
        both.check()
        self.assertEqual([c["id"] for c in both.calendars()], ["me@example.com"])
        self.assertEqual(both.auth_hint({}), "sign in to graph")

    def test_a_backend_that_fails_while_listing_is_skipped_too(self):
        _google, _graph, both = pair(fail_calendars=SyncError("down"))
        both.check()
        self.assertEqual([c["id"] for c in both.calendars()], ["me@example.com"])

    def test_every_backend_failing_fails_the_sync(self):
        google = Fake("gws", [], fail_check=SyncError("no gws"))
        graph = Prefixed("graph", [], fail_check=SyncError("expired"))
        with self.assertRaises(SyncError) as caught:
            multi.Multi([google, graph]).check()
        self.assertEqual(str(caught.exception), "no gws")

    def test_a_read_only_backends_calendars_are_never_writable(self):
        feed = Fake("ics", [calendar("https://x.example/a.ics")])
        feed.can_write = False
        graph = Prefixed("graph", [calendar("graph:CAL1")])
        both = multi.Multi([feed, graph])
        listed = {c["id"]: c["writable"] for c in both.calendars()}
        self.assertEqual(listed, {"https://x.example/a.ics": False, "graph:CAL1": True})
        with self.assertRaises(SyncError):
            both.create("https://x.example/a.ics", {})


class TestBuildClient(unittest.TestCase):
    GRAPH = {"clientId": "11111111-2222-3333-4444-555555555555", "tenant": "organizations"}

    def cfg(self, backend):
        return {**config.DEFAULTS, "backend": backend, "graph": self.GRAPH}

    def test_a_list_of_backends_builds_one_client_over_all(self):
        client = cli.build_client(self.cfg(["gws", "Graph", "gws"]))
        self.assertIsInstance(client, multi.Multi)
        self.assertEqual(client.SOURCE_NAME, "gws+graph")

    def test_several_google_accounts_are_several_clients_under_one_name(self):
        cfg = {**self.cfg(["gws", "graph"]), "profile": "/p/one", "profiles": ["/p/two", "/p/one"]}
        client = cli.build_client(cfg)
        self.assertEqual(client.SOURCE_NAME, "gws+graph")
        self.assertEqual([one.profile for one in client._clients[:-1]], ["/p/one", "/p/two"])
        alone = cli.build_client({**self.cfg("gws"), "profiles": ["/p/two"]})
        self.assertIsInstance(alone, multi.Multi)
        self.assertEqual(alone.SOURCE_NAME, "gws")
        self.assertIn("/p/two", alone._clients[1].auth_hint(cfg))

    def test_profiles_that_are_not_folder_paths_are_a_config_error(self):
        for profiles in ("/p/two", [""], [3]):
            with self.assertRaises(config.ConfigError):
                cli.build_client({**self.cfg("gws"), "profiles": profiles})

    def test_a_list_of_one_is_that_backend(self):
        self.assertEqual(cli.build_client(self.cfg(["graph"])).SOURCE_NAME, "graph")
        self.assertEqual(cli.build_client(self.cfg("graph")).SOURCE_NAME, "graph")

    def test_an_empty_or_unknown_list_is_a_config_error(self):
        for backend in ([], ["gws", "exchange"]):
            with self.assertRaises(config.ConfigError):
                cli.build_client(self.cfg(backend))

    def test_graph_without_a_client_id_is_a_config_error(self):
        with self.assertRaises(config.ConfigError):
            cli.build_client({**config.DEFAULTS, "backend": "graph"})


if __name__ == "__main__":
    unittest.main()
