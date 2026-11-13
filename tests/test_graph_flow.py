"""Outlook through the same sync and the same event command as Google."""

import contextlib
import io
import json
import tempfile
import unittest
from datetime import datetime, timezone
from pathlib import Path
from zoneinfo import ZoneInfo

from omarchy_calendar_sync import cli, config, contract, event_cli, graph, multi
from tests.test_graph import FakeHttp, client, outlook_event

NEW_YORK = ZoneInfo("America/New_York")
NOW = datetime(2026, 11, 13, 12, 0, tzinfo=timezone.utc)
JOIN = "https://teams.microsoft.com/l/meetup-join/19%3ameeting_abc%40thread.v2/0"
CALENDARS = {"value": [{
    "id": "CAL1", "name": "Calendar", "hexColor": "#0078d4", "isDefaultCalendar": True, "canEdit": True,
    "owner": {"address": "me@contoso.example"}, "defaultOnlineMeetingProvider": "teamsForBusiness",
}]}


class FakeGoogle:
    SOURCE_NAME = "gws"
    can_write = True

    def version(self):
        return (0, 13, 2)

    def check(self):
        pass

    def calendars(self):
        return [{"id": "me@example.com", "name": "me@example.com", "color": "#7bd148", "primary": True,
                 "writable": True, "defaultReminders": []}]

    def events(self, calendar_id, time_min, time_max):
        return [{"id": "g1", "iCalUID": "g1@google.com", "summary": "Lunch",
                 "start": {"dateTime": "2026-11-16T12:00:00-05:00"},
                 "end": {"dateTime": "2026-11-16T13:00:00-05:00"}}]

    def auth_hint(self, _cfg):
        return ""


class TestSyncBoth(unittest.TestCase):
    def test_google_and_outlook_land_in_one_file(self):
        http = FakeHttp((200, CALENDARS), (200, {"value": [
            outlook_event(isOnlineMeeting=True, onlineMeeting={"joinUrl": JOIN})]}))
        both = multi.Multi([FakeGoogle(), client(http)])
        cfg = {**config.DEFAULTS, "write": True}
        with tempfile.TemporaryDirectory() as tmp:
            out = Path(tmp) / "events.json"
            with contextlib.redirect_stdout(io.StringIO()):
                code = cli.run(both, cfg, NOW, out, NEW_YORK)
            doc = json.loads(out.read_text())
        self.assertEqual(code, cli.EXIT_OK)
        self.assertEqual(contract.validate(doc), [])
        self.assertEqual(doc["source"], "gws+graph/2")
        self.assertEqual([(row["title"], row["calendarId"]) for row in doc["events"]],
                         [("Planning", "graph:CAL1"), ("Lunch", "me@example.com")])
        self.assertEqual(doc["events"][0]["meetingUrl"], JOIN)
        self.assertEqual(doc["writableCalendars"], [
            {"id": "me@example.com", "name": "me@example.com", "color": "#7bd148"},
            {"id": "graph:CAL1", "name": "me@contoso.example", "color": "#0078d4",
             "lacks": ["color"], "video": "teams"},
        ])


class TestEventCommand(unittest.TestCase):
    def perform(self, raw, http):
        doc = {"version": 1, "syncedAt": NOW.isoformat(), "source": "gws+graph/2", "events": [],
               "writableCalendars": [
            {"id": "graph:CAL1", "name": "Work", "color": "#0078d4", "lacks": ["color"]}]}
        with tempfile.TemporaryDirectory() as tmp:
            path = Path(tmp) / "events.json"
            path.write_text(json.dumps(doc))
            both = multi.Multi([FakeGoogle(), client(http)])
            code, reply = event_cli.perform(raw, {**config.DEFAULTS, "write": True}, both, path, NEW_YORK,
                                            start_sync=lambda: None, now=NOW)
            return code, reply, json.loads(path.read_text())

    def test_the_form_opens_on_an_outlook_event(self):
        http = FakeHttp((200, outlook_event(isOnlineMeeting=True, onlineMeeting={"joinUrl": JOIN})))
        code, reply, _doc = self.perform(
            {"action": "get", "calendarId": "graph:CAL1", "eventId": "EV1"}, http)
        self.assertEqual(code, 0)
        form = reply["event"]
        self.assertEqual((form["title"], form["startTime"], form["endTime"]), ("Planning", "10:00", "11:00"))
        self.assertEqual(form["startDate"], "2026-11-16")
        self.assertTrue(form["meet"])
        self.assertEqual(form["reminders"]["overrides"], [{"method": "popup", "minutes": 15}])

    def test_saving_a_renamed_event_patches_the_title_alone(self):
        current = outlook_event()
        http = FakeHttp((200, current), (200, current), (200, outlook_event(subject="Renamed")))
        _code, reply, _doc = self.perform({"action": "get", "calendarId": "graph:CAL1", "eventId": "EV1"}, http)
        form = {**reply["event"], "title": "Renamed"}
        code, reply, doc = self.perform({"action": "update", "event": form}, http)
        self.assertEqual((code, reply), (0, {"ok": True, "eventId": "EV1"}))
        self.assertEqual(http.calls[-1][0], "PATCH")
        self.assertEqual(http.body(), {"subject": "Renamed"})
        self.assertEqual([row["title"] for row in doc["events"]], ["Renamed"])

    def test_a_new_event_with_video_asks_for_a_teams_meeting(self):
        http = FakeHttp((200, CALENDARS["value"][0]),
                        (201, outlook_event(id="NEW", isOnlineMeeting=True, onlineMeeting={"joinUrl": JOIN})))
        form = {"calendarId": "graph:CAL1", "title": "Sync", "startDate": "2026-11-17", "startTime": "09:00",
                "endDate": "2026-11-17", "endTime": "09:30", "meet": True}
        code, reply, doc = self.perform({"action": "create", "event": form}, http)
        self.assertEqual((code, reply), (0, {"ok": True, "eventId": "NEW"}))
        self.assertTrue(http.body()["isOnlineMeeting"])
        self.assertEqual(doc["events"][0]["meetingUrl"], JOIN)

    def test_turning_video_off_says_outlook_cannot(self):
        current = outlook_event(isOnlineMeeting=True, onlineMeeting={"joinUrl": JOIN})
        http = FakeHttp((200, current), (200, current), (200, current))
        _code, reply, _doc = self.perform({"action": "get", "calendarId": "graph:CAL1", "eventId": "EV1"}, http)
        code, reply, _doc = self.perform({"action": "update", "event": {**reply["event"], "meet": False}}, http)
        self.assertEqual((code, reply), (1, {"ok": False, "error": graph.KEEPS_TEAMS}))

    def test_an_expired_sign_in_names_the_outlook_setup(self):
        http = FakeHttp((401, {}), (400, {"error": "invalid_grant"}))
        code, reply, _doc = self.perform({"action": "get", "calendarId": "graph:CAL1", "eventId": "EV1"}, http)
        self.assertEqual((code, reply), (1, {"ok": False, "error": graph.SIGN_IN_EXPIRED}))

    def test_deleting_one_occurrence_deletes_that_id(self):
        http = FakeHttp((204, ""))
        code, _reply, _doc = self.perform(
            {"action": "delete", "calendarId": "graph:CAL1", "eventId": "OCC1", "recurringEventId": "MASTER"}, http)
        self.assertEqual(code, 0)
        self.assertEqual(http.calls[0][:2], ("DELETE", "https://graph.microsoft.com/v1.0/me/events/OCC1"))


if __name__ == "__main__":
    unittest.main()
