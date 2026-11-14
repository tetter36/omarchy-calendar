import base64
import hashlib
import io
import json
import threading
import unittest
import urllib.error
import urllib.parse
import urllib.request
from datetime import date
from zoneinfo import ZoneInfo

from omarchy_calendar_sync import config, event_form, graph, normalize

NEW_YORK = ZoneInfo("America/New_York")
CLIENT = "11111111-2222-3333-4444-555555555555"
SETTINGS = {"clientId": CLIENT, "tenant": "contoso.example"}
CALENDAR = {"id": "graph:CAL1", "name": "Work", "color": "#0078d4"}


class FakeStore:
    def __init__(self, tokens=None):
        self.tokens = tokens
        self.saved = []

    def load(self):
        return self.tokens

    def save(self, tokens):
        self.tokens = tokens
        self.saved.append(tokens)


class FakeHttp:
    """Replays one reply per call, in order, and records every request."""

    def __init__(self, *replies):
        self.replies = list(replies)
        self.calls = []

    def __call__(self, method, url, headers, data):
        self.calls.append((method, url, headers, data))
        status, payload = self.replies.pop(0)
        return status, payload if isinstance(payload, str) else json.dumps(payload)

    def body(self, index=-1):
        return json.loads(self.calls[index][3])


def fresh():
    return FakeStore({"refresh_token": "R1", "access_token": "A1", "expires_at": 10_000})


def client(http, store=None, now=1_000):
    return graph.Graph(SETTINGS, local_tz=NEW_YORK, http=http, store=store or fresh(), clock=lambda: now)


def outlook_event(**changes):
    event = {
        "id": "EV1",
        "iCalUId": "UID1",
        "subject": "Planning",
        "body": {"contentType": "text", "content": "Agenda\r\nitems"},
        "start": {"dateTime": "2026-11-16T15:00:00.0000000", "timeZone": "UTC"},
        "end": {"dateTime": "2026-11-16T16:00:00.0000000", "timeZone": "UTC"},
        "isAllDay": False,
        "location": {"displayName": "Room 4"},
        "webLink": "https://outlook.office365.com/owa/?itemid=EV1",
        "showAs": "busy",
        "sensitivity": "normal",
        "isReminderOn": True,
        "reminderMinutesBeforeStart": 15,
        "type": "singleInstance",
        "isOrganizer": True,
        "organizer": {"emailAddress": {"address": "me@contoso.example"}},
        "attendees": [],
    }
    event.update(changes)
    return event


class TestSettings(unittest.TestCase):
    def test_a_missing_client_id_is_a_config_error(self):
        with self.assertRaises(config.ConfigError):
            graph.Graph({"tenant": "organizations"}, store=FakeStore(), http=FakeHttp())

    def test_a_tenant_cannot_carry_a_path(self):
        with self.assertRaises(config.ConfigError):
            graph.Graph({"clientId": CLIENT, "tenant": "a/b"}, store=FakeStore(), http=FakeHttp())

    def test_owns_only_its_own_calendar_ids(self):
        self.assertTrue(graph.Graph.owns("graph:CAL1"))
        self.assertFalse(graph.Graph.owns("me@example.com"))


class TestSignIn(unittest.TestCase):
    def test_a_kept_token_that_lasts_is_used_without_a_request(self):
        http = FakeHttp((200, {"value": []}))
        client(http).calendars()
        self.assertEqual(len(http.calls), 1)
        self.assertEqual(http.calls[0][2]["Authorization"], "Bearer A1")

    def test_a_token_about_to_end_is_renewed_and_kept(self):
        store = fresh()
        http = FakeHttp((200, {"access_token": "A2", "refresh_token": "R2", "expires_in": 3600}),
                        (200, {"value": []}))
        client(http, store, now=9_900).calendars()
        method, url, _headers, data = http.calls[0]
        self.assertEqual(url, "https://login.microsoftonline.com/contoso.example/oauth2/v2.0/token")
        sent = urllib.parse.parse_qs(data.decode())
        self.assertEqual(sent["refresh_token"], ["R1"])
        self.assertEqual(sent["client_id"], [CLIENT])
        self.assertEqual(store.tokens, {"refresh_token": "R2", "access_token": "A2", "expires_at": 13_500})
        self.assertEqual(http.calls[1][2]["Authorization"], "Bearer A2")

    def test_a_renewal_without_a_new_refresh_token_keeps_the_old_one(self):
        store = fresh()
        http = FakeHttp((200, {"access_token": "A2", "expires_in": 3600}), (200, {"value": []}))
        client(http, store, now=9_900).calendars()
        self.assertEqual(store.tokens["refresh_token"], "R1")

    def test_no_sign_in_names_the_setup(self):
        with self.assertRaises(graph.GraphSignInError) as caught:
            client(FakeHttp(), FakeStore(None)).check()
        self.assertIn("setup --outlook", str(caught.exception))

    def test_a_refused_renewal_is_an_expired_sign_in(self):
        http = FakeHttp((400, {"error": "invalid_grant"}))
        with self.assertRaises(graph.GraphSignInError) as caught:
            client(http, now=9_900).check()
        self.assertEqual(str(caught.exception), graph.SIGN_IN_EXPIRED)

    def test_a_401_renews_once_then_gives_up(self):
        http = FakeHttp((401, {}), (200, {"access_token": "A2", "expires_in": 3600}), (401, {}))
        with self.assertRaises(graph.GraphSignInError):
            client(http).calendars()
        self.assertEqual(len(http.calls), 3)

    def test_no_message_quotes_a_token(self):
        http = FakeHttp((500, {"error": {"message": "boom"}}))
        with self.assertRaises(graph.GraphError) as caught:
            client(http).calendars()
        self.assertNotIn("A1", str(caught.exception))
        self.assertNotIn("R1", str(caught.exception))

    def test_login_polls_until_the_sign_in_finishes(self):
        store = FakeStore()
        out = io.StringIO()
        http = FakeHttp(
            (200, {"device_code": "D", "user_code": "ABCD", "interval": 5, "expires_in": 900,
                   "message": "Open https://microsoft.com/devicelogin and enter ABCD"}),
            (400, {"error": "authorization_pending"}),
            (400, {"error": "slow_down"}),
            (200, {"access_token": "A", "refresh_token": "R", "expires_in": 3600}),
        )
        waits = []
        graph.login(CLIENT, "contoso.example", True, http=http, store=store, out=out,
                    sleep=waits.append, clock=lambda: 100)
        self.assertIn("ABCD", out.getvalue())
        self.assertEqual(waits, [5, 5, 10])
        asked = urllib.parse.parse_qs(http.calls[0][3].decode())["scope"][0]
        self.assertIn("Calendars.ReadWrite", asked)
        self.assertIn("offline_access", asked)
        self.assertEqual(store.tokens["refresh_token"], "R")

    def test_login_without_write_asks_to_read_only(self):
        http = FakeHttp((200, {"device_code": "D", "message": "m"}),
                        (200, {"access_token": "A", "refresh_token": "R", "expires_in": 3600}))
        graph.login(CLIENT, "organizations", False, http=http, store=FakeStore(), out=io.StringIO(),
                    sleep=lambda _s: None, clock=lambda: 100)
        asked = urllib.parse.parse_qs(http.calls[0][3].decode())["scope"][0]
        self.assertIn("Calendars.Read ", asked)

    def test_a_declined_sign_in_says_why(self):
        http = FakeHttp((200, {"device_code": "D", "message": "m"}),
                        (400, {"error": "authorization_declined",
                               "error_description": "AADSTS1: Declined. Trace ID: abc Correlation ID: def"}))
        with self.assertRaises(graph.GraphSignInError) as caught:
            graph.login(CLIENT, "organizations", True, http=http, store=FakeStore(), out=io.StringIO(),
                        sleep=lambda _s: None, clock=lambda: 100)
        self.assertIn("Declined.", str(caught.exception))
        self.assertNotIn("Trace", str(caught.exception))


class FakeCallback:
    """The browser coming back, without a browser."""

    redirect_uri = "http://localhost:4321"

    def __init__(self, answer):
        self.answer, self.closed = answer, False

    def __call__(self, state):
        self.state = state
        return self

    def wait(self, _seconds):
        return self.answer

    def close(self):
        self.closed = True


class TestBrowserSignIn(unittest.TestCase):
    def sign_in(self, callback, http, store=None, write=True):
        opened, out = [], io.StringIO()
        try:
            graph.login_browser(CLIENT, "contoso.example", write, http=http, store=store or FakeStore(),
                                out=out, open_browser=opened.append, callback=callback, clock=lambda: 100)
        finally:
            self.opened, self.printed = opened, out.getvalue()

    def test_the_browser_opens_at_microsoft_and_the_code_becomes_a_sign_in(self):
        store, callback = FakeStore(), FakeCallback({"code": "C"})
        http = FakeHttp((200, {"access_token": "A", "refresh_token": "R", "expires_in": 3600}))
        self.sign_in(callback, http, store)
        address = urllib.parse.urlsplit(self.opened[0])
        asked = {key: values[0] for key, values in urllib.parse.parse_qs(address.query).items()}
        self.assertEqual(address.netloc, "login.microsoftonline.com")
        self.assertEqual(address.path, "/contoso.example/oauth2/v2.0/authorize")
        self.assertEqual(asked["redirect_uri"], "http://localhost:4321")
        self.assertEqual(asked["state"], callback.state)
        self.assertEqual(asked["code_challenge_method"], "S256")
        self.assertIn("Calendars.ReadWrite", asked["scope"])
        self.assertIn("offline_access", asked["scope"])
        self.assertIn(self.opened[0], self.printed)
        sent = {key: values[0] for key, values in urllib.parse.parse_qs(http.calls[0][3].decode()).items()}
        self.assertEqual(sent["grant_type"], "authorization_code")
        self.assertEqual(sent["code"], "C")
        self.assertEqual(sent["redirect_uri"], "http://localhost:4321")
        # The verifier sent with the code is the one the challenge was made from.
        digest = hashlib.sha256(sent["code_verifier"].encode()).digest()
        self.assertEqual(asked["code_challenge"], base64.urlsafe_b64encode(digest).rstrip(b"=").decode())
        self.assertEqual(store.tokens["refresh_token"], "R")
        self.assertTrue(callback.closed)

    def test_reading_only_asks_to_read_only(self):
        http = FakeHttp((200, {"access_token": "A", "refresh_token": "R", "expires_in": 3600}))
        self.sign_in(FakeCallback({"code": "C"}), http, write=False)
        self.assertIn("Calendars.Read+offline_access", self.opened[0])

    def test_a_sign_in_microsoft_refused_says_why_and_asks_for_no_token(self):
        http = FakeHttp()
        callback = FakeCallback({"error": "access_denied",
                                 "error_description": "AADSTS65004: User declined. Trace ID: abc"})
        with self.assertRaises(graph.GraphSignInError) as caught:
            self.sign_in(callback, http)
        self.assertIn("User declined.", str(caught.exception))
        self.assertEqual(http.calls, [])
        self.assertTrue(callback.closed)

    def test_nobody_signing_in_gives_up(self):
        with self.assertRaises(graph.GraphSignInError) as caught:
            self.sign_in(FakeCallback(None), FakeHttp())
        self.assertIn("Nobody signed in", str(caught.exception))

    def test_a_reply_with_no_lasting_sign_in_is_not_kept(self):
        store = FakeStore()
        http = FakeHttp((200, {"access_token": "A", "expires_in": 3600}))
        with self.assertRaises(graph.GraphSignInError):
            self.sign_in(FakeCallback({"code": "C"}), http, store)
        self.assertIsNone(getattr(store, "tokens", None))


class TestCallback(unittest.TestCase):
    def setUp(self):
        self.callback = graph._Callback("the-state")
        self.addCleanup(self.callback.close)
        self.assertRegex(self.callback.redirect_uri, r"^http://localhost:\d+$")
        self.base = self.callback.redirect_uri.replace("localhost", "127.0.0.1")

    def visit(self, query):
        def fetch():
            try:
                with urllib.request.urlopen(self.base + "/?" + query, timeout=5) as response:
                    self.reply = response.status, response.read().decode()
            except urllib.error.HTTPError as error:
                self.reply = error.code, ""
                error.close()
        thread = threading.Thread(target=fetch)
        thread.start()
        self.callback._server.handle_request()
        thread.join()

    def test_it_listens_on_this_machine_only(self):
        self.assertEqual(self.callback._server.server_address[0], "127.0.0.1")

    def test_the_answer_with_this_sign_ins_state_is_taken(self):
        self.visit("code=C&state=the-state")
        self.assertEqual(self.reply[0], 200)
        self.assertIn("Signed in", self.reply[1])
        self.assertEqual(self.callback.wait(1)["code"], "C")

    def test_an_answer_with_another_state_is_refused(self):
        self.visit("code=EVIL&state=other")
        self.assertEqual(self.reply[0], 400)
        self.assertIsNone(self.callback.answer)

    def test_the_page_shows_nothing_that_was_sent(self):
        self.visit("error=%3Cscript%3E&state=the-state")
        self.assertIn("Not signed in", self.reply[1])
        self.assertNotIn("script", self.reply[1])


class TestReading(unittest.TestCase):
    def test_calendars_take_outlooks_names_colours_and_rights(self):
        http = FakeHttp((200, {"value": [
            {"id": "CAL1", "name": "Calendar", "hexColor": "", "color": "lightGreen",
             "isDefaultCalendar": True, "canEdit": True,
             "owner": {"address": "me@contoso.example"},
             "defaultOnlineMeetingProvider": "teamsForBusiness"},
            {"id": "CAL2", "name": "Holidays", "hexColor": "#AA00BB", "canEdit": False,
             "owner": {"address": "me@contoso.example"}, "defaultOnlineMeetingProvider": "unknown",
             "allowedOnlineMeetingProviders": []},
        ]}))
        holidays, work = client(http).calendars()
        self.assertEqual(work["id"], "graph:CAL1")
        # Every mailbox's main calendar is called "Calendar".
        self.assertEqual(work["name"], "me@contoso.example")
        self.assertEqual(work["color"], "#498205")
        self.assertTrue(work["primary"])
        self.assertTrue(work["writable"])
        self.assertEqual(work["lacks"], ["color"])
        self.assertEqual(work["video"], "teams")
        self.assertEqual(holidays["color"], "#AA00BB")
        self.assertFalse(holidays["writable"])
        self.assertEqual(holidays["lacks"], ["color", "meet"])
        self.assertNotIn("video", holidays)

    def test_events_ask_for_the_window_in_utc_and_follow_pages(self):
        page_two = "https://graph.microsoft.com/v1.0/me/calendars/CAL1/calendarView?$skip=250"
        http = FakeHttp(
            (200, {"id": "CAL1", "name": "Work", "owner": {"address": "me@contoso.example"}}),
            (200, {"value": [outlook_event()], "@odata.nextLink": page_two}),
            (200, {"value": [outlook_event(id="EV2")]}),
        )
        events = client(http).events("graph:CAL1", "2026-11-09T12:00:00+00:00", "2027-01-15T07:00:00-05:00")
        self.assertEqual([event["id"] for event in events], ["EV1", "EV2"])
        query = urllib.parse.parse_qs(urllib.parse.urlsplit(http.calls[1][1]).query)
        self.assertEqual(query["startDateTime"], ["2026-11-09T12:00:00"])
        self.assertEqual(query["endDateTime"], ["2027-01-15T12:00:00"])
        self.assertIn('outlook.timezone="UTC"', http.calls[1][2]["Prefer"])
        self.assertEqual(http.calls[2][1], page_two)

    def test_a_next_page_on_another_host_gets_no_token(self):
        http = FakeHttp(
            (200, {"id": "CAL1", "name": "Work"}),
            (200, {"value": [], "@odata.nextLink": "https://graph.microsoft.com.evil.example/v1.0/x"}),
        )
        with self.assertRaises(graph.GraphError):
            client(http).events("graph:CAL1", "2026-11-09T12:00:00+00:00", "2026-11-10T12:00:00+00:00")
        self.assertEqual(len(http.calls), 2)

    def test_an_event_with_unreadable_times_costs_only_itself(self):
        http = FakeHttp(
            (200, {"id": "CAL1", "name": "Work"}),
            (200, {"value": [outlook_event(start={"dateTime": "soon"}), outlook_event(id="EV2")]}),
        )
        events = client(http).events("graph:CAL1", "2026-11-09T12:00:00+00:00", "2026-11-10T12:00:00+00:00")
        self.assertEqual([event["id"] for event in events], ["EV2"])

    def test_a_missing_event_is_not_found(self):
        http = FakeHttp((404, {"error": {"code": "ErrorItemNotFound", "message": "gone"}}))
        with self.assertRaises(graph.GraphNotFound):
            client(http).get("graph:CAL1", "EV1")


class TestToGoogle(unittest.TestCase):
    def test_a_timed_event_becomes_a_row(self):
        resource = graph.to_google(outlook_event())
        row, = normalize.normalize_event(resource, CALENDAR, NEW_YORK)
        self.assertEqual(row["start"], "2026-11-16T10:00:00-05:00")
        self.assertEqual(row["title"], "Planning")
        self.assertEqual(row["location"], "Room 4")
        self.assertEqual(row["description"], "Agenda\nitems")
        self.assertEqual(row["eventUrl"], "https://outlook.office365.com/owa/?itemid=EV1")
        self.assertEqual(row["reminders"], [15])
        self.assertFalse(row["allDay"])

    def test_a_teams_meeting_gets_its_join_link(self):
        join = "https://teams.microsoft.com/l/meetup-join/19%3ameeting_abc%40thread.v2/0"
        resource = graph.to_google(outlook_event(isOnlineMeeting=True, onlineMeeting={"joinUrl": join}))
        row, = normalize.normalize_event(resource, CALENDAR, NEW_YORK)
        self.assertEqual(row["meetingUrl"], join)
        form = event_form.resource_to_form(resource, "graph:CAL1", NEW_YORK)
        self.assertTrue(form["meet"])
        self.assertEqual(form["meetUrl"], join)

    def test_an_all_day_event_keeps_its_days(self):
        resource = graph.to_google(outlook_event(
            isAllDay=True,
            start={"dateTime": "2026-11-16T00:00:00.0000000", "timeZone": "UTC"},
            end={"dateTime": "2026-11-18T00:00:00.0000000", "timeZone": "UTC"}))
        self.assertEqual(resource["start"], {"date": "2026-11-16"})
        self.assertEqual(resource["end"], {"date": "2026-11-18"})
        rows = normalize.normalize_event(resource, CALENDAR, NEW_YORK)
        self.assertEqual([row["dateKey"] for row in rows], ["2026-11-16", "2026-11-17"])

    def test_an_all_day_event_moved_by_a_zone_conversion_keeps_its_days(self):
        # Midnight in Sydney and midnight in New York, each shown in UTC.
        for start, end in (("2026-11-15T13:00:00", "2026-11-16T13:00:00"),
                           ("2026-11-16T05:00:00", "2026-11-17T05:00:00")):
            resource = graph.to_google(outlook_event(
                isAllDay=True, start={"dateTime": start, "timeZone": "UTC"},
                end={"dateTime": end, "timeZone": "UTC"}))
            self.assertEqual(resource["start"], {"date": "2026-11-16"}, start)
            self.assertEqual(resource["end"], {"date": "2026-11-17"}, start)

    def test_a_cancelled_event_gets_no_row(self):
        resource = graph.to_google(outlook_event(isCancelled=True))
        self.assertEqual(normalize.normalize_event(resource, CALENDAR, NEW_YORK), [])

    def test_guests_carry_their_answers_and_the_owner_is_marked(self):
        resource = graph.to_google(outlook_event(
            isOrganizer=False,
            organizer={"emailAddress": {"address": "boss@contoso.example"}},
            attendees=[
                {"type": "required", "status": {"response": "none"},
                 "emailAddress": {"address": "boss@contoso.example", "name": "The Boss"}},
                {"type": "optional", "status": {"response": "tentativelyAccepted"},
                 "emailAddress": {"address": "Me@Contoso.example", "name": "Me"}},
                {"type": "resource", "status": {"response": "accepted"},
                 "emailAddress": {"address": "room4@contoso.example", "name": "Room 4"}},
            ]), owner="me@contoso.example")
        boss, me, room = resource["attendees"]
        self.assertEqual(boss, {"email": "boss@contoso.example", "responseStatus": "needsAction",
                                "displayName": "The Boss", "organizer": True})
        self.assertTrue(me["self"])
        self.assertTrue(me["optional"])
        self.assertEqual(me["responseStatus"], "tentative")
        self.assertTrue(room["resource"])
        row, = normalize.normalize_event(resource, CALENDAR, NEW_YORK)
        self.assertEqual(row["responseStatus"], "tentative")

    def test_the_sync_reads_the_owners_answer_kept_outside_the_guest_list(self):
        item = outlook_event(isOrganizer=False, responseStatus={"response": "declined"})
        listed = graph.to_google(item, listing=True)
        self.assertEqual(listed["attendees"], [{"self": True, "responseStatus": "declined"}])
        # The form shows guests, and that row is not one.
        self.assertEqual(graph.to_google(item)["attendees"], [])

    def test_free_and_private_map_to_googles_words(self):
        resource = graph.to_google(outlook_event(showAs="free", sensitivity="private"))
        self.assertEqual(resource["transparency"], "transparent")
        self.assertEqual(resource["visibility"], "private")

    def test_a_reminder_turned_off_is_no_reminder(self):
        resource = graph.to_google(outlook_event(isReminderOn=False))
        self.assertEqual(resource["reminders"], {"useDefault": False, "overrides": []})

    def test_an_html_body_reaches_the_form_as_text(self):
        resource = graph.to_google(outlook_event(body={"contentType": "html", "content": "<p>One</p><p>Two</p>"}))
        self.assertNotIn("<", resource["description"])
        self.assertIn("One", resource["description"])

    def test_an_occurrence_names_its_series(self):
        resource = graph.to_google(outlook_event(type="occurrence", seriesMasterId="MASTER"))
        self.assertEqual(resource["recurringEventId"], "MASTER")
        self.assertNotIn("recurrence", resource)


def master(pattern, span=None, start="2026-11-12T15:00:00.0000000"):
    return outlook_event(type="seriesMaster", start={"dateTime": start, "timeZone": "UTC"},
                         end={"dateTime": start[:11] + "16:00:00", "timeZone": "UTC"},
                         recurrence={"pattern": pattern,
                                     "range": {"type": "noEnd", "startDate": "2026-11-12", **(span or {})}})


class TestRecurrence(unittest.TestCase):
    # 2026-11-12 is the second Thursday of November.
    FIRST = date(2026, 11, 12)

    def preset(self, pattern, span=None):
        resource = graph.to_google(master(pattern, span))
        return event_form.repeat_preset(resource["recurrence"], self.FIRST)

    def test_outlooks_patterns_read_as_the_forms_presets(self):
        self.assertEqual(self.preset({"type": "daily", "interval": 1}), "daily")
        self.assertEqual(self.preset({"type": "weekly", "interval": 1, "daysOfWeek": ["thursday"]}), "weekly")
        self.assertEqual(self.preset({"type": "weekly", "interval": 1, "daysOfWeek": [
            "friday", "monday", "wednesday", "tuesday", "thursday"]}), "weekdays")
        self.assertEqual(self.preset({"type": "relativeMonthly", "interval": 1, "index": "second",
                                      "daysOfWeek": ["thursday"]}), "monthly")
        self.assertEqual(self.preset({"type": "absoluteMonthly", "interval": 1, "dayOfMonth": 12}), "monthlyDate")
        self.assertEqual(self.preset({"type": "absoluteYearly", "interval": 1, "dayOfMonth": 12, "month": 11}), "yearly")

    def test_an_interval_or_an_end_reads_as_custom(self):
        self.assertEqual(self.preset({"type": "weekly", "interval": 2, "daysOfWeek": ["thursday"]}), "custom")
        self.assertEqual(self.preset({"type": "daily", "interval": 1},
                                     {"type": "endDate", "endDate": "2026-12-31"}), "custom")

    def test_the_series_day_is_its_own_zones_not_utcs(self):
        # 20:00 in New York on the 12th is already the 13th in UTC.
        item = master({"type": "absoluteMonthly", "interval": 1, "dayOfMonth": 12}, start="2026-11-13T01:00:00.0000000")
        self.assertEqual(graph.to_google(item)["recurrence"], ["RRULE:FREQ=MONTHLY"])

    def test_every_rule_this_writes_reads_back_the_same(self):
        patterns = [
            ({"type": "daily", "interval": 3}, None),
            ({"type": "weekly", "interval": 2, "daysOfWeek": ["monday", "thursday"]}, None),
            ({"type": "absoluteMonthly", "interval": 1, "dayOfMonth": 28}, None),
            ({"type": "relativeMonthly", "interval": 1, "index": "last", "daysOfWeek": ["friday"]}, None),
            ({"type": "relativeYearly", "interval": 1, "index": "first", "daysOfWeek": ["monday"], "month": 9}, None),
            ({"type": "absoluteYearly", "interval": 1, "dayOfMonth": 4, "month": 7}, None),
            ({"type": "daily", "interval": 1}, {"type": "numbered", "numberOfOccurrences": 10}),
            ({"type": "weekly", "interval": 1, "daysOfWeek": ["thursday"]},
             {"type": "endDate", "endDate": "2027-03-01"}),
        ]
        for pattern, span in patterns:
            lines = graph.to_google(master(pattern, span))["recurrence"]
            back = graph.to_recurrence(lines, self.FIRST, "America/New_York")
            for key, value in pattern.items():
                self.assertEqual(back["pattern"][key], value, lines)
            for key, value in (span or {"type": "noEnd"}).items():
                self.assertEqual(back["range"][key], value, lines)
            self.assertEqual(back["range"]["startDate"], "2026-11-12")
            self.assertEqual(back["range"]["recurrenceTimeZone"], "America/New_York")

    def test_the_forms_presets_become_outlooks_patterns(self):
        def pattern(preset):
            return graph.to_recurrence(event_form.repeat_rule(preset, self.FIRST), self.FIRST, "UTC")["pattern"]

        self.assertEqual(pattern("daily")["type"], "daily")
        self.assertEqual(pattern("weekly")["daysOfWeek"], ["thursday"])
        self.assertEqual(pattern("monthly")["index"], "second")
        self.assertEqual(pattern("monthlyDate"), {"interval": 1, "type": "absoluteMonthly", "dayOfMonth": 12})
        self.assertEqual(pattern("yearly"), {"interval": 1, "type": "absoluteYearly", "dayOfMonth": 12, "month": 11})
        self.assertEqual(len(pattern("weekdays")["daysOfWeek"]), 5)

    def test_a_rule_outlook_cannot_store_is_refused(self):
        for lines in (["RRULE:FREQ=HOURLY"], ["RRULE:FREQ=MONTHLY;BYMONTHDAY=1,15"],
                      ["RRULE:FREQ=MONTHLY;BYDAY=5MO"], ["RRULE:FREQ=DAILY;BYHOUR=9"],
                      ["RRULE:FREQ=DAILY", "EXDATE:20261113"], ["FREQ=DAILY"]):
            with self.assertRaises(graph.GraphError, msg=lines):
                graph.to_recurrence(lines, self.FIRST, "UTC")


class TestWriting(unittest.TestCase):
    FORM = {
        "title": "Review", "startDate": "2026-11-16", "startTime": "10:00",
        "endDate": "2026-11-16", "endTime": "10:30", "location": "Room 4", "description": "Notes",
        "guests": [{"email": "Ana@Example.com", "optional": True}], "meet": True,
        "repeat": "weekly", "reminders": {"useDefault": False, "overrides": [
            {"method": "popup", "minutes": 30}, {"method": "popup", "minutes": 10},
            {"method": "email", "minutes": 5}]},
        "busy": False, "visibility": "private", "colorId": "5",
    }

    def test_create_sends_outlooks_fields_to_the_calendar(self):
        http = FakeHttp((200, {"id": "CAL1", "name": "Work", "defaultOnlineMeetingProvider": "teamsForBusiness"}),
                        (201, outlook_event(id="NEW")))
        body = event_form.form_to_body(self.FORM, NEW_YORK)
        resource = client(http).create("graph:CAL1", body, send_updates="all")
        self.assertEqual(resource["id"], "NEW")
        method, url, _headers, _data = http.calls[1]
        self.assertEqual((method, url), ("POST", "https://graph.microsoft.com/v1.0/me/calendars/CAL1/events"))
        sent = http.body()
        self.assertEqual(sent["subject"], "Review")
        self.assertEqual(sent["start"], {"dateTime": "2026-11-16T10:00:00", "timeZone": "America/New_York"})
        self.assertEqual(sent["end"], {"dateTime": "2026-11-16T10:30:00", "timeZone": "America/New_York"})
        self.assertFalse(sent["isAllDay"])
        self.assertEqual(sent["body"], {"contentType": "text", "content": "Notes"})
        self.assertEqual(sent["location"], {"displayName": "Room 4"})
        self.assertEqual(sent["attendees"], [
            {"emailAddress": {"address": "ana@example.com", "name": ""}, "type": "optional"}])
        self.assertTrue(sent["isOnlineMeeting"])
        self.assertEqual(sent["onlineMeetingProvider"], "teamsForBusiness")
        self.assertEqual(sent["recurrence"]["pattern"]["daysOfWeek"], ["monday"])
        self.assertEqual(sent["recurrence"]["range"]["startDate"], "2026-11-16")
        # One reminder: the earliest popup. An email reminder is not one.
        self.assertTrue(sent["isReminderOn"])
        self.assertEqual(sent["reminderMinutesBeforeStart"], 10)
        self.assertEqual(sent["showAs"], "free")
        self.assertEqual(sent["sensitivity"], "private")
        # Google's own: a colour and guest permissions.
        for key in ("colorId", "guestsCanModify", "conferenceData", "summary"):
            self.assertNotIn(key, sent)

    def test_a_new_event_on_the_default_reminder_leaves_it_to_outlook(self):
        body = event_form.form_to_body({**self.FORM, "reminders": {"useDefault": True}}, NEW_YORK)
        sent = graph.from_google(body, NEW_YORK, creating=True)
        self.assertNotIn("isReminderOn", sent)

    def test_an_all_day_event_runs_midnight_to_midnight(self):
        body = event_form.form_to_body({"title": "Trip", "allDay": True, "startDate": "2026-11-16",
                                        "endDate": "2026-11-17"}, NEW_YORK)
        sent = graph.from_google(body, NEW_YORK, creating=True)
        self.assertTrue(sent["isAllDay"])
        self.assertEqual(sent["start"], {"dateTime": "2026-11-16T00:00:00", "timeZone": "America/New_York"})
        self.assertEqual(sent["end"], {"dateTime": "2026-11-18T00:00:00", "timeZone": "America/New_York"})

    def test_update_patches_only_what_it_is_given(self):
        http = FakeHttp((200, outlook_event(subject="Renamed")))
        resource = client(http).update("graph:CAL1", "EV/1", {"summary": "Renamed"})
        self.assertEqual(resource["summary"], "Renamed")
        method, url, _headers, _data = http.calls[0]
        self.assertEqual((method, url), ("PATCH", "https://graph.microsoft.com/v1.0/me/events/EV%2F1"))
        self.assertEqual(http.body(), {"subject": "Renamed"})

    def test_an_update_of_only_googles_fields_sends_no_patch(self):
        http = FakeHttp((200, outlook_event()))
        client(http).update("graph:CAL1", "EV1", {"colorId": "5", "guestsCanModify": True})
        self.assertEqual([call[0] for call in http.calls], ["GET"])

    def test_removing_a_repeat_sends_null(self):
        http = FakeHttp((200, outlook_event()))
        client(http).update("graph:CAL1", "EV1", {"recurrence": []})
        self.assertEqual(http.body(), {"recurrence": None})

    def test_a_new_repeat_without_a_start_reads_the_series_first_day(self):
        http = FakeHttp((200, outlook_event()), (200, outlook_event()))
        client(http).update("graph:CAL1", "EV1", {"recurrence": ["RRULE:FREQ=MONTHLY"]})
        self.assertEqual([call[0] for call in http.calls], ["GET", "PATCH"])
        sent = http.body()["recurrence"]
        self.assertEqual(sent["pattern"], {"interval": 1, "type": "absoluteMonthly", "dayOfMonth": 16})
        self.assertEqual(sent["range"]["startDate"], "2026-11-16")

    def test_replace_sends_only_what_differs(self):
        current = outlook_event()
        http = FakeHttp((200, current), (200, outlook_event(isAllDay=True)))
        whole = {**graph.to_google(current), "start": {"date": "2026-11-16"}, "end": {"date": "2026-11-17"}}
        client(http).replace("graph:CAL1", "EV1", whole)
        sent = http.body()
        self.assertEqual(sorted(sent), ["end", "isAllDay", "start"])
        self.assertTrue(sent["isAllDay"])

    def test_replace_refuses_to_drop_a_teams_meeting(self):
        current = outlook_event(isOnlineMeeting=True, onlineMeeting={"joinUrl": "https://teams.microsoft.com/l/meetup-join/x"})
        http = FakeHttp((200, current))
        whole = graph.to_google(current)
        whole.pop("hangoutLink")
        with self.assertRaises(graph.GraphError) as caught:
            client(http).replace("graph:CAL1", "EV1", whole)
        self.assertEqual(str(caught.exception), graph.KEEPS_TEAMS)
        self.assertEqual(len(http.calls), 1)

    def test_delete_expects_no_body(self):
        http = FakeHttp((204, ""))
        self.assertIsNone(client(http).delete("graph:CAL1", "EV1"))
        self.assertEqual(http.calls[0][0], "DELETE")

    def test_a_refused_change_says_why(self):
        http = FakeHttp((403, {"error": {"code": "ErrorAccessDenied", "message": "Access is denied."}}))
        with self.assertRaises(graph.GraphError) as caught:
            client(http).update("graph:CAL1", "EV1", {"summary": "x"})
        self.assertIn("Access is denied.", str(caught.exception))


if __name__ == "__main__":
    unittest.main()
