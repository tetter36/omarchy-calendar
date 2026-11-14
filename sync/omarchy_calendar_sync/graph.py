"""Outlook and Microsoft 365 calendars over Microsoft Graph, read and write.

Microsoft offers no CalDAV, and a published iCal feed is read only, so this
talks to Graph directly over HTTPS with the standard library. Like the gws
backend it needs a client of your own: an app registration in Microsoft
Entra with the delegated permission Calendars.ReadWrite (Calendars.Read is
enough without writing). sync/setup --outlook creates it with the Azure
CLI when that is installed, and signs in through the browser on this
machine (authorization code with PKCE, back to a localhost address), or
with a typed device code where there is no browser. No client secret is
involved either way.

The sign-in lives in the keyring (secret-tool), never in the config file.
The bearer token goes to graph.microsoft.com only: redirects are not
followed, and a next-page address on another host is refused.

Everything this module returns is a Google-shaped event resource, like the
other backends, so normalization, the event form and the contract are
shared. What Outlook cannot store (Google's event colours) is listed per
calendar under "lacks". Three differences a person can notice:

- Outlook decides by itself whether guests get an email, so the panel's
  "notify guests" answer is not sent.
- A Teams meeting cannot be taken off an event once it has one.
- An event holds one reminder, so the earliest popup is the one kept.
"""

import base64
import hashlib
import hmac
import json
import os
import re
import secrets
import subprocess
import sys
import time
import urllib.error
import urllib.parse
import urllib.request
from datetime import date, datetime, timedelta, timezone
from http.server import BaseHTTPRequestHandler, HTTPServer
from zoneinfo import ZoneInfo, ZoneInfoNotFoundError

from . import budget, event_form
from .bounded import MAX_ERROR_BYTES, ResponseTooLarge, read_at_most, read_body
from .config import ConfigError
from .errors import EventNotFound, SyncError
from .plaintext import _LOOKS_LIKE_HTML, description_text

GRAPH = "https://graph.microsoft.com/v1.0"
LOGIN = "https://login.microsoftonline.com"
ID_PREFIX = "graph:"
READ_SCOPE = "https://graph.microsoft.com/Calendars.Read"
WRITE_SCOPE = "https://graph.microsoft.com/Calendars.ReadWrite"
HTTP_TIMEOUT_SECONDS = 30
MAX_PAGES = 50
PAGE_SIZE = 250
# An access token lasts about an hour. One with less than this left is
# renewed, so a run never starts with a token that ends halfway.
RENEW_BEFORE_SECONDS = 300
BROWSER_SIGN_IN_SECONDS = 300
FALLBACK_COLOR = "#0078d4"
DEFAULT_REMINDER_MINUTES = 15
DEFAULT_PROVIDER = "teamsForBusiness"
KEYRING_SERVICE = "omarchy-calendar-graph"

NO_SIGN_IN = "No Outlook sign-in. Run sync/setup --outlook."
SIGN_IN_EXPIRED = "Your Outlook sign-in expired. Run sync/setup --outlook."
NO_KEYRING = "secret-tool is needed to keep the Outlook sign-in. Install libsecret."
KEEPS_TEAMS = "Outlook cannot take a Teams meeting off an event once it has one."
BAD_REPEAT = "Outlook cannot store this repeat rule."

# Marks a resource whose event is an online meeting, so replace() can tell
# a removed Teams meeting from an event that never had one.
ONLINE_KEY = "graphOnlineMeeting"

LACKS = ["color"]

_SIGNED_IN_PAGE = ("<!doctype html><meta charset=utf-8><title>Omarchy Calendar</title>"
                   "<body style='font-family:sans-serif;margin:4em auto;max-width:30em'>"
                   "<h2>Signed in</h2><p>Omarchy Calendar has your Outlook sign-in. "
                   "You can close this tab.</p>")
_NOT_SIGNED_IN_PAGE = _SIGNED_IN_PAGE.replace("Signed in", "Not signed in").replace(
    "Omarchy Calendar has your Outlook sign-in. You can close this tab.",
    "The sign-in did not finish. The terminal says why.")

_CLIENT_ID = re.compile(r"^[0-9a-fA-F-]{32,40}$")
_ACCOUNT = re.compile(r"^[A-Za-z0-9_-]{0,32}$")
_TENANT = re.compile(r"^[A-Za-z0-9][A-Za-z0-9.-]{0,252}$")
_HEX_COLOR = re.compile(r"^#[0-9a-fA-F]{6}$")
_GRAPH_TIME = re.compile(r"^(\d{4}-\d{2}-\d{2})T(\d{2}:\d{2}(?::\d{2})?)")

# Outlook's named calendar colours, for a calendar with no hexColor.
_COLORS = {
    "lightBlue": "#0078d4",
    "lightGreen": "#498205",
    "lightOrange": "#da3b01",
    "lightGray": "#69797e",
    "lightYellow": "#c19c00",
    "lightTeal": "#038387",
    "lightPink": "#e3008c",
    "lightBrown": "#8e562e",
    "lightRed": "#d13438",
}

_RESPONSES = {
    "accepted": "accepted",
    "declined": "declined",
    "tentativelyAccepted": "tentative",
    "organizer": "accepted",
    "notResponded": "needsAction",
    "none": "needsAction",
}

_DAY_CODES = {
    "monday": "MO", "tuesday": "TU", "wednesday": "WE", "thursday": "TH",
    "friday": "FR", "saturday": "SA", "sunday": "SU",
}
_DAY_NAMES = {code: name for name, code in _DAY_CODES.items()}
_INDEX_NUMBERS = {"first": 1, "second": 2, "third": 3, "fourth": 4, "last": -1}
_INDEX_NAMES = {number: name for name, number in _INDEX_NUMBERS.items()}
_RULE_KEYS = {"FREQ", "INTERVAL", "BYDAY", "BYMONTHDAY", "BYMONTH", "BYSETPOS", "COUNT", "UNTIL", "WKST"}
_BYDAY = re.compile(r"^([+-]?\d+)?(MO|TU|WE|TH|FR|SA|SU)$")


class GraphError(SyncError):
    """Base class for every failure this backend reports."""


class GraphSignInError(GraphError):
    """No usable sign-in. The message is for a person."""


class GraphNotFound(GraphError, EventNotFound):
    """The event or the calendar is not there (HTTP 404)."""


# ---- HTTP ----------------------------------------------------------------


class _NoRedirect(urllib.request.HTTPRedirectHandler):
    """A redirect would carry the bearer token to wherever it points."""

    def redirect_request(self, *_args, **_kwargs):
        return None


_OPENER = urllib.request.build_opener(_NoRedirect)


def _https(method, url, headers, data):
    """One request: (status, body text). An HTTP error status is a reply,
    not an exception; only a request that never got one raises."""
    request = urllib.request.Request(url, data=data, headers=headers, method=method)
    try:
        with _OPENER.open(request, timeout=HTTP_TIMEOUT_SECONDS) as response:
            return response.status, read_body(response).decode("utf-8", "replace")
    except urllib.error.HTTPError as error:
        return error.code, read_at_most(error, MAX_ERROR_BYTES).decode("utf-8", "replace")
    except ResponseTooLarge as error:
        raise GraphError(f"Outlook answered: {error}") from error
    except (urllib.error.URLError, OSError) as error:
        raise GraphError(f"cannot reach Outlook: {getattr(error, 'reason', error)}") from error


def _quote(value):
    return urllib.parse.quote(str(value), safe="")


def _json(text):
    try:
        payload = json.loads(text) if text.strip() else None
    except json.JSONDecodeError:
        return None
    return payload if isinstance(payload, dict) else None


# ---- The sign-in ---------------------------------------------------------


class Keyring:
    """The sign-in for one app registration, kept with secret-tool.

    It goes in on stdin and comes out on stdout, never as an argument:
    arguments show to every user on the machine through ps.
    """

    def __init__(self, client_id, account=""):
        # A lookup matches every entry that has these attributes, so a
        # further account is told apart in the value, not by one more.
        client = "%s/%s" % (client_id, account) if account else str(client_id)
        self._attributes = ["service", KEYRING_SERVICE, "client", client]

    def _run(self, args, text=None):
        try:
            return subprocess.run(
                ["secret-tool", *args, *self._attributes], input=text, capture_output=True,
                text=True, timeout=budget.HELPER_SECONDS, preexec_fn=budget.uncap_memory,
            )
        except FileNotFoundError as error:
            raise GraphSignInError(NO_KEYRING) from error
        except subprocess.TimeoutExpired as error:
            raise GraphError("the keyring did not answer; is it locked?") from error

    def load(self):
        completed = self._run(["lookup"])
        try:
            found = json.loads(completed.stdout)
        except json.JSONDecodeError:
            return None
        return found if isinstance(found, dict) else None

    def save(self, tokens):
        completed = self._run(["store", "--label=Omarchy Calendar Outlook sign-in"], json.dumps(tokens))
        if completed.returncode != 0:
            raise GraphError("the keyring refused the Outlook sign-in: "
                             + (completed.stderr.strip()[:200] or "no reason given"))


def _settings(raw):
    """(client id, tenant) from the config's "graph" object."""
    raw = raw if isinstance(raw, dict) else {}
    client_id = str(raw.get("clientId") or "").strip()
    tenant = str(raw.get("tenant") or "organizations").strip()
    if not _CLIENT_ID.match(client_id):
        raise ConfigError("graph.clientId must be the Application (client) ID of your "
                          "app registration; run sync/setup --outlook")
    if not _TENANT.match(tenant):
        raise ConfigError("graph.tenant must be a Directory (tenant) ID, a domain, "
                          "or \"organizations\"")
    return client_id, tenant


def _account(raw):
    """Which sign-in of several: "" for the first, a short label after."""
    account = str((raw if isinstance(raw, dict) else {}).get("account") or "").strip()
    if not _ACCOUNT.match(account):
        raise ConfigError("graph.account must be a short label of letters and digits")
    return account


def _token_url(tenant):
    return f"{LOGIN}/{_quote(tenant)}/oauth2/v2.0/token"


def _form(http, url, fields):
    status, text = http("POST", url, {"Content-Type": "application/x-www-form-urlencoded"},
                        urllib.parse.urlencode(fields).encode("utf-8"))
    return status, _json(text) or {}


def _kept(reply, previous, now):
    """What the keyring holds after a token reply. Microsoft sends a new
    refresh token with most replies; the one before stays valid."""
    try:
        lifetime = int(reply.get("expires_in") or 0)
    except (TypeError, ValueError):
        lifetime = 0
    return {
        "refresh_token": reply.get("refresh_token") or (previous or {}).get("refresh_token") or "",
        "access_token": reply["access_token"],
        "expires_at": int(now + lifetime),
    }


def _keep_first(status, reply, store, now):
    """Keeps the reply that ends a sign-in, or says why there is none."""
    if status != 200 or not isinstance(reply.get("access_token"), str) or not reply["access_token"]:
        raise GraphSignInError("The sign-in did not finish: " + _reason(reply, status))
    if not reply.get("refresh_token"):
        raise GraphSignInError(
            "Microsoft gave no lasting sign-in. Check that the app registration "
            "allows public client flows.")
    store.save(_kept(reply, None, now))


def _scope(write):
    return "%s offline_access" % (WRITE_SCOPE if write else READ_SCOPE)


class _Callback:
    """Waits on this machine for the browser to come back from Microsoft.

    Listens on the loopback address only, on a port the system picks, and
    takes no answer but the one that carries this sign-in's state.
    """

    def __init__(self, state):
        callback = self
        self.answer = None

        class Handler(BaseHTTPRequestHandler):
            def do_GET(self):
                query = urllib.parse.parse_qs(urllib.parse.urlsplit(self.path).query)
                answer = {key: values[0] for key, values in query.items()}
                if not hmac.compare_digest(answer.get("state", ""), state):
                    self.send_error(400)
                    return
                callback.answer = answer
                page = (_SIGNED_IN_PAGE if "code" in answer else _NOT_SIGNED_IN_PAGE).encode("utf-8")
                self.send_response(200)
                self.send_header("Content-Type", "text/html; charset=utf-8")
                self.send_header("Content-Length", str(len(page)))
                self.end_headers()
                self.wfile.write(page)

            def log_message(self, *_args):
                pass

        self._server = HTTPServer(("127.0.0.1", 0), Handler)
        self._server.timeout = 1
        # Microsoft matches a localhost address whatever its port.
        self.redirect_uri = "http://localhost:%d" % self._server.server_address[1]

    def wait(self, seconds, clock=time.time):
        give_up = clock() + seconds
        while self.answer is None and clock() < give_up:
            self._server.handle_request()
        return self.answer

    def close(self):
        self._server.server_close()


def _open_browser(address):
    try:
        subprocess.Popen(["xdg-open", address], stdin=subprocess.DEVNULL,
                         stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
    except OSError:
        pass


def login_browser(client_id, tenant, write, http=None, store=None, out=None,
                  open_browser=None, callback=None, clock=time.time, account=""):
    """Sign in through the browser on this machine: pick the account there
    and the sign-in lands in the keyring, with nothing to type.

    The authorization code flow with PKCE: the code Microsoft hands the
    browser is worth nothing without the verifier that never leaves here.
    """
    http = http or _https
    store = store or Keyring(client_id, account)
    out = out or sys.stdout
    state = secrets.token_urlsafe(24)
    verifier = secrets.token_urlsafe(64)
    challenge = base64.urlsafe_b64encode(
        hashlib.sha256(verifier.encode("ascii")).digest()).rstrip(b"=").decode("ascii")
    listener = (callback or _Callback)(state)
    try:
        address = f"{LOGIN}/{_quote(tenant)}/oauth2/v2.0/authorize?" + urllib.parse.urlencode({
            "client_id": client_id,
            "response_type": "code",
            "response_mode": "query",
            "redirect_uri": listener.redirect_uri,
            "scope": _scope(write),
            "state": state,
            "code_challenge": challenge,
            "code_challenge_method": "S256",
            "prompt": "select_account",
        })
        print("Your browser is opening for the Microsoft sign-in. If it does not, open:\n\n  %s\n"
              % address, file=out, flush=True)
        (open_browser or _open_browser)(address)
        answer = listener.wait(BROWSER_SIGN_IN_SECONDS)
    finally:
        listener.close()
    if answer is None:
        raise GraphSignInError("Nobody signed in within %d minutes. Run the setup again."
                               % (BROWSER_SIGN_IN_SECONDS // 60))
    if not answer.get("code"):
        raise GraphSignInError("The sign-in did not finish: " + _reason(answer, 0))
    status, reply = _form(http, _token_url(tenant), {
        "grant_type": "authorization_code",
        "client_id": client_id,
        "code": answer["code"],
        "redirect_uri": listener.redirect_uri,
        "code_verifier": verifier,
    })
    _keep_first(status, reply, store, clock())


def login(client_id, tenant, write, http=None, store=None, out=None, sleep=time.sleep,
          clock=time.time, account=""):
    """Sign in with the device code flow and keep the result in the keyring.

    Prints the address and the code to type there, then waits for the
    sign-in to finish in a browser, on this machine or any other. For a
    machine with no browser of its own; login_browser is the usual way.
    """
    http = http or _https
    store = store or Keyring(client_id, account)
    out = out or sys.stdout

    status, reply = _form(http, f"{LOGIN}/{_quote(tenant)}/oauth2/v2.0/devicecode",
                          {"client_id": client_id, "scope": _scope(write)})
    code = reply.get("device_code")
    if status != 200 or not isinstance(code, str) or not code:
        raise GraphSignInError("Microsoft refused to start the sign-in: " + _reason(reply, status))
    print(reply.get("message") or "Open %s and enter the code %s" % (
        reply.get("verification_uri"), reply.get("user_code")), file=out, flush=True)

    try:
        interval = max(1, int(reply.get("interval") or 5))
        give_up = clock() + int(reply.get("expires_in") or 900)
    except (TypeError, ValueError):
        interval, give_up = 5, clock() + 900
    while clock() < give_up:
        sleep(interval)
        status, reply = _form(http, _token_url(tenant), {
            "grant_type": "urn:ietf:params:oauth:grant-type:device_code",
            "client_id": client_id,
            "device_code": code,
        })
        problem = reply.get("error")
        if problem == "authorization_pending":
            continue
        if problem == "slow_down":
            interval += 5
            continue
        _keep_first(status, reply, store, clock())
        return
    raise GraphSignInError("The sign-in code expired before it was used. Run the setup again.")


def _reason(reply, status):
    text = str(reply.get("error_description") or reply.get("error") or "").strip()
    # The description runs on with trace ids, on the same line or the next.
    text = re.split(r"\s*(?:\r?\n|Trace ID:)", text, maxsplit=1)[0]
    return text[:300] or f"HTTP {status}"


# ---- Outlook's event, as Google's ----------------------------------------


def _zone(name):
    if not name or name in ("UTC", "Etc/UTC", "Etc/GMT", "GMT"):
        return timezone.utc
    try:
        return ZoneInfo(name)
    except (ZoneInfoNotFoundError, ValueError) as error:
        raise ValueError("unknown time zone %r" % name) from error


def _moment(node):
    """The aware datetime of Outlook's {"dateTime", "timeZone"}."""
    match = _GRAPH_TIME.match(str((node or {}).get("dateTime") or ""))
    if not match:
        raise ValueError("no dateTime")
    naive = datetime.fromisoformat("%sT%s" % match.groups())
    return naive.replace(tzinfo=_zone((node or {}).get("timeZone")))


def _whole_day(node):
    """The date of an all-day endpoint. It is midnight in the event's own
    zone; if a zone conversion moved it, the nearest midnight is the day."""
    moment = _moment(node)
    return moment.date() + timedelta(days=1 if moment.hour >= 12 else 0)


def _nodes(item):
    """Google's (start, end) nodes for Outlook's event."""
    if item.get("isAllDay"):
        first = _whole_day(item.get("start"))
        last = _whole_day(item.get("end") or item.get("start"))
        if last <= first:
            last = first + timedelta(days=1)
        return {"date": first.isoformat()}, {"date": last.isoformat()}
    start = _moment(item.get("start"))
    end = _moment(item.get("end") or item.get("start"))
    return ({"dateTime": start.astimezone(timezone.utc).isoformat()},
            {"dateTime": end.astimezone(timezone.utc).isoformat()})


def _address(holder):
    return str(((holder or {}).get("emailAddress") or {}).get("address") or "").strip()


def _attendees(item, owner):
    organizer = _address(item.get("organizer")).lower()
    people = []
    for entry in item.get("attendees") or []:
        email = _address(entry)
        if not email:
            continue
        person = {
            "email": email,
            "responseStatus": _RESPONSES.get(str((entry.get("status") or {}).get("response")), "needsAction"),
        }
        name = str((entry.get("emailAddress") or {}).get("name") or "").strip()
        if name and name.lower() != email.lower():
            person["displayName"] = name
        kind = entry.get("type")
        if kind == "optional":
            person["optional"] = True
        elif kind == "resource":
            person["resource"] = True
        if email.lower() == organizer:
            person["organizer"] = True
        if owner and email.lower() == owner.lower():
            person["self"] = True
        people.append(person)
    return people


def _reminders(item):
    minutes = item.get("reminderMinutesBeforeStart")
    if item.get("isReminderOn") and isinstance(minutes, int) and not isinstance(minutes, bool) and minutes >= 0:
        return {"useDefault": False, "overrides": [{"method": "popup", "minutes": minutes}]}
    return {"useDefault": False, "overrides": []}


def _description(item):
    body = item.get("body") or {}
    content = body.get("content")
    if content is None:
        return str(item.get("bodyPreview") or "")
    content = str(content)
    # The request asks for text. A server that sends HTML anyway must not
    # put markup in the form, where a save would store the tags as text.
    if body.get("contentType") == "html" or _LOOKS_LIKE_HTML.search(content):
        return description_text(content, limit=len(content) + 1)
    return content.replace("\r\n", "\n").strip()


def to_google(item, owner="", listing=False):
    """Outlook's event as a Google event resource.

    `owner` is the calendar owner's address, which marks their own row among
    the guests. `listing` is for the sync: it adds the owner's answer when
    Outlook keeps it outside the guest list, which the form must not show.
    Raises ValueError for an event with no usable times.
    """
    start, end = _nodes(item)
    resource = {
        "id": str(item.get("id") or ""),
        "summary": str(item.get("subject") or ""),
        "location": str((item.get("location") or {}).get("displayName") or ""),
        "description": _description(item),
        "start": start,
        "end": end,
        "htmlLink": str(item.get("webLink") or ""),
        "transparency": "transparent" if item.get("showAs") == "free" else "opaque",
        "visibility": {"private": "private", "personal": "private",
                       "confidential": "confidential"}.get(item.get("sensitivity"), "default"),
        "reminders": _reminders(item),
        "attendees": _attendees(item, owner),
    }
    if item.get("iCalUId"):
        resource["iCalUID"] = str(item["iCalUId"])
    if item.get("isCancelled"):
        resource["status"] = "cancelled"
    join = str((item.get("onlineMeeting") or {}).get("joinUrl") or item.get("onlineMeetingUrl") or "")
    if join:
        resource["hangoutLink"] = join
    if item.get("isOnlineMeeting"):
        resource[ONLINE_KEY] = True
    if item.get("seriesMasterId"):
        resource["recurringEventId"] = str(item["seriesMasterId"])
    if item.get("type") == "seriesMaster" and item.get("recurrence"):
        # The series' first day in its own zone. The start above is in UTC,
        # where an evening event is already on the next day.
        begins = str((item["recurrence"].get("range") or {}).get("startDate") or "")
        try:
            first = date.fromisoformat(begins)
        except ValueError:
            first = date.fromisoformat(start.get("date") or start["dateTime"][:10])
        resource["recurrence"] = to_rrule(item["recurrence"], first)

    if listing and not any(person.get("self") for person in resource["attendees"]):
        answer = str((item.get("responseStatus") or {}).get("response") or "")
        if not item.get("isOrganizer") and answer in ("accepted", "declined", "tentativelyAccepted", "notResponded"):
            resource["attendees"] = resource["attendees"] + [{"self": True, "responseStatus": _RESPONSES[answer]}]
    return resource


def to_rrule(recurrence, first):
    """Outlook's recurrence as RRULE lines, spelled the way the form's
    presets are, so "every week" reads as weekly and not as custom.
    `first` is the date the series starts on. [] for one it cannot spell."""
    pattern = recurrence.get("pattern") or {}
    span = recurrence.get("range") or {}
    kind = pattern.get("type")
    days = [_DAY_CODES[day] for day in pattern.get("daysOfWeek") or [] if day in _DAY_CODES]
    days.sort(key=event_form.WEEKDAYS.index)
    index = _INDEX_NUMBERS.get(pattern.get("index") or "first", 1)
    day_of_month = pattern.get("dayOfMonth") or first.day
    month = pattern.get("month") or first.month

    if kind == "daily":
        parts = ["FREQ=DAILY"]
    elif kind == "weekly":
        parts = ["FREQ=WEEKLY", "BYDAY=" + ",".join(days or [event_form.WEEKDAYS[first.weekday()]])]
    elif kind == "absoluteMonthly":
        parts = ["FREQ=MONTHLY"] + ([] if day_of_month == first.day else ["BYMONTHDAY=%d" % day_of_month])
    elif kind in ("relativeMonthly", "relativeYearly"):
        parts = ["FREQ=MONTHLY" if kind == "relativeMonthly" else "FREQ=YEARLY"]
        if kind == "relativeYearly":
            parts.append("BYMONTH=%d" % month)
        if len(days) == 1:
            parts.append("BYDAY=%d%s" % (index, days[0]))
        elif days:
            parts += ["BYDAY=" + ",".join(days), "BYSETPOS=%d" % index]
        else:
            return []
    elif kind == "absoluteYearly":
        parts = ["FREQ=YEARLY"]
        if (month, day_of_month) != (first.month, first.day):
            parts += ["BYMONTH=%d" % month, "BYMONTHDAY=%d" % day_of_month]
    else:
        return []

    try:
        interval = int(pattern.get("interval") or 1)
    except (TypeError, ValueError):
        interval = 1
    if interval > 1:
        parts.insert(1, "INTERVAL=%d" % interval)
    if span.get("type") == "endDate" and span.get("endDate"):
        parts.append("UNTIL=" + str(span["endDate"]).replace("-", "")[:8])
    elif span.get("type") == "numbered" and span.get("numberOfOccurrences"):
        parts.append("COUNT=%d" % int(span["numberOfOccurrences"]))
    return ["RRULE:" + ";".join(parts)]


# ---- Google's event body, as Outlook's -----------------------------------


def to_recurrence(lines, first, zone_name):
    """RRULE lines as Outlook's recurrence, for a series starting on `first`."""
    lines = [str(line).strip() for line in lines or [] if str(line).strip()]
    if len(lines) != 1 or not lines[0].upper().startswith("RRULE:"):
        # EXDATE and RDATE lines have no place in Outlook's recurrence.
        raise GraphError(BAD_REPEAT)
    rule = {}
    for part in lines[0][len("RRULE:"):].split(";"):
        key, _, value = part.partition("=")
        if key.upper() not in _RULE_KEYS or not value:
            raise GraphError(BAD_REPEAT)
        rule[key.upper()] = value.upper()

    try:
        interval = int(rule.get("INTERVAL", "1"))
        month_days = [int(day) for day in rule["BYMONTHDAY"].split(",")] if "BYMONTHDAY" in rule else []
        months = [int(month) for month in rule["BYMONTH"].split(",")] if "BYMONTH" in rule else []
        position = int(rule["BYSETPOS"]) if "BYSETPOS" in rule else None
        count = int(rule["COUNT"]) if "COUNT" in rule else None
        until = datetime.strptime(rule["UNTIL"][:8], "%Y%m%d").date() if "UNTIL" in rule else None
    except ValueError:
        raise GraphError(BAD_REPEAT) from None
    if interval < 1 or len(month_days) > 1 or len(months) > 1:
        raise GraphError(BAD_REPEAT)

    days, ordinals = [], set()
    for entry in rule["BYDAY"].split(",") if "BYDAY" in rule else []:
        match = _BYDAY.match(entry)
        if not match:
            raise GraphError(BAD_REPEAT)
        days.append(_DAY_NAMES[match.group(2)])
        if match.group(1):
            ordinals.add(int(match.group(1)))
    if position is not None:
        ordinals.add(position)
    if len(ordinals) > 1 or (ordinals and next(iter(ordinals)) not in _INDEX_NAMES):
        raise GraphError(BAD_REPEAT)
    index = _INDEX_NAMES[next(iter(ordinals))] if ordinals else None

    freq = rule.get("FREQ")
    pattern = {"interval": interval}
    if freq == "DAILY" and not days and not index:
        pattern["type"] = "daily"
    elif freq == "WEEKLY" and not index:
        pattern["type"] = "weekly"
        pattern["daysOfWeek"] = days or [_DAY_NAMES[event_form.WEEKDAYS[first.weekday()]]]
        pattern["firstDayOfWeek"] = _DAY_NAMES.get(rule.get("WKST", "SU"), "sunday")
    elif freq in ("MONTHLY", "YEARLY") and index and days:
        pattern["type"] = "relativeMonthly" if freq == "MONTHLY" else "relativeYearly"
        pattern["daysOfWeek"] = days
        pattern["index"] = index
        if freq == "YEARLY":
            pattern["month"] = months[0] if months else first.month
    elif freq == "MONTHLY" and not days:
        pattern["type"] = "absoluteMonthly"
        pattern["dayOfMonth"] = month_days[0] if month_days else first.day
    elif freq == "YEARLY" and not days:
        pattern["type"] = "absoluteYearly"
        pattern["dayOfMonth"] = month_days[0] if month_days else first.day
        pattern["month"] = months[0] if months else first.month
    else:
        raise GraphError(BAD_REPEAT)

    span = {"type": "noEnd", "startDate": first.isoformat(), "recurrenceTimeZone": zone_name}
    if count is not None:
        span.update(type="numbered", numberOfOccurrences=count)
    elif until is not None:
        span.update(type="endDate", endDate=until.isoformat())
    return {"pattern": pattern, "range": span}


def _zone_name(tz):
    return getattr(tz, "key", None) or "UTC"


def _graph_time(node, tz):
    """Outlook's {"dateTime", "timeZone"} for Google's start or end node."""
    if "date" in node:
        # An all-day event starts and ends at midnight in one zone.
        return {"dateTime": node["date"] + "T00:00:00", "timeZone": _zone_name(tz)}
    moment = datetime.fromisoformat(node["dateTime"])
    if moment.tzinfo is None:
        moment = moment.replace(tzinfo=tz)
    name = node.get("timeZone")
    try:
        # The wall time in a named zone, so a series keeps its hour across
        # daylight saving. Without a name, the instant in UTC.
        zone = ZoneInfo(name) if name else timezone.utc
    except (ZoneInfoNotFoundError, ValueError):
        name, zone = None, timezone.utc
    return {"dateTime": moment.astimezone(zone).strftime("%Y-%m-%dT%H:%M:%S"), "timeZone": name or "UTC"}


def from_google(body, tz, creating=False, provider=DEFAULT_PROVIDER, first=None):
    """Outlook's event fields for the keys a Google event body carries.

    Only the keys present are translated, so a patch stays a patch. `first`
    is the date a repeat rule starts on, when the body has no start.
    """
    out = {}
    if "summary" in body:
        out["subject"] = str(body.get("summary") or "")
    if "location" in body:
        out["location"] = {"displayName": str(body.get("location") or "")}
    if "description" in body:
        out["body"] = {"contentType": "text", "content": str(body.get("description") or "")}
    for key in ("start", "end"):
        if isinstance(body.get(key), dict) and body[key]:
            out[key] = _graph_time(body[key], tz)
    if "start" in out:
        out["isAllDay"] = "date" in body["start"]

    if "attendees" in body:
        out["attendees"] = [
            {
                "emailAddress": {"address": str(person["email"]).strip(),
                                 "name": str(person.get("displayName") or "")},
                "type": "resource" if person.get("resource")
                        else "optional" if person.get("optional") else "required",
            }
            for person in body.get("attendees") or [] if str(person.get("email") or "").strip()
        ]

    if (body.get("conferenceData") or {}).get("createRequest"):
        out["isOnlineMeeting"] = True
        out["onlineMeetingProvider"] = provider or DEFAULT_PROVIDER

    if "recurrence" in body:
        if body.get("recurrence"):
            start = body.get("start") or {}
            if start.get("date"):
                first = date.fromisoformat(start["date"])
            elif start.get("dateTime"):
                first = date.fromisoformat(out["start"]["dateTime"][:10])
            if first is None:
                raise GraphError(BAD_REPEAT)
            out["recurrence"] = to_recurrence(body["recurrence"], first, _zone_name(tz))
        else:
            out["recurrence"] = None

    reminders = body.get("reminders")
    if isinstance(reminders, dict):
        if reminders.get("useDefault", True):
            # A new event gets Outlook's own default by leaving it out.
            if not creating:
                out["isReminderOn"] = True
                out["reminderMinutesBeforeStart"] = DEFAULT_REMINDER_MINUTES
        else:
            popups = [int(entry.get("minutes") or 0) for entry in reminders.get("overrides") or []
                      if entry.get("method", "popup") == "popup"]
            out["isReminderOn"] = bool(popups)
            if popups:
                out["reminderMinutesBeforeStart"] = min(popups)

    if "transparency" in body:
        out["showAs"] = "free" if body["transparency"] == "transparent" else "busy"
    if "visibility" in body:
        out["sensitivity"] = {"private": "private", "confidential": "confidential"}.get(
            body["visibility"], "normal")
    return out


# ---- The backend ---------------------------------------------------------


def _calendar(item, alone_default_name="Calendar"):
    owner = _address({"emailAddress": item.get("owner") or {}})
    name = str(item.get("name") or "").strip()
    if item.get("isDefaultCalendar") and owner and name in ("", alone_default_name):
        # Every mailbox calls its main calendar "Calendar", which says
        # nothing next to another account's. Google names its by address.
        name = owner
    color = str(item.get("hexColor") or "")
    if not _HEX_COLOR.match(color):
        color = _COLORS.get(item.get("color"), FALLBACK_COLOR)
    providers = item.get("allowedOnlineMeetingProviders") or []
    provider = item.get("defaultOnlineMeetingProvider")
    if provider in (None, "", "unknown"):
        provider = providers[0] if providers else ""
    calendar = {
        "id": ID_PREFIX + str(item["id"]),
        "name": name or str(item["id"]),
        "color": color,
        "primary": item.get("isDefaultCalendar") is True,
        "writable": item.get("canEdit") is True,
        "defaultReminders": [],
        "lacks": list(LACKS) + ([] if provider else ["meet"]),
        "owner": owner,
        "provider": provider,
    }
    if provider:
        # What the form's video switch adds on this calendar.
        calendar["video"] = "teams" if "teams" in provider.lower() else "skype"
    return calendar


class Graph:
    SOURCE_NAME = "graph"
    can_write = True

    def __init__(self, settings, local_tz=None, http=None, store=None, clock=time.time):
        self.client_id, self.tenant = _settings(settings)
        self.account = _account(settings)
        self._local_tz = local_tz or timezone.utc
        self._http = http or _https
        self._store = store or Keyring(self.client_id, self.account)
        self._clock = clock
        self._access_token = None
        self._calendars = {}

    @staticmethod
    def owns(calendar_id):
        """Whether a calendar id is one this backend made, so several
        backends can share one events file without asking the network."""
        return str(calendar_id).startswith(ID_PREFIX)

    def version(self):
        return (1, 0)

    def check(self):
        self._token()

    def auth_hint(self, _cfg):
        return self._named("if this is a sign-in error, run sync/setup --outlook.")[:-1]

    def _named(self, message):
        """A sign-in message, naming the account when there are several."""
        if not self.account:
            return message
        return (message.replace("sign-in", "sign-in (account %s)" % self.account, 1)
                .replace("--outlook.", "--outlook=%s." % self.account))

    # -- reading

    def calendars(self):
        found = []
        url = f"{GRAPH}/me/calendars?{urllib.parse.urlencode({'$top': 100})}"
        for item in self._pages(url):
            if not item.get("id"):
                continue
            calendar = _calendar(item)
            self._calendars[calendar["id"]] = calendar
            found.append(calendar)
        return sorted(found, key=lambda calendar: calendar["name"])

    def events(self, calendar_id, time_min, time_max):
        query = urllib.parse.urlencode({
            # No offset means UTC to Graph, whatever the Prefer header says.
            "startDateTime": _utc(time_min),
            "endDateTime": _utc(time_max),
            "$top": PAGE_SIZE,
        })
        owner = self._calendar(calendar_id).get("owner", "")
        resources = []
        for item in self._pages(f"{GRAPH}/me/calendars/{_raw(calendar_id)}/calendarView?{query}"):
            try:
                resources.append(to_google(item, owner, listing=True))
            except (ValueError, KeyError, TypeError):
                # One event with times this cannot read costs that event.
                continue
            if len(resources) > budget.MAX_EVENTS:
                raise budget.TooManyEvents("more than %d events in the window" % budget.MAX_EVENTS)
        return resources

    def get(self, _calendar_id, event_id):
        return self._resource(self._call("GET", f"/me/events/{_quote(event_id)}"))

    # -- writing. Outlook mails the guests by its own rules: send_updates
    # has nowhere to go.

    def create(self, calendar_id, body, send_updates="none"):
        provider = DEFAULT_PROVIDER
        if (body.get("conferenceData") or {}).get("createRequest"):
            provider = self._calendar(calendar_id).get("provider") or DEFAULT_PROVIDER
        fields = from_google(body, self._local_tz, creating=True, provider=provider)
        return self._resource(self._call("POST", f"/me/calendars/{_raw(calendar_id)}/events", fields))

    def update(self, calendar_id, event_id, body, send_updates="none"):
        first = None
        if body.get("recurrence") and not body.get("start"):
            first = event_form.starts_on(self.get(calendar_id, event_id), self._local_tz)
        provider = DEFAULT_PROVIDER
        if (body.get("conferenceData") or {}).get("createRequest"):
            provider = self._calendar(calendar_id).get("provider") or DEFAULT_PROVIDER
        fields = from_google(body, self._local_tz, provider=provider, first=first)
        if not fields:
            # Nothing Outlook stores changed (a Google colour, say).
            return self.get(calendar_id, event_id)
        return self._resource(self._call("PATCH", f"/me/events/{_quote(event_id)}", fields))

    def replace(self, calendar_id, event_id, resource, send_updates="none"):
        # Graph has no whole-event replace, and needs none: a patch can turn
        # a timed event into an all-day one. Only what differs is sent, so a
        # formatted description is not flattened by an edit of the times.
        current = self.get(calendar_id, event_id)
        if (current.get("hangoutLink") or current.get(ONLINE_KEY)) and not resource.get("hangoutLink"):
            raise GraphError(KEEPS_TEAMS)
        return self.update(calendar_id, event_id,
                           event_form.changed_only(resource, current, self._local_tz), send_updates)

    def delete(self, _calendar_id, event_id, send_updates="none"):
        self._call("DELETE", f"/me/events/{_quote(event_id)}")
        return None

    # -- plumbing

    def _resource(self, item):
        if not isinstance(item, dict):
            raise GraphError("Outlook returned unparseable output")
        try:
            return to_google(item)
        except (ValueError, KeyError, TypeError) as error:
            raise GraphError(f"Outlook returned an event this cannot read: {error}") from error

    def _calendar(self, calendar_id):
        """One calendar's details. The event command runs without a sync
        first, so it asks for the one calendar it needs."""
        if calendar_id not in self._calendars:
            item = self._call("GET", f"/me/calendars/{_raw(calendar_id)}")
            if not isinstance(item, dict) or not item.get("id"):
                raise GraphNotFound("that calendar is not in Outlook")
            self._calendars[calendar_id] = _calendar({**item, "id": calendar_id[len(ID_PREFIX):]})
        return self._calendars[calendar_id]

    def _pages(self, url):
        for _ in range(MAX_PAGES):
            payload = self._call("GET", url) or {}
            yield from (item for item in payload.get("value") or [] if isinstance(item, dict))
            url = payload.get("@odata.nextLink")
            if not url:
                return
        raise GraphError(f"Outlook did not finish listing within {MAX_PAGES} pages")

    def _call(self, method, target, body=None, retry=True):
        url = target if "://" in target else GRAPH + target
        if not url.startswith(GRAPH + "/"):
            # A next-page address is the server's to name. The token is not
            # going anywhere else.
            raise GraphError("Outlook pointed at another server, which gets no sign-in")
        headers = {
            "Authorization": "Bearer " + self._token(),
            "Accept": "application/json",
            "Prefer": 'outlook.timezone="UTC", outlook.body-content-type="text"',
        }
        data = None
        if body is not None:
            headers["Content-Type"] = "application/json"
            data = json.dumps(body).encode("utf-8")
        status, text = self._http(method, url, headers, data)
        if status == 401 and retry:
            # The kept access token was revoked before it ran out.
            self._access_token = None
            self._renew(self._store.load() or {})
            return self._call(method, target, body, retry=False)
        payload = _json(text)
        if 200 <= status < 300:
            if payload is None and text.strip():
                raise GraphError("Outlook returned unparseable output")
            return payload
        error = (payload or {}).get("error")
        error = error if isinstance(error, dict) else {}
        message = str(error.get("message") or "").strip() or f"HTTP {status}"
        if status == 401:
            raise GraphSignInError(self._named(SIGN_IN_EXPIRED))
        if status == 404 or error.get("code") == "ErrorItemNotFound":
            raise GraphNotFound(f"404: {message}")
        if status == 403:
            raise GraphError(f"Outlook refused that: {message}")
        if status == 429:
            raise GraphError("Outlook is asking for fewer requests. The next sync tries again.")
        raise GraphError(f"Outlook answered HTTP {status}: {message}")

    def _token(self):
        """An access token, from the keyring while it lasts, renewed there
        when it does not. No message here may quote it."""
        if self._access_token:
            return self._access_token
        kept = self._store.load()
        if not kept or not kept.get("refresh_token"):
            raise GraphSignInError(self._named(NO_SIGN_IN))
        try:
            lasts = float(kept.get("expires_at") or 0) - self._clock()
        except (TypeError, ValueError):
            lasts = 0
        if kept.get("access_token") and lasts > RENEW_BEFORE_SECONDS:
            self._access_token = str(kept["access_token"])
            return self._access_token
        return self._renew(kept)

    def _renew(self, kept):
        if not kept.get("refresh_token"):
            raise GraphSignInError(self._named(NO_SIGN_IN))
        status, reply = _form(self._http, _token_url(self.tenant), {
            "grant_type": "refresh_token",
            "client_id": self.client_id,
            "refresh_token": kept["refresh_token"],
        })
        token = reply.get("access_token")
        if status != 200 or not isinstance(token, str) or not token:
            # invalid_grant: the sign-in expired, was revoked, or the
            # organisation's policy asks for a new one.
            if status in (400, 401) or reply.get("error") == "invalid_grant":
                raise GraphSignInError(self._named(SIGN_IN_EXPIRED))
            raise GraphError(f"Microsoft answered the sign-in renewal with HTTP {status}")
        self._store.save(_kept(reply, kept, self._clock()))
        self._access_token = token
        return token


def _raw(calendar_id):
    text = str(calendar_id)
    return _quote(text[len(ID_PREFIX):] if text.startswith(ID_PREFIX) else text)


def _utc(text):
    moment = datetime.fromisoformat(text)
    if moment.tzinfo is not None:
        moment = moment.astimezone(timezone.utc)
    return moment.strftime("%Y-%m-%dT%H:%M:%S")


def main(argv=None):
    """`login`: the sign-in sync/setup --outlook runs."""
    import argparse

    parser = argparse.ArgumentParser(prog="omarchy_calendar_sync.graph")
    parser.add_argument("command", choices=["login"])
    parser.add_argument("--client-id", required=True)
    parser.add_argument("--tenant", default="organizations")
    parser.add_argument("--write", action="store_true")
    parser.add_argument("--account", default="", help="a further account's label")
    parser.add_argument("--device-code", action="store_true",
                        help="type a code in a browser on any machine, for one with no browser")
    args = parser.parse_args(argv)
    # Without a desktop there is no browser to open here.
    desktop = os.environ.get("WAYLAND_DISPLAY") or os.environ.get("DISPLAY")
    try:
        client_id, tenant = _settings({"clientId": args.client_id, "tenant": args.tenant})
        account = _account({"account": args.account})
        if args.device_code or not desktop:
            login(client_id, tenant, args.write, account=account)
        else:
            login_browser(client_id, tenant, args.write, account=account)
    except (ConfigError, SyncError) as error:
        print(str(error), file=sys.stderr)
        return 1
    print("Signed in. The sign-in is in your keyring.")
    return 0


if __name__ == "__main__":
    sys.exit(main())
