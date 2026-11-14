"""The one command the panel runs to list accounts and to connect one.

In: one JSON object on stdin, {"action": "list"} or {"action": "connect",
"type": "google" | "outlook"}. Out: one JSON line on stdout, and exit 0 or
1: {"ok": true, "accounts": [...]} for list, {"ok": true} for connect,
{"ok": false, "error": ...} on failure, written for a person.

Connecting is a sign-in in a browser, with questions the first time, so
this opens sync/setup in a terminal window with the right option for what
is already connected. Nothing from the request reaches a command line: the
type picks one of the fixed options below.
"""

import json
import shlex
import shutil
import subprocess
import sys
from pathlib import Path

from . import config as config_module
from .errors import SyncError
from .graph import Graph
from .gws import Gws

SETUP = Path(__file__).resolve().parent.parent / "setup"
TYPES = ("google", "outlook")
NO_TERMINAL = "No terminal to open. Run this in one yourself: %s"
CLOSE_PROMPT = "Press any key to close this window."


def _raw(path=None):
    """The config as written, or {}: a missing file is no accounts, not
    the defaults' Google one."""
    path = Path(path) if path is not None else config_module.CONFIG_PATH
    try:
        raw = json.loads(path.read_text())
    except (OSError, ValueError):
        return {}
    return raw if isinstance(raw, dict) else {}


def _backends(raw):
    named = raw.get("backend")
    if isinstance(named, list):
        return [str(name).strip().lower() for name in named]
    if named:
        return [str(named).strip().lower()]
    # The Google setup of older versions named no backend.
    return ["gws"] if raw.get("gwsPath") else []


def configured(raw):
    """Every connected account: [{"type", "number", "settings"}]."""
    found = []
    names = _backends(raw)
    if "gws" in names:
        first = str(raw.get("profile") or config_module.DEFAULTS["profile"])
        more = [one for one in raw.get("profiles") or [] if isinstance(one, str) and one.strip()]
        for number, profile in enumerate(dict.fromkeys([first, *more]), start=1):
            found.append({"type": "google", "number": number, "settings": profile})
    if "graph" in names:
        held = raw.get("graph")
        held = held if isinstance(held, list) else [held]
        for number, entry in enumerate([one for one in held if isinstance(one, dict)], start=1):
            found.append({"type": "outlook", "number": number, "settings": entry})
    return found


def _client(account, raw):
    if account["type"] == "google":
        return Gws(account["settings"], binary=str(raw.get("gwsPath") or "gws"))
    return Graph(account["settings"])


def _address(calendars):
    """The account's own address, as its main calendar gives it."""
    for calendar in calendars:
        if calendar.get("primary"):
            # Google's main calendar has the address as its id; Outlook's
            # is named after its owner.
            for value in (calendar.get("id"), calendar.get("name")):
                if isinstance(value, str) and "@" in value and not value.startswith("graph:"):
                    return value
    return ""


def list_accounts(raw, client_for=_client):
    """Each account with its address and whether its sign-in still works."""
    listed = []
    for account in configured(raw):
        row = {"type": account["type"], "number": account["number"], "address": "",
               "calendars": 0, "ok": True, "error": ""}
        try:
            calendars = client_for(account, raw).calendars()
            row["address"] = _address(calendars)
            row["calendars"] = len(calendars)
        except (SyncError, config_module.ConfigError) as error:
            row["ok"] = False
            row["error"] = str(error)[:300]
        listed.append(row)
    return listed


def setup_option(kind, raw):
    """The sync/setup option that connects one more account of a kind."""
    kinds = [account["type"] for account in configured(raw)]
    if kind == "google":
        profile = Path(str(raw.get("profile") or config_module.DEFAULTS["profile"]))
        # The first account's setup makes the OAuth client the others use.
        return "--add-google" if "google" in kinds and (profile / "client_secret.json").exists() else "--write"
    if kind == "outlook":
        return "--add-outlook" if "outlook" in kinds else "--outlook"
    raise ValueError(kind)


def open_terminal(option, setup=SETUP, which=shutil.which, spawn=subprocess.Popen):
    """Runs sync/setup with one option in a terminal window of its own,
    which stays open until a key is pressed so the outcome can be read."""
    command = "%s %s" % (shlex.quote(str(setup)), option)
    if not which("xdg-terminal-exec"):
        raise SyncError(NO_TERMINAL % command)
    script = "%s; echo; read -n 1 -s -r -p %s" % (command, shlex.quote(CLOSE_PROMPT))
    argv = ["xdg-terminal-exec", "--app-id=org.omarchy.terminal", "--title=Omarchy Calendar",
            "-e", "bash", "-c", script]
    if which("uwsm-app"):
        argv = ["uwsm-app", "--", *argv]
    try:
        spawn(argv, stdin=subprocess.DEVNULL, stdout=subprocess.DEVNULL,
              stderr=subprocess.DEVNULL, start_new_session=True)
    except OSError as error:
        raise SyncError(NO_TERMINAL % command) from error


def perform(request, raw, client_for=_client, terminal=open_terminal):
    """Run one request. Returns (exit code, reply)."""
    if not isinstance(request, dict):
        return 1, {"ok": False, "error": "The request must be a JSON object."}
    action = request.get("action")
    if action == "list":
        return 0, {"ok": True, "accounts": list_accounts(raw, client_for)}
    if action == "connect":
        kind = request.get("type")
        if kind not in TYPES:
            return 1, {"ok": False, "error": "Pick Google or Outlook."}
        try:
            terminal(setup_option(kind, raw))
        except SyncError as error:
            return 1, {"ok": False, "error": str(error)}
        return 0, {"ok": True}
    return 1, {"ok": False, "error": "Unknown action."}


def main():
    try:
        request = json.loads(sys.stdin.read())
    except ValueError:
        request = None
    code, reply = perform(request, _raw())
    print(json.dumps(reply))
    return code


if __name__ == "__main__":
    sys.exit(main())
