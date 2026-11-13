"""Several backends as one, so Google and Outlook share one events file.

Set "backend" to a list in calendar-sync.json: ["gws", "graph"]. Each
calendar belongs to the backend that listed it, and a change goes to that
backend only. A backend that fails (an expired sign-in, a server that is
down) costs its own calendars for that sync, not the others'.
"""

import sys

from .errors import SyncError

NOT_WRITABLE = "This calendar's backend cannot write events."


class Multi:
    def __init__(self, clients):
        self._clients = list(clients)
        if not self._clients:
            raise ValueError("no backends")
        self._active = list(self._clients)
        self._owners = {}
        self._failed = []
        self.SOURCE_NAME = "+".join(client.SOURCE_NAME for client in self._clients)
        self.can_write = any(getattr(client, "can_write", False) for client in self._clients)

    def version(self):
        return (len(self._clients),)

    def check(self):
        self._active, self._failed, first = [], [], None
        for client in self._clients:
            try:
                client.check()
            except SyncError as error:
                first = first or error
                self._skip(client, error)
                continue
            self._active.append(client)
        if not self._active:
            raise first

    def calendars(self):
        found, first = [], None
        for client in list(self._active):
            try:
                listed = client.calendars()
            except SyncError as error:
                first = first or error
                self._active.remove(client)
                self._skip(client, error)
                continue
            for calendar in listed:
                if calendar["id"] in self._owners and self._owners[calendar["id"]] is not client:
                    continue
                self._owners[calendar["id"]] = client
                if not getattr(client, "can_write", False):
                    calendar = {**calendar, "writable": False}
                found.append(calendar)
        if not self._active:
            raise first
        return found

    def events(self, calendar_id, time_min, time_max):
        return self._owner(calendar_id).events(calendar_id, time_min, time_max)

    def get(self, calendar_id, event_id):
        return self._writer(calendar_id).get(calendar_id, event_id)

    def create(self, calendar_id, body, send_updates="none"):
        return self._writer(calendar_id).create(calendar_id, body, send_updates=send_updates)

    def update(self, calendar_id, event_id, body, send_updates="none"):
        return self._writer(calendar_id).update(calendar_id, event_id, body, send_updates=send_updates)

    def replace(self, calendar_id, event_id, resource, send_updates="none"):
        return self._writer(calendar_id).replace(calendar_id, event_id, resource, send_updates=send_updates)

    def delete(self, calendar_id, event_id, send_updates="none"):
        return self._writer(calendar_id).delete(calendar_id, event_id, send_updates=send_updates)

    def auth_hint(self, cfg):
        hints = [client.auth_hint(cfg) for client in self._failed or self._clients]
        return "\n".join(hint for hint in hints if hint)

    def _skip(self, client, error):
        self._failed.append(client)
        print(f"skipping {client.SOURCE_NAME}: {error}", file=sys.stderr)

    def _writer(self, calendar_id):
        client = self._owner(calendar_id)
        if not getattr(client, "can_write", False):
            raise SyncError(NOT_WRITABLE)
        return client

    def _owner(self, calendar_id):
        """The backend a calendar belongs to. The event command runs without
        a sync first: a backend that can tell its ids apart says so, and
        only when several cannot are the calendars listed."""
        if calendar_id in self._owners:
            return self._owners[calendar_id]
        rest = []
        for client in self._clients:
            owns = getattr(client, "owns", None)
            if owns is None:
                rest.append(client)
            elif owns(calendar_id):
                return client
        if len(rest) == 1:
            return rest[0]
        if rest:
            self.calendars()
            if calendar_id in self._owners:
                return self._owners[calendar_id]
        raise SyncError("no backend lists that calendar")
