#!/usr/bin/env python3
"""Google Calendar REST client — stdlib only (build steps 3.2–3.4).

Deliberately NOT google-api-python-client: the tool needs six endpoints,
and for a 10-year horizon (SPEC §8) a hand-rolled 200-line stdlib client
rots slower than a pinned SDK stack. Everything network-capable in this
project lives in this module, and nothing imports it except the
test/real-mode code paths — `local` mode can never touch it (AC-M1/D2).

OAuth: the standard installed-app flow. The one-time browser consent
lives in `interactive_authorize` (used only by gsetup.py); normal runs
just exchange the stored refresh token for an access token.

Test hook: when the ``CRISISPLAN_GCAL_FAKE`` environment variable names a
state file, `get_transport` returns an on-disk fake implementing the
same endpoints (plus a request log). The regression suites use it to
exercise the REAL setup/reconcile code offline; it is not a mock of our
code, only of Google's server.
"""

import json
import os
import urllib.error
import urllib.parse
import urllib.request
from pathlib import Path

API = "https://www.googleapis.com/calendar/v3"
TOKEN_URL = "https://oauth2.googleapis.com/token"
AUTH_URL = "https://accounts.google.com/o/oauth2/v2/auth"
SCOPE = "https://www.googleapis.com/auth/calendar"
FAKE_ENV = "CRISISPLAN_GCAL_FAKE"


class GcalError(RuntimeError):
    """Any Google-side failure. Callers treat it as the fragile tier:
    report and continue, never block local artifacts (AC-D3)."""


# --- transports --------------------------------------------------------------

class Transport:
    """Real HTTPS transport with a Bearer token."""

    def __init__(self, access_token: str):
        self.access_token = access_token

    def request(self, method, path, params=None, body=None):
        url = API + path
        if params:
            url += "?" + urllib.parse.urlencode(params, doseq=True)
        req = urllib.request.Request(
            url, method=method,
            data=json.dumps(body).encode("utf-8") if body else None,
            headers={"Authorization": f"Bearer {self.access_token}",
                     "Content-Type": "application/json"})
        try:
            with urllib.request.urlopen(req, timeout=30) as resp:
                raw = resp.read()
                return json.loads(raw) if raw else {}
        except urllib.error.HTTPError as e:
            detail = e.read()[:300].decode("utf-8", "replace")
            raise GcalError(f"{method} {path} -> HTTP {e.code}: {detail}")
        except urllib.error.URLError as e:
            raise GcalError(f"Google unreachable: {e.reason}")


class FakeTransport:
    """Offline stand-in for Google's server (test hook, see module doc).

    State (calendars/acls/events + a request log) persists in a JSON
    file so multi-process CLI tests can assert against it."""

    def __init__(self, state_path):
        self.path = Path(state_path)
        if self.path.exists():
            self.state = json.loads(self.path.read_text(encoding="utf-8"))
        else:
            self.state = {"calendars": {}, "acls": {}, "events": {},
                          "counter": 0, "log": [], "fail": False}

    def _save(self):
        self.path.parent.mkdir(parents=True, exist_ok=True)
        self.path.write_text(json.dumps(self.state, indent=1,
                                        ensure_ascii=False),
                             encoding="utf-8")

    def _next(self, prefix):
        self.state["counter"] += 1
        return f"{prefix}{self.state['counter']}"

    @staticmethod
    def _merge_patch(target, patch):
        for key, value in patch.items():
            if isinstance(value, dict) and isinstance(target.get(key),
                                                      dict):
                FakeTransport._merge_patch(target[key], value)
            elif value is None:
                target.pop(key, None)   # Google: null deletes the field
            else:
                target[key] = value

    def request(self, method, path, params=None, body=None):
        if self.state.get("fail"):
            raise GcalError("simulated Google outage (fake transport)")
        params = params or {}
        self.state["log"].append({"method": method, "path": path,
                                  "params": params})
        try:
            return self._route(method, path, params, body or {})
        finally:
            self._save()

    def _route(self, method, path, params, body):
        parts = path.strip("/").split("/")
        if method == "POST" and path == "/users/me/calendarList":
            cal = self.state["calendars"].get(body.get("id"))
            if cal is None:
                raise GcalError(f"POST {path} -> HTTP 404: calendar "
                                "not found")
            listed = self.state.setdefault("calendarList", [])
            if cal["id"] not in listed:
                listed.append(cal["id"])
            return cal
        if method == "POST" and path == "/calendars":
            cal_id = self._next("fake-cal-")
            cal = {"id": cal_id, "summary": body.get("summary", "")}
            self.state["calendars"][cal_id] = cal
            self.state["events"].setdefault(cal_id, {})
            return cal
        if method == "GET" and len(parts) == 2 and parts[0] == "calendars":
            cal = self.state["calendars"].get(parts[1])
            if cal is None:
                raise GcalError(f"GET {path} -> HTTP 404: calendar not "
                                "found")
            return cal
        if method == "POST" and len(parts) == 3 and parts[2] == "acl":
            acl = self.state["acls"].setdefault(parts[1], [])
            email = body.get("scope", {}).get("value")
            acl[:] = [a for a in acl
                      if a.get("scope", {}).get("value") != email]
            acl.append(body)
            return body
        if len(parts) >= 3 and parts[0] == "calendars" \
                and parts[2] == "events":
            cal_id = parts[1]
            if cal_id not in self.state["calendars"]:
                raise GcalError(f"{method} {path} -> HTTP 404: calendar "
                                "not found")
            events = self.state["events"].setdefault(cal_id, {})
            if method == "GET":
                items = list(events.values())
                for want in ([params["privateExtendedProperty"]]
                             if isinstance(
                                 params.get("privateExtendedProperty"),
                                 str)
                             else params.get("privateExtendedProperty",
                                             [])):
                    key, _, value = want.partition("=")
                    items = [e for e in items
                             if e.get("extendedProperties", {})
                             .get("private", {}).get(key) == value]
                return {"items": items}
            if method == "POST":
                event = dict(body, id=self._next("ev"))
                events[event["id"]] = event
                return event
            if method == "PATCH" and len(parts) == 4:
                event = events.get(parts[3])
                if event is None:
                    raise GcalError(f"PATCH {path} -> HTTP 404")
                self._merge_patch(event, body)
                return event
        raise GcalError(f"fake transport: unhandled {method} {path}")


# --- OAuth -------------------------------------------------------------------

def _client_info(credentials_file):
    creds = json.loads(Path(credentials_file).read_text(encoding="utf-8"))
    info = creds.get("installed") or creds.get("web")
    if not info:
        raise GcalError(f"{credentials_file} is not an OAuth client JSON "
                        "(expected an 'installed' section) — see "
                        "tools/GOOGLE-SETUP.md")
    return info


def _token_post(data):
    req = urllib.request.Request(
        TOKEN_URL, data=urllib.parse.urlencode(data).encode("utf-8"),
        headers={"Content-Type": "application/x-www-form-urlencoded"})
    try:
        with urllib.request.urlopen(req, timeout=30) as resp:
            return json.loads(resp.read())
    except urllib.error.HTTPError as e:
        raise GcalError(f"token endpoint HTTP {e.code}: "
                        f"{e.read()[:300].decode('utf-8', 'replace')}")
    except urllib.error.URLError as e:
        raise GcalError(f"token endpoint unreachable: {e.reason}")


def refresh_access_token(credentials_file, token_file) -> str:
    client = _client_info(credentials_file)
    token = json.loads(Path(token_file).read_text(encoding="utf-8"))
    resp = _token_post({
        "client_id": client["client_id"],
        "client_secret": client["client_secret"],
        "refresh_token": token["refresh_token"],
        "grant_type": "refresh_token",
    })
    return resp["access_token"]


def interactive_authorize(credentials_file, token_file) -> None:
    """One-time browser consent (gsetup only): loopback redirect, code
    exchange, refresh token stored at token_file (outside the repo)."""
    import webbrowser
    from http.server import BaseHTTPRequestHandler, HTTPServer

    client = _client_info(credentials_file)
    holder = {}

    class Handler(BaseHTTPRequestHandler):
        def do_GET(self):
            query = urllib.parse.parse_qs(
                urllib.parse.urlparse(self.path).query)
            for field in ("code", "error"):
                if field in query:
                    holder[field] = query[field][0]
            self.send_response(200)
            self.send_header("Content-Type", "text/plain; charset=utf-8")
            self.end_headers()
            self.wfile.write("Authorised — you can close this window and "
                             "return to the terminal.".encode("utf-8"))

        def log_message(self, *args):
            pass

    server = HTTPServer(("127.0.0.1", 0), Handler)
    redirect = f"http://127.0.0.1:{server.server_address[1]}/"
    url = AUTH_URL + "?" + urllib.parse.urlencode({
        "client_id": client["client_id"],
        "redirect_uri": redirect,
        "response_type": "code",
        "scope": SCOPE,
        "access_type": "offline",
        "prompt": "consent",
    })
    print("\nOpening the Google consent page in your browser…")
    print(f"(If nothing opens, paste this URL yourself:)\n{url}\n")
    webbrowser.open(url)
    while "code" not in holder and "error" not in holder:
        server.handle_request()
    server.server_close()
    if holder.get("error"):
        raise GcalError(f"consent refused: {holder['error']}")
    resp = _token_post({
        "client_id": client["client_id"],
        "client_secret": client["client_secret"],
        "code": holder["code"],
        "redirect_uri": redirect,
        "grant_type": "authorization_code",
    })
    if "refresh_token" not in resp:
        raise GcalError("Google returned no refresh token — remove the "
                        "app's access at myaccount.google.com/permissions "
                        "and run setup again")
    Path(token_file).parent.mkdir(parents=True, exist_ok=True)
    Path(token_file).write_text(
        json.dumps({"refresh_token": resp["refresh_token"]}),
        encoding="utf-8")


def get_transport(cfg):
    """Fake if the test hook is set; otherwise a real transport with a
    freshly refreshed access token."""
    fake = os.environ.get(FAKE_ENV)
    if fake:
        return FakeTransport(fake)
    if cfg is None or not cfg.credentials_file or not cfg.token_file:
        raise GcalError("Google is not configured — run "
                        "python tools/gsetup.py (see tools/GOOGLE-SETUP.md)")
    if not Path(cfg.token_file).exists():
        raise GcalError("not authorised yet — run python tools/gsetup.py "
                        "to sign in once")
    return Transport(refresh_access_token(cfg.credentials_file,
                                          cfg.token_file))


# --- thin API wrapper --------------------------------------------------------

class CalendarClient:
    def __init__(self, transport):
        self.t = transport

    def create_calendar(self, summary) -> str:
        return self.t.request("POST", "/calendars",
                              body={"summary": summary})["id"]

    def calendar_exists(self, cal_id) -> bool:
        try:
            self.t.request("GET", f"/calendars/{cal_id}")
            return True
        except GcalError as e:
            if "404" in str(e):
                return False
            raise

    def ensure_listed(self, cal_id) -> None:
        """Re-add a calendar to the owner's visible list. Owners can
        'unsubscribe' (hide) a calendar without deleting it — the
        calendar survives, so setup keeps it, but it looks vanished.
        calendarList insert is an upsert: harmless when already
        listed."""
        self.t.request("POST", "/users/me/calendarList",
                       body={"id": cal_id})

    def share(self, cal_id, email, role="writer") -> None:
        self.t.request("POST", f"/calendars/{cal_id}/acl",
                       body={"role": role,
                             "scope": {"type": "user", "value": email}})

    def list_tagged_events(self, cal_id) -> dict:
        """All events we manage (tagged crisisplan=1), keyed by planKey.
        Untagged events are never even fetched (AC-I5)."""
        events, token = {}, None
        while True:
            params = {"privateExtendedProperty": "crisisplan=1",
                      "maxResults": "2500", "showDeleted": "false"}
            if token:
                params["pageToken"] = token
            page = self.t.request("GET", f"/calendars/{cal_id}/events",
                                  params=params)
            for event in page.get("items", []):
                key = event.get("extendedProperties", {}) \
                    .get("private", {}).get("planKey")
                if key:
                    events[key] = event
            token = page.get("nextPageToken")
            if not token:
                return events

    def insert_event(self, cal_id, body):
        return self.t.request("POST", f"/calendars/{cal_id}/events",
                              params={"sendUpdates": "none"}, body=body)

    def patch_event(self, cal_id, event_id, body):
        return self.t.request("PATCH",
                              f"/calendars/{cal_id}/events/{event_id}",
                              params={"sendUpdates": "none"}, body=body)
