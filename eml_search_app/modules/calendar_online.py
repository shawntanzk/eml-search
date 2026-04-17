"""
Online calendar fetchers — iCal URL (Apple iCloud, Google Calendar, Outlook.com,
any CalDAV feed) and Microsoft Graph API (Microsoft 365 / Outlook).

Multiple accounts are supported simultaneously; `fetch_all_accounts` merges them
into one sorted list, tagging each event with `_account_id`, `_account_name`,
and `_account_color` so the UI can colour-code events per source.

Normalised event keys
---------------------
id, subject, start_time, end_time, start_dt, end_dt, body,
organizer, time_zone ("UTC"), required_attendees, optional_attendees, all_emails,
_account_id, _account_name, _account_color
"""

import re
import threading
import time
from datetime import date, datetime, timedelta, timezone
from typing import Optional

import requests

# ── Per-account TTL caches ────────────────────────────────────────────────────
_cache_lock    = threading.Lock()
# account_id → {"events": [...], "fetched_at": float, "key": str}
_ical_caches:  dict[str, dict] = {}
_graph_caches: dict[str, dict] = {}

GRAPH_CALENDAR_SCOPE = ["https://graph.microsoft.com/Calendars.Read"]
MICROSOFT_AUTHORITY  = "https://login.microsoftonline.com/consumers"

# Default color palette — assigned round-robin when user doesn't pick
ACCOUNT_COLORS = [
    "#4a6cf7",  # blue
    "#e64a4a",  # red
    "#2da44e",  # green
    "#f0a500",  # amber
    "#9b59b6",  # purple
    "#1abc9c",  # teal
    "#e67e22",  # orange
    "#3498db",  # sky blue
]


def next_account_color(existing_accounts: list[dict]) -> str:
    """Return the next color from the palette not yet used by existing accounts."""
    used = {a.get("color") for a in existing_accounts}
    for c in ACCOUNT_COLORS:
        if c not in used:
            return c
    return ACCOUNT_COLORS[len(existing_accounts) % len(ACCOUNT_COLORS)]


# ── High-level: fetch all accounts and merge ──────────────────────────────────

def fetch_all_accounts(
    accounts: list[dict],
    refresh_minutes: int = 15,
) -> tuple[list[dict], dict[str, str]]:
    """
    Fetch events from every enabled account and return a merged, date-sorted list.

    Returns ``(all_events, errors)`` where *errors* maps account ``id`` to an
    error string for any account that failed.  Successful accounts still appear
    in *all_events* even if others failed.

    Each event dict gains three extra keys:
    ``_account_id``, ``_account_name``, ``_account_color``.
    """
    ttl = refresh_minutes * 60
    all_events: list[dict] = []
    errors: dict[str, str] = {}

    for acc in accounts:
        if not acc.get("enabled", True):
            continue

        acc_id    = acc.get("id", "")
        acc_name  = acc.get("name", "Calendar")
        acc_color = acc.get("color", ACCOUNT_COLORS[0])
        acc_type  = acc.get("type", "ical")

        events: list[dict] = []
        err = ""

        if acc_type == "ical":
            events, err = fetch_ical(
                acc.get("url", ""),
                acc.get("username", ""),
                acc.get("password", ""),
                ttl=ttl,
                account_id=acc_id,
            )
        elif acc_type == "graph":
            events, err = fetch_graph_calendar(
                acc.get("access_token", ""),
                days_back=int(acc.get("days_back", 30)),
                days_forward=int(acc.get("days_forward", 90)),
                ttl=ttl,
                account_id=acc_id,
            )
        elif acc_type == "json":
            from modules import calendar_reader as _cr
            events = _cr.load_events(acc.get("path", ""))

        if err:
            errors[acc_id] = err

        for ev in events:
            tagged = dict(ev)
            tagged["_account_id"]    = acc_id
            tagged["_account_name"]  = acc_name
            tagged["_account_color"] = acc_color
            all_events.append(tagged)

    all_events.sort(key=lambda e: e.get("start_dt") or datetime.min)
    return all_events, errors


def invalidate_account_cache(account_id: str = "") -> None:
    """Invalidate the iCal and Graph caches for one account (or all if id is empty)."""
    with _cache_lock:
        if account_id:
            _ical_caches.pop(account_id, None)
            _graph_caches.pop(account_id, None)
        else:
            _ical_caches.clear()
            _graph_caches.clear()


def last_fetched_str(account_id: str = "") -> str:
    """'5m ago' / 'never' for a specific account, or the oldest across all accounts."""
    with _cache_lock:
        caches = list(_ical_caches.values()) + list(_graph_caches.values())
    if account_id:
        with _cache_lock:
            entry = _ical_caches.get(account_id) or _graph_caches.get(account_id)
        ts = entry["fetched_at"] if entry else 0.0
    elif caches:
        ts = min(c["fetched_at"] for c in caches)
    else:
        ts = 0.0
    if ts == 0.0:
        return "never"
    elapsed = int(time.time() - ts)
    if elapsed < 60:
        return f"{elapsed}s ago"
    if elapsed < 3600:
        return f"{elapsed // 60}m ago"
    return f"{elapsed // 3600}h ago"


# ── iCal URL ──────────────────────────────────────────────────────────────────

def fetch_ical(
    url: str,
    username: str = "",
    password: str = "",
    ttl: int = 900,
    account_id: str = "",
) -> tuple[list[dict], str]:
    """
    Fetch and parse an ICS feed from *url*.  Returns ``(events, error_str)``.
    webcal:// is automatically promoted to https://.
    Results are cached per account_id (or per url) for *ttl* seconds.
    """
    url = re.sub(r"^webcal://", "https://", url.strip())
    cache_key = f"{url}|{username}"
    key = account_id or cache_key

    with _cache_lock:
        cached = _ical_caches.get(key, {})
        if cached.get("key") == cache_key and time.time() - cached.get("fetched_at", 0) < ttl:
            return list(cached["events"]), ""

    try:
        auth = (username, password) if username else None
        resp = requests.get(url, auth=auth, timeout=30)
        resp.raise_for_status()
    except Exception as exc:
        with _cache_lock:
            return list((_ical_caches.get(key) or {}).get("events", [])), f"Fetch failed: {exc}"

    events, err = _parse_ics_bytes(resp.content)

    with _cache_lock:
        _ical_caches[key] = {"events": events, "fetched_at": time.time(), "key": cache_key}

    return events, err


def _parse_ics_bytes(data: bytes) -> tuple[list[dict], str]:
    try:
        from icalendar import Calendar
    except ImportError:
        return [], "icalendar not installed — run: pip install icalendar"
    try:
        cal = Calendar.from_ical(data)
    except Exception as exc:
        return [], f"Could not parse ICS data: {exc}"

    events: list[dict] = []
    for component in cal.walk():
        if component.name != "VEVENT":
            continue
        ev = _parse_vevent(component)
        if ev:
            events.append(ev)

    events.sort(key=lambda e: e["start_dt"] or datetime.min)
    return events, ""


def _ical_to_utc(val) -> Optional[datetime]:
    if val is None:
        return None
    if type(val) is date:
        return datetime(val.year, val.month, val.day, 0, 0, 0)
    if isinstance(val, datetime):
        if val.tzinfo is not None:
            val = val.astimezone(timezone.utc).replace(tzinfo=None)
        return val
    try:
        return _ical_to_utc(val.dt)
    except AttributeError:
        return None


def _ical_addr(val) -> str:
    if val is None:
        return ""
    return re.sub(r"^mailto:", "", str(val), flags=re.IGNORECASE).strip().lower()


def _parse_vevent(comp) -> Optional[dict]:
    subject  = (str(comp.get("SUMMARY") or "")).strip() or "(no subject)"
    dtstart  = comp.get("DTSTART")
    dtend    = comp.get("DTEND")
    start_dt = _ical_to_utc(dtstart.dt if dtstart else None)
    end_dt   = _ical_to_utc(dtend.dt   if dtend   else None)
    uid      = (str(comp.get("UID") or "")).strip()
    body     = (str(comp.get("DESCRIPTION") or "")).strip()
    organizer = _ical_addr(comp.get("ORGANIZER"))

    raw_att = comp.get("ATTENDEE")
    att_values = raw_att if isinstance(raw_att, list) else ([raw_att] if raw_att else [])
    required = [a for a in (_ical_addr(v) for v in att_values) if a]

    seen: set[str] = set()
    all_emails: list[str] = []
    for addr in [organizer] + required:
        if addr and addr not in seen:
            seen.add(addr)
            all_emails.append(addr)

    return {
        "id":                  uid or f"ical-{hash(subject + str(start_dt))}",
        "subject":             subject,
        "start_time":          start_dt.isoformat() if start_dt else "",
        "end_time":            end_dt.isoformat()   if end_dt   else "",
        "start_dt":            start_dt,
        "end_dt":              end_dt,
        "body":                body,
        "organizer":           organizer,
        "time_zone":           "UTC",
        "required_attendees":  required,
        "optional_attendees":  [],
        "all_emails":          all_emails,
    }


# ── Microsoft Graph Calendar ──────────────────────────────────────────────────

def fetch_graph_calendar(
    access_token: str,
    days_back: int = 30,
    days_forward: int = 90,
    ttl: int = 900,
    account_id: str = "",
) -> tuple[list[dict], str]:
    """
    Fetch events from the signed-in user's primary calendar via Microsoft Graph.
    Returns ``(events, error_str)``.  Results cached per account_id for *ttl* seconds.
    """
    cache_key = access_token[:32] if access_token else ""
    key = account_id or cache_key

    with _cache_lock:
        cached = _graph_caches.get(key, {})
        if cached.get("key") == cache_key and time.time() - cached.get("fetched_at", 0) < ttl:
            return list(cached["events"]), ""

    now   = datetime.now(timezone.utc)
    start = (now - timedelta(days=days_back)).strftime("%Y-%m-%dT%H:%M:%SZ")
    end   = (now + timedelta(days=days_forward)).strftime("%Y-%m-%dT%H:%M:%SZ")

    headers = {"Authorization": f"Bearer {access_token}"}
    url: Optional[str] = "https://graph.microsoft.com/v1.0/me/calendarView"
    params: Optional[dict] = {
        "startDateTime": start,
        "endDateTime":   end,
        "$select":       "id,subject,start,end,bodyPreview,organizer,attendees",
        "$top":          "999",
    }

    events: list[dict] = []
    err = ""
    while url:
        try:
            resp = requests.get(url, headers=headers, params=params, timeout=30)
            resp.raise_for_status()
        except Exception as exc:
            err = str(exc)
            break
        data = resp.json()
        if "error" in data:
            err = data["error"].get("message", str(data["error"]))
            break
        for item in data.get("value", []):
            ev = _parse_graph_event(item)
            if ev:
                events.append(ev)
        url    = data.get("@odata.nextLink")
        params = None

    events.sort(key=lambda e: e["start_dt"] or datetime.min)

    if events or not err:
        with _cache_lock:
            _graph_caches[key] = {"events": events, "fetched_at": time.time(), "key": cache_key}

    with _cache_lock:
        return list((_graph_caches.get(key) or {}).get("events", [])), err


def _graph_dt(obj: dict) -> Optional[datetime]:
    if not obj:
        return None
    s  = obj.get("dateTime", "")
    tz = obj.get("timeZone", "UTC")
    if not s:
        return None
    s = re.sub(r"(\.\d{1,6})\d*", r"\1", s)
    s = re.sub(r"Z$|[+-]\d{2}:\d{2}$", "", s)
    for fmt in ("%Y-%m-%dT%H:%M:%S.%f", "%Y-%m-%dT%H:%M:%S"):
        try:
            dt = datetime.strptime(s, fmt)
            break
        except ValueError:
            continue
    else:
        return None
    try:
        from zoneinfo import ZoneInfo
        dt = dt.replace(tzinfo=ZoneInfo(tz)).astimezone(timezone.utc).replace(tzinfo=None)
    except Exception:
        pass
    return dt


def _parse_graph_event(item: dict) -> Optional[dict]:
    subject   = (item.get("subject") or "(no subject)").strip()
    start_dt  = _graph_dt(item.get("start") or {})
    end_dt    = _graph_dt(item.get("end")   or {})
    uid       = item.get("id", "")
    body      = (item.get("bodyPreview") or "").strip()
    org_addr  = ((item.get("organizer") or {}).get("emailAddress") or {})
    organizer = (org_addr.get("address") or "").lower().strip()

    required: list[str] = []
    optional: list[str] = []
    for att in item.get("attendees") or []:
        addr = ((att.get("emailAddress") or {}).get("address") or "").lower().strip()
        if not addr:
            continue
        if (att.get("type") or "required").lower() == "optional":
            optional.append(addr)
        else:
            required.append(addr)

    seen: set[str] = set()
    all_emails: list[str] = []
    for addr in [organizer] + required + optional:
        if addr and addr not in seen:
            seen.add(addr)
            all_emails.append(addr)

    return {
        "id":                  uid,
        "subject":             subject,
        "start_time":          start_dt.isoformat() if start_dt else "",
        "end_time":            end_dt.isoformat()   if end_dt   else "",
        "start_dt":            start_dt,
        "end_dt":              end_dt,
        "body":                body,
        "organizer":           organizer,
        "time_zone":           "UTC",
        "required_attendees":  required,
        "optional_attendees":  optional,
        "all_emails":          all_emails,
    }


# ── Graph token acquisition (device-code flow) ────────────────────────────────

def start_graph_device_flow(client_id: str) -> tuple[dict, str]:
    try:
        import msal
    except ImportError:
        return {}, "msal not installed — run: pip install msal"
    try:
        app  = msal.PublicClientApplication(client_id, authority=MICROSOFT_AUTHORITY)
        flow = app.initiate_device_flow(scopes=GRAPH_CALENDAR_SCOPE)
        if "user_code" not in flow:
            return {}, flow.get("error_description", "Could not start device flow")
        return flow, ""
    except Exception as exc:
        return {}, str(exc)


def complete_graph_device_flow(client_id: str, flow: dict) -> tuple[dict, str]:
    try:
        import msal
    except ImportError:
        return {}, "msal not installed"
    try:
        app    = msal.PublicClientApplication(client_id, authority=MICROSOFT_AUTHORITY)
        result = app.acquire_token_by_device_flow(flow)
        if "access_token" not in result:
            return {}, result.get("error_description", "Authentication failed")
        return result, ""
    except Exception as exc:
        return {}, str(exc)
