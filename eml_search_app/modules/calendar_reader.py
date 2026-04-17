"""
Calendar event reader, HTML renderer, and email correlator for EML Search.

Reads a JSON file of calendar events (produced by an external automation).
Expected fields per event:
    subject, start_time, end_time, body, id, organizer, time_zone,
    required_attendees, optional_attendees
"""
import json
import re
import threading
from datetime import date, datetime, timedelta
from pathlib import Path
from typing import Optional
from zoneinfo import ZoneInfo

_cache_lock = threading.Lock()
_events_cache: list[dict] = []
_cache_path: str = ""
_cache_mtime: float = -1.0


# ── Loading & parsing ─────────────────────────────────────────────────────────

def load_events(json_path: str) -> list[dict]:
    """Load and parse calendar events from a JSON file. Result is mtime-cached."""
    global _events_cache, _cache_mtime, _cache_path

    p = Path(json_path)
    if not p.exists():
        return []
    try:
        mtime = p.stat().st_mtime
    except OSError:
        return []

    with _cache_lock:
        if json_path == _cache_path and mtime == _cache_mtime:
            return list(_events_cache)

        try:
            raw = json.loads(p.read_text(encoding="utf-8"))
        except Exception:
            return list(_events_cache)
        
        # Accept: plain list, {"events": [...]}, {"value": [...]} (Graph API style),
        # {"body": [...]} (local automation export)
        if isinstance(raw, list):
            items = raw
        elif isinstance(raw, dict):
            items = raw.get("body") or raw.get("events") or raw.get("value") or []
        else:
            items = []


        parsed: list[dict] = []
        for item in items:
            ev = _parse_event(item)
            if ev:
                parsed.append(ev)

        parsed.sort(key=lambda e: e["start_dt"] or datetime.min)

        _events_cache = parsed
        _cache_mtime = mtime
        _cache_path = json_path
        return list(parsed)


def _to_str_list(v) -> list[str]:
    """Coerce attendees field to a flat list of strings."""
    if not v:
        return []
    if isinstance(v, list):
        return [str(x).strip() for x in v if x and str(x).strip()]
    if isinstance(v, str):
        return [e.strip() for e in re.split(r"[;,]", v) if e.strip()]
    return []


def _parse_event(raw: dict) -> Optional[dict]:
    """Normalise a raw event dict into a consistent structure."""
    start_dt = parse_dt(raw.get("start_time", ""))
    end_dt   = parse_dt(raw.get("end_time",   ""))

    req       = _to_str_list(raw.get("required_attendees"))
    opt       = _to_str_list(raw.get("optional_attendees"))
    organizer = (raw.get("organizer") or "").strip()

    # Deduplicated, order-preserving list of all participant emails
    seen: set[str] = set()
    all_emails: list[str] = []
    for addr in ([organizer] + req + opt):
        if addr and addr not in seen:
            seen.add(addr)
            all_emails.append(addr)

    return {
        "id":                  str(raw.get("id", "")),
        "subject":             (raw.get("subject") or "(no subject)").strip(),
        "start_time":          raw.get("start_time", ""),
        "end_time":            raw.get("end_time",   ""),
        "start_dt":            start_dt,
        "end_dt":              end_dt,
        "body":                (raw.get("body") or "").strip(),
        "organizer":           organizer,
        "time_zone":           (raw.get("time_zone") or "").strip(),
        "required_attendees":  req,
        "optional_attendees":  opt,
        "all_emails":          all_emails,
    }


def parse_dt(s: str) -> Optional[datetime]:
    """
    Parse calendar datetime strings such as '2026-03-18T01:00:00.0000000'.
    Returns a naive datetime (timezone info stripped).
    """
    if not s:
        return None
    # Trim sub-second precision to 6 digits (microseconds)
    s = re.sub(r"(\.\d{6})\d+", r"\1", s)
    # Strip trailing Z or ±HH:MM offset
    s = re.sub(r"Z$", "", s)
    s = re.sub(r"[+-]\d{2}:\d{2}$", "", s)
    for fmt in ("%Y-%m-%dT%H:%M:%S.%f", "%Y-%m-%dT%H:%M:%S", "%Y-%m-%d"):
        try:
            return datetime.strptime(s, fmt)
        except ValueError:
            continue
    return None


def convert_display_tz(
    events: list[dict],
    display_tz: str = "Asia/Singapore",
) -> list[dict]:
    """
    Return a new list of events with start_dt/end_dt converted from each
    event's own time_zone field (default UTC) to display_tz.
    The original event dicts are not mutated.
    """
    try:
        dst = ZoneInfo(display_tz)
    except Exception:
        return events

    result = []
    for ev in events:
        src_name = ev.get("time_zone") or "UTC"
        try:
            src = ZoneInfo(src_name)
        except Exception:
            src = ZoneInfo("UTC")

        ev2 = dict(ev)
        if ev2.get("start_dt"):
            ev2["start_dt"] = (
                ev2["start_dt"].replace(tzinfo=src).astimezone(dst).replace(tzinfo=None)
            )
        if ev2.get("end_dt"):
            ev2["end_dt"] = (
                ev2["end_dt"].replace(tzinfo=src).astimezone(dst).replace(tzinfo=None)
            )
        result.append(ev2)
    return result


# ── Filtering helpers ─────────────────────────────────────────────────────────

def events_for_date(events: list[dict], d: date) -> list[dict]:
    """Return events that span the given calendar date, sorted by start time."""
    result = []
    for ev in events:
        if ev["start_dt"] is None:
            continue
        start_d = ev["start_dt"].date()
        end_d   = ev["end_dt"].date() if ev["end_dt"] else start_d
        if start_d <= d <= end_d:
            result.append(ev)
    return sorted(result, key=lambda e: e["start_dt"])


def events_in_range(events: list[dict], start: date, end: date) -> list[dict]:
    """Return events that overlap [start, end] (inclusive), sorted by start time."""
    result = []
    for ev in events:
        if ev["start_dt"] is None:
            continue
        start_d = ev["start_dt"].date()
        end_d   = ev["end_dt"].date() if ev["end_dt"] else start_d
        if start_d <= end and end_d >= start:
            result.append(ev)
    return sorted(result, key=lambda e: e["start_dt"])


def fmt_time(dt: Optional[datetime]) -> str:
    """Format a datetime as HH:MM (24-hr)."""
    return dt.strftime("%H:%M") if dt else ""


def fmt_duration(ev: dict) -> str:
    """Return a human-readable duration string, e.g. '1 h 30 min'."""
    if not ev["start_dt"] or not ev["end_dt"]:
        return ""
    delta = ev["end_dt"] - ev["start_dt"]
    total_minutes = int(delta.total_seconds() // 60)
    if total_minutes <= 0:
        return ""
    hours, mins = divmod(total_minutes, 60)
    if hours and mins:
        return f"{hours} h {mins} min"
    if hours:
        return f"{hours} h"
    return f"{mins} min"


# ── HTML month calendar renderer ──────────────────────────────────────────────

def render_month_html(year: int, month: int, events: list[dict]) -> tuple[str, int]:
    """
    Render a month calendar as a self-contained HTML string.
    Events are shown as coloured chips inside each day cell.
    Past events are grey; today's events are dark blue; future events are blue.
    Returns (html_str, height_px) where height_px is the recommended iframe height.
    """
    import calendar as _cal

    today = date.today()

    # Build a map: date → sorted list of events for that day
    event_map: dict[date, list[dict]] = {}
    for ev in events:
        if ev["start_dt"] is None:
            continue
        start_d = ev["start_dt"].date()
        end_d   = ev["end_dt"].date() if ev["end_dt"] else start_d
        cur = start_d
        while cur <= end_d:
            if cur.year == year and cur.month == month:
                event_map.setdefault(cur, []).append(ev)
            cur += timedelta(days=1)

    for d in event_map:
        event_map[d].sort(key=lambda e: e["start_dt"])

    cal_weeks = _cal.monthcalendar(year, month)

    css = """
<style>
.ecal { font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', sans-serif;
        border-collapse: collapse; width: 100%; table-layout: fixed; }
.ecal th { background: #f0f2f6; padding: 6px 4px; text-align: center;
           font-size: 12px; font-weight: 600; color: #555; border: 1px solid #e0e4ea; }
.ecal td { vertical-align: top; border: 1px solid #e0e4ea;
           padding: 4px 5px; min-height: 84px; height: auto; width: 14.28%; overflow: visible; }
.ecal td.empty { background: #fafbfc; }
.ecal td.today { background: #eef3ff; border: 1.5px solid #4a6cf7; }
.ecal td.past  { background: #fafafa; }
.day-num { font-size: 12px; font-weight: 600; color: #333; margin-bottom: 3px; }
.day-num.today-num { color: #4a6cf7; }
.day-num.past-num  { color: #aaa; }
.ev-chip { display: block; border-radius: 3px; padding: 1px 5px;
           font-size: 10px; line-height: 1.5; margin-bottom: 2px;
           white-space: nowrap; overflow: hidden; text-overflow: ellipsis;
           cursor: default; color: #fff; }
.ev-more   { font-size: 10px; color: #4a6cf7; padding-left: 3px; }
</style>
"""

    day_names = ["Mon", "Tue", "Wed", "Thu", "Fri", "Sat", "Sun"]
    header = "".join(f"<th>{d}</th>" for d in day_names)

    rows_html = ""
    for week in cal_weeks:
        row = ""
        for day_num in week:
            if day_num == 0:
                row += '<td class="empty"></td>'
                continue

            d      = date(year, month, day_num)
            is_today = d == today
            is_past  = d < today

            if is_today:
                cell_cls = "today"
                num_cls  = "today-num"
            elif is_past:
                cell_cls = "past"
                num_cls  = "past-num"
            else:
                cell_cls = ""
                num_cls  = ""

            day_evs = event_map.get(d, [])
            chips = ""
            for ev in day_evs[:3]:
                t      = fmt_time(ev["start_dt"])
                label  = ev["subject"][:20] + ("…" if len(ev["subject"]) > 20 else "")
                prefix = f"{t} " if t else ""
                color  = ev.get("_account_color", "#4a6cf7")
                opacity = "0.45" if is_past else "1"
                chip_style = f"background:{color};opacity:{opacity}"
                chips += (
                    f'<span class="ev-chip" style="{chip_style}" title="{ev["subject"]}">'
                    f'{prefix}{label}</span>'
                )
            if len(day_evs) > 3:
                chips += f'<span class="ev-more">+{len(day_evs) - 3} more</span>'

            row += (
                f'<td class="{cell_cls}">'
                f'<div class="day-num {num_cls}">{day_num}</div>'
                f'{chips}</td>'
            )
        rows_html += f"<tr>{row}</tr>"

    html = f"{css}<table class='ecal'><thead><tr>{header}</tr></thead><tbody>{rows_html}</tbody></table>"
    # 38px header row + ~104px per week row + 16px bottom padding
    height = 38 + len(cal_weeks) * 104 + 16
    return html, height


# ── Email correlation ─────────────────────────────────────────────────────────

def _recency_multiplier(date_str: str, today: datetime) -> float:
    """Decay score for emails that are old relative to today."""
    if not date_str:
        return 0.5
    try:
        em_dt = datetime.fromisoformat(date_str[:19])
        days_old = max(0, (today - em_dt).days)
    except Exception:
        return 0.5
    if days_old <= 60:
        return 1.0
    if days_old <= 180:
        return 0.65
    if days_old <= 365:
        return 0.30
    return 0.08


def find_related_emails(event: dict, limit: int = 15, user_email: str = "") -> list[dict]:
    """
    Find emails related to a calendar event using a multi-signal approach:

    1. FTS search on event subject
    2. Semantic search on subject + body snippet
    3. Named entity matching (people / orgs in event text vs email_entities)
    4. Tag keyword matching (tags whose name appears in subject/body)
    5. Direct attendee / organizer email match (strongest signal)

    Results are RRF-merged then re-ranked by a recency multiplier that
    strongly penalises emails older than 60 days.
    """
    from modules import indexer, semantic_search, nlp_engine

    subject      = event.get("subject", "")
    body_snippet = (event.get("body") or "")[:600]
    all_emails   = event.get("all_emails", [])
    search_text  = f"{subject} {body_snippet}".strip()

    conn  = indexer._get_conn()
    today = datetime.utcnow()

    # ── 1. FTS ────────────────────────────────────────────────────────────────
    fts_results = indexer.search_fts(subject, filters={}, limit=200) if subject.strip() else []
    fts_rank = {r["id"]: i for i, r in enumerate(fts_results)}

    # ── 2. Semantic ───────────────────────────────────────────────────────────
    sem_rank: dict[str, int] = {}
    sem_ok, _ = semantic_search.model_status()
    if sem_ok and search_text:
        try:
            ids, matrix = indexer.get_cached_embeddings()
            if len(ids) > 0:
                vec   = semantic_search.embed_text(search_text)
                pairs = semantic_search.cosine_search(vec, ids, matrix, top_k=200)
                sem_rank = {eid: i for i, (eid, _) in enumerate(pairs)}
        except Exception:
            pass

    # ── 3. Named entity matching ──────────────────────────────────────────────
    entity_email_ids: set[str] = set()
    if nlp_engine.NLP_AVAILABLE() and search_text:
        try:
            entities = nlp_engine.extract_entities(search_text)
            for ent in entities:
                if ent["label"] in ("PERSON", "ORG", "GPE") and len(ent["text"]) > 2:
                    rows = conn.execute(
                        "SELECT DISTINCT email_id FROM email_entities WHERE entity_text LIKE ?",
                        (f"%{ent['text']}%",),
                    ).fetchall()
                    for r in rows:
                        entity_email_ids.add(r["email_id"])
        except Exception:
            pass

    # ── 4. Tag keyword matching ───────────────────────────────────────────────
    tag_email_ids: set[str] = set()
    subject_lower = subject.lower()
    body_lower    = body_snippet.lower()
    all_tags = conn.execute("SELECT id, name FROM tags").fetchall()
    for tag in all_tags:
        if tag["name"].lower() in subject_lower or tag["name"].lower() in body_lower:
            rows = conn.execute(
                "SELECT email_id FROM email_tags WHERE tag_id = ?", (tag["id"],)
            ).fetchall()
            for r in rows:
                tag_email_ids.add(r["email_id"])

    # ── 5. Attendee / organizer email match ───────────────────────────────────
    # Exclude the user's own address — they attend everything, zero signal.
    _user_addr = user_email.lower().strip()
    attendee_addrs = [
        addr.lower().strip() for addr in all_emails
        if addr and addr.lower().strip() and addr.lower().strip() != _user_addr
    ]

    attendee_email_ids: set[str] = set()
    if attendee_addrs:
        # Batch sender lookup (indexed column — fast)
        ph = ",".join("?" * len(attendee_addrs))
        for r in conn.execute(
            f"SELECT id FROM emails WHERE sender_email IN ({ph})", attendee_addrs
        ).fetchall():
            attendee_email_ids.add(r["id"])
        # JSON recipients/cc lookup — quote-wrapped for precision
        for addr in attendee_addrs:
            pattern = f'%"{addr}"%'
            for r in conn.execute(
                "SELECT id FROM emails WHERE recipients LIKE ? OR cc LIKE ?",
                (pattern, pattern),
            ).fetchall():
                attendee_email_ids.add(r["id"])

    # ── RRF merge ─────────────────────────────────────────────────────────────
    all_candidate_ids = (
        set(fts_rank)
        | set(sem_rank)
        | entity_email_ids
        | tag_email_ids
        | attendee_email_ids
    )

    K = 60
    raw_scores: dict[str, float] = {}
    for eid in all_candidate_ids:
        s = 0.0
        if eid in fts_rank:
            s += 2.0 / (K + fts_rank[eid])
        if eid in sem_rank:
            s += 1.0 / (K + sem_rank[eid])
        if eid in entity_email_ids:
            s += 0.005
        if eid in tag_email_ids:
            s += 0.003
        if eid in attendee_email_ids:
            s += 0.50
        raw_scores[eid] = s

    # Pre-fetch top candidates (3× limit gives room to re-rank by recency)
    candidate_ids = sorted(raw_scores, key=lambda x: -raw_scores[x])[: limit * 3]
    emails_map = indexer.get_emails_by_ids(candidate_ids)

    # Apply recency multiplier and re-rank
    final_scores = {
        eid: raw_scores[eid] * _recency_multiplier(
            (emails_map[eid].get("date") or "") if eid in emails_map else "", today
        )
        for eid in candidate_ids
        if eid in emails_map
    }

    ranked_ids = sorted(final_scores, key=lambda x: -final_scores[x])[:limit]

    results = []
    for eid in ranked_ids:
        em = dict(emails_map[eid])
        signals: list[str] = []
        if eid in attendee_email_ids:
            signals.append("👥 Attendee")
        if eid in fts_rank:
            signals.append("📝 Subject")
        if eid in sem_rank:
            signals.append("🔍 Semantic")
        if eid in entity_email_ids:
            signals.append("🏷 Entity")
        if eid in tag_email_ids:
            signals.append("🔖 Tag")
        em["_match_signals"] = signals
        results.append(em)
    return results


def tag_summary(email_ids: list[str]) -> list[dict]:
    """
    Return the most common tags across a list of email IDs,
    sorted by frequency descending.
    """
    from modules import indexer
    if not email_ids:
        return []
    conn = indexer._get_conn()
    ph   = ",".join("?" * len(email_ids))
    rows = conn.execute(
        f"""SELECT t.id, t.name, COUNT(*) AS cnt
            FROM email_tags et
            JOIN tags t ON t.id = et.tag_id
            WHERE et.email_id IN ({ph})
            GROUP BY t.id
            ORDER BY cnt DESC""",
        email_ids,
    ).fetchall()
    return [dict(r) for r in rows]
