"""
MedSoc UNSW 2026 Calendar — Google Sheets → ICS converter
==========================================================
Usage:
    python medsoc_calendar.py --sheet-id SHEET_ID --api-key API_KEY [--output-dir ./output]

Outputs four ICS files:
    MedSoc_UNSW_2026_general.ics
    MedSoc_UNSW_2026_phase1.ics
    MedSoc_UNSW_2026_phase2.ics
    MedSoc_UNSW_2026_phase3.ics
"""

import argparse
import os
import re
import sys
import uuid
from datetime import datetime, timedelta, date
from typing import Optional
from googleapiclient.discovery import build

# ---------------------------------------------------------------------------
# Constants
# ---------------------------------------------------------------------------

YEAR = 2026
TIMEZONE = "Australia/Sydney"

# Portfolios that map to each phase feed
PHASE1_LABELS = {"p1 acads", "p1 acads+wcsoc", "2nd year reps", "second year reps"}
PHASE2_LABELS = {"p2 acads", "p2/p3 acads", "3rd year reps", "third year reps"}
PHASE3_LABELS = {"p2/p3 acads"}

# Known label-only / non-event cells to skip entirely
SKIP_EVENTS = {
    "o-week stalls", "o-day", "p-day", "love week", "ski trip",
    "amsa convention", "rahms rural high school visit", "rahms outreach trip",
    "medshow", "subcommittee interviews", "medcamp", "medcamp registration",
    "cultural week: wallace wurth foyer - hr + international",
    "student vs faculty week", "love letters",
}

# Day/month header patterns
DAY_NAMES = {"monday", "tuesday", "wednesday", "thursday", "friday", "saturday", "sunday"}
MONTH_NAMES = {
    "january", "february", "march", "april", "may", "june",
    "july", "august", "september", "october", "november", "december",
}

DATE_HEADER_RE = re.compile(r"^\d{1,2}/\d{1,2}$")

# Time parsing regex
# Matches formats like: 6:00-8:00pm  5:15pm-7:30pm  10:00am-5:00pm  9:30pm-Late
TIME_RE = re.compile(
    r"(\d{1,2}:\d{2})(am|pm)?\s*[-–]\s*(\d{1,2}:\d{2}|late)(am|pm)?",
    re.IGNORECASE,
)

# ---------------------------------------------------------------------------
# Time parsing helpers
# ---------------------------------------------------------------------------

def _to_24h(h: int, m: int, ampm: Optional[str]) -> tuple[int, int]:
    """Convert h/m + optional am/pm to 24-hour h/m."""
    if ampm:
        ampm = ampm.lower()
        if ampm == "pm" and h != 12:
            h += 12
        elif ampm == "am" and h == 12:
            h = 0
    else:
        # Fallback heuristic: assume PM for 1–11, AM for 12
        if 1 <= h <= 11:
            h += 12
    return h, m


def parse_time_range(raw: str) -> Optional[tuple[tuple[int, int], tuple[int, int]]]:
    """
    Parse a time range string into ((start_h, start_m), (end_h, end_m)) in 24h.
    Returns None if no time found.
    """
    m = TIME_RE.search(raw)
    if not m:
        return None

    start_str, start_ap, end_str, end_ap = m.group(1), m.group(2), m.group(3), m.group(4)

    # If start has no AM/PM marker, inherit from end
    effective_start_ap = start_ap or end_ap

    sh, sm = map(int, start_str.split(":"))
    sh, sm = _to_24h(sh, sm, effective_start_ap)

    if end_str.lower() == "late":
        eh, em = sh + 3, sm
        if eh >= 24:
            # Clamp to 23:59 — crossing midnight would require a next-day DTEND
            # which adds complexity not worth it for a single "Late" token.
            eh = 23
            em = 59
    else:
        eh, em = map(int, end_str.split(":"))
        eh, em = _to_24h(eh, em, end_ap or effective_start_ap)

    return (sh, sm), (eh, em)


def is_time_unknown(raw: str) -> bool:
    """True if the time field is TBC / TBD / ? or absent."""
    cleaned = raw.strip().upper()
    return cleaned in ("TBC", "TBD", "?", "")


# ---------------------------------------------------------------------------
# Event cell parsing
# ---------------------------------------------------------------------------

def parse_event_cell(cell: str) -> Optional[dict]:
    """
    Parse a single event cell into a dict with keys:
        title, time_raw, place, portfolio, has_time
    Returns None if the cell should be skipped.
    """
    cell = cell.strip()
    if not cell:
        return None

    lower = cell.lower()

    # Skip day/month headers
    if lower in DAY_NAMES or lower in MONTH_NAMES:
        return None
    if lower.startswith("key:") or lower.startswith("phase "):
        return None

    # Extract structured fields using greedy splits
    # Fields: Time: ... Place: ... Portfolio: ...
    time_raw = ""
    place = ""
    portfolio = ""

    time_match = re.search(r"Time:\s*(.+?)(?=\s+Place:|\s+Portfolio:|$)", cell, re.IGNORECASE)
    place_match = re.search(r"Place:\s*(.+?)(?=\s+Portfolio:|$)", cell, re.IGNORECASE)
    portfolio_match = re.search(r"Portfolio:\s*(.+?)$", cell, re.IGNORECASE)

    if time_match:
        time_raw = time_match.group(1).strip()
    if place_match:
        place = place_match.group(1).strip()
    if portfolio_match:
        portfolio = portfolio_match.group(1).strip()

    # Title = everything before the first known field keyword
    title = re.split(r"\s+(?:Time:|Place:|Portfolio:)", cell, maxsplit=1, flags=re.IGNORECASE)[0].strip()

    if not title:
        return None

    # Skip known non-event labels
    if title.lower() in SKIP_EVENTS:
        return None

    return {
        "title": title,
        "time_raw": time_raw,
        "place": place,
        "portfolio": portfolio,
    }


# ---------------------------------------------------------------------------
# Feed assignment
# ---------------------------------------------------------------------------

def assign_feeds(portfolio: str) -> list[str]:
    """Return which feeds this event belongs to. Always includes 'general' if not phase-specific."""
    p = portfolio.lower().strip()
    feeds = []
    if p in PHASE1_LABELS:
        feeds.append("phase1")
    if p in PHASE2_LABELS:
        feeds.append("phase2")
    if p in PHASE3_LABELS:
        feeds.append("phase3")
    if not feeds:
        feeds.append("general")
    return feeds


# ---------------------------------------------------------------------------
# ICS formatting
# ---------------------------------------------------------------------------

def fmt_dt(dt: datetime) -> str:
    return dt.strftime("%Y%m%dT%H%M%S")


def make_vevent(event_date: date, title: str, start_hm: tuple[int, int],
                end_hm: tuple[int, int], place: str) -> str:
    start = datetime(event_date.year, event_date.month, event_date.day, *start_hm)
    end = datetime(event_date.year, event_date.month, event_date.day, *end_hm)
    uid = str(uuid.uuid4())

    lines = [
        "BEGIN:VEVENT",
        f"UID:{uid}",
        f"DTSTART;TZID={TIMEZONE}:{fmt_dt(start)}",
        f"DTEND;TZID={TIMEZONE}:{fmt_dt(end)}",
        f"SUMMARY:{title}",
    ]
    if place:
        lines.append(f"LOCATION:{place}")
    # Suppress default notifications
    lines += [
        "BEGIN:VALARM",
        "ACTION:NONE",
        "TRIGGER:-PT0S",
        "END:VALARM",
        "END:VEVENT",
    ]
    return "\r\n".join(lines)


def make_ics(feed_name: str, vevents: list[str]) -> str:
    header = "\r\n".join([
        "BEGIN:VCALENDAR",
        "VERSION:2.0",
        "PRODID:-//MedSoc UNSW//Calendar 2026//EN",
        f"X-WR-CALNAME:MedSoc UNSW 2026 — {feed_name.title()}",
        "X-WR-TIMEZONE:" + TIMEZONE,
        "X-APPLE-DEFAULT-ALARM:FALSE",
        "CALSCALE:GREGORIAN",
        "METHOD:PUBLISH",
    ])
    footer = "END:VCALENDAR"
    body = "\r\n".join(vevents)
    return header + "\r\n" + body + "\r\n" + footer + "\r\n"


# ---------------------------------------------------------------------------
# Google Sheets fetching
# ---------------------------------------------------------------------------

def fetch_sheet_rows(sheet_id: str, api_key: str) -> list[list[str]]:
    """Fetch all rows from the first sheet tab via Sheets API v4."""
    service = build("sheets", "v4", developerKey=api_key)
    result = (
        service.spreadsheets()
        .values()
        .get(spreadsheetId=sheet_id, range="A:Z")
        .execute()
    )
    raw_rows = result.get("values", [])
    # Normalise: ensure every row has at least 7 columns
    return [row + [""] * max(0, 7 - len(row)) for row in raw_rows]


# ---------------------------------------------------------------------------
# Main parsing logic
# ---------------------------------------------------------------------------

def parse_rows(rows: list[list[str]]) -> list[dict]:
    """
    Walk the visual grid and return a flat list of parsed event dicts:
        { title, event_date, start_hm, end_hm, place, portfolio, feeds }
    """
    events = []
    # Maps column index (0-6) → date object for the current week
    week_dates: dict[int, Optional[date]] = {i: None for i in range(7)}

    for row in rows:
        # Detect date header row: any cell matching DD/MM
        date_cells = [DATE_HEADER_RE.match(str(c).strip()) for c in row[:7]]
        if any(date_cells):
            for col_idx, cell in enumerate(row[:7]):
                cell = str(cell).strip()
                if DATE_HEADER_RE.match(cell):
                    day, month = map(int, cell.split("/"))
                    try:
                        week_dates[col_idx] = date(YEAR, month, day)
                    except ValueError:
                        week_dates[col_idx] = None
                else:
                    week_dates[col_idx] = None
            continue

        # Event row — scan each column
        for col_idx, cell in enumerate(row[:7]):
            cell = str(cell).strip()
            if not cell:
                continue

            event_date = week_dates.get(col_idx)
            if event_date is None:
                continue  # No date context for this column

            parsed = parse_event_cell(cell)
            if parsed is None:
                continue

            # Determine time
            time_raw = parsed["time_raw"]
            if is_time_unknown(time_raw) or not time_raw:
                start_hm = (8, 0)
                end_hm = (9, 0)
            else:
                result = parse_time_range(time_raw)
                if result is None:
                    start_hm = (8, 0)
                    end_hm = (9, 0)
                else:
                    start_hm, end_hm = result

            feeds = assign_feeds(parsed["portfolio"])

            events.append({
                "title": parsed["title"],
                "event_date": event_date,
                "start_hm": start_hm,
                "end_hm": end_hm,
                "place": parsed["place"],
                "portfolio": parsed["portfolio"],
                "feeds": feeds,
            })

    return events


# ---------------------------------------------------------------------------
# Entry point
# ---------------------------------------------------------------------------

def main():
    parser = argparse.ArgumentParser(description="MedSoc UNSW 2026 Google Sheets → ICS")
    parser.add_argument("--sheet-id", required=True, help="Google Sheet ID (from URL)")
    parser.add_argument("--api-key", required=True, help="Google Sheets API key")
    parser.add_argument("--output-dir", default=".", help="Directory for ICS output files")
    args = parser.parse_args()

    os.makedirs(args.output_dir, exist_ok=True)

    print("Fetching sheet data...")
    rows = fetch_sheet_rows(args.sheet_id, args.api_key)
    print(f"  Fetched {len(rows)} rows")

    print("Parsing events...")
    events = parse_rows(rows)
    print(f"  Parsed {len(events)} events")

    # Bucket events into feeds
    feed_vevents: dict[str, list[str]] = {
        "general": [], "phase1": [], "phase2": [], "phase3": []
    }

    for ev in events:
        vevent = make_vevent(
            ev["event_date"], ev["title"],
            ev["start_hm"], ev["end_hm"], ev["place"]
        )
        for feed in ev["feeds"]:
            feed_vevents[feed].append(vevent)

    # Write ICS files
    feed_labels = {
        "general": "General",
        "phase1": "Phase 1",
        "phase2": "Phase 2",
        "phase3": "Phase 3",
    }
    for feed, label in feed_labels.items():
        filename = f"MedSoc_UNSW_2026_{feed}.ics"
        filepath = os.path.join(args.output_dir, filename)
        ics_content = make_ics(label, feed_vevents[feed])
        with open(filepath, "w", encoding="utf-8") as f:
            f.write(ics_content)
        print(f"  Written {filepath}  ({len(feed_vevents[feed])} events)")

    print("\nDone! Four ICS files generated.")
    print("\nNext steps:")
    print("  1. Push the ICS files to your GitHub repo")
    print("  2. Use raw GitHub URLs for WordPress download links:")
    print("     https://raw.githubusercontent.com/[username]/medsoc-calendar/main/MedSoc_UNSW_2026_general.ics")


if __name__ == "__main__":
    main()
