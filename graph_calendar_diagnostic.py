#!/usr/bin/env python3
"""Microsoft Graph Calendar Diagnostic.

Uses Microsoft Graph application OAuth (client credentials) to read a user's
default calendar via the ``calendarView`` endpoint, which expands recurring
series into individual occurrences and exceptions for the requested period.

The events returned by ``calendarView`` are grouped by their series master
ID (``seriesMasterId``). Non-recurring events do not have a series master, so
each one is treated as its own master record.

The script produces:
  * A text report with per-master details and overall totals.
  * A CSV report listing recurring masters with more than 20 modified
    exceptions.
  * A CSV report listing masters whose series spans more than one year.
"""

from __future__ import annotations

import argparse
import csv
import os
import sys
from dataclasses import dataclass, field
from datetime import datetime, timedelta, timezone
from typing import Any, Dict, Iterable, List, Optional, Tuple

import requests

GRAPH_BASE_URL = "https://graph.microsoft.com/v1.0"
GRAPH_SCOPE = "https://graph.microsoft.com/.default"

# A recurring master is flagged if it has more than this many modified
# (type == "exception") occurrences.
EXCEPTION_THRESHOLD = 20

# A master is flagged if the span between its earliest start and latest end
# exceeds this many days (roughly one year).
DURATION_THRESHOLD_DAYS = 365

DATETIME_FORMATS = (
    "%Y-%m-%dT%H:%M:%S.%f",
    "%Y-%m-%dT%H:%M:%S",
)


def parse_graph_datetime(value: Optional[str]) -> Optional[datetime]:
    """Parse a Graph ``dateTime`` string (naive, timezone info is separate)."""
    if not value:
        return None
    for fmt in DATETIME_FORMATS:
        try:
            return datetime.strptime(value, fmt)
        except ValueError:
            continue
    # Fall back to fromisoformat for other variants (e.g. with offsets).
    try:
        return datetime.fromisoformat(value)
    except ValueError:
        return None


def get_access_token(tenant_id: str, client_id: str, client_secret: str) -> str:
    """Acquire an application (client credentials) OAuth token for Graph."""
    import msal

    authority = f"https://login.microsoftonline.com/{tenant_id}"
    app = msal.ConfidentialClientApplication(
        client_id, authority=authority, client_credential=client_secret
    )
    result = app.acquire_token_for_client(scopes=[GRAPH_SCOPE])
    if not result or "access_token" not in result:
        error = (result or {}).get("error_description", "unknown error")
        raise RuntimeError(f"Failed to acquire access token: {error}")
    return result["access_token"]


def fetch_calendar_view(
    access_token: str,
    user_id: str,
    start_date_time: str,
    end_date_time: str,
    session: Optional[requests.Session] = None,
    page_size: int = 100,
) -> List[Dict[str, Any]]:
    """Fetch all events in the calendar view for a user, following paging."""
    session = session or requests.Session()
    headers = {
        "Authorization": "Bearer " + access_token,
        "Prefer": 'outlook.timezone="UTC"',
    }
    url = f"{GRAPH_BASE_URL}/users/{user_id}/calendarView"
    params = {
        "startDateTime": start_date_time,
        "endDateTime": end_date_time,
        "$top": page_size,
        "$orderby": "start/dateTime",
    }

    events: List[Dict[str, Any]] = []
    while url:
        response = session.get(url, headers=headers, params=params, timeout=30)
        response.raise_for_status()
        payload = response.json()
        events.extend(payload.get("value", []))
        url = payload.get("@odata.nextLink")
        # Query params are already encoded into nextLink; don't resend them.
        params = None

    return events


@dataclass
class MasterRecord:
    """Aggregated information about a series master (or single event)."""

    master_id: str
    subject: str
    is_recurring: bool
    occurrence_count: int = 0
    exception_count: int = 0
    cancelled_count: int = 0
    first_start: Optional[datetime] = None
    last_end: Optional[datetime] = None
    event_ids: List[str] = field(default_factory=list)

    @property
    def total_events(self) -> int:
        return self.occurrence_count + self.exception_count

    @property
    def duration_days(self) -> Optional[float]:
        if self.first_start is None or self.last_end is None:
            return None
        return (self.last_end - self.first_start).total_seconds() / 86400.0

    @property
    def exceeds_exception_threshold(self) -> bool:
        return self.is_recurring and self.exception_count > EXCEPTION_THRESHOLD

    @property
    def exceeds_duration_threshold(self) -> bool:
        duration = self.duration_days
        return duration is not None and duration > DURATION_THRESHOLD_DAYS


def group_events(events: Iterable[Dict[str, Any]]) -> Dict[str, MasterRecord]:
    """Group calendarView events by series master, keyed by master id.

    Non-recurring events (``type`` == ``singleInstance`` and no
    ``seriesMasterId``) are treated as their own master record, keyed by
    their own event id.
    """
    masters: Dict[str, MasterRecord] = {}

    for event in events:
        event_type = event.get("type", "singleInstance")
        master_id = event.get("seriesMasterId") or event.get("id")
        is_recurring = event_type in ("occurrence", "exception") and bool(
            event.get("seriesMasterId")
        )

        record = masters.get(master_id)
        if record is None:
            record = MasterRecord(
                master_id=master_id,
                subject=event.get("subject") or "(no subject)",
                is_recurring=is_recurring,
            )
            masters[master_id] = record
        else:
            record.is_recurring = record.is_recurring or is_recurring

        record.event_ids.append(event.get("id"))

        if event.get("isCancelled"):
            record.cancelled_count += 1

        if event_type == "exception":
            record.exception_count += 1
        else:
            record.occurrence_count += 1

        start = parse_graph_datetime((event.get("start") or {}).get("dateTime"))
        end = parse_graph_datetime((event.get("end") or {}).get("dateTime"))

        if start is not None and (record.first_start is None or start < record.first_start):
            record.first_start = start
        if end is not None and (record.last_end is None or end > record.last_end):
            record.last_end = end

    return masters


def sorted_masters(masters: Dict[str, MasterRecord]) -> List[MasterRecord]:
    return sorted(
        masters.values(),
        key=lambda m: (m.first_start or datetime.max, m.subject or ""),
    )


def format_dt(dt: Optional[datetime]) -> str:
    return dt.strftime("%Y-%m-%d %H:%M:%S") if dt else "n/a"


def write_text_report(masters: Iterable[MasterRecord], path: str) -> None:
    records = sorted_masters({m.master_id: m for m in masters})

    total_masters = len(records)
    total_recurring = sum(1 for r in records if r.is_recurring)
    total_events = sum(r.total_events for r in records)
    total_exceptions = sum(r.exception_count for r in records)
    total_cancelled = sum(r.cancelled_count for r in records)

    with open(path, "w", encoding="utf-8") as fh:
        fh.write("Microsoft Graph Calendar Diagnostic Report\n")
        fh.write("=" * 43 + "\n\n")

        for record in records:
            fh.write(f"Master ID: {record.master_id}\n")
            fh.write(f"  Subject: {record.subject}\n")
            fh.write(f"  Recurring: {'yes' if record.is_recurring else 'no'}\n")
            fh.write(f"  Occurrences: {record.occurrence_count}\n")
            fh.write(f"  Modified exceptions: {record.exception_count}\n")
            fh.write(f"  Cancelled instances: {record.cancelled_count}\n")
            fh.write(f"  First start: {format_dt(record.first_start)}\n")
            fh.write(f"  Last end: {format_dt(record.last_end)}\n")
            duration = record.duration_days
            fh.write(
                f"  Span (days): {duration:.1f}\n" if duration is not None else "  Span (days): n/a\n"
            )
            if record.exceeds_exception_threshold:
                fh.write("  ** Exceeds modified exception threshold **\n")
            if record.exceeds_duration_threshold:
                fh.write("  ** Series spans more than one year **\n")
            fh.write("\n")

        fh.write("Totals\n")
        fh.write("-" * 6 + "\n")
        fh.write(f"Total masters: {total_masters}\n")
        fh.write(f"Recurring masters: {total_recurring}\n")
        fh.write(f"Non-recurring masters: {total_masters - total_recurring}\n")
        fh.write(f"Total events (occurrences + exceptions): {total_events}\n")
        fh.write(f"Total modified exceptions: {total_exceptions}\n")
        fh.write(f"Total cancelled instances: {total_cancelled}\n")


def write_csv_report(
    masters: Iterable[MasterRecord],
    path: str,
    predicate,
    extra_field: str,
    extra_value_fn,
) -> None:
    records = [m for m in sorted_masters({m.master_id: m for m in masters}) if predicate(m)]

    fieldnames = [
        "master_id",
        "subject",
        "occurrence_count",
        "exception_count",
        "first_start",
        "last_end",
        "duration_days",
        extra_field,
    ]

    with open(path, "w", newline="", encoding="utf-8") as fh:
        writer = csv.DictWriter(fh, fieldnames=fieldnames)
        writer.writeheader()
        for record in records:
            duration = record.duration_days
            writer.writerow(
                {
                    "master_id": record.master_id,
                    "subject": record.subject,
                    "occurrence_count": record.occurrence_count,
                    "exception_count": record.exception_count,
                    "first_start": format_dt(record.first_start),
                    "last_end": format_dt(record.last_end),
                    "duration_days": f"{duration:.1f}" if duration is not None else "",
                    extra_field: extra_value_fn(record),
                }
            )


def write_exception_report(masters: Iterable[MasterRecord], path: str) -> None:
    write_csv_report(
        masters,
        path,
        predicate=lambda m: m.exceeds_exception_threshold,
        extra_field="exception_threshold",
        extra_value_fn=lambda m: EXCEPTION_THRESHOLD,
    )


def write_long_series_report(masters: Iterable[MasterRecord], path: str) -> None:
    write_csv_report(
        masters,
        path,
        predicate=lambda m: m.exceeds_duration_threshold,
        extra_field="duration_threshold_days",
        extra_value_fn=lambda m: DURATION_THRESHOLD_DAYS,
    )


def default_date_range(days: int = 30) -> Tuple[str, str]:
    start = datetime.now(timezone.utc).replace(tzinfo=None)
    end = start + timedelta(days=days)
    fmt = "%Y-%m-%dT%H:%M:%S"
    return start.strftime(fmt), end.strftime(fmt)


def parse_args(argv: Optional[List[str]] = None) -> argparse.Namespace:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--tenant-id", required=True, help="Azure AD tenant ID")
    parser.add_argument("--client-id", required=True, help="App registration client ID")
    parser.add_argument(
        "--client-secret",
        default=None,
        help=(
            "App registration client secret. Prefer setting the "
            "GRAPH_CLIENT_SECRET environment variable instead of passing this "
            "on the command line, since command-line arguments can be visible "
            "in shell history and process listings."
        ),
    )
    parser.add_argument("--user", required=True, help="User principal name or object ID")
    parser.add_argument("--start", help="Start of the calendar view range (ISO 8601)")
    parser.add_argument("--end", help="End of the calendar view range (ISO 8601)")
    parser.add_argument(
        "--days", type=int, default=30, help="Number of days to look ahead when --start/--end are omitted"
    )
    parser.add_argument(
        "--report", default="calendar_report.txt", help="Path for the text report output"
    )
    parser.add_argument(
        "--exceptions-csv",
        default="recurring_exceptions.csv",
        help="Path for the CSV report of recurring masters with too many modified exceptions",
    )
    parser.add_argument(
        "--long-series-csv",
        default="long_series.csv",
        help="Path for the CSV report of masters spanning more than one year",
    )
    return parser.parse_args(argv)


def main(argv: Optional[List[str]] = None) -> int:
    args = parse_args(argv)

    client_secret = args.client_secret or os.environ.get("GRAPH_CLIENT_SECRET")
    if not client_secret:
        print(
            "Error: a client secret must be provided via --client-secret or "
            "the GRAPH_CLIENT_SECRET environment variable.",
            file=sys.stderr,
        )
        return 2

    start, end = args.start, args.end
    if not start or not end:
        default_start, default_end = default_date_range(args.days)
        start = start or default_start
        end = end or default_end

    access_token = get_access_token(args.tenant_id, args.client_id, client_secret)
    events = fetch_calendar_view(access_token, args.user, start, end)
    masters = group_events(events)

    write_text_report(masters.values(), args.report)
    write_exception_report(masters.values(), args.exceptions_csv)
    write_long_series_report(masters.values(), args.long_series_csv)

    print(f"Processed {len(events)} events into {len(masters)} master records.")
    print(f"Text report: {args.report}")
    print(f"Exception CSV report: {args.exceptions_csv}")
    print(f"Long series CSV report: {args.long_series_csv}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
