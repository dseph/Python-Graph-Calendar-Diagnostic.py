# Python-Graph-Calendar-Diagnostic.py

"""Report calendar master records and recurrence exceptions via Microsoft Graph.

Required Microsoft Graph application permission: Calendars.Read (admin consent).
Set GRAPH_TENANT_ID, GRAPH_CLIENT_ID, GRAPH_CLIENT_SECRET, and GRAPH_USER_ID in the
environment before running this script. No third-party Python packages are required.
"""

from __future__ import annotations

import csv
import json
import os
import sys
import time
from collections import defaultdict
from dataclasses import dataclass, field
from datetime import datetime, timezone
from pathlib import Path
from typing import Any, Iterable
from urllib.error import HTTPError, URLError
from urllib.parse import quote, urlencode
from urllib.request import Request, urlopen


# -----------------------------------------------------------------------------
# CONFIGURATION
# -----------------------------------------------------------------------------
#TENANT_ID = os.environ.get("GRAPH_TENANT_ID", "YOUR_TENANT_ID_GUID")
#CLIENT_ID = os.environ.get("GRAPH_CLIENT_ID", "YOUR_APPLICATION_CLIENT_ID_GUID")
#CLIENT_SECRET = os.environ.get("GRAPH_CLIENT_SECRET", "YOUR_APPLICATION_CLIENT_SECRET")
#USER_ID = os.environ.get("GRAPH_USER_ID", "YOUR_TARGET_USER_OBJECT_GUID")

START_DATE_TIME = "2006-01-01T00:00:00Z"
END_DATE_TIME = "2036-12-31T23:59:59Z"

SCRIPT_DIRECTORY = Path(__file__).resolve().parent
REPORT_TEXT_PATH = SCRIPT_DIRECTORY / "CalendarMasterDiagnostic.txt"
CREATE_HIGH_EXCEPTION_CSV = True
HIGH_EXCEPTION_THRESHOLD = 20
HIGH_EXCEPTION_CSV_PATH = SCRIPT_DIRECTORY / "CalendarMastersOver20Exceptions.csv"
LONG_MEETING_CSV_PATH = SCRIPT_DIRECTORY / "CalendarMastersOverOneYear.csv"
MASTER_BREAKDOWN_CSV_PATH = SCRIPT_DIRECTORY / "CalendarMasterBreakdown.csv"
DUPLICATE_ICAL_UID_CSV_PATH = SCRIPT_DIRECTORY / "CalendarMastersWithDuplicateIcalUid.csv"
PAGE_SIZE = 500
CALENDAR_VIEW_WINDOW_YEARS = 3
GRAPH_BASE_URI = "https://graph.microsoft.com/v1.0"
# -----------------------------------------------------------------------------


JsonObject = dict[str, Any]


@dataclass
class MasterStats:
    graph_id: str
    ical_uid: str
    master_kind: str
    subject: str
    start: str
    end: str
    calendar_items: int = 0
    occurrences: int = 0
    exceptions: int = 0
    cancelled_items: int = 0
    master_object: JsonObject = field(default_factory=dict)


def log(message: str) -> None:
    print(message, flush=True)


def parse_iso_datetime(value: str, setting_name: str) -> datetime:
    try:
        parsed = datetime.fromisoformat(value.replace("Z", "+00:00"))
    except ValueError as error:
        raise ValueError(f"{setting_name} is not a valid ISO 8601 date/time: {value}") from error
    if parsed.tzinfo is None:
        raise ValueError(f"{setting_name} must include a UTC offset or Z suffix: {value}")
    return parsed


def add_years(value: datetime, years: int) -> datetime:
    try:
        return value.replace(year=value.year + years)
    except ValueError:
        return value.replace(month=2, day=28, year=value.year + years)


def graph_request(
    uri: str,
    *,
    method: str,
    headers: dict[str, str] | None = None,
    body: bytes | None = None,
    maximum_attempts: int = 5,
) -> JsonObject:
    request = Request(uri, data=body, headers=headers or {}, method=method)
    transient_statuses = {429, 500, 502, 503, 504}

    for attempt in range(1, maximum_attempts + 1):
        try:
            with urlopen(request, timeout=120) as response:
                payload = response.read().decode("utf-8")
                result = json.loads(payload)
                if not isinstance(result, dict):
                    raise RuntimeError(f"Expected a JSON object from {uri}")
                return result
        except HTTPError as error:
            if error.code not in transient_statuses or attempt == maximum_attempts:
                detail = error.read().decode("utf-8", errors="replace")
                raise RuntimeError(
                    f"Graph request failed with HTTP {error.code}: {detail}"
                ) from error
            retry_after = error.headers.get("Retry-After")
            delay_seconds = int(retry_after) if retry_after and retry_after.isdigit() else min(2**attempt, 30)
            log(
                f"Warning: Graph returned HTTP {error.code}. Retrying in "
                f"{delay_seconds} seconds (attempt {attempt} of {maximum_attempts})."
            )
            time.sleep(delay_seconds)
        except URLError as error:
            raise RuntimeError(f"Graph request failed: {error.reason}") from error

    raise RuntimeError(f"Graph request failed after {maximum_attempts} attempts: {uri}")


def get_graph_access_token(tenant_id: str, client_id: str, client_secret: str) -> str:
    token_uri = f"https://login.microsoftonline.com/{quote(tenant_id, safe='')}/oauth2/v2.0/token"
    body = urlencode(
        {
            "client_id": client_id,
            "client_secret": client_secret,
            "scope": "https://graph.microsoft.com/.default",
            "grant_type": "client_credentials",
        }
    ).encode("ascii")
    response = graph_request(
        token_uri,
        method="POST",
        headers={"Content-Type": "application/x-www-form-urlencoded"},
        body=body,
    )
    access_token = response.get("access_token")
    if not isinstance(access_token, str) or not access_token:
        raise RuntimeError("The OAuth token response did not contain an access token.")
    return access_token


def invoke_graph_get(uri: str, access_token: str) -> JsonObject:
    return graph_request(
        uri,
        method="GET",
        headers={
            "Authorization": f"Bearer {access_token}",
            "Accept": "application/json",
            "Prefer": f"odata.maxpagesize={PAGE_SIZE}",
        },
    )


def get_all_graph_items(uri: str, access_token: str) -> list[JsonObject]:
    items: list[JsonObject] = []
    next_uri: str | None = uri
    page_number = 0
    while next_uri:
        page_number += 1
        log(f"Reading Graph page {page_number}...")
        response = invoke_graph_get(next_uri, access_token)
        response_items = response.get("value", [])
        if not isinstance(response_items, list):
            raise RuntimeError("Graph response 'value' property was not an array.")
        items.extend(item for item in response_items if isinstance(item, dict))
        next_link = response.get("@odata.nextLink")
        next_uri = next_link if isinstance(next_link, str) and next_link else None
    return items


def event_datetime_text(value: Any) -> str:
    if not isinstance(value, dict) or not value.get("dateTime"):
        return ""
    date_time = str(value["dateTime"])
    time_zone = value.get("timeZone")
    return f"{date_time} [{time_zone}]" if time_zone else date_time


def nested_value(value: Any, *keys: str) -> Any:
    current = value
    for key in keys:
        if not isinstance(current, dict):
            return None
        current = current.get(key)
    return current


def parse_graph_date(value: Any) -> datetime | None:
    if not value:
        return None
    try:
        return datetime.fromisoformat(str(value).replace("Z", "+00:00"))
    except ValueError:
        return None


def master_range_date(master: JsonObject, boundary: str) -> str:
    recurrence_range = nested_value(master, "recurrence", "range")
    if isinstance(recurrence_range, dict):
        if boundary == "End" and recurrence_has_no_end_date(master):
            return ""
        date_value = recurrence_range.get("startDate" if boundary == "Start" else "endDate")
        parsed_date = parse_graph_date(date_value)
        return parsed_date.strftime("%m/%d/%Y") if parsed_date else ""
    key = "start" if boundary == "Start" else "end"
    parsed_date = parse_graph_date(nested_value(master, key, "dateTime"))
    return parsed_date.strftime("%m/%d/%Y") if parsed_date else ""


def recurrence_longer_than_one_year(master: JsonObject) -> bool:
    recurrence_range = nested_value(master, "recurrence", "range")
    if not isinstance(recurrence_range, dict):
        return False
    start_date = parse_graph_date(recurrence_range.get("startDate"))
    end_date = parse_graph_date(recurrence_range.get("endDate"))
    return bool(start_date and end_date and end_date > add_years(start_date, 1))


def meeting_master_longer_than_one_year(master: JsonObject) -> bool:
    if not master:
        return False
    if master.get("recurrence"):
        return recurrence_longer_than_one_year(master)
    start_date = parse_graph_date(nested_value(master, "start", "dateTime"))
    end_date = parse_graph_date(nested_value(master, "end", "dateTime"))
    return bool(start_date and end_date and end_date > add_years(start_date, 1))


def recurrence_has_no_end_date(master: JsonObject) -> bool:
    recurrence_range = nested_value(master, "recurrence", "range")
    if not isinstance(recurrence_range, dict):
        return False
    if str(recurrence_range.get("type") or "").casefold() == "noend":
        return True
    end_date = parse_graph_date(recurrence_range.get("endDate"))
    return end_date is None or end_date.date() == datetime.min.date()


def validate_configuration() -> tuple[datetime, datetime]:
    required_settings = {
        "GRAPH_TENANT_ID": TENANT_ID,
        "GRAPH_CLIENT_ID": CLIENT_ID,
        "GRAPH_CLIENT_SECRET": CLIENT_SECRET,
        "GRAPH_USER_ID": USER_ID,
    }
    for name, value in required_settings.items():
        if not value.strip() or value.startswith("YOUR_"):
            raise ValueError(f"Set {name} in the environment before running this script.")

    parsed_start = parse_iso_datetime(START_DATE_TIME, "START_DATE_TIME")
    parsed_end = parse_iso_datetime(END_DATE_TIME, "END_DATE_TIME")
    if parsed_start >= parsed_end:
        raise ValueError("START_DATE_TIME must be earlier than END_DATE_TIME.")
    if not 1 <= PAGE_SIZE <= 999:
        raise ValueError("PAGE_SIZE must be between 1 and 999.")
    if CALENDAR_VIEW_WINDOW_YEARS < 1:
        raise ValueError("CALENDAR_VIEW_WINDOW_YEARS must be at least 1.")
    return parsed_start, parsed_end


def format_graph_datetime(value: datetime) -> str:
    return value.isoformat(timespec="microseconds")


def read_calendar_items(
    access_token: str, user_id: str, parsed_start: datetime, parsed_end: datetime
) -> list[JsonObject]:
    encoded_user_id = quote(user_id, safe="")
    event_select = "id,subject,type,seriesMasterId,start,end,isCancelled,iCalUId"
    items_by_id: dict[str, JsonObject] = {}
    window_start = parsed_start

    while window_start < parsed_end:
        try:
            window_end = add_years(window_start, CALENDAR_VIEW_WINDOW_YEARS)
        except (OverflowError, ValueError):
            window_end = parsed_end
        window_end = min(window_end, parsed_end)
        start_text = format_graph_datetime(window_start)
        end_text = format_graph_datetime(window_end)
        query = urlencode(
            {
                "startDateTime": start_text,
                "endDateTime": end_text,
                "$select": event_select,
            }
        )
        uri = f"{GRAPH_BASE_URI}/users/{encoded_user_id}/calendarView?{query}"
        log(f"Reading calendar window {start_text} through {end_text}...")
        for item in get_all_graph_items(uri, access_token):
            item_id = item.get("id")
            if isinstance(item_id, str) and item_id:
                items_by_id[item_id] = item
        window_start = window_end

    return list(items_by_id.values())


def build_master_stats(calendar_items: Iterable[JsonObject]) -> dict[str, MasterStats]:
    stats_by_id: dict[str, MasterStats] = {}
    for item in calendar_items:
        recurring_item = item.get("type") in {"occurrence", "exception"}
        series_master_id = item.get("seriesMasterId")
        if recurring_item and isinstance(series_master_id, str) and series_master_id:
            master_id = series_master_id
            master_kind = "Recurring series"
        else:
            master_id = str(item.get("id") or "")
            master_kind = "Single event"

        if master_id not in stats_by_id:
            stats_by_id[master_id] = MasterStats(
                graph_id=master_id,
                ical_uid=str(item.get("iCalUId") or ""),
                master_kind=master_kind,
                subject=str(item.get("subject") or ""),
                start=event_datetime_text(item.get("start")),
                end=event_datetime_text(item.get("end")),
                master_object=item,
            )

        stats = stats_by_id[master_id]
        stats.calendar_items += 1
        if item.get("type") == "occurrence":
            stats.occurrences += 1
        elif item.get("type") == "exception":
            stats.exceptions += 1
        if item.get("isCancelled"):
            stats.cancelled_items += 1
    return stats_by_id


def populate_recurring_masters(
    recurring_stats: Iterable[MasterStats], access_token: str, user_id: str
) -> None:
    encoded_user_id = quote(user_id, safe="")
    master_select = "id,subject,type,start,end,recurrence,iCalUId"
    for stats in recurring_stats:
        encoded_master_id = quote(stats.graph_id, safe="")
        query = urlencode({"$select": master_select})
        uri = f"{GRAPH_BASE_URI}/users/{encoded_user_id}/events/{encoded_master_id}?{query}"
        master = invoke_graph_get(uri, access_token)
        stats.master_object = master
        stats.ical_uid = str(master.get("iCalUId") or "")
        stats.subject = str(master.get("subject") or "")
        stats.start = event_datetime_text(master.get("start"))
        stats.end = event_datetime_text(master.get("end"))


def duplicate_maps(
    master_stats: list[MasterStats],
) -> tuple[list[tuple[str, list[MasterStats]]], dict[str, int], dict[str, list[MasterStats]]]:
    grouped: defaultdict[str, list[MasterStats]] = defaultdict(list)
    for stats in master_stats:
        if stats.ical_uid.strip():
            grouped[stats.ical_uid].append(stats)
    duplicate_groups = [(uid, group) for uid, group in grouped.items() if len(group) > 1]
    counts: dict[str, int] = {}
    peers: dict[str, list[MasterStats]] = {}
    for _, group in duplicate_groups:
        for stats in group:
            counts[stats.graph_id] = len(group)
            peers[stats.graph_id] = [peer for peer in group if peer.graph_id != stats.graph_id]
    return duplicate_groups, counts, peers


def warning_messages(stats: MasterStats, duplicate_counts: dict[str, int]) -> list[str]:
    warnings: list[str] = []
    if stats.exceptions > HIGH_EXCEPTION_THRESHOLD:
        warnings.append("More than 20 exceptions for this master event.")
    if meeting_master_longer_than_one_year(stats.master_object):
        warnings.append("The meeting is over a year long.")
    if recurrence_has_no_end_date(stats.master_object):
        warnings.append("The meeting is over has no end date.")
    if stats.graph_id in duplicate_counts:
        warnings.append(
            f"Duplicate iCalUID shared by {duplicate_counts[stats.graph_id]} master records."
        )
    return warnings


def write_csv(path: Path, fieldnames: list[str], rows: Iterable[dict[str, Any]]) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    row_list = list(rows)
    with path.open("w", encoding="utf-8-sig", newline="") as output_file:
        if row_list:
            writer = csv.DictWriter(output_file, fieldnames=fieldnames, quoting=csv.QUOTE_ALL)
            writer.writeheader()
            writer.writerows(row_list)
        else:
            output_file.write(",".join(fieldnames) + "\n")


def write_reports(
    calendar_items: list[JsonObject],
    master_stats: list[MasterStats],
    recurring_stats: list[MasterStats],
    parsed_start: datetime,
    parsed_end: datetime,
) -> None:
    duplicate_groups, duplicate_counts, duplicate_peers = duplicate_maps(master_stats)
    high_exception_stats = [
        stats for stats in recurring_stats if stats.exceptions > HIGH_EXCEPTION_THRESHOLD
    ]
    longer_than_one_year_stats = [
        stats for stats in recurring_stats if recurrence_longer_than_one_year(stats.master_object)
    ]
    without_end_date_stats = [
        stats for stats in recurring_stats if recurrence_has_no_end_date(stats.master_object)
    ]
    long_meeting_stats = [
        stats
        for stats in master_stats
        if meeting_master_longer_than_one_year(stats.master_object)
        or recurrence_has_no_end_date(stats.master_object)
    ]

    total_occurrences = sum(item.get("type") == "occurrence" for item in calendar_items)
    total_exceptions = sum(item.get("type") == "exception" for item in calendar_items)
    total_single_events = sum(item.get("type") == "singleInstance" for item in calendar_items)
    total_cancelled_items = sum(bool(item.get("isCancelled")) for item in calendar_items)
    highest_recurrences = max((stats.calendar_items for stats in recurring_stats), default=0)
    highest_exceptions = max((stats.exceptions for stats in recurring_stats), default=0)

    report_lines = [
        "MICROSOFT GRAPH CALENDAR MASTER DIAGNOSTIC",
        f"Generated (UTC): {datetime.now(timezone.utc).strftime('%Y-%m-%dT%H:%M:%SZ')}",
        f"User GUID: {USER_ID}",
        f"Period: {format_graph_datetime(parsed_start)} through {format_graph_datetime(parsed_end)}",
        "",
        "TOTALS",
        f"Master records with events in period: {len(master_stats)}",
        f"  Recurring series masters: {len(recurring_stats)}",
        f"  Recurring meeting masters with more than {HIGH_EXCEPTION_THRESHOLD} exceptions: {len(high_exception_stats)}",
        f"  Recurring meeting masters longer than one year: {len(longer_than_one_year_stats)}",
        f"  Recurring meeting masters without an end date: {len(without_end_date_stats)}",
        f"  Single-event masters: {len(master_stats) - len(recurring_stats)}",
        f"  Duplicate iCalUID values: {len(duplicate_groups)}",
        f"  Master records sharing duplicate iCalUIDs: {len(duplicate_counts)}",
        f"Events in period, including recurring instances: {len(calendar_items)}",
        f"  Single-instance events: {total_single_events}",
        f"  Normal recurring occurrences: {total_occurrences}",
        f"  Modified recurring exceptions: {total_exceptions}",
        f"  Cancelled items returned by Graph: {total_cancelled_items}",
        f"Highest recurring instances for one meeting master in period: {highest_recurrences}",
        f"Highest modified exception count for one recurring meeting master in period: {highest_exceptions}",
        "",
        "MASTER RECORD BREAKDOWN",
    ]
    if not master_stats:
        report_lines.append("No calendar records were returned for the requested period.")
    else:
        for stats in master_stats:
            report_lines.extend(
                [
                    "-" * 100,
                    f"Graph ID: {stats.graph_id}",
                    f"iCalUID: {stats.ical_uid}",
                    f"Subject: {stats.subject}",
                    f"Master type: {stats.master_kind}",
                    f"Master/first event start: {stats.start}",
                    f"Master/first event end: {stats.end}",
                    f"Events in period: {stats.calendar_items}",
                    f"Normal recurring occurrences in period: {stats.occurrences}",
                    f"Modified recurring exceptions in period: {stats.exceptions}",
                    f"Cancelled items returned by Graph in period: {stats.cancelled_items}",
                ]
            )
            for warning in warning_messages(stats, duplicate_counts):
                report_lines.append(f"*** Warning: {warning} ***")
            for peer in duplicate_peers.get(stats.graph_id, []):
                report_lines.append(f"    Duplicate Graph ID: {peer.graph_id}")
                report_lines.append(f"    Duplicate Subject: {peer.subject}")

    REPORT_TEXT_PATH.parent.mkdir(parents=True, exist_ok=True)
    REPORT_TEXT_PATH.write_text("\n".join(report_lines) + "\n", encoding="utf-8-sig")

    breakdown_fields = [
        "GraphId", "IcalUid", "Subject", "MasterType", "MasterFirstEventStart",
        "MasterFirstEventEnd", "EventsInPeriod", "NormalRecurringOccurrencesInPeriod",
        "ModifiedRecurringExceptionsInPeriod", "CancelledItemsReturnedByGraphInPeriod", "Warnings",
    ]
    write_csv(
        MASTER_BREAKDOWN_CSV_PATH,
        breakdown_fields,
        (
            {
                "GraphId": stats.graph_id,
                "IcalUid": stats.ical_uid,
                "Subject": stats.subject,
                "MasterType": stats.master_kind,
                "MasterFirstEventStart": stats.start,
                "MasterFirstEventEnd": stats.end,
                "EventsInPeriod": stats.calendar_items,
                "NormalRecurringOccurrencesInPeriod": stats.occurrences,
                "ModifiedRecurringExceptionsInPeriod": stats.exceptions,
                "CancelledItemsReturnedByGraphInPeriod": stats.cancelled_items,
                "Warnings": " | ".join(warning_messages(stats, duplicate_counts)),
            }
            for stats in master_stats
        ),
    )

    duplicate_fields = [
        "IcalUid", "DuplicateMasterCount", "GraphId", "Subject", "DuplicateGraphIds",
        "DuplicateSubjects", "MasterType", "MasterFirstEventStart", "MasterFirstEventEnd",
    ]
    duplicate_rows: list[dict[str, Any]] = []
    for ical_uid, group in duplicate_groups:
        for stats in sorted(group, key=lambda value: (value.subject, value.graph_id)):
            peers = duplicate_peers[stats.graph_id]
            duplicate_rows.append(
                {
                    "IcalUid": ical_uid,
                    "DuplicateMasterCount": len(group),
                    "GraphId": stats.graph_id,
                    "Subject": stats.subject,
                    "DuplicateGraphIds": " | ".join(peer.graph_id for peer in peers),
                    "DuplicateSubjects": " | ".join(peer.subject for peer in peers),
                    "MasterType": stats.master_kind,
                    "MasterFirstEventStart": stats.start,
                    "MasterFirstEventEnd": stats.end,
                }
            )
    write_csv(DUPLICATE_ICAL_UID_CSV_PATH, duplicate_fields, duplicate_rows)

    summary_fields = ["GraphId", "Subject", "StartDate", "EndDate"]
    if CREATE_HIGH_EXCEPTION_CSV:
        write_csv(
            HIGH_EXCEPTION_CSV_PATH,
            summary_fields,
            (
                {
                    "GraphId": stats.graph_id,
                    "Subject": stats.subject,
                    "StartDate": master_range_date(stats.master_object, "Start"),
                    "EndDate": master_range_date(stats.master_object, "End"),
                }
                for stats in sorted(high_exception_stats, key=lambda value: value.exceptions, reverse=True)
            ),
        )
    write_csv(
        LONG_MEETING_CSV_PATH,
        summary_fields,
        (
            {
                "GraphId": stats.graph_id,
                "Subject": stats.subject,
                "StartDate": master_range_date(stats.master_object, "Start"),
                "EndDate": master_range_date(stats.master_object, "End"),
            }
            for stats in sorted(long_meeting_stats, key=lambda value: (value.subject, value.graph_id))
        ),
    )


def main() -> int:
    try:
        parsed_start, parsed_end = validate_configuration()
        for output_path in (
            REPORT_TEXT_PATH,
            HIGH_EXCEPTION_CSV_PATH,
            LONG_MEETING_CSV_PATH,
            MASTER_BREAKDOWN_CSV_PATH,
            DUPLICATE_ICAL_UID_CSV_PATH,
        ):
            output_path.parent.mkdir(parents=True, exist_ok=True)

        log("Acquiring application OAuth token...")
        access_token = get_graph_access_token(TENANT_ID, CLIENT_ID, CLIENT_SECRET)
        log(
            f"Reading expanded calendar events from {format_graph_datetime(parsed_start)} "
            f"through {format_graph_datetime(parsed_end)}..."
        )
        calendar_items = read_calendar_items(access_token, USER_ID, parsed_start, parsed_end)
        stats_by_id = build_master_stats(calendar_items)
        recurring_stats = [
            stats for stats in stats_by_id.values() if stats.master_kind == "Recurring series"
        ]
        populate_recurring_masters(recurring_stats, access_token, USER_ID)
        master_stats = sorted(stats_by_id.values(), key=lambda value: (value.subject, value.graph_id))
        write_reports(calendar_items, master_stats, recurring_stats, parsed_start, parsed_end)

        log(f"Text report written to: {REPORT_TEXT_PATH}")
        log(f"Master breakdown CSV written to: {MASTER_BREAKDOWN_CSV_PATH}")
        log(f"Duplicate iCalUID CSV written to: {DUPLICATE_ICAL_UID_CSV_PATH}")
        if CREATE_HIGH_EXCEPTION_CSV:
            log(f"High-exception CSV written to: {HIGH_EXCEPTION_CSV_PATH}")
        log(f"Over-one-year CSV written to: {LONG_MEETING_CSV_PATH}")
        return 0
    except (OSError, RuntimeError, ValueError) as error:
        print(f"Error: {error}", file=sys.stderr)
        return 1


if __name__ == "__main__":
    raise SystemExit(main())
