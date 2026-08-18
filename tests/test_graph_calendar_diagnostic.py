import csv
import os
import sys

sys.path.insert(0, os.path.abspath(os.path.join(os.path.dirname(__file__), "..")))

from graph_calendar_diagnostic import (
    DURATION_THRESHOLD_DAYS,
    EXCEPTION_THRESHOLD,
    group_events,
    parse_graph_datetime,
    write_exception_report,
    write_long_series_report,
    write_text_report,
)


def make_event(event_id, subject, start, end, event_type="singleInstance", series_master_id=None, cancelled=False):
    return {
        "id": event_id,
        "subject": subject,
        "type": event_type,
        "seriesMasterId": series_master_id,
        "isCancelled": cancelled,
        "start": {"dateTime": start, "timeZone": "UTC"},
        "end": {"dateTime": end, "timeZone": "UTC"},
    }


def test_parse_graph_datetime():
    dt = parse_graph_datetime("2024-01-15T09:00:00.0000000")
    assert dt is not None
    assert dt.year == 2024 and dt.month == 1 and dt.day == 15

    assert parse_graph_datetime(None) is None
    assert parse_graph_datetime("") is None


def test_non_recurring_event_is_its_own_master():
    events = [
        make_event("evt1", "One-off meeting", "2024-01-15T09:00:00.0000000", "2024-01-15T10:00:00.0000000"),
    ]
    masters = group_events(events)

    assert len(masters) == 1
    master = masters["evt1"]
    assert master.is_recurring is False
    assert master.occurrence_count == 1
    assert master.exception_count == 0
    assert master.subject == "One-off meeting"


def test_recurring_series_groups_by_series_master_id():
    events = []
    for day in range(1, 6):
        events.append(
            make_event(
                f"occ{day}",
                "Weekly sync",
                f"2024-01-{day:02d}T09:00:00.0000000",
                f"2024-01-{day:02d}T10:00:00.0000000",
                event_type="occurrence",
                series_master_id="series1",
            )
        )
    events.append(
        make_event(
            "exc1",
            "Weekly sync (moved)",
            "2024-01-06T11:00:00.0000000",
            "2024-01-06T12:00:00.0000000",
            event_type="exception",
            series_master_id="series1",
        )
    )

    masters = group_events(events)

    assert len(masters) == 1
    master = masters["series1"]
    assert master.is_recurring is True
    assert master.occurrence_count == 5
    assert master.exception_count == 1
    assert master.total_events == 6


def test_exception_threshold_flagging():
    events = [
        make_event(
            f"exc{i}",
            "Overridden series",
            "2024-01-01T09:00:00.0000000",
            "2024-01-01T10:00:00.0000000",
            event_type="exception",
            series_master_id="series-exceptions",
        )
        for i in range(EXCEPTION_THRESHOLD + 5)
    ]
    masters = group_events(events)
    master = masters["series-exceptions"]

    assert master.exception_count == EXCEPTION_THRESHOLD + 5
    assert master.exceeds_exception_threshold is True


def test_duration_threshold_flagging():
    events = [
        make_event(
            "occ-start",
            "Long series",
            "2023-01-01T09:00:00.0000000",
            "2023-01-01T10:00:00.0000000",
            event_type="occurrence",
            series_master_id="series-long",
        ),
        make_event(
            "occ-end",
            "Long series",
            "2024-06-01T09:00:00.0000000",
            "2024-06-01T10:00:00.0000000",
            event_type="occurrence",
            series_master_id="series-long",
        ),
    ]
    masters = group_events(events)
    master = masters["series-long"]

    assert master.duration_days > DURATION_THRESHOLD_DAYS
    assert master.exceeds_duration_threshold is True


def test_short_recurring_series_not_flagged():
    events = [
        make_event(
            f"occ{day}",
            "Weekly sync",
            f"2024-01-{day:02d}T09:00:00.0000000",
            f"2024-01-{day:02d}T10:00:00.0000000",
            event_type="occurrence",
            series_master_id="series-short",
        )
        for day in range(1, 6)
    ]
    masters = group_events(events)
    master = masters["series-short"]

    assert master.exceeds_exception_threshold is False
    assert master.exceeds_duration_threshold is False


def test_write_text_report_contains_totals(tmp_path):
    events = [
        make_event("evt1", "One-off", "2024-01-15T09:00:00.0000000", "2024-01-15T10:00:00.0000000"),
        make_event(
            "occ1",
            "Recurring",
            "2024-02-01T09:00:00.0000000",
            "2024-02-01T10:00:00.0000000",
            event_type="occurrence",
            series_master_id="series1",
        ),
    ]
    masters = group_events(events)
    report_path = tmp_path / "report.txt"
    write_text_report(masters.values(), str(report_path))

    content = report_path.read_text()
    assert "Master ID: evt1" in content
    assert "Master ID: series1" in content
    assert "Total masters: 2" in content
    assert "Recurring masters: 1" in content
    assert "Non-recurring masters: 1" in content


def test_write_exception_report_filters_by_threshold(tmp_path):
    flagged_events = [
        make_event(
            f"exc{i}",
            "Overridden series",
            "2024-01-01T09:00:00.0000000",
            "2024-01-01T10:00:00.0000000",
            event_type="exception",
            series_master_id="series-flagged",
        )
        for i in range(EXCEPTION_THRESHOLD + 1)
    ]
    normal_events = [
        make_event(
            "occ1",
            "Normal series",
            "2024-01-01T09:00:00.0000000",
            "2024-01-01T10:00:00.0000000",
            event_type="occurrence",
            series_master_id="series-normal",
        )
    ]
    masters = group_events(flagged_events + normal_events)

    csv_path = tmp_path / "exceptions.csv"
    write_exception_report(masters.values(), str(csv_path))

    with open(csv_path, newline="") as fh:
        rows = list(csv.DictReader(fh))

    assert len(rows) == 1
    assert rows[0]["master_id"] == "series-flagged"


def test_write_long_series_report_filters_by_duration(tmp_path):
    events = [
        make_event(
            "occ-start",
            "Long series",
            "2023-01-01T09:00:00.0000000",
            "2023-01-01T10:00:00.0000000",
            event_type="occurrence",
            series_master_id="series-long",
        ),
        make_event(
            "occ-end",
            "Long series",
            "2024-06-01T09:00:00.0000000",
            "2024-06-01T10:00:00.0000000",
            event_type="occurrence",
            series_master_id="series-long",
        ),
        make_event(
            "occ-short",
            "Short series",
            "2024-01-01T09:00:00.0000000",
            "2024-01-01T10:00:00.0000000",
            event_type="occurrence",
            series_master_id="series-short",
        ),
    ]
    masters = group_events(events)

    csv_path = tmp_path / "long_series.csv"
    write_long_series_report(masters.values(), str(csv_path))

    with open(csv_path, newline="") as fh:
        rows = list(csv.DictReader(fh))

    assert len(rows) == 1
    assert rows[0]["master_id"] == "series-long"
