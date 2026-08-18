# Python-Graph-Calendar-Diagnostic.py

A diagnostic tool that uses Microsoft Graph application OAuth (client
credentials) to read a user's default calendar. It queries the
`calendarView` endpoint, which expands recurring series into individual
occurrences and exceptions for the requested period. Events are then grouped
by their series master ID (`seriesMasterId`); non-recurring events are
treated as their own master record.

The tool produces:

* A text report with per-master details and overall totals.
* A CSV report listing recurring masters with more than 20 modified
  exceptions.
* A CSV report listing masters whose series spans more than one year.

## Setup

```bash
pip install -r requirements.txt
```

You'll need an Azure AD app registration with application permissions for
`Calendars.Read` (admin consent granted), along with its tenant ID, client
ID, and client secret. Prefer setting the client secret via the
`GRAPH_CLIENT_SECRET` environment variable rather than `--client-secret`, to
avoid exposing it in shell history or process listings.

## Usage

```bash
export GRAPH_CLIENT_SECRET=<CLIENT_SECRET>
python graph_calendar_diagnostic.py \
  --tenant-id <TENANT_ID> \
  --client-id <CLIENT_ID> \
  --user <user@example.com> \
  --start 2024-01-01T00:00:00 \
  --end 2024-12-31T23:59:59
```

If `--start`/`--end` are omitted, the tool defaults to a window starting now
and extending `--days` (default 30) days into the future.

Output paths can be customized with `--report`, `--exceptions-csv`, and
`--long-series-csv` (see `--help` for defaults).

## Tests

```bash
pip install pytest
pytest tests/
```
