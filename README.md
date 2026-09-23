# sellercloud-sync

Daily ETL that loads the SellerCloud catalog export into SQL Server. Reads `SellerCloud.xlsx` from the configured OneDrive path, checks its headers against the known schema, normalizes column dtypes, clears the `Reports.SellerCloud` table, and bulk-inserts every row via `seller_automation_utils.database_utils.insert_dataframe`. Runs at 05:15 daily via APScheduler — ~15 minutes after the SellerCloud daily report typically lands in OneDrive.

This script was extracted from `amzn-catalog-health` in May 2026 so the daily catalog refresh runs independently of the larger nightly catalog/health scrape job. The `Reports.SellerCloud` table feeds many downstream reports across the suite; decoupling its refresh removes a hidden dependency on a long-running job that occasionally crashes early.

## Daily flow

1. **Freshness check** — stat `SellerCloud.xlsx` itself; if it is missing or not dated today, leave the table untouched and alert. See [Stale-export alerts](#stale-export-alerts).
2. **Read export** — read `SellerCloud.xlsx` from the configured OneDrive path.
3. **Schema check** — compare headers against `COLUMN_TYPES`; on drift, email a one-time alert and (for removed columns) skip the write. See [Schema-drift alerts](#schema-drift-alerts).
4. **Empty check** — a fresh file with headers but no data rows skips the write and alerts once per file date. See [Stale-export alerts](#stale-export-alerts).
5. **Normalize** — normalize column dtypes for the SQL schema.
6. **Reload table** — clear the `Reports.SellerCloud` table and bulk-insert every row via `seller_automation_utils.database_utils.insert_dataframe`.

## Architecture

```mermaid
flowchart LR
    sched[APScheduler<br/>daily 05:15] --> fresh{SellerCloud.xlsx<br/>itself dated today?}
    fresh -->|no / missing| skip[WARNING + email once<br/>per file date —<br/>table untouched]
    fresh -->|yes| read[Read SellerCloud.xlsx<br/>dtype-typed]
    read --> drift{Headers match<br/>COLUMN_TYPES?}
    drift -->|new column| alert[Email one-time<br/>alert] --> norm
    drift -->|removed column| halt[Email alert +<br/>skip write — keep<br/>yesterday's data]
    drift -->|yes| empty{Any data<br/>rows?}
    empty -->|no| skip2[WARNING + email once<br/>per file date —<br/>table untouched]
    empty -->|yes| norm[Normalize text / int / float columns]
    norm --> reload[DELETE all rows<br/>+ bulk insert]
    reload --> db[(SQL Server<br/>Reports.SellerCloud)]
```

## Stale-export alerts

`SellerCloud.xlsx` is produced by a SellerCloud scheduled report, not by this
repo, so the script cannot fix a missing delivery, only refuse to load it and
say so. Before opening a DB connection, `export_is_fresh` stats the file
**itself** and compares its local modification date with today:

- **Fresh** — the sync proceeds.
- **Stale or missing** — no connection is opened, so `Reports.SellerCloud` keeps
  its current rows. The run logs a `WARNING` and emails `ALERT_EMAIL` with the
  file path, its LastWriteTime, how many days stale it is, and a note that the
  table was left untouched. Values in the email are HTML-escaped.

Why the file and not the folder: the folder also holds a workbook another
automation rewrites every day, so a folder-wide "newest file" check always read
as today and let a weeks-old export be reloaded daily.

The email is debounced via `logs/freshness_alert_state.json`: one email per
stale file date (or one for a missing file), while the `WARNING` repeats every
run. A fresh run clears the state, so the next stale date alerts again. The state
is written only after Outlook accepts the email; a failed send (or a missing
`ALERT_EMAIL`) goes to the crash handler instead and is retried next run. An
unreadable state file (bad JSON, non-UTF-8 bytes, a directory, no permission)
logs a `WARNING` and reads as empty, so it errs toward alerting; a failure to
clear it on a fresh run is logged and never blocks the load.

**Empty export.** A fresh-dated file with headers but no data rows would
otherwise `DELETE` the table and insert nothing. After the schema check and
before the `DELETE`, the run skips the write, logs a `WARNING`, and emails
`ALERT_EMAIL` once per file date (debounced via
`logs/empty_export_alert_state.json`).

## Schema-drift alerts

The SellerCloud export occasionally gains or loses columns. Rather than silently
dropping new data or crashing on a vanished column, the script diffs the report's
headers against the `COLUMN_TYPES` dict on every run:

- **New column** (in the report, not in `COLUMN_TYPES`) — harmless to the load
  (the insert writes an explicit column list), so the sync continues and emails
  an alert. The email samples the column's values, infers a likely type, and
  includes ready-to-paste `COLUMN_TYPES` and `ALTER TABLE` lines plus numbered
  steps for wiring it into both the script and the SQL table.
- **Removed column** (in `COLUMN_TYPES`, not in the report) — would break the
  insert, so the script emails an alert and **skips the DB write**, leaving
  yesterday's data intact. The email gives rename/remove steps.

The read happens *before* the `DELETE`, so a drifted schema (or an unreadable
file) never clears the table and then fails to refill it.

To avoid a daily nag, alerts are deduped via `logs/schema_alert_state.json`: each
column is emailed once, and entries clear automatically once the drift is
resolved. Alerts go to `ALERT_EMAIL` (the same address as crash reports).

## Performance notes

No browser, no Excel automation — just a fast, defensive file-to-SQL load:

- **Freshness guard.** Before touching the database the script checks
  `SellerCloud.xlsx`'s own modification date; if it isn't today's (or the file
  is missing) it alerts and skips the load. It would rather miss a day than
  overwrite `Reports.SellerCloud` with stale data that many downstream reports
  depend on, and it never misses one silently. See
  [Stale-export alerts](#stale-export-alerts).
- **One-stop column schema.** A single ordered `COLUMN_TYPES` dict maps every
  Excel column to how it's coerced and drives reading, normalization, and
  insertion. Adding a column the export starts emitting is a one-line edit.
- **Explicit dtype normalization.** Text columns are read as `str` (so
  leading-zero and long numeric IDs like UPC / SKU / eBayItemID survive without
  a `.0`); `int`/`float` columns coerce blanks to `0`; `float_null` analytical
  columns (P&L, shipping cost, FBAFee, Rebate, TotalCost) and the `datetime`
  `LastReceived` keep blanks as SQL `NULL`. Names with spaces/symbols like
  `P&L (30 days)` are bracketed for SQL automatically by the shared
  `insert_dataframe` (seller-automation-utils 1.8.3 or later is required).
- **Injection-safe table name.** `DB_TABLE_SELLERCLOUD` is validated
  (`isalnum` after stripping underscores) before use.
- **Full-replace load.** A single `DELETE` then a bulk `insert_dataframe` of all
  46 columns in their SQL-schema order.
- **Decoupled by design.** Extracted from `amzn-catalog-health` so this daily
  refresh can't be blocked by the long-running nightly scrape.

## Logging

```text
05:15:02 INFO     Uploading SellerCloud items to database.
05:15:03 INFO     Table rows deleted successfully.
05:15:48 INFO     Data inserted successfully.
05:15:48 SUCCESS  Done!
```

Configured once via the shared helper:

```python
from seller_automation_utils.logging_utils import setup_logging
log = setup_logging("sellercloud_sync")
```

`setup_logging` wires a Rich console handler (colorized output, markup
rendering, rich tracebacks) and a 1 MB rotating file handler writing to
`logs/sellercloud_sync.log`. Available to every automation that imports
`seller_automation_utils`.

## Project layout

```
sellercloud-sync/
├── run_sellercloud_sync.py     # entry point (single script)
├── config/
│   └── paths.json.example      # OneDrive folder holding SellerCloud.xlsx
├── logs/                       # rotating run logs + schema/freshness alert state (gitignored)
├── tests/                      # pytest: freshness guard (email and SQL mocked)
├── pytest.ini
├── .env.example
├── requirements.txt
├── LICENSE
└── README.md
```

## Setup

### 1. Clone and create the venv

```powershell
git clone https://github.com/dominicci13/sellercloud-sync.git
cd sellercloud-sync
py -3.12 -m venv .venv
.venv\Scripts\pip install -r requirements.txt
```

`requirements.txt` pins seller-automation-utils to a commit (1.8.5, `f8defe9`); bump the pin deliberately, never install the library unpinned.

### 2. Configure

```powershell
copy .env.example .env
copy config\paths.json.example config\paths.json
```

Edit both files with real values. Both are gitignored.

### 3. Run

```powershell
.venv\Scripts\python run_sellercloud_sync.py
```

The script prompts "Run now?" — answer **Y** to execute immediately, or **N** to register the APScheduler job and idle until the next **daily 05:15** trigger.

### 4. Test

```powershell
.venv\Scripts\python.exe -m pytest -q -p no:cacheprovider
```

Email, SQL, the crash handler and logging are mocked in `tests\conftest.py`, and every file the tests stat is a temp file, so nothing real is sent or touched. Importing the script still needs a local `config\paths.json`. `pytest.ini` disables the `pytest-html` and `seleniumbase` plugins, which this suite does not use.

## Environment variables

| Variable | Description |
|---|---|
| `DB_TABLE_SELLERCLOUD` | SQL Server table name (default: `SellerCloud`) |
| `ALERT_EMAIL` | Outlook account used to send crash reports (`seller_automation_utils.alert_utils.handle_crash`), schema-drift alerts and stale- or empty-export alerts (required once a stale, missing or empty export is seen; its absence crashes the run loudly) |

## Author

Built by **Brian Ramirez** ([@dominicci13](https://github.com/dominicci13)) — automation & AI workflow specialist. More on my [GitHub profile](https://github.com/dominicci13) and [LinkedIn](https://linkedin.com/in/bdramirez).

## License

[MIT](LICENSE)
