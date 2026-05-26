# sellercloud-sync

Daily ETL that loads the SellerCloud catalog export into SQL Server. Reads `SellerCloud.xlsx` from the configured OneDrive path, normalizes column dtypes, clears the `Reports.SellerCloud` table, and bulk-inserts every row via `seller_automation_utils.database_utils.insert_dataframe`. Runs at 05:15 daily via APScheduler — ~15 minutes after the SellerCloud daily report typically lands in OneDrive.

This script was extracted from `amzn-catalog-health` in May 2026 so the daily catalog refresh runs independently of the larger nightly catalog/health scrape job. The `Reports.SellerCloud` table feeds many downstream reports across the suite; decoupling its refresh removes a hidden dependency on a long-running job that occasionally crashes early.

## Daily flow

1. **Read export** — read `SellerCloud.xlsx` from the configured OneDrive path.
2. **Normalize** — normalize column dtypes for the SQL schema.
3. **Reload table** — clear the `Reports.SellerCloud` table and bulk-insert every row via `seller_automation_utils.database_utils.insert_dataframe`.

## Architecture

```mermaid
flowchart LR
    sched[APScheduler<br/>daily 05:15] --> fresh{File dated<br/>today?}
    fresh -->|no| skip[Skip — keep<br/>yesterday's data]
    fresh -->|yes| read[Read SellerCloud.xlsx<br/>dtype-typed]
    read --> norm[Normalize text / int / float columns]
    norm --> reload[DELETE all rows<br/>+ bulk insert]
    reload --> db[(SQL Server<br/>Reports.SellerCloud)]
```

## Performance notes

No browser, no Excel automation — just a fast, defensive file-to-SQL load:

- **Freshness guard.** Before touching the database the script checks the
  latest file's modification date via `latest_modified_date`; if it isn't
  today's (or the folder is empty) it logs and exits cleanly — it would
  rather miss a day than overwrite `Reports.SellerCloud` with stale data that
  many downstream reports depend on.
- **Explicit dtype normalization.** Text columns are forced to `str` (so
  leading-zero IDs like UPC / SKU / eBayItemID survive), int and float columns
  are coerced with `pd.to_numeric(errors="coerce").fillna(0)`.
- **Injection-safe table name.** `DB_TABLE_SELLERCLOUD` is validated
  (`isalnum` after stripping underscores) before use.
- **Full-replace load.** A single `DELETE` then a bulk `insert_dataframe`,
  with columns renamed to the SQL schema (`Manufacturer` → `BrandName`,
  `MFNQuantity` → `AmazonQuantity`).
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
├── logs/                       # rotating run logs (gitignored)
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
.venv\Scripts\pip install git+https://github.com/dominicci13/shared-python-utils.git
```

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

## Environment variables

| Variable | Description |
|---|---|
| `DB_TABLE_SELLERCLOUD` | SQL Server table name (default: `SellerCloud`) |
| `ALERT_EMAIL` | Outlook account used to send crash reports via `seller_automation_utils.alert_utils.handle_crash` |

## Author

Built by **Brian Ramirez** ([@dominicci13](https://github.com/dominicci13)) — automation & AI workflow specialist. More on my [GitHub profile](https://github.com/dominicci13) and [LinkedIn](https://linkedin.com/in/bdramirez).

## License

[MIT](LICENSE)
