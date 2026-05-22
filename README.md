# sellercloud-sync

Daily ETL that loads the SellerCloud catalog export into SQL Server. Reads `SellerCloud.xlsx` from the configured OneDrive path, normalizes column dtypes, clears the `Reports.SellerCloud` table, and bulk-inserts every row via `fc_utils.database_utils.insert_dataframe`. Runs at 05:15 daily via APScheduler — ~15 minutes after the SellerCloud daily report typically lands in OneDrive.

This script was extracted from `amzn-catalog-health` in May 2026 so the daily catalog refresh runs independently of the larger nightly catalog/health scrape job. The `Reports.SellerCloud` table feeds many downstream reports across the suite; decoupling its refresh removes a hidden dependency on a long-running job that occasionally crashes early.

## Setup

### 1. Install dependencies

```bash
pip install -r requirements.txt
pip install git+https://github.com/dominicci13/shared-python-utils.git
```

### 2. Configure environment

```bash
cp .env.example .env
cp config/paths.json.example config/paths.json
```

Edit both files with your local paths and SQL table name.

## Run

```bash
python run_sellercloud_sync.py
```

Prompts whether to run immediately, then schedules itself to run at 05:15 daily via APScheduler.

## Environment Variables

| Variable | Description |
|---|---|
| `DB_TABLE_SELLERCLOUD` | SQL Server table name (default: `SellerCloud`) |
| `ALERT_EMAIL` | Outlook account used to send crash reports via `fc_utils.alert_utils.handle_crash` |

## License

[MIT](LICENSE)
