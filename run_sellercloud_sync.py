"""Daily SellerCloud catalog sync.

Once per day at 05:15 local time (~15 minutes after the SellerCloud daily report
typically lands in OneDrive), this script:

1. Reads the most recent ``SellerCloud.xlsx`` from the configured OneDrive path.
2. Confirms its modification date is today (otherwise skips — yesterday's data
   is already in the table, and we'd rather miss a day than overwrite with
   stale data).
3. Clears the ``Reports.SellerCloud`` table.
4. Bulk-inserts every row via ``seller_automation_utils.database_utils.insert_dataframe``.

Extracted from ``amzn-catalog-health`` in May 2026 so this daily refresh runs
independently of the larger nightly catalog/health scrape job at 04:00.
"""
from __future__ import annotations

import os
import traceback
from datetime import datetime
from pathlib import Path

import pandas as pd
from dotenv import load_dotenv

from seller_automation_utils import alert_utils, custom_functions, database_utils
from seller_automation_utils.config_utils import load_config_safe
from seller_automation_utils.file_utils import latest_modified_date
from seller_automation_utils.logging_utils import setup_logging
from seller_automation_utils.schedule_utils import run_on_schedule
from seller_automation_utils.ui_utils import ask_user


log = setup_logging("sellercloud_sync")

load_dotenv()
table_sellercloud: str = os.getenv("DB_TABLE_SELLERCLOUD", "SellerCloud")
if not table_sellercloud.replace("_", "").isalnum():
    raise ValueError(f"Invalid table name: {table_sellercloud!r}")

# Resolve config relative to this file so the script works regardless of CWD
# (Task Scheduler / cron wrappers etc.)
_paths = load_config_safe(Path(__file__).resolve().parent / "config" / "paths.json")
sellercloud_file_path: str = _paths["sellercloud_file_path"]


# --- Column schema --------------------------------------------------------
# The single place to edit when the SellerCloud.xlsx export gains/loses a
# column. Keys are the *exact* Excel header (which also matches the SQL column
# name); the value picks how the cell is coerced before insert:
#
#   "text"       -> string; blanks become "" (read as str so leading zeros and
#                   long numeric IDs like eBayItemID survive without a ".0").
#   "int"        -> whole number; blanks/garbage become 0.
#   "float"      -> decimal; blanks/garbage become 0.
#   "float_null" -> decimal; blanks stay NULL (use for analytical metrics where
#                   "no data" must stay distinct from a real 0).
#   "datetime"   -> timestamp; blanks/unparseable stay NULL.
#
# Order here is the SQL table's column order. To add a new column, drop one
# line in the right spot with its type — it flows through read, normalize,
# and insert automatically. Names with spaces/symbols are bracketed for SQL
# on the fly, so write them plainly (e.g. "P&L (30 days)").
COLUMN_TYPES: dict[str, str] = {
    "SKU": "text",
    "ASIN": "text",
    "CompanyID": "int",
    "CompanyName": "text",
    "WalmartAPIItemID": "text",
    "eBayItemID": "text",
    "ProductName": "text",
    "Manufacturer": "text",
    "Vendor": "text",
    "BuyerEmail": "text",
    "UPC": "text",
    "AggregateQty": "int",
    "MFNQuantity": "int",
    "FBAQuantity": "int",
    "AmazonPrice": "float",
    "AmazonBusinessPrice": "float",
    "AmazonBusinessPriceDiscountQty1": "int",
    "Rebate": "float_null",
    "ListPrice": "float",
    "SitePrice": "float",
    "MAPPrice": "float",
    "SiteCost": "float",
    "TotalCost": "float_null",
    "CountryofOrigin": "text",
    "WeightLbs": "float",
    "WeightOz": "float",
    "Length": "float",
    "Width": "float",
    "Height": "float",
    "OnOrder": "int",
    "QtySold30": "int",
    "P&L (30 days)": "float_null",
    "P&L (90 days)": "float_null",
    "AverageShippingCost (30 Days)": "float_null",
    "AverageShippingCost (90 Days)": "float_null",
    "SearchTerms": "text",
    "AmazonShippingTemplate": "text",
    "ASIN1": "text",
    "LastReceived": "datetime",
    "FBAFee": "float_null",
    "MinPrice": "float",
}


def _sql_identifier(name: str) -> str:
    """Bracket a column name for SQL Server when it isn't a plain identifier.

    ``insert_dataframe`` interpolates column names straight into the INSERT
    statement, so names containing spaces or symbols (e.g. ``P&L (30 days)``)
    must be wrapped in brackets to be valid T-SQL. Plain names pass through
    unchanged.
    """
    return name if name.replace("_", "").isalnum() else f"[{name}]"


def sellercloud_db(reports_cursor) -> None:
    """Replace every row in the SellerCloud SQL table with the latest export.

    Reads ``SellerCloud.xlsx`` from ``sellercloud_file_path``, normalizes
    column dtypes, ``DELETE``s all existing rows in the SellerCloud table,
    and bulk-inserts every row from the Excel file via ``insert_dataframe``.

    Args:
        reports_cursor: Active pyodbc cursor for the Reports database.

    Raises:
        RuntimeError: If a row fails to insert (the traceback carries the
            failing row's full column → value mapping via seller-automation-utils 0.7.1).
    """
    reports_cursor.execute(f"DELETE FROM {table_sellercloud}")
    log.info("Table rows deleted successfully.")

    text_cols = [c for c, t in COLUMN_TYPES.items() if t == "text"]
    int_cols = [c for c, t in COLUMN_TYPES.items() if t == "int"]
    float_cols = [c for c, t in COLUMN_TYPES.items() if t == "float"]
    float_null_cols = [c for c, t in COLUMN_TYPES.items() if t == "float_null"]
    datetime_cols = [c for c, t in COLUMN_TYPES.items() if t == "datetime"]

    # Read text columns as str so leading zeros and long IDs (eBayItemID,
    # WalmartAPIItemID) don't get coerced to floats and pick up a ".0".
    df = pd.read_excel(
        f"{sellercloud_file_path}/SellerCloud.xlsx",
        dtype={c: str for c in text_cols},
    )

    for col in float_cols:
        df[col] = pd.to_numeric(df[col], errors="coerce").fillna(0)
    for col in int_cols:
        df[col] = pd.to_numeric(df[col], errors="coerce").fillna(0).astype(int)
    for col in text_cols:
        # fillna before astype so blank cells become "" rather than "nan".
        df[col] = df[col].fillna("").astype(str)
    for col in float_null_cols:
        s = pd.to_numeric(df[col], errors="coerce")
        df[col] = s.astype(object).where(s.notna(), None)
    for col in datetime_cols:
        s = pd.to_datetime(df[col], errors="coerce")
        df[col] = s.astype(object).where(s.notna(), None)

    # insert_dataframe uses each name as both the SQL identifier and the
    # DataFrame key, so bracket the special-character columns and rename the
    # matching DataFrame columns to keep the two in sync.
    sql_columns = [_sql_identifier(c) for c in COLUMN_TYPES]
    df = df.rename(columns={
        c: ident for c, ident in zip(COLUMN_TYPES, sql_columns) if ident != c
    })
    database_utils.insert_dataframe(reports_cursor, table_sellercloud, df, sql_columns)
    log.info("Data inserted successfully.")


def main() -> None:
    """Daily SellerCloud-to-SQL sync pipeline.

    Checks that today's ``SellerCloud.xlsx`` has actually arrived in the
    configured folder before touching the DB. If the latest file is older
    than today, logs a warning and exits cleanly — downstream reports run
    on yesterday's data instead of corrupted partial state.

    Raises:
        SystemExit: On KeyboardInterrupt (clean) or unhandled exception
            (after sending a crash report via Outlook).
    """
    try:
        today = datetime.now().strftime("%Y-%m-%d")
        latest = latest_modified_date(sellercloud_file_path)

        if latest is None:
            log.warning(
                f"No files found in [cyan]{sellercloud_file_path}[/cyan]. "
                f"Skipping SellerCloud upload."
            )
            return

        if latest.strftime("%Y-%m-%d") != today:
            log.warning(
                f"Latest file in [cyan]{sellercloud_file_path}[/cyan] is dated "
                f"[cyan]{latest.strftime('%Y-%m-%d')}[/cyan], not today. "
                f"Skipping to preserve table integrity."
            )
            return

        log.info("Uploading [cyan]SellerCloud[/cyan] items to database.")
        reports_conn = custom_functions.sql_connection("Reports")
        reports_cursor = reports_conn.cursor()
        sellercloud_db(reports_cursor)
        reports_conn.close()
        log.success("Done!")

    except (KeyboardInterrupt, SystemExit):
        log.warning("Script interrupted by user.")
        raise SystemExit(0)

    except Exception:
        alert_utils.handle_crash(None, traceback.format_exc(), "SellerCloud Sync")
        raise SystemExit(1)


if ask_user("Run now?", "SellerCloud Sync"):
    main()
run_on_schedule(main, hour=5, minute=15, day_of_week="mon-sun")
