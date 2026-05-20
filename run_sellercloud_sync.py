"""Daily SellerCloud catalog sync.

Once per day at 05:15 local time (~15 minutes after the SellerCloud daily report
typically lands in OneDrive), this script:

1. Reads the most recent ``SellerCloud.xlsx`` from the configured OneDrive path.
2. Confirms its modification date is today (otherwise skips — yesterday's data
   is already in the table, and we'd rather miss a day than overwrite with
   stale data).
3. Clears the ``Reports.SellerCloud`` table.
4. Bulk-inserts every row via ``fc_utils.database_utils.insert_dataframe``.

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

from fc_utils import alert_utils, custom_functions, database_utils
from fc_utils.config_utils import load_config_safe
from fc_utils.file_utils import latest_modified_date
from fc_utils.logging_utils import setup_logging
from fc_utils.schedule_utils import run_on_schedule
from fc_utils.ui_utils import ask_user


log = setup_logging("sellercloud_sync")

load_dotenv()
table_sellercloud: str = os.getenv("DB_TABLE_SELLERCLOUD", "SellerCloud")
if not table_sellercloud.replace("_", "").isalnum():
    raise ValueError(f"Invalid table name: {table_sellercloud!r}")

# Resolve config relative to this file so the script works regardless of CWD
# (Task Scheduler / cron wrappers etc.)
_paths = load_config_safe(Path(__file__).resolve().parent / "config" / "paths.json")
sellercloud_file_path: str = _paths["sellercloud_file_path"]


def sellercloud_db(reports_cursor) -> None:
    """Replace every row in the SellerCloud SQL table with the latest export.

    Reads ``SellerCloud.xlsx`` from ``sellercloud_file_path``, normalizes
    column dtypes, ``DELETE``s all existing rows in the SellerCloud table,
    and bulk-inserts every row from the Excel file via ``insert_dataframe``.

    Args:
        reports_cursor: Active pyodbc cursor for the Reports database.

    Raises:
        RuntimeError: If a row fails to insert (the traceback carries the
            failing row's full column → value mapping via fc-utils 0.7.1).
    """
    reports_cursor.execute(f"DELETE FROM {table_sellercloud}")
    log.info("Table rows deleted successfully.")

    df = pd.read_excel(
        f"{sellercloud_file_path}/SellerCloud.xlsx",
        dtype={
            "SKU": str, "ASIN": str, "CompanyName": str, "WalmartAPIItemID": str,
            "eBayItemID": str, "ProductName": str, "Manufacturer": str,
            "Vendor": str, "BuyerEmail": str, "UPC": str, "CountryofOrigin": str,
        },
    )

    text_columns = [
        "SKU", "ASIN", "CompanyName", "WalmartAPIItemID", "eBayItemID",
        "ProductName", "Manufacturer", "Vendor", "BuyerEmail", "UPC", "CountryofOrigin",
    ]
    int_columns = ["MFNQuantity", "AggregateQty", "OnOrder", "CompanyID"]
    float_columns = [
        "AmazonPrice", "AmazonBusinessPrice", "ListPrice", "SitePrice", "MinPrice",
        "MAPPrice", "SiteCost", "WeightLbs", "WeightOz", "Length", "Width", "Height",
    ]

    for col in float_columns:
        df[col] = pd.to_numeric(df[col], errors="coerce").fillna(0)
    for col in text_columns:
        df[col] = df[col].astype(str)
    df[text_columns] = df[text_columns].fillna("")
    for col in int_columns:
        df[col] = pd.to_numeric(df[col], errors="coerce").fillna(0).astype(int)

    df_for_insert = df.rename(columns={
        "Manufacturer": "BrandName",
        "MFNQuantity": "AmazonQuantity",
    })
    columns = [
        "SKU", "ASIN", "CompanyName", "WalmartAPIItemID", "eBayItemID",
        "ProductName", "BrandName", "Vendor", "BuyerEmail", "UPC",
        "AmazonQuantity", "AggregateQty", "AmazonPrice", "AmazonBusinessPrice",
        "ListPrice", "SitePrice", "MAPPrice", "SiteCost", "CountryofOrigin",
        "WeightLbs", "WeightOz", "Length", "Width", "Height", "OnOrder",
        "CompanyID", "MinPrice",
    ]
    database_utils.insert_dataframe(reports_cursor, table_sellercloud, df_for_insert, columns)
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
