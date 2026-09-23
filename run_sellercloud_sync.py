"""Daily SellerCloud catalog sync.

Once per day at 05:15 local time (~15 minutes after the SellerCloud daily report
typically lands in OneDrive), this script:

1. Stats ``SellerCloud.xlsx`` itself (not the folder) in the configured OneDrive
   path and confirms its modification date is today. A stale or missing file
   leaves the table untouched, logs a WARNING and emails ``ALERT_EMAIL`` once
   per file date (``_alert_stale_export``) — we'd rather miss a day than
   overwrite with stale data, but never silently.
2. Reads the file.
3. Compares the report's headers against ``COLUMN_TYPES`` and emails a one-time
   alert (to ``ALERT_EMAIL``) when columns are added or removed, with copy-paste
   steps to wire them into this script and the SQL table. A *removed* mapped
   column aborts the write so a stale schema can't empty the table, and so does
   an export with no data rows (alerted once per file date).
4. Clears the ``Reports.SellerCloud`` table.
5. Bulk-inserts every row via ``seller_automation_utils.database_utils.insert_dataframe``.

Extracted from ``amzn-catalog-health`` in May 2026 so this daily refresh runs
independently of the larger nightly catalog/health scrape job at 04:00.
"""
from __future__ import annotations

import html
import json
import os
import traceback
import warnings
from datetime import date, datetime
from pathlib import Path

import pandas as pd
from dotenv import load_dotenv

from seller_automation_utils import alert_utils, custom_functions, database_utils, outlook
from seller_automation_utils.config_utils import get_env, load_config_safe
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
# The one file both the freshness guard and the reader use, so they cannot
# drift apart. Guarding the folder instead is what let a daily-rewritten
# sibling workbook pass the check for weeks while this file sat stale.
SELLERCLOUD_FILE: Path = Path(sellercloud_file_path) / "SellerCloud.xlsx"


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
# by the shared insert_dataframe (1.8.3+), so write them plainly
# (e.g. "P&L (30 days)").
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
    "IsBundle": "int",
    "QtySold60": "int",
    "Sales60": "float",
    "ConditionName": "text",
    "EnableSellingBelowCost": "int",
}


# How each COLUMN_TYPES type maps to a SQL Server column type for the ALTER
# TABLE suggestion emailed on schema drift. "{width}" is filled per-column from
# the sampled data length for text columns.
SCHEMA_TYPE_TO_SQL: dict[str, str] = {
    "text": "NVARCHAR({width})",
    "int": "INT",
    "float": "FLOAT",
    "float_null": "FLOAT",
    "datetime": "DATETIME",
}

# Tracks which drifted columns we've already emailed about so a daily run alerts
# once per column instead of every morning until it's mapped.
_ALERT_STATE_PATH = Path(__file__).resolve().parent / "logs" / "schema_alert_state.json"

# Separate file, not a key in the schema one: _save_alert_state rewrites that
# file whole, so a sibling key there would be erased on every schema check.
_FRESHNESS_STATE_PATH = Path(__file__).resolve().parent / "logs" / "freshness_alert_state.json"

# Its own file because main clears the freshness state on every fresh run,
# which would defeat this debounce on a same-day rerun.
_EMPTY_EXPORT_STATE_PATH = Path(__file__).resolve().parent / "logs" / "empty_export_alert_state.json"

_EXPORT_ORIGIN_NOTE = (
    "SellerCloud.xlsx is produced by a SellerCloud scheduled report, not by any "
    "automation in this repo."
)


def _sql_identifier(name: str) -> str:
    """Bracket a column name for SQL Server when it isn't a plain identifier.

    Used only to build the ``ALTER TABLE`` suggestion in the schema-drift
    alert, where a name with spaces or symbols (e.g. ``P&L (30 days)``) must be
    wrapped in brackets to be valid T-SQL. Plain names pass through unchanged.
    The INSERT path does not use it: ``insert_dataframe`` brackets every
    column itself since seller-automation-utils 1.8.3.

    Args:
        name (str): Column name as it appears in the report header.

    Returns:
        str: ``name`` unchanged, or ``[name]`` when it is not a plain identifier.
    """
    return name if name.replace("_", "").isalnum() else f"[{name.replace(']', ']]')}]"


def _diff_schema(
    excel_cols, schema_cols
) -> tuple[list[str], list[str]]:
    """Compare the report's headers against the known schema.

    Args:
        excel_cols: Column names read from ``SellerCloud.xlsx``.
        schema_cols: Column names this script knows about (``COLUMN_TYPES`` keys).

    Returns:
        ``(new_cols, missing_cols)`` — headers in the report but not the schema,
        and schema columns no longer present in the report.
    """
    excel = list(excel_cols)
    schema = list(schema_cols)
    new_cols = [c for c in excel if c not in schema]
    missing_cols = [c for c in schema if c not in excel]
    return new_cols, missing_cols


def _infer_column_type(series: pd.Series) -> str:
    """Guess a ``COLUMN_TYPES`` type from a column's actual values.

    Numeric is checked before datetime so numeric IDs (e.g. ``"12345"``) classify
    as ``int``/``float`` rather than being parsed as dates. Blank/all-null columns
    fall back to ``text``.

    Args:
        series: The report column to inspect.

    Returns:
        One of ``"text"``, ``"int"``, ``"float"``, ``"datetime"``.
    """
    sample = series.dropna()
    sample = sample[sample.astype(str).str.strip() != ""]
    if sample.empty:
        return "text"

    numeric = pd.to_numeric(sample, errors="coerce")
    if numeric.notna().all():
        return "int" if (numeric == numeric.round()).all() else "float"

    # A failed parse is the expected path for text columns; the per-element
    # "could not infer format" warning is just noise here.
    with warnings.catch_warnings():
        warnings.simplefilter("ignore", UserWarning)
        if pd.to_datetime(sample, errors="coerce").notna().all():
            return "datetime"

    return "text"


def _describe_new_column(df: pd.DataFrame, col: str, table: str) -> dict[str, object]:
    """Build the email facts for one newly-seen report column.

    Args:
        df: The freshly-read report.
        col: The new column's header.
        table: SQL table name, used in the suggested ALTER statement.

    Returns:
        Dict with sample values, inferred type, and ready-to-paste
        ``COLUMN_TYPES`` line and ``ALTER TABLE`` statement.
    """
    series = df[col]
    values = series.dropna().astype(str).map(str.strip)
    values = values[values != ""]
    samples = values.unique()[:5].tolist()

    inferred = _infer_column_type(series)
    if inferred == "text":
        max_len = int(values.map(len).max()) if not values.empty else 0
        width = max(50, ((max_len // 50) + 1) * 50)
        sql_type = SCHEMA_TYPE_TO_SQL["text"].format(width=width)
    else:
        sql_type = SCHEMA_TYPE_TO_SQL[inferred]

    return {
        "name": col,
        "samples": samples,
        "inferred": inferred,
        "column_line": f'    "{col}": "{inferred}",',
        "alter_stmt": f"ALTER TABLE {table} ADD {_sql_identifier(col)} {sql_type} NULL;",
    }


def _load_alert_state(path: Path) -> dict[str, list[str]]:
    """Read the already-alerted column lists, tolerating a missing/corrupt file."""
    try:
        with open(path, "r", encoding="utf-8") as fh:
            data = json.load(fh)
    except (FileNotFoundError, json.JSONDecodeError):
        return {"new": [], "missing": []}
    return {"new": list(data.get("new", [])), "missing": list(data.get("missing", []))}


def _save_alert_state(path: Path, new_cols: list[str], missing_cols: list[str]) -> None:
    """Persist the current drift sets so resolved columns drop out next run."""
    path.parent.mkdir(parents=True, exist_ok=True)
    with open(path, "w", encoding="utf-8") as fh:
        json.dump({"new": new_cols, "missing": missing_cols}, fh, indent=2)


def _build_alert_body(
    new_described: list[dict[str, object]], missing_cols: list[str], table: str
) -> str:
    """Render the schema-drift email body as HTML with copy-paste fix steps."""
    parts = [
        "<p>The latest <b>SellerCloud.xlsx</b> no longer matches the columns this sync "
        "knows about (<code>COLUMN_TYPES</code> in <code>run_sellercloud_sync.py</code>).</p>"
    ]

    if new_described:
        parts.append("<h3>New columns in the report</h3>")
        for d in new_described:
            samples = ", ".join(d["samples"]) if d["samples"] else "(no sample values)"
            parts.append(
                f"<p><b>{html.escape(str(d['name']))}</b><br>"
                f"Sample values: <code>{html.escape(samples)}</code><br>"
                f"Inferred type: <b>{d['inferred']}</b></p>"
            )

        py_lines = "<br>".join(html.escape(d["column_line"].strip()) for d in new_described)
        alter_lines = "<br>".join(html.escape(str(d["alter_stmt"])) for d in new_described)

        parts.append(
            "<h4>1) Add to the Python script</h4>"
            "<ol>"
            "<li>Open <code>run_sellercloud_sync.py</code>.</li>"
            "<li>In the <code>COLUMN_TYPES</code> dict, add a line for each new column "
            "<b>at the end of the dict</b> (the SQL ALTER below appends the column to the "
            "end of the table, so keep both in the same order):</li>"
            "</ol>"
            f"<pre>{py_lines}</pre>"
            "<p>Pick the type deliberately: <code>text</code> for IDs/codes (even ones that "
            "look numeric, to keep any leading zeros), <code>int</code>/<code>float</code> "
            "for counts/money, <code>datetime</code> for dates. Use <code>float_null</code> "
            "instead of <code>float</code> when a blank cell must stay NULL (an analytical "
            "metric where blank is not the same as 0).</p>"
        )
        parts.append(
            "<h4>2) Add to the SQL database</h4>"
            "<ol>"
            "<li>Open SQL Server Management Studio and connect to the <b>Reports</b> database.</li>"
            f"<li>Run the statement(s) below — one per new column. This adds each to the "
            f"<code>{table}</code> table as nullable, so existing rows stay valid:</li>"
            "</ol>"
            f"<pre>{alter_lines}</pre>"
            "<p>The <code>NVARCHAR(n)</code> width is sized from the sample data; raise it if "
            "you expect longer values. Once both edits are in, the next daily run picks the "
            "column up automatically.</p>"
        )

    if missing_cols:
        cols = ", ".join(missing_cols)
        py_remove = "<br>".join(f'"{c}": ...,' for c in missing_cols)
        parts.append("<h3 style='color:#b00'>Columns that disappeared from the report</h3>")
        parts.append(
            f"<p>These are still in <code>COLUMN_TYPES</code> but no longer in the report: "
            f"<b>{cols}</b>. The sync was <b>skipped</b> to protect the table — it can't build "
            f"these columns' values without them, and would otherwise empty the table.</p>"
            "<h4>How to fix</h4>"
            "<ol>"
            "<li>If a column was <b>renamed</b>, change its key in <code>COLUMN_TYPES</code> to "
            "the new header.</li>"
            "<li>If it was <b>removed for good</b>, delete its line(s) from "
            "<code>COLUMN_TYPES</code>:"
            f"<pre>{py_remove}</pre></li>"
            f"<li>Optionally drop it in SQL once you're sure: "
            f"<pre>ALTER TABLE {table} DROP COLUMN &lt;column&gt;;</pre></li>"
            "</ol>"
        )

    return "\n".join(parts)


def _alert_schema_drift(
    df: pd.DataFrame, table: str, state_path: Path = _ALERT_STATE_PATH
) -> list[str]:
    """Detect schema drift, email new/missing columns once each, persist state.

    New columns are harmless to the insert (it writes an explicit column list),
    so the caller can keep syncing. Missing columns can't be built, so the
    returned list signals the caller to skip the DB write.

    Args:
        df: The freshly-read report.
        table: SQL table name for the suggested DDL.
        state_path: Where the already-alerted column lists live.

    Returns:
        The list of missing (schema-but-not-report) columns; empty when safe.
    """
    new_cols, missing_cols = _diff_schema(df.columns, COLUMN_TYPES.keys())
    if not new_cols and not missing_cols:
        return missing_cols

    state = _load_alert_state(state_path)
    unreported_new = [c for c in new_cols if c not in state["new"]]
    unreported_missing = [c for c in missing_cols if c not in state["missing"]]

    if unreported_new or unreported_missing:
        alert_email = get_env("ALERT_EMAIL", required=True)
        described = [_describe_new_column(df, c, table) for c in unreported_new]
        subject_bits = []
        if unreported_new:
            subject_bits.append(f"{len(unreported_new)} new")
        if unreported_missing:
            subject_bits.append(f"{len(unreported_missing)} missing")
        outlook.send_email(
            account=alert_email,
            subject=f"[SCHEMA] SellerCloud report columns changed — {', '.join(subject_bits)}",
            body=_build_alert_body(described, unreported_missing, table),
            to=[alert_email],
            show=False,
            send=True,
        )
        log.warning(
            f"Schema drift detected — new={unreported_new} missing={unreported_missing}. "
            f"Alert emailed to [cyan]{alert_email}[/cyan]."
        )

    _save_alert_state(state_path, new_cols, missing_cols)
    return missing_cols


def export_is_fresh(path: Path, today: date) -> tuple[bool, datetime | None]:
    """Decide whether the export file itself was written today.

    Stats ``path`` directly. Never infer freshness from the folder: another
    workbook in the same folder is rewritten daily and would always read as
    "today".

    Args:
        path: The export file, normally ``SELLERCLOUD_FILE``.
        today: The run's local date.

    Returns:
        ``(fresh, modified)`` — ``fresh`` is True only when the file exists and
        its local modification date equals ``today``. ``modified`` is the file's
        LastWriteTime, or None when the file does not exist.

    Raises:
        OSError: Any stat failure other than a missing file (permissions, an
            unreachable drive), so the crash handler reports it instead of the
            run treating it as "missing".
    """
    try:
        mtime = path.stat().st_mtime
    except FileNotFoundError:
        return False, None
    modified = datetime.fromtimestamp(mtime)
    return modified.date() == today, modified


def _stale_alert_key(modified: datetime | None) -> str:
    """Debounce key: one alert per stale file date, or one for a missing file.

    Args:
        modified: The file's LastWriteTime, or None when it is missing.

    Returns:
        ``"missing"`` or the ISO date of ``modified``.
    """
    return "missing" if modified is None else modified.date().isoformat()


def _load_freshness_state(path: Path) -> str | None:
    """Read the key of the last stale-export alert sent.

    A missing, unreadable or corrupt file reads as "nothing alerted yet", which
    errs toward sending the alert again rather than staying quiet.

    Args:
        path: The freshness state file.

    Returns:
        The last alerted key, or None.
    """
    try:
        with open(path, "r", encoding="utf-8") as fh:
            data = json.load(fh)
    except FileNotFoundError:
        return None
    # OSError: permissions or a directory at the path; ValueError: bad JSON or
    # non-UTF-8 bytes (UnicodeDecodeError).
    except (OSError, ValueError) as exc:
        log.warning(f"Alert state [cyan]{path}[/cyan] unreadable ({exc!r}); treating as empty.")
        return None
    alerted = data.get("alerted") if isinstance(data, dict) else None
    return alerted if isinstance(alerted, str) else None


def _save_freshness_state(path: Path, alerted: str | None) -> None:
    """Persist the last alerted key (None clears it once the export is fresh again).

    Args:
        path: The freshness state file.
        alerted: Key from ``_stale_alert_key``, or None.
    """
    path.parent.mkdir(parents=True, exist_ok=True)
    with open(path, "w", encoding="utf-8") as fh:
        json.dump({"alerted": alerted}, fh, indent=2)


def _build_stale_alert(
    path: Path, modified: datetime | None, today: date, table: str
) -> tuple[str, str]:
    """Render the stale/missing-export alert.

    Args:
        path: The export file that was checked.
        modified: Its LastWriteTime, or None when it is missing.
        today: The run's local date.
        table: SQL table the sync would have reloaded.

    Returns:
        ``(subject, html_body)``. Every interpolated value is HTML-escaped.
    """
    e_path = html.escape(str(path))
    e_table = html.escape(table)
    if modified is None:
        subject = "[STALE] SellerCloud export is missing - sync skipped"
        facts = (
            f"<p><b>File:</b> <code>{e_path}</code><br>"
            f"<b>Status:</b> not found<br>"
            f"<b>Run date:</b> {today.isoformat()}</p>"
        )
    else:
        days = (today - modified.date()).days
        age = f"{days} day(s) stale" if days > 0 else "dated in the future"
        subject = (
            f"[STALE] SellerCloud export dated {modified.date().isoformat()} "
            f"({age}) - sync skipped"
        )
        facts = (
            f"<p><b>File:</b> <code>{e_path}</code><br>"
            f"<b>LastWriteTime:</b> {html.escape(modified.strftime('%Y-%m-%d %H:%M:%S'))}<br>"
            f"<b>Age:</b> {html.escape(age)}<br>"
            f"<b>Run date:</b> {today.isoformat()}</p>"
        )
    body = "\n".join(
        [
            "<p>The daily SellerCloud sync did <b>not</b> reload the table because "
            "the export file was not written today.</p>",
            facts,
            f"<p><b>The <code>{e_table}</code> table was left untouched</b> (no DELETE, "
            "no insert). It still holds whatever the last successful load wrote, so "
            "anything reading it is working from that date's catalog.</p>",
            f"<p>{html.escape(_EXPORT_ORIGIN_NOTE)}</p>",
            "<p>No action is needed in this repo: the next 05:15 run reloads "
            "automatically once a file dated that day lands. This email is sent once "
            "per file date; the log repeats the WARNING every run until then.</p>",
        ]
    )
    return subject, body


def _alert_stale_export(
    path: Path,
    modified: datetime | None,
    today: date,
    table: str,
    state_path: Path = _FRESHNESS_STATE_PATH,
) -> bool:
    """Log a stale/missing export every run and email it once per file date.

    The state is written only after the email is handed to Outlook, so a failed
    send propagates to the crash handler and the alert is retried next run.

    Args:
        path: The export file that was checked.
        modified: Its LastWriteTime, or None when it is missing.
        today: The run's local date.
        table: SQL table the sync would have reloaded.
        state_path: Where the last alerted key lives.

    Returns:
        True when an email was sent this run, False when debounced.

    Raises:
        ValueError: If ``ALERT_EMAIL`` is not set.
        Exception: Whatever ``outlook.send_email`` raises on a failed send.
    """
    what = "is missing" if modified is None else f"is dated {modified:%Y-%m-%d %H:%M:%S}"
    log.warning(
        f"[cyan]{path}[/cyan] {what}, not today ({today.isoformat()}). "
        f"Table [cyan]{table}[/cyan] left untouched."
    )

    key = _stale_alert_key(modified)
    if _load_freshness_state(state_path) == key:
        log.warning(f"Stale-export alert for '{key}' already emailed; not re-sending.")
        return False

    alert_email = get_env("ALERT_EMAIL", required=True)
    subject, body = _build_stale_alert(path, modified, today, table)
    outlook.send_email(
        account=alert_email,
        subject=subject,
        body=body,
        to=[alert_email],
        show=False,
        send=True,
    )
    _save_freshness_state(state_path, key)
    log.warning(f"Stale-export alert emailed to [cyan]{alert_email}[/cyan].")
    return True


def _alert_empty_export(path: Path, table: str, state_path: Path) -> bool:
    """Log a fresh-dated export with no data rows every run; email it once per file date.

    Uses the same state-file helpers and send-then-save ordering as
    ``_alert_stale_export``.

    Args:
        path: The export file that was read.
        table: SQL table the sync would have reloaded.
        state_path: Where the last alerted key lives.

    Returns:
        True when an email was sent this run, False when debounced.

    Raises:
        ValueError: If ``ALERT_EMAIL`` is not set.
        Exception: Whatever ``outlook.send_email`` raises on a failed send.
    """
    modified = datetime.fromtimestamp(path.stat().st_mtime)
    log.warning(
        f"[cyan]{path}[/cyan] (dated {modified:%Y-%m-%d %H:%M:%S}) has no data rows. "
        f"Table [cyan]{table}[/cyan] left untouched."
    )

    key = modified.date().isoformat()
    if _load_freshness_state(state_path) == key:
        log.warning(f"Empty-export alert for '{key}' already emailed; not re-sending.")
        return False

    alert_email = get_env("ALERT_EMAIL", required=True)
    body = "\n".join(
        [
            "<p>The daily SellerCloud sync did <b>not</b> reload the table because "
            "the export file has headers but no data rows.</p>",
            f"<p><b>File:</b> <code>{html.escape(str(path))}</code><br>"
            f"<b>LastWriteTime:</b> {html.escape(modified.strftime('%Y-%m-%d %H:%M:%S'))}</p>",
            f"<p><b>The <code>{html.escape(table)}</code> table was left untouched</b> "
            "(no DELETE, no insert).</p>",
            f"<p>{html.escape(_EXPORT_ORIGIN_NOTE)}</p>",
        ]
    )
    outlook.send_email(
        account=alert_email,
        subject=f"[EMPTY] SellerCloud export dated {key} has no rows - sync skipped",
        body=body,
        to=[alert_email],
        show=False,
        send=True,
    )
    _save_freshness_state(state_path, key)
    log.warning(f"Empty-export alert emailed to [cyan]{alert_email}[/cyan].")
    return True


def sellercloud_db(reports_cursor) -> None:
    """Replace every row in the SellerCloud SQL table with the latest export.

    Reads ``SELLERCLOUD_FILE`` (the same file ``main`` freshness-checked) and checks its
    headers against ``COLUMN_TYPES`` first. New or missing columns trigger a
    one-time email (``_alert_schema_drift``); missing columns also abort the
    write so a stale-schema run never empties the table. An export with headers
    but no data rows also aborts the write and alerts once per file date
    (``_alert_empty_export``). Otherwise it normalizes
    column dtypes, ``DELETE``s all existing rows, and bulk-inserts every row via
    ``insert_dataframe``.

    The read happens *before* the ``DELETE`` on purpose: if the file is unreadable
    or its schema drifted, the table keeps yesterday's data instead of being
    cleared and then left empty by a crash.

    Column names go to ``insert_dataframe`` exactly as ``COLUMN_TYPES`` declares
    them; seller-automation-utils 1.8.3+ brackets each one, so ``P&L (30 days)``
    needs no renaming here. An older library interpolates names verbatim and
    the INSERT fails, so the venv must be on 1.8.3 or later.

    Args:
        reports_cursor: Active pyodbc cursor for the Reports database.

    Raises:
        RuntimeError: If a row fails to insert (the traceback carries the
            failing row's full column → value mapping, added in seller-automation-utils 0.7.1).
    """
    text_cols = [c for c, t in COLUMN_TYPES.items() if t == "text"]
    int_cols = [c for c, t in COLUMN_TYPES.items() if t == "int"]
    float_cols = [c for c, t in COLUMN_TYPES.items() if t == "float"]
    float_null_cols = [c for c, t in COLUMN_TYPES.items() if t == "float_null"]
    datetime_cols = [c for c, t in COLUMN_TYPES.items() if t == "datetime"]

    # Read text columns as str so leading zeros and long IDs (eBayItemID,
    # WalmartAPIItemID) don't get coerced to floats and pick up a ".0".
    df = pd.read_excel(
        SELLERCLOUD_FILE,
        dtype={c: str for c in text_cols},
    )

    missing_cols = _alert_schema_drift(df, table_sellercloud)
    if missing_cols:
        log.warning(
            f"Mapped column(s) missing from the report: {missing_cols}. "
            f"Skipping the DB write to keep the table intact."
        )
        return

    if df.empty:
        _alert_empty_export(
            SELLERCLOUD_FILE, table_sellercloud, state_path=_EMPTY_EXPORT_STATE_PATH
        )
        return

    reports_cursor.execute(f"DELETE FROM {table_sellercloud}")
    log.info("Table rows deleted successfully.")

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

    database_utils.insert_dataframe(reports_cursor, table_sellercloud, df, list(COLUMN_TYPES))
    log.info("Data inserted successfully.")


def main() -> None:
    """Daily SellerCloud-to-SQL sync pipeline.

    Checks that ``SELLERCLOUD_FILE`` itself was written today before opening a
    DB connection. A stale or missing file goes to ``_alert_stale_export``
    (WARNING every run, email once per file date) and the run returns without
    connecting, so the table keeps its current rows. A fresh file clears the
    alert state, so the next stale date alerts again.

    Raises:
        SystemExit: On KeyboardInterrupt (clean) or unhandled exception
            (after sending a crash report via Outlook). A failed stale-export
            email is such an exception, so it is never swallowed.
    """
    try:
        today = date.today()
        fresh, modified = export_is_fresh(SELLERCLOUD_FILE, today)
        if not fresh:
            _alert_stale_export(
                SELLERCLOUD_FILE, modified, today, table_sellercloud,
                state_path=_FRESHNESS_STATE_PATH,
            )
            return
        try:
            if _load_freshness_state(_FRESHNESS_STATE_PATH) is not None:
                _save_freshness_state(_FRESHNESS_STATE_PATH, None)
        except OSError as exc:
            # Debounce bookkeeping must never cost a day's load.
            log.warning(f"Could not clear the stale-export alert state ({exc!r}); loading anyway.")

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


if __name__ == "__main__":
    if ask_user("Run now?", "SellerCloud Sync"):
        main()
    run_on_schedule(main, hour=5, minute=15, day_of_week="mon-sun")
