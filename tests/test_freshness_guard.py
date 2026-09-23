"""Freshness guard: a stale, missing or empty export must alert and leave the table alone."""
from __future__ import annotations

import json
import re
from datetime import date, datetime, timedelta
from pathlib import Path
from unittest.mock import MagicMock

import pandas as pd
import pytest
from seller_automation_utils.file_utils import latest_modified_date

import run_sellercloud_sync as rss
from conftest import RUN_DATE, set_mtime

STALE = datetime(2026, 8, 25, 5, 3, 48)
FRESH = datetime(2026, 9, 22, 5, 1, 0)

# Captured at import, before the autouse fixture swaps it for a mock.
_REAL_SELLERCLOUD_DB = rss.sellercloud_db


def _assert_table_untouched(sync) -> None:
    sync.sql_connection.assert_not_called()
    sync.sellercloud_db.assert_not_called()
    sync.insert_dataframe.assert_not_called()


# --- export_is_fresh (pure decision) -------------------------------------


@pytest.mark.parametrize(
    ("mtime", "today", "expected"),
    [
        (FRESH, RUN_DATE, True),
        (datetime(2026, 9, 22, 0, 0, 1), RUN_DATE, True),
        (datetime(2026, 9, 21, 23, 59, 59), RUN_DATE, False),
        (STALE, RUN_DATE, False),
        (datetime(2026, 9, 23, 5, 0, 0), RUN_DATE, False),
    ],
    ids=["today", "just-after-midnight", "just-before-midnight", "28-days-old", "future"],
)
def test_export_is_fresh_compares_the_files_own_date(sync, mtime, today, expected):
    set_mtime(sync.export, mtime)

    fresh, modified = rss.export_is_fresh(sync.export, today)

    assert fresh is expected
    assert modified == mtime


def test_missing_export_is_not_fresh(sync):
    assert rss.export_is_fresh(sync.export, RUN_DATE) == (False, None)


def test_fresh_sibling_file_does_not_make_a_stale_export_fresh(sync):
    set_mtime(sync.export, STALE)
    set_mtime(sync.folder / "All Items.xlsm", FRESH)
    # The old folder-wide guard sees the sibling's date; that is the bug.
    assert latest_modified_date(str(sync.folder)).date() == RUN_DATE

    fresh, modified = rss.export_is_fresh(sync.export, RUN_DATE)

    assert fresh is False
    assert modified == STALE


def test_stat_failure_other_than_missing_is_raised_not_treated_as_missing():
    class Unreadable:
        def stat(self):
            raise PermissionError("denied")

    with pytest.raises(PermissionError):
        rss.export_is_fresh(Unreadable(), RUN_DATE)


# --- main(): the guard's effect on the run --------------------------------


def test_fresh_export_reloads_the_table_without_alerting(sync):
    set_mtime(sync.export, FRESH)

    rss.main()

    sync.sql_connection.assert_called_once_with("Reports")
    sync.sellercloud_db.assert_called_once()
    sync.send_email.assert_not_called()
    sync.handle_crash.assert_not_called()


def test_stale_export_alerts_and_leaves_the_table_untouched(sync):
    set_mtime(sync.export, STALE)

    rss.main()

    _assert_table_untouched(sync)
    sync.send_email.assert_called_once()
    kwargs = sync.send_email.call_args.kwargs
    assert kwargs["to"] == ["alerts@example.com"]
    assert kwargs["send"] is True and kwargs["show"] is False
    assert "2026-08-25" in kwargs["subject"]
    assert sync.log.warning.called
    sync.handle_crash.assert_not_called()


def test_stale_export_beside_a_fresh_sibling_still_alerts_and_skips(sync):
    set_mtime(sync.export, STALE)
    set_mtime(sync.folder / "All Items.xlsm", FRESH)

    rss.main()

    _assert_table_untouched(sync)
    sync.send_email.assert_called_once()


def test_missing_export_alerts_and_leaves_the_table_untouched(sync):
    set_mtime(sync.folder / "All Items.xlsm", FRESH)

    rss.main()

    _assert_table_untouched(sync)
    sync.send_email.assert_called_once()
    assert "missing" in sync.send_email.call_args.kwargs["subject"]
    sync.handle_crash.assert_not_called()


# --- debounce ---------------------------------------------------------------


def test_second_stale_run_with_same_file_date_does_not_re_alert(sync):
    set_mtime(sync.export, STALE)

    rss.main()
    rss.main()

    assert sync.send_email.call_count == 1
    _assert_table_untouched(sync)
    # Still loud in the log on the debounced run.
    assert sync.log.warning.call_count >= 2


def test_new_stale_file_date_alerts_again(sync):
    set_mtime(sync.export, STALE)
    rss.main()
    set_mtime(sync.export, STALE + timedelta(days=1))
    rss.main()

    assert sync.send_email.call_count == 2


def test_missing_file_alerts_once_then_again_after_a_fresh_run(sync):
    rss.main()
    rss.main()
    assert sync.send_email.call_count == 1

    set_mtime(sync.export, FRESH)
    rss.main()
    assert json.loads(sync.state.read_text(encoding="utf-8")) == {"alerted": None}

    sync.export.unlink()
    rss.main()
    assert sync.send_email.call_count == 2


def test_corrupt_state_file_errs_toward_alerting(sync):
    set_mtime(sync.export, STALE)
    sync.state.parent.mkdir(parents=True, exist_ok=True)
    sync.state.write_text("{not json", encoding="utf-8")

    rss.main()

    sync.send_email.assert_called_once()


def test_non_utf8_state_file_errs_toward_alerting(sync):
    set_mtime(sync.export, STALE)
    sync.state.parent.mkdir(parents=True, exist_ok=True)
    sync.state.write_bytes(b"\xff\xfe\x80 not utf-8")

    rss.main()

    sync.send_email.assert_called_once()
    sync.handle_crash.assert_not_called()


def test_directory_at_state_path_reads_as_empty_and_fresh_load_proceeds(sync):
    sync.state.mkdir(parents=True)

    assert rss._load_freshness_state(sync.state) is None
    assert sync.log.warning.called

    set_mtime(sync.export, FRESH)
    rss.main()

    sync.sellercloud_db.assert_called_once()
    sync.handle_crash.assert_not_called()


def test_failing_state_clear_on_fresh_run_still_loads(sync, monkeypatch):
    set_mtime(sync.export, STALE)
    rss.main()
    set_mtime(sync.export, FRESH)
    monkeypatch.setattr(
        rss, "_save_freshness_state", MagicMock(side_effect=PermissionError("locked"))
    )

    rss.main()

    sync.sellercloud_db.assert_called_once()
    sync.handle_crash.assert_not_called()
    assert any("Could not clear" in str(c) for c in sync.log.warning.call_args_list)


def test_header_only_fresh_export_alerts_and_does_not_delete(sync):
    pd.DataFrame(columns=list(rss.COLUMN_TYPES)).to_excel(sync.export, index=False)
    set_mtime(sync.export, FRESH)
    cursor = MagicMock(name="cursor")

    _REAL_SELLERCLOUD_DB(cursor)

    cursor.execute.assert_not_called()
    sync.insert_dataframe.assert_not_called()
    sync.send_email.assert_called_once()
    kwargs = sync.send_email.call_args.kwargs
    assert kwargs["to"] == ["alerts@example.com"]
    assert "2026-09-22" in kwargs["subject"]
    assert "no rows" in kwargs["subject"]
    assert sync.log.warning.called


def test_failed_send_crashes_loudly_and_retries_next_run(sync):
    set_mtime(sync.export, STALE)
    sync.send_email.side_effect = RuntimeError("Outlook unavailable")

    with pytest.raises(SystemExit) as exc:
        rss.main()

    assert exc.value.code == 1
    sync.handle_crash.assert_called_once()
    assert not sync.state.exists()
    _assert_table_untouched(sync)

    sync.send_email.side_effect = None
    rss.main()
    assert sync.send_email.call_count == 2


def test_missing_alert_email_crashes_rather_than_skipping_quietly(sync, monkeypatch):
    set_mtime(sync.export, STALE)
    monkeypatch.delenv("ALERT_EMAIL")

    with pytest.raises(SystemExit):
        rss.main()

    sync.handle_crash.assert_called_once()
    sync.send_email.assert_not_called()
    assert not sync.state.exists()
    _assert_table_untouched(sync)


# --- alert content ----------------------------------------------------------


def test_stale_alert_states_the_facts():
    path = Path("C:/data/Multiplatform/SellerCloud.xlsx")

    subject, body = rss._build_stale_alert(path, STALE, RUN_DATE, "SellerCloud")

    assert "28 day(s) stale" in subject
    assert "C:/data/Multiplatform/SellerCloud.xlsx" in body.replace("\\", "/")
    assert "2026-08-25 05:03:48" in body
    assert "28 day(s) stale" in body
    assert "left untouched" in body
    assert rss._EXPORT_ORIGIN_NOTE in body
    assert "scheduled report" in body
    # Public repo: no ticket system or number in the alert, named or numbered.
    assert "ticket" not in body.lower()
    assert re.search(r"#\s*\d", body) is None


def test_missing_alert_states_the_facts():
    subject, body = rss._build_stale_alert(
        Path("C:/data/SellerCloud.xlsx"), None, RUN_DATE, "SellerCloud"
    )

    assert "missing" in subject
    assert "not found" in body
    assert "left untouched" in body


@pytest.mark.parametrize("modified", [STALE, None], ids=["stale", "missing"])
def test_alert_body_escapes_interpolated_values(modified):
    # No slash inside the markup: Path would turn it into a Windows separator.
    path = Path("C:/data/R&D <x>/SellerCloud.xlsx")

    _, body = rss._build_stale_alert(path, modified, RUN_DATE, "Sync<&>Table")

    assert "R&amp;D &lt;x&gt;" in body
    assert "R&D <x>" not in body
    assert "Sync&lt;&amp;&gt;Table" in body
    assert "Sync<&>Table" not in body


def test_stale_alert_key():
    assert rss._stale_alert_key(None) == "missing"
    assert rss._stale_alert_key(STALE) == "2026-08-25"


def test_run_date_is_pinned_by_the_fixture():
    assert rss.date.today() == RUN_DATE
    assert isinstance(RUN_DATE, date)
