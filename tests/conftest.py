"""Shared fixtures: every outbound side effect of the sync is replaced by a mock.

Importing ``run_sellercloud_sync`` needs a local ``config/paths.json`` (copied
from the ``.example``); nothing under that path is read or written by the tests.
"""
from __future__ import annotations

import os
from datetime import date, datetime
from pathlib import Path
from types import SimpleNamespace
from unittest.mock import MagicMock

import pytest

import run_sellercloud_sync as rss

RUN_DATE = date(2026, 9, 22)


class _FixedDate(date):
    """``date`` whose ``today()`` is pinned, so no test depends on the clock."""

    @classmethod
    def today(cls) -> date:
        return RUN_DATE


def set_mtime(path: Path, when: datetime) -> Path:
    """Create ``path`` if needed and set its modification time.

    Args:
        path: File to create or touch.
        when: Local datetime to stamp as its mtime.

    Returns:
        ``path``, for chaining.
    """
    path.parent.mkdir(parents=True, exist_ok=True)
    if not path.exists():
        path.write_bytes(b"fabricated")
    ts = when.timestamp()
    os.utime(path, (ts, ts))
    return path


@pytest.fixture(autouse=True)
def sync(tmp_path: Path, monkeypatch: pytest.MonkeyPatch) -> SimpleNamespace:
    """Point the module at a temp folder and mock email, SQL, crash mail and logging.

    Autouse so no test can reach Outlook, SQL Server or the production log by
    forgetting a patch.

    Returns:
        Namespace exposing the mocks and the temp paths.
    """
    folder = tmp_path / "Multiplatform"
    folder.mkdir()
    ns = SimpleNamespace(
        folder=folder,
        export=folder / "SellerCloud.xlsx",
        state=tmp_path / "logs" / "freshness_alert_state.json",
        send_email=MagicMock(name="send_email"),
        sql_connection=MagicMock(name="sql_connection"),
        insert_dataframe=MagicMock(name="insert_dataframe"),
        handle_crash=MagicMock(name="handle_crash"),
        sellercloud_db=MagicMock(name="sellercloud_db"),
        log=MagicMock(name="log"),
    )
    monkeypatch.setenv("ALERT_EMAIL", "alerts@example.com")
    monkeypatch.setattr(rss.outlook, "send_email", ns.send_email)
    monkeypatch.setattr(rss.custom_functions, "sql_connection", ns.sql_connection)
    monkeypatch.setattr(rss.database_utils, "insert_dataframe", ns.insert_dataframe)
    monkeypatch.setattr(rss.alert_utils, "handle_crash", ns.handle_crash)
    monkeypatch.setattr(rss, "sellercloud_db", ns.sellercloud_db)
    monkeypatch.setattr(rss, "log", ns.log)
    monkeypatch.setattr(rss, "date", _FixedDate)
    monkeypatch.setattr(rss, "SELLERCLOUD_FILE", ns.export)
    monkeypatch.setattr(rss, "_FRESHNESS_STATE_PATH", ns.state)
    monkeypatch.setattr(rss, "_EMPTY_EXPORT_STATE_PATH", tmp_path / "logs" / "empty_export_alert_state.json")
    monkeypatch.setattr(rss, "_ALERT_STATE_PATH", tmp_path / "logs" / "schema_alert_state.json")
    return ns
