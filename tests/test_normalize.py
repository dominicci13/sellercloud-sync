"""Normalization of the Informed columns, run through the real ``sellercloud_db``.

The export is written as a real ``.xlsx`` so the read path (dtypes pandas infers
from the cells) is exercised too, not only the coercion loops.
"""
from __future__ import annotations

from types import SimpleNamespace
from unittest.mock import MagicMock

import pandas as pd
import pytest

import run_sellercloud_sync as rss

# Captured at import: the autouse fixture replaces the module attribute with a mock.
REAL_SELLERCLOUD_DB = rss.sellercloud_db

STRATEGY = "Informed.Co Strategy ID"
ENABLED = "Enable For Informed.Co"


def _write_export(sync: SimpleNamespace, strategy: list[object], enabled: list[bool]) -> None:
    """Write a fabricated export carrying every ``COLUMN_TYPES`` column.

    Args:
        sync: The autouse fixture namespace (``sync.export`` is the target path).
        strategy: Values for the strategy-ID column, one per row.
        enabled: Values for the Informed enable flag, one per row.
    """
    rows = len(strategy)
    filler = {
        "text": "x",
        "int": 1,
        "float": 1.5,
        "float_null": 2.5,
        "datetime": pd.Timestamp("2026-09-01 10:00"),
    }
    data = {
        col: [filler.get(kind)] * rows
        for col, kind in rss.COLUMN_TYPES.items()
    }
    data[STRATEGY] = strategy
    data[ENABLED] = enabled
    pd.DataFrame(data).to_excel(sync.export, sheet_name="Table", index=False)


def _inserted(sync: SimpleNamespace) -> pd.DataFrame:
    """Return the DataFrame ``sellercloud_db`` handed to ``insert_dataframe``."""
    sync.insert_dataframe.assert_called_once()
    _cursor, _table, df, columns = sync.insert_dataframe.call_args.args
    assert columns == list(rss.COLUMN_TYPES)
    return df


@pytest.mark.parametrize(
    ("raw", "expected"),
    [
        ("900000101", 900000101),
        (" 900000102", 900000102),
        ("900000101  ", 900000101),
        (None, None),
        ("XYZ-0001", None),
    ],
    ids=["plain", "leading-space", "trailing-space", "blank-stays-null", "garbage-stays-null"],
)
def test_strategy_id_keeps_blank_as_null_never_zero(
    sync: SimpleNamespace, raw: object, expected: int | None
) -> None:
    # The blank row matters: with no blank, pandas reads the column as int64 and
    # a float-typed coercion would pass unnoticed.
    _write_export(sync, [raw, None, "900000103"], [True, False, True])

    REAL_SELLERCLOUD_DB(MagicMock(name="cursor"))

    value = _inserted(sync)[STRATEGY].iloc[0]
    if expected is None:
        assert value is None
    else:
        assert value == expected
        # A float would bind as 900000101.0; a numpy int may not bind at all.
        assert type(value) is int


def test_enable_for_informed_lands_as_one_and_zero(sync: SimpleNamespace) -> None:
    _write_export(sync, ["900000101", None], [True, False])

    REAL_SELLERCLOUD_DB(MagicMock(name="cursor"))

    assert _inserted(sync)[ENABLED].tolist() == [1, 0]


@pytest.mark.parametrize(
    "raw",
    [900000101.5, 1e20, float("inf")],
    ids=["non-whole", "out-of-range", "infinite"],
)
def test_bad_number_strategy_id_loads_as_null_and_the_load_continues(
    sync: SimpleNamespace, raw: float
) -> None:
    _write_export(sync, [raw, None, "900000103"], [True, False, True])

    REAL_SELLERCLOUD_DB(MagicMock(name="cursor"))

    assert _inserted(sync)[STRATEGY].tolist() == [None, None, 900000103]
    warnings = [c.args[0] for c in sync.log.warning.call_args_list]
    assert any(w.startswith("1 non-blank") and STRATEGY in w for w in warnings)


def test_non_numeric_strategy_ids_are_counted_in_a_warning(sync: SimpleNamespace) -> None:
    _write_export(sync, ["XYZ-0001", "XYZ-0002", None, "900000103"], [True, True, False, True])

    REAL_SELLERCLOUD_DB(MagicMock(name="cursor"))

    warnings = [c.args[0] for c in sync.log.warning.call_args_list]
    assert any(w.startswith("2 non-blank") and STRATEGY in w for w in warnings)
