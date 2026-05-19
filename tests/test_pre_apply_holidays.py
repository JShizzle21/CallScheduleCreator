"""Tests for scheduler_main._pre_apply_holidays.

Exercises manual holiday assignment pre-apply: name/PGY validation, counter
bumps via apply_assignment, blank-cell behaviour, and deduping with
completed_calls.
"""

from __future__ import annotations

from datetime import date

import pytest

import scheduler_main as sm
from scheduler_main import (
    SLOT_INTERN_WEEKDAY,
    SLOT_INTERN_WEEKEND,
    SLOT_UPPER_WEEKDAY,
    SLOT_UPPER_WEEKEND,
    _pre_apply_holidays,
)


class _StubLookup:
    """Minimal lookup replacement — only rotation_on_date() is called."""

    def rotation_on_date(self, name, d):
        return "WARDS"


def make_resident(pgy: int) -> dict:
    return {
        "pgy": pgy,
        "assigned_dates": [],
        "total_calls": 0,
        "weekday_calls": 0,
        "weekend_calls": 0,
        "friday_calls": 0,
        "saturday_calls": 0,
        "upper_calls": 0,
        "intern_calls": 0,
        "Jul_Dec_calls": 0,
        "Jan_Jun_calls": 0,
    }


# A Friday (weekday) well inside the academic year. Using ACADEMIC_YEAR_START
# means it lands in the Jul_Dec half, which lets us check that counter too.
WEEKDAY = date(sm.ACADEMIC_YEAR_START, 12, 25)  # Dec 25 2026 is a Friday
WEEKEND = date(sm.ACADEMIC_YEAR_START, 12, 26)  # Saturday

assert WEEKDAY.weekday() == 4, "test date expected to be Friday"
assert WEEKEND.weekday() == 5, "test date expected to be Saturday"


def test_weekend_holiday_bumps_both_counters():
    residents = {"Maria": make_resident(2), "Bob": make_resident(1)}
    holidays = {WEEKEND: {"name": "Boxing Day", "upper": "Maria", "intern": "Bob"}}

    rows, unassigned, assignments = _pre_apply_holidays(
        holidays=holidays,
        residents=residents,
        lookup=_StubLookup(),
        block1_end=None,
    )

    assert residents["Maria"]["weekend_calls"] == 1
    assert residents["Maria"]["upper_calls"] == 1
    assert residents["Maria"]["total_calls"] == 1
    assert residents["Bob"]["weekend_calls"] == 1
    assert residents["Bob"]["intern_calls"] == 1
    assert residents["Bob"]["total_calls"] == 1

    assert len(rows) == 2
    assert {r["slot"] for r in rows} == {SLOT_UPPER_WEEKEND, SLOT_INTERN_WEEKEND}
    assert all(r["note"] == "HOLIDAY: Boxing Day" for r in rows)
    assert unassigned == []
    assert assignments == {
        (WEEKEND, SLOT_UPPER_WEEKEND): "Maria",
        (WEEKEND, SLOT_INTERN_WEEKEND): "Bob",
    }


def test_weekday_holiday_uses_weekday_slot_and_counter():
    residents = {"Maria": make_resident(2)}
    holidays = {WEEKDAY: {"name": "Christmas", "upper": "Maria", "intern": None}}

    rows, unassigned, _ = _pre_apply_holidays(
        holidays=holidays,
        residents=residents,
        lookup=_StubLookup(),
        block1_end=None,
    )

    assert residents["Maria"]["weekday_calls"] == 1
    assert residents["Maria"]["weekend_calls"] == 0
    assert residents["Maria"]["upper_calls"] == 1
    assert rows[0]["slot"] == SLOT_UPPER_WEEKDAY
    # Blank intern on non-Block-1 weekday → silently dropped, no unassigned row.
    assert all(r["slot"] != SLOT_INTERN_WEEKDAY for r in rows)
    assert unassigned == []


def test_blank_upper_produces_unassigned_row():
    residents = {"Bob": make_resident(1)}
    holidays = {WEEKEND: {"name": "NYD", "upper": None, "intern": "Bob"}}

    rows, unassigned, _ = _pre_apply_holidays(
        holidays=holidays,
        residents=residents,
        lookup=_StubLookup(),
        block1_end=None,
    )

    upper_rows = [r for r in rows if r["slot"] == SLOT_UPPER_WEEKEND]
    assert len(upper_rows) == 1
    assert upper_rows[0]["resident"] == ""
    assert len(unassigned) == 1
    assert unassigned[0]["reasons"] == "holiday_manual_assignment"
    assert unassigned[0]["holiday"] == "NYD"


def test_unknown_name_raises():
    residents = {"Maria": make_resident(2)}
    holidays = {WEEKEND: {"name": "X", "upper": "Ghost", "intern": None}}

    with pytest.raises(ValueError, match="unknown resident"):
        _pre_apply_holidays(
            holidays=holidays,
            residents=residents,
            lookup=_StubLookup(),
            block1_end=None,
        )


def test_intern_in_upper_column_raises():
    residents = {"Bob": make_resident(1)}
    holidays = {WEEKEND: {"name": "X", "upper": "Bob", "intern": None}}

    with pytest.raises(ValueError, match="PGY1"):
        _pre_apply_holidays(
            holidays=holidays,
            residents=residents,
            lookup=_StubLookup(),
            block1_end=None,
        )


def test_upper_in_intern_column_raises():
    residents = {"Maria": make_resident(2)}
    holidays = {WEEKEND: {"name": "X", "upper": None, "intern": "Maria"}}

    with pytest.raises(ValueError, match="Interns must be PGY1"):
        _pre_apply_holidays(
            holidays=holidays,
            residents=residents,
            lookup=_StubLookup(),
            block1_end=None,
        )


def test_skip_dates_prevents_double_count():
    residents = {"Maria": make_resident(2)}
    holidays = {WEEKEND: {"name": "X", "upper": "Maria", "intern": None}}

    rows, unassigned, assignments = _pre_apply_holidays(
        holidays=holidays,
        residents=residents,
        lookup=_StubLookup(),
        block1_end=None,
        skip_dates={WEEKEND},
    )

    # Completely skipped — counters untouched, no rows, no assignments.
    assert residents["Maria"]["total_calls"] == 0
    assert rows == []
    assert unassigned == []
    assert assignments == {}


def test_block1_weekday_intern_slot_is_created():
    """When Block 1 has intern weekday calls enabled, a blank intern cell on a
    Block 1 weekday holiday should still emit an unassigned row (vs. being
    silently dropped like non-Block-1 weekdays)."""
    if not sm.INTERN_BLOCK1_WEEKDAY_CALLS:
        pytest.skip("INTERN_BLOCK1_WEEKDAY_CALLS disabled in this config")

    residents = {"Maria": make_resident(2)}
    # Use a weekday inside Block 1. Block 1 starts at ACADEMIC_DATE_START.
    block1_weekday = sm.ACADEMIC_DATE_START
    while block1_weekday.weekday() >= 5:
        block1_weekday = block1_weekday.replace(day=block1_weekday.day + 1)
    block1_end = block1_weekday  # keep this date inside Block 1

    holidays = {block1_weekday: {"name": "X", "upper": "Maria", "intern": None}}

    rows, unassigned, _ = _pre_apply_holidays(
        holidays=holidays,
        residents=residents,
        lookup=_StubLookup(),
        block1_end=block1_end,
    )

    intern_rows = [r for r in rows if r["slot"] == SLOT_INTERN_WEEKDAY]
    assert len(intern_rows) == 1
    assert intern_rows[0]["resident"] == ""
    assert len(unassigned) == 1
