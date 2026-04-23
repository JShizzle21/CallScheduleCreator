"""Tests for validation.audit_schedule no-call-day enforcement.

The scheduler's eligibility filter is supposed to prevent any resident from
being scheduled on a date listed in their no_call_days. The audit is the
backstop that flags it if a regression ever lets one through. These tests
exercise the audit directly with a hand-built schedule containing a known
violation, so we know the check itself works regardless of which schedule
the Monte Carlo loop happens to produce.

Holiday-named residents are intentionally exempt — holidays.xlsx is treated
as a manual override that wins over the resident's own no_call_days. We
lock that exemption in too.
"""

from __future__ import annotations

from datetime import date, timedelta

import validation
from validation import audit_schedule


# Pick a Saturday well inside the academic year so the audit's per-day
# slot-count loop processes it as a normal weekend day. Using
# validation.ACADEMIC_DATE_START guarantees we're inside the configured
# year regardless of what config.yaml currently holds.
def _first_saturday_in_year() -> date:
    d = validation.ACADEMIC_DATE_START + timedelta(days=30)
    while d.weekday() != 5:  # Saturday
        d += timedelta(days=1)
    return d


WEEKEND = _first_saturday_in_year()


class _StubLookup:
    """Minimal lookup — audit calls .blocks and .rotation_on_date()."""

    blocks = []  # empty so block1_end is None and Block-1-intern logic is off
    skip_rows: set = set()

    def rotation_on_date(self, name, d):
        return "WARDS"

    def rotation_segments_for_resident(self, name):
        # Audit's rotation_date_summary builder calls this; an empty
        # iterable is fine for our purposes (we only assert on errors).
        return []


def _make_resident(pgy: int) -> dict:
    return {
        "pgy": pgy,
        "assigned_dates": [],
        "total_calls": 0,
        "weekday_calls": 0,
        "weekend_calls": 0,
        "upper_calls": 0,
        "intern_calls": 0,
        "Jul_Dec_calls": 0,
        "Jan_Jun_calls": 0,
    }


def _wards_allowed_rules() -> dict:
    """Rotation rules that ALLOW WARDS for every PGY so 'WARDS' assignments
    don't get flagged as NO_CALL violations — keeps the test isolated to
    the no-call-day check."""
    return {
        ("WARDS", 1): {"preference": "ALLOW"},
        ("WARDS", 2): {"preference": "ALLOW"},
        ("WARDS", 3): {"preference": "ALLOW"},
    }


def _audit(schedule_rows, residents, no_call_days, holidays):
    return audit_schedule(
        schedule_rows=schedule_rows,
        residents=residents,
        lookup=_StubLookup(),
        rules=_wards_allowed_rules(),
        no_call_days=no_call_days,
        unassigned_rows=[],
        holidays=holidays,
        seed=0,
        tiebreaker_count=0,
    )


def test_audit_flags_resident_assigned_on_no_call_day():
    """A resident scheduled on a date in their no_call_days dict must
    produce an error containing 'assigned on no-call day'."""
    residents = {"Maria": _make_resident(2), "Bob": _make_resident(1)}
    no_call_days = {"Maria": {WEEKEND: "vacation"}}

    schedule_rows = [
        {"date": WEEKEND.isoformat(), "slot": "UPPER_WEEKEND",
         "resident": "Maria", "note": ""},
        {"date": WEEKEND.isoformat(), "slot": "INTERN_WEEKEND",
         "resident": "Bob", "note": ""},
    ]

    result = _audit(schedule_rows, residents, no_call_days, holidays={})
    expected = f"{WEEKEND}: Maria assigned on no-call day"
    assert expected in result["errors"], (
        f"Expected no-call violation in errors. Got: {result['errors']}"
    )


def test_audit_does_not_flag_resident_with_unrelated_no_call_days():
    """Residents whose no_call_days don't include the assignment date
    must NOT produce a no-call violation."""
    residents = {"Maria": _make_resident(2), "Bob": _make_resident(1)}
    other_day = WEEKEND + timedelta(days=14)
    no_call_days = {"Maria": {other_day: "vacation"}}

    schedule_rows = [
        {"date": WEEKEND.isoformat(), "slot": "UPPER_WEEKEND",
         "resident": "Maria", "note": ""},
        {"date": WEEKEND.isoformat(), "slot": "INTERN_WEEKEND",
         "resident": "Bob", "note": ""},
    ]

    result = _audit(schedule_rows, residents, no_call_days, holidays={})
    no_call_errors = [e for e in result["errors"] if "no-call day" in e]
    assert no_call_errors == [], (
        f"Did not expect any no-call errors. Got: {no_call_errors}"
    )


def test_holiday_override_exempts_no_call_day():
    """Per design, holidays.xlsx assignments override a resident's
    no_call_days. A row with note='HOLIDAY: ...' on a no-call date must
    NOT trigger the no-call violation."""
    residents = {"Maria": _make_resident(2), "Bob": _make_resident(1)}
    no_call_days = {"Maria": {WEEKEND: "vacation"}}
    holidays = {WEEKEND: {"name": "Christmas",
                          "upper": "Maria", "intern": "Bob"}}

    schedule_rows = [
        {"date": WEEKEND.isoformat(), "slot": "UPPER_WEEKEND",
         "resident": "Maria", "note": "HOLIDAY: Christmas"},
        {"date": WEEKEND.isoformat(), "slot": "INTERN_WEEKEND",
         "resident": "Bob", "note": "HOLIDAY: Christmas"},
    ]

    result = _audit(schedule_rows, residents, no_call_days, holidays=holidays)
    no_call_errors = [e for e in result["errors"] if "no-call day" in e]
    assert no_call_errors == [], (
        f"Holiday override should exempt no-call check. Got: {no_call_errors}"
    )


def test_completed_row_exempts_no_call_day():
    """COMPLETED rows are accepted as ground truth — the audit must skip
    constraint checks (including no-call) for them so partial-year
    restarts don't re-flag pre-existing data the user can't change."""
    residents = {"Maria": _make_resident(2), "Bob": _make_resident(1)}
    no_call_days = {"Maria": {WEEKEND: "vacation"}}

    schedule_rows = [
        {"date": WEEKEND.isoformat(), "slot": "UPPER_WEEKEND",
         "resident": "Maria", "note": "COMPLETED"},
        {"date": WEEKEND.isoformat(), "slot": "INTERN_WEEKEND",
         "resident": "Bob", "note": ""},
    ]

    result = _audit(schedule_rows, residents, no_call_days, holidays={})
    no_call_errors = [e for e in result["errors"] if "no-call day" in e]
    assert no_call_errors == [], (
        f"COMPLETED rows should bypass no-call check. Got: {no_call_errors}"
    )
