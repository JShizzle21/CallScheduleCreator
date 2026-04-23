"""End-to-end smoke test against the real data files.

Runs `generate_schedule_once(seed=0)` against `data/*.xlsx` exactly as the
production scheduler would, then asserts the high-level invariants every
shipped schedule must satisfy:

  - audit reports zero hard errors
  - every scheduler-produced required slot is filled (unassigned rows
    that come from blank manual-assignment cells in holidays.xlsx are
    expected and reported as warnings, not failures — those cells are
    filled in by hand right before shipping)
  - per-resident call-count gap doesn't exceed MAX_DIFF_HARD

This is intentionally a wide net, not a precise check. It catches
catastrophic regressions — a refactor that breaks slot assignment, a
bad merge that drops a constraint, a config change that makes the
schedule unsolvable. It does NOT validate optimization quality.

NOTE: this test can fail for benign data-state reasons (e.g. holiday
manual assignments that conflict with post-call rest of an existing
auto-assignment). Those failures are real and intended — they tell
you holidays.xlsx needs to be edited before the schedule can ship.
Skip the test locally if your data is mid-edit:

    pytest tests/ --deselect tests/test_end_to_end_smoke.py
"""

from __future__ import annotations

from pathlib import Path

import pytest

import scheduler_main as sm
from config import load_default_config


# Skip the entire module if data files aren't present (e.g. in a fresh
# checkout that hasn't been populated yet). Avoids confusing failures
# unrelated to code regressions.
DATA_DIR = Path(__file__).resolve().parent.parent / "data"
REQUIRED_FILES = ["flow.xlsx", "rotation_rules.xlsx"]
missing = [f for f in REQUIRED_FILES if not (DATA_DIR / f).exists()]
pytestmark = pytest.mark.skipif(
    bool(missing),
    reason=f"Required data files missing: {missing}. "
    f"Smoke test only runs when data/ is populated.",
)


# Use a single fixed seed so the test is deterministic. seed=0 is what the
# Monte Carlo loop tries first, so it's also the "default" run people see.
SEED = 0


@pytest.fixture(scope="module")
def schedule_result():
    """Build the schedule once per test session — generate_schedule_once is
    deterministic in its seed, so all assertions can share one result."""
    config, paths = load_default_config()
    return sm.generate_schedule_once(seed=SEED, config=config, paths=paths)


def test_audit_reports_no_hard_errors(schedule_result):
    """The audit's `errors` list must be empty. Any entry here means a
    hard constraint (no-call day, post-call day, NO_CALL rotation, slot
    count mismatch, PGY3 cutoff, etc.) was violated. Non-empty → bug."""
    errors = schedule_result["audit_data"].get("errors", [])
    assert errors == [], (
        f"Audit reported {len(errors)} hard error(s) in seed={SEED} run. "
        f"First few: {errors[:5]}"
    )


def test_no_scheduler_unassigned_slots(schedule_result):
    """Every required slot the SCHEDULER is responsible for must be filled.

    Rows with reason 'holiday_manual_assignment' come from blank upper/intern
    cells in holidays.xlsx — those are filled in by hand right before
    shipping, so we ignore them here. Any other unassigned row means the
    scheduler ran out of eligible candidates — that's a real bug.
    """
    unassigned = schedule_result.get("unassigned_rows", [])
    scheduler_unassigned = [
        r for r in unassigned if r.get("reasons") != "holiday_manual_assignment"
    ]
    assert scheduler_unassigned == [], (
        f"{len(scheduler_unassigned)} scheduler-produced unassigned slot(s) "
        f"in seed={SEED} run (holiday-manual ignored). "
        f"First few: {scheduler_unassigned[:5]}"
    )

    # Surface the holiday-manual count as an informational message so the
    # test output reminds you what's still pending without failing.
    holiday_pending = [
        r for r in unassigned if r.get("reasons") == "holiday_manual_assignment"
    ]
    if holiday_pending:
        print(
            f"\n[info] {len(holiday_pending)} holiday slot(s) await manual "
            f"assignment in holidays.xlsx — fill these in before shipping."
        )


def test_fairness_gap_within_hard_limit(schedule_result):
    """No fairness metric should exceed MAX_DIFF_HARD. The hard fairness
    flag is the upper bound the scheduler is allowed to leave open;
    crossing it post-swap-pass means the optimization broke."""
    fairness = schedule_result["audit_data"].get("fairness_summary", {})
    diff_keys = [k for k in fairness if k.endswith("_diff")]

    overshoots = {
        k: fairness[k]
        for k in diff_keys
        if isinstance(fairness[k], (int, float)) and fairness[k] > sm.MAX_DIFF_HARD
    }
    assert not overshoots, (
        f"Fairness diff(s) exceed MAX_DIFF_HARD={sm.MAX_DIFF_HARD}: "
        f"{overshoots}. Full fairness_summary: {fairness}"
    )


def test_schedule_has_assignments(schedule_result):
    """Sanity check: a year of scheduling should produce hundreds of rows.
    Catches the degenerate case where the loop exits immediately (e.g.
    academic dates misconfigured so start > end)."""
    rows = schedule_result.get("schedule_rows", [])
    # 365 days × ~2 slots/day on weekends + 1 on weekdays ≈ 500+ rows minimum
    assert len(rows) > 300, (
        f"Only {len(rows)} schedule rows produced — suspiciously low. "
        f"Check ACADEMIC_DATE_START/END configuration."
    )
