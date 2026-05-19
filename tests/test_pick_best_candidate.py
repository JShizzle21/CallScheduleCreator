"""Tests for scheduler_main.pick_best_candidate.

These tests exercise the ranking/tiebreak logic in isolation, without
touching the Excel rotation lookup, validation, or Monte Carlo loop.

The function reads several module-level constants from CONFIG (MAX_DIFF_*,
MIN_SPACING_DAYS_*, weight constants). Tests reference those constants by
name rather than hardcoding values, so they remain valid if config.yaml
is retuned.
"""

from __future__ import annotations

import random
from datetime import date, timedelta

import scheduler_main as sm
from scheduler_main import (
    SLOT_INTERN_WEEKEND,
    SLOT_UPPER_WEEKDAY,
    SLOT_UPPER_WEEKEND,
    pick_best_candidate,
)


# A weekday well inside the academic year, far from year boundaries so
# year_bias is roughly 0.5 and won't dominate.
TEST_DATE = date(sm.ACADEMIC_YEAR_START + 1, 1, 5)


def make_resident(
    pgy: int,
    *,
    weekday_calls: int = 0,
    weekend_calls: int = 0,
    intern_calls: int = 0,
    last_call: date | None = None,
) -> dict:
    """Build a minimal resident dict matching loader.load_residents shape."""
    return {
        "pgy": pgy,
        "assigned_dates": [last_call] if last_call else [],
        "total_calls": 0,
        "weekday_calls": weekday_calls,
        "weekend_calls": weekend_calls,
        "friday_calls": 0,
        "saturday_calls": 0,
        "sunday_calls": 0,
        "upper_calls": 0,
        "intern_calls": intern_calls,
        "Jul_Dec_calls": 0,
        "Jan_Jun_calls": 0,
    }


def eligible_entry(name: str, pref: str = "ELIGIBLE", rotation: str = "WARDS"):
    return (name, pref, rotation)


# ---------------------------------------------------------------------------
# Basic contract
# ---------------------------------------------------------------------------

def test_empty_eligible_returns_none():
    result = pick_best_candidate(
        residents={"A": make_resident(2)},
        eligible=[],
        d=TEST_DATE,
        slot=SLOT_UPPER_WEEKDAY,
        rng=random.Random(0),
    )
    assert result is None


def test_single_candidate_picked_no_tiebreak():
    residents = {"A": make_resident(2, weekday_calls=0)}
    result = pick_best_candidate(
        residents=residents,
        eligible=[eligible_entry("A")],
        d=TEST_DATE,
        slot=SLOT_UPPER_WEEKDAY,
        rng=random.Random(0),
    )
    assert result is not None
    name, rotation, was_tiebreak = result
    assert name == "A"
    assert rotation == "WARDS"
    assert was_tiebreak is False


def test_identical_candidates_set_tiebreak_flag():
    residents = {
        "A": make_resident(2, weekday_calls=0),
        "B": make_resident(2, weekday_calls=0),
    }
    result = pick_best_candidate(
        residents=residents,
        eligible=[eligible_entry("A"), eligible_entry("B")],
        d=TEST_DATE,
        slot=SLOT_UPPER_WEEKDAY,
        rng=random.Random(0),
    )
    assert result is not None
    name, _, was_tiebreak = result
    assert was_tiebreak is True
    assert name in {"A", "B"}


def test_rng_is_deterministic_under_tie():
    """Same seed → same pick when candidates are tied."""
    residents = {
        "A": make_resident(2, weekday_calls=0),
        "B": make_resident(2, weekday_calls=0),
    }
    eligible = [eligible_entry("A"), eligible_entry("B")]

    pick1 = pick_best_candidate(residents, eligible, TEST_DATE, SLOT_UPPER_WEEKDAY, random.Random(42))
    pick2 = pick_best_candidate(residents, eligible, TEST_DATE, SLOT_UPPER_WEEKDAY, random.Random(42))
    assert pick1[0] == pick2[0]


# ---------------------------------------------------------------------------
# Lexicographic gates: hard_diff_flag > soft_diff_flag > weighted_score
# ---------------------------------------------------------------------------

def test_hard_diff_flag_dominates_lower_weighted_score():
    """A candidate that trips hard_diff is dispreferred even if their
    weighted score is lower than a non-hard-diff candidate."""
    pool_min = 0
    hard_gap = sm.MAX_DIFF_HARD + 1   # trips hard
    soft_gap = sm.MAX_DIFF_SOFT + 1   # trips soft only

    residents = {
        # Pool-min anchor (not eligible, just sets min_in_pool=0)
        "ANCHOR": make_resident(2, weekday_calls=pool_min),
        # A: hard-flag, but otherwise pristine (good spacing, no avoid)
        "A": make_resident(2, weekday_calls=hard_gap, last_call=TEST_DATE - timedelta(days=30)),
        # B: soft-flag only, bigger weighted score than A would naively imply,
        # but lex-dominated by hard flag.
        "B": make_resident(2, weekday_calls=soft_gap, last_call=TEST_DATE - timedelta(days=30)),
    }
    result = pick_best_candidate(
        residents=residents,
        eligible=[eligible_entry("A"), eligible_entry("B")],
        d=TEST_DATE,
        slot=SLOT_UPPER_WEEKDAY,
        rng=random.Random(0),
    )
    assert result[0] == "B"


def test_soft_diff_flag_dominates_when_hard_equal():
    """When hard_diff_flag is tied at 0, soft_diff_flag breaks rank
    before weighted_score is consulted."""
    residents = {
        "ANCHOR": make_resident(2, weekday_calls=0),
        # A: trips soft only
        "A": make_resident(2, weekday_calls=sm.MAX_DIFF_SOFT + 1,
                           last_call=TEST_DATE - timedelta(days=30)),
        # B: no soft trip, but slightly behind on calls
        "B": make_resident(2, weekday_calls=sm.MAX_DIFF_SOFT,
                           last_call=TEST_DATE - timedelta(days=30)),
    }
    result = pick_best_candidate(
        residents=residents,
        eligible=[eligible_entry("A"), eligible_entry("B")],
        d=TEST_DATE,
        slot=SLOT_UPPER_WEEKDAY,
        rng=random.Random(0),
    )
    assert result[0] == "B"


def test_lower_fairness_gap_wins_in_weighted_score():
    """When both gates are tied, the weighted score breaks rank.
    Holding spacing/avoid/year_bias equal, lower fairness_gap wins."""
    residents = {
        "A": make_resident(2, weekday_calls=0,
                           last_call=TEST_DATE - timedelta(days=30)),
        "B": make_resident(2, weekday_calls=2,
                           last_call=TEST_DATE - timedelta(days=30)),
    }
    result = pick_best_candidate(
        residents=residents,
        eligible=[eligible_entry("A"), eligible_entry("B")],
        d=TEST_DATE,
        slot=SLOT_UPPER_WEEKDAY,
        rng=random.Random(0),
    )
    assert result[0] == "A"


def test_better_spacing_wins_when_fairness_equal():
    """Same fairness_gap, same avoid, same pgy → spacing tier breaks tie."""
    residents = {
        # A: violates strong spacing
        "A": make_resident(2, weekday_calls=0,
                           last_call=TEST_DATE - timedelta(days=sm.MIN_SPACING_DAYS_STRONG - 1)),
        # B: well spaced
        "B": make_resident(2, weekday_calls=0,
                           last_call=TEST_DATE - timedelta(days=sm.MIN_SPACING_DAYS_MILD + 5)),
    }
    result = pick_best_candidate(
        residents=residents,
        eligible=[eligible_entry("A"), eligible_entry("B")],
        d=TEST_DATE,
        slot=SLOT_UPPER_WEEKDAY,
        rng=random.Random(0),
    )
    assert result[0] == "B"


def test_avoid_pref_dispreferred():
    """Identical state otherwise → ELIGIBLE preferred over AVOID."""
    residents = {
        "A": make_resident(2, weekday_calls=0,
                           last_call=TEST_DATE - timedelta(days=30)),
        "B": make_resident(2, weekday_calls=0,
                           last_call=TEST_DATE - timedelta(days=30)),
    }
    result = pick_best_candidate(
        residents=residents,
        eligible=[
            eligible_entry("A", pref="AVOID"),
            eligible_entry("B", pref="ELIGIBLE"),
        ],
        d=TEST_DATE,
        slot=SLOT_UPPER_WEEKDAY,
        rng=random.Random(0),
    )
    assert result[0] == "B"


# ---------------------------------------------------------------------------
# Pool / counter selection by slot
# ---------------------------------------------------------------------------

def test_intern_weekend_uses_pgy1_pool_and_intern_calls_counter():
    """Intern slot must rank against the PGY1 pool only, and use
    intern_calls (total intern calls, weekday + weekend) for fairness_gap
    so Block 1 weekday calls are factored in alongside weekend calls."""
    residents = {
        # I1: fewer total intern_calls but more weekday_calls — intern_calls
        #     is what matters, weekday_calls should be ignored.
        "I1": make_resident(1, intern_calls=0, weekday_calls=10,
                            last_call=TEST_DATE - timedelta(days=30)),
        "I2": make_resident(1, intern_calls=2, weekday_calls=0,
                            last_call=TEST_DATE - timedelta(days=30)),
        # An upper resident in the dict — must not affect the intern pool min.
        "U1": make_resident(2, weekend_calls=0,
                            last_call=TEST_DATE - timedelta(days=30)),
    }
    result = pick_best_candidate(
        residents=residents,
        eligible=[eligible_entry("I1"), eligible_entry("I2")],
        d=TEST_DATE,
        slot=SLOT_INTERN_WEEKEND,
        rng=random.Random(0),
    )
    # I1 has fewer intern_calls, so wins regardless of weekday_calls.
    assert result[0] == "I1"


def test_upper_weekday_ignores_pgy1_pool():
    """Upper weekday slot pool is PGY 2/3 only — a PGY1 with 0 calls
    must NOT pull min_in_pool down."""
    residents = {
        "I1": make_resident(1, weekday_calls=0),  # ignored by pool
        "U1": make_resident(2, weekday_calls=5,
                            last_call=TEST_DATE - timedelta(days=30)),
        "U2": make_resident(2, weekday_calls=5,
                            last_call=TEST_DATE - timedelta(days=30)),
    }
    # If PGY1 were (incorrectly) in the pool, min_in_pool=0 and gap=5 → both
    # would trip hard_diff_flag. With the correct pool, min=5, gap=0 → no flag.
    result = pick_best_candidate(
        residents=residents,
        eligible=[eligible_entry("U1"), eligible_entry("U2")],
        d=TEST_DATE,
        slot=SLOT_UPPER_WEEKDAY,
        rng=random.Random(0),
    )
    name, _, was_tiebreak = result
    assert name in {"U1", "U2"}
    assert was_tiebreak is True  # confirms neither tripped a gate
