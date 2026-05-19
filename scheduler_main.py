from __future__ import annotations

import bisect
import concurrent.futures
import logging
import os
import random
import sys
import threading
import time
from datetime import date, datetime, timedelta
from functools import partial
from typing import Callable, Dict, List, Tuple, Optional

# All internal modules live in src/ — keep them off the project root to reduce
# clutter for end users. Add src/ to the import path so existing flat-style
# imports (`from config import X`) keep working without per-file changes.
sys.path.insert(0, os.path.join(os.path.dirname(os.path.abspath(__file__)), "src"))

from config import CONFIG, load_default_config, legacy_gui_config_warning
from data_bundle import DataBundle, load_data_bundle
from errors import ConfigError, DataValidationError, ScheduleError
from exports import write_call_totals_xlsx, write_call_schedule_xlsx, write_audit
from loader import load_completed_calls
import validation
from validation import validate_rotations_against_rules, validate_no_call_days, audit_schedule

logger = logging.getLogger(__name__)

# Slot constants — truly static, not user-configurable.
SLOT_UPPER_WEEKDAY = "UPPER_WEEKDAY"
SLOT_UPPER_WEEKEND = "UPPER_WEEKEND"
SLOT_INTERN_WEEKEND = "INTERN_WEEKEND"
SLOT_INTERN_WEEKDAY = "INTERN_WEEKDAY"

VALID_MONTE_CARLO_SCORE_KEYS = {
    "errors",
    "unassigned",
    "upper_weekend_diff",
    "upper_weekday_diff",
    "upper_total_diff",
    "pgy2_total_diff",
    "pgy3_total_diff",
    "intern_weekend_diff",
    "avoid_assignments",
    "warnings",
}

VALID_PICK_CANDIDATE_RANK_KEYS = {
    "hard_diff_flag",
    "soft_diff_flag",
    "weighted_score",
}

_DEFAULT_MONTE_CARLO_SCORE_ORDER = [
    "errors",
    "unassigned",
    "upper_weekend_diff",
    "upper_weekday_diff",
    "upper_total_diff",
    "intern_weekend_diff",
    "avoid_assignments",
    "warnings",
]

_DEFAULT_PICK_CANDIDATE_RANK_ORDER = [
    "hard_diff_flag",
    "soft_diff_flag",
    "weighted_score",
]


# Module-level constants below are populated by _apply_config(). They are
# declared here with sentinel values so the module can be imported without a
# config dict — _apply_config(CONFIG) is called at the bottom of this module
# to populate them from config.yaml for backward compat with code and tests
# that reference e.g. scheduler_main.FAIRNESS_GAP_WEIGHT directly.
#
# generate_schedule_once / run_simulation call _apply_config again with the
# caller-supplied config so GUI overrides take effect per run.
POST_CALL_DAYS: int = 0
SIMULATION_RUNS: int = 0
ACADEMIC_DATE_START_STRING: str = ""
ACADEMIC_DATE_END_STRING: str = ""
ACADEMIC_DATE_START: date = date(1970, 1, 1)
ACADEMIC_DATE_END: date = date(1970, 1, 1)
ACADEMIC_YEAR_START: int = 1970
TOTAL_YEAR_DAYS: int = 0
FIRST_HALF_END: date = date(1970, 12, 31)
MIN_SPACING_DAYS_STRONG: int = 0
MIN_SPACING_DAYS_MILD: int = 0
MAX_CALLS_IN_WINDOW: int = 0
ROLLING_WINDOW_DAYS: int = 0
MAX_DIFF_SOFT: int = 0
MAX_DIFF_HARD: int = 0
FAIRNESS_GAP_WEIGHT: float = 0.0
SPACING_WEIGHT: float = 0.0
AVOID_WEIGHT: float = 0.0
YEAR_BIAS_WEIGHT: float = 0.0
PACE_WEIGHT: float = 0.0
LOOKAHEAD_WEIGHT: float = 0.0
INTERN_BLOCK1_WEEKDAY_CALLS: int = 0
PGY3_CUTOFF_DATE: Optional[date] = None
MONTE_CARLO_SCORE_ORDER: List[str] = list(_DEFAULT_MONTE_CARLO_SCORE_ORDER)
PICK_CANDIDATE_RANK_ORDER: List[str] = list(_DEFAULT_PICK_CANDIDATE_RANK_ORDER)


def _apply_config(config: dict) -> None:
    """Populate module-level scheduler constants from `config`.

    Mutates this module's global namespace. Callers pass the config dict
    returned by `load_default_config()` (CLI) or built from `st.session_state`
    (GUI). The function is idempotent — calling it twice with the same config
    is a no-op.

    Rationale: the scheduler has ~25 tuning parameters threaded through many
    small helpers (spacing, window cap, year bias, weights). Passing `config`
    through every signature would be a large, noisy change for no real
    benefit since a single simulation run uses one config. Instead, we apply
    config once per run and let the helpers read module globals.
    """
    global POST_CALL_DAYS, SIMULATION_RUNS
    global ACADEMIC_DATE_START_STRING, ACADEMIC_DATE_END_STRING
    global ACADEMIC_DATE_START, ACADEMIC_DATE_END
    global ACADEMIC_YEAR_START, TOTAL_YEAR_DAYS, FIRST_HALF_END
    global MIN_SPACING_DAYS_STRONG, MIN_SPACING_DAYS_MILD
    global MAX_CALLS_IN_WINDOW, ROLLING_WINDOW_DAYS
    global MAX_DIFF_SOFT, MAX_DIFF_HARD
    global FAIRNESS_GAP_WEIGHT, SPACING_WEIGHT, AVOID_WEIGHT
    global YEAR_BIAS_WEIGHT, PACE_WEIGHT, LOOKAHEAD_WEIGHT
    global INTERN_BLOCK1_WEEKDAY_CALLS, PGY3_CUTOFF_DATE
    global MONTE_CARLO_SCORE_ORDER, PICK_CANDIDATE_RANK_ORDER

    POST_CALL_DAYS = int(config["POST_CALL_DAYS"])
    SIMULATION_RUNS = int(config.get("SIMULATION_RUNS", 1000))

    ACADEMIC_DATE_START_STRING = config["ACADEMIC_DATE_START_STRING"]
    ACADEMIC_DATE_END_STRING = config["ACADEMIC_DATE_END_STRING"]
    ACADEMIC_DATE_START = datetime.strptime(ACADEMIC_DATE_START_STRING, "%Y-%m-%d").date()
    ACADEMIC_DATE_END = datetime.strptime(ACADEMIC_DATE_END_STRING, "%Y-%m-%d").date()
    ACADEMIC_YEAR_START = ACADEMIC_DATE_START.year
    TOTAL_YEAR_DAYS = (ACADEMIC_DATE_END - ACADEMIC_DATE_START).days
    FIRST_HALF_END = date(ACADEMIC_YEAR_START, 12, 31)

    MIN_SPACING_DAYS_STRONG = int(config["MIN_SPACING_DAYS_STRONG"])
    MIN_SPACING_DAYS_MILD = int(config["MIN_SPACING_DAYS_MILD"])
    MAX_CALLS_IN_WINDOW = int(config.get("MAX_CALLS_IN_WINDOW", 0))
    ROLLING_WINDOW_DAYS = int(config.get("ROLLING_WINDOW_DAYS", 14))
    MAX_DIFF_SOFT = int(config["MAX_DIFF_SOFT"])
    MAX_DIFF_HARD = int(config["MAX_DIFF_HARD"])

    FAIRNESS_GAP_WEIGHT = float(config["FAIRNESS_GAP_WEIGHT"])
    SPACING_WEIGHT = float(config["SPACING_WEIGHT"])
    AVOID_WEIGHT = float(config["AVOID_WEIGHT"])
    YEAR_BIAS_WEIGHT = float(config["YEAR_BIAS_WEIGHT"])
    PACE_WEIGHT = float(config.get("PACE_WEIGHT", config.get("FUTURE_AVAIL_WEIGHT", 1.0)))
    LOOKAHEAD_WEIGHT = float(config.get("LOOKAHEAD_WEIGHT", 1.0))

    INTERN_BLOCK1_WEEKDAY_CALLS = int(config.get("INTERN_BLOCK1_WEEKDAY_CALLS", 0))

    # PGY3 graduation cutoff — inclusive: PGY3s are excluded on this date and
    # after. If omitted, blank, or outside the academic year range the cutoff
    # is disabled.
    _pgy3_cutoff_raw = config.get("PGY3_CUTOFF_DATE", "")
    PGY3_CUTOFF_DATE = None
    if _pgy3_cutoff_raw and str(_pgy3_cutoff_raw).strip():
        try:
            _parsed_cutoff = datetime.strptime(str(_pgy3_cutoff_raw).strip(), "%Y-%m-%d").date()
            if ACADEMIC_DATE_START <= _parsed_cutoff <= ACADEMIC_DATE_END:
                PGY3_CUTOFF_DATE = _parsed_cutoff
            else:
                logger.warning(
                    "WARNING: PGY3_CUTOFF_DATE (%s) is outside the academic year "
                    "(%s – %s). Cutoff disabled; PGY3s will be scheduled through "
                    "the end of the academic year.",
                    _parsed_cutoff, ACADEMIC_DATE_START, ACADEMIC_DATE_END,
                )
        except ValueError:
            logger.warning(
                "WARNING: PGY3_CUTOFF_DATE '%s' is not a valid date (expected "
                "YYYY-MM-DD). Cutoff disabled; PGY3s will be scheduled through "
                "the end of the academic year.",
                _pgy3_cutoff_raw,
            )

    MONTE_CARLO_SCORE_ORDER = list(
        config.get("MONTE_CARLO_SCORE_ORDER", _DEFAULT_MONTE_CARLO_SCORE_ORDER)
    )
    invalid_score_keys = [
        k for k in MONTE_CARLO_SCORE_ORDER if k not in VALID_MONTE_CARLO_SCORE_KEYS
    ]
    if invalid_score_keys:
        raise ConfigError(
            f"Invalid MONTE_CARLO_SCORE_ORDER entries: {invalid_score_keys}. "
            f"Valid options are: {sorted(VALID_MONTE_CARLO_SCORE_KEYS)}"
        )

    # Refresh validation.py's module-level constants too — the audit reads
    # PGY3_CUTOFF_DATE etc. from validation's own globals, so without this
    # call GUI config overrides are silently ignored at audit time.
    validation._apply_config(config)

    PICK_CANDIDATE_RANK_ORDER = list(
        config.get("PICK_CANDIDATE_RANK_ORDER", _DEFAULT_PICK_CANDIDATE_RANK_ORDER)
    )
    invalid_pick_rank_keys = [
        k for k in PICK_CANDIDATE_RANK_ORDER
        if k not in VALID_PICK_CANDIDATE_RANK_KEYS
    ]
    if invalid_pick_rank_keys:
        raise ConfigError(
            f"Invalid PICK_CANDIDATE_RANK_ORDER entries: {invalid_pick_rank_keys}. "
            f"Valid options are: {sorted(VALID_PICK_CANDIDATE_RANK_KEYS)}"
        )


# Populate module globals from config.yaml at import time. Backward-compat
# shim — keeps `from scheduler_main import X` and `sm.X` references working
# without requiring callers to invoke _apply_config manually.
_apply_config(CONFIG)





def is_weekend(d: date) -> bool:
    return d.weekday() >= 5


def required_slots(d: date, block1_end: Optional[date] = None) -> List[str]:
    if is_weekend(d):
        return [SLOT_INTERN_WEEKEND, SLOT_UPPER_WEEKEND]
    if INTERN_BLOCK1_WEEKDAY_CALLS and block1_end is not None and d <= block1_end:
        return [SLOT_UPPER_WEEKDAY, SLOT_INTERN_WEEKDAY]
    return [SLOT_UPPER_WEEKDAY]


def year_progress(d: date) -> float:
    return (d - ACADEMIC_DATE_START).days / TOTAL_YEAR_DAYS


def days_since_last_call(resident_data: dict, d: date) -> int:
    if not resident_data["assigned_dates"]:
        return 9999
    last = max(resident_data["assigned_dates"])
    return (d - last).days


def is_post_call(resident_data, d):
    """True if d is within POST_CALL_DAYS of any existing assignment, in
    either direction.

    Forward direction matters because `assigned_dates` may already contain
    dates the greedy day loop hasn't processed yet — pre-applied holiday
    assignments and (in partial-year mode) completed calls. Without the
    forward check, the loop would happily place a candidate two days
    before their pre-assigned holiday call, producing a schedule the
    audit then has to flag.
    """
    for i in range(1, POST_CALL_DAYS + 1):
        if (d - timedelta(days=i)) in resident_data["assigned_dates"]:
            return True
        if (d + timedelta(days=i)) in resident_data["assigned_dates"]:
            return True
    return False


def would_exceed_window_cap(assigned_dates, d: date) -> bool:
    """Return True if adding d to assigned_dates would create any rolling
    window of ROLLING_WINDOW_DAYS consecutive days containing more than
    MAX_CALLS_IN_WINDOW call assignments.

    Checks every window that contains d (there are ROLLING_WINDOW_DAYS such
    windows, with start dates ranging from d-(W-1) to d). If any such window
    would end up with more than the cap, return True.

    Set MAX_CALLS_IN_WINDOW to 0 in config.yaml to disable this check.
    """
    if MAX_CALLS_IN_WINDOW <= 0 or ROLLING_WINDOW_DAYS <= 0:
        return False

    already_present = d in assigned_dates
    for offset in range(ROLLING_WINDOW_DAYS):
        start = d - timedelta(days=ROLLING_WINDOW_DAYS - 1 - offset)
        end = start + timedelta(days=ROLLING_WINDOW_DAYS - 1)
        count = sum(1 for ad in assigned_dates if start <= ad <= end)
        if not already_present:
            count += 1
        if count > MAX_CALLS_IN_WINDOW:
            return True
    return False


_SLOT_COUNTER_KEY: Dict[str, str] = {
    SLOT_UPPER_WEEKDAY: "weekday_calls",
    SLOT_UPPER_WEEKEND: "weekend_calls",
    SLOT_INTERN_WEEKEND: "intern_calls",
    SLOT_INTERN_WEEKDAY: "intern_calls",
}


def _static_pool_for_slot(
    lookup,
    residents: Dict[str, dict],
    rules: Dict[Tuple[str, int], str],
    no_call_days: Dict[str, dict],
    d: date,
    slot: str,
    intern_names,
    upper_names,
) -> List[str]:
    """Residents statically eligible for `slot` on day `d`.

    Same checks as eligible_for_slot minus dynamic exclusions (post_call,
    window_cap): this is "who *could* take this call ignoring recent
    assignment history," which is what we need to compute an expected
    share of each day's call load.
    """
    names = intern_names if slot in (SLOT_INTERN_WEEKEND, SLOT_INTERN_WEEKDAY) else upper_names
    pool: List[str] = []
    for name in names:
        pgy = residents[name]["pgy"]
        if slot in (SLOT_INTERN_WEEKEND, SLOT_INTERN_WEEKDAY):
            if pgy != 1:
                continue
        else:
            if pgy == 1:
                continue
        # PGY3 graduation cutoff mirrors eligible_for_slot.
        if pgy == 3 and PGY3_CUTOFF_DATE is not None and d >= PGY3_CUTOFF_DATE:
            continue
        if no_call_days.get(name, {}).get(d) is not None:
            continue
        rotation = lookup.rotation_on_date(name, d)
        if rotation is None:
            continue
        pref = rules.get((rotation, pgy))
        if pref is None or pref == "NO_CALL":
            continue
        pool.append(name)
    return pool


def _precompute_expected_calls(
    lookup,
    residents: Dict[str, dict],
    rules: Dict[Tuple[str, int], str],
    no_call_days: Dict[str, dict],
    holidays: Dict[date, dict],
    intern_names,
    upper_names,
    block1_end: Optional[date],
    holiday_assignments: Optional[Dict[Tuple[date, str], str]] = None,
) -> Dict[str, Dict[str, Dict[date, float]]]:
    """Per-resident cumulative expected-call counts for each counter type.

    For every non-holiday day, each required slot contributes 1/|pool|
    to each statically-eligible resident's counter, where pool is the
    static-eligibility pool for that slot.

    For holiday days with a manual assignment, the assigned resident
    receives +1.0 to the relevant counter (and only them). This keeps
    pace_value neutral on holiday days — their actual counter is also +1
    via apply_assignment — so pacing neither rewards nor penalises the
    holiday-assigned resident after the fact.

    The returned structure gives the cumulative expected value **as of the
    start of day d** (i.e., before day-d assignments are made), so it aligns
    directly with residents[name][counter_key] inside pick_best_candidate.
    """
    holiday_assignments = holiday_assignments or {}
    counter_keys = ("weekday_calls", "weekend_calls", "intern_calls")
    expected_cum: Dict[str, Dict[str, Dict[date, float]]] = {
        name: {ck: {} for ck in counter_keys} for name in residents
    }
    running: Dict[str, Dict[str, float]] = {
        name: {ck: 0.0 for ck in counter_keys} for name in residents
    }

    d = ACADEMIC_DATE_START
    while d <= ACADEMIC_DATE_END:
        # Snapshot BEFORE adding today's share: expected_cum[...][d] reflects
        # the state at the start of day d, matching residents[name][ck] when
        # pick_best_candidate is called for day d.
        for name in residents:
            for ck in counter_keys:
                expected_cum[name][ck][d] = running[name][ck]

        if d in holidays:
            # Manual holiday assignment: +1 expected for the named resident on
            # the counter that apply_assignment bumped. Other pool members get
            # nothing (they're not competing for this day).
            for slot in (SLOT_UPPER_WEEKDAY, SLOT_UPPER_WEEKEND,
                         SLOT_INTERN_WEEKDAY, SLOT_INTERN_WEEKEND):
                name = holiday_assignments.get((d, slot))
                if name is not None:
                    running[name][_SLOT_COUNTER_KEY[slot]] += 1.0
        else:
            for slot in required_slots(d, block1_end=block1_end):
                ck = _SLOT_COUNTER_KEY[slot]
                pool = _static_pool_for_slot(
                    lookup, residents, rules, no_call_days, d, slot,
                    intern_names, upper_names,
                )
                if pool:
                    share = 1.0 / len(pool)
                    for name in pool:
                        running[name][ck] += share

        d += timedelta(days=1)

    return expected_cum


def _precompute_eligible_dates(
    lookup,
    residents: Dict[str, dict],
    rules: Dict[Tuple[str, int], str],
    no_call_days: Dict[str, dict],
    holidays: Dict[date, dict],
    block1_end: Optional[date],
) -> Dict[str, list]:
    """Per-resident sorted list of dates on which the resident is statically
    eligible for any slot they can serve (holiday-free, rotation ≠ NO_CALL,
    no_call_days clean, slot PGY matches resident PGY).

    Used by the lookahead component of the weighted score: the count of
    eligible dates still ahead of `d` measures how much runway the resident
    has to spread their remaining calls over. Low runway → prefer now.
    """
    eligible_dates: Dict[str, list] = {name: [] for name in residents}
    d = ACADEMIC_DATE_START
    while d <= ACADEMIC_DATE_END:
        if d not in holidays:
            slots = required_slots(d, block1_end=block1_end)
            for name, data in residents.items():
                pgy = data["pgy"]
                # PGY3 graduation cutoff: no eligible dates on/after cutoff.
                if pgy == 3 and PGY3_CUTOFF_DATE is not None and d >= PGY3_CUTOFF_DATE:
                    continue
                if no_call_days.get(name, {}).get(d) is not None:
                    continue
                rotation = lookup.rotation_on_date(name, d)
                if rotation is None:
                    continue
                pref = rules.get((rotation, pgy))
                if pref is None or pref == "NO_CALL":
                    continue
                # Only count the day if at least one required slot matches this
                # resident's PGY — e.g. a PGY1 isn't "eligible" on a weekday in
                # Block 2+ because only UPPER_WEEKDAY is required.
                for slot in slots:
                    if slot in (SLOT_INTERN_WEEKEND, SLOT_INTERN_WEEKDAY):
                        if pgy == 1:
                            eligible_dates[name].append(d)
                            break
                    else:
                        if pgy in (2, 3):
                            eligible_dates[name].append(d)
                            break
        d += timedelta(days=1)
    return eligible_dates


def _compute_weighted_score(
    fairness_gap: int,
    spacing_value: int,
    avoid_value: int,
    year_value: float,
    pace_value: float = 0.0,
    lookahead_value: float = 0.0,
) -> float:
    """Combine ranking components into a single scalar.

    Each raw component lives on a different scale (fairness_gap is an
    unbounded int, spacing_value is {0,1,2}, avoid/year/pace are already
    in [0,1]). We normalize each to roughly [0, 1] so the *_WEIGHT
    constants express genuine relative importance.

    fairness_gap is clipped at MAX_DIFF_SOFT: above that threshold the
    soft/hard lexicographic gates have already separated candidates, so
    losing resolution there is acceptable.

    pace_value: two-sided [0, 1] pacing signal. 0 = far behind expected
    pace (preferred now), 0.5 = on pace (neutral), 1 = far ahead of pace
    (deprioritized). Expected pace is the resident's share of every day's
    statically-eligible pool, accumulated through the current date.
    Defaults to 0.0 (no contribution) for callers that don't pass pacing
    data — pick_best_candidate always passes a computed value.

    lookahead_value: [0, 1] ratio of the resident's remaining eligible
    days through year-end vs. the pool maximum. 0 = no runway left
    (strongly preferred now), 1 = most runway in the pool (can be
    deferred). Anticipatory: pushes residents with shrinking future
    availability forward before fairness has to catch up.
    """
    if MAX_DIFF_SOFT > 0:
        fairness_norm = min(fairness_gap, MAX_DIFF_SOFT) / MAX_DIFF_SOFT
    else:
        fairness_norm = 1.0 if fairness_gap > 0 else 0.0

    spacing_norm = spacing_value / 2  # spacing_value ∈ {0, 1, 2}

    return (
        FAIRNESS_GAP_WEIGHT * fairness_norm
        + SPACING_WEIGHT * spacing_norm
        + AVOID_WEIGHT * avoid_value
        + YEAR_BIAS_WEIGHT * year_value
        + PACE_WEIGHT * pace_value
        + LOOKAHEAD_WEIGHT * lookahead_value
    )


def eligible_for_slot(
    lookup,
    residents: Dict[str, dict],
    rules: Dict[Tuple[str, int], str],
    no_call_days: Dict[str, dict],
    d: date,
    slot: str,
    intern_names,
    upper_names,
) -> Tuple[List[Tuple[str, str, str]], Dict[str, int]]:
    eligible: List[Tuple[str, str, str]] = []
    # Residents that pass every check except window_cap. Used as a fallback
    # so the slot can still be filled when the cap would otherwise leave it
    # unassigned (e.g. end-of-year when only a handful of residents are
    # eligible and all have hit the rolling cap simultaneously).
    fallback_capped: List[Tuple[str, str, str]] = []
    reasons: Dict[str, int] = {}

    names = intern_names if slot in (SLOT_INTERN_WEEKEND, SLOT_INTERN_WEEKDAY) else upper_names
    for name in names:
        data = residents[name]
        pgy = data["pgy"]

        if slot in (SLOT_INTERN_WEEKEND, SLOT_INTERN_WEEKDAY):
            if pgy != 1:
                reasons["pgy_mismatch"] = reasons.get("pgy_mismatch", 0) + 1
                continue
        else:
            if pgy == 1:
                reasons["pgy_mismatch"] = reasons.get("pgy_mismatch", 0) + 1
                continue

        # PGY3 graduation cutoff: PGY3s are excluded on/after the cutoff date.
        if pgy == 3 and PGY3_CUTOFF_DATE is not None and d >= PGY3_CUTOFF_DATE:
            reasons["pgy3_graduation_cutoff"] = reasons.get("pgy3_graduation_cutoff", 0) + 1
            continue

        no_call_entry = no_call_days.get(name, {}).get(d)
        if no_call_entry is not None:
            # Use the stored reason so "pre_clinic_day" shows as a distinct
            # bucket from a regular no-call day in the unassigned report.
            reason_key = no_call_entry if no_call_entry else "no_call_day"
            reasons[reason_key] = reasons.get(reason_key, 0) + 1
            continue

        if is_post_call(data, d):
            reasons["post_call"] = reasons.get("post_call", 0) + 1
            continue

        # Check window cap — but don't hard-reject yet; collect separately.
        is_capped = would_exceed_window_cap(data["assigned_dates"], d)

        rotation = lookup.rotation_on_date(name, d)
        if rotation is None:
            reasons["missing_rotation"] = reasons.get("missing_rotation", 0) + 1
            continue

        pref = rules.get((rotation, pgy))
        if pref is None:
            reasons["missing_rule"] = reasons.get("missing_rule", 0) + 1
            continue

        if pref == "NO_CALL":
            reasons["rotation_no_call"] = reasons.get("rotation_no_call", 0) + 1
            continue

        if is_capped:
            fallback_capped.append((name, pref, rotation))
        else:
            eligible.append((name, pref, rotation))

    # Always record how many candidates were window-capped (for audit/reasons).
    if fallback_capped:
        reasons["window_cap"] = reasons.get("window_cap", 0) + len(fallback_capped)

    # Soft fallback: if no non-capped candidate exists but capped ones do,
    # use them rather than leaving the slot unassigned. The rolling-window cap
    # is designed to prevent burst patterns mid-year where non-capped
    # alternatives exist; forcing an unassigned slot when the cap is the sole
    # blocker serves no protective purpose.
    if not eligible and fallback_capped:
        return fallback_capped, reasons

    return eligible, reasons


def pick_best_candidate(
    residents: Dict[str, dict],
    eligible: List[Tuple[str, str, str]],
    d: date,
    slot: str,
    rng: random.Random,
    expected_cum: Optional[Dict[str, Dict[str, Dict[date, float]]]] = None,
    eligible_dates: Optional[Dict[str, list]] = None,
) -> Optional[Tuple[str, str, bool]]:
    if not eligible:
        return None

    if slot in (SLOT_INTERN_WEEKEND, SLOT_INTERN_WEEKDAY):
        counter_key = "intern_calls"
    elif slot == SLOT_UPPER_WEEKEND:
        counter_key = "weekend_calls"
    else:
        counter_key = "weekday_calls"

    if slot in (SLOT_INTERN_WEEKEND, SLOT_INTERN_WEEKDAY):
        pool = [n for n, r in residents.items() if r["pgy"] == 1]
    elif slot in (SLOT_UPPER_WEEKDAY, SLOT_UPPER_WEEKEND):
        pool = [n for n, r in residents.items() if r["pgy"] in (2, 3)]
    else:
        pool = list(residents.keys())

    min_in_pool = min(residents[n][counter_key] for n in pool)
    prog = year_progress(d)

    def spacing_tier(spacing: int) -> int:
        if spacing < MIN_SPACING_DAYS_STRONG:
            return 2
        elif spacing < MIN_SPACING_DAYS_MILD:
            return 1
        return 0

    # PGY3 year_bias uses a shorter year when a graduation cutoff is set:
    # they ramp 0→1 over [ACADEMIC_DATE_START, PGY3_CUTOFF_DATE) instead of
    # the full academic year, so they aren't pushed further into the late
    # dates they can no longer serve.
    if PGY3_CUTOFF_DATE is not None:
        _pgy3_year_days = (PGY3_CUTOFF_DATE - ACADEMIC_DATE_START).days
        pgy3_prog = (d - ACADEMIC_DATE_START).days / _pgy3_year_days if _pgy3_year_days > 0 else 1.0
        pgy3_prog = min(pgy3_prog, 1.0)
    else:
        pgy3_prog = prog

    def year_bias(pgy: int) -> float:
        if slot in (SLOT_INTERN_WEEKEND, SLOT_INTERN_WEEKDAY):
            return 0.0
        if pgy == 3:
            return pgy3_prog
        elif pgy == 2:
            return 1 - prog
        return 0.0

    # Rate-based pacing: compare each candidate's actual counter to the
    # expected cumulative value by date d (their share of every day's static
    # pool so far). Behind pace → preferred; ahead of pace → deprioritized.
    # Normalization scale: MAX_DIFF_SOFT, same threshold used for fairness_gap.
    pace_norm_scale = 2.0 * MAX_DIFF_SOFT if MAX_DIFF_SOFT > 0 else 2.0

    # Look-ahead: remaining eligible days (today through year-end) per
    # candidate, normalized against the pool maximum. Residents with
    # shrinking runway get preferred before fairness has to catch up.
    if eligible_dates is not None:
        remaining_days = {
            name: len(eligible_dates.get(name, []))
                  - bisect.bisect_right(eligible_dates.get(name, []), d)
            for name, _, _ in eligible
        }
        max_remaining = max(remaining_days.values(), default=0)
    else:
        remaining_days = {}
        max_remaining = 0

    ranked_candidates = []

    for name, pref, rotation in eligible:
        data = residents[name]
        pgy = data["pgy"]

        fairness_gap = data[counter_key] - min_in_pool
        spacing = days_since_last_call(data, d)

        hard_diff_flag = 1 if fairness_gap > MAX_DIFF_HARD else 0
        soft_diff_flag = 1 if fairness_gap > MAX_DIFF_SOFT else 0
        spacing_value = spacing_tier(spacing)
        avoid_value = 1 if pref == "AVOID" else 0
        year_value = year_bias(pgy)

        if expected_cum is not None:
            expected = expected_cum.get(name, {}).get(counter_key, {}).get(d, 0.0)
            ahead = data[counter_key] - expected
            # Two-sided: 0 = far behind (preferred), 0.5 = on pace, 1 = far ahead.
            pace_value = max(0.0, min(1.0, (ahead + MAX_DIFF_SOFT) / pace_norm_scale))
        else:
            pace_value = 0.0

        if max_remaining > 0:
            lookahead_value = remaining_days.get(name, 0) / max_remaining
        else:
            lookahead_value = 0.0

        weighted_score = _compute_weighted_score(
            fairness_gap=fairness_gap,
            spacing_value=spacing_value,
            avoid_value=avoid_value,
            year_value=year_value,
            pace_value=pace_value,
            lookahead_value=lookahead_value,
        )

        rank_components = {
            "hard_diff_flag": hard_diff_flag,
            "soft_diff_flag": soft_diff_flag,
            "weighted_score": weighted_score,
        }

        rank = tuple(rank_components[key] for key in PICK_CANDIDATE_RANK_ORDER)
        ranked_candidates.append((rank, name, rotation))

    best_rank = min(rank for rank, _, _ in ranked_candidates)
    best = [(name, rotation) for rank, name, rotation in ranked_candidates if rank == best_rank]

    was_tiebreak = len(best) > 1
    name, rotation = rng.choice(best)
    return name, rotation, was_tiebreak


def apply_assignment(residents: Dict[str, dict], name: str, slot: str, d: date) -> None:
    data = residents[name]
    data["assigned_dates"].append(d)

    data["total_calls"] += 1

    if d <= FIRST_HALF_END:
        data["Jul_Dec_calls"] += 1
    else:
        data["Jan_Jun_calls"] += 1

    if is_weekend(d):
        data["weekend_calls"] += 1
    else:
        data["weekday_calls"] += 1

    # Friday, Saturday, and Sunday counted separately (in addition to
    # weekday/weekend) because they are the least-favoured days and tracked
    # explicitly in the call_totals output.
    if d.weekday() == 4:
        data["friday_calls"] += 1
    elif d.weekday() == 5:
        data["saturday_calls"] += 1
    elif d.weekday() == 6:
        data["sunday_calls"] += 1

    if slot in (SLOT_INTERN_WEEKEND, SLOT_INTERN_WEEKDAY):
        data["intern_calls"] += 1
    else:
        data["upper_calls"] += 1


def _undo_assignment(residents: Dict[str, dict], name: str, slot: str, d: date) -> None:
    """Mirror of apply_assignment — remove one call and decrement all counters."""
    data = residents[name]
    data["assigned_dates"].remove(d)  # list.remove: O(n) but safe (each day assigned once)

    data["total_calls"] -= 1

    if d <= FIRST_HALF_END:
        data["Jul_Dec_calls"] -= 1
    else:
        data["Jan_Jun_calls"] -= 1

    if is_weekend(d):
        data["weekend_calls"] -= 1
    else:
        data["weekday_calls"] -= 1

    # Mirror the friday/saturday/sunday increments in apply_assignment.
    if d.weekday() == 4:
        data["friday_calls"] -= 1
    elif d.weekday() == 5:
        data["saturday_calls"] -= 1
    elif d.weekday() == 6:
        data["sunday_calls"] -= 1

    if slot in (SLOT_INTERN_WEEKEND, SLOT_INTERN_WEEKDAY):
        data["intern_calls"] -= 1
    else:
        data["upper_calls"] -= 1


def local_swap_pass(
    schedule_rows: list,
    residents: Dict[str, dict],
    lookup,
    rules: Dict[Tuple[str, int], str],
    no_call: Dict[str, dict],
    max_iterations: int = 10,
) -> int:
    """Post-generation fairness repair: replace overscheduled residents with
    underscheduled ones on days where both are eligible.

    A swap is accepted when the candidate has at least 2 fewer calls in the
    relevant counter (weekend_calls / weekday_calls) than the assigned resident.
    The gap-of-2 requirement prevents oscillation across passes.

    Returns the total number of swaps made across all iterations.
    """
    # Build index: (date_str, slot) → row position for O(1) lookup.
    # COMPLETED rows are historical ground-truth — never swap them.
    # HOLIDAY rows are manual overrides — never swap them either.
    row_index: Dict[Tuple[str, str], int] = {}
    for i, row in enumerate(schedule_rows):
        note = row.get("note") or ""
        if row["resident"] and note != "COMPLETED" and not note.startswith("HOLIDAY:"):
            row_index[(row["date"], row["slot"])] = i

    intern_names = [n for n, r in residents.items() if r["pgy"] == 1]
    upper_names = [n for n, r in residents.items() if r["pgy"] in (2, 3)]

    total_swaps = 0

    for _iteration in range(max_iterations):
        iteration_swaps = 0

        for (date_str, slot), row_idx in sorted(row_index.items()):
            row = schedule_rows[row_idx]
            if not row["resident"]:
                continue

            d = date.fromisoformat(date_str)
            assigned = row["resident"]

            if slot in (SLOT_INTERN_WEEKEND, SLOT_INTERN_WEEKDAY):
                pool = intern_names
                counter_key = "intern_calls"
            elif slot == SLOT_UPPER_WEEKEND:
                pool = upper_names
                counter_key = "weekend_calls"
            else:
                pool = upper_names
                counter_key = "weekday_calls"

            assigned_count = residents[assigned][counter_key]

            # Find the most underscheduled eligible candidate (gap ≥ 2 required)
            best_candidate: Optional[Tuple[str, str]] = None
            best_count = assigned_count - 1  # threshold: candidate must beat this

            for candidate in pool:
                if candidate == assigned:
                    continue
                cand_count = residents[candidate][counter_key]
                if cand_count >= assigned_count - 1:
                    continue  # gap < 2, skip to avoid oscillation
                if cand_count >= best_count:
                    continue  # not the most underscheduled so far

                # Eligibility: no_call_day
                if d in no_call.get(candidate, {}):
                    continue

                # Eligibility: post-call (candidate not assigned at d-1 or d-2)
                if is_post_call(residents[candidate], d):
                    continue

                # Eligibility: PGY3 graduation cutoff
                pgy = residents[candidate]["pgy"]
                if pgy == 3 and PGY3_CUTOFF_DATE is not None and d >= PGY3_CUTOFF_DATE:
                    continue

                # Eligibility: rotation and preference
                rotation = lookup.rotation_on_date(candidate, d)
                if rotation is None:
                    continue
                pref = rules.get((rotation, pgy))
                if pref is None or pref == "NO_CALL":
                    continue

                # Forward post-call spillover: adding d to candidate must not
                # conflict with their existing assignments at d+1 or d+2
                spillover = any(
                    (d + timedelta(days=i)) in residents[candidate]["assigned_dates"]
                    for i in range(1, POST_CALL_DAYS + 1)
                )
                if spillover:
                    continue

                # Rolling-window cap: adding d must not push the candidate
                # over MAX_CALLS_IN_WINDOW in any ROLLING_WINDOW_DAYS window
                if would_exceed_window_cap(residents[candidate]["assigned_dates"], d):
                    continue

                best_count = cand_count
                best_candidate = (candidate, rotation)

            if best_candidate is None:
                continue

            candidate, rotation = best_candidate
            _undo_assignment(residents, assigned, slot, d)
            apply_assignment(residents, candidate, slot, d)

            row["resident"] = candidate
            row["pgy"] = residents[candidate]["pgy"]
            row["rotation"] = rotation

            iteration_swaps += 1

        total_swaps += iteration_swaps
        if iteration_swaps == 0:
            break

    return total_swaps


def _pre_apply_holidays(
    holidays: Dict[date, dict],
    residents: Dict[str, dict],
    lookup,
    block1_end: Optional[date],
    skip_dates: Optional[set] = None,
) -> Tuple[List[dict], List[dict], Dict[Tuple[date, str], str]]:
    """Pre-apply manual holiday assignments before the main day loop.

    For each holiday date, reads the `upper` and `intern` names from
    holidays.xlsx and:
      - Validates the name exists and has the correct PGY (errors if not).
      - Calls apply_assignment() so every counter (total_calls, weekday/weekend,
        intern/upper, Jul_Dec/Jan_Jun, assigned_dates) is bumped. This makes
        holiday calls participate in fairness, pacing, and post-call logic for
        every subsequent day.
      - Emits a schedule_rows entry with note='HOLIDAY: <name>' so the writer
        can render the name and apply the holiday colour.
      - For blank upper/intern cells, emits an empty schedule row + an
        unassigned_rows entry so the audit still flags them (same as today's
        behaviour for holidays with no staffing).

    Slot inference mirrors load_completed_calls:
      - Weekend → UPPER_WEEKEND / INTERN_WEEKEND.
      - Weekday → UPPER_WEEKDAY / INTERN_WEEKDAY (INTERN_WEEKDAY is used on
        any weekday holiday, including outside Block 1, because the writer
        already prefers the assigned intern over the Night Float display).

    Conflicts (no_call_days, NO_CALL rotation, adjacent-holiday post-call) are
    intentionally NOT rejected — manual holiday assignments win per user
    policy. The audit may still flag them for human review.

    Parameters:
      skip_dates: dates already handled elsewhere (e.g. by completed_calls).
        Holiday rows for these dates are skipped entirely to avoid
        double-counting.

    Returns:
      schedule_rows, unassigned_rows, and a map
      (date, slot) → resident_name of applied holiday assignments — used by
      _precompute_expected_calls to keep pace_value neutral on holiday days.
    """
    skip_dates = skip_dates or set()
    schedule_rows: List[dict] = []
    unassigned_rows: List[dict] = []
    holiday_assignments: Dict[Tuple[date, str], str] = {}

    for d in sorted(holidays.keys()):
        if d in skip_dates:
            continue

        info = holidays[d]
        display_name = info["name"]
        upper_name = info.get("upper")
        intern_name = info.get("intern")

        is_wknd = is_weekend(d)
        upper_slot = SLOT_UPPER_WEEKEND if is_wknd else SLOT_UPPER_WEEKDAY
        intern_slot = SLOT_INTERN_WEEKEND if is_wknd else SLOT_INTERN_WEEKDAY

        for slot, name in ((upper_slot, upper_name), (intern_slot, intern_name)):
            if name:
                if name not in residents:
                    raise DataValidationError(
                        f"Holiday file references unknown resident '{name}' on "
                        f"{d} ({display_name}). Check that names match the flow "
                        f"sheet exactly."
                    )
                pgy = residents[name]["pgy"]
                if slot in (SLOT_INTERN_WEEKEND, SLOT_INTERN_WEEKDAY) and pgy != 1:
                    raise DataValidationError(
                        f"Holiday file assigns '{name}' (PGY{pgy}) to the intern "
                        f"slot on {d} ({display_name}). Interns must be PGY1."
                    )
                if slot in (SLOT_UPPER_WEEKDAY, SLOT_UPPER_WEEKEND) and pgy == 1:
                    raise DataValidationError(
                        f"Holiday file assigns '{name}' (PGY1) to the upper slot "
                        f"on {d} ({display_name}). Upper residents must be PGY2 or PGY3."
                    )

                apply_assignment(residents, name, slot, d)
                holiday_assignments[(d, slot)] = name
                schedule_rows.append({
                    "date": d.isoformat(),
                    "day_of_week": d.strftime("%a"),
                    "slot": slot,
                    "resident": name,
                    "pgy": pgy,
                    "rotation": lookup.rotation_on_date(name, d) or "",
                    "note": f"HOLIDAY: {display_name}",
                })
            else:
                # Blank cell for this slot — leave unassigned so the audit
                # flags it (matches today's behaviour when holidays.xlsx has
                # no names). Non-Block-1 weekday intern slot is skipped
                # silently because that slot doesn't normally exist.
                is_optional_intern_slot = (
                    slot == SLOT_INTERN_WEEKDAY
                    and not (INTERN_BLOCK1_WEEKDAY_CALLS and block1_end is not None and d <= block1_end)
                )
                if is_optional_intern_slot:
                    continue

                schedule_rows.append({
                    "date": d.isoformat(),
                    "day_of_week": d.strftime("%a"),
                    "slot": slot,
                    "resident": "",
                    "pgy": "",
                    "rotation": "",
                    "note": f"HOLIDAY: {display_name}",
                })
                unassigned_rows.append({
                    "date": d.isoformat(),
                    "slot": slot,
                    "holiday": display_name,
                    "reasons": "holiday_manual_assignment",
                })

    return schedule_rows, unassigned_rows, holiday_assignments


def generate_schedule_once(
    seed=None,
    completed_assignments=None,
    config: Optional[dict] = None,
    paths: Optional[dict] = None,
    data_bundle: Optional[DataBundle] = None,
):
    """Generate one schedule for a given seed.

    Parameters:
      seed: rng seed; `None` for nondeterministic.
      completed_assignments: optional partial-year seed data (overrides the
        bundle's completed_calls). Pass the bundle's own completed_calls to
        honor the partial-year mode from paths.
      config: behavior parameters (weights, dates, thresholds). If None, the
        last-applied config is reused (module globals). CLI + GUI both pass
        this explicitly; tests and the auto-apply at import cover the None case.
      paths: file locations. Required unless `data_bundle` is passed.
      data_bundle: pre-built DataBundle. Pass this to avoid re-reading input
        files between calls. If None, a fresh bundle is built from `paths`.
        NOTE: the bundle's `residents` dict is mutated by apply_assignment.
        Callers supplying their own bundle must rebuild it per call or the
        state leaks across runs.
    """
    if config is not None:
        _apply_config(config)

    if data_bundle is None:
        if paths is None:
            _, paths = load_default_config()
        data_bundle = load_data_bundle(
            paths,
            academic_year_start=ACADEMIC_YEAR_START,
            intern_block1_weekday_calls=bool(INTERN_BLOCK1_WEEKDAY_CALLS),
            use_completed_calls=False,  # completed_assignments param wins
            academic_start_date=ACADEMIC_DATE_START,
            academic_end_date=ACADEMIC_DATE_END,
        )

    rng = random.Random(seed)
    tiebreaker_count = 0

    lookup = data_bundle.lookup
    residents = data_bundle.residents
    rules = data_bundle.rules
    no_call = data_bundle.no_call
    holidays = data_bundle.holidays
    block1_end = data_bundle.block1_end

    validate_rotations_against_rules(lookup, residents, rules)
    validate_no_call_days(no_call, residents)  # validates merged dict (both sources)

    intern_names = [n for n, r in residents.items() if r["pgy"] == 1]
    upper_names = [n for n, r in residents.items() if r["pgy"] in (2, 3)]

    # ── Partial year: seed residents with historical calls ────────────────────
    schedule_rows: List[dict] = []
    unassigned_rows: List[dict] = []

    completed_dates: set = set()
    if completed_assignments:
        for ca_date, ca_name, ca_slot in completed_assignments:
            if ca_name not in residents:
                raise DataValidationError(
                    f"Completed call file references unknown resident '{ca_name}' "
                    f"on {ca_date}. Check that names match the flow sheet exactly."
                )
            apply_assignment(residents, ca_name, ca_slot, ca_date)
            schedule_rows.append({
                "date": ca_date.isoformat(),
                "day_of_week": ca_date.strftime("%a"),
                "slot": ca_slot,
                "resident": ca_name,
                "pgy": residents[ca_name]["pgy"],
                "rotation": lookup.rotation_on_date(ca_name, ca_date) or "",
                "note": "COMPLETED",
            })
            completed_dates.add(ca_date)
        restart_date = max(ca_date for ca_date, _, _ in completed_assignments) + timedelta(days=1)
    else:
        restart_date = ACADEMIC_DATE_START
    # ─────────────────────────────────────────────────────────────────────────

    # ── Pre-apply manually-assigned holiday calls ────────────────────────────
    # Each row in holidays.xlsx may name an upper and/or intern resident.
    # Those names are applied here, BEFORE the main day loop, so their
    # counters are bumped and their post-call days correctly block them from
    # subsequent assignments. Completed dates win to avoid double-counting.
    holiday_rows, holiday_unassigned, holiday_assignments = _pre_apply_holidays(
        holidays=holidays,
        residents=residents,
        lookup=lookup,
        block1_end=block1_end,
        skip_dates=completed_dates,
    )
    schedule_rows.extend(holiday_rows)
    unassigned_rows.extend(holiday_unassigned)
    # ─────────────────────────────────────────────────────────────────────────

    # Precompute cumulative expected-call counts per (resident, counter_key, date)
    # using each day's static-eligibility pool. Each eligible resident is "owed"
    # 1/|pool| per slot. The pacing signal compares actual counters to these
    # expected values to prevent first-half/second-half concentration when a
    # resident has alternating eligible and NO_CALL stretches.
    expected_cum = _precompute_expected_calls(
        lookup=lookup,
        residents=residents,
        rules=rules,
        no_call_days=no_call,
        holidays=holidays,
        intern_names=intern_names,
        upper_names=upper_names,
        block1_end=block1_end,
        holiday_assignments=holiday_assignments,
    )

    # Per-resident sorted list of statically-eligible dates — used for the
    # lookahead component (remaining runway through year-end).
    eligible_dates = _precompute_eligible_dates(
        lookup=lookup,
        residents=residents,
        rules=rules,
        no_call_days=no_call,
        holidays=holidays,
        block1_end=block1_end,
    )

    d = restart_date
    while d <= ACADEMIC_DATE_END:
        # Holiday days are handled by _pre_apply_holidays above. Counters are
        # already bumped and schedule_rows already populated; just advance.
        if d in holidays:
            d += timedelta(days=1)
            continue

        for slot in required_slots(d, block1_end=block1_end):
            eligible, reasons = eligible_for_slot(
                lookup, residents, rules, no_call, d, slot, intern_names, upper_names
            )
            picked = pick_best_candidate(residents, eligible, d, slot, rng, expected_cum, eligible_dates)

            if picked is None:
                schedule_rows.append({
                    "date": d.isoformat(),
                    "day_of_week": d.strftime("%a"),
                    "slot": slot,
                    "resident": "",
                    "pgy": "",
                    "rotation": "",
                    "note": "UNASSIGNED",
                })
                unassigned_rows.append({
                    "date": d.isoformat(),
                    "slot": slot,
                    "holiday": "",
                    "reasons": str(reasons),
                })
            else:
                name, rotation, was_tiebreak = picked
                if was_tiebreak:
                    tiebreaker_count += 1
                apply_assignment(residents, name, slot, d)
                schedule_rows.append({
                    "date": d.isoformat(),
                    "day_of_week": d.strftime("%a"),
                    "slot": slot,
                    "resident": name,
                    "pgy": residents[name]["pgy"],
                    "rotation": rotation,
                    "note": "",
                })

        d += timedelta(days=1)

    # Post-generation local swap pass: repair fairness imbalances the greedy missed.
    swap_count = local_swap_pass(schedule_rows, residents, lookup, rules, no_call)

    audit_data = audit_schedule(
        schedule_rows=schedule_rows,
        residents=residents,
        lookup=lookup,
        rules=rules,
        no_call_days=no_call,
        unassigned_rows=unassigned_rows,
        holidays=holidays,
        seed=seed,
        tiebreaker_count=tiebreaker_count,
    )

    audit_data["intern_block1_weekday_calls"] = bool(INTERN_BLOCK1_WEEKDAY_CALLS)
    audit_data["block1_end"] = block1_end.isoformat() if block1_end else None
    audit_data["pick_candidate_rank_order"] = PICK_CANDIDATE_RANK_ORDER
    audit_data["pick_candidate_weights"] = {
        "FAIRNESS_GAP_WEIGHT": FAIRNESS_GAP_WEIGHT,
        "SPACING_WEIGHT": SPACING_WEIGHT,
        "AVOID_WEIGHT": AVOID_WEIGHT,
        "YEAR_BIAS_WEIGHT": YEAR_BIAS_WEIGHT,
        "PACE_WEIGHT": PACE_WEIGHT,
        "LOOKAHEAD_WEIGHT": LOOKAHEAD_WEIGHT,
    }
    audit_data["monte_carlo_score_order"] = MONTE_CARLO_SCORE_ORDER
    audit_data["swap_improvements"] = swap_count
    audit_data["restart_date"] = restart_date.isoformat() if restart_date != ACADEMIC_DATE_START else None
    audit_data["completed_call_count"] = len(completed_assignments) if completed_assignments else 0

    return {
        "schedule_rows": schedule_rows,
        "residents": residents,
        "rules": rules,
        "lookup": lookup,
        "holidays": holidays,
        "no_call": no_call,
        "unassigned_rows": unassigned_rows,
        "audit_data": audit_data,
    }


def monte_carlo_score(result):
    audit = result["audit_data"]
    fairness = audit["fairness_summary"]

    score_components = {
        "errors": len(audit["errors"]),
        "unassigned": len(result["unassigned_rows"]),
        "upper_total_diff": fairness["upper_total_diff"],
        "pgy2_total_diff": fairness["pgy2_total_diff"],
        "pgy3_total_diff": fairness["pgy3_total_diff"],
        "upper_weekend_diff": fairness["upper_weekend_diff"],
        "upper_weekday_diff": fairness["upper_weekday_diff"],
        "intern_weekend_diff": fairness["intern_weekend_diff"],
        "avoid_assignments": len(audit.get("avoid_assignments", [])),
        "warnings": len(audit["warnings"]),
    }

    return tuple(score_components[key] for key in MONTE_CARLO_SCORE_ORDER)


def format_monte_carlo_score(score):
    return ", ".join(f"{key}={value}" for key, value in zip(MONTE_CARLO_SCORE_ORDER, score))

def _score_seed(
    seed: int,
    config: dict,
    paths: dict,
    completed_assignments=None,
) -> tuple:
    """Worker function: returns (seed, score_tuple). Must be top-level for pickling.

    Runs in a ProcessPoolExecutor child process. On Windows (spawn) the child
    re-imports this module, re-runs _apply_config(CONFIG), then
    generate_schedule_once re-applies the caller-supplied config. The data
    bundle is rebuilt fresh per call — residents state cannot leak between
    seeds because each call builds its own bundle.
    """
    result = generate_schedule_once(
        seed=seed,
        config=config,
        paths=paths,
        completed_assignments=completed_assignments,
    )
    return seed, monte_carlo_score(result)


def run_simulation(
    num_runs: int,
    config: dict,
    paths: dict,
    completed_assignments=None,
    progress_callback: Optional[Callable[[int, int, dict], None]] = None,
    cancel_event: Optional[threading.Event] = None,
):
    # Apply config in the driver process too — format_monte_carlo_score reads
    # MONTE_CARLO_SCORE_ORDER from module globals, and workers' results come
    # back here for logging.
    _apply_config(config)

    start = time.time()
    seed_scores: list[tuple[int, tuple]] = []

    worker = partial(
        _score_seed,
        config=config,
        paths=paths,
        completed_assignments=completed_assignments or [],
    )

    cancelled = False
    with concurrent.futures.ProcessPoolExecutor() as executor:
        futures = {executor.submit(worker, seed): seed for seed in range(num_runs)}
        for future in concurrent.futures.as_completed(futures):
            seed, score = future.result()
            seed_scores.append((seed, score))
            logger.info(f"[{len(seed_scores)}/{num_runs}] seed={seed} | {format_monte_carlo_score(score)}")

            if progress_callback is not None:
                best_seed_so_far, best_score_so_far = min(seed_scores, key=lambda x: x[1])
                progress_callback(
                    len(seed_scores),
                    num_runs,
                    {
                        "seed": seed,
                        "score": score,
                        "best_seed": best_seed_so_far,
                        "best_score": best_score_so_far,
                    },
                )

            if cancel_event is not None and cancel_event.is_set():
                cancelled = True
                executor.shutdown(wait=False, cancel_futures=True)
                break

    if cancelled:
        logger.info("Simulation cancelled.")
        return None

    best_seed, best_score = min(seed_scores, key=lambda x: x[1])
    best_result = generate_schedule_once(
        seed=best_seed,
        config=config,
        paths=paths,
        completed_assignments=completed_assignments,
    )

    end = time.time()
    logger.info(f"\nSimulation completed in {end - start:.2f} seconds")
    logger.info(f"Best run: seed={best_seed} | {format_monte_carlo_score(best_score)}\n")

    return best_result


def export_result(result, paths: dict):
    schedule_rows = result["schedule_rows"]
    residents = result["residents"]
    lookup = result["lookup"]
    holidays = result["holidays"]
    no_call = result["no_call"]

    audit_data = result["audit_data"]
    seed = audit_data.get("seed", "N/A")

    logger.info("\nFINAL SCHEDULE SELECTION")
    logger.info(f"Seed used: {seed}")
    logger.info(f"Tie-break decisions: {audit_data.get('tiebreaker_count', 0)}")

    intern_names = [n for n, r in residents.items() if r["pgy"] == 1]

    out_dir = f"{paths['data_dir']}/{paths['output_dir']}"
    write_call_totals_xlsx(residents, f"{out_dir}/call_totals.xlsx")
    write_call_schedule_xlsx(
        schedule_rows,
        holidays,
        no_call,
        f"{out_dir}/call_schedule.xlsx",
        lookup,
        intern_names,
    )
    logger.info("Excel files exported:")
    logger.info(f"  {out_dir}/call_totals.xlsx")
    logger.info(f"  {out_dir}/call_schedule.xlsx")

    audit_data = result["audit_data"]
    write_audit(audit_data, path=f"{out_dir}/audit_report.txt")
    logger.info(f"Wrote to: {out_dir}/audit_report.txt")


def _main() -> int:
    # Stream plain messages to stdout so output matches the pre-logging
    # print()-based CLI byte-for-byte. force=True in case an ancestor import
    # already installed a root handler (e.g. a library-style basicConfig).
    logging.basicConfig(
        level=logging.INFO,
        format="%(message)s",
        stream=sys.stdout,
        force=True,
    )

    warning = legacy_gui_config_warning()
    if warning is not None:
        logger.warning(warning)

    config, paths = load_default_config()
    _apply_config(config)

    use_completed = bool(int(config.get("USE_COMPLETED_CALLS", 0)))

    if use_completed:
        # Building a DataBundle here would be wasteful (run_simulation rebuilds
        # one per seed); we only need block1_end to classify weekday intern
        # entries correctly. Build it via a throwaway bundle load, or just
        # read the flow sheet once.
        if INTERN_BLOCK1_WEEKDAY_CALLS:
            bundle = load_data_bundle(
                paths,
                academic_year_start=ACADEMIC_YEAR_START,
                intern_block1_weekday_calls=True,
                use_completed_calls=False,
                academic_start_date=ACADEMIC_DATE_START,
                academic_end_date=ACADEMIC_DATE_END,
            )
            _block1_end = bundle.block1_end
        else:
            _block1_end = None
        completed = load_completed_calls(paths["completed_calls_xlsx"], block1_end=_block1_end)
    else:
        completed = []

    if completed:
        restart = max(d for d, _, _ in completed) + timedelta(days=1)
        logger.info(f"Partial year mode: {len(completed)} completed call(s) loaded.")
        logger.info(f"Generating schedule from {restart} onward.\n")
    else:
        logger.info("Full year mode: generating schedule from scratch.\n")

    best_result = run_simulation(
        num_runs=SIMULATION_RUNS,
        config=config,
        paths=paths,
        completed_assignments=completed,
    )
    export_result(best_result, paths=paths)
    return 0


if __name__ == "__main__":
    try:
        sys.exit(_main())
    except ScheduleError as exc:
        # Friendly prose for known, user-actionable failures.
        print(f"Error: {exc}", file=sys.stderr)
        sys.exit(1)
