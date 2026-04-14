from __future__ import annotations

import bisect
import concurrent.futures
import random
import time
from datetime import date, datetime, timedelta
from functools import partial
from typing import Dict, List, Tuple, Optional

from config import CONFIG
from excel_reader import ExcelRotationLookup
from exports import write_call_totals_xlsx, write_call_schedule_xlsx, write_audit
from loader import (
    load_residents, load_no_call_days, load_holidays, load_rotation_rules,
    load_clinic_days, load_completed_calls,
)
from validation import validate_rotations_against_rules, validate_no_call_days, audit_schedule

DATA_DIR = CONFIG.get("DATA_DIR", "data")
OUTPUT_DIR = CONFIG.get("OUTPUT_DIR", "output")
FLOW_XLSX = CONFIG.get("FLOW_XLSX", "data/flow.xlsx")
SHEET_NAME = CONFIG.get("SHEET_NAME", "master_block_calendar")
CLINIC_DAYS_XLSX = CONFIG.get("CLINIC_DAYS_XLSX", "data/clinic_days.xlsx")
COMPLETED_CALLS_XLSX = CONFIG.get("COMPLETED_CALLS_XLSX", "")

POST_CALL_DAYS = CONFIG.get("POST_CALL_DAYS")
SIMULATION_RUNS = CONFIG.get("SIMULATION_RUNS")

ACADEMIC_DATE_START_STRING = CONFIG.get("ACADEMIC_DATE_START_STRING")
ACADEMIC_DATE_END_STRING = CONFIG.get("ACADEMIC_DATE_END_STRING")
MIN_SPACING_DAYS_STRONG = CONFIG.get("MIN_SPACING_DAYS_STRONG")
MIN_SPACING_DAYS_MILD = CONFIG.get("MIN_SPACING_DAYS_MILD")

MAX_CALLS_IN_WINDOW = int(CONFIG.get("MAX_CALLS_IN_WINDOW", 0))
ROLLING_WINDOW_DAYS = int(CONFIG.get("ROLLING_WINDOW_DAYS", 14))

MAX_DIFF_SOFT = CONFIG.get("MAX_DIFF_SOFT")
MAX_DIFF_HARD = CONFIG.get("MAX_DIFF_HARD")

FAIRNESS_GAP_WEIGHT = CONFIG.get("FAIRNESS_GAP_WEIGHT")
SPACING_WEIGHT = CONFIG.get("SPACING_WEIGHT")
AVOID_WEIGHT = CONFIG.get("AVOID_WEIGHT")
YEAR_BIAS_WEIGHT = CONFIG.get("YEAR_BIAS_WEIGHT")
FUTURE_AVAIL_WEIGHT = CONFIG.get("FUTURE_AVAIL_WEIGHT")


ACADEMIC_DATE_START = datetime.strptime(ACADEMIC_DATE_START_STRING, "%Y-%m-%d").date()
ACADEMIC_DATE_END = datetime.strptime(ACADEMIC_DATE_END_STRING, "%Y-%m-%d").date()
ACADEMIC_YEAR_START = ACADEMIC_DATE_START.year
TOTAL_YEAR_DAYS = (ACADEMIC_DATE_END - ACADEMIC_DATE_START).days
FIRST_HALF_END = date(ACADEMIC_YEAR_START, 12, 31)

SLOT_UPPER_WEEKDAY = "UPPER_WEEKDAY"
SLOT_UPPER_WEEKEND = "UPPER_WEEKEND"
SLOT_INTERN_WEEKEND = "INTERN_WEEKEND"
SLOT_INTERN_WEEKDAY = "INTERN_WEEKDAY"

INTERN_BLOCK1_WEEKDAY_CALLS = int(CONFIG.get("INTERN_BLOCK1_WEEKDAY_CALLS", 0))


MONTE_CARLO_SCORE_ORDER = CONFIG.get(
    "MONTE_CARLO_SCORE_ORDER",
    [
        "errors",
        "unassigned",
        "upper_weekend_diff",
        "upper_weekday_diff",
        "upper_total_diff",
        "intern_weekend_diff",
        "avoid_assignments",
        "warnings",
    ],
)

VALID_MONTE_CARLO_SCORE_KEYS = {
    "errors",
    "unassigned",
    "upper_weekend_diff",
    "upper_weekday_diff",
    "upper_total_diff",
    "intern_weekend_diff",
    "avoid_assignments",
    "warnings",
}

invalid_score_keys = [k for k in MONTE_CARLO_SCORE_ORDER if k not in VALID_MONTE_CARLO_SCORE_KEYS]
if invalid_score_keys:
    raise ValueError(
        f"Invalid MONTE_CARLO_SCORE_ORDER entries: {invalid_score_keys}. "
        f"Valid options are: {sorted(VALID_MONTE_CARLO_SCORE_KEYS)}"
    )

PICK_CANDIDATE_RANK_ORDER = CONFIG.get(
    "PICK_CANDIDATE_RANK_ORDER",
    [
        "hard_diff_flag",
        "soft_diff_flag",
        "weighted_score",
    ],
)

VALID_PICK_CANDIDATE_RANK_KEYS = {
    "hard_diff_flag",
    "soft_diff_flag",
    "weighted_score",
}

invalid_pick_rank_keys = [
    k for k in PICK_CANDIDATE_RANK_ORDER
    if k not in VALID_PICK_CANDIDATE_RANK_KEYS
]

if invalid_pick_rank_keys:
    raise ValueError(
        f"Invalid PICK_CANDIDATE_RANK_ORDER entries: {invalid_pick_rank_keys}. "
        f"Valid options are: {sorted(VALID_PICK_CANDIDATE_RANK_KEYS)}"
    )





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
    for i in range(1, POST_CALL_DAYS + 1):
        if (d - timedelta(days=i)) in resident_data["assigned_dates"]:
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


def _compute_weighted_score(
    fairness_gap: int,
    spacing_value: int,
    avoid_value: int,
    year_value: float,
    future_avail_value: float = 0.0,
) -> float:
    """Combine ranking components into a single scalar.

    Each raw component lives on a different scale (fairness_gap is an
    unbounded int, spacing_value is {0,1,2}, avoid/year/future_avail are
    already in [0,1]). We normalize each to roughly [0, 1] so the *_WEIGHT
    constants express genuine relative importance.

    fairness_gap is clipped at MAX_DIFF_SOFT: above that threshold the
    soft/hard lexicographic gates have already separated candidates, so
    losing resolution there is acceptable.

    future_avail_value: fraction of remaining call-eligible days this resident
    has vs. the pool maximum. Higher = more days remaining = can be deferred =
    less preferred now. Defaults to 0.0 (no look-ahead contribution).
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
        + FUTURE_AVAIL_WEIGHT * future_avail_value
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

        if would_exceed_window_cap(data["assigned_dates"], d):
            reasons["window_cap"] = reasons.get("window_cap", 0) + 1
            continue

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

        eligible.append((name, pref, rotation))

    return eligible, reasons


def pick_best_candidate(
    residents: Dict[str, dict],
    eligible: List[Tuple[str, str, str]],
    d: date,
    slot: str,
    rng: random.Random,
    future_eligible: Optional[Dict[str, list]] = None,
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

    def year_bias(pgy: int) -> float:
        if slot in (SLOT_INTERN_WEEKEND, SLOT_INTERN_WEEKDAY):
            return 0.0
        if pgy == 3:
            return prog
        elif pgy == 2:
            return 1 - prog
        return 0.0

    # Precompute future availability for each eligible candidate.
    # remaining = # dates after d in their precomputed eligible list (static constraints only).
    # Normalize within this eligible pool so the component is always in [0, 1].
    # Higher remaining → more deferrable → higher score → less preferred now.
    if future_eligible is not None:
        future_remaining = {
            name: len(future_eligible.get(name, [])) - bisect.bisect(future_eligible.get(name, []), d)
            for name, _, _ in eligible
        }
        max_future_remaining = max(future_remaining.values(), default=0)
    else:
        future_remaining = {}
        max_future_remaining = 0

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

        if max_future_remaining > 0:
            future_avail_value = future_remaining.get(name, 0) / max_future_remaining
        else:
            future_avail_value = 0.0

        weighted_score = _compute_weighted_score(
            fairness_gap=fairness_gap,
            spacing_value=spacing_value,
            avoid_value=avoid_value,
            year_value=year_value,
            future_avail_value=future_avail_value,
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
    row_index: Dict[Tuple[str, str], int] = {}
    for i, row in enumerate(schedule_rows):
        if row["resident"] and row.get("note") != "COMPLETED":
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

                # Eligibility: rotation and preference
                rotation = lookup.rotation_on_date(candidate, d)
                if rotation is None:
                    continue
                pgy = residents[candidate]["pgy"]
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


def _merge_no_call(
    base: Dict[str, dict],
    overlay: Dict[str, dict],
) -> Dict[str, dict]:
    """Return a new dict that is base updated with overlay entries.
    Neither input is mutated.
    """
    merged = {name: dict(dates) for name, dates in base.items()}
    for name, dates in overlay.items():
        merged.setdefault(name, {}).update(dates)
    return merged


def generate_schedule_once(seed=None, completed_assignments=None):
    rng = random.Random(seed)
    tiebreaker_count = 0

    lookup = ExcelRotationLookup(FLOW_XLSX, SHEET_NAME, ACADEMIC_YEAR_START)
    block1_end = lookup.blocks[0].end if (INTERN_BLOCK1_WEEKDAY_CALLS and lookup.blocks) else None
    residents = load_residents(lookup)
    rules = load_rotation_rules()
    no_call_base = load_no_call_days()
    clinic_pre_blocks = load_clinic_days(CLINIC_DAYS_XLSX)
    # Merge clinic-derived pre-call blocks into no_call so the same eligibility
    # path enforces both; the distinct reason string ("pre_clinic_day") is
    # preserved so the unassigned report can distinguish the two sources.
    no_call = _merge_no_call(no_call_base, clinic_pre_blocks)
    holidays = load_holidays()

    validate_rotations_against_rules(lookup, residents, rules)
    validate_no_call_days(no_call, residents)  # validates merged dict (both sources)

    intern_names = [n for n, r in residents.items() if r["pgy"] == 1]
    upper_names = [n for n, r in residents.items() if r["pgy"] in (2, 3)]

    # ── Partial year: seed residents with historical calls ────────────────────
    schedule_rows: List[dict] = []
    unassigned_rows: List[dict] = []

    if completed_assignments:
        for ca_date, ca_name, ca_slot in completed_assignments:
            if ca_name not in residents:
                raise ValueError(
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
        restart_date = max(ca_date for ca_date, _, _ in completed_assignments) + timedelta(days=1)
    else:
        restart_date = ACADEMIC_DATE_START
    # ─────────────────────────────────────────────────────────────────────────

    # Precompute sorted lists of future call-eligible dates per resident.
    # Uses static constraints only (rotation preference + no_call_days, which
    # now includes pre-clinic blocks); post-call exclusions are dynamic.
    future_eligible: Dict[str, list] = {}
    for name, data in residents.items():
        pgy = data["pgy"]
        dates = []
        di = ACADEMIC_DATE_START
        while di <= ACADEMIC_DATE_END:
            if di not in holidays and di not in no_call.get(name, {}):
                rotation = lookup.rotation_on_date(name, di)
                if rotation is not None:
                    pref = rules.get((rotation, pgy))
                    if pref not in (None, "NO_CALL"):
                        dates.append(di)
            di += timedelta(days=1)
        future_eligible[name] = dates  # ascending order — ready for bisect

    d = restart_date
    while d <= ACADEMIC_DATE_END:
        if d in holidays:
            for slot in required_slots(d, block1_end=block1_end):
                schedule_rows.append({
                    "date": d.isoformat(),
                    "day_of_week": d.strftime("%a"),
                    "slot": slot,
                    "resident": "",
                    "pgy": "",
                    "rotation": "",
                    "note": f"HOLIDAY: {holidays[d]}",
                })
                unassigned_rows.append({
                    "date": d.isoformat(),
                    "slot": slot,
                    "holiday": holidays[d],
                    "reasons": "holiday_manual_assignment",
                })
            d += timedelta(days=1)
            continue

        for slot in required_slots(d, block1_end=block1_end):
            eligible, reasons = eligible_for_slot(
                lookup, residents, rules, no_call, d, slot, intern_names, upper_names
            )
            picked = pick_best_candidate(residents, eligible, d, slot, rng, future_eligible)

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
        "FUTURE_AVAIL_WEIGHT": FUTURE_AVAIL_WEIGHT,
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
        "upper_weekend_diff": fairness["upper_weekend_diff"],
        "upper_weekday_diff": fairness["upper_weekday_diff"],
        "intern_weekend_diff": fairness["intern_weekend_diff"],
        "avoid_assignments": len(audit.get("avoid_assignments", [])),
        "warnings": len(audit["warnings"]),
    }

    return tuple(score_components[key] for key in MONTE_CARLO_SCORE_ORDER)


def format_monte_carlo_score(score):
    return ", ".join(f"{key}={value}" for key, value in zip(MONTE_CARLO_SCORE_ORDER, score))

def _score_seed(seed: int, completed_assignments=None) -> tuple:
    """Worker function: returns (seed, score_tuple). Must be top-level for pickling."""
    result = generate_schedule_once(seed=seed, completed_assignments=completed_assignments)
    return seed, monte_carlo_score(result)


def run_simulation(num_runs=50, completed_assignments=None):
    start = time.time()
    seed_scores: list[tuple[int, tuple]] = []

    worker = partial(_score_seed, completed_assignments=completed_assignments or [])

    with concurrent.futures.ProcessPoolExecutor() as executor:
        futures = {executor.submit(worker, seed): seed for seed in range(num_runs)}
        for future in concurrent.futures.as_completed(futures):
            seed, score = future.result()
            seed_scores.append((seed, score))
            print(f"[{len(seed_scores)}/{num_runs}] seed={seed} | {format_monte_carlo_score(score)}")

    best_seed, best_score = min(seed_scores, key=lambda x: x[1])
    best_result = generate_schedule_once(seed=best_seed, completed_assignments=completed_assignments)

    end = time.time()
    print(f"\nSimulation completed in {end - start:.2f} seconds")
    print(f"Best run: seed={best_seed} | {format_monte_carlo_score(best_score)}\n")

    return best_result


def export_result(result):
    schedule_rows = result["schedule_rows"]
    residents = result["residents"]
    lookup = result["lookup"]
    holidays = result["holidays"]
    no_call = result["no_call"]

    audit_data = result["audit_data"]
    seed = audit_data.get("seed", "N/A")

    print("\nFINAL SCHEDULE SELECTION")
    print(f"Seed used: {seed}")
    print(f"Tie-break decisions: {audit_data.get('tiebreaker_count', 0)}")

    intern_names = [n for n, r in residents.items() if r["pgy"] == 1]

    write_call_totals_xlsx(residents, f"{DATA_DIR}/{OUTPUT_DIR}/call_totals.xlsx")
    write_call_schedule_xlsx(
        schedule_rows,
        holidays,
        no_call,
        f"{DATA_DIR}/{OUTPUT_DIR}/call_schedule.xlsx",
        lookup,
        intern_names,
    )
    print("Excel files exported:")
    print(f"  {DATA_DIR}/{OUTPUT_DIR}/call_totals.xlsx")
    print(f"  {DATA_DIR}/{OUTPUT_DIR}/call_schedule.xlsx")

    audit_data = result["audit_data"]
    write_audit(audit_data, path=f"{DATA_DIR}/{OUTPUT_DIR}/audit_report.txt")
    print(f"Wrote to: {DATA_DIR}/{OUTPUT_DIR}/audit_report.txt")


if __name__ == "__main__":
    # When Block 1 intern weekday calls are enabled, we need the block 1 end date
    # so load_completed_calls can correctly classify weekday intern entries.
    if INTERN_BLOCK1_WEEKDAY_CALLS and COMPLETED_CALLS_XLSX:
        _lu = ExcelRotationLookup(FLOW_XLSX, SHEET_NAME, ACADEMIC_YEAR_START)
        _block1_end = _lu.blocks[0].end if _lu.blocks else None
    else:
        _block1_end = None
    completed = load_completed_calls(COMPLETED_CALLS_XLSX, block1_end=_block1_end) if COMPLETED_CALLS_XLSX else []

    if completed:
        from datetime import timedelta as _td
        restart = max(d for d, _, _ in completed) + _td(days=1)
        print(f"Partial year mode: {len(completed)} completed call(s) loaded.")
        print(f"Generating schedule from {restart} onward.\n")
    else:
        print("Full year mode: generating schedule from scratch.\n")

    best_result = run_simulation(num_runs=SIMULATION_RUNS, completed_assignments=completed)
    export_result(best_result)
