from __future__ import annotations

import random
import time
from datetime import date, datetime, timedelta
from typing import Dict, List, Tuple, Optional

from config import CONFIG
from excel_reader import ExcelRotationLookup
from exports import write_call_totals_xlsx, write_call_schedule_xlsx, write_audit
from loader import load_residents, load_no_call_days, load_holidays, load_rotation_rules
from validation import validate_rotations_against_rules, validate_no_call_days, audit_schedule

TIEBREAKER_COUNT = 0

# ---------------- Config ----------------

#See descriptions in config document
# Variables related to file names and locations
DATA_DIR = CONFIG.get("DATA_DIR", "data")
OUTPUT_DIR = CONFIG.get("OUTPUT_DIR","output")
FLOW_XLSX = CONFIG.get("FLOW_XLSX","data/flow.xlsx")
SHEET_NAME = CONFIG.get("SHEET_NAME", "master_block_calendar")

POST_CALL_DAYS = CONFIG.get("POST_CALL_DAYS", 1)

SIMULATION_RUNS = CONFIG.get("SIMULATION_RUNS", 100)

# Variables related to academic year start/end dates
ACADEMIC_DATE_START_STRING = CONFIG.get("ACADEMIC_DATE_START_STRING","2026-07-01")
ACADEMIC_DATE_END_STRING = CONFIG.get("ACADEMIC_DATE_END_STRING","2027-06-30")
AVOID_ROTATION_PENALTY = CONFIG.get("AVOID_ROTATION_PENALTY", 200)
MIN_SPACING_DAYS_STRONG = CONFIG.get("MIN_SPACING_DAYS_STRONG", 7)
MIN_SPACING_DAYS_MILD = CONFIG.get("MIN_SPACING_DAYS_MILD", 14)
SPACING_STRONG_PENALTY = CONFIG.get("SPACING_STRONG_PENALTY", 100)
SPACING_MILD_PENALTY = CONFIG.get("SPACING_MILD_PENALTY",25)
YEAR_PROGRESS_MODIFIER = CONFIG.get("YEAR_PROGRESS_MODIFIER", 50)
MAX_DIFF_SOFT = CONFIG.get("MAX_DIFF_SOFT", 3)
MAX_DIFF_HARD = CONFIG.get("MAX_DIFF_HARD", 5)
TIGHTNESS_PENALTY = CONFIG.get("TIGHTNESS_PENALTY", 100)
HARD_DIFF_PENALTY = CONFIG.get("HARD_DIFF_PENALTY", 2000)


# --------------- DATE Formatting ------------------------------------------

ACADEMIC_DATE_START = datetime.strptime(ACADEMIC_DATE_START_STRING, "%Y-%m-%d").date()
ACADEMIC_DATE_END = datetime.strptime(ACADEMIC_DATE_END_STRING, "%Y-%m-%d").date()
ACADEMIC_YEAR_START = ACADEMIC_DATE_START.year
TOTAL_YEAR_DAYS = (ACADEMIC_DATE_END - ACADEMIC_DATE_START).days

# ---------------- Slots ----------------

SLOT_UPPER_WEEKDAY = "UPPER_WEEKDAY"
SLOT_UPPER_WEEKEND = "UPPER_WEEKEND"
SLOT_INTERN_WEEKEND = "INTERN_WEEKEND"

#----------------------------------------


def is_weekend(d: date) -> bool:
    return d.weekday() >= 5  # Sat/Sun


def required_slots(d: date) -> List[str]:
    if is_weekend(d):
        return [SLOT_INTERN_WEEKEND, SLOT_UPPER_WEEKEND]
    return [SLOT_UPPER_WEEKDAY]


def year_progress(d: date) -> float:
    # 0.0 near July 1; 1.0 near June 30
    return (d - ACADEMIC_DATE_START).days / TOTAL_YEAR_DAYS


# ---------------- Core logic ----------------

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


def eligible_for_slot(
    lookup,
    residents: Dict[str, dict],
    rules: Dict[Tuple[str, int], str],
    no_call_days: Dict[str, dict],
    d: date,
    slot: str,
    intern_names,
    upper_names
) -> Tuple[List[Tuple[str, str, str]], Dict[str, int]]:
    """
    Returns:
      eligible: list of (name, preference, rotation)
      reasons: counts for why excluded
    """
    eligible: List[Tuple[str, str, str]] = []
    reasons: Dict[str, int] = {}

    names = intern_names if slot == SLOT_INTERN_WEEKEND else upper_names
    for name in names:
        data = residents[name]

        pgy = data["pgy"]

        # Slot ↔ PGY hard rules
        if slot == SLOT_INTERN_WEEKEND:
            if pgy != 1:
                reasons["pgy_mismatch"] = reasons.get("pgy_mismatch", 0) + 1
                continue
        else:
            if pgy == 1:
                reasons["pgy_mismatch"] = reasons.get("pgy_mismatch", 0) + 1
                continue

        # No-call day overrides everything
        if d in no_call_days.get(name, set()):
            reasons["no_call_day"] = reasons.get("no_call_day", 0) + 1
            continue

        # Post-call
        if is_post_call(data, d):
            reasons["post_call"] = reasons.get("post_call", 0) + 1
            continue

        # Rotation lookup from Excel
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
) -> Optional[Tuple[str, str]]:
    """
    Returns (chosen_name, rotation) or None.
    """
    if not eligible:
        return None

    # Choose call counter group
    if slot == SLOT_INTERN_WEEKEND:
        counter_key = "weekend_calls"
    elif slot == SLOT_UPPER_WEEKEND:
        counter_key = "weekend_calls"
    else:
        counter_key = "weekday_calls"

    # Determine which pool we are balancing
    if slot == SLOT_INTERN_WEEKEND:
        pool = [n for n, r in residents.items() if r["pgy"] == 1]
    elif slot in (SLOT_UPPER_WEEKDAY, SLOT_UPPER_WEEKEND):
        pool = [n for n, r in residents.items() if r["pgy"] in (2, 3)]
    else:
        pool = list(residents.keys())

    min_in_pool = min(residents[n][counter_key] for n in pool)

    best_score: Optional[float] = None
    best: List[Tuple[str, str]] = []

    prog = year_progress(d)

    for name, pref, rotation in eligible:
        data = residents[name]
        pgy = data["pgy"]

        diff = data[counter_key] - min_in_pool
        score = diff * TIGHTNESS_PENALTY

        if diff > MAX_DIFF_SOFT:
            score += (diff - MAX_DIFF_SOFT) * TIGHTNESS_PENALTY

        if diff > MAX_DIFF_HARD:
            score += HARD_DIFF_PENALTY

        # Spacing penalty (soft)
        spacing = days_since_last_call(data, d)
        if spacing < MIN_SPACING_DAYS_STRONG:
            score += SPACING_STRONG_PENALTY
        elif spacing < MIN_SPACING_DAYS_MILD:
            score += SPACING_MILD_PENALTY

        # AVOID penalty
        if pref == "AVOID":
            score += AVOID_ROTATION_PENALTY

        # Front/back loading (upper only)
        if slot != SLOT_INTERN_WEEKEND:
            if pgy == 3:
                score += YEAR_PROGRESS_MODIFIER * prog
            elif pgy == 2:
                score += YEAR_PROGRESS_MODIFIER * (1 - prog)

        if best_score is None or score < best_score:
            best_score = score
            best = [(name, rotation)]
        elif score == best_score:
            best.append((name, rotation))

    global TIEBREAKER_COUNT
    if len(best) > 1:
        TIEBREAKER_COUNT += 1

    # Deterministic tie-break (alphabetical)
    #chosen = sorted(best, key=lambda x: x[0])[0]

    # Random Seed tie-breaker
    chosen = random.choice(best)

    return chosen


def apply_assignment(residents: Dict[str, dict], name: str, slot: str, d: date) -> None:
    data = residents[name]
    data["assigned_dates"].append(d)

    data["total_calls"] += 1
    if is_weekend(d):
        data["weekend_calls"] += 1
    else:
        data["weekday_calls"] += 1

    # optional (keep if useful elsewhere)
    if slot == SLOT_INTERN_WEEKEND:
        data["intern_calls"] += 1
    else:
        data["upper_calls"] += 1








# ---------------- Full-year generation ----------------


def generate_schedule_once(seed=None):
    global TIEBREAKER_COUNT
    TIEBREAKER_COUNT = 0

    if seed is not None:
        random.seed(seed)

    lookup = ExcelRotationLookup(FLOW_XLSX, SHEET_NAME, ACADEMIC_YEAR_START)
    residents = load_residents(lookup)
    rules = load_rotation_rules()
    no_call = load_no_call_days()
    holidays = load_holidays()

    validate_rotations_against_rules(lookup, residents, rules)
    validate_no_call_days(no_call, residents)

    intern_names = [n for n, r in residents.items() if r["pgy"] == 1]
    upper_names = [n for n, r in residents.items() if r["pgy"] in (2, 3)]

    schedule_rows = []
    unassigned_rows = []

    d = ACADEMIC_DATE_START
    while d <= ACADEMIC_DATE_END:
        if d in holidays:
            for slot in required_slots(d):
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

        for slot in required_slots(d):
            eligible, reasons = eligible_for_slot(
                lookup, residents, rules, no_call, d, slot, intern_names, upper_names
            )
            picked = pick_best_candidate(residents, eligible, d, slot)

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
                name, rotation = picked
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

    audit_data = audit_schedule(
        schedule_rows=schedule_rows,
        residents=residents,
        lookup=lookup,
        rules=rules,
        no_call_days=no_call,
        unassigned_rows=unassigned_rows,
        holidays=holidays,
        seed=seed,
        tiebreaker_count=TIEBREAKER_COUNT,
    )

    return {
        "schedule_rows": schedule_rows,
        "residents": residents,
        "lookup": lookup,
        "holidays": holidays,
        "no_call": no_call,
        "unassigned_rows": unassigned_rows,
        "audit_data": audit_data,
    }

def monte_carlo_score(result):
    audit = result["audit_data"]
    fairness = audit["fairness_summary"]

    return (
        len(audit["errors"]),
        len(result["unassigned_rows"]),
        fairness["upper_weekday_diff"],
        fairness["upper_weekend_diff"],
        fairness["intern_weekend_diff"],
        len(audit.get("avoid_assignments", [])),
        len(audit["warnings"]),
    )

def run_simulation(num_runs=50):
    start = time.time()
    best_result = None
    best_score = None

    for seed in range(num_runs):
        result = generate_schedule_once(seed=seed)
        score = monte_carlo_score(result)

        print(f"Run {seed+1}/{num_runs} | seed={seed} | score={score}")

        if best_score is None or score < best_score:
            best_score = score
            best_result = result

    best_seed = best_result['audit_data']['seed']

    end = time.time()
    print(f"\nSimulation completed in {end-start:.2f} seconds")
    print("Best run selected:")
    print(f"Seed: {best_seed}")
    print(f"Score: {best_score}\n")



    return best_result

def export_result(result):
    schedule_rows = result["schedule_rows"]
    residents = result["residents"]
    lookup = result["lookup"]
    holidays = result["holidays"]
    no_call = result["no_call"]

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
    print(f"  {DATA_DIR}/{OUTPUT_DIR}/call_totals.csv")
    print(f"  {DATA_DIR}/{OUTPUT_DIR}/call_schedule.xlsx")

    audit_data = result["audit_data"]
    write_audit(audit_data, path=f"{DATA_DIR}/{OUTPUT_DIR}/audit_report.txt")
    print(f"Wrote to: {DATA_DIR}/{OUTPUT_DIR}/audit_report.txt")

if __name__ == "__main__":

    #best_result = run_simulation(num_runs=SIMULATION_RUNS)
    #export_result(best_result)

    export_result(generate_schedule_once(seed=211))
