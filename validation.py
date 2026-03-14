from typing import Dict, Tuple
from datetime import datetime
from collections import defaultdict
from datetime import date, timedelta

from config import CONFIG

DATA_DIR = CONFIG.get("DATA_DIR", "data")
OUTPUT_DIR = CONFIG.get("OUTPUT_DIR","output")
POST_CALL_DAYS = CONFIG.get("POST_CALL_DAYS", 2)
ACADEMIC_DATE_START_STRING = CONFIG.get("ACADEMIC_DATE_START_STRING","2026-07-01")
ACADEMIC_DATE_END_STRING = CONFIG.get("ACADEMIC_DATE_END_STRING","2027-06-30")
ACADEMIC_DATE_START = datetime.strptime(ACADEMIC_DATE_START_STRING, "%Y-%m-%d").date()
ACADEMIC_DATE_END = datetime.strptime(ACADEMIC_DATE_END_STRING, "%Y-%m-%d").date()

def is_weekend(d: date) -> bool:
    return d.weekday() >= 5  # Sat/Sun

def validate_rotations_against_rules(lookup, residents: Dict[str, dict], rules: Dict[Tuple[str, int], str]) -> None:
    missing = set()

    for name, data in residents.items():
        pgy = data["pgy"]

        for block in lookup.blocks:
            raw_rotation = lookup.rotation_grid.get(name, {}).get(block.col, "")
            if not raw_rotation:
                continue

            parts = [part.strip() for part in raw_rotation.split("/")]

            for rotation in parts:
                key = (rotation, pgy)
                if key not in rules:
                    missing.add((name, pgy, rotation))

    if missing:
        print("\nWARNING: Missing rotation rules found:\n")
        for name, pgy, rotation in sorted(missing):
            print(f"  Resident: {name}, PGY: {pgy}, Rotation: {rotation}")

        raise ValueError(
            "\nValidation failed: one or more rotations in the flow sheet do not exist "
            "in rotation_rules.csv for the correct PGY."
        )
    else:
        #print("Rotation validation passed.")
        return


def validate_no_call_days(no_call_days: Dict[str, dict], residents: Dict[str, dict]) -> None:
    missing_names = sorted(name for name in no_call_days.keys() if name not in residents)

    if missing_names:
        print("\nWARNING: Names in no_call_days.csv not found in flow sheet:\n")
        for name in missing_names:
            print(f"  {name}")

        raise ValueError(
            "\nValidation failed: one or more names in no_call_days.csv do not match "
            "the resident names in the flow sheet."
        )
    else:
        #print("No-call name validation passed.\n")
        return

def audit_schedule(
        schedule_rows,
        residents,
        lookup,
        rules,
        no_call_days,
        unassigned_rows,
        holidays,
        seed,
        tiebreaker_count
):
    """
    Audits the generated schedule for:
    - coverage issues
    - PGY/slot rule violations
    - no-call day violations
    - NO_CALL rotation violations
    - post-call violations
    - duplicate same-day assignments
    - spacing metrics
    - fairness summaries
    """


    errors = []
    warnings = []

    # Organize schedule by date
    by_date = defaultdict(list)
    by_resident = defaultdict(list)

    for row in schedule_rows:
        d = date.fromisoformat(row["date"])
        resident = (row.get("resident") or "").strip()
        slot = row["slot"]

        by_date[d].append(row)

        if resident:
            by_resident[resident].append((d, slot))

    # --------------------------------------------------
    # 1. Coverage audit
    # --------------------------------------------------
    d = ACADEMIC_DATE_START
    while d <= ACADEMIC_DATE_END:

        #Skip holidays while verifying dates
        if d in holidays:
            d += timedelta(days=1)
            continue

        rows = by_date.get(d, [])

        upper_count = sum(1 for r in rows if r["slot"] in ("UPPER_WEEKDAY", "UPPER_WEEKEND") and r.get("resident"))
        intern_count = sum(1 for r in rows if r["slot"] == "INTERN_WEEKEND" and r.get("resident"))



        if is_weekend(d):
            if upper_count != 1:
                errors.append(f"{d}: expected 1 upper weekend assignment, found {upper_count}")
            if intern_count != 1:
                errors.append(f"{d}: expected 1 intern weekend assignment, found {intern_count}")
        else:
            if upper_count != 1:
                errors.append(f"{d}: expected 1 upper weekday assignment, found {upper_count}")
            if intern_count != 0:
                errors.append(f"{d}: weekday should not have intern assignment, found {intern_count}")

        d += timedelta(days=1)

    # --------------------------------------------------
    # 2. Per-assignment hard rule audit
    # --------------------------------------------------
    for d, rows in by_date.items():
        assigned_today = set()

        for row in rows:
            resident = (row.get("resident") or "").strip()
            if not resident:
                continue

            slot = row["slot"]

            if resident not in residents:
                errors.append(f"{d}: resident '{resident}' not found in resident list")
                continue

            pgy = residents[resident]["pgy"]

            # Duplicate same-day assignment
            if resident in assigned_today:
                errors.append(f"{d}: {resident} assigned more than once on same day")
            assigned_today.add(resident)

            # PGY/slot check
            if slot == "INTERN_WEEKEND" and pgy != 1:
                errors.append(f"{d}: {resident} assigned to intern weekend slot but is PGY{pgy}")
            if slot in ("UPPER_WEEKDAY", "UPPER_WEEKEND") and pgy == 1:
                errors.append(f"{d}: {resident} assigned to upper slot but is PGY1")

            # No-call day check
            resident_days = no_call_days.get(resident, {})
            if d in resident_days:
                errors.append(f"{d}: {resident} assigned on no-call day")

            # Rotation rule check
            rotation = lookup.rotation_on_date(resident, d)
            if rotation is None:
                errors.append(f"{d}: {resident} has no rotation found")
                continue

            rule = rules.get((rotation, pgy))
            if rule is None:
                errors.append(f"{d}: {resident} on rotation '{rotation}' PGY{pgy} missing from rotation rules")
                continue

            pref = rule["preference"] if isinstance(rule, dict) else rule
            if pref == "NO_CALL":
                errors.append(f"{d}: {resident} assigned on NO_CALL rotation '{rotation}'")

    # --------------------------------------------------
    # 3. Post-call / consecutive day audit
    # --------------------------------------------------
    for resident, assignments in by_resident.items():
        assignments = sorted(assignments, key=lambda x: x[0])

        for i in range(1, len(assignments)):
            prev_date = assignments[i - 1][0]
            curr_date = assignments[i][0]

            if (curr_date - prev_date).days <= POST_CALL_DAYS:
                errors.append(f"{resident}: assigned on consecutive days {prev_date} and {curr_date}")

    # --------------------------------------------------
    # 4. Spacing audit (more than one call in 7 days)
    # --------------------------------------------------
    under_7 = []

    for resident, assignments in by_resident.items():
        assignments = sorted(assignments, key=lambda x: x[0])

        for i in range(1, len(assignments)):
            prev_date = assignments[i - 1][0]
            curr_date = assignments[i][0]
            spacing = (curr_date - prev_date).days

            if spacing < 7:
                under_7.append((resident, prev_date, curr_date, spacing))

    for resident, prev_date, curr_date, spacing in under_7:
        warnings.append(f"{resident}: spacing < 7 days ({prev_date} -> {curr_date}, {spacing} days)")

    # --------------------------------------------------
    # 5. Fairness summary
    # --------------------------------------------------
    uppers = [r for r in residents.values() if r["pgy"] in (2, 3)]
    interns = [r for r in residents.values() if r["pgy"] == 1]

    upper_weekday_counts = [r["weekday_calls"] for r in uppers]
    upper_weekend_counts = [r["weekend_calls"] for r in uppers]
    intern_weekend_counts = [r["weekend_calls"] for r in interns]

    fairness_summary = {
        "upper_weekday_min": min(upper_weekday_counts) if upper_weekday_counts else 0,
        "upper_weekday_max": max(upper_weekday_counts) if upper_weekday_counts else 0,
        "upper_weekday_diff": (max(upper_weekday_counts) - min(upper_weekday_counts)) if upper_weekday_counts else 0,

        "upper_weekend_min": min(upper_weekend_counts) if upper_weekend_counts else 0,
        "upper_weekend_max": max(upper_weekend_counts) if upper_weekend_counts else 0,
        "upper_weekend_diff": (max(upper_weekend_counts) - min(upper_weekend_counts)) if upper_weekend_counts else 0,

        "intern_weekend_min": min(intern_weekend_counts) if intern_weekend_counts else 0,
        "intern_weekend_max": max(intern_weekend_counts) if intern_weekend_counts else 0,
        "intern_weekend_diff": (max(intern_weekend_counts) - min(intern_weekend_counts)) if intern_weekend_counts else 0,
    }

    # --------------------------------------------------
    # 7. UNASSIGNED COUNT
    # --------------------------------------------------
    unassigned_count = len(unassigned_rows)

    # --------------------------------------------------
    # 8. AVOID ASSIGNMENTS
    # --------------------------------------------------

    avoid_assignments = []

    for d, rows in by_date.items():
        for row in rows:
            resident = (row.get("resident") or "").strip()
            if not resident:
                continue

            pgy = residents[resident]["pgy"]
            rotation = lookup.rotation_on_date(resident, d)
            if rotation is None:
                continue

            rule = rules.get((rotation, pgy))
            if rule is None:
                continue

            pref = rule["preference"] if isinstance(rule, dict) else rule
            if pref == "AVOID":
                avoid_assignments.append((d, resident, rotation, row["slot"]))



    return {
        "errors": errors,
        "warnings": warnings,
        "fairness_summary": fairness_summary,
        "avoid_assignments": avoid_assignments,
        "seed": seed,
        "tiebreaker_count": tiebreaker_count,
        "unassigned_rows": unassigned_rows
    }
