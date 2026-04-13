from typing import Dict, Tuple
from datetime import datetime
from collections import defaultdict
from datetime import date, timedelta

from config import CONFIG

DATA_DIR = CONFIG.get("DATA_DIR", "data")
OUTPUT_DIR = CONFIG.get("OUTPUT_DIR", "output")
POST_CALL_DAYS = CONFIG.get("POST_CALL_DAYS", 2)
INTERN_BLOCK1_WEEKDAY_CALLS = int(CONFIG.get("INTERN_BLOCK1_WEEKDAY_CALLS", 0))
ACADEMIC_DATE_START_STRING = CONFIG.get("ACADEMIC_DATE_START_STRING", "2026-07-01")
ACADEMIC_DATE_END_STRING = CONFIG.get("ACADEMIC_DATE_END_STRING", "2027-06-30")
ACADEMIC_DATE_START = datetime.strptime(ACADEMIC_DATE_START_STRING, "%Y-%m-%d").date()
ACADEMIC_DATE_END = datetime.strptime(ACADEMIC_DATE_END_STRING, "%Y-%m-%d").date()


def is_weekend(d: date) -> bool:
    return d.weekday() >= 5


def validate_rotations_against_rules(lookup, residents: Dict[str, dict], rules: Dict[Tuple[str, int], str]) -> None:
    missing = set()

    for name, data in residents.items():
        pgy = data["pgy"]

        for block in lookup.blocks:
            raw_rotation = lookup.rotation_grid.get(name, {}).get(block.col, "")
            if not raw_rotation:
                continue

            parts = [part.strip() for part in __import__("re").split(r"[\\/]", raw_rotation) if part.strip()]

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
            "in rotation rules for the correct PGY."
        )


def validate_no_call_days(no_call_days: Dict[str, dict], residents: Dict[str, dict]) -> None:
    missing_names = sorted(name for name in no_call_days.keys() if name not in residents)

    if missing_names:
        print("\nWARNING: Names in no-call file not found in flow sheet:\n")
        for name in missing_names:
            print(f"  {name}")

        raise ValueError(
            "\nValidation failed: one or more names in the no-call input do not match "
            "the resident names in the flow sheet."
        )


def build_rotation_date_summary(lookup, residents: Dict[str, dict]) -> Dict[str, list[dict]]:
    summary: Dict[str, list[dict]] = {}
    for resident in sorted(residents.keys()):
        rows = []
        for seg in lookup.rotation_segments_for_resident(resident):
            rows.append(
                {
                    "block": seg.block_index,
                    "part": seg.part_index,
                    "parts_total": seg.total_parts,
                    "rotation": seg.rotation,
                    "start": seg.start.isoformat(),
                    "end": seg.end.isoformat(),
                }
            )
        summary[resident] = rows
    return summary


def audit_schedule(
    schedule_rows,
    residents,
    lookup,
    rules,
    no_call_days,
    unassigned_rows,
    holidays,
    seed,
    tiebreaker_count,
):
    errors = []
    warnings = []

    by_date = defaultdict(list)
    by_resident = defaultdict(list)

    for row in schedule_rows:
        d = date.fromisoformat(row["date"])
        resident = (row.get("resident") or "").strip()
        slot = row["slot"]

        by_date[d].append(row)

        if resident:
            by_resident[resident].append((d, slot))

    block1_end = lookup.blocks[0].end if (INTERN_BLOCK1_WEEKDAY_CALLS and lookup.blocks) else None

    d = ACADEMIC_DATE_START
    while d <= ACADEMIC_DATE_END:
        if d in holidays:
            d += timedelta(days=1)
            continue

        rows = by_date.get(d, [])

        upper_count = sum(1 for r in rows if r["slot"] in ("UPPER_WEEKDAY", "UPPER_WEEKEND") and r.get("resident"))
        intern_count = sum(1 for r in rows if r["slot"] == "INTERN_WEEKEND" and r.get("resident"))
        intern_weekday_count = sum(1 for r in rows if r["slot"] == "INTERN_WEEKDAY" and r.get("resident"))

        if is_weekend(d):
            if upper_count != 1:
                errors.append(f"{d}: expected 1 upper weekend assignment, found {upper_count}")
            if intern_count != 1:
                errors.append(f"{d}: expected 1 intern weekend assignment, found {intern_count}")
        else:
            if upper_count != 1:
                errors.append(f"{d}: expected 1 upper weekday assignment, found {upper_count}")
            if block1_end is not None and d <= block1_end:
                if intern_weekday_count != 1:
                    errors.append(f"{d}: expected 1 intern weekday assignment (Block 1), found {intern_weekday_count}")
            else:
                if intern_count != 0 or intern_weekday_count != 0:
                    errors.append(
                        f"{d}: weekday should not have intern assignment, "
                        f"found {intern_count + intern_weekday_count}"
                    )

        d += timedelta(days=1)

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

            if resident in assigned_today:
                errors.append(f"{d}: {resident} assigned more than once on same day")
            assigned_today.add(resident)

            if slot in ("INTERN_WEEKEND", "INTERN_WEEKDAY") and pgy != 1:
                errors.append(f"{d}: {resident} assigned to intern slot ({slot}) but is PGY{pgy}")
            if slot in ("UPPER_WEEKDAY", "UPPER_WEEKEND") and pgy == 1:
                errors.append(f"{d}: {resident} assigned to upper slot but is PGY1")

            # COMPLETED rows are accepted as ground truth — skip constraint
            # checks that the user may have manually overridden.
            is_completed = row.get("note") == "COMPLETED"
            if is_completed:
                continue

            resident_days = no_call_days.get(resident, {})
            if d in resident_days:
                errors.append(f"{d}: {resident} assigned on no-call day")

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

    for resident, assignments in by_resident.items():
        assignments = sorted(assignments, key=lambda x: x[0])

        for i in range(1, len(assignments)):
            prev_date = assignments[i - 1][0]
            curr_date = assignments[i][0]

            if (curr_date - prev_date).days <= POST_CALL_DAYS:
                errors.append(f"{resident}: assigned on consecutive/post-call restricted days {prev_date} and {curr_date}")

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

    uppers = [r for r in residents.values() if r["pgy"] in (2, 3)]
    interns = [r for r in residents.values() if r["pgy"] == 1]

    upper_weekday_counts = [r["weekday_calls"] for r in uppers]
    upper_weekend_counts = [r["weekend_calls"] for r in uppers]
    intern_weekend_counts = [r["weekend_calls"] for r in interns]
    upper_total_counts = [r["total_calls"] for r in uppers]

    fairness_summary = {
        "upper_total_min": min(upper_total_counts) if upper_total_counts else 0,
        "upper_total_max": max(upper_total_counts) if upper_total_counts else 0,
        "upper_total_diff": (max(upper_total_counts) - min(upper_total_counts)) if upper_total_counts else 0,

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

    avoid_assignments = []

    for d, rows in by_date.items():
        for row in rows:
            if row.get("note") == "COMPLETED":
                continue

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

    weekend_call_monthly = defaultdict(int)
    weekend_call_overages = []

    for resident, assignments in by_resident.items():
        for assignment_date, slot in assignments:
            if not is_weekend(assignment_date):
                continue
            month_key = assignment_date.strftime("%Y-%m")
            weekend_call_monthly[(resident, month_key)] += 1

    for (resident, month_key), count in sorted(weekend_call_monthly.items()):
        if count > 4:
            weekend_call_overages.append(
                {
                    "resident": resident,
                    "month": month_key,
                    "weekend_calls": count,
                }
            )
            warnings.append(
                f"{resident}: {count} weekend call shifts in {month_key} (possible <4 days off concern)"
            )

    rotation_date_summary = build_rotation_date_summary(lookup, residents)

    skipped_rows = sorted(lookup.skip_rows)

    return {
        "errors": errors,
        "warnings": warnings,
        "fairness_summary": fairness_summary,
        "avoid_assignments": avoid_assignments,
        "seed": seed,
        "tiebreaker_count": tiebreaker_count,
        "unassigned_rows": unassigned_rows,
        "weekend_call_overages": weekend_call_overages,
        "rotation_date_summary": rotation_date_summary,
        "skipped_rows": skipped_rows,
    }
