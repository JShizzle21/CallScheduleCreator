import csv
from datetime import date, timedelta
from typing import Dict, Tuple

from config import CONFIG

ROTATION_RULES_CSV = CONFIG.get("ROTATION_RULES_CSV","data/rotation_rules.csv")
NO_CALL_DAYS_CSV = CONFIG.get("NO_CALL_DAYS_CSV","data/no_call_days.csv")
HOLIDAYS_CSV = CONFIG.get("HOLIDAYS_CSV","data/holidays.csv")

# ---------------- Loaders ----------------

def load_residents(lookup) -> Dict[str, dict]:
    ws = lookup.ws
    skip_rows = lookup.skip_rows

    residents = {}
    pgy = 1

    for r in range(3, ws.max_row + 1):
        if r in skip_rows:
            pgy += 1
            continue

        name = ws.cell(row=r, column=1).value
        if not name:
            continue

        name = str(name).strip()

        residents[name] = {
            "pgy": pgy,
            "assigned_dates": [],
            "total_calls": 0,
            "weekday_calls": 0,
            "weekend_calls": 0,
            "upper_calls": 0,
            "intern_calls": 0,
        }

    return residents

def load_rotation_rules(path: str = ROTATION_RULES_CSV) -> Dict[Tuple[str, int], str]:
    """
    rotation_rules.csv columns:
      rotation_name,preference,pgy

    preference: ELIGIBLE | AVOID | NO_CALL
    """
    rules: Dict[Tuple[str, int], str] = {}
    with open(path, newline="", encoding="utf-8") as f:
        reader = csv.DictReader(f)
        for row in reader:
            rotation = row["rotation_name"].strip()
            pgy = int(row["pgy"])
            pref = row["preference"].strip().upper()
            rules[(rotation, pgy)] = pref
    return rules


def load_no_call_days(path: str = NO_CALL_DAYS_CSV) -> Dict[str, dict]:
    out: Dict[str, dict] = {}

    try:
        with open(path, newline="", encoding="utf-8") as f:
            reader = csv.DictReader(f)
            for row in reader:
                name = row["name"].strip()
                start = date.fromisoformat(row["start_date"].strip())
                end = date.fromisoformat(row["end_date"].strip())
                reason = row.get("type", "").strip()

                cur = start
                while cur <= end:
                    out.setdefault(name, {})[cur] = reason
                    cur += timedelta(days=1)

    except FileNotFoundError:
        pass

    return out


def load_holidays(path: str = HOLIDAYS_CSV) -> Dict[date, str]:
    """
    Optional holidays.csv columns:
      date,name
    """
    out: Dict[date, str] = {}
    try:
        with open(path, newline="", encoding="utf-8") as f:
            reader = csv.DictReader(f)
            for row in reader:
                d = date.fromisoformat(row["date"].strip())
                out[d] = row.get("name", "").strip() or "Holiday"
    except FileNotFoundError:
        pass
    return out