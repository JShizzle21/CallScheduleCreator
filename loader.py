from datetime import date, datetime, timedelta
from typing import Dict, Tuple, Iterable
from openpyxl import load_workbook

from config import CONFIG

ROTATION_RULES_XLSX = CONFIG.get("ROTATION_RULES_XLSX", "data/rotation_rules.xlsx")
NO_CALL_DAYS_XLSX = CONFIG.get("NO_CALL_DAYS_XLSX", "data/no_call_days.xlsx")
HOLIDAYS_XLSX = CONFIG.get("HOLIDAYS_XLSX", "data/holidays.xlsx")


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


def _normalize_header(value) -> str:
    return str(value).strip().lower() if value is not None else ""


def _parse_date(value) -> date:
    if value is None or value == "":
        raise ValueError("Blank date cell encountered while loading Excel input file.")

    if isinstance(value, datetime):
        return value.date()
    if isinstance(value, date):
        return value

    s = str(value).strip()
    for fmt in ("%Y-%m-%d", "%m/%d/%Y", "%m/%d/%y"):
        try:
            return datetime.strptime(s, fmt).date()
        except ValueError:
            pass

    raise ValueError(f"Could not parse date value '{value}' from Excel input file.")


def _iter_excel_dict_rows(path: str) -> Iterable[dict]:
    wb = load_workbook(path, data_only=True)
    ws = wb.active

    headers = [_normalize_header(cell.value) for cell in ws[1]]
    if not any(headers):
        return

    for r in range(2, ws.max_row + 1):
        row_values = [ws.cell(row=r, column=c).value for c in range(1, len(headers) + 1)]
        if all(v is None or str(v).strip() == "" for v in row_values):
            continue
        yield {headers[i]: row_values[i] for i in range(len(headers))}


def load_rotation_rules(path: str = ROTATION_RULES_XLSX) -> Dict[Tuple[str, int], str]:
    rules: Dict[Tuple[str, int], str] = {}
    for row in _iter_excel_dict_rows(path):
        rotation = str(row["rotation_name"]).strip()
        pgy = int(row["pgy"])
        pref = str(row["preference"]).strip().upper()
        rules[(rotation, pgy)] = pref
    return rules


def load_no_call_days(path: str = NO_CALL_DAYS_XLSX) -> Dict[str, dict]:
    out: Dict[str, dict] = {}
    try:
        for row in _iter_excel_dict_rows(path):
            name = str(row["name"]).strip()
            start = _parse_date(row["start_date"])
            end = _parse_date(row["end_date"])
            reason = str(row.get("type", "") or "").strip()

            cur = start
            while cur <= end:
                out.setdefault(name, {})[cur] = reason
                cur += timedelta(days=1)
    except FileNotFoundError:
        pass

    return out


def load_holidays(path: str = HOLIDAYS_XLSX) -> Dict[date, str]:
    out: Dict[date, str] = {}
    try:
        for row in _iter_excel_dict_rows(path):
            d = _parse_date(row["date"])
            out[d] = str(row.get("name", "") or "").strip() or "Holiday"
    except FileNotFoundError:
        pass
    return out
