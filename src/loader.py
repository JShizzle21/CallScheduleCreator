import warnings
from datetime import date, datetime, timedelta
from typing import Dict, List, Optional, Tuple, Iterable
from openpyxl import load_workbook

from config import CONFIG
from errors import DataValidationError

ROTATION_RULES_XLSX = CONFIG.get("ROTATION_RULES_XLSX", "data/rotation_rules.xlsx")
NO_CALL_DAYS_XLSX = CONFIG.get("NO_CALL_DAYS_XLSX", "data/no_call_days.xlsx")
HOLIDAYS_XLSX = CONFIG.get("HOLIDAYS_XLSX", "data/holidays.xlsx")
CLINIC_DAYS_XLSX = CONFIG.get("CLINIC_DAYS_XLSX", "data/clinic_days.xlsx")
COMPLETED_CALLS_XLSX = CONFIG.get("COMPLETED_CALLS_XLSX", "")


def load_residents(lookup) -> Dict[str, dict]:
    ws = lookup.ws
    skip_rows = lookup.skip_rows
    separator_pgy_labels: dict = getattr(lookup, "separator_pgy_labels", {})

    residents = {}
    pgy = 1
    separator_count = 0

    for r in range(3, ws.max_row + 1):
        if r in skip_rows:
            separator_count += 1
            explicit_pgy = separator_pgy_labels.get(r)
            if explicit_pgy is not None:
                pgy = explicit_pgy
            else:
                pgy += 1
                if pgy > 3:
                    raise ValueError(
                        f"Flow sheet has {separator_count} PGY separator rows, implying "
                        f"PGY{pgy} which is unexpected. Expected at most 2 separators "
                        f"(for PGY1/PGY2/PGY3). Check for extra header rows, or add "
                        f"explicit PGY labels (e.g. 'PGY2') in column A of separator rows."
                    )
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
            "friday_calls": 0,
            "saturday_calls": 0,
            "sunday_calls": 0,
            "upper_calls": 0,
            "intern_calls": 0,
            "Jul_Dec_calls": 0,
            "Jan_Jun_calls": 0,
        }

    return residents


def _normalize_header(value) -> str:
    return str(value).strip().lower() if value is not None else ""


_DATE_FORMATS = (
    "%Y-%m-%d",        # 2026-07-01
    "%Y/%m/%d",        # 2026/07/01
    "%m/%d/%Y",        # 7/1/2026
    "%m/%d/%y",        # 7/1/26
    "%m-%d-%Y",        # 7-1-2026
    "%m-%d-%y",        # 7-1-26
    "%B %d %Y",        # July 1 2026
    "%B %d, %Y",       # July 1, 2026
    "%d %B %Y",        # 1 July 2026
    "%d-%B-%Y",        # 1-July-2026
    "%b %d %Y",        # Jul 1 2026
    "%b %d, %Y",       # Jul 1, 2026
    "%d %b %Y",        # 1 Jul 2026
    "%d-%b-%Y",        # 1-Jul-2026
)


def _parse_date(value) -> date:
    if value is None or value == "":
        raise ValueError("Blank date cell encountered while loading Excel input file.")

    if isinstance(value, datetime):
        return value.date()
    if isinstance(value, date):
        return value

    s = str(value).strip()
    # Collapse internal whitespace so "July  1,  2026" still parses.
    s_clean = " ".join(s.split())
    for fmt in _DATE_FORMATS:
        try:
            return datetime.strptime(s_clean, fmt).date()
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


def load_holidays(path: str = HOLIDAYS_XLSX) -> Dict[date, dict]:
    """Load holiday dates plus optional manually-assigned residents.

    Columns:
      - ``date`` (required): the holiday date.
      - ``name`` (optional): display name (e.g. "Christmas"); defaults to "Holiday".
      - ``upper`` (optional): name of the upper-level resident who works this
        holiday. Blank = unassigned (audit will flag).
      - ``intern`` (optional): name of the intern who works this holiday. On
        weekday holidays outside Block 1 the intern override replaces Night
        Float for that day. Blank = unassigned.

    Returns ``Dict[date, {"name": str, "upper": str|None, "intern": str|None}]``.
    Membership checks (``if d in holidays``) keep working as before.
    """
    out: Dict[date, dict] = {}
    try:
        for row in _iter_excel_dict_rows(path):
            d = _parse_date(row["date"])
            name = str(row.get("name", "") or "").strip() or "Holiday"
            upper = str(row.get("upper", "") or "").strip() or None
            intern = str(row.get("intern", "") or "").strip() or None
            out[d] = {"name": name, "upper": upper, "intern": intern}
    except FileNotFoundError:
        pass
    return out


CLINIC_HEADER_SCAN_ROWS = 10
EXPECTED_CLINIC_BLOCK_COUNT = 13

# Body-row cell values that match these patterns are silently ignored — they
# are header-like labels (or stray copy-paste artefacts) from the
# supervisor's working doc, not resident names.
_CLINIC_IGNORED_BODY_LABELS = {"date", "intern", "interns"}


# Module-level set of clinic-warning keys already emitted in this process.
# Prevents the same '?' warning from firing once per Monte Carlo seed when
# run_simulation rebuilds the data bundle per run.
# (Each ProcessPoolExecutor worker has its own copy, so the warning will
# still fire once per worker process — acceptably small.)
_clinic_warning_seen: set = set()

# Also register Python's own per-process "once" filter as a belt-and-braces
# guard in case our explicit set is somehow bypassed.
warnings.filterwarnings(
    "once",
    message=r".*uncertain entry.*Confirm with supervisor\.",
    category=UserWarning,
)


def _is_ignored_clinic_label(value: str) -> bool:
    """True if a body-row cell value is a known non-name label to skip."""
    lower = value.lower()
    if lower in _CLINIC_IGNORED_BODY_LABELS:
        return True
    # 'Amb - <whatever>' is a rotation/group marker the supervisor places in
    # the sheet — never a resident name on its own.
    if lower.startswith("amb -") or lower.startswith("amb-"):
        return True
    return False


def _is_parseable_date(value) -> bool:
    """Lightweight helper: True iff value parses as a date via _parse_date."""
    if value is None or value == "":
        return False
    try:
        _parse_date(value)
        return True
    except ValueError:
        return False


def _find_clinic_date_column(ws, sheet_name: str) -> Optional[Tuple[int, int]]:
    """Locate the 'date' header in a clinic-days block sheet.

    Scans the first CLINIC_HEADER_SCAN_ROWS rows for any cell whose
    case-insensitive trimmed value equals 'date'. Each candidate is
    classified by the cell directly below:
      - Parses as a date            → CONFIRMED (the right column to use).
      - Blank/empty                 → AMBIGUOUS (could be an empty sheet).
      - Non-empty, not a date       → BAD (suspicious — header is misplaced
                                       or the column is mislabelled).

    Returns:
      - (header_row, date_col) for the first confirmed candidate.
      - None if no 'date' header exists at all OR all 'date' headers are
        AMBIGUOUS (treated as "this block has no clinics yet").

    Raises DataValidationError only when a 'date' header has non-date data
    below it AND no other column confirms — that pattern usually means the
    user has the date column in the wrong place.
    """
    confirmed: Optional[Tuple[int, int]] = None
    bad: List[Tuple[int, int, object]] = []
    ambiguous_count = 0

    max_scan_row = min(CLINIC_HEADER_SCAN_ROWS, ws.max_row)
    for r in range(1, max_scan_row + 1):
        for c in range(1, ws.max_column + 1):
            v = ws.cell(row=r, column=c).value
            if v is None:
                continue
            if str(v).strip().lower() != "date":
                continue
            below = ws.cell(row=r + 1, column=c).value
            if below is None or str(below).strip() == "":
                ambiguous_count += 1
                continue
            if _is_parseable_date(below):
                if confirmed is None:
                    confirmed = (r, c)
            else:
                bad.append((r, c, below))

    if confirmed is not None:
        return confirmed

    if bad:
        details = ", ".join(
            f"row {r} col {c} (below = {value!r})" for r, c, value in bad
        )
        raise DataValidationError(
            f"clinic_days.xlsx sheet '{sheet_name}': found 'Date' header(s) "
            f"at {details}, but the cell below is not a recognizable date. "
            f"Either fix the value or move the 'Date' header to the correct "
            f"column."
        )

    # No confirmed and no bad. Either there were only ambiguous 'date'
    # headers (empty block), or no 'date' header at all (also treated as
    # empty — the loader will still error if the sheet has stray names
    # elsewhere).
    return None


def load_clinic_days(
    path: str = CLINIC_DAYS_XLSX,
    *,
    valid_residents: Optional[Iterable[str]] = None,
    academic_start: Optional[date] = None,
    academic_end: Optional[date] = None,
    blocks: Optional[List] = None,
) -> Dict[str, Dict[date, str]]:
    """Load clinic days from a 13-sheet workbook → day-before no-call map.

    Residency rule: a resident cannot take call the night before a clinic
    because the following morning they must be present in clinic. The day
    before each clinic date is blocked with reason 'pre_clinic_day'.

    Workbook format:
      - Sheets named 'Block 1' through 'Block 13' (one per academic block).
      - Each sheet has a 'Date' column (case-insensitive header) somewhere in
        the first 10 rows; the column may be preceded by other columns (e.g.
        day-of-week) that are ignored.
      - The cell directly below the 'Date' header confirms the column —
        it must parse as a date.
      - Rows below the header: column = a clinic date; cells to the right
        list the resident names with clinic that day. Empty cells are
        skipped. Cells to the left of the Date column are ignored.
      - If a row's date cell is blank or unparseable, the row is skipped
        silently (allows in-progress edits).

    Validation (only when the corresponding params are supplied):
      - `valid_residents`: any name not in this set → DataValidationError.
        Also caps the rightward scan at `len(valid_residents)` cells past
        the Date column.
      - `academic_start`/`academic_end`: clinic dates outside this inclusive
        range → DataValidationError.
      - `blocks`: list of Block objects (with .start/.end). A date on sheet
        'Block N' that falls outside blocks[N-1]'s range → DataValidationError.

    If the file does not exist, returns an empty dict (clinics are optional).
    """
    out: Dict[str, Dict[date, str]] = {}
    try:
        wb = load_workbook(path, data_only=True)
    except FileNotFoundError:
        return out

    name_set = set(valid_residents) if valid_residents is not None else None
    resident_count = len(name_set) if name_set is not None else None

    sheetnames_lower = {s.lower(): s for s in wb.sheetnames}

    for block_idx in range(1, EXPECTED_CLINIC_BLOCK_COUNT + 1):
        expected = f"Block {block_idx}"
        actual_name = sheetnames_lower.get(expected.lower())
        if actual_name is None:
            raise DataValidationError(
                f"clinic_days.xlsx is missing sheet '{expected}'. "
                f"The workbook must contain 13 sheets named 'Block 1' "
                f"through 'Block 13'."
            )

        ws = wb[actual_name]
        located = _find_clinic_date_column(ws, actual_name)
        if located is None:
            # No confirmed date column. Treat as empty block UNLESS the sheet
            # has stray non-header content past the scan rows, which would
            # suggest data is present but the Date column is missing.
            scan_floor = CLINIC_HEADER_SCAN_ROWS + 1
            has_stray = False
            for r in range(scan_floor, ws.max_row + 1):
                for c in range(1, ws.max_column + 1):
                    v = ws.cell(row=r, column=c).value
                    if v is not None and str(v).strip() != "":
                        has_stray = True
                        break
                if has_stray:
                    break
            if has_stray:
                raise DataValidationError(
                    f"clinic_days.xlsx sheet '{actual_name}': contains data "
                    f"below row {CLINIC_HEADER_SCAN_ROWS} but no 'Date' "
                    f"header with a recognizable date value could be found "
                    f"in the first {CLINIC_HEADER_SCAN_ROWS} rows. Add a "
                    f"'Date' column header with a valid date below it."
                )
            continue
        header_row, date_col = located

        # Restrict rightward scan: never more cells than total residents.
        if resident_count is not None:
            scan_end = min(ws.max_column, date_col + resident_count)
        else:
            scan_end = ws.max_column

        block = blocks[block_idx - 1] if blocks is not None and len(blocks) >= block_idx else None

        for r in range(header_row + 1, ws.max_row + 1):
            date_cell = ws.cell(row=r, column=date_col).value
            if date_cell is None or str(date_cell).strip() == "":
                continue
            try:
                clinic_date = _parse_date(date_cell)
            except ValueError:
                # Non-date value in a body row — treat as in-progress / ignore.
                continue

            if academic_start is not None and academic_end is not None:
                if clinic_date < academic_start or clinic_date > academic_end:
                    raise DataValidationError(
                        f"clinic_days.xlsx sheet '{actual_name}' row {r}: "
                        f"clinic date {clinic_date} is outside the academic "
                        f"year ({academic_start} to {academic_end})."
                    )

            if block is not None:
                if clinic_date < block.start or clinic_date > block.end:
                    raise DataValidationError(
                        f"clinic_days.xlsx sheet '{actual_name}' row {r}: "
                        f"clinic date {clinic_date} is outside this block's "
                        f"calendar range ({block.start} to {block.end}). "
                        f"Move this row to the correct block sheet."
                    )

            blocked_date = clinic_date - timedelta(days=1)
            for c in range(date_col + 1, scan_end + 1):
                v = ws.cell(row=r, column=c).value
                if v is None or str(v).strip() == "":
                    continue
                raw = str(v).strip()

                # Skip header-like labels that bleed into body cells
                # (working-doc artefacts: stray 'Date', 'Intern',
                # 'Amb - <something>', etc.).
                if _is_ignored_clinic_label(raw):
                    continue

                # Trailing '?' marks an uncertain entry from the supervisor's
                # working doc. Strip the '?' for matching purposes (the
                # source file is left untouched), still apply the call, and
                # emit a one-time warning so the user can confirm later.
                # The '?' stays in the supervisor's spreadsheet; only our
                # in-memory copy is stripped.
                name = raw
                if name.endswith("?"):
                    name = name.rstrip("?").strip()
                    key = (actual_name, r, raw)
                    if key not in _clinic_warning_seen:
                        _clinic_warning_seen.add(key)
                        warnings.warn(
                            f"clinic_days.xlsx sheet '{actual_name}' row {r}: "
                            f"uncertain entry '{raw}' - treating as '{name}'. "
                            f"Confirm with supervisor.",
                            stacklevel=2,
                        )

                if name_set is not None and name not in name_set:
                    raise DataValidationError(
                        f"clinic_days.xlsx sheet '{actual_name}' row {r}: "
                        f"resident name '{name}' does not match any resident "
                        f"in the flow sheet. Names must match exactly."
                    )
                out.setdefault(name, {})[blocked_date] = "pre_clinic_day"

    return out


def load_completed_calls(path: str = COMPLETED_CALLS_XLSX, block1_end: Optional[date] = None) -> List[Tuple[date, str, str]]:
    """Load a partially-completed call schedule to seed a mid-year restart.

    File format: one row per call day, with columns:
      - a 'date' column (any header containing 'date')
      - an upper-level column (any header containing 'upper')
      - an intern column (any header containing 'intern')
    Blank cells mean no assignment for that slot on that date.

    Slot types are inferred from the date: weekday → UPPER_WEEKDAY,
    weekend upper → UPPER_WEEKEND, weekend intern → INTERN_WEEKEND.

    Returns a list of (date, resident_name, slot) tuples sorted by date.
    Returns an empty list if the file does not exist or has no data.
    """
    if not path:
        return []

    out: List[Tuple[date, str, str]] = []
    try:
        wb = load_workbook(path, data_only=True)
        ws = wb.active

        headers = [_normalize_header(cell.value) for cell in ws[1]]
        if not any(headers):
            return []

        date_col: Optional[int] = next(
            (i for i, h in enumerate(headers) if "date" in h), None
        )
        upper_col: Optional[int] = next(
            (i for i, h in enumerate(headers) if "upper" in h), None
        )
        intern_col: Optional[int] = next(
            (i for i, h in enumerate(headers) if "intern" in h), None
        )

        if date_col is None:
            raise ValueError(
                f"No column containing 'date' found in {path}. "
                f"Headers found: {headers}"
            )

        for r in range(2, ws.max_row + 1):
            row_vals = [
                ws.cell(row=r, column=c + 1).value for c in range(len(headers))
            ]
            if all(v is None or str(v).strip() == "" for v in row_vals):
                continue

            try:
                d = _parse_date(row_vals[date_col])
            except (ValueError, TypeError):
                continue

            if upper_col is not None:
                upper_name = row_vals[upper_col]
                if upper_name and str(upper_name).strip():
                    slot = "UPPER_WEEKEND" if d.weekday() >= 5 else "UPPER_WEEKDAY"
                    out.append((d, str(upper_name).strip(), slot))

            if intern_col is not None:
                intern_name = row_vals[intern_col]
                if intern_name and str(intern_name).strip():
                    if d.weekday() >= 5:
                        out.append((d, str(intern_name).strip(), "INTERN_WEEKEND"))
                    elif block1_end is not None and d <= block1_end:
                        # Block 1 intern weekday call (feature enabled)
                        out.append((d, str(intern_name).strip(), "INTERN_WEEKDAY"))
                    # else: weekday intern = Night Float — silently dropped

    except FileNotFoundError:
        pass

    return sorted(out, key=lambda x: x[0])
