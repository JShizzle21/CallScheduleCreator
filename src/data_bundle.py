"""Data bundle loading — consolidates every input-file read into one call.

The CLI builds a bundle from paths derived from `config.yaml`; the GUI builds
one from user-uploaded file paths. In both cases `generate_schedule_once`
consumes the same shape, so the scheduler itself is oblivious to where the
files came from.
"""
from __future__ import annotations

from dataclasses import dataclass
from datetime import date
from typing import Dict, List, Optional, Tuple

from excel_reader import ExcelRotationLookup
from loader import (
    load_clinic_days,
    load_completed_calls,
    load_holidays,
    load_no_call_days,
    load_residents,
    load_rotation_rules,
)


@dataclass
class DataBundle:
    lookup: ExcelRotationLookup
    residents: Dict[str, dict]
    rules: Dict[Tuple[str, int], str]
    # no_call is the merged view (raw no_call_days + clinic-derived pre-call
    # blocks). Eligibility checks only need this merged dict; the distinct
    # reason strings are preserved inside it.
    no_call: Dict[str, dict]
    holidays: Dict[date, dict]
    completed_calls: List[Tuple[date, str, str]]
    block1_end: Optional[date]
    # Original time-off range entries, preserved for the audit report so
    # the user can sanity-check that every supervisor-submitted request
    # was loaded and respected.
    no_call_entries: List[dict] = None


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


def load_data_bundle(
    paths: dict,
    *,
    academic_year_start: int,
    intern_block1_weekday_calls: bool,
    use_completed_calls: bool,
    academic_start_date: Optional[date] = None,
    academic_end_date: Optional[date] = None,
) -> DataBundle:
    """Load every input file referenced by `paths` into a DataBundle.

    `academic_year_start` is the calendar year of ACADEMIC_DATE_START — the
    flow-sheet reader needs it to resolve ambiguous MM/DD date cells.

    `intern_block1_weekday_calls` gates Block 1 intern weekday call parsing:
    when False, weekday intern cells in completed_calls.xlsx are dropped
    silently (they represent Night Float, not scheduled calls).

    `use_completed_calls` toggles the partial-year restart: when False,
    `completed_calls.xlsx` is not read even if the path exists.
    """
    flow_xlsx = paths["flow_xlsx"]
    sheet_name = paths["sheet_name"]

    lookup = ExcelRotationLookup(flow_xlsx, sheet_name, academic_year_start)
    block1_end = (
        lookup.blocks[0].end
        if (intern_block1_weekday_calls and lookup.blocks)
        else None
    )

    residents = load_residents(lookup)
    rules = load_rotation_rules(paths["rotation_rules_xlsx"])
    no_call_base, no_call_entries = load_no_call_days(
        paths["no_call_days_xlsx"],
        valid_residents=residents.keys(),
        academic_start=academic_start_date,
        academic_end=academic_end_date,
        blocks=lookup.blocks,
    )
    clinic_pre_blocks = load_clinic_days(
        paths["clinic_days_xlsx"],
        valid_residents=residents.keys(),
        academic_start=academic_start_date,
        academic_end=academic_end_date,
        blocks=lookup.blocks,
    )
    no_call = _merge_no_call(no_call_base, clinic_pre_blocks)
    holidays = load_holidays(paths["holidays_xlsx"])

    if use_completed_calls:
        completed_calls = load_completed_calls(
            paths["completed_calls_xlsx"],
            block1_end=block1_end,
        )
    else:
        completed_calls = []

    return DataBundle(
        lookup=lookup,
        residents=residents,
        rules=rules,
        no_call=no_call,
        holidays=holidays,
        completed_calls=completed_calls,
        block1_end=block1_end,
        no_call_entries=no_call_entries,
    )
