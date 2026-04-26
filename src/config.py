from __future__ import annotations

from pathlib import Path

from ruamel.yaml import YAML

# Resolve config.yaml relative to this file (src/config.yaml), not cwd.
# scheduler_main.py runs from project root with cwd=root, but the GUI is
# launched via `streamlit run src/app.py` from project root too — and tests
# invoke from various cwds. Anchoring to __file__ removes the ambiguity.
CONFIG_PATH = str(Path(__file__).resolve().parent / "config.yaml")
LEGACY_GUI_CONFIG_PATH = str(Path(__file__).resolve().parent.parent / "gui_config.yaml")

REQUIRED_KEYS = [
    "ACADEMIC_DATE_START_STRING",
    "ACADEMIC_DATE_END_STRING",
    "POST_CALL_DAYS",
    "SIMULATION_RUNS",
    "MIN_SPACING_DAYS_STRONG",
    "MIN_SPACING_DAYS_MILD",
    "MAX_DIFF_SOFT",
    "MAX_DIFF_HARD",
    "FAIRNESS_GAP_WEIGHT",
    "SPACING_WEIGHT",
    "AVOID_WEIGHT",
    "YEAR_BIAS_WEIGHT",
    "PACE_WEIGHT",
    "LOOKAHEAD_WEIGHT",
]

# Mapping from config.yaml path keys to the internal lowercase names used in the
# paths dict. Keeping this in one place makes the CLI → paths conversion explicit
# and lets the GUI build paths from user uploads using the same lowercase keys.
_PATH_KEY_MAP = {
    "FLOW_XLSX": "flow_xlsx",
    "SHEET_NAME": "sheet_name",
    "ROTATION_RULES_XLSX": "rotation_rules_xlsx",
    "NO_CALL_DAYS_XLSX": "no_call_days_xlsx",
    "HOLIDAYS_XLSX": "holidays_xlsx",
    "CLINIC_DAYS_XLSX": "clinic_days_xlsx",
    "COMPLETED_CALLS_XLSX": "completed_calls_xlsx",
    "DATA_DIR": "data_dir",
    "OUTPUT_DIR": "output_dir",
}

_PATH_DEFAULTS = {
    "flow_xlsx": "data/flow.xlsx",
    "sheet_name": "master_block_calendar",
    "rotation_rules_xlsx": "data/rotation_rules.xlsx",
    "no_call_days_xlsx": "data/no_call_days.xlsx",
    "holidays_xlsx": "data/holidays.xlsx",
    "clinic_days_xlsx": "data/clinic_days.xlsx",
    "completed_calls_xlsx": "",
    "data_dir": "data",
    "output_dir": "output",
}


def _safe_yaml() -> YAML:
    """YAML instance for read-only loads — returns plain Python dicts."""
    return YAML(typ="safe")


def _round_trip_yaml() -> YAML:
    """YAML instance that preserves comments and formatting on dump.

    `indent(mapping=2, sequence=4, offset=2)` matches the indentation style
    of the hand-written config.yaml — list items appear as `  - item` with
    the dash at column 2. Without this, ruamel defaults to flat style
    (`- item` at column 0), which round-trips comments correctly but
    visually rewrites every list in the file.
    """
    y = YAML()
    y.preserve_quotes = True
    y.indent(mapping=2, sequence=4, offset=2)
    return y


def load_config(path: str = CONFIG_PATH) -> dict:
    with open(path, "r", encoding="utf-8") as f:
        return _safe_yaml().load(f)


def validate_config(config: dict) -> None:
    missing = [k for k in REQUIRED_KEYS if config.get(k) is None]
    if missing:
        raise KeyError(
            f"config.yaml is missing required key(s): {missing}. "
            f"Add them before running the scheduler."
        )


def load_default_config(path: str = CONFIG_PATH) -> tuple[dict, dict]:
    """Read config.yaml and split into (behavior_config, paths).

    behavior_config: scheduler tuning params (weights, dates, thresholds, flags).
                     The GUI persists this via "Save as defaults" and treats it
                     as the editable-in-session surface.
    paths:           file and directory locations (lowercase keys). The CLI
                     derives these from config.yaml; the GUI builds them from
                     user-uploaded file paths and does NOT persist them.
    """
    raw = load_config(path)
    validate_config(raw)

    paths = dict(_PATH_DEFAULTS)
    for yaml_key, internal_key in _PATH_KEY_MAP.items():
        if yaml_key in raw and raw[yaml_key] is not None:
            paths[internal_key] = raw[yaml_key]

    config = {k: v for k, v in raw.items() if k not in _PATH_KEY_MAP}
    return config, paths


def save_config(values: dict, path: str = CONFIG_PATH) -> None:
    """Overwrite `path` with updated values, preserving comments and formatting.

    Only keys present in `values` are updated; all other keys (including path
    keys not exposed to the GUI) and every comment in the original file are
    preserved via ruamel.yaml round-tripping.

    Raises FileNotFoundError if `path` does not exist — we never create a
    config file from scratch, because that would discard the explanatory
    comments the original file carries.
    """
    p = Path(path)
    if not p.exists():
        raise FileNotFoundError(
            f"Cannot save config: {path} not found. save_config updates an "
            f"existing file in-place to preserve comments; create the file "
            f"manually first."
        )

    y = _round_trip_yaml()
    with open(p, "r", encoding="utf-8") as f:
        data = y.load(f)

    for key, value in values.items():
        existing = data.get(key)
        # Lists need in-place mutation, not reassignment, to preserve
        # ruamel.yaml's CommentedSeq wrapper — it carries the original
        # block style (one item per line, indented) and any inline or
        # trailing comments. Plain `data[key] = [...]` replaces the
        # CommentedSeq with a bare Python list, which ruamel then dumps
        # in flow style and drops adjacent section-header comments.
        if isinstance(existing, list) and isinstance(value, list):
            # Snapshot the CommentedSeq's per-item comment metadata, do
            # the in-place replacement, then restore it. Slice-assign
            # (`existing[:] = ...`) wipes existing.ca.items, which is
            # where ruamel stores trailing comments — and those trailing
            # comments often belong to the *next* top-level key's section
            # header (e.g. the "# Pick-candidate rank priority" block
            # that lives between MONTE_CARLO_SCORE_ORDER's last item and
            # PICK_CANDIDATE_RANK_ORDER). Without this restore, every
            # save erases those headers.
            saved_ca_items = None
            if hasattr(existing, "ca") and existing.ca.items:
                saved_ca_items = dict(existing.ca.items)
            existing[:] = value
            if saved_ca_items and hasattr(existing, "ca"):
                existing.ca.items.update(saved_ca_items)
        else:
            data[key] = value

    with open(p, "w", encoding="utf-8") as f:
        y.dump(data, f)


def legacy_gui_config_warning(path: str = LEGACY_GUI_CONFIG_PATH) -> str | None:
    """Return a warning message if a legacy gui_config.yaml is on disk, else None.

    Earlier GUI designs considered persisting overrides to gui_config.yaml. The
    final design (docs/gui_plan.md §4.5) writes back to config.yaml directly,
    so gui_config.yaml should never exist. If it does — left behind by a
    crashed future version or manually placed — tell the user it is ignored.
    """
    if Path(path).exists():
        return (
            f"{path} was found on disk. The CLI does not read this file. "
            f"Delete it or run the GUI to clean up."
        )
    return None


# Backward compatibility: existing modules (loader.py) import CONFIG at module
# level. Keep this so imports don't break. New code should call
# load_default_config() instead.
CONFIG = load_config()
validate_config(CONFIG)
