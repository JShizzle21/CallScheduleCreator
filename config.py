import yaml

CONFIG_PATH = "config.yaml"

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


def load_config(path: str = CONFIG_PATH) -> dict:
    with open(path, "r", encoding="utf-8") as f:
        return yaml.safe_load(f)


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


# Backward compatibility: existing modules (loader.py) import CONFIG at module
# level. Keep this so imports don't break. New code should call
# load_default_config() instead.
CONFIG = load_config()
validate_config(CONFIG)
