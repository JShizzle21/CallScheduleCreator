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


CONFIG = load_config()
validate_config(CONFIG)
