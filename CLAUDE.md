# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Commands

All commands use the project venv at `.venv/`:

```bash
# Run the full scheduler (Monte Carlo + Excel export)
.venv/Scripts/python.exe scheduler_main.py

# Run all tests
.venv/Scripts/python.exe -m pytest tests/

# Run a single test
.venv/Scripts/python.exe -m pytest tests/test_pick_best_candidate.py::test_hard_diff_flag_dominates_lower_weighted_score

# Run tests with verbose output
.venv/Scripts/python.exe -m pytest tests/ -v
```

Outputs land in `data/output/` (gitignored): `call_schedule.xlsx`, `call_totals.xlsx`, `audit_report.txt`.

## Architecture

### Data flow

```
config.yaml
    └── config.py (CONFIG dict, loaded at module import)

data/flow.xlsx          ──► excel_reader.ExcelRotationLookup
data/rotation_rules.xlsx ──► loader.load_rotation_rules()
data/no_call_days.xlsx  ──► loader.load_no_call_days()
data/holidays.xlsx      ──► loader.load_holidays()

scheduler_main.generate_schedule_once(seed)
    ├── validate inputs (validation.py)
    ├── day loop: eligible_for_slot → pick_best_candidate → apply_assignment
    └── audit_schedule (validation.py) → returns result dict

scheduler_main.run_simulation(num_runs)
    ├── calls generate_schedule_once(seed=0..N)
    ├── scores each via monte_carlo_score() (lexicographic tuple)
    └── returns best result

export_result(result)
    ├── exports.write_call_schedule_xlsx
    ├── exports.write_call_totals_xlsx
    └── exports.write_audit
```

### Flow sheet format (`data/flow.xlsx`, sheet `master_block_calendar`)

- Row 1: ignored header
- Row 2: block start dates (columns B onward), used to define rotation blocks
- Row 3+: resident names in column A, rotation names in the block columns
- **PGY separator rows**: rows with ≥3 date-like cell values are treated as separators between PGY cohorts — PGY increments each time one is encountered (PGY1 first, then PGY2, then PGY3)
- **Split blocks**: a cell value like `WARDS/ED` means the resident is on WARDS for the first ~2 weeks and ED for the second ~2 weeks of that block

### Candidate ranking (`pick_best_candidate`)

Two-level system applied lexicographically per `PICK_CANDIDATE_RANK_ORDER`:

1. **Gates** (`hard_diff_flag`, `soft_diff_flag`): hard/soft fairness thresholds compared against pool min
2. **Weighted score** (`_compute_weighted_score`): all four components normalized to [0,1] before weighting so `*_WEIGHT` constants behave as true relative importance

All four weighted-score components are normalized:
- `fairness_norm = min(gap, MAX_DIFF_SOFT) / MAX_DIFF_SOFT`
- `spacing_norm = spacing_tier / 2` (tier ∈ {0,1,2})
- `avoid_value` and `year_bias` already in [0,1]

### Monte Carlo scoring

`monte_carlo_score()` returns a lexicographic tuple ordered by `MONTE_CARLO_SCORE_ORDER` (configurable). The run with the minimum tuple wins. All randomness comes from `random.Random(seed)` — `generate_schedule_once` is a pure function of its seed.

## Key config knobs

| Key | Effect |
|---|---|
| `SIMULATION_RUNS` | Number of Monte Carlo seeds to try |
| `POST_CALL_DAYS` | Hard floor: days after call where no assignment is allowed (enforced at eligibility) |
| `MIN_SPACING_DAYS_STRONG/MILD` | Soft spacing thresholds — only affect `spacing_tier` in weighted score, not eligibility |
| `MAX_DIFF_SOFT/HARD` | Fairness gate thresholds for `soft_diff_flag` / `hard_diff_flag` |
| `FAIRNESS_GAP_WEIGHT` etc. | Relative importance of each weighted-score component (all normalized to [0,1]) |
| `PICK_CANDIDATE_RANK_ORDER` | Order of rank components; gates must precede `weighted_score` or the hard-diff gate is bypassed |
| `MONTE_CARLO_SCORE_ORDER` | Priority of schedule-level metrics when selecting the best MC run |

## Known limitations / future work

- **PGY assignment is fragile**: PGY is inferred by counting skip-rows in `loader.py`; reordering the flow sheet silently produces wrong PGYs
- **Greedy day-by-day has no recovery**: a local swap post-pass or CP-SAT solver would improve schedule quality
- **`CONFIG.get` silently returns `None`** for missing required keys, causing cryptic downstream errors — a required-key check at startup is missing
- **Parallel Monte Carlo**: `generate_schedule_once` is now a pure function of seed, so `ProcessPoolExecutor` over seeds is straightforward
- Holidays are marked unassigned and require manual assignment
