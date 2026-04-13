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
    ├── precompute future_eligible per resident (static rotation + no_call_days)
    ├── day loop: eligible_for_slot → pick_best_candidate → apply_assignment
    ├── local_swap_pass (post-generation fairness repair, up to 10 iterations)
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
2. **Weighted score** (`_compute_weighted_score`): all five components normalized to [0,1] before weighting so `*_WEIGHT` constants behave as true relative importance

All five weighted-score components are normalized:
- `fairness_norm = min(gap, MAX_DIFF_SOFT) / MAX_DIFF_SOFT`
- `spacing_norm = spacing_tier / 2` (tier ∈ {0,1,2})
- `avoid_value`, `year_bias`, and `future_avail_norm` already in [0,1]

`future_avail_norm` is the fraction of remaining call-eligible days this resident has vs. the pool maximum on the current day. Precomputed once per seed using static constraints (rotation + no_call_days, not post-call). Residents about to enter long NO_CALL rotations score low (preferred now).

### Post-generation local swap pass (`local_swap_pass`)

After the greedy day loop, runs up to 10 passes over all assigned slots looking for improving swaps. A swap replaces resident A with resident B on a given day when:
- B has ≥ 2 fewer calls of the relevant type (gap-of-2 prevents oscillation across passes)
- B is eligible: correct rotation, no no_call_day, not post-call, no forward spillover conflict at D+1/D+2
- `_undo_assignment` mirrors `apply_assignment` to reverse all six counters

`swap_improvements` count is recorded in `audit_data`. The audit runs after all swap passes so reported metrics reflect the improved schedule.

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
| `FUTURE_AVAIL_WEIGHT` | Look-ahead bias toward residents with fewer remaining eligible call days |
| `PICK_CANDIDATE_RANK_ORDER` | Order of rank components; gates must precede `weighted_score` or the hard-diff gate is bypassed |
| `MONTE_CARLO_SCORE_ORDER` | Priority of schedule-level metrics when selecting the best MC run |


## Applied Learning

When something fails repeatedly, when Johnny has to re-explain, or when a workaround is found for a a platform/tool limitation, add a one-line bullet here. Keep each bullet under 15 words. No explanations. Only add things that will save time in future sessions.