# Claude Code session handoff

Paste this whole file into a fresh Claude Code session so it can continue
work without re-deriving context.

---

## What this project is

**Call Schedule Creator** — a Windows desktop app (Streamlit GUI + Python
scheduler) that builds a year-long medical residency call schedule from
Excel inputs. Monte Carlo simulation picks the best schedule across
1000 seeds; greedy day-loop with weighted candidate ranking; post-
generation local-swap pass for fairness repair. Outputs: a daily call
schedule (`call_schedule.xlsx`), per-resident totals
(`call_totals.xlsx`), and an audit report (`audit_report.txt`).

Repo: **https://github.com/JShizzle21/CallScheduleCreator**

User: Johnny McMurray (johnnymcmurray123@gmail.com). Windows machine.
Project lives at `C:\Users\johnn\PycharmProjects\CallScheduleCreator`
on the original device. End-user audience is a chief resident /
program scheduler, not a developer.

## Tech stack

- Python 3.14 (embedded distribution on end-user machines via
  `install.bat`; dev uses a `.venv/` at the project root with
  `.venv/Scripts/python.exe`).
- Streamlit GUI (`src/app.py`).
- openpyxl for Excel I/O, ruamel.yaml for round-tripping `config.yaml`
  with comments, pandas for the in-GUI totals table.
- pytest for the 67-test suite.
- ProcessPoolExecutor for parallel Monte Carlo runs.

## Project layout (post-May-refactor)

```
README.md   install.bat   run.bat   uninstall.bat       ← end-user
input_files/      output_files/                         ← end-user data
  flow.xlsx         call_schedule.xlsx                  (gitignored)
  rotation_rules.xlsx  call_totals.xlsx
  no_call_days.xlsx    audit_report.txt
  holidays.xlsx
  clinic_days.xlsx
  completed_calls.xlsx
src/                                                    ← all Python
  app.py                — Streamlit GUI
  scheduler_main.py     — CLI entry, MC loop, day loop, swap pass
  config.py + config.yaml
  data_bundle.py        — single-call input loading
  excel_reader.py       — flow-sheet → rotation lookup
  loader.py             — all per-file loaders
  exports.py            — write_call_schedule_xlsx etc.
  validation.py         — audit_schedule
  errors.py             — custom exceptions
  tests/                — pytest suite (conftest walks 2 parents up)
docs/
  requirements.txt
  gui_plan.md           — authoritative GUI design doc
  claude_handoff.md     — this file
.gitignore  .git/  .venv/  CLAUDE.md (gitignored, personal scratchpad)
```

## How to run things

```bash
# Tests
.venv/Scripts/python.exe -m pytest src/tests/ -q

# CLI scheduler end-to-end (writes to output_files/)
.venv/Scripts/python.exe src/scheduler_main.py

# GUI dev launch
.venv/Scripts/python.exe -m streamlit run src/app.py
```

CLI assumes cwd = project root. GUI ditto. All 67 tests pass on master
as of the last commit.

## Important conventions

- **Resident names**: the flow sheet (`input_files/flow.xlsx`) is the
  source of truth. Last-name-only, case-sensitive in the canonical
  form. Other input files match last names case-insensitively +
  whitespace-trimmed; unknown last names hard-error with a sheet/row
  pointer.
- **Date format**: 14 formats accepted (ISO, US slashes, named months,
  abbreviated months — see `_DATE_FORMATS` in `loader.py`).
- **Working-doc tolerance**: supervisors edit the workbooks live, so
  loaders prefer warnings + skips over hard stops for non-fatal issues
  (extra/blank rows, asterisk-wrapped tentative dates, trailing `?` on
  uncertain entries).
- **CLAUDE.md** at the root is a gitignored scratchpad with extra
  architectural detail (scoring/ranking design, swap pass mechanics,
  candidate ranking knobs).
- **Drift warning** in `src/validation.py`: any new module-level
  constant that reads from CONFIG must ALSO be added to
  `validation._apply_config()` or it goes stale on GUI overrides.

## Recent major changes (most recent first)

1. **Root-cleanup refactor** (commit `dad8658`) — moved
   `scheduler_main.py` into `src/`, `tests/` into `src/tests/`,
   `requirements.txt` into `docs/`, and split `data/` into
   `input_files/` + `output_files/` as siblings. Pure file moves +
   path updates, no functional change.
2. **Holiday workbook redesign** (commit `f29374d`) — replaced the old
   simple holidays.xlsx with a working-doc grid (holiday columns ×
   resident rows, cell = rotation). ER-prefixed rotations (case-
   insensitive `ER` followed by space or end) become actual call
   assignments; non-ER cells just block day-of + day-before. Added the
   intern-day-before-NF rule (`pre_nf` no-call block). 22 new tests in
   `test_holidays_loader.py`.
3. **Multi-sheet no_call_days workbook** (commit `bdd4c3d`) — 13-sheet
   workbook with Interns/Uppers sections, header scanning, last-name
   canonicalisation, range-crosses-block warning. Audit now lists every
   loaded range. 5 new tests in `test_no_call_loader.py`.
4. **Sunday calls + bold dividers** (commit `3d58d25`) —
   `call_totals.xlsx` got a `sunday_calls` column and thick vertical
   dividers grouping name/pgy ‖ totals/weekday/weekend ‖ fri/sat/sun ‖
   halves. Schedule writer's block divider thickened.
5. **Friday/Saturday call counters** (commit `13dc104`) — Fri + Sat
   tracked separately on top of weekend totals because they're the
   least-favoured days.
6. **Multi-sheet clinic_days workbook** (also `13dc104`) — 13-sheet
   layout matching the supervisor's working doc. Date column found by
   case-insensitive header scan in the first 10 rows.
7. **Name corrections** — `Prahbu → Prabhu` and `Avila → Loera`
   propagated across all input files.

## Open items / known TODOs

- **Clean-room testing** on a different Windows machine still pending.
  The user planned to test SmartScreen / Defender / firewall popups on
  a machine that has never seen this codebase. README's Quick Start
  was just rewritten (with step-by-step GitHub download instructions)
  to support a less-technical co-resident going through this.
- **`gui_config.yaml` migration cleanup** — `src/config.py` still has
  `legacy_gui_config_warning()` for an old design that was never
  shipped. Safe to delete if nobody has a stray `gui_config.yaml`
  lying around.
- **Smoke-test xfails** in `test_end_to_end_smoke.py` are
  `strict=True`. They'll self-clear (XPASSED → forces marker removal)
  once placeholder clinic data is replaced. May or may not still apply
  — verify.
- **Year-typo bulk fix** in `input_files/no_call_days.xlsx` (3 entries
  bumped from 2026→2027 locally). Supervisor's master should be
  fixed upstream.
- **GUI logic test coverage** with `streamlit.testing.v1.AppTest` —
  deferred during Phase 6. Worth a few hours to lock in the upload
  validators and the run-lifecycle queue.

## Useful pointers

- **Architecture details**: `CLAUDE.md` at the project root
  (gitignored personal scratchpad with scoring formulas, swap-pass
  invariants, candidate ranking knobs).
- **GUI design doc**: `docs/gui_plan.md` — authoritative spec with
  as-built notes from Phases 5b/5c/6.
- **User wants to commit only when explicitly asked.** They're
  technical enough to follow git ops but prefer explicit "push this"
  before any commit/push happens.
- **Communication style**: Johnny is a chief medical resident — give
  concise summaries, surface trade-offs as numbered options when
  there's a real choice to make, but don't ask permission for obvious
  defensive fixes.
