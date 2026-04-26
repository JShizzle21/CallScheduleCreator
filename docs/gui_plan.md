# Call Schedule Creator — GUI Planning Document

**Status:** Design spec. No code written. Intended to be executed after the current academic year's schedule is delivered.

**Target architecture:** Local web app (Streamlit) launched from a `.bat` file. Single user, Windows-only, session-only persistence for uploaded files. Config changes persist across sessions via a GUI override file. Distributed via OneDrive/Teams as source + bundled embeddable Python + `install.bat`.

**Python version:** 3.11+. Pinned in `install.bat` and `requirements.txt`. Listed in README. An embeddable Python build is bundled with the source (no separate install required for the user — see §1).

**Browser:** Windows Edge (organization default) and Chrome both supported. No other browsers tested.

**Guiding principle:** The primary user is a senior/junior resident who runs this 1–3 times per year and will not remember how it works between runs. Every field needs a label and tooltip. Every error message needs to be prose, not a stack trace.

**Dev launch command** (run from a terminal — opens the GUI in the default browser at http://localhost:8501):

```bash
cd C:\Users\johnn\PycharmProjects\CallScheduleCreator
.venv/Scripts/python.exe -m streamlit run src/app.py
```

Press `Ctrl+C` in the terminal to stop the server. End users will get a packaged `run.bat` (see §1) — this command is for development only.

---

## 1. Distribution & install model

### What lives on OneDrive/Teams
- Source code (the whole project folder except `.venv/` and `data/output/`).
- `install.bat` (one-time setup — downloads embeddable Python on first run).
- `run.bat` (double-click to launch the app).
- `README.md` (the handoff doc — see §6).

> **As-built note:** the original plan called for shipping `python-embed/` *inside* the OneDrive folder. Phase 6 implementation chose a different approach: install.bat downloads embeddable Python 3.14.3 from python.org into `%LOCALAPPDATA%\CallScheduler\python_embed\` on first run. Rationale: (a) ~300MB of pyarrow/pandas wheels lands in the same place — keeping all of it off OneDrive avoids sync churn AND avoids syncing the binaries to every other user's machine, (b) install.bat becomes idempotent (delete `%LOCALAPPDATA%\CallScheduler` to force clean reinstall), (c) the OneDrive folder stays small (source + data + .bat files only). Trade-off: first install requires internet access to python.org and pypi.org.

### What install.bat does (as built)
1. Creates `%LOCALAPPDATA%\CallScheduler\` (off OneDrive — avoids sync churn on the ~300 MB of dependency wheels).
2. Downloads `python-3.14.3-embed-amd64.zip` from python.org and extracts to `%LOCALAPPDATA%\CallScheduler\python_embed\`.
3. Edits `python314._pth` to uncomment `import site` (otherwise pip-installed packages are unimportable — embeddable distribution quirk).
4. Downloads `get-pip.py` from `bootstrap.pypa.io` and bootstraps pip.
5. Runs `pip install -r requirements.txt` against the embedded interpreter.
6. Prints "Setup complete. Double-click run.bat to start."

Each step is idempotent — re-running skips work that's already done. Downloads use `curl` (bundled with Win10 1803+) with PowerShell fallback.

### What run.bat does (as built)
1. Cleans up `%TEMP%\tmp*` directories older than 7 days (Streamlit doesn't fire a session-end hook so upload staging dirs accumulate).
2. Verifies `%LOCALAPPDATA%\CallScheduler\python_embed\python.exe` exists; if not, prompts user to run install.bat first.
3. Scans for the first free port in 8501-8520 via `netstat -an | findstr LISTENING`. Aborts with a friendly message if all 20 are taken.
4. Starts Streamlit headless on that port: `python -m streamlit run src\app.py --server.port=<P> --server.headless=true --browser.gatherUsageStats=false --server.fileWatcherType=none`.
5. Opens the user's default browser to `http://localhost:<P>` after a 4-second delay (gives the server time to bind).
6. Server runs in the foreground of the terminal window — closing the window or Ctrl+C stops it.

### Failure modes to handle gracefully
- Bundled `python-embed/` missing → install.bat tells user to re-copy project from OneDrive.
- Port 8501 in use → try 8502, 8503, etc. Streamlit accepts `--server.port` but does not auto-fallback; run.bat needs to scan for a free port (e.g. `netstat -an | findstr :8501`) and pass the first free one explicitly.
- Venv missing → run.bat prompts user to run install.bat first.
- Corrupt `gui_config.yaml` (§4.5) → log warning, delete it, fall back to defaults.

### Phase 6 implementation notes (gotchas confirmed during dev testing)

These are the non-obvious things that will bite while building install.bat / run.bat. Capture them here so the implementer doesn't re-discover each one from scratch.

**Embeddable Python:**
- Ships **without `pip`** — install.bat must download `get-pip.py` (`https://bootstrap.pypa.io/get-pip.py`) and run it against `python-embed\python.exe` to bootstrap pip before any `pip install` works.
- Ships with `python311._pth` that **blocks `site-packages`** by default. install.bat must edit `python311._pth` and uncomment the `import site` line, otherwise pip-installed packages won't be importable at runtime.

**Streamlit first-launch:**
- Telemetry prompt hangs the terminal waiting for an email address. Pre-create `.streamlit/config.toml` in the project root with:
  ```toml
  [browser]
  gatherUsageStats = false
  [server]
  headless = true
  ```
  Ship this file with the source so the prompt never appears.

**Windows Firewall:**
- First time Python opens port 8501, Windows Defender Firewall pops a permission dialog. Users will instinctively click "Block." README must include a screenshot + "click Allow on private networks" instruction.

**Temp directory cleanup:**
- The GUI calls `tempfile.mkdtemp()` per browser session for uploaded-file staging. Streamlit doesn't fire a reliable session-end hook (browser close ≠ session end), so these dirs accumulate in `%TEMP%`. Add an explicit cleanup step at the top of `run.bat`:
  ```bat
  forfiles /p "%TEMP%" /m "tmp*" /d -7 /c "cmd /c rmdir /s /q @path" 2>nul
  ```
  Or do it in Python at app startup (safer — only deletes dirs we know we created, by checking for a sentinel file we drop into each one).

### Deferred (post-Phase-6 polish)

**README handoff doc (§6):** intentionally deferred until install.bat / run.bat are working and dogfooded on a clean Windows machine. Doing it earlier means rewriting it once the install steps stabilize.

**GUI logic test coverage:** zero unit-tests today on `app.py` validators, `_seed_widget_state_from_config`, `_build_paths_from_uploads`, `_preflight_completed_assignments`, the queue/log handler. `streamlit.testing.v1.AppTest` (Streamlit ≥1.28) provides a programmatic API for driving the app: instantiate `AppTest.from_file("app.py")`, set widget values, click buttons, then assert on session_state and rendered elements. Worth a few hours after Phase 6 settles to lock in the upload-validator, reset-to-defaults, and run-lifecycle behaviors.

---

## 2. Page layout

Single-page app with four collapsible sections, expanded in this order on launch:

```
┌─────────────────────────────────────────────────────────┐
│  Call Schedule Creator                                   │
│  [Help] [Reset session]                                  │
├─────────────────────────────────────────────────────────┤
│  ▼ 1. Upload input files                                │
│     (file slots, last-uploaded indicators)              │
├─────────────────────────────────────────────────────────┤
│  ▼ 2. Settings   [Reset to defaults] [Save as defaults] │
│     "3 values differ from defaults"                     │
│     ▼ Common (always expanded)                          │
│     ▷ Advanced (collapsed)                              │
│     ▷ Expert — tuned values, modify with caution        │
│         (collapsed, yellow banner)                      │
├─────────────────────────────────────────────────────────┤
│  ▼ 3. Run schedule                                      │
│     [big green button]   Progress bar                   │
│     Live log output (scrolling text area)               │
├─────────────────────────────────────────────────────────┤
│  ▼ 4. Results (only visible after a run completes)      │
│     Status banner (success/warnings/errors)             │
│     [Download all three files] buttons row              │
│     ▼ Audit — structured view                           │
│         Errors / Warnings / Fairness / Unassigned /     │
│         Avoid assignments (collapsible sections)        │
│     ▼ Call totals — sortable table                      │
│     ▼ Call schedule — filterable table                  │
│         (month / slot type / resident filters)          │
└─────────────────────────────────────────────────────────┘

**Results replace on each run.** No in-app run history. When the user clicks Run
again, the previous results are cleared immediately and replaced with a
"Generating new schedule..." placeholder. New results populate once the run
completes. On-disk output files in `data/output/` are overwritten each run
(existing CLI behavior preserved).
```

---

## 3. Upload slots specification

Temporary directory created on launch (`tempfile.mkdtemp()`), deleted on session end.
Files persist within a session — user only re-uploads when a file actually changes.

| Slot label | File | Required | Validation on upload | Error message template |
|---|---|---|---|---|
| **Rotation flow sheet** | `flow.xlsx` | Yes | Must have sheet `master_block_calendar`; row 2 must contain parseable dates in columns B+; at least one resident row | "Flow sheet error: [specific issue]. Expected sheet named 'master_block_calendar' with block start dates in row 2." |
| **Rotation rules** | `rotation_rules.xlsx` | Yes | Columns: rotation, PGY1, PGY2, PGY3; values ∈ {ELIGIBLE, AVOID, NO_CALL} | "Rotation rules error: row N has invalid value '[X]'. Allowed values: ELIGIBLE, AVOID, NO_CALL." |
| **No-call days** | `no_call_days.xlsx` | Yes (can be empty) | Columns: name, date; names must match flow sheet; dates parseable | "No-call days error: name '[X]' in row N not found in flow sheet." |
| **Holidays** | `holidays.xlsx` | Yes (can be empty) | Columns: date, name, upper, intern; assigned names must match flow sheet PGY constraints | "Holidays error: '[X]' is listed as an intern holiday assignment but is a PGY[Y]." |
| **Clinic days** | `clinic_days.xlsx` | Optional | Columns: name, date; if present, validated like no-call days | Same as no-call |
| **Completed calls** | `completed_calls.xlsx` | Optional (required if partial-year mode on) | Columns: date, upper, intern | "Completed calls error: [specific issue]." |

**Upload UX:**
- Each slot shows either: "No file loaded" (gray), or filename + "Uploaded [timestamp]" + [X to remove].
- Green checkmark once validation passes; red X with message if it fails.
- "Validate all files" button at the bottom — runs all validators, shows a summary before the user tries to run the schedule.
- **Disable the Run button until all required slots are green.**

---

## 4. Config inventory (Settings page spec)

All values are editable. Organized into three accordions:

### 4.1 Common (always expanded)

These change every year or are high-impact.

| Key | Label | Widget | Validation | Tooltip |
|---|---|---|---|---|
| `ACADEMIC_DATE_START_STRING` | Academic year start | Date picker | Must be a valid date | "First day of the academic year. All scheduling begins from this date." |
| `ACADEMIC_DATE_END_STRING` | Academic year end | Date picker | Must be > start; typically ~365 days after start | "Last day of the academic year (inclusive)." |
| `PGY3_CUTOFF_DATE` | PGY3 graduation cutoff | Date picker (nullable) | Must be within academic year, or empty | "PGY3s are excluded from call on this date and after. Leave blank to keep PGY3s eligible all year." |
| `SIMULATION_RUNS` | Number of simulation runs | Slider: 1–5000, default 1000 | Integer ≥ 1 | "More runs = better schedule but slower. 1000 is usually sufficient." |
| `USE_COMPLETED_CALLS` | Partial year mode | Toggle (off/on) | Boolean | "ON: seed the schedule from completed_calls.xlsx and generate from the next day. OFF: generate a full year from scratch." |
| `INTERN_BLOCK1_WEEKDAY_CALLS` | Interns take weekday calls in Block 1 | Toggle | Boolean | "ON: interns cover both weekday and weekend calls in Block 1. OFF: weekday calls in Block 1 are covered by Night Float." |

### 4.2 Advanced (collapsed by default)

Constraints that rarely need adjustment but could reasonably change.

| Key | Label | Widget | Validation | Tooltip |
|---|---|---|---|---|
| `POST_CALL_DAYS` | Post-call rest days | Slider: 0–5, default 2 | Integer ≥ 0 | "Minimum days off after a call before another can be assigned. Hard constraint." |
| `MIN_SPACING_DAYS_STRONG` | Strong spacing threshold (days) | Slider: 1–21, default 7 | Integer ≥ 1 | "Calls within this many days incur a strong penalty (soft — not hard-blocked)." |
| `MIN_SPACING_DAYS_MILD` | Mild spacing threshold (days) | Slider: 1–30, default 14 | Integer > strong | "Calls within this many days incur a mild penalty." |
| `MAX_CALLS_IN_WINDOW` | Max calls per window | Slider: 0–10, default 3 | Integer ≥ 0; 0 = disabled | "Rolling cap: no more than this many calls per N days. Set to 0 to disable." |
| `ROLLING_WINDOW_DAYS` | Rolling window (days) | Slider: 7–30, default 14 | Integer ≥ 1 | "Size of the rolling window for the cap above." |
| `MAX_DIFF_SOFT` | Soft fairness threshold | Slider: 1–10, default 3 | Integer ≥ 1 | "Call-count gap that triggers the soft fairness flag in ranking." |
| `MAX_DIFF_HARD` | Hard fairness threshold | Slider: 1–10, default 4 | Integer > soft | "Call-count gap that triggers the hard fairness flag." |
| `NIGHT_FLOAT_ROTATION_NAME` | Night Float rotation code | Text | Non-empty string | "The exact name used for Night Float in the flow sheet (e.g. 'NF')." |

### 4.3 Expert — tuned values, modify with caution (collapsed, yellow banner)

**Banner text:** ⚠️ These values control how the scheduler ranks candidates. They have been tuned through Monte Carlo testing. Modifying them may produce worse schedules. Change only if you understand the weighted-score system — see README §Scoring.

| Key | Label | Widget | Validation | Tooltip |
|---|---|---|---|---|
| `FAIRNESS_GAP_WEIGHT` | Fairness gap weight | Slider: 0–10, step 0.25, default 3.0 | Float ≥ 0 | "How strongly to prefer residents behind in call count." |
| `SPACING_WEIGHT` | Spacing weight | Slider: 0–10, step 0.25, default 1.0 | Float ≥ 0 | "How strongly to prefer residents with longer spacing since last call." |
| `AVOID_WEIGHT` | Avoid-rotation weight | Slider: 0–5, step 0.25, default 0.25 | Float ≥ 0 | "Penalty for assigning call while on an AVOID rotation." |
| `YEAR_BIAS_WEIGHT` | Year-bias weight | Slider: 0–5, step 0.25, default 1.5 | Float ≥ 0 | "How strongly to front-load PGY3s and back-load PGY2s." |
| `PACE_WEIGHT` | Pace weight | Slider: 0–5, step 0.25, default 1.0 | Float ≥ 0 | "Corrective: penalizes residents ahead of their expected call pace." |
| `LOOKAHEAD_WEIGHT` | Lookahead weight | Slider: 0–5, step 0.25, default 3.0 | Float ≥ 0 | "Anticipatory: prefers residents with less remaining eligibility runway." |
| `PICK_CANDIDATE_RANK_ORDER` | Candidate rank order | Drag-to-reorder list | Must include all of: hard_diff_flag, soft_diff_flag, weighted_score | "Order in which candidate ranking criteria are applied (lexicographic)." |
| `MONTE_CARLO_SCORE_ORDER` | Monte Carlo score order | Drag-to-reorder list | Must include at least one scoring key | "Order in which schedule-level metrics are prioritized when selecting the best MC run." |

### 4.4 Fixed (not exposed in UI)

These are structural and should not be edited from the GUI. Shown in a read-only "About this run" panel if the user expands it.

| Key | Reason for hiding |
|---|---|
| `DATA_DIR`, `OUTPUT_DIR`, `*_XLSX` paths | Managed by the GUI via upload slots and download buttons. |
| `SHEET_NAME` | Structural; changing it breaks flow sheet parsing. |

### 4.5 Config persistence model

**Per-session overrides, in-memory only.** Settings changes live in Streamlit's `st.session_state` while the server is running. They do **not** persist to disk between sessions. Each new session starts from a fresh copy of `config.yaml`.

The only way to persist changes across sessions is the explicit "Save as defaults" button, which overwrites `config.yaml` with the current values.

| Storage | Role | Lifetime |
|---|---|---|
| `config.yaml` | Baseline defaults. Preserved with comments and original formatting. | Permanent on disk. Only mutated by "Save as defaults". |
| Streamlit `session_state` | User's in-session overrides (current effective values). | In-memory; gone when server stops or browser session ends. |

**The scheduler always reads from `session_state` (the GUI's effective config).** At session start, `session_state` is seeded from `config.yaml`. Every edit updates `session_state`. No merge file, no disk cache.

**Three user actions:**
- **Any edit in the GUI** → updates `session_state` immediately. The current value shown is always the effective value. No save button for individual fields.
- **"Reset to defaults"** button (Settings header) → reloads `session_state` from `config.yaml`, discarding all in-session changes. Confirmation dialog: *"Discard all your changes and return to the saved defaults?"*
- **"Save as defaults"** button (Settings header) → writes the current `session_state` values into `config.yaml` (via ruamel.yaml to preserve comments). Confirmation dialog: *"Overwrite config.yaml so your current values become the new defaults? The baseline will be permanently replaced."*

**UI indicators:**
- Each field shows a small "●" marker next to its label when it differs from `config.yaml`.
- Settings header shows "3 values differ from defaults" as a reminder.

**Comment preservation (important):** `config.yaml` contains extensive explanatory comments that are valuable for future maintainers. The default Python YAML library (`PyYAML`) strips them on write. The GUI must use **`ruamel.yaml`** (drop-in replacement that round-trips comments and formatting) so "Save as defaults" preserves the annotations. Add `ruamel.yaml` to `requirements.txt`.

**CLI interaction:**
- The CLI reads `config.yaml` directly (unchanged behavior).
- Because GUI overrides never persist to disk, there is no cross-tool drift risk by design.
- **Safety check:** if for any reason a `gui_config.yaml` file is ever found on disk at CLI startup (e.g. left behind by a crashed future version, or manually placed), the CLI prints a warning: *"gui_config.yaml was found on disk. The CLI does not read this file. Delete it or run the GUI to clean up."* — then proceeds with `config.yaml` unchanged.

**Browser refresh note:** Streamlit `session_state` persists across page reruns but may be lost on hard browser refresh (Ctrl+F5). If this becomes a frequent annoyance in practice, a fallback — persist to a temp file, reload on refresh — can be added. Start without it.

### 4.6 Input validation and reasonableness warnings

Two tiers of input checking on the Settings page:

**Hard validation (errors, block the Run button):**

These produce a red error message and disable the Run buttons until fixed.

| Rule | Example error |
|---|---|
| Dates must parse as real dates | "Academic year start is not a valid date." |
| `ACADEMIC_DATE_END` > `ACADEMIC_DATE_START` | "Academic year end must be after academic year start." |
| `PGY3_CUTOFF_DATE` must be within academic year, or empty | "PGY3 cutoff must be between the academic start and end dates." |
| `MAX_DIFF_HARD` ≥ `MAX_DIFF_SOFT` | "Hard threshold must be ≥ soft threshold." |
| `MIN_SPACING_DAYS_MILD` ≥ `MIN_SPACING_DAYS_STRONG` | "Mild spacing must be ≥ strong spacing." |
| All numeric fields non-negative | "Weight values cannot be negative." |
| `SIMULATION_RUNS` ≥ 1 | "Simulation runs must be at least 1." |
| `NIGHT_FLOAT_ROTATION_NAME` non-empty | "Night Float rotation name cannot be blank." |

**Soft warnings (yellow banner, allow Run but flag suspicious values):**

These are odd-but-not-invalid values that the user might have entered by mistake. They appear in a collapsible "Configuration warnings (2)" panel above the Run button.

| Trigger | Warning message |
|---|---|
| Academic year duration < 300 days or > 400 days | "Academic year is [X] days. Typical value is ~365." |
| `PGY3_CUTOFF_DATE` more than 45 days before year end | "PGY3s will be excluded for [X] days. That may strain PGY2 coverage." |
| `SIMULATION_RUNS` > 5000 | "Simulation will be slow — expect several minutes. Consider the Quick preview button if iterating." |
| `SIMULATION_RUNS` < 100 | "Low run count may produce suboptimal schedules. 1000 is the tuned default." |
| `POST_CALL_DAYS` = 0 | "No rest days enforced between calls — residents could be assigned on consecutive days." |
| Any weight set to 0 | "[WEIGHT_NAME] is 0 — this component will be ignored during candidate ranking." |
| Any weight > 10 | "[WEIGHT_NAME] is [X] — far above the default. This component will dominate ranking." |
| `MAX_CALLS_IN_WINDOW` = 0 | "Rolling-window cap disabled — burst patterns (multiple calls within a short span) will not be prevented." |

Warnings do not block Run. They are a sanity check for the non-technical user: "are you sure?" — not a gate.

---

## 5. Run page & results behavior

### 5.1 Run

- **Two run buttons** side by side:
  - **"Quick preview"** (secondary style) — runs with `SIMULATION_RUNS = 50`. Returns in ~10–30 seconds. For fast iteration while tuning config values.
  - **"Full run"** (primary, big green) — uses the configured `SIMULATION_RUNS` value (default 1000). For the final schedule.
- Both buttons disabled until all required input files validate green.
- On click:
  - Progress bar fills from 0 → N. Label shows "Run 347 / 1000 — best score so far: (0, 0, 1, 2, ...)".
  - Live log area below streams info-level messages (holiday warnings, cutoff warnings, etc.).
- **Cancel button** interrupts the simulation cleanly — last completed best is discarded, previous results (if any) are restored.
- **Behavior on subsequent runs:** previous results are cleared immediately and replaced with a "Generating new schedule..." placeholder. New results populate once the run completes. Outputs are not shown partially.

### 5.2 Results — overall structure

Shown immediately when a run completes. All three outputs render inline in the GUI so the user can review without leaving the app. Download buttons are always available (no "finalize" step needed — user runs as many times as they want, downloads whenever satisfied).

Top of section:
1. **Status banner** — green "Schedule generated successfully" / yellow "Completed with warnings" / red "Completed with errors". One sentence summary: "12 warnings, 0 errors, 5 unassigned holiday slots." and Monte Carlo score tuple.
2. **Download row** — three buttons side by side:
   `[⬇ call_schedule.xlsx]` `[⬇ call_totals.xlsx]` `[⬇ audit_report.txt]`
   Buttons stream the latest `data/output/*` file. User can download any or all, any time after the run completes.

Three expandable result sections below (all expanded by default on first run):

### 5.3 Audit — structured view

The plain-text audit file still gets written to disk. The GUI renders a prettier, structured version from the `audit_data` dict returned by `generate_schedule_once`. This is much easier for non-technical users to scan than the raw text.

Rendered as collapsible subsections, in this order:

| Section | Color | Default state | Content |
|---|---|---|---|
| **Errors** | Red | Expanded if non-empty | Bulleted list of error strings. Empty state: "No errors ✓" |
| **Warnings** | Yellow | Collapsed | Bulleted list. Count shown in header ("Warnings (12)") |
| **Fairness summary** | Neutral | Expanded | Table: metric / min / max / diff. Rows for upper_total, pgy2_total, pgy3_total, upper_weekday, upper_weekend, intern_weekend. |
| **Unassigned slots** | Yellow if non-empty | Expanded if non-empty | Table: date / slot / reasons. Empty state: "All required slots assigned ✓" |
| **Avoid-rotation assignments** | Neutral | Collapsed | Table: date / resident / rotation / slot. Count in header. |
| **Run metadata** | Neutral | Collapsed | Seed, tiebreaker count, swap improvements, MC score order, weights used. |

### 5.4 Call totals — sortable table

The `call_totals.xlsx` content rendered as an interactive table. Columns match the xlsx:

- `name`, `pgy` | `total_calls` | `weekday_calls`, `weekend_calls` | `Jul_Dec_calls`, `Jan_Jun_calls`

**Features:**
- Sortable on any column (click header to toggle).
- PGY rows color-coded (PGY1 / PGY2 / PGY3) matching the xlsx.
- Group dividers between PGY cohorts.
- Totals row at bottom showing column sums.
- Fits on screen without scrolling (~14–20 residents).

### 5.5 Call schedule — filterable table

The `call_schedule.xlsx` content rendered as a filterable table. Columns: date, day_of_week, slot, resident, pgy, rotation, note.

This table is ~500–700 rows (one per slot per day). Filters along the top let the user narrow it down:

- **Month** dropdown (All / July / August / …)
- **Slot type** dropdown (All / Upper weekday / Upper weekend / Intern weekday / Intern weekend)
- **Resident** dropdown (All / individual names)
- **Note filter** checkbox: "Show only rows with notes (HOLIDAY, COMPLETED, UNASSIGNED)"

Table features:
- Sticky header row.
- Rows highlighted by `note`: UNASSIGNED → red, HOLIDAY → blue, COMPLETED → gray.
- Weekend rows subtly tinted to distinguish from weekdays.
- Sortable on any column.

### 5.6 No in-session run history

Only the most recent run's results are displayed. Previous runs are not retained for in-app comparison. Rationale:

- Users who want to compare two configurations can download the outputs from run 1 before triggering run 2.
- Keeps the UI simple — no "which run am I looking at?" confusion.
- Matches user intent: outputs are meant to be reviewed and finalized, not archived in-browser.

**Trade-off to accept:** fine-grained "did raising PACE_WEIGHT to 2.0 help?" comparisons require downloading + diffing xlsx files manually, or relying on memory of the fairness summary between runs. Quick-preview mode (§5.1) mitigates this by making iteration cheap.

---

## 6. README handoff document (what must be in it)

The README is read by someone who hasn't seen the project in a year. It should answer, in this order:

1. **What is this?** (2 sentences)
2. **How do I run it?** (install.bat once, run.bat thereafter; screenshots)
3. **What files do I need to prepare?** (the 4 required + 2 optional; link to example files)
4. **How do I read the output?** (what call_schedule vs call_totals vs audit_report mean)
5. **What are the common errors and how do I fix them?** (precomputed list of the most likely issues with fixes)
6. **What do I do if a slot is unassigned?** (use holidays.xlsx manual override; §why)
7. **What is the PGY3 cutoff and when should I change it?**
8. **What are the weights and should I touch them?** (short answer: no)
9. **Who wrote this and where can I find the source?** (link to repo, original author)

Keep it short. 3–5 pages of plain English with screenshots. Detailed architecture doc (CLAUDE.md) stays separate for future devs.

---

## 7. Required code refactor before GUI work begins

The GUI build is blocked on these changes to the existing code. **CLI behavior is preserved at every phase** — `python scheduler_main.py` continues to work identically from start to finish. Each phase is independently mergeable, all tests pass after each phase, and the CLI is exercised in tests at every step. No "big bang" — every change is additive.

**Phase 1 — zero-risk, do first:**
1. Replace `print()` calls with `logging.getLogger(__name__).info()`. `scheduler_main.py` installs a default stream handler when run as `__main__` so CLI output is unchanged.
2. Introduce `ScheduleError`, `ConfigError`, `DataValidationError` exception types. CLI catches them at the top level and prints the same messages as today.

**Phase 2 — moderate, do after Phase 1 is stable:**
3. Remove module-level `CONFIG` evaluation from `scheduler_main.py`. Pass `config` as a parameter to `generate_schedule_once` and `run_simulation`. Keep a `load_default_config()` helper that reads `config.yaml` (for CLI).
4. Extract data loading into `load_data_bundle(paths: dict) -> DataBundle` returning an object with `lookup`, `rules`, `no_call`, `holidays`, `completed_calls`, `clinic_days`. GUI builds this from uploaded file paths; CLI builds it from CONFIG paths via the helper.

**Phase 3 — for polish, do last:**
5. Add `progress_callback: Optional[Callable[[int, int, dict], None]]` parameter to `run_simulation`. Defaults to None → CLI uses tqdm as today; GUI passes a callback that updates `st.progress`.
6. Add cancel mechanism (`threading.Event` checked between simulation runs; defaults to None → CLI ignores).

**Phase 4 — ruamel.yaml round-trip helper (for §4.5):**
7. Replace PyYAML load/dump calls with ruamel.yaml in `config.py`, preserving comments on write. Add `save_config(values: dict, path: str)` helper used by "Save as defaults". Add a startup check: if a legacy `gui_config.yaml` is present on disk, log a warning and ignore it. CLI and GUI both read `config.yaml` as the sole persisted source.

**After all four phases:** the GUI is a thin Streamlit layer on top of an unchanged core. Building `app.py` becomes straightforward and low-risk.

---

## 8. Deferred / explicit non-goals

- **Edit-in-browser of input files.** Users edit xlsx files in Excel and re-upload. Nice-to-have for a v2.
- **Multi-user / concurrent access.** Single user only.
- **Long-term run history / database.** Not even session-level — only the most recent run is displayed. On-disk outputs are overwritten each run.
- **Side-by-side run comparison in GUI.** Explicitly rejected. User downloads files if comparison is needed.
- **Hosted deployment.** Local-only.
- **Non-Windows support.** Windows only for v1.
- **Mobile/tablet view.** Desktop browser only.
- **Calendar-grid view of the schedule.** Filterable table is sufficient for v1. Calendar view is a v2 polish item.

---

## 9. Resolved decisions

Decisions finalized with user — frozen before code changes begin. Reference for future maintainers.

| # | Decision | Chosen |
|---|---|---|
| 1 | Framework | Streamlit |
| 2 | Distribution | Bundled embeddable Python + install.bat + run.bat |
| 3 | Config persistence | In-memory session_state only. No gui_config.yaml on disk. "Reset to defaults" reloads from config.yaml; "Save as defaults" overwrites config.yaml. CLI always reads config.yaml as single source of truth (§4.5). |
| 3a | Input validation | Two tiers: hard validation (blocks Run) and soft warnings (yellow banner, allow Run) — see §4.6 |
| 4 | YAML library | `ruamel.yaml` (preserves comments on round-trip) |
| 5 | Validation | Structural-only on upload; cross-file at run time |
| 6 | Results behavior | Replace-on-run; no history; old results clear immediately |
| 7 | Schedule display | Filterable single table (month / slot / resident) |
| 8 | Audit display | Structured collapsible sections; plain-text audit_report.txt still downloadable |
| 9 | Quick-preview toggle | Included — two buttons, "Quick preview" (fewer seeds) + "Full run" |
| 10 | CLI backward compatibility | Preserved at every refactor phase; all tests pass after each phase |
| 11 | Python version | 3.11+ (embeddable build bundled) |
| 12 | Platform | Windows only; Edge or Chrome |
| 13 | Error recovery | Top-level catch in GUI, prose message shown, previous results remain visible |
| 14 | Rank-order widgets (§4.3) | Downgraded from drag-to-reorder to **multiline text area with fixed-set validation**. Streamlit core has no drag-reorder widget; using a 3rd-party component was rejected for simplicity. The allowed items for each list are baked in (users may only reorder, never add/drop) because the scheduler requires every component to be present. Required items shown as a caption above each field. |
| 15 | SIMULATION_RUNS widget (§4.1) | Downgraded from slider to `st.number_input`. Sliders are imprecise for the 1–5000 range — hard to land on round numbers like 1000. |
| 16 | PGY3 cutoff picker UX (§4.1) | Implemented as **checkbox gate + date picker**. `st.date_input` cannot itself be "blank" once rendered, so the checkbox toggles the whole feature. On first enable from empty state the picker falls back to `academic_end - 14 days` (not today). Stored as ISO string, or `""` when disabled. |
| 17 | Reset/Save confirmation (§4.5) | Implemented with `st.dialog` modal (Streamlit ≥1.35). Save is disabled while any hard validation error is present — writing an invalid config would break the CLI. |
| 18 | Widget-key naming | All Settings widgets use `key="w_<CONFIG_KEY>"` so the sync routine can mirror widget state → cfg in one loop. Checkbox + date-picker components use `cb_<KEY>` / `dp_<KEY>` prefixes. |
| 19 | Progress callback signature (§7 Phase 3) | `Callable[[completed, total, info], None]` where `info = {"seed", "score", "best_seed", "best_score"}`. `logger.info` per-seed log is preserved alongside the callback so CLI output stays byte-identical. |
| 20 | Cancel semantics (§7 Phase 3) | `cancel_event: threading.Event` checked after each `as_completed` iteration. On set: `executor.shutdown(cancel_futures=True)`, loop break, `logger.info("Simulation cancelled.")`, return `None`. Caller restores prior results. |
| 21 | `gui_config.yaml` warning location (§7 Phase 4) | Fires in `scheduler_main._main()` only (not at `config.py` import time) so it only surfaces on actual CLI run, not on every test import or future GUI startup. |

## 10. Open questions (to measure or resolve during build)

1. **Quick preview seed count** (§5.1) — default of 50 is a guess. Needs empirical tuning during build: smallest seed count that's directionally accurate compared to a full 1000-run result.
2. **Call schedule table row count** — at ~500–700 rows, even a filterable table may feel heavy. If performance or UX becomes an issue during build, fall back to monthly tabs (one tab per month, ~30 rows per tab).
3. **Streamlit temp directory cleanup** — Streamlit has no clean shutdown hook. Strategy: on each launch, scan the temp dir prefix for orphaned `CallScheduler_*` folders older than 24 hours and delete them. This catches the "user closed the browser tab without stopping the server" case without risking active-session cleanup.
4. **"Save as defaults" confirmation strength** — Resolved: single confirmation dialog is sufficient for v1 (single trusted user). Flagging here in case this ever ships to a shared-credential environment.
