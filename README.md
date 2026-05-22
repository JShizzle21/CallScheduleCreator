# Call Schedule Creator

A Windows desktop app that builds a year-long medical residency call schedule from a few Excel input files. Optimizes for fairness, post-call rest, rotation eligibility, and spacing — and produces an Excel schedule, a per-resident call totals report, and an audit report.

---

## Quick start

You only need to do steps 1–3 the **first time**. After that, just double-click `run.bat` whenever you want to use the app.

### Step 1 — Download the project from GitHub

1. Open this link in any web browser: **https://github.com/JShizzle21/CallScheduleCreator**
2. On that page, look for the green **`< > Code`** button (top right of the file list). Click it.
3. In the dropdown that appears, click **`Download ZIP`** at the bottom.
4. Your browser downloads a file called `CallScheduleCreator-master.zip` (usually into your **Downloads** folder).

### Step 2 — Extract the ZIP somewhere convenient

1. Open your **Downloads** folder (File Explorer → Downloads on the left sidebar).
2. **Right-click** `CallScheduleCreator-master.zip` and pick **Extract All…**
3. When the dialog appears, click **Browse…** and choose a place you'll remember — your **Desktop** or **Documents** folder both work well. Click **Extract**.
4. You'll end up with a folder called `CallScheduleCreator-master`. Open it. Inside you should see `install.bat`, `run.bat`, `README.md`, and a few subfolders. **This is your project folder** — don't move individual files out of it.

> **Tip:** if Windows blocks the ZIP, right-click the ZIP → **Properties** → tick the **Unblock** checkbox at the bottom → click **OK**, then extract.

### Step 3 — Install (one time only, ~5 minutes)

1. Inside your project folder, **double-click `install.bat`**.
2. A black console window opens. It downloads Python and the libraries the app needs (~400 MB total). Don't close the window — just leave it alone.
3. When you see **`Setup complete.`** at the bottom, you're done. Close the window.

### Step 4 — Use the app

1. **Double-click `run.bat`** any time you want to use the app.
2. A console window opens and your default browser pops up after a few seconds, showing the app.
3. When you're finished, click the red **Exit** button at the top of the page (it shuts down the app cleanly). You can also just close the console window.

### Uninstall (optional)

Double-click **`uninstall.bat`** — it removes the ~400 MB of Python files from your computer. The project folder itself you delete by hand from File Explorer if you don't want it any more.

---

### First-launch warnings (all normal — these aren't errors)

Windows is cautious about programs it hasn't seen before. If you see any of these, here's how to dismiss them:

| What you see | What to do |
|---|---|
| **"Windows protected your PC"** (blue SmartScreen popup) | Click **More info**, then **Run anyway**. |
| **"Allow Python to communicate on these networks"** (firewall popup) | Tick **Private networks** and click **Allow access**. |
| Your antivirus flags `install.bat` or `run.bat` | Add the project folder to your antivirus's "exclusions" list. The scripts are plain text — you can open them in Notepad to verify. |

---

### Something went wrong?

- **`install.bat` failed with a download error** — your network is blocking `python.org` or `pypi.org` (common on hospital networks). Try a personal network or hotspot, or ask IT to whitelist those domains.
- **`run.bat` says "Python is not installed"** — you skipped Step 3. Run `install.bat` first, then try again.
- **Browser doesn't open automatically** — look at the console window for a URL like `http://localhost:8501`. Copy that into your browser manually.
- **Anything else weird** — double-click `uninstall.bat`, then `install.bat` to start clean.

---

## What you need to prepare

All input files live in the `input_files/` folder. You can edit them in Excel; **close the file before running the app** (Excel locks open files). The app also accepts uploads in the GUI — those override the files in `input_files/` for that run only.

> **Date format flexibility.** Every date column accepts any of these formats interchangeably, and Excel-native date cells are always accepted:
> `2026-07-01`, `7/1/2026`, `7/1/26`, `July 1 2026`, `July 1, 2026`, `1-Jul-2026`, `Jul 1, 2026`, `1 July 2026`. Use whichever is easiest in your workflow.

> **Resident names must match exactly** across all input files (case-sensitive). `Prabhu` is not the same as `prabhu`. The flow sheet is the source of truth; if a name appears in another file but not in `flow.xlsx`, the app will stop with an error pointing at the offending row.

> **Tip:** the files currently in `input_files/` are real working examples. Copy one as a template rather than starting from scratch.

---

### Required input files

#### 1. `flow.xlsx` — rotation block calendar

Defines which rotation each resident is on for each ~2-week block of the academic year.

- **Sheet name:** `master_block_calendar`
- **Row 1:** ignored (free-form header).
- **Row 2:** **block start dates** in columns B onward. One date per block — typically 13 blocks for a 12-month academic year. The block ends the day before the next block's start date (or the academic-year end date for the last block).
- **Column A (row 3+):** resident names, one per row, grouped by PGY cohort (PGY1 first, then PGY2, then PGY3).
- **Cells (row 3+ × column B+):** rotation codes that match `rotation_rules.xlsx` — e.g. `WARDS`, `ED`, `NF`, `ICU`, `ELECTIVE`.
- **PGY separator rows:** a row in which column A is blank/non-name and at least three cells in columns B+ are date-like values. The app uses these to detect cohort boundaries. The first separator ends PGY1 / starts PGY2; the second ends PGY2 / starts PGY3.
- **Split blocks:** a cell value like `WARDS/ED` means the resident is on `WARDS` for the first ~2 weeks of the block and `ED` for the second ~2 weeks.

Example layout:

| A | B | C | D | ... |
|---|---|---|---|---|
| | 2026-07-01 | 2026-07-15 | 2026-07-29 | ... |
| McMurray | WARDS | ED | NF | ... |
| Lovell | ICU | WARDS | ELECTIVE/WARDS | ... |
| *(PGY2 separator row with dates)* | | | | |
| Avila | WARDS | NF | ED | ... |

#### 2. `rotation_rules.xlsx` — call eligibility per rotation × PGY

Tells the scheduler which rotations may take call at each PGY level.

- **Columns (row 1 = header):** `rotation_name`, `pgy`, `preference`
- **Values:**
  - `ELIGIBLE` — can take call.
  - `AVOID` — eligible, but try not to assign (only used when no `ELIGIBLE` candidate exists).
  - `NO_CALL` — never takes call on this rotation.
- One row per `(rotation, PGY)` pair. Missing rows default to `NO_CALL`.

Example:

| rotation_name | pgy | preference |
|---|---|---|
| WARDS | 1 | ELIGIBLE |
| WARDS | 2 | ELIGIBLE |
| WARDS | 3 | ELIGIBLE |
| NF | 1 | NO_CALL |
| ELECTIVE | 2 | AVOID |

#### 3. `no_call_days.xlsx` — vacations, conferences, days off

Per-resident date ranges where the resident is unavailable for call.

- **Columns (row 1 = header):** `name`, `start_date`, `end_date`, `type` *(optional)*
- One row per range. `type` is a free-form label that appears in the audit report (e.g. `vacation`, `conference`, `wedding`). May be omitted.
- For a single day, set `start_date` and `end_date` to the same date.

Example:

| name | start_date | end_date | type |
|---|---|---|---|
| McMurray | 2026-12-22 | 2026-12-29 | vacation |
| Lovell | 2027-03-15 | 2027-03-15 | conference |

#### 4. `holidays.xlsx` — holiday dates + optional manual assignments

Holidays are visually highlighted in the output. The `upper` and `intern` columns let you hand-assign specific residents for holidays (otherwise the scheduler picks).

- **Columns (row 1 = header):** `date`, `name`, `upper`, `intern`
- `name` is a free-form label (`Christmas`, `Thanksgiving`); defaults to `Holiday` if blank.
- `upper` / `intern`: blank cells are left **unassigned** and flagged in the audit so you can hand-fill before shipping. Pre-filled cells are locked in.
- Upper assignments must be PGY2 or PGY3; intern assignments must be PGY1. The app errors at load time if a PGY mismatch is detected.

Example:

| date | name | upper | intern |
|---|---|---|---|
| 2026-12-25 | Christmas | Lovell | McMurray |
| 2027-01-01 | New Year | | |

---

### Optional input files

#### 5. `clinic_days.xlsx` — per-block continuity clinic schedule

The day **before** each clinic is automatically blocked as a no-call day for the resident in clinic (so they're not post-call the morning of clinic). This is a "working document" format — designed to evolve through the year.

- **Workbook structure:** exactly 13 sheets named `Block 1`, `Block 2`, …, `Block 13`. Sheet name matching is case-insensitive.
- **Per-sheet structure:**
  - The "Date" column is found by scanning the **first 10 rows** for a cell whose value (case-insensitive, trimmed) is exactly `Date`. The match is confirmed by checking that the cell directly below it contains a parseable date.
  - **Multiple header rows are fine** — anything above the confirmed "Date" header is ignored (so columns labelled "Amb - Davis", "Interns", day-of-week, etc. are no problem).
  - **Cells to the left of the Date column are ignored** (commonly used for day-of-week labels like `MON`/`TUES`).
  - **Cells to the right of the Date column** list the resident names who have clinic that day. Empty cells are skipped. The app scans up to N cells right, where N = total resident count.
  - **Body cells that match header-like labels** (`Date`, `Intern`, `Interns`, `Amb - <anything>`) are silently ignored — handy for stray cells in a working doc.
  - **Trailing `?` on a name** (e.g. `Payne?`) means "supervisor is unsure." The app strips the `?` for matching, applies the call block normally, and emits a one-time warning so you can verify later. The `?` is **not** removed from the source file.
  - **Empty block sheets are OK** — a `Block N` sheet with no clinics yet is accepted silently.

**Hard-error validation** (these all stop the run with a clear sheet/row pointer):
- A `Date` header is found, but the cell below it contains a non-date value (likely a misplaced column).
- A clinic date falls outside the academic year.
- A clinic date is on the wrong `Block N` sheet (i.e. the date is outside that block's calendar range).
- A resident name does not match any name in `flow.xlsx`.

Example `Block 1` layout:

| A | B | C | D | E | ... |
|---|---|---|---|---|---|
| *(free-form header row)* | | | Amb - Davis | | |
| | **Date** | | | | Interns |
| MON | 2026-07-06 | Clark | McMurray | Snell | |
| TUES | 2026-07-07 | Behaj | Green | | |

The columns labeled "Amb - Davis" and "Interns" in the header row above the Date header are noise — they're ignored. Only resident names in body rows count.

#### 6. `completed_calls.xlsx` — mid-year handoff seed

Used only when **"Partial year"** is enabled in the GUI. Tells the scheduler which calls are already done so it picks up after the last entry.

- **Columns (row 1 = header):** `date`, `upper`, `intern`
- One row per day. Blank `upper` or `intern` cells are tolerated (the scheduler infers from context).
- The scheduler restarts on the day after the latest date in this file.

---

## Reading the output

After a successful run, three files land in `output_files/` (a sibling folder to `input_files/` at the project root) and are also offered as download buttons in the GUI.

| File | Audience | What's in it |
|---|---|---|
| **`call_schedule.xlsx`** | The residents — this is the published schedule | Day-by-day calendar with assigned upper-level and intern. Holidays highlighted. |
| **`call_totals.xlsx`** | Chief / scheduler — fairness sanity check | Per-resident counts: `total_calls`, `weekday_calls`, `weekend_calls`, `friday_calls`, `saturday_calls`, `Jul_Dec_calls`, `Jan_Jun_calls`. Friday and Saturday are tracked separately because they're the least-favoured days. Color-coded by PGY. |
| **`audit_report.txt`** | Anyone debugging an issue | Plain-text summary: errors, warnings, fairness gaps, list of unassigned slots, and reasoning. **Read this first if anything looks off.** |

---

## Day-to-day workflow

1. **Update `input_files/flow.xlsx`** with the new academic year's rotations.
2. **Update `input_files/no_call_days.xlsx`** with the year's vacations and conferences as you collect them.
3. **Update `input_files/holidays.xlsx`** with the year's holidays. Leave the upper/intern columns blank if you want the scheduler to assign them, or fill them in to lock specific people.
4. **Run the app** (`run.bat`).
5. **Click "Run schedule."** Wait for it to finish (~30 seconds for 1000 simulations).
6. **Open `audit_report.txt`** and skim the Errors and Unassigned sections. Fix anything flagged.
7. **Open `call_schedule.xlsx`** to fill in any unassigned holiday slots by hand.
8. **Publish** the finalized schedule.

---

## Common questions

### A slot is unassigned. What do I do?

Open `audit_report.txt` and look in the **Unassigned** section. Each unassigned row says why:

- `holiday_manual_assignment` — that holiday cell was left blank in `holidays.xlsx` on purpose. Fill it in by hand in `call_schedule.xlsx` before publishing. Most common case.
- Anything else — the scheduler ran out of eligible candidates that day. Check the audit's Errors section for the specific constraint (no-call day, post-call, rotation conflict). Usually means too many people are off / on incompatible rotations on the same day. Either move someone's no-call day or accept a manual assignment that bends a soft rule.

### What is the "PGY3 graduation cutoff" and when should I change it?

PGY3s often graduate before the academic year ends. The cutoff date is the first day they're no longer scheduled — typically mid-June. Set it in the GUI's **Common settings** section. Toggle it off if PGY3s are scheduled through year-end.

### Should I touch the Expert weights?

**No, almost certainly not.** They're tuned. Leaving them alone produces good schedules. The only weight a chief might reasonably touch is `MAX_DIFF_HARD` (raise it if the scheduler can't find a feasible schedule and the audit reports residents being skipped). Everything else is fine on defaults.

If you do change weights and find a setting you like, click **"Save as defaults"** in the Settings section to overwrite `config.yaml`. Click **"Reset to defaults"** to undo without saving.

### How do I generate a partial-year schedule (mid-year handoff)?

1. Fill in `input_files/completed_calls.xlsx` with all calls already assigned (columns `date`, `upper`, `intern` — see the [Input files](#optional-input-files) section above for full format details).
2. In the GUI, toggle **"Use completed calls"** on.
3. Run. The scheduler picks up from the day after the last completed call.

### The app won't start / Browser doesn't open / "Address already in use"

- **Browser doesn't open automatically:** open it manually and go to the URL printed in the console window (e.g. `http://localhost:8501`).
- **Port conflict:** `run.bat` automatically tries 8501 through 8520. If all are taken, close other apps and retry.
- **`run.bat` says Python is not installed:** you skipped `install.bat`. Double-click it first.
- **`install.bat` fails with download errors:** corporate firewall or antivirus blocking `python.org` / `pypi.org`. Get IT to whitelist them, or run the installer on an unrestricted network.
- **Anything else:** double-click `uninstall.bat`, then `install.bat`.

### Where do I get help with the code?

Source lives at: *(project repo URL)*. Architecture details are in `CLAUDE.md` (for developers). The GUI design history is in `docs/gui_plan.md`.

---

## What's installed where

| Location | Contents | Safe to delete? |
|---|---|---|
| Project folder (this directory) | Code (`src/`), your inputs (`input_files/`), generated outputs (`output_files/`), docs (`docs/`) | Back up `input_files/` first. Everything else is replaceable from the repo. |
| `%LOCALAPPDATA%\CallScheduler\` | Embedded Python + dependencies (~400 MB) | Yes — use `uninstall.bat` (then `install.bat` to recreate). |
| `%TEMP%\tmp*` | Streamlit upload staging dirs | Yes — `run.bat` cleans up dirs older than 7 days at startup. |

---

## For developers

- `CLAUDE.md` — architecture, scheduler internals, scoring/ranking design.
- `docs/gui_plan.md` — GUI design spec and as-built notes.
- `src/tests/` — pytest suite. Run with `.venv/Scripts/python.exe -m pytest src/tests/`.
- Dev launch (skips `run.bat`): `.venv/Scripts/python.exe -m streamlit run src/app.py` from the project root.
- CLI scheduler (no GUI): `.venv/Scripts/python.exe src/scheduler_main.py` from the project root.
- Project layout: end-user-facing files (`README.md`, `install.bat`, `run.bat`, `uninstall.bat`) and the two data folders (`input_files/`, `output_files/`) live at the root. All Python source, `config.yaml`, and the test suite live in `src/` (`src/tests/`). `requirements.txt` and design docs live in `docs/`. `src/scheduler_main.py` and `src/app.py` each add `src/` to `sys.path` at import time so existing flat imports (`from config import X`) work unchanged.
