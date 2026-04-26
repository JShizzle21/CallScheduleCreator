# Call Schedule Creator

A Windows desktop app that builds a year-long medical residency call schedule from a few Excel input files. Optimizes for fairness, post-call rest, rotation eligibility, and spacing — and produces an Excel schedule, a per-resident call totals report, and an audit report.

---

## Quick start

1. **Copy the entire project folder** to your computer (anywhere — Desktop, Documents, etc.). Keep all files together.
2. **Double-click `install.bat`.** A console window opens and runs for ~5 minutes the first time. Wait for the `Setup complete.` message, then close the window.
3. **Double-click `run.bat`** whenever you want to use the app. Your default browser opens automatically after a few seconds.
4. **Stop the app** when done — easiest is the **Exit** button at the top right of the browser page (it shuts down the server cleanly). You can also just close the console window.

To uninstall, double-click **`uninstall.bat`** — it removes the embedded Python and ~400 MB of dependencies from `%LOCALAPPDATA%\CallScheduler\`. The project folder itself you delete by hand from Explorer.

> First-launch warnings you may see — these are normal:
> - **Windows SmartScreen** ("Windows protected your PC"): click "More info" → "Run anyway."
> - **Windows Defender Firewall** ("allow Python to communicate"): tick **Private networks** and click Allow.
> - **Antivirus** flag on `.bat` files: add the project folder to your AV's exclusions if needed.

If you ever need a clean reinstall, double-click `uninstall.bat`, then `install.bat`.

---

## What you need to prepare

All input files live in the `data/` folder. You can edit them in Excel; close the file before running the app (Excel locks open files).

### Required (4)

| File | What it holds |
|---|---|
| **`flow.xlsx`** | Block calendar — each resident's rotation for each ~2-week block of the academic year. Sheet name `master_block_calendar`. Row 2 = block start dates; column A = resident names; cells = rotation codes (`WARDS`, `ED`, `NF`, etc.). PGY cohorts are separated by rows containing date-like cells. |
| **`rotation_rules.xlsx`** | For each `(rotation, PGY level)` pair: whether they take call (`AVOID`, `NO_CALL`, etc.). |
| **`no_call_days.xlsx`** | Per-resident days off (vacation, conferences). Columns: `name`, `date`. |
| **`holidays.xlsx`** | Holiday dates with optional manual upper/intern assignments. Blank cells will be flagged as unassigned in the output for you to fill in by hand before shipping. |

### Optional (2)

| File | What it holds |
|---|---|
| **`clinic_days.xlsx`** | Clinic days. The day BEFORE each clinic is automatically blocked as a no-call day. Leave the file empty or omit it if you don't track clinics. |
| **`completed_calls.xlsx`** | Mid-year handoff: existing assignments to seed from. Only used when "Partial year" is enabled in the GUI. |

> **Tip:** the files currently in `data/` are working examples. Copy one and edit, rather than starting from scratch.

---

## Reading the output

After a successful run, three files land in `data/output/` and are also offered as download buttons in the GUI.

| File | Audience | What's in it |
|---|---|---|
| **`call_schedule.xlsx`** | The residents — this is the published schedule | Day-by-day calendar with assigned upper-level and intern. Holidays highlighted. |
| **`call_totals.xlsx`** | Chief / scheduler — fairness sanity check | Per-resident counts: total weekday calls, weekend calls, holidays. Color-coded gaps. |
| **`audit_report.txt`** | Anyone debugging an issue | Plain-text summary: errors, warnings, fairness gaps, list of unassigned slots, and reasoning. **Read this first if anything looks off.** |

---

## Day-to-day workflow

1. **Update `data/flow.xlsx`** with the new academic year's rotations.
2. **Update `data/no_call_days.xlsx`** with the year's vacations and conferences as you collect them.
3. **Update `data/holidays.xlsx`** with the year's holidays. Leave the upper/intern columns blank if you want the scheduler to assign them, or fill them in to lock specific people.
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

1. Fill in `data/completed_calls.xlsx` with all calls already assigned (one row per day, columns: `date`, upper name, intern name).
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
| Project folder (this directory) | Source code, your input files (`data/`), output files (`data/output/`) | Backup `data/` first. Everything else is replaceable from the repo. |
| `%LOCALAPPDATA%\CallScheduler\` | Embedded Python + dependencies (~400 MB) | Yes — use `uninstall.bat` (then `install.bat` to recreate). |
| `%TEMP%\tmp*` | Streamlit upload staging dirs | Yes — `run.bat` cleans up dirs older than 7 days at startup. |

---

## For developers

- `CLAUDE.md` — architecture, scheduler internals, scoring/ranking design.
- `docs/gui_plan.md` — GUI design spec and as-built notes.
- `tests/` — pytest suite. Run with `.venv/Scripts/python.exe -m pytest tests/`.
- Dev launch (skips `run.bat`): `.venv/Scripts/python.exe -m streamlit run src/app.py` from the project root.
- Project layout: end-user-facing files (`README.md`, `install.bat`, `run.bat`, `requirements.txt`, `scheduler_main.py`) live at the root; internal modules and `config.yaml` live in `src/`. `scheduler_main.py` and `src/app.py` each prepend the appropriate path to `sys.path` at import time, so existing flat imports (`from config import X`) work unchanged.
