---

FAMILY MEDICINE CALL SCHEDULER
Author: Johnny McMurray
Purpose: Generate a balanced resident call schedule for an academic year.

---

OVERVIEW

This program automatically generates a full academic-year call schedule for a residency program.

It accounts for:

• PGY level (intern vs upper level)
• Rotation call eligibility
• Resident vacation / no-call requests
• Post-call restrictions
• Call spacing preferences
• Fair distribution of weekday and weekend call
• Holiday highlighting
• Night float intern display

The scheduler assigns call shifts using a penalty-based scoring system that chooses the resident with the lowest penalty score for each assignment.

The output is written to Excel files for easy review and editing by chief residents.

---

PROJECT STRUCTURE

The program is organized into several Python modules.

scheduler_core.py
Main scheduling engine and scoring logic.

excel_reader.py
Handles reading and interpreting the master flow sheet.

loaders.py
Loads configuration and input files.

validators.py
Runs data validation checks before scheduling begins.

exports.py
Writes Excel output files.

data/
Folder containing all configuration and input files.

---

REQUIRED INPUT FILES

All input files are stored inside the "data" folder.

1. FLOW SHEET

File: flow.xlsx
Sheet name: master_block_calendar

Structure:

Row 1
Block titles

Row 2
Block date ranges (example: JUL 1 - JUL 27)

Row 3+
Resident rows

Column A
Resident name

Remaining columns
Rotation assignments for each block

Notes:

• Rotation names must match rotation_rules.csv exactly
• Resident names must match across all files exactly
• Blocks with two rotations should be written like:
NF/ENDO

---

2. ROTATION RULES

File: rotation_rules.csv

Defines whether a rotation allows call.

Columns:

rotation_name
preference
pgy

Preference values:

ELIGIBLE
Resident can take call normally

AVOID
Call is discouraged but still allowed

NO_CALL
Resident cannot take call

Example:

rotation_name,preference,pgy
ICU,NO_CALL,1
INP,ELIGIBLE,2
NF,ELIGIBLE,1

---

3. NO CALL DAYS

File: no_call_days.csv

Defines vacation or unavailable dates.

Columns:

name
start_date
end_date
type

Example:

name,start_date,end_date,type
Hall,2026-08-10,2026-08-16,vacation
Prabhu,2027-03-05,2027-03-05,sick

Notes:

• start_date and end_date use YYYY-MM-DD format
• A single-day entry should use the same start and end date
• The "type" column is informational only

---

4. HOLIDAYS

File: holidays.csv

Defines holidays to highlight in the output schedule.

Columns:

date
name

Example:

date,name
2026-07-04,Independence Day
2026-12-25,Christmas

Dates must be in YYYY-MM-DD format.

---

5. CONFIGURATION

File: config.yaml

This file contains all adjustable scheduler parameters.

Examples include:

• fairness weighting
• call spacing rules
• avoid penalties
• year progression modifiers

Comments inside this file explain the purpose of each setting.

Users can adjust scheduler behavior by editing this file.

---

OUTPUT FILES

The program produces Excel output files inside the data folder.

daily_call_list.xlsx

Columns:

Block
Date
Day
Upper Level
Intern
No Call

Features:

• Weekend days highlighted yellow
• Holidays highlighted red
• New blocks separated visually
• Night float interns shown on weekdays
• Residents on no-call days listed

This file is the primary schedule used by chiefs.

---

output_summary.xlsx

Columns:

name
pgy
weekday_calls
weekend_calls
total_calls

Features:

• Each PGY level is color coded
• Used to confirm fairness of call distribution

---

SCHEDULING RULES

Interns (PGY1)

• Weekend call only
• No weekday call

Upper levels (PGY2 and PGY3)

• Weekday and weekend call

Other rules:

• Post-call residents cannot take call the following day
• Calls are spaced whenever possible
• AVOID rotations add penalties but do not block call
• NO_CALL rotations prevent assignments entirely

---

SCORING PHILOSOPHY

For each call slot:

1. Determine which residents are eligible
2. Calculate a penalty score for each resident
3. Assign the resident with the lowest score

Lower score = more favorable assignment.

Penalties include:

• fairness imbalance
• recent call spacing
• avoid rotations
• PGY2 / PGY3 year balancing

This system ensures the scheduler always selects the least penalized candidate.

---

VALIDATION CHECKS

Before scheduling begins the program checks for:

• rotations missing from rotation_rules.csv
• resident names in no_call_days not found in flow sheet
• blank rotation cells in the flow sheet
• duplicate resident names

These checks prevent common data entry mistakes.

---

RUNNING THE PROGRAM

Run:

python scheduler_core.py

The program will:

1. Load all inputs
2. Validate data
3. Generate the schedule
4. Export Excel output files

---

IMPORTANT NOTES

Names must match exactly between all files.

Rotation names must match exactly between the flow sheet and rotation_rules.csv.

Dates must use the YYYY-MM-DD format.

---

END OF DOCUMENT

---
