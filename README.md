# NL Stats — American Football Officiating Analysis

This project collects and analyses officiating grades from American football games in the Danish National League (NL). The goal is to provide officials, crew chiefs, and supervisors with structured performance data based on graded penalty calls.

---

## How It Works

Each game is recorded in an Excel file with one row per play. When a penalty occurs, the row includes a foul code, an overall crew grade (FLAG), and a code that links individual grades to the specific official(s) involved. A separate schedule file tracks which officials worked each game and in what positions.

### Grading System

Each penalty call is graded on the following scale:

| Code | Name | Score |
|------|------|-------|
| CC / C | Correct Call | 100 |
| MC / M | Marginal Call | 75 |
| NC / N | No Call | 50 |
| IC / I | Incorrect Call | 0 |
| NG / G | Non-gradeable | Excluded |
| W | Waived | Excluded |

The FLAG column holds the overall crew grade — did the crew collectively handle the penalty correctly? The GRADE OFFICIAL column holds the individual breakdown — which officials were involved and how did each perform? These two are kept separate in the data so both crew-level and individual analysis is possible.

### Official Position Codes

Grades are linked to officials using single-letter position codes — for example `BC` means the Back Judge received a Correct Call grade, and `BCHC` means both the Back Judge and Head Linesman received Correct Call grades.

| Code | Position |
|------|----------|
| R | Referee |
| U | Umpire |
| H | Head Linesman |
| L | Line Judge |
| S | Side Judge |
| F | Field Judge |
| B | Back Judge |
| C | Center Judge |

Note: The letter `C` is used both as a position code (Center Judge) and a grade code (Correct). The parser handles this automatically by always reading characters in pairs — the first letter is the position, the second is the grade.

---

## Folder Structure

All scripts must be run from the project root folder (`nlstats/`). The folder structure must look like this:

```
nlstats/
├── data/               ← game CSV files (produced by 02_convert_to_csv.py)
│   └── *.xlsx          ← original game Excel files go here
├── nlplan/             ← schedule Excel file
│   └── NL_dommerplan_2025.xlsx
├── output/             ← all reports are written here (auto-created)
├── 01_check_files.py
├── 02_convert_to_csv.py
├── 03_build_flat_file.py
├── 04_generate_reports.py
└── run_all.py
```

---

## Pipeline

The analysis is built as a series of single-purpose scripts that are run in order. The easiest way to run everything is:

```bash
python3 run_all.py
```

Or run individual steps:

```bash
python3 01_check_files.py
python3 02_convert_to_csv.py
python3 03_build_flat_file.py
python3 04_generate_reports.py
```

Once game files have been converted to CSV you can skip step 02 on subsequent runs:

```bash
python3 run_all.py --skip-convert
```

All scripts use Python's standard library only and require no external packages. If you see NumPy or pandas warnings at startup, they are harmless — add `-W ignore` to suppress them:

```bash
python3 -W ignore 03_build_flat_file.py
```

---

## Script 01 — Check Files

`01_check_files.py` checks that your game Excel files match the games listed in the schedule. It generates an HTML report showing which games have files and which are missing.

**Output:** `output/troubleshooting_report.html`

Open the report in your browser after running. The report shows:

- **Found** (green) — game is in the schedule and a matching file exists
- **Missing** (red) — game is in the schedule but no matching file was found
- **Unmatched** (yellow) — a file exists but doesn't match any scheduled game

If the schedule file cannot be read, a diagnostic error report is written instead, showing the first 10 rows of the file so you can spot formatting problems.

### Common mistakes

**Game file name does not match the Game ID in the schedule.** The file name (without `.xlsx`) must exactly match the `GameID` column in the schedule. Capitalisation and spacing matter.

```
✓  31August-Towers-v-Razorbacks.xlsx    (matches GameID exactly)
✗  31august-towers-v-razorbacks.xlsx    (wrong capitalisation)
✗  31 August - Towers v Razorbacks.xlsx (spaces not allowed)
```

**The `GameID` column header is missing.** The schedule sheet must have a column with the header `GameID` (case sensitive, no spaces). If that cell is blank the script will not find any game IDs. Add `GameID` as the header in that column in the Excel file.

**Junk rows at the bottom of the schedule.** Formula results or summary rows at the bottom of the schedule sheet may be picked up as game IDs. Delete or move them outside the data range.

---

## Script 02 — Convert to CSV

`02_convert_to_csv.py` converts game Excel files in `data/` to CSV format. This is a one-time conversion step needed because the game Excel files have a formatting quirk (missing `styles.xml`) that prevents standard Excel readers from opening them. Converting to CSV removes all formatting and produces clean data files.

**Output:** one `.csv` file per `.xlsx` file, written into `data/`

The original Excel files are not modified or deleted.

**The schedule file in `nlplan/` does not need converting** — it reads fine as-is.

### Common mistakes

**Running the script from the wrong folder.** Always run scripts from the `nlstats/` project root, not from inside `data/` or any subfolder.

**Converting the schedule file.** Only game files in `data/` are converted. Do not move the schedule file into `data/`.

---

## Script 03 — Build Flat File

`03_build_flat_file.py` reads all game CSV files and the schedule file, then produces a single flat CSV with one row per graded official call. This file is the input for all reporting.

**Output:** `output/flat_calls.csv`

For each game the script:
1. Matches the filename to a Game ID in the schedule to get teams, date and officials
2. Reads every play and finds rows with penalties
3. Reads the FLAG column (overall crew grade) and GRADE OFFICIAL column (individual grades)
4. Splits multi-official codes (e.g. `LCHC`) into one row per official
5. Looks up the official's name from the officials database in the schedule file
6. Writes one row per graded official call

Plays with two separate penalties (PENALTY-CAT 1 and PENALTY CAT 2) are both processed. If a penalty has no GRADE OFFICIAL code the row is still included with blank position and grade fields so no penalty data is silently lost.

### Output columns

| Column | Description |
|--------|-------------|
| game_id | Matches the filename and GameID in the schedule |
| date | Date of the game |
| home_team | Home team |
| away_team | Away team |
| play_number | Play number within the game |
| qtr | Quarter |
| foul_code | Penalty code (e.g. DOF-NZI, FST, OFH-TD) |
| flag | Overall crew grade for this penalty (e.g. CC, MC, IC) |
| position | Single letter position code (R, U, H, L, S, F, B, C) |
| official_initials | Initials of the official in that position |
| official_name | Full name of the official |
| grade_code | Single letter individual grade (C, M, I, N, G, W) |

### Common mistakes

**Official initials or names are blank.** This means the game was not found in the schedule, or the officials were not assigned in the schedule yet. Check that the game file name matches the `GameID` exactly.

**No rows in output.** All plays may have empty PENALTY-CAT columns. Check that the game file has data in the `PENALTY-CAT 1` and `GRADE OFFICIAL 1` columns.

---

## Script 04 — Generate Reports

`04_generate_reports.py` reads `output/flat_calls.csv` and generates HTML reports.

**Output:**
- `output/combined_report.html` — season overview for all audiences
- `output/officials/{initials}.html` — one individual report per official

### Combined report sections

- **Game Summary** — one row per game with penalty count, crew accuracy and flag breakdown
- **Game by Game Breakdown** — detailed section per game with officials table (sorted by position) and full penalty list
- **Flag Breakdown** — counts of CC, MC, IC etc. across all games
- **Penalty Analysis** — fouls grouped by category (PF, OFH, DPI, OPI, UC, DOF) with flag breakdown and subcode counts
- **Officials List** — all officials alphabetically with games, positions, accuracy and grade breakdown
- **Season Accuracy Ranking** — officials ranked by accuracy (minimum 3 games to qualify)
- **Position Rankings** — best official at each position (minimum 2 games at that position to qualify)

All tables are interactive — click any column header to sort, and use the filter box above each table to search.

### Individual report sections

- Summary cards (accuracy, games, graded calls, positions worked)
- Grade breakdown (C, M, I, N, G, W counts and percentages)
- Performance by game (accuracy trend with visual bar)
- Game by game breakdown with full call list sorted by position then play number

### Foul code display

Foul codes are shown in full wherever possible:

- Known exact code → `OFH-TD — Holding, offense, takedown`
- Known parent with unknown subcode → `DOF-NZI — Offside, defense (NZI)`
- Unknown code → shown as-is

### Scoring

Accuracy is calculated as the weighted average of all scorable grades:

| Grade | Score |
|-------|-------|
| C | 100 |
| M | 75 |
| N | 50 |
| I | 0 |
| G | Excluded |
| W | Excluded |

Colour coding: green ≥ 90%, yellow ≥ 75%, orange ≥ 60%, red < 60%.

### Common mistakes

**No individual reports generated.** If officials show as 0 it means no officials were matched from the schedule. Check that the schedule has officials assigned in the position columns (R, U, H, L, S, F, B, C) and that the game file names match the Game IDs.

**Links between combined report and individual reports are broken.** The combined report links to `officials/{initials}.html` using relative paths. Both files must remain in their generated locations — do not move the combined report out of `output/` or the individual reports out of `output/officials/`.

---

## Schedule File Format

The schedule file must be an Excel file (`.xlsx`) placed in the `nlplan/` folder. It must contain a sheet named exactly `Plan - NL` with the following columns in the first row:

| Column | Description |
|--------|-------------|
| `GameID` | Unique identifier — must match the game file name exactly |
| `Dato` | Day number (e.g. `31`) |
| `Måned` | Month name (e.g. `August`) |
| `Hjemme` | Home team name |
| `Ude` | Away team name |
| `R` | Referee initials |
| `U` | Umpire initials |
| `H` | Head Linesman initials |
| `L` | Line Judge initials |
| `S` | Side Judge initials |
| `F` | Field Judge initials |
| `B` | Back Judge initials |
| `C` | Center Judge initials |

The officials database must be on a second sheet named exactly `Officials and games` with columns `Initialer` (initials) and `Navn` (full name).

Rows without a `GameID` value are skipped automatically, so pre-season or practice game rows do not need to be deleted.

---

## Game File Format

Game files must be Excel files (`.xlsx`) placed in the `data/` folder. The file name (without `.xlsx`) must exactly match the `GameID` in the schedule.

The file must have a single sheet with the following columns:

| Column | Description |
|--------|-------------|
| `PLAY #` | Play number |
| `QTR` | Quarter |
| `PENALTY-CAT 1` | Foul code for first penalty on the play |
| `FLAG 1` | Crew grade for first penalty (e.g. CC, MC) |
| `GRADE OFFICIAL 1` | Individual grades for first penalty (e.g. LCHC) |
| `PENALTY CAT 2` | Foul code for second penalty (if any) |
| `FLAG 2` | Crew grade for second penalty |
| `GRADE OFFICIAL 2` | Individual grades for second penalty |

Plays with no penalty are left blank and are skipped automatically.
