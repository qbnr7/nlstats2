# NL Stats - American Football Officiating Analysis

This project collects and analyses officiating grades from American football games in the Danish National League (NL). The goal is to provide officials, crew chiefs, and supervisors with structured performance data based on graded penalty calls.

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
| B | Back Judge |
| F | Field Judge |
| S | Side Judge |
| C | Center Judge |

## Pipeline

The analysis is built as a series of single-purpose scripts that are run in order:
```
01_check_files.py       ← check game files match schedule
02_convert_to_csv.py    ← convert game Excel files to CSV
03_build_flat_file.py   ← build flat_calls.csv from CSVs and schedule
04_generate_report.py   ← generate reports from flat_calls.csv
```

Run them with:
```bash
python3 -W ignore 01_check_files.py
python3 -W ignore 02_convert_to_csv.py
python3 -W ignore 03_build_flat_file.py
python3 -W ignore 04_generate_report.py
```

The `-W ignore` flag suppresses NumPy version warnings that appear on some systems. The scripts run correctly regardless.

---

## Script 01 - Check Files

The script `01_check_files.py` checks that your game Excel files match the games listed in the schedule. It generates an HTML report showing which games have files and which are missing.

### Folder Structure
```
nlstats2/
├── data/          ← game Excel files (one per game)
├── nlplan/        ← schedule Excel file with officials assignments
├── reports/       ← generated reports are saved here (auto-created)
└── 01_check_files.py
```

### Usage
```bash
# Generate troubleshooting report only
python3 01_check_files.py --troubleshooting-only

# Generate troubleshooting report alongside full analysis
python3 01_check_files.py --troubleshooting
```

Open `reports/troubleshooting_report.html` in your browser to see the results.

### What the Report Shows

- **Found** (green) — game is in the schedule and a matching Excel file exists
- **Missing** (red) — game is in the schedule but no matching file was found
- **Extra** (yellow) — an Excel file exists but doesn't match any scheduled game

### Game File Naming

Game files must match the Game ID in the schedule. For example:
```
13April-Oaks-v-Razorbacks.xlsx
12April-Tigers-v-89ers.xlsx
```

---

## Script 02 - Convert to CSV

The script `02_convert_to_csv.py` converts game Excel files to CSV format. This is a one-time conversion step needed because the game Excel files have a formatting issue that prevents standard Excel readers from opening them. Converting to CSV removes all formatting and produces clean data files that work reliably on any system.

The schedule file in `nlplan/` does not need converting as it reads fine.

### Folder Structure
```
nlstats2/
├── data/          ← game Excel files go in, CSV files come out
└── 02_convert_to_csv.py
```

### Usage
```bash
python3 02_convert_to_csv.py
```

Each `.xlsx` file in `data/` will produce a matching `.csv` file in the same folder. The original Excel files are not modified or deleted.

---

## Script 03 - Build Flat File

The script `03_build_flat_file.py` reads all game CSV files and the schedule file, then produces a single flat CSV file with one row per graded official call. This CSV is the input for all subsequent reporting scripts.

### How It Works

For each game CSV file the script:
1. Matches the filename to a Game ID in the schedule to get teams, date and officials
2. Reads every play and finds rows with penalties
3. Reads the FLAG column which holds the overall crew grade for the penalty
4. Parses the GRADE OFFICIAL code to extract individual position and grade pairs
5. Looks up the official's initials and full name from the officials database
6. Writes one row per graded official call to the output CSV

A penalty involving multiple officials (e.g. `LCHC`) is split into separate rows — one for the Line Judge and one for the Head Linesman. The crew flag is repeated on both rows. Plays with two separate penalties (PENALTY-CAT 1 and PENALTY CAT 2) are both processed.

### Folder Structure
```
nlstats2/
├── data/          ← game CSV files (from 02_convert_to_csv.py)
├── nlplan/        ← schedule Excel file
├── output/        ← flat_calls.csv is written here (auto-created)
└── 03_build_flat_file.py
```

### Usage
```bash
python3 -W ignore 03_build_flat_file.py
```

Output is written to `output/flat_calls.csv`.

### Output Format

One row per graded official call with the following columns:

| Column | Description |
|--------|-------------|
| game_id | Matches the filename and GameID in the schedule |
| date | Date of the game |
| home_team | Home team |
| away_team | Away team |
| play_number | Play number within the game |
| qtr | Quarter |
| foul_code | Penalty code (e.g. DOF, FST) |
| flag | Overall crew grade for this penalty (e.g. CC, MC, IC) |
| position | Single letter position code (R, U, H, L, B, F, S, C) |
| official_initials | Initials of the official in that position |
| official_name | Full name of the official |
| grade_code | Single letter individual grade (C, M, I, N, G, W) |

### Notes

- The `C` position code (Center Judge) and the `C` grade code (Correct) use the same letter. The parser handles this automatically by reading characters in pairs — the first letter is always the position and the second is always the grade.
- If a penalty has no GRADE OFFICIAL code the row is still included in the output with blank position and grade fields, so no penalty data is silently lost.
- Official initials and names will be blank if the schedule has no officials assigned to that game yet.
