# NL Stats - American Football Officiating Analysis

This project collects and analyses officiating grades from American football games in the Danish National League (NL). The goal is to provide officials, crew chiefs, and supervisors with structured performance data based on graded penalty calls.

## How It Works

Each game is recorded in an Excel file with one row per play. When a penalty occurs, the row includes a foul code, a grade for the call, and a code that links the grade to the specific official(s) involved. A separate schedule file tracks which officials worked each game and in what positions.

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

## Scripts

The analysis is built as a series of single-purpose scripts that can be run individually or in sequence. Each script is described in its own section below.

---

## Troubleshooting Script

The troubleshooting script (`troubleshooting_script.py`) checks that your game Excel files match the games listed in the schedule. It generates an HTML report showing which games have files and which are missing.

### Folder Structure

The script expects the following folder structure:
```
nlstats2/
├── data/          ← game Excel files (one per game)
├── nlplan/        ← schedule Excel file with officials assignments
├── reports/       ← generated reports are saved here (auto-created)
└── troubleshooting_script.py
```

### Usage

Run the script from the terminal in the repo folder:
```bash
# Generate troubleshooting report only
python troubleshooting_script.py --troubleshooting-only

# Generate troubleshooting report alongside full analysis
python troubleshooting_script.py --troubleshooting
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


## Script 02 - Build Flat File

The script `02_build_flat_file.py` reads all game Excel files and the schedule file, then produces a single flat CSV file with one row per graded official call. This CSV is the input for all subsequent reporting scripts.

### How It Works

For each game file the script:
1. Matches the filename to a Game ID in the schedule to get teams, date and officials
2. Reads every play and finds rows with penalties
3. Parses the GRADE OFFICIAL code to extract individual position and grade pairs
4. Looks up the official's initials and full name from the officials database
5. Writes one row per graded official call to the output CSV

A penalty involving multiple officials (e.g. `LCHC`) is split into separate rows — one for the Line Judge and one for the Head Linesman. Plays with two separate penalties (PENALTY-CAT 1 and PENALTY CAT 2) are both processed.

### Folder Structure
```
nlstats2/
├── data/          ← one Excel file per game, filename must match GameID in schedule
├── nlplan/        ← schedule Excel file
├── output/        ← flat_calls.csv is written here (auto-created)
└── 02_build_flat_file.py
```

### Usage
```bash
python 02_build_flat_file.py
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
| position | Single letter position code (R, U, H, L, B, F, S, C) |
| official_initials | Initials of the official in that position |
| official_name | Full name of the official |
| grade_code | Single letter grade (C, M, I, N, G, W) |

### Game File Naming

Game files must be named exactly after the GameID in the schedule. For example if the schedule has GameID `31August-Towers-v-Razorbacks` the file must be named:
```
31August-Towers-v-Razorbacks.xlsx
```

If a file does not match any GameID in the schedule the script will still process it but date, teams and official names will be blank. A warning is printed to the terminal.

### Notes

- The `C` position code (Center Judge) and the `C` grade code (Correct) use the same letter. The parser handles this automatically by reading characters in pairs — the first letter is always the position and the second is always the grade.
- If a penalty has no GRADE OFFICIAL code the row is still included in the output with blank position and grade fields, so no penalty data is silently lost.
- Official initials and names will be blank if the schedule has no officials assigned to that game yet.
