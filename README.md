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

Grades are linked to officials using single-letter position codes — for example `BC` means the Back Judge received a Correct Call grade, and `BCDC` means both the Back Judge and Down Judge received Correct Call grades.

| Code | Position |
|------|----------|
| R | Referee |
| U | Umpire |
| D | Down Judge |
| L | Line Judge |
| S | Side Judge |
| F | Field Judge |
| B | Back Judge |
| C | Center Judge |

Note: `D` (Down Judge) used to be called `H` (Head Linesman) in older schedules and game files. Both codes are still accepted as input — the scripts normalise `H` to `D` automatically, so every generated report only ever shows `D` / Down Judge, never `H`. `H` should not appear in current schedules or game files, though, so every time it's found (in a schedule position column or a GRADE OFFICIAL code) `03_build_flat_file.py` and `04_generate_reports.py` print a warning naming the exact game/play/column it came from, plus a one-line total count at the end of the run. Processing is not stopped — the value is still normalised to `D` and the run completes — but the warnings are there so stale `H` usage in the source files gets noticed and cleaned up rather than silently passing through.

Note: The letter `C` is used both as a position code (Center Judge) and a grade code (Correct). The parser handles this automatically by always reading characters in pairs — the first letter is the position, the second is the grade.

---

## Folder Structure

All scripts must be run from the project root folder (`nlstats2/`). The folder structure must look like this:

```
nlstats2/
├── data/               ← original game Excel files go here
│   ├── *.xlsx          ← one file per game, named to match the schedule's GameID
│   └── *.csv           ← optional CSV copies (produced by 02_convert_to_csv.py, not required by the pipeline)
├── nlplan/             ← schedule Excel file
│   └── NL_dommerplan_2025.xlsx
├── output/             ← all reports are written here (auto-created)
│   ├── flat_calls.csv
│   ├── combined_report.html
│   ├── officials/
│   └── games/
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

`02_convert_to_csv.py` converts game Excel files in `data/` to CSV format. This is useful as a quick, dependency-free way to eyeball a game file's raw contents (e.g. in a text editor or spreadsheet app) when troubleshooting.

**Output:** one `.csv` file per `.xlsx` file, written into `data/`

The original Excel files are not modified or deleted.

**The schedule file in `nlplan/` does not need converting** — it reads fine as-is.

**Note:** `03_build_flat_file.py` reads the original `.xlsx` game files directly — it does not use the `.csv` files this script produces. Both game files and the schedule can have a formatting quirk (a missing or malformed `styles.xml`) that trips up standard Excel readers; `03_build_flat_file.py` detects and repairs this itself before reading, so running this step first is optional and only needed if you want CSV copies of the game files for your own reference.

### Common mistakes

**Running the script from the wrong folder.** Always run scripts from the `nlstats2/` project root, not from inside `data/` or any subfolder.

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

The script reads each game file's raw `.xlsx` contents directly (via `zipfile` + `openpyxl`) and repairs two common real-world quirks automatically, with no need to touch the original files:
- A missing or malformed `styles.xml` (some exports omit it, or include a stub `<fill></fill>` with no colour/pattern info that trips up strict Excel readers) is replaced with a minimal valid one before the file is parsed.
- Rows with sparsely-filled trailing columns (e.g. an empty "notes" column near the end of the sheet) are padded out so every row lines up with the header row, instead of raising a column-count mismatch.

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
| position | Single letter position code (R, U, D, L, S, F, B, C) -- legacy `H` is normalised to `D` |
| official_initials | Initials of the official in that position |
| official_name | Full name of the official |
| grade_code | Single letter individual grade (C, M, I, N, G, W) |

### Common mistakes

**Official initials or names are blank.** This means the game was not found in the schedule, or the officials were not assigned in the schedule yet. Check that the game file name matches the `GameID` exactly.

**No rows in output.** All plays may have empty PENALTY-CAT columns. Check that the game file has data in the `PENALTY-CAT 1` and `GRADE OFFICIAL 1` columns.

---

## Script 04 -- Generate Reports

`04_generate_reports.py` reads `output/flat_calls.csv` and generates HTML reports. It also reads the schedule file from `nlplan/` to cross-reference the full assigned crew per game.

**Dependencies:** `openpyxl` only. pandas is not used or required.

**Output:**
- `output/combined_report.html` -- season overview for all audiences
- `output/officials/{initials}.html` -- one individual report per official
- `output/games/{game_id}.html` -- one shareable report per game

### Full crew display

For each game, the script reads the schedule to find the complete list of assigned officials. Any official who was assigned but had no recorded calls is still shown in the officials table, greyed out with a circle marker and `--` in the accuracy column. This ensures the full crew is always visible even when some officials were not involved in any flagged play.

The schedule is matched to flat file game IDs automatically. If the schedule has a `GameID` column that column is used directly, and rows with a blank `GameID` cell (e.g. junk or formula/summary rows at the bottom of the sheet) are skipped. Only when the schedule has no `GameID` column at all (older schedules) is the game ID constructed from `Dato + Maaned + Hjemme + Ude` to match the filename format.

**Every "games" count in every report is based on schedule assignment, not on whether a flag was thrown.** An official's Games Officiated total, their row in the Officials List and Season Accuracy Ranking, and their Games at Position count in Position Rankings all include every game they were scheduled for in `nlplan/`, including games where they had zero recorded penalties at their position. Accuracy percentages are unaffected by this -- those are still only ever calculated from actual graded calls -- but the games/games-at-position counts themselves reflect assignment, not activity.

The same rule applies to **which positions show up** for an official, not just how many games are counted at each one. The Officials List "Positions" column and an individual report's "Positions Worked" summary card both list every position an official was ever scheduled for, even a position where they went every game without throwing a single flag -- it still appears, just with `Flags 0`, rather than being left off the list entirely.

### Combined report sections

- **Game Summary** -- one row per game with penalty count, crew accuracy and flag breakdown. Games are displayed as `10 Maj -- 89ers vs Oaks` (day and month taken directly from the game ID, underscores replaced with spaces).
- **Game by Game Breakdown** -- detailed section per game with officials table (sorted by position) and full penalty list. Assigned officials with no calls are shown greyed out.
- **Flag Breakdown** -- counts of CC, MC, IC etc. across all games
- **Foul Breakdown** -- interactive table with one row per foul type, showing crew flag percentages (CC%, IC% etc.) and individual accuracy. Filterable by category, foul name and minimum flag count.
- **Penalty Analysis** -- fouls grouped by category (PF, OFH, DPI, OPI, UC, DOF) with flag breakdown and subcode counts
- **Officials List** -- all officials alphabetically with games, positions, accuracy and grade breakdown
- **Season Accuracy Ranking** -- officials ranked by accuracy (minimum 3 games to qualify)
- **Position Rankings** -- best official at each position (minimum 2 games at that position to qualify)

All tables are interactive -- click any column header to sort, and use the filter box above each table to search. A sticky navigation panel on the right edge lets you jump between sections without scrolling.

### Individual report sections

- Summary cards: Overall Accuracy, Games Officiated, Flags Thrown (all calls including G and W), Graded Calls (C/M/N/I only), and Positions Worked (one line per position showing games and flags at that position)
- Grade breakdown (C, M, I, N, G, W counts and percentages)
- Performance by game (accuracy trend with visual bar; game shown as `10 Maj -- 89ers vs Oaks`)
- Game by game breakdown with full call list sorted by position then play number

### Per-game reports

One standalone HTML file is generated per game into `output/games/`. Each file contains:

- A crew accuracy banner at the top
- **Officials table** -- all assigned crew sorted by position, with accuracy and grade breakdown per official. Officials who were assigned but had no recorded calls are shown greyed out with a circle marker.
- **Penalties table** -- every flagged play in the game showing quarter, play number, foul, crew grade, official name, position and individual grade

The files are self-contained and can be shared directly -- for example posted to a Discord channel for the crew. All CSS is inline and the file has no external dependencies. It links back to `combined_report.html` and to individual official reports when opened locally.

The Game Summary table in the combined report has a `View` link for each game that opens the corresponding per-game file directly.

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

**No individual reports generated.** If officials show as 0 it means no officials were matched from the schedule. Check that the schedule has officials assigned in the position columns (R, U, D, L, S, F, B, C) and that the game file names match the Game IDs.

**An official is in the schedule but not shown in the game.** Make sure the game file name matches the `GameID` in the schedule exactly. Run `01_check_files.py` to diagnose mismatches.

**An official appears greyed out with a circle marker.** This is expected -- it means they were assigned in the schedule but had no calls recorded in the game file. It is not an error.

**Links between combined report and individual reports are broken.** The combined report links to `officials/{initials}.html` using relative paths. Both files must remain in their generated locations -- do not move the combined report out of `output/` or the individual reports out of `output/officials/`.

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
| `D` | Down Judge initials (legacy schedules may use `H` for this column instead — both are accepted) |
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

---

## GitHub — Pushing and Pulling

The project is hosted at `https://github.com/qbnr7/nlstats2`, on the `main` branch. Use the workflow below whenever you update scripts or add new game files.

### First-time setup (already done)

The repository is initialised and connected. You should not need to repeat these steps. If you ever need to clone it fresh onto a new machine:

```bash
git clone https://github.com/qbnr7/nlstats2.git
cd nlstats2
```

### Logging in — authenticating with GitHub

The first time you run `git push` or `git pull` on a machine (or after your saved credentials expire or get cleared), Git will stop and ask you to log in before it can talk to GitHub. What that looks like depends on your setup:

- **In a plain terminal**, you'll see prompts like:
  ```
  Username for 'https://github.com': your-github-username
  Password for 'https://your-github-username@github.com': 
  ```
- **On macOS/Windows with a Git GUI or credential manager installed**, a small login popup window may appear instead of a terminal prompt.

Either way, enter the same two things:
- **Username** — your GitHub username (not your email).
- **Password** — your **Personal Access Token**, not your actual GitHub account password. GitHub stopped accepting real account passwords for this years ago. See [Creating a Personal Access Token](#creating-a-personal-access-token) below if you don't have one yet.

Once you've logged in successfully once, see [Saving your token so you don't have to retype it](#saving-your-token-so-you-dont-have-to-retype-it) so you aren't asked again on every single push/pull.

If a login attempt fails with something like `Authentication failed` or `403`, the most common causes are: you typed your GitHub password instead of a token, the token expired, or the token's `repo` scope wasn't ticked when it was created — regenerate a token (below) and try again.

### Pulling — getting the latest version from GitHub

Run this before you start working to make sure your local copy is up to date:

```bash
cd ~/nlstats2
git pull origin main
```

If nothing has changed on GitHub since your last push it will say `Already up to date.`

### Pushing — sending your changes to GitHub

After updating scripts or adding new game files, run these three commands:

```bash
cd ~/nlstats2
git add .
git commit -m "Short description of what changed"
git push origin main
```

**Examples of good commit messages:**
- `"Add September game files"`
- `"Update 04_generate_reports with foul breakdown table"`
- `"Fix schedule loader picking up junk rows as games"`

If Git prompts you to log in at this point, see [Logging in — authenticating with GitHub](#logging-in--authenticating-with-github) above.

### Checking what has changed

To see which files have been modified or added before committing:

```bash
git status
```

To see a summary of recent commits:

```bash
git log --oneline -10
```

### If you only want to push specific files

Instead of `git add .` (which stages everything), you can stage individual files:

```bash
git add 04_generate_reports.py README.md
git commit -m "Update report script and docs"
git push origin main
```

### Creating a Personal Access Token

GitHub no longer accepts your account password for `git push` / `git pull` over HTTPS — you need a **Personal Access Token (PAT)** instead. A token acts like a password that's scoped just to this purpose and can be revoked at any time without changing your GitHub login.

1. Sign in to GitHub, then go to **Settings** (click your profile picture, top right → *Settings*).
2. Scroll down to **Developer settings** (bottom of the left-hand menu).
3. Click **Personal access tokens** → **Tokens (classic)** — the classic flow is simplest for a single personal repo like this one. (GitHub also offers "Fine-grained tokens" with more granular per-repo permissions, if you prefer.)
4. Click **Generate new token** → **Generate new token (classic)**.
5. Give it a descriptive name, e.g. `nlstats2-laptop`.
6. Set an **Expiration** (e.g. 90 days, or a custom date — GitHub recommends against "No expiration" for security, but it's your choice).
7. Under **Select scopes**, tick **`repo`** (this covers push/pull access to your repositories). No other scopes are needed for this project.
8. Click **Generate token** at the bottom.
9. **Copy the token immediately** — it's shown only once. Store it somewhere safe (a password manager is ideal).

You'll use this token as your password the next time Git asks for authentication (over HTTPS). If a token expires or is lost, just repeat the steps above to generate a new one.

### Saving your token so you don't have to retype it

By default, Git will ask for your username and token every time you push or pull. To avoid that:

**macOS** — Git usually already uses the macOS Keychain automatically. If not:
```bash
git config --global credential.helper osxkeychain
```

**Windows** — Git for Windows installs "Git Credential Manager" by default, which handles this automatically. If not:
```bash
git config --global credential.helper manager
```

**Linux** — cache your credentials in memory for a while (default 15 minutes; the example below extends it to 8 hours):
```bash
git config --global credential.helper 'cache --timeout=28800'
```

After setting this, the next `git push`/`git pull` that asks for a password will remember your token for future commands.
