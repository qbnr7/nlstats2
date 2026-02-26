#!/usr/bin/env python3
"""
03_build_flat_file.py
=====================
Reads all game CSV files and the schedule file, then produces a single
flat CSV file with one row per graded official call.

Each penalty on a play that involves multiple officials (e.g. LCHC) is
split into individual rows - one per official. Plays with two penalties
(PENALTY-CAT 1 and PENALTY CAT 2) are both processed.

Input:
    data/           - folder with one CSV file per game (from 02_convert_to_csv.py)
                      filename must match the GameID in the schedule
    nlplan/         - folder with the schedule Excel file
                      requires sheets: 'Plan - NL' and 'Officials and games'

Output:
    output/flat_calls.csv  - one row per graded official call

Columns in output:
    game_id            - matches the filename / GameID in schedule
    date               - date of the game (from schedule)
    home_team          - home team (from schedule)
    away_team          - away team (from schedule)
    play_number        - play number within the game
    qtr                - quarter
    foul_code          - foul/penalty code (e.g. DOF, FST)
    flag               - overall crew grade for this penalty (e.g. CC, MC)
    position           - single letter position code (R, U, H, L, B, F, S, C)
    official_initials  - initials of official in that position (from schedule)
    official_name      - full name of official (from officials database)
    grade_code         - single letter grade (C, M, I, N, G, W)

Usage:
    python 03_build_flat_file.py
"""

import pandas as pd
import zipfile
import io
import xml.etree.ElementTree as ET
from pathlib import Path
from openpyxl import load_workbook

# ── Configuration ─────────────────────────────────────────────────────────────

DATA_FOLDER     = Path("data")
SCHEDULE_FOLDER = Path("nlplan")
OUTPUT_FOLDER   = Path("output")

SCHEDULE_SHEET  = "Plan - NL"
OFFICIALS_SHEET = "Officials and games"

POSITION_CODES = {'R', 'U', 'H', 'L', 'B', 'F', 'S', 'C'}
GRADE_CODES    = {'C', 'M', 'I', 'N', 'G', 'W'}

PENALTY_COLUMNS = [
    ('PENALTY-CAT 1', 'FLAG 1', 'GRADE OFFICIAL 1'),
    ('PENALTY CAT 2',  'FLAG 2', 'GRADE OFFICIAL 2'),
]

# ── Schedule reader (Excel - works fine) ──────────────────────────────────────

def load_xlsx_sheet(file_path, sheet_name):
    """
    Read a sheet from an xlsx file using openpyxl.
    Returns a pandas DataFrame.
    """
    wb = load_workbook(file_path, read_only=True, data_only=True)
    if sheet_name not in wb.sheetnames:
        raise ValueError(f"Sheet '{sheet_name}' not found. "
                         f"Available: {wb.sheetnames}")
    ws = wb[sheet_name]
    data = [[cell.value for cell in row] for row in ws.rows]
    wb.close()

    if not data:
        return pd.DataFrame()

    headers = [str(h).strip() if h is not None else '' for h in data[0]]
    return pd.DataFrame(data[1:], columns=headers)

# ── Data loaders ──────────────────────────────────────────────────────────────

def load_officials(schedule_file):
    """
    Load officials database. Returns dict: { initials -> full name }
    """
    df = load_xlsx_sheet(schedule_file, OFFICIALS_SHEET)

    # Row 0 is a merged title row, row 1 contains the actual column headers
    df.columns = [str(c).strip() for c in df.iloc[1]]
    df = df[2:].reset_index(drop=True)

    officials = {}
    for _, row in df.iterrows():
        initials = str(row.get('Initialer', '') or '').strip()
        name     = str(row.get('Navn',      '') or '').strip()
        if initials and name and initials.lower() != 'nan' and name.lower() != 'nan':
            officials[initials] = name

    print(f"  Loaded {len(officials)} officials")
    return officials


def load_schedule(schedule_file):
    """
    Load game schedule. Returns dict: { game_id -> { date, home_team,
    away_team, positions } }
    """
    df = load_xlsx_sheet(schedule_file, SCHEDULE_SHEET)

    schedule = {}
    for _, row in df.iterrows():
        game_id = str(row.get('GameID', '') or '').strip()
        if not game_id or game_id.lower() == 'nan':
            continue

        dato  = str(row.get('Dato',  '') or '').strip()
        maned = str(row.get('Måned', '') or '').strip()
        date  = f"{dato}-{maned}" if dato and maned else ''

        positions = {}
        for pos in POSITION_CODES:
            val = str(row.get(pos, '') or '').strip()
            if val and val.lower() != 'nan':
                positions[pos] = val.split('+')[0].strip()

        schedule[game_id] = {
            'date':      date,
            'home_team': str(row.get('Hjemme', '') or '').strip(),
            'away_team': str(row.get('Ude',    '') or '').strip(),
            'positions': positions,
        }

    print(f"  Loaded {len(schedule)} games from schedule")
    return schedule

# ── Grade official parser ──────────────────────────────────────────────────────

def parse_grade_official(code):
    """
    Parse a GRADE OFFICIAL string into a list of (position, grade) pairs.

    e.g. 'RC'   -> [('R', 'C')]
         'LCHC' -> [('L', 'C'), ('H', 'C')]
         'LCRN' -> [('L', 'C'), ('R', 'N')]
    """
    if code is None or str(code).strip() == '':
        return []

    code = str(code).strip().upper()
    pairs = []
    i = 0

    while i < len(code) - 1:
        pos   = code[i]
        grade = code[i + 1]

        if pos in POSITION_CODES and grade in GRADE_CODES:
            pairs.append((pos, grade))
            i += 2
        else:
            print(f"    Warning: unexpected characters '{code[i:i+2]}' "
                  f"in '{code}', skipping")
            i += 1

    return pairs

# ── Game file processor ────────────────────────────────────────────────────────

def process_game_file(file_path, game_id, game_info, officials):
    """
    Process a single game CSV file.
    Returns list of flat row dicts, one per graded official call.
    """
    df = pd.read_csv(file_path, dtype=str).fillna('')
    df.columns = df.columns.str.strip()

    rows = []

    for _, play in df.iterrows():
        play_number = play.get('PLAY #', '')
        qtr         = play.get('QTR',    '')

        for penalty_col, flag_col, grade_col in PENALTY_COLUMNS:
            foul_code      = play.get(penalty_col, '').strip()
            flag           = play.get(flag_col,    '').strip()
            grade_official = play.get(grade_col,   '').strip()

            if not foul_code:
                continue

            pairs = parse_grade_official(grade_official)

            if not pairs:
                rows.append(build_row(game_id, game_info, play_number, qtr,
                                      foul_code, flag, '', '', '', ''))
                continue

            for position, grade in pairs:
                initials = game_info['positions'].get(position, '')
                name     = officials.get(initials, '') if initials else ''
                rows.append(build_row(game_id, game_info, play_number, qtr,
                                      foul_code, flag, position, initials,
                                      name, grade))
    return rows


def build_row(game_id, game_info, play_number, qtr,
              foul_code, flag, position, initials, name, grade):
    return {
        'game_id':           game_id,
        'date':              game_info['date'],
        'home_team':         game_info['home_team'],
        'away_team':         game_info['away_team'],
        'play_number':       play_number,
        'qtr':               qtr,
        'foul_code':         foul_code,
        'flag':              flag,
        'position':          position,
        'official_initials': initials,
        'official_name':     name,
        'grade_code':        grade,
    }

# ── Main ──────────────────────────────────────────────────────────────────────

def main():
    print("03_build_flat_file.py")
    print("=" * 50)

    for folder in [DATA_FOLDER, SCHEDULE_FOLDER]:
        if not folder.exists():
            print(f"ERROR: Folder '{folder}' not found")
            return

    OUTPUT_FOLDER.mkdir(exist_ok=True)

    schedule_files = list(SCHEDULE_FOLDER.glob("*.xlsx")) + \
                     list(SCHEDULE_FOLDER.glob("*.xls"))
    if not schedule_files:
        print(f"ERROR: No Excel files found in '{SCHEDULE_FOLDER}'")
        return
    schedule_file = schedule_files[0]
    print(f"\nSchedule file : {schedule_file.name}")

    officials = load_officials(schedule_file)
    schedule  = load_schedule(schedule_file)

    game_files = sorted(DATA_FOLDER.glob("*.csv"))
    if not game_files:
        print(f"ERROR: No CSV files found in '{DATA_FOLDER}'")
        print(f"       Run 02_convert_to_csv.py first")
        return
    print(f"  Found {len(game_files)} game CSV file(s)\n")

    all_rows  = []
    matched   = 0
    unmatched = []

    for file_path in game_files:
        game_id = file_path.stem

        if game_id in schedule:
            game_info = schedule[game_id]
            matched += 1
        else:
            print(f"  Warning: '{game_id}' not found in schedule")
            game_info = {'date': '', 'home_team': '', 'away_team': '',
                         'positions': {}}
            unmatched.append(game_id)

        print(f"  Processing : {file_path.name}")
        rows = process_game_file(file_path, game_id, game_info, officials)
        all_rows.extend(rows)
        print(f"  Rows output: {len(rows)}\n")

    if not all_rows:
        print("No graded calls found - nothing written")
        return

    output_path = OUTPUT_FOLDER / "flat_calls.csv"
    pd.DataFrame(all_rows, columns=[
        'game_id', 'date', 'home_team', 'away_team',
        'play_number', 'qtr', 'foul_code', 'flag',
        'position', 'official_initials', 'official_name', 'grade_code'
    ]).to_csv(output_path, index=False)

    print("-" * 50)
    print(f"Output file   : {output_path}")
    print(f"Total rows    : {len(all_rows)}")
    print(f"Games matched : {matched}")
    if unmatched:
        print(f"Not in schedule ({len(unmatched)}):")
        for g in unmatched:
            print(f"  - {g}")


if __name__ == "__main__":
    main()
