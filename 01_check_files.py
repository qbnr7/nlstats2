#!/usr/bin/env python3
"""
01_check_files.py
=================
Checks that game Excel files in data/ match the games listed in the
schedule file. Generates an HTML report showing what was found,
what is missing, and what files don't match any scheduled game.

No external dependencies - uses Python stdlib only.

Input:
    data/      - folder containing game Excel files
    nlplan/    - folder containing the schedule Excel file

Output:
    output/troubleshooting_report.html

Usage:
    python3 01_check_files.py
"""

import zipfile
import xml.etree.ElementTree as ET
from pathlib import Path
from datetime import datetime

# ── Configuration ─────────────────────────────────────────────────────────────

DATA_FOLDER     = Path("data")
SCHEDULE_FOLDER = Path("nlplan")
OUTPUT_FOLDER   = Path("output")

SCHEDULE_SHEET  = "Plan - NL"

NS = {
    'ss': 'http://schemas.openxmlformats.org/spreadsheetml/2006/main',
}

# ── xlsx reader (stdlib only) ─────────────────────────────────────────────────

def get_shared_strings(zf):
    if 'xl/sharedStrings.xml' not in zf.namelist():
        return []
    tree = ET.parse(zf.open('xl/sharedStrings.xml'))
    strings = []
    for si in tree.findall('.//ss:si', NS):
        text = ''.join(t.text or '' for t in si.findall('.//ss:t', NS))
        strings.append(text)
    return strings


def get_sheet_by_name(zf, sheet_name):
    """Find the sheet XML file for a given sheet name."""
    # Read workbook to find sheet names and their IDs
    wb_tree = ET.parse(zf.open('xl/workbook.xml'))
    wb_ns   = {'ss': 'http://schemas.openxmlformats.org/spreadsheetml/2006/main',
               'r':  'http://schemas.openxmlformats.org/officeDocument/2006/relationships'}

    sheet_id = None
    for sheet in wb_tree.findall('.//ss:sheet', wb_ns):
        if sheet.get('name') == sheet_name:
            sheet_id = sheet.get(
                '{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id'
            )
            break

    if not sheet_id:
        return None

    # Read relationships to find file path for this sheet ID
    rels_tree = ET.parse(zf.open('xl/_rels/workbook.xml.rels'))
    for rel in rels_tree.getroot():
        if rel.get('Id') == sheet_id:
            target = rel.get('Target', '')
            path   = f"xl/{target}" if not target.startswith('xl/') else target
            if path in zf.namelist():
                return zf.open(path)

    return None


def read_sheet(zf, sheet_name):
    """
    Read a named sheet from an xlsx zip file.
    Returns list of row dicts using first row as headers.
    """
    shared = get_shared_strings(zf)
    sheet  = get_sheet_by_name(zf, sheet_name)
    if sheet is None:
        return []

    tree  = ET.parse(sheet)
    cells = {}
    max_col = 0

    for row_el in tree.findall('.//ss:row', NS):
        row_idx = int(row_el.get('r', 0))
        for cell_el in row_el.findall('ss:c', NS):
            ref       = cell_el.get('r', '')
            cell_type = cell_el.get('t', '')
            val_el    = cell_el.find('ss:v', NS)

            if val_el is None or val_el.text is None:
                value = ''
            elif cell_type == 's':
                try:
                    value = shared[int(val_el.text)]
                except (IndexError, ValueError):
                    value = ''
            else:
                value = val_el.text

            col_letters = ''.join(c for c in ref if c.isalpha())
            col_idx = 0
            for ch in col_letters:
                col_idx = col_idx * 26 + (ord(ch.upper()) - ord('A') + 1)

            cells[(row_idx, col_idx)] = value
            max_col = max(max_col, col_idx)

    if not cells:
        return []

    max_row = max(r for r, c in cells)
    rows    = []
    for r in range(1, max_row + 1):
        rows.append([cells.get((r, c), '') for c in range(1, max_col + 1)])

    if not rows:
        return []

    # First row as headers
    headers = [str(h).strip() for h in rows[0]]
    result  = []
    for row in rows[1:]:
        result.append({
            headers[i]: row[i] if i < len(row) else ''
            for i in range(len(headers))
        })
    return result


# ── Schedule loader ────────────────────────────────────────────────────────────

def read_raw_rows(schedule_file, max_rows=10):
    """
    Read the first max_rows rows of the schedule sheet as raw lists.
    Used for diagnostic output when loading fails.
    """
    try:
        with zipfile.ZipFile(schedule_file, 'r') as zf:
            shared = get_shared_strings(zf)
            sheet  = get_sheet_by_name(zf, SCHEDULE_SHEET)
            if sheet is None:
                return [], f"Sheet '{SCHEDULE_SHEET}' not found"

            tree    = ET.parse(sheet)
            cells   = {}
            max_col = 0

            for row_el in tree.findall('.//ss:row', NS):
                row_idx = int(row_el.get('r', 0))
                if row_idx > max_rows:
                    continue
                for cell_el in row_el.findall('ss:c', NS):
                    ref       = cell_el.get('r', '')
                    cell_type = cell_el.get('t', '')
                    val_el    = cell_el.find('ss:v', NS)
                    if val_el is None or val_el.text is None:
                        value = ''
                    elif cell_type == 's':
                        try:
                            value = shared[int(val_el.text)]
                        except (IndexError, ValueError):
                            value = ''
                    else:
                        value = val_el.text
                    col_letters = ''.join(c for c in ref if c.isalpha())
                    col_idx = 0
                    for ch in col_letters:
                        col_idx = col_idx * 26 + (ord(ch.upper()) - ord('A') + 1)
                    cells[(row_idx, col_idx)] = value
                    max_col = max(max_col, col_idx)

            if not cells:
                return [], "No data found in sheet"

            actual_max_row = min(max_rows, max(r for r, c in cells))
            rows = []
            for r in range(1, actual_max_row + 1):
                rows.append([cells.get((r, c), '')
                             for c in range(1, max_col + 1)])
            return rows, None

    except Exception as e:
        return [], str(e)


def build_error_report(schedule_file, error_message, raw_rows):
    """Build an HTML error report showing what the script expected vs what it found."""
    timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

    rows_html = ''
    if raw_rows:
        # Header row
        rows_html += '<tr>' + ''.join(
            f'<th>Col {i+1}</th>' for i in range(len(raw_rows[0]))
        ) + '</tr>\n'
        for i, row in enumerate(raw_rows):
            style = ' style="background:#fffde7"' if i == 0 else ''
            cells_html = ''
            for c in row:
                if i == 0 and c:
                    cells_html += '<td><strong>' + str(c) + '</strong></td>'
                elif c:
                    cells_html += '<td>' + str(c) + '</td>'
                else:
                    cells_html += '<td><em style="color:#aaa">—</em></td>'
            rows_html += '<tr' + style + '>' + cells_html + '</tr>\n'

    return f"""<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<title>Schedule Read Error</title>
<style>
  body {{ font-family: Arial, sans-serif; margin: 30px; color: #222; line-height: 1.5; }}
  h1   {{ color: #dc3545; border-bottom: 3px solid #dc3545; padding-bottom: 8px; }}
  h2   {{ color: #34495e; border-bottom: 1px solid #ccc; padding-bottom: 4px; margin-top: 36px; }}
  table {{ border-collapse: collapse; width: 100%; margin: 12px 0; font-size: 0.88em; }}
  th {{ background: #2c3e50; color: white; padding: 6px 10px; text-align: left; }}
  td {{ padding: 5px 10px; border-bottom: 1px solid #eee; word-break: break-word; }}
  .error-box {{ background: #f8d7da; border: 1px solid #f5c6cb; border-radius: 6px;
                padding: 16px; margin: 16px 0; }}
  .expected-box {{ background: #d4edda; border: 1px solid #c3e6cb; border-radius: 6px;
                   padding: 16px; margin: 16px 0; }}
  code {{ background: #eee; padding: 2px 6px; border-radius: 3px; font-size: 0.95em; }}
  .footer {{ margin-top: 40px; padding: 12px 16px; background: #f8f9fa;
             border-left: 4px solid #dc3545; font-size: 0.9em; color: #555; }}
</style>
</head>
<body>
<h1>⚠ Error Reading Schedule File</h1>

<div class="error-box">
  <strong>Error:</strong> {error_message}<br>
  <strong>File:</strong> {schedule_file}
</div>

<h2>What the script expects</h2>
<div class="expected-box">
  <p>The schedule Excel file must have a sheet named exactly <code>Plan - NL</code>
  with the following columns in row 1:</p>
  <ul>
    <li><code>GameID</code> — unique identifier matching the game file name
        (e.g. <code>31August-Towers-v-Razorbacks</code>)</li>
    <li><code>Dato</code> — day number (e.g. <code>31</code>)</li>
    <li><code>Måned</code> — month name (e.g. <code>August</code>)</li>
    <li><code>Hjemme</code> — home team name</li>
    <li><code>Ude</code> — away team name</li>
    <li><code>R, U, H, L, B, F, S, C</code> — official initials per position</li>
  </ul>
  <p>The <code>GameID</code> column header must not be blank. Rows without a
  <code>GameID</code> value are skipped automatically.</p>
</div>

<h2>First {len(raw_rows)} rows found in sheet '{SCHEDULE_SHEET}'</h2>
<p>Row 1 (highlighted) should contain column headers. Check that
<code>GameID</code> appears as a header and that the column names above are present.</p>
<div style="overflow-x:auto">
<table>{rows_html}</table>
</div>

<div class="footer">
  Report generated: {timestamp}<br>
  Schedule file: {schedule_file}
</div>
</body>
</html>
"""


def load_schedule(schedule_file):
    """
    Load game schedule from Excel file.
    Returns dict: { game_id -> { date, home_team, away_team } }
    """
    schedule = {}
    with zipfile.ZipFile(schedule_file, 'r') as zf:
        rows = read_sheet(zf, SCHEDULE_SHEET)

    if not rows:
        raise ValueError(
            f"No data rows found. Check that the sheet is named '{SCHEDULE_SHEET}' "
            f"and has a header row with 'GameID'."
        )

    # Check for GameID column
    if 'GameID' not in rows[0]:
        found_headers = [k for k in rows[0].keys() if k.strip()]
        raise ValueError(
            f"No 'GameID' column found in header row. "
            f"Headers found: {found_headers}"
        )

    for row in rows:
        game_id = str(row.get('GameID', '') or '').strip()
        if not game_id or game_id.lower() == 'nan':
            continue

        dato  = str(row.get('Dato',  '') or '').strip()
        maned = str(row.get('Måned', '') or '').strip()
        date  = f"{dato} {maned}" if dato and maned else '—'

        schedule[game_id] = {
            'date':      date,
            'home_team': str(row.get('Hjemme', '') or '').strip(),
            'away_team': str(row.get('Ude',    '') or '').strip(),
        }

    return schedule


# ── HTML report ────────────────────────────────────────────────────────────────

def build_report(schedule, available_files):
    matched   = []
    missing   = []
    extra     = set(available_files)

    for game_id, info in schedule.items():
        if game_id in available_files:
            matched.append(game_id)
            extra.discard(game_id)
        else:
            missing.append(game_id)

    n_total   = len(schedule)
    n_matched = len(matched)
    n_missing = len(missing)
    n_extra   = len(extra)
    timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

    html = f"""<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<title>File Check Report</title>
<style>
  body {{ font-family: Arial, sans-serif; margin: 30px; color: #222; line-height: 1.5; }}
  h1   {{ color: #2c3e50; border-bottom: 3px solid #3498db; padding-bottom: 8px; }}
  h2   {{ color: #34495e; border-bottom: 1px solid #ccc; padding-bottom: 4px;
           margin-top: 36px; }}
  table {{ border-collapse: collapse; width: 100%; margin: 12px 0;
           font-size: 0.94em; }}
  th {{ background: #2c3e50; color: white; padding: 8px 12px; text-align: left; }}
  td {{ padding: 7px 12px; border-bottom: 1px solid #eee; }}
  tr:hover td {{ background: #f5f5f5; }}
  .grid {{ display: grid; grid-template-columns: repeat(auto-fit, minmax(160px,1fr));
           gap: 16px; margin: 20px 0; }}
  .card {{ background: #f8f9fa; border-radius: 8px; padding: 16px;
           border-left: 4px solid #3498db; }}
  .card .value {{ font-size: 2em; font-weight: bold; }}
  .card .label {{ color: #666; font-size: 0.9em; }}
  .ok   {{ color: #28a745; font-weight: bold; }}
  .bad  {{ color: #dc3545; font-weight: bold; }}
  .warn {{ color: #ffc107; font-weight: bold; }}
  .row-ok   {{ background: #d4edda !important; }}
  .row-miss {{ background: #f8d7da !important; }}
  .row-xtra {{ background: #fff3cd !important; }}
  .footer {{ margin-top: 40px; padding: 12px 16px; background: #f8f9fa;
             border-left: 4px solid #3498db; font-size: 0.9em; color: #555; }}
</style>
</head>
<body>
<h1>File Check Report</h1>

<div class="grid">
  <div class="card">
    <div class="value">{n_total}</div>
    <div class="label">Scheduled games</div>
  </div>
  <div class="card">
    <div class="value ok">{n_matched}</div>
    <div class="label">Files found</div>
  </div>
  <div class="card">
    <div class="value bad">{n_missing}</div>
    <div class="label">Files missing</div>
  </div>
  <div class="card">
    <div class="value warn">{n_extra}</div>
    <div class="label">Unmatched files</div>
  </div>
</div>

<h2>Schedule vs Files</h2>
<table>
  <tr>
    <th>Status</th><th>Game ID</th><th>Date</th>
    <th>Home</th><th>Away</th>
  </tr>
"""

    # Matched first, then missing
    for game_id in sorted(matched):
        info = schedule[game_id]
        html += (f'<tr class="row-ok">'
                 f'<td class="ok">✓ Found</td>'
                 f'<td>{game_id}</td>'
                 f'<td>{info["date"]}</td>'
                 f'<td>{info["home_team"]}</td>'
                 f'<td>{info["away_team"]}</td></tr>\n')

    for game_id in sorted(missing):
        info = schedule[game_id]
        html += (f'<tr class="row-miss">'
                 f'<td class="bad">✗ Missing</td>'
                 f'<td>{game_id}</td>'
                 f'<td>{info["date"]}</td>'
                 f'<td>{info["home_team"]}</td>'
                 f'<td>{info["away_team"]}</td></tr>\n')

    html += '</table>\n'

    if extra:
        html += '<h2>Unmatched Files</h2>'
        html += '<p>These files exist in data/ but do not match any scheduled game ID:</p>'
        html += '<table><tr><th>Filename</th></tr>\n'
        for name in sorted(extra):
            html += f'<tr class="row-xtra"><td>{name}.xlsx</td></tr>\n'
        html += '</table>\n'

    if missing:
        html += '<h2>How to Fix Missing Files</h2>'
        html += '<p>The file name must exactly match the Game ID shown above. For example:</p>'
        html += '<ul>'
        for game_id in sorted(missing)[:3]:
            html += f'<li><code>{game_id}.xlsx</code></li>'
        html += '</ul>'

    html += f"""
<div class="footer">
  Report generated: {timestamp}<br>
  Schedule file: {SCHEDULE_FOLDER}<br>
  Data folder: {DATA_FOLDER}
</div>
</body>
</html>
"""
    return html


# ── Main ──────────────────────────────────────────────────────────────────────

def main():
    print("01_check_files.py")
    print("=" * 50)

    # Validate folders
    for folder in [DATA_FOLDER, SCHEDULE_FOLDER]:
        if not folder.exists():
            print(f"ERROR: Folder '{folder}' not found")
            return

    OUTPUT_FOLDER.mkdir(exist_ok=True)

    # Find schedule file
    schedule_files = (list(SCHEDULE_FOLDER.glob("*.xlsx")) +
                      list(SCHEDULE_FOLDER.glob("*.xls")))
    if not schedule_files:
        print(f"ERROR: No Excel files found in '{SCHEDULE_FOLDER}'")
        return

    schedule_file = schedule_files[0]
    print(f"\nSchedule file : {schedule_file.name}")

    # Load schedule - write diagnostic error report if it fails
    try:
        schedule = load_schedule(schedule_file)
    except Exception as e:
        print(f"ERROR reading schedule: {e}")
        print(f"  Writing diagnostic report...")
        raw_rows, _  = read_raw_rows(schedule_file)
        error_html   = build_error_report(schedule_file, str(e), raw_rows)
        output_path  = OUTPUT_FOLDER / "troubleshooting_report.html"
        try:
            output_path.write_text(error_html, encoding='utf-8')
            print(f"  Diagnostic report: {output_path}")
        except Exception as write_err:
            print(f"  Could not write report: {write_err}")
        return
    print(f"  Loaded {len(schedule)} scheduled games")

    # Find available game files (by stem name)
    xlsx_files = set(f.stem for f in DATA_FOLDER.glob("*.xlsx"))
    csv_files  = set(f.stem for f in DATA_FOLDER.glob("*.csv"))
    available  = xlsx_files | csv_files
    print(f"  Found {len(available)} file(s) in data/")

    # Build and write report
    output_path = OUTPUT_FOLDER / "troubleshooting_report.html"
    try:
        html = build_report(schedule, available)
        output_path.write_text(html, encoding='utf-8')
    except Exception as e:
        print(f"ERROR writing report: {e}")
        return

    # Print summary
    matched = sum(1 for g in schedule if g in available)
    missing = sum(1 for g in schedule if g not in available)
    extra   = sum(1 for f in available if f not in schedule)

    print(f"\n{'─' * 50}")
    print(f"Found    : {matched}")
    print(f"Missing  : {missing}")
    print(f"Unmatched: {extra}")
    print(f"\nReport saved : {output_path}")


if __name__ == "__main__":
    main()
