#!/usr/bin/env python3
"""
02_convert_to_csv.py
====================
Converts game Excel files to CSV format.

This is a one-time conversion step needed because the game Excel files
have a formatting issue that prevents standard Excel readers from opening
them. Converting to CSV removes all formatting and produces clean data
files that work reliably on any system.

The schedule file (in nlplan/) does NOT need converting as it reads fine.

Input:
    data/    - folder containing game Excel files (.xlsx)

Output:
    data/    - CSV files written alongside the Excel files
               e.g. 31August-Towers-v-Razorbacks.xlsx
                 -> 31August-Towers-v-Razorbacks.csv

Usage:
    python 02_convert_to_csv.py

After running this script, the Excel files can be kept or deleted.
Script 03_build_flat_file.py will read the CSV files.
"""

import csv
import zipfile
import xml.etree.ElementTree as ET
from pathlib import Path

# ── Configuration ─────────────────────────────────────────────────────────────

DATA_FOLDER = Path("data")

# ── xlsx to CSV converter (no dependencies beyond stdlib) ─────────────────────

NS = {
    'ss': 'http://schemas.openxmlformats.org/spreadsheetml/2006/main',
    'r':  'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
}


def get_shared_strings(zf):
    """Extract shared strings table from xlsx zip."""
    if 'xl/sharedStrings.xml' not in zf.namelist():
        return []
    tree = ET.parse(zf.open('xl/sharedStrings.xml'))
    strings = []
    for si in tree.findall('.//ss:si', NS):
        # Concatenate all text nodes within the si element
        text = ''.join(t.text or '' for t in si.findall('.//ss:t', NS))
        strings.append(text)
    return strings


def get_sheet_xml(zf):
    """Get the first sheet XML from xlsx zip."""
    # Find sheet path from workbook relationships
    rels_xml = zf.read('xl/_rels/workbook.xml.rels').decode()
    rels_tree = ET.fromstring(rels_xml)
    
    for rel in rels_tree:
        target = rel.get('Target', '')
        if 'sheet' in target.lower():
            path = f"xl/{target}" if not target.startswith('xl/') else target
            if path in zf.namelist():
                return zf.open(path)
    
    # Fallback - try direct path
    for name in zf.namelist():
        if 'worksheets/sheet' in name:
            return zf.open(name)
    
    raise ValueError("No sheet found in xlsx file")


def xlsx_to_rows(file_path):
    """
    Read an xlsx file using only stdlib (zipfile + xml).
    Returns a list of rows, each row is a list of cell values.
    """
    with zipfile.ZipFile(file_path, 'r') as zf:
        shared_strings = get_shared_strings(zf)
        sheet_file = get_sheet_xml(zf)
        tree = ET.parse(sheet_file)

    rows = []
    max_col = 0

    # Build a dict of {(row, col): value} 
    cells = {}

    for row_el in tree.findall('.//ss:row', NS):
        row_idx = int(row_el.get('r', 0))
        for cell_el in row_el.findall('ss:c', NS):
            ref = cell_el.get('r', '')  # e.g. A1, B2
            cell_type = cell_el.get('t', '')
            val_el = cell_el.find('ss:v', NS)

            if val_el is None or val_el.text is None:
                value = ''
            elif cell_type == 's':
                # Shared string
                try:
                    value = shared_strings[int(val_el.text)]
                except (IndexError, ValueError):
                    value = ''
            else:
                value = val_el.text

            # Parse column from ref (e.g. 'AB3' -> col index)
            col_letters = ''.join(c for c in ref if c.isalpha())
            col_idx = 0
            for ch in col_letters:
                col_idx = col_idx * 26 + (ord(ch.upper()) - ord('A') + 1)

            cells[(row_idx, col_idx)] = value
            max_col = max(max_col, col_idx)

    if not cells:
        return []

    max_row = max(r for r, c in cells)

    for r in range(1, max_row + 1):
        row = [cells.get((r, c), '') for c in range(1, max_col + 1)]
        rows.append(row)

    return rows


def convert_file(xlsx_path):
    """Convert a single xlsx file to CSV in the same folder."""
    csv_path = xlsx_path.with_suffix('.csv')

    try:
        rows = xlsx_to_rows(xlsx_path)
        if not rows:
            print(f"  Warning: no data found in {xlsx_path.name}")
            return False

        with open(csv_path, 'w', newline='', encoding='utf-8') as f:
            writer = csv.writer(f)
            writer.writerows(rows)

        print(f"  Converted : {xlsx_path.name} -> {csv_path.name} "
              f"({len(rows)} rows)")
        return True

    except Exception as e:
        print(f"  ERROR converting {xlsx_path.name}: {e}")
        return False


# ── Main ──────────────────────────────────────────────────────────────────────

def main():
    print("02_convert_to_csv.py")
    print("=" * 50)

    if not DATA_FOLDER.exists():
        print(f"ERROR: Data folder '{DATA_FOLDER}' not found")
        return

    xlsx_files = sorted(DATA_FOLDER.glob("*.xlsx"))
    if not xlsx_files:
        print(f"No .xlsx files found in '{DATA_FOLDER}'")
        return

    print(f"\nFound {len(xlsx_files)} Excel file(s) to convert\n")

    success = 0
    failed  = 0

    for xlsx_path in xlsx_files:
        if convert_file(xlsx_path):
            success += 1
        else:
            failed += 1

    print(f"\n{'-' * 50}")
    print(f"Converted : {success}")
    if failed:
        print(f"Failed    : {failed}")
    print(f"\nCSV files are ready in '{DATA_FOLDER}'")
    print("You can now run 03_build_flat_file.py")


if __name__ == "__main__":
    main()
