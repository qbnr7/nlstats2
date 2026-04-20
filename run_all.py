#!/usr/bin/env python3
"""
run_all.py
==========
Runs the full NL Stats pipeline in order:

    01_check_files.py      - check game files match schedule
    02_convert_to_csv.py   - convert game Excel files to CSV
    03_build_flat_file.py  - build flat_calls.csv
    04_generate_reports.py - generate HTML reports

The script stops if any step fails.

Usage:
    python3 run_all.py                 # run all steps
    python3 run_all.py --skip-convert  # skip step 02 (already converted)
    python3 run_all.py --skip-check    # skip step 01 (skip file check)
"""

import sys
import importlib.util
from pathlib import Path

# ── Configuration ─────────────────────────────────────────────────────────────

SCRIPTS = [
    ("01", "01_check_files.py",      "Check files"),
    ("02", "02_convert_to_csv.py",   "Convert to CSV"),
    ("03", "03_build_flat_file.py",  "Build flat file"),
    ("04", "04_generate_reports.py", "Generate reports"),
]

# ── Runner ────────────────────────────────────────────────────────────────────

def load_and_run(script_path):
    """Load a script as a module and call its main() function."""
    spec   = importlib.util.spec_from_file_location("module", script_path)
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    module.main()


def run_step(number, filename, label):
    script_path = Path(filename)
    if not script_path.exists():
        print(f"  ERROR: {filename} not found")
        return False
    try:
        load_and_run(script_path)
        return True
    except SystemExit:
        return True  # scripts that call sys.exit(0) are fine
    except Exception as e:
        print(f"\n  FAILED: {e}")
        return False


# ── Main ──────────────────────────────────────────────────────────────────────

def main():
    args         = sys.argv[1:]
    skip_convert = '--skip-convert' in args
    skip_check   = '--skip-check'   in args

    skip = set()
    if skip_convert:
        skip.add("02")
    if skip_check:
        skip.add("01")

    print("=" * 50)
    print("  NL Stats — Full Pipeline")
    print("=" * 50)
    if skip:
        skipping = ', '.join(f"step {s}" for s in sorted(skip))
        print(f"  Skipping: {skipping}")
    print()

    results = []
    for number, filename, label in SCRIPTS:
        if number in skip:
            print(f"[ SKIP ] Step {number} — {label}")
            results.append((label, 'skipped'))
            continue

        print(f"[ RUN  ] Step {number} — {label}")
        print("-" * 50)
        ok = run_step(number, filename, label)
        print("-" * 50)

        if ok:
            print(f"[ OK   ] Step {number} — {label}\n")
            results.append((label, 'ok'))
        else:
            print(f"[ FAIL ] Step {number} — {label}")
            print(f"\nPipeline stopped at step {number}.")
            results.append((label, 'failed'))
            break

    # Summary
    print()
    print("=" * 50)
    print("  Summary")
    print("=" * 50)
    for label, status in results:
        icon = {'ok': '✓', 'skipped': '–', 'failed': '✗'}.get(status, '?')
        print(f"  {icon}  {label}")

    all_ok = all(s in ('ok', 'skipped') for _, s in results)
    print()
    if all_ok:
        print("  Done. Open output/combined_report.html to view reports.")
    else:
        print("  Pipeline did not complete. Check errors above.")
    print("=" * 50)


if __name__ == "__main__":
    main()
