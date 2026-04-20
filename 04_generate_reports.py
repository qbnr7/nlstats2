#!/usr/bin/env python3
"""
04_generate_reports.py
======================
Reads flat_calls.csv and generates HTML reports with interactive tables.

Output:
    output/combined_report.html        - season overview for all audiences
    output/officials/{initials}.html   - individual report per official

Usage:
    python3 -W ignore 04_generate_reports.py
"""

import csv
from pathlib import Path
from collections import defaultdict

import re as _re

def format_game(game_id):
    """10Maj-89ers-v-Oaks  →  10 Maj -- 89ers vs Oaks"""
    m = _re.match(r'^(\d+)([A-Za-z]+)-(.+)-v-(.+)$', game_id)
    if not m:
        return game_id
    day, month, home, away = m.groups()
    return f"{day} {month} -- {home.replace('_', ' ')} vs {away.replace('_', ' ')}"

MONTH_ORDER = {
    'Januar': 1, 'February': 2, 'Marts': 3, 'April': 4,
    'Maj': 5, 'Juni': 6, 'Juli': 7, 'August': 8,
    'September': 9, 'Oktober': 10, 'November': 11, 'December': 12,
}

def game_sort_key(game_id):
    """Sort key for game IDs -- by month then day number."""
    m = _re.match(r'^(\d+)([A-Za-z]+)-', game_id)
    if not m:
        return (99, 99, game_id)
    day, month = m.groups()
    return (MONTH_ORDER.get(month, 50), int(day), game_id)



# ── Configuration ─────────────────────────────────────────────────────────────

INPUT_FILE    = Path("output/flat_calls.csv")
OUTPUT_FOLDER = Path("output")
OFFICIALS_DIR = OUTPUT_FOLDER / "officials"

MIN_GAMES_RANKING  = 3
MIN_GAMES_POSITION = 2

GRADE_SCORES   = {'C': 100, 'M': 75, 'N': 50, 'I': 0}
GRADE_EXCLUDED = {'G', 'W'}

GRADE_NAMES = {
    'C': 'Correct', 'M': 'Marginal', 'I': 'Incorrect',
    'N': 'No Call', 'G': 'Non-gradeable', 'W': 'Waived',
}

POSITION_NAMES = {
    'R': 'Referee', 'U': 'Umpire', 'D': 'Deep Judge',
    'L': 'Line Judge', 'B': 'Back Judge', 'F': 'Field Judge',
    'S': 'Side Judge', 'C': 'Center Judge',
}

# Official sort order in game reports
POSITION_ORDER = ['R', 'U', 'D', 'L', 'S', 'F', 'B', 'C']

FLAG_ORDER = ['CC', 'MC', 'IC', 'NC', 'NG', 'W']

# ── Foul code lookup ───────────────────────────────────────────────────────────

FOUL_NAMES = {
    'APS': 'Altering playing surface',
    'ATR': 'Assisting the runner',
    'BAT': 'Illegal batting',
    'DEH': 'Holding, defense',
    'DOD': 'Delay of game, defense',
    'DOF': 'Offside, defense',
    'DOG': 'Delay of game, offense',
    'DPI-AB':  'Pass interference, defense, arm bar',
    'DPI-CO':  'Pass interference, defense, cut off',
    'DPI-GR':  'Pass interference, defense, grab and restrict',
    'DPI-HT':  'Pass interference, defense, hook and turn',
    'DPI-NPB': 'Pass interference, defense, not playing the ball',
    'DPI-PTO': 'Pass interference, defense, playing through opponent',
    'DSH': 'Delay of game, start of half',
    'DSQ': 'Disqualification',
    'ENC': 'Encroachment (offense)',
    'FGT': 'Fighting',
    'FST': 'False start',
    'IBB': 'Illegal block in the back',
    'IBK': 'Illegal block during kick',
    'IBP': 'Illegal backward pass',
    'IDP': 'Ineligible downfield on pass',
    'IFD': 'Illegal formation, defense (3-on-1)',
    'IFH': 'Illegal forward handing',
    'IFK': 'Illegal free kick formation',
    'IFP': 'Illegal forward pass',
    'IKB': 'Illegally kicking ball',
    'ILF': 'Illegal formation',
    'ILM': 'Illegal motion',
    'ING': 'Intentional grounding',
    'IPN': 'Improper numbering',
    'IPR': 'Illegal procedure',
    'ISH': 'Illegal shift',
    'ISP': 'Illegal snap',
    'ITP': 'Illegal touching of a forward pass',
    'IUH': 'Illegal use of hands',
    'IWK': 'Illegal wedge on kickoff',
    'KCI': 'Kick-catch interference',
    'KIK': 'Illegal kick',
    'KOB': 'Free kick out of bounds',
    'OBK': 'Out of bounds during kick',
    'OFH-GR': 'Holding, offense, grab and restrict',
    'OFH-HR': 'Holding, offense, hook and restrict',
    'OFH-TD': 'Holding, offense, takedown',
    'OFK': 'Offside, free kick',
    'OPI-BK': 'Pass interference, offense, blocking',
    'OPI-DT': 'Pass interference, offense, driving through',
    'OPI-PK': 'Pass interference, offense, pick',
    'OPI-PO': 'Pass interference, offense, pushing off',
    'PF-BBW': 'Personal foul, blocking below the waist',
    'PF-BOB': 'Personal foul, blocking out of bounds',
    'PF-BSB': 'Personal foul, blind-side block',
    'PF-BTH': 'Personal foul, blow to the head',
    'PF-CHB': 'Personal foul, chop block',
    'PF-CLP': 'Personal foul, clipping',
    'PF-FMM': 'Personal foul, face mask',
    'PF-HCT': 'Personal foul, horse collar tackle',
    'PF-HDR': 'Personal foul, hit on defenseless receiver',
    'PF-HTF': 'Personal foul, hands to the face',
    'PF-HUR': 'Personal foul, hurdling',
    'PF-ICS': 'Personal foul, illegal contact with snapper',
    'PF-LEA': 'Personal foul, leaping',
    'PF-LHP': 'Personal foul, late hit/piling on',
    'PF-LTO': 'Personal foul, late hit out of bounds',
    'PF-OTH': 'Personal foul, other',
    'PF-RFK': 'Personal foul, roughing free kicker',
    'PF-RTH': 'Personal foul, roughing the holder',
    'PF-RTK': 'Personal foul, roughing the kicker',
    'PF-RTP': 'Personal foul, roughing the passer',
    'PF-SKE': 'Personal foul, striking/kneeing/elbowing',
    'PF-TGB': 'Personal foul, targeting (both rules)',
    'PF-TGC': 'Personal foul, targeting (crown of helmet)',
    'PF-TGD': 'Personal foul, targeting (defenceless player)',
    'PF-TRP': 'Personal foul, tripping',
    'PF-UNR': 'Personal foul, unnecessary roughness',
    'RNH': 'Running into the holder',
    'RNK': 'Running into the kicker',
    'SLI': 'Sideline interference, 5 yards',
    'SLM': 'Sideline interference, 15 yards',
    'SLW': 'Sideline interference, warning',
    'SUB': 'Illegal substitution',
    'UC-2PN': 'Unsportsmanlike conduct, two players with same number',
    'UC-ABL': 'Unsportsmanlike conduct, abusive language',
    'UC-BCH': 'Unsportsmanlike conduct, bench',
    'UC-DBS': 'Unsportsmanlike conduct, dead ball shoving',
    'UC-DEA': 'Unsportsmanlike conduct, delayed/excessive act',
    'UC-FCO': 'Unsportsmanlike conduct, forcibly contacting an official',
    'UC-RHT': 'Unsportsmanlike conduct, removal of helmet',
    'UC-SBR': 'Unsportsmanlike conduct, simulating being roughed',
    'UC-STB': 'Unsportsmanlike conduct, spiking/throwing ball',
    'UC-TAU': 'Unsportsmanlike conduct, taunting/baiting',
    'UC-UNS': 'Unsportsmanlike conduct, other',
    'UFA': 'Unfair acts',
    'UFT': 'Unfair tactics',
}

FOUL_GROUPS = {
    'PF-':  'Personal foul',
    'OFH-': 'Holding, offense',
    'DPI-': 'Pass interference, defense',
    'OPI-': 'Pass interference, offense',
    'UC-':  'Unsportsmanlike conduct',
    'DOF-': 'Offside, defense',
}


def foul_display(code):
    if not code:
        return '--'
    if code in FOUL_NAMES:
        return f"{code} -- {FOUL_NAMES[code]}"
    for prefix, parent_name in FOUL_GROUPS.items():
        if code.startswith(prefix):
            sub = code[len(prefix):]
            return f"{code} -- {parent_name} ({sub})"
    return code


def foul_group(code):
    if not code:
        return '--'
    for prefix, parent_name in FOUL_GROUPS.items():
        if code.startswith(prefix):
            return parent_name
    if code in FOUL_NAMES:
        return f"{code} -- {FOUL_NAMES[code]}"
    return code


# ── Scoring ────────────────────────────────────────────────────────────────────

def calc_accuracy(grades):
    scorable = [GRADE_SCORES[g] for g in grades if g in GRADE_SCORES]
    if not scorable:
        return None
    return round(sum(scorable) / len(scorable), 1)


def score_colour(score):
    if score is None:
        return '#888888'
    if score >= 90:
        return '#28a745'
    if score >= 75:
        return '#ffc107'
    if score >= 60:
        return '#fd7e14'
    return '#dc3545'


def pos_sort_key(pos):
    try:
        return POSITION_ORDER.index(pos)
    except ValueError:
        return len(POSITION_ORDER)


# ── Data loading ───────────────────────────────────────────────────────────────

def load_data():
    games     = {}
    officials = {}

    with open(INPUT_FILE, newline='', encoding='utf-8') as f:
        reader = csv.DictReader(f)
        for row in reader:
            game_id  = row['game_id'].strip()
            initials = row['official_initials'].strip()
            name     = row['official_name'].strip()
            grade    = row['grade_code'].strip().upper()
            position = row['position'].strip().upper()
            foul     = row['foul_code'].strip()
            flag     = row['flag'].strip().upper()
            play     = row['play_number'].strip()
            qtr      = row['qtr'].strip()

            if game_id not in games:
                games[game_id] = {
                    'date':      row['date'].strip(),
                    'home_team': row['home_team'].strip(),
                    'away_team': row['away_team'].strip(),
                    'rows':      [],
                }
            games[game_id]['rows'].append(row)

            if not initials:
                continue

            if initials not in officials:
                officials[initials] = {
                    'name':          name or initials,
                    'games':         set(),
                    'calls_by_game': defaultdict(list),
                }

            officials[initials]['games'].add(game_id)
            officials[initials]['calls_by_game'][game_id].append({
                'play': play, 'qtr': qtr, 'foul': foul,
                'flag': flag, 'position': position, 'grade': grade,
            })

    return games, officials


# ── HTML helpers ───────────────────────────────────────────────────────────────

CSS = """
body { font-family: Arial, sans-serif; margin: 30px; color: #222; line-height: 1.5; }
h1 { color: #2c3e50; border-bottom: 3px solid #3498db; padding-bottom: 8px; }
h2 { color: #34495e; border-bottom: 1px solid #ccc; padding-bottom: 4px; margin-top: 40px; }
h3 { color: #555; margin-top: 28px; }
h4 { color: #666; margin-top: 16px; margin-bottom: 4px; }
.sortable-table { border-collapse: collapse; width: 100%; margin: 4px 0 12px 0;
                  font-size: 0.93em; }
.sortable-table th { background: #2c3e50; color: white; padding: 8px 12px;
                     text-align: left; cursor: pointer; user-select: none;
                     white-space: nowrap; }
.sortable-table th:hover { background: #34495e; }
.sortable-table th.sort-asc::after  { content: ' ▲'; font-size: 0.75em; }
.sortable-table th.sort-desc::after { content: ' ▼'; font-size: 0.75em; }
.sortable-table td { padding: 7px 12px; border-bottom: 1px solid #eee; }
.sortable-table tr:hover td { background: #f5f5f5; }
.filter-input { width: 100%; padding: 6px 10px; margin: 6px 0 2px 0;
                border: 1px solid #ccc; border-radius: 4px;
                font-size: 0.9em; box-sizing: border-box; }
.badge { display: inline-block; padding: 2px 8px; border-radius: 4px;
         font-size: 0.85em; font-weight: bold; color: white; }
.tag-C { background: #28a745; }
.tag-M { background: #ffc107; color: #333; }
.tag-I { background: #dc3545; }
.tag-N { background: #fd7e14; }
.tag-G { background: #888; }
.tag-W { background: #6c757d; }
.summary-grid { display: grid; grid-template-columns: repeat(auto-fit, minmax(180px,1fr));
                gap: 16px; margin: 20px 0; }
.summary-card { background: #f8f9fa; border-radius: 8px; padding: 16px;
                border-left: 4px solid #3498db; }
.summary-card .value { font-size: 1.9em; font-weight: bold; }
.summary-card .label { color: #666; font-size: 0.9em; }
.game-section { background: #fafafa; border: 1px solid #ddd; border-radius: 6px;
                padding: 16px; margin: 20px 0; }
a { color: #3498db; }
.trend-bar { display: inline-block; height: 14px; border-radius: 3px;
             vertical-align: middle; }
.toc { background: #f0f4f8; border-radius: 6px; padding: 16px; margin: 20px 0; }
.toc a { display: block; margin: 4px 0; }
.explainer { background: #eaf3fb; border-left: 4px solid #3498db; border-radius: 6px;
             padding: 18px 22px; margin: 20px 0; font-size: 0.95em; line-height: 1.7; }
.explainer h3 { color: #2c6e9e; margin: 0 0 10px 0; font-size: 1.05em; }
.explainer p  { margin: 6px 0; }
.explainer table { border-collapse: collapse; margin: 10px 0; font-size: 0.92em; }
.explainer table td, .explainer table th { padding: 4px 14px; border: 1px solid #c3daf0; }
.explainer table th { background: #c3daf0; font-weight: bold; }
.explainer .flag-row  { background: #fef9e7; }
.explainer .grade-row { background: #eafaf1; }
.summary-card .sublabel { color: #888; font-size: 0.8em; margin-top: 2px; }
/* Sticky side nav */
.sidenav { position: fixed; right: 0; top: 50%; transform: translateY(-50%);
           background: #2c3e50; border-radius: 8px 0 0 8px; padding: 10px 0;
           z-index: 1000; min-width: 44px; box-shadow: -2px 0 10px rgba(0,0,0,.2); }
.sidenav a { display: flex; align-items: center; padding: 8px 14px 8px 12px;
             color: #ccc; text-decoration: none; font-size: 0.78em;
             white-space: nowrap; overflow: hidden; transition: all .2s; }
.sidenav a .nav-icon { font-size: 1.1em; min-width: 20px; }
.sidenav a .nav-label { max-width: 0; overflow: hidden; transition: max-width .25s ease;
                         padding-left: 0; }
.sidenav:hover a .nav-label { max-width: 200px; padding-left: 8px; }
.sidenav a:hover { background: #3498db; color: white; }
.sidenav a.active { background: #3498db; color: white; }
.sidenav .nav-divider { height: 1px; background: #455; margin: 4px 10px; }
/* Foul breakdown table */
.pct-bar-cell { white-space: nowrap; }
.pct-bar-cell .mini-bar { display:inline-block; height:10px; border-radius:3px;
                           vertical-align:middle; margin-right:4px; }
.cell-cc  { background: #d4edda; }
.cell-ic  { background: #f8d7da; }
.cell-mc  { background: #fff3cd; }
.cell-nc  { background: #fde8d3; }
.cell-ng  { background: #e9ecef; }
.cell-w   { background: #e9ecef; }
.foul-filter-bar { display:flex; gap:10px; flex-wrap:wrap; margin:10px 0 4px 0;
                   align-items:center; font-size:0.9em; }
.foul-filter-bar select, .foul-filter-bar input
  { padding:5px 8px; border:1px solid #ccc; border-radius:4px; font-size:0.9em; }
"""

JS = """
<script>
// ── Sticky nav scroll spy ─────────────────────────────────────────────────
document.addEventListener('DOMContentLoaded', function() {
  var navLinks = document.querySelectorAll('.sidenav a[href^="#"]');
  if (!navLinks.length) return;
  var sections = Array.from(navLinks).map(function(a) {
    return document.querySelector(a.getAttribute('href'));
  }).filter(Boolean);
  function onScroll() {
    var scrollY = window.scrollY + 120;
    var current = sections[0];
    sections.forEach(function(s) { if (s.offsetTop <= scrollY) current = s; });
    navLinks.forEach(function(a) {
      a.classList.toggle('active', a.getAttribute('href') === '#' + current.id);
    });
  }
  window.addEventListener('scroll', onScroll, { passive: true });
  onScroll();
});
// ── Sortable tables ────────────────────────────────────────────────────────
function getCellValue(row, colIndex) {
  return row.cells[colIndex] ? row.cells[colIndex].innerText.trim() : '';
}

function compareValues(a, b) {
  // Try numeric comparison first
  const numA = parseFloat(a.replace('%', ''));
  const numB = parseFloat(b.replace('%', ''));
  if (!isNaN(numA) && !isNaN(numB)) return numA - numB;
  return a.localeCompare(b, undefined, { numeric: true });
}

function makeSortable(table) {
  const headers = table.querySelectorAll('th');
  let currentCol = -1;
  let ascending  = true;

  headers.forEach((th, colIndex) => {
    th.addEventListener('click', () => {
      if (currentCol === colIndex) {
        ascending = !ascending;
      } else {
        ascending  = true;
        currentCol = colIndex;
      }
      headers.forEach(h => h.classList.remove('sort-asc', 'sort-desc'));
      th.classList.add(ascending ? 'sort-asc' : 'sort-desc');

      const tbody  = table.querySelector('tbody') || table;
      const rows   = Array.from(tbody.querySelectorAll('tr'));
      const sorted = rows.sort((a, b) => {
        const valA = getCellValue(a, colIndex);
        const valB = getCellValue(b, colIndex);
        return ascending ? compareValues(valA, valB) : compareValues(valB, valA);
      });
      sorted.forEach(r => tbody.appendChild(r));
    });
  });
}

// ── Filter inputs ──────────────────────────────────────────────────────────
function addFilter(table) {
  const wrapper = document.createElement('div');
  const input   = document.createElement('input');
  input.type        = 'text';
  input.placeholder = 'Filter table...';
  input.className   = 'filter-input';
  input.addEventListener('input', () => {
    const term  = input.value.toLowerCase();
    const tbody = table.querySelector('tbody') || table;
    tbody.querySelectorAll('tr').forEach(row => {
      const text = row.innerText.toLowerCase();
      row.style.display = text.includes(term) ? '' : 'none';
    });
  });
  wrapper.appendChild(input);
  table.parentNode.insertBefore(wrapper, table);
}

// ── Apply to all tables ────────────────────────────────────────────────────
document.addEventListener('DOMContentLoaded', () => {
  document.querySelectorAll('table.sortable-table').forEach(table => {
    makeSortable(table);
    addFilter(table);
  });
});
</script>
"""


def html_header(title, back_link=None):
    back = (f'<p><a href="{back_link}">← Back to overview</a></p>'
            if back_link else '')
    return (f'<!DOCTYPE html>\n<html lang="en">\n<head>\n'
            f'<meta charset="UTF-8">\n<title>{title}</title>\n'
            f'<style>{CSS}</style>\n</head>\n<body>\n{JS}\n{back}\n'
            f'<h1>{title}</h1>\n')


def html_footer():
    return "</body></html>\n"


def table_start(headers):
    """Open a sortable table with given header labels."""
    ths  = ''.join(f'<th>{h}</th>' for h in headers)
    return f'<table class="sortable-table"><thead><tr>{ths}</tr></thead><tbody>\n'


def table_end():
    return '</tbody></table>\n'


def grade_badge(grade):
    name = GRADE_NAMES.get(grade, grade)
    return f'<span class="badge tag-{grade}">{name}</span>'


def grade_breakdown_cells(grades):
    counts = defaultdict(int)
    for g in grades:
        counts[g] += 1
    return ''.join(f'<td>{counts.get(g, 0)}</td>'
                   for g in ['C', 'M', 'I', 'N', 'G', 'W'])


# ── Individual official report ─────────────────────────────────────────────────


def explainer_box():
    return """
<div class="explainer">
  <h3>&#x1F4CB; Understanding Flags and Grades</h3>
  <p>Each penalty call produces two separate pieces of information tracked independently:</p>
  <table>
    <tr><th>Concept</th><th>What it measures</th><th>Who it applies to</th><th>Codes used</th></tr>
    <tr class="flag-row">
      <td><strong>Flag</strong>&nbsp;(crew grade)</td>
      <td>Did the crew as a whole handle the situation correctly?</td>
      <td>The entire officiating crew on the play</td>
      <td>CC &nbsp;MC &nbsp;IC &nbsp;NC &nbsp;NG &nbsp;W</td>
    </tr>
    <tr class="grade-row">
      <td><strong>Grade</strong>&nbsp;(individual)</td>
      <td>Did this specific official perform correctly -- positioning, mechanics, judgement?</td>
      <td>Each individual official involved in the call</td>
      <td>C &nbsp;M &nbsp;N &nbsp;I &nbsp;G &nbsp;W</td>
    </tr>
  </table>
  <p><strong>Accuracy percentages are based on individual grades only.</strong>
     G (non-gradeable) and W (waived) are excluded from the calculation.</p>
  <table>
    <tr><th>Grade</th><th>Meaning</th><th>Score</th></tr>
    <tr><td><span class="badge tag-C">C</span></td><td>Correct</td><td>100</td></tr>
    <tr><td><span class="badge tag-M">M</span></td><td>Marginal</td><td>75</td></tr>
    <tr><td><span class="badge tag-N">N</span></td><td>No-call (should have been flagged)</td><td>50</td></tr>
    <tr><td><span class="badge tag-I">I</span></td><td>Incorrect</td><td>0</td></tr>
    <tr><td><span class="badge tag-G">G</span></td><td>Non-gradeable -- excluded</td><td>--</td></tr>
    <tr><td><span class="badge tag-W">W</span></td><td>Waived (flag picked up) -- excluded</td><td>--</td></tr>
  </table>
</div>
"""

def build_official_report(initials, data, games):
    name          = data['name']
    games_worked  = sorted(data['games'], key=game_sort_key)
    calls_by_game = data['calls_by_game']

    all_grades        = []
    per_game_accuracy = {}
    positions_worked  = defaultdict(int)

    pos_games = defaultdict(set)   # position -> set of game_ids
    for game_id in games_worked:
        calls       = calls_by_game[game_id]
        game_grades = [c['grade'] for c in calls if c['grade']]
        all_grades.extend(game_grades)
        per_game_accuracy[game_id] = calc_accuracy(game_grades)
        for c in calls:
            if c['position']:
                positions_worked[c['position']] += 1
                pos_games[c['position']].add(game_id)

    overall_accuracy = calc_accuracy(all_grades)
    grade_counts     = defaultdict(int)
    for g in all_grades:
        grade_counts[g] += 1

    colour      = score_colour(overall_accuracy)
    acc_display = f"{overall_accuracy}%" if overall_accuracy is not None else "N/A"

    html = html_header(f"Official Report: {name}",
                       back_link="../combined_report.html")

    html += sidenav_html([
        ('summary',       '📊', 'Summary'),
        ('grade-breakdown','🎯', 'Grade Breakdown'),
        ('perf-by-game',  '📅', 'By Game'),
        ('game-breakdown','🔍', 'Game Detail'),
    ])

    html += explainer_box()

    # Summary cards
    html += '<div id="summary">'
    n_flags_total = len(all_grades)
    n_graded      = sum(1 for g in all_grades if g in GRADE_SCORES)
    # Build positions card: one line per position -- name, games, flags
    pos_lines = ''.join(
        f'<div style="margin:3px 0">'
        f'<strong>{POSITION_NAMES.get(p, p)}</strong>'
        f', Games {len(pos_games[p])}'
        f', Flags {n}</div>'
        for p, n in sorted(positions_worked.items(), key=lambda x: pos_sort_key(x[0]))
    ) or '<div>N/A</div>'
    html += '<div class="summary-grid">'
    html += (f'<div class="summary-card"><div class="value" style="color:{colour}">'
             f'{acc_display}</div>'
             f'<div class="label">Overall Accuracy</div>'
             f'<div class="sublabel">Based on {n_graded} graded calls (C / M / N / I)</div></div>')
    html += (f'<div class="summary-card"><div class="value">{len(games_worked)}</div>'
             f'<div class="label">Games Officiated</div></div>')
    html += (f'<div class="summary-card"><div class="value">{n_flags_total}</div>'
             f'<div class="label">Flags Thrown</div>'
             f'<div class="sublabel">All calls including non-gradeable (G) and waived (W)</div></div>')
    html += (f'<div class="summary-card"><div class="value">{n_graded}</div>'
             f'<div class="label">Graded Calls</div>'
             f'<div class="sublabel">C / M / N / I only -- used for accuracy</div></div>')
    html += (f'<div class="summary-card" style="min-width:220px">'
             f'<div class="label" style="margin-bottom:6px;font-weight:bold">Positions Worked</div>'
             f'<div style="font-size:0.93em">{pos_lines}</div></div>')
    html += '</div>'

    html += '</div>'
    # Grade breakdown
    html += '<h2 id="grade-breakdown">Grade Breakdown</h2>'
    html += table_start(['Grade', 'Count', '% of graded calls'])
    total_scorable = sum(1 for g in all_grades if g in GRADE_SCORES)
    for g in ['C', 'M', 'N', 'I', 'G', 'W']:
        count = grade_counts.get(g, 0)
        pct   = (f"{round(count / total_scorable * 100, 1)}%"
                 if g in GRADE_SCORES and total_scorable > 0 else 'Excluded')
        html += f'<tr><td>{grade_badge(g)}</td><td>{count}</td><td>{pct}</td></tr>'
    html += table_end()

    # Performance by game
    html += '<h2 id="perf-by-game">Performance by Game</h2>'
    html += table_start(['Game', 'Position', 'Calls', 'Accuracy', 'Bar'])
    for game_id in games_worked:
        calls     = calls_by_game[game_id]
        info      = games.get(game_id, {})
        home      = info.get('home_team', '')
        away      = info.get('away_team', '')
        date      = info.get('date', '')
        pos_list  = ', '.join(sorted(
            set(c['position'] for c in calls if c['position']),
            key=pos_sort_key
        ))
        n_calls   = len([c for c in calls if c['grade'] in GRADE_SCORES])
        acc       = per_game_accuracy[game_id]
        acc_str   = f"{acc}%" if acc is not None else "N/A"
        col       = score_colour(acc)
        bar_width = int(acc) if acc is not None else 0
        bar       = (f'<div class="trend-bar" style="width:{bar_width}px;'
                     f'background:{col}"></div>')
        html += (f'<tr><td>{format_game(game_id)}</td>'
                 f'<td>{pos_list}</td><td>{n_calls}</td>'
                 f'<td style="color:{col};font-weight:bold">{acc_str}</td>'
                 f'<td>{bar}</td></tr>')
    html += table_end()

    # Game by game breakdown
    html += '<h2 id="game-breakdown">Game by Game Breakdown</h2>'
    for game_id in games_worked:
        calls   = calls_by_game[game_id]
        info    = games.get(game_id, {})
        home    = info.get('home_team', '')
        away    = info.get('away_team', '')
        date    = info.get('date', '')
        acc     = per_game_accuracy[game_id]
        col     = score_colour(acc)
        acc_str = f"{acc}%" if acc is not None else "N/A"

        html += '<div class="game-section">'
        html += (f'<h3>{format_game(game_id)} '
                 f'<span style="color:{col}">({acc_str})</span></h3>')
        html += table_start(['Qtr', 'Play', 'Position', 'Foul', 'Flag', 'Grade'])
        # Sort calls by position order, then play number
        sorted_calls = sorted(
            calls,
            key=lambda c: (pos_sort_key(c['position']),
                           int(c['play']) if c['play'].isdigit() else 0)
        )
        for c in sorted_calls:
            pos_name   = POSITION_NAMES.get(c['position'], c['position'])
            grade_cell = grade_badge(c['grade']) if c['grade'] else '--'
            html += (f"<tr><td>{c['qtr']}</td><td>{c['play']}</td>"
                     f"<td>{pos_name}</td><td>{foul_display(c['foul'])}</td>"
                     f"<td>{c['flag']}</td><td>{grade_cell}</td></tr>")
        html += table_end()
        html += '</div>'

    html += html_footer()
    return html


# ── Combined overview report ───────────────────────────────────────────────────


def sidenav_html(sections):
    """sections = list of (anchor, icon, label)"""
    html  = '<nav class="sidenav">'
    for i, (anchor, icon, label) in enumerate(sections):
        if i and i % 4 == 0:
            html += '<div class="nav-divider"></div>'
        html += (f'<a href="#{anchor}">'
                 f'<span class="nav-icon">{icon}</span>'
                 f'<span class="nav-label">{label}</span>'
                 f'</a>')
    html += '</nav>'
    return html





def build_foul_table(games):
    """One row per foul type with crew flag % columns and individual accuracy."""
    from collections import defaultdict

    GRADE_SCORES_LOCAL = {'C': 100, 'M': 75, 'N': 50, 'I': 0}
    FLAG_COLS = ['CC', 'MC', 'IC', 'NC', 'NG', 'W']

    foul_flags  = defaultdict(lambda: defaultdict(int))
    foul_grades = defaultdict(lambda: defaultdict(int))

    seen_plays = set()
    for info in games.values():
        for r in info['rows']:
            foul  = r['foul_code'].strip()
            flag  = r['flag'].strip().upper()
            grade = r['grade_code'].strip().upper()
            if not foul:
                continue
            play_key = (r['game_id'], r['play_number'], foul)
            if play_key not in seen_plays:
                seen_plays.add(play_key)
                if flag:
                    foul_flags[foul][flag] += 1
            if grade:
                foul_grades[foul][grade] += 1

    def ind_acc(foul):
        denom  = sum(foul_grades[foul][g] for g in GRADE_SCORES_LOCAL)
        if not denom:
            return None
        total_score = sum(GRADE_SCORES_LOCAL[g] * foul_grades[foul][g]
                          for g in GRADE_SCORES_LOCAL)
        return round(total_score / denom, 1)

    def flag_pct(foul, flag):
        total = sum(foul_flags[foul].values())
        return round(foul_flags[foul][flag] / total * 100, 1) if total else 0.0

    def pct_bar(value):
        w = round(value * 0.55)
        bar = '<span style="display:inline-block;height:10px;border-radius:3px;'
        bar += 'background:#aaa;vertical-align:middle;margin-right:5px;'
        bar += 'width:' + str(w) + 'px"></span>'
        return bar + str(value) + '%'

    def acc_bar(value):
        if value is None:
            return '<span style="color:#aaa">N/A</span>'
        col = score_colour(value)
        w   = round(value * 0.55)
        bar = '<span style="display:inline-block;height:12px;border-radius:3px;'
        bar += 'background:' + col + ';vertical-align:middle;margin-right:6px;'
        bar += 'width:' + str(w) + 'px"></span>'
        return bar + '<strong style="color:' + col + '">' + str(value) + '%</strong>'

    fouls = sorted(foul_flags.keys(), key=lambda f: -sum(foul_flags[f].values()))
    cats  = sorted(set(foul_group(f) for f in fouls))

    # Colour backgrounds per column
    col_bg = {
        'CC': '#d4edda', 'MC': '#fff3cd', 'IC': '#f8d7da',
        'NC': '#fde8d3', 'NG': '#e9ecef', 'W':  '#e9ecef',
    }

    lines = []
    lines.append('<h2 id="foul-table">Foul Breakdown by Call Type</h2>')
    lines.append('<p>One row per foul type. '
                 '<strong>Crew flag %</strong> columns show how often the crew\'s collective '
                 'call received each grade. '
                 '<strong>Ind. Accuracy</strong> is the weighted average of individual '
                 'official grades (C=100, M=75, N=50, I=0) -- G and W excluded. '
                 'Click any column header to sort. Use the filters below to narrow the list.</p>')

    # Filter bar
    lines.append('<div class="foul-filter-bar">')
    lines.append('<label>Category: <select id="foulCatFilter" onchange="filterFoulTable()">'
                 '<option value="">All</option>')
    for cat in cats:
        lines.append('<option value="' + cat + '">' + cat + '</option>')
    lines.append('</select></label>')
    lines.append('<label>Search foul: <input type="text" id="foulSearch" '
                 'placeholder="e.g. OFH, PF..." oninput="filterFoulTable()" '
                 'style="width:160px"></label>')
    lines.append('<label>Min flags: <input type="number" id="foulMinFlags" value="1" min="1" '
                 'oninput="filterFoulTable()" style="width:60px"></label>')
    lines.append('</div>')

    # Table header
    lines.append('<table class="sortable-table" id="foulBreakdownTable">')
    lines.append('<thead><tr>')
    lines.append('<th>Foul</th>')
    lines.append('<th>Category</th>')
    lines.append('<th>Flags</th>')
    for flag in FLAG_COLS:
        title_map = {
            'CC': 'Crew: Correct Call', 'MC': 'Crew: Marginal Call',
            'IC': 'Crew: Incorrect Call', 'NC': 'Crew: No-Call',
            'NG': 'Crew: Non-gradeable', 'W': 'Crew: Waived',
        }
        lines.append('<th title="' + title_map[flag] + '">' + flag + ' %</th>')
    lines.append('<th title="Individual official accuracy (C/M/N/I grades)">Ind. Accuracy</th>')
    lines.append('</tr></thead>')
    lines.append('<tbody>')

    for foul in fouls:
        total = sum(foul_flags[foul].values())
        cat   = foul_group(foul)
        acc   = ind_acc(foul)
        lines.append('<tr data-cat="' + cat + '" data-flags="' + str(total) + '">')
        lines.append('<td><strong>' + foul_display(foul) + '</strong></td>')
        lines.append('<td>' + cat + '</td>')
        lines.append('<td>' + str(total) + '</td>')
        for flag in FLAG_COLS:
            p = flag_pct(foul, flag)
            bg = col_bg.get(flag, '#fff')
            cell = '<td style="background:' + bg + '">' + pct_bar(p) + '</td>'
            lines.append(cell)
        lines.append('<td>' + acc_bar(acc) + '</td>')
        lines.append('</tr>')

    lines.append('</tbody></table>')

    # Filter JS
    lines.append('<script>')
    lines.append('function filterFoulTable() {')
    lines.append('  var cat      = document.getElementById("foulCatFilter").value.toLowerCase();')
    lines.append('  var search   = document.getElementById("foulSearch").value.toLowerCase();')
    lines.append('  var minFlags = parseInt(document.getElementById("foulMinFlags").value) || 1;')
    lines.append('  var rows     = document.querySelectorAll("#foulBreakdownTable tbody tr");')
    lines.append('  rows.forEach(function(row) {')
    lines.append('    var rowCat   = (row.dataset.cat   || "").toLowerCase();')
    lines.append('    var rowFlags = parseInt(row.dataset.flags) || 0;')
    lines.append('    var rowText  = row.textContent.toLowerCase();')
    lines.append('    var show = (!cat    || rowCat.indexOf(cat)    >= 0)')
    lines.append('            && (!search || rowText.indexOf(search) >= 0)')
    lines.append('            && rowFlags >= minFlags;')
    lines.append('    row.style.display = show ? "" : "none";')
    lines.append('  });')
    lines.append('}')
    lines.append('</script>')

    return '\n'.join(lines) + '\n'



def build_combined_report(games, officials):
    html = html_header("NL Officiating -- Season Overview")

    NAV_SECTIONS = [
        ('game-summary',   '📅', 'Game Summary'),
        ('game-breakdown', '🔍', 'Game by Game'),
        ('flag-breakdown', '🚩', 'Flag Breakdown'),
        ('foul-table',     '📊', 'Foul Breakdown'),
        ('foul-analysis',  '📋', 'Penalty Analysis'),
        ('officials-list', '👤', 'Officials List'),
        ('season-ranking', '🏆', 'Season Ranking'),
        ('pos-rankings',   '📍', 'Position Rankings'),
    ]
    html += sidenav_html(NAV_SECTIONS)

    html += '<div class="toc">'
    html += '<strong>Contents</strong>'
    for anchor, icon, label in NAV_SECTIONS:
        html += f'<a href="#{anchor}">{icon} {label}</a>'
    html += '</div>'

    html += explainer_box()

    # ── Game summary ──────────────────────────────────────────────────────────
    html += '<h2 id="game-summary">Game Summary</h2>'
    html += table_start(['Game', 'Penalties', 'Crew Accuracy', 'Flag Breakdown'])
    for game_id in sorted(games, key=game_sort_key):
        info = games[game_id]
        rows = info['rows']
        home = info['home_team']
        away = info['away_team']
        date = info['date']

        seen        = set()
        flag_counts = defaultdict(int)
        all_grades  = []

        for r in rows:
            foul = r['foul_code'].strip()
            if not foul:
                continue
            key = (r['play_number'], r['foul_code'])
            if key not in seen:
                seen.add(key)
                flag = r['flag'].strip().upper()
                if flag:
                    flag_counts[flag] += 1
            grade = r['grade_code'].strip().upper()
            if grade:
                all_grades.append(grade)

        n_penalties = len(seen)
        acc         = calc_accuracy(all_grades)
        acc_str     = f"{acc}%" if acc is not None else "N/A"
        col         = score_colour(acc)
        flag_str    = ', '.join(
            f"{k}: {v}" for k, v in sorted(flag_counts.items())
        )
        html += (f'<tr><td>{format_game(game_id)}</td>'
                 f'<td>{n_penalties}</td>'
                 f'<td style="color:{col};font-weight:bold">{acc_str}</td>'
                 f'<td>{flag_str or "--"}</td></tr>')
    html += table_end()

    # ── Game by game breakdown ────────────────────────────────────────────────
    html += '<h2 id="game-breakdown">Game by Game Breakdown</h2>'
    for game_id in sorted(games, key=game_sort_key):
        info = games[game_id]
        home = info['home_team']
        away = info['away_team']
        date = info['date']
        rows = info['rows']

        all_grades = [r['grade_code'].strip().upper()
                      for r in rows if r['grade_code'].strip()]
        acc     = calc_accuracy(all_grades)
        acc_str = f"{acc}%" if acc is not None else "N/A"
        col     = score_colour(acc)

        html += '<div class="game-section">'
        html += (f'<h3>{format_game(game_id)} '
                 f'<span style="color:{col}">Crew accuracy: {acc_str}</span></h3>')

        # Officials table - sorted by position order
        game_officials = {}
        for initials, data in officials.items():
            if game_id in data['calls_by_game']:
                calls     = data['calls_by_game'][game_id]
                positions = sorted(
                    set(c['position'] for c in calls if c['position']),
                    key=pos_sort_key
                )
                grades = [c['grade'] for c in calls if c['grade']]
                game_officials[initials] = {
                    'name':      data['name'],
                    'positions': positions,
                    'grades':    grades,
                    'n_fouls':   len(calls),
                }

        if game_officials:
            html += '<h4>Officials</h4>'
            html += table_start(['Official', 'Position', 'Fouls', 'Accuracy',
                                 'C', 'M', 'I', 'N', 'G', 'W'])
            # Sort officials by their primary (first) position
            sorted_officials = sorted(
                game_officials.items(),
                key=lambda x: pos_sort_key(
                    x[1]['positions'][0] if x[1]['positions'] else 'Z'
                )
            )
            for initials, d in sorted_officials:
                pos_str = ', '.join(
                    POSITION_NAMES.get(p, p) for p in d['positions']
                )
                acc     = calc_accuracy(d['grades'])
                acc_str = f"{acc}%" if acc is not None else "N/A"
                col     = score_colour(acc)
                link    = f"officials/{initials}.html"
                html += (f'<tr><td><a href="{link}">{d["name"]}</a></td>'
                         f'<td>{pos_str}</td><td>{d["n_fouls"]}</td>'
                         f'<td style="color:{col};font-weight:bold">{acc_str}</td>'
                         f'{grade_breakdown_cells(d["grades"])}</tr>')
            html += table_end()

        # Penalties table
        html += '<h4>Penalties</h4>'
        html += table_start(['Qtr', 'Play', 'Foul', 'Flag',
                             'Official', 'Position', 'Grade'])
        for r in rows:
            foul = r['foul_code'].strip()
            if not foul:
                continue
            play      = r['play_number'].strip()
            qtr       = r['qtr'].strip()
            flag      = r['flag'].strip().upper()
            initials  = r['official_initials'].strip()
            position  = r['position'].strip().upper()
            grade     = r['grade_code'].strip().upper()
            off_name  = r['official_name'].strip() or initials or '--'
            pos_name  = POSITION_NAMES.get(position, position) if position else '--'
            grade_cell = grade_badge(grade) if grade else '--'
            html += (f'<tr><td>{qtr}</td><td>{play}</td>'
                     f'<td>{foul_display(foul)}</td><td>{flag}</td>'
                     f'<td>{off_name}</td><td>{pos_name}</td>'
                     f'<td>{grade_cell}</td></tr>')
        html += table_end()
        html += '</div>'

    # ── Flag breakdown ────────────────────────────────────────────────────────
    html += '<h2 id="flag-breakdown">Flag Breakdown (All Games)</h2>'
    all_flags = defaultdict(int)
    for info in games.values():
        seen = set()
        for r in info['rows']:
            foul = r['foul_code'].strip()
            if not foul:
                continue
            key = (r['play_number'], r['foul_code'])
            if key not in seen:
                seen.add(key)
                flag = r['flag'].strip().upper()
                if flag:
                    all_flags[flag] += 1

    total_flags = sum(all_flags.values())
    html += table_start(['Flag', 'Count', '% of penalties'])
    for flag in FLAG_ORDER:
        count = all_flags.get(flag, 0)
        pct   = (f"{round(count / total_flags * 100, 1)}%"
                 if total_flags > 0 else '--')
        html += (f'<tr><td><strong>{flag}</strong></td>'
                 f'<td>{count}</td><td>{pct}</td></tr>')
    html += table_end()

    # ── Penalty analysis ──────────────────────────────────────────────────────
    html += build_foul_table(games)

    html += '<h2 id="foul-analysis">Penalty Analysis</h2>'
    html += '<p>Subcategories (PF, OFH, DPI, OPI, UC, DOF) are grouped together.</p>'

    group_data = defaultdict(lambda: {
        'subcodes':    defaultdict(int),
        'flag_counts': defaultdict(int),
        'total':       0,
    })

    for info in games.values():
        seen = set()
        for r in info['rows']:
            foul = r['foul_code'].strip()
            if not foul:
                continue
            key = (r['play_number'], r['foul_code'], r['game_id'])
            if key not in seen:
                seen.add(key)
                group = foul_group(foul)
                flag  = r['flag'].strip().upper()
                group_data[group]['subcodes'][foul] += 1
                group_data[group]['total'] += 1
                if flag:
                    group_data[group]['flag_counts'][flag] += 1

    for group, gdata in sorted(
        group_data.items(), key=lambda x: -x[1]['total']
    ):
        fc     = gdata['flag_counts']
        total  = gdata['total']
        t_flag = sum(fc.values())

        html += f'<h3>{group} ({total} total)</h3>'
        html += table_start(['Flag', 'Count', '%'])
        for flag in FLAG_ORDER:
            count = fc.get(flag, 0)
            pct   = (f"{round(count / t_flag * 100, 1)}%"
                     if t_flag > 0 else '--')
            html += (f'<tr><td><strong>{flag}</strong></td>'
                     f'<td>{count}</td><td>{pct}</td></tr>')
        html += table_end()

        if len(gdata['subcodes']) > 1:
            html += table_start(['Subcode', 'Count'])
            for subcode, count in sorted(
                gdata['subcodes'].items(), key=lambda x: -x[1]
            ):
                html += (f'<tr><td>{foul_display(subcode)}</td>'
                         f'<td>{count}</td></tr>')
            html += table_end()

    # ── Officials list ────────────────────────────────────────────────────────
    html += '<h2 id="officials-list">Officials List</h2>'
    html += table_start(['Official', 'Games', 'Positions', 'Accuracy',
                         'C', 'M', 'I', 'N', 'G', 'W'])
    for initials, data in sorted(
        officials.items(), key=lambda x: x[1]['name']
    ):
        all_grades = [
            c['grade']
            for calls in data['calls_by_game'].values()
            for c in calls if c['grade']
        ]
        positions = sorted(set(
            c['position']
            for calls in data['calls_by_game'].values()
            for c in calls if c['position']
        ), key=pos_sort_key)
        pos_str = ', '.join(POSITION_NAMES.get(p, p) for p in positions)
        acc     = calc_accuracy(all_grades)
        acc_str = f"{acc}%" if acc is not None else "N/A"
        col     = score_colour(acc)
        link    = f"officials/{initials}.html"
        html += (f'<tr><td><a href="{link}">{data["name"]}</a></td>'
                 f'<td>{len(data["games"])}</td><td>{pos_str}</td>'
                 f'<td style="color:{col};font-weight:bold">{acc_str}</td>'
                 f'{grade_breakdown_cells(all_grades)}</tr>')
    html += table_end()

    # ── Season ranking ────────────────────────────────────────────────────────
    html += '<h2 id="season-ranking">Season Accuracy Ranking</h2>'
    html += (f'<p>Minimum {MIN_GAMES_RANKING} games to qualify. '
             f'G and W excluded from score.</p>')
    html += table_start(['Rank', 'Official', 'Games', 'Graded Calls', 'Accuracy'])

    ranking = []
    for initials, data in officials.items():
        if len(data['games']) < MIN_GAMES_RANKING:
            continue
        all_grades = [
            c['grade']
            for calls in data['calls_by_game'].values()
            for c in calls if c['grade']
        ]
        acc        = calc_accuracy(all_grades)
        n_scorable = sum(1 for g in all_grades if g in GRADE_SCORES)
        ranking.append((initials, data['name'], len(data['games']),
                        n_scorable, acc))

    ranking.sort(key=lambda x: (x[4] is None, -(x[4] or 0)))
    for rank, (initials, name, n_games, n_calls, acc) in enumerate(ranking, 1):
        acc_str = f"{acc}%" if acc is not None else "N/A"
        col     = score_colour(acc)
        link    = f"officials/{initials}.html"
        html += (f'<tr><td>{rank}</td>'
                 f'<td><a href="{link}">{name}</a></td>'
                 f'<td>{n_games}</td><td>{n_calls}</td>'
                 f'<td style="color:{col};font-weight:bold">{acc_str}</td></tr>')
    html += table_end()

    # ── Position rankings ─────────────────────────────────────────────────────
    html += '<h2 id="pos-rankings">Position Rankings</h2>'
    html += (f'<p>Minimum {MIN_GAMES_POSITION} games at that position to qualify.</p>')

    for pos_code in POSITION_ORDER:
        pos_name    = POSITION_NAMES[pos_code]
        pos_ranking = []

        for initials, data in officials.items():
            pos_games = [
                gid for gid, calls in data['calls_by_game'].items()
                if any(c['position'] == pos_code for c in calls)
            ]
            if len(pos_games) < MIN_GAMES_POSITION:
                continue
            pos_grades = [
                c['grade']
                for gid in pos_games
                for c in data['calls_by_game'][gid]
                if c['position'] == pos_code and c['grade']
            ]
            acc        = calc_accuracy(pos_grades)
            n_scorable = sum(1 for g in pos_grades if g in GRADE_SCORES)
            pos_ranking.append((initials, data['name'], len(pos_games),
                                n_scorable, acc))

        if not pos_ranking:
            continue

        pos_ranking.sort(key=lambda x: (x[4] is None, -(x[4] or 0)))
        html += f'<h3>{pos_name} ({pos_code})</h3>'
        html += table_start(['Rank', 'Official', 'Games at Position',
                             'Graded Calls', 'Accuracy'])
        for rank, (initials, name, n_games, n_calls, acc) in \
                enumerate(pos_ranking, 1):
            acc_str = f"{acc}%" if acc is not None else "N/A"
            col     = score_colour(acc)
            link    = f"officials/{initials}.html"
            html += (f'<tr><td>{rank}</td>'
                     f'<td><a href="{link}">{name}</a></td>'
                     f'<td>{n_games}</td><td>{n_calls}</td>'
                     f'<td style="color:{col};font-weight:bold">{acc_str}</td></tr>')
        html += table_end()

    html += html_footer()
    return html


# ── Main ──────────────────────────────────────────────────────────────────────

def main():
    print("04_generate_reports.py")
    print("=" * 50)

    if not INPUT_FILE.exists():
        print(f"ERROR: {INPUT_FILE} not found. Run 03_build_flat_file.py first.")
        return

    OUTPUT_FOLDER.mkdir(exist_ok=True)
    OFFICIALS_DIR.mkdir(exist_ok=True)

    print(f"\nLoading data from {INPUT_FILE}...")
    games, officials = load_data()
    print(f"  Games loaded    : {len(games)}")
    print(f"  Officials found : {len(officials)}")

    print(f"\nGenerating individual reports...")
    for initials, data in sorted(officials.items()):
        html = build_official_report(initials, data, games)
        path = OFFICIALS_DIR / f"{initials}.html"
        path.write_text(html, encoding='utf-8')
        print(f"  {data['name']:30s} -> officials/{initials}.html")

    print(f"\nGenerating combined report...")
    html = build_combined_report(games, officials)
    path = OUTPUT_FOLDER / "combined_report.html"
    path.write_text(html, encoding='utf-8')
    print(f"  -> combined_report.html")

    print(f"\n{'─' * 50}")
    print(f"Done. Open output/combined_report.html to view.")


if __name__ == "__main__":
    main()
