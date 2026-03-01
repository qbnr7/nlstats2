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
    'R': 'Referee', 'U': 'Umpire', 'H': 'Head Linesman',
    'L': 'Line Judge', 'B': 'Back Judge', 'F': 'Field Judge',
    'S': 'Side Judge', 'C': 'Center Judge',
}

# Official sort order in game reports
POSITION_ORDER = ['R', 'U', 'H', 'L', 'S', 'F', 'B', 'C']

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
        return '—'
    if code in FOUL_NAMES:
        return f"{code} — {FOUL_NAMES[code]}"
    for prefix, parent_name in FOUL_GROUPS.items():
        if code.startswith(prefix):
            sub = code[len(prefix):]
            return f"{code} — {parent_name} ({sub})"
    return code


def foul_group(code):
    if not code:
        return '—'
    for prefix, parent_name in FOUL_GROUPS.items():
        if code.startswith(prefix):
            return parent_name
    if code in FOUL_NAMES:
        return f"{code} — {FOUL_NAMES[code]}"
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
"""

JS = """
<script>
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

def build_official_report(initials, data, games):
    name          = data['name']
    games_worked  = sorted(data['games'])
    calls_by_game = data['calls_by_game']

    all_grades        = []
    per_game_accuracy = {}
    positions_worked  = defaultdict(int)

    for game_id in games_worked:
        calls       = calls_by_game[game_id]
        game_grades = [c['grade'] for c in calls if c['grade']]
        all_grades.extend(game_grades)
        per_game_accuracy[game_id] = calc_accuracy(game_grades)
        for c in calls:
            if c['position']:
                positions_worked[c['position']] += 1

    overall_accuracy = calc_accuracy(all_grades)
    grade_counts     = defaultdict(int)
    for g in all_grades:
        grade_counts[g] += 1

    colour      = score_colour(overall_accuracy)
    acc_display = f"{overall_accuracy}%" if overall_accuracy is not None else "N/A"

    html = html_header(f"Official Report: {name}",
                       back_link="../combined_report.html")

    # Summary cards
    html += '<div class="summary-grid">'
    html += (f'<div class="summary-card"><div class="value" style="color:{colour}">'
             f'{acc_display}</div><div class="label">Overall Accuracy</div></div>')
    html += (f'<div class="summary-card"><div class="value">{len(games_worked)}</div>'
             f'<div class="label">Games Officiated</div></div>')
    html += (f'<div class="summary-card"><div class="value">{len(all_grades)}</div>'
             f'<div class="label">Graded Calls</div></div>')
    pos_str = ', '.join(
        f"{POSITION_NAMES.get(p, p)} ({n})"
        for p, n in sorted(positions_worked.items(), key=lambda x: pos_sort_key(x[0]))
    )
    html += (f'<div class="summary-card"><div class="value" style="font-size:1em">'
             f'{pos_str or "N/A"}</div><div class="label">Positions Worked</div></div>')
    html += '</div>'

    # Grade breakdown
    html += '<h2>Grade Breakdown</h2>'
    html += table_start(['Grade', 'Count', '% of graded calls'])
    total_scorable = sum(1 for g in all_grades if g in GRADE_SCORES)
    for g in ['C', 'M', 'N', 'I', 'G', 'W']:
        count = grade_counts.get(g, 0)
        pct   = (f"{round(count / total_scorable * 100, 1)}%"
                 if g in GRADE_SCORES and total_scorable > 0 else 'Excluded')
        html += f'<tr><td>{grade_badge(g)}</td><td>{count}</td><td>{pct}</td></tr>'
    html += table_end()

    # Performance by game
    html += '<h2>Performance by Game</h2>'
    html += table_start(['Game', 'Date', 'Position', 'Calls', 'Accuracy', 'Bar'])
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
        html += (f'<tr><td>{home} v {away}</td><td>{date}</td>'
                 f'<td>{pos_list}</td><td>{n_calls}</td>'
                 f'<td style="color:{col};font-weight:bold">{acc_str}</td>'
                 f'<td>{bar}</td></tr>')
    html += table_end()

    # Game by game breakdown
    html += '<h2>Game by Game Breakdown</h2>'
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
        html += (f'<h3>{home} v {away} — {date} '
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
            grade_cell = grade_badge(c['grade']) if c['grade'] else '—'
            html += (f"<tr><td>{c['qtr']}</td><td>{c['play']}</td>"
                     f"<td>{pos_name}</td><td>{foul_display(c['foul'])}</td>"
                     f"<td>{c['flag']}</td><td>{grade_cell}</td></tr>")
        html += table_end()
        html += '</div>'

    html += html_footer()
    return html


# ── Combined overview report ───────────────────────────────────────────────────

def build_combined_report(games, officials):
    html = html_header("NL Officiating — Season Overview")

    html += '<div class="toc">'
    html += '<strong>Contents</strong>'
    for anchor, label in [
        ('#game-summary',    'Game Summary'),
        ('#game-breakdown',  'Game by Game Breakdown'),
        ('#flag-breakdown',  'Flag Breakdown'),
        ('#foul-analysis',   'Penalty Analysis'),
        ('#officials-list',  'Officials List'),
        ('#season-ranking',  'Season Accuracy Ranking'),
        ('#pos-rankings',    'Position Rankings'),
    ]:
        html += f'<a href="{anchor}">{label}</a>'
    html += '</div>'

    # ── Game summary ──────────────────────────────────────────────────────────
    html += '<h2 id="game-summary">Game Summary</h2>'
    html += table_start(['Game', 'Date', 'Penalties', 'Crew Accuracy',
                         'Flag Breakdown'])
    for game_id in sorted(games):
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
        html += (f'<tr><td>{home} v {away}</td><td>{date}</td>'
                 f'<td>{n_penalties}</td>'
                 f'<td style="color:{col};font-weight:bold">{acc_str}</td>'
                 f'<td>{flag_str or "—"}</td></tr>')
    html += table_end()

    # ── Game by game breakdown ────────────────────────────────────────────────
    html += '<h2 id="game-breakdown">Game by Game Breakdown</h2>'
    for game_id in sorted(games):
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
        html += (f'<h3>{home} v {away} — {date} '
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
            off_name  = r['official_name'].strip() or initials or '—'
            pos_name  = POSITION_NAMES.get(position, position) if position else '—'
            grade_cell = grade_badge(grade) if grade else '—'
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
                 if total_flags > 0 else '—')
        html += (f'<tr><td><strong>{flag}</strong></td>'
                 f'<td>{count}</td><td>{pct}</td></tr>')
    html += table_end()

    # ── Penalty analysis ──────────────────────────────────────────────────────
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
                     if t_flag > 0 else '—')
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
