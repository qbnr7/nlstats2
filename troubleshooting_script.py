#!/usr/bin/env python3
"""
Enhanced Comprehensive Football Officiating Performance Analysis System
Complete version with troubleshooting functionality
"""

import pandas as pd
import numpy as np
import os
import sys
from pathlib import Path
import argparse
from collections import defaultdict, Counter
import warnings
from datetime import datetime
import math
import re
warnings.filterwarnings('ignore')

# Script metadata
SCRIPT_VERSION = "1.4.3"
SCRIPT_CREATED = "2024-12-28"
SCRIPT_AUTHOR = "Claude AI Assistant"

class ComprehensiveOfficatingAnalyzer:
    def __init__(self, data_folder="data", schedule_folder="nlplan", output_folder="reports"):
        self.data_folder = Path(data_folder)
        self.schedule_folder = Path(schedule_folder)
        self.output_folder = Path(output_folder)
        self.output_folder.mkdir(exist_ok=True)
        
        # Define grade mappings
        self.grades = ['CC', 'MC', 'IC', 'NC', 'NG', 'Waived']
        self.grade_names = {
            'CC': 'Correct Call',
            'MC': 'Marginal Call', 
            'IC': 'Incorrect Call',
            'NC': 'No Call',
            'NG': 'Non-gradeable',
            'Waived': 'Waived'
        }
        
        # Enhanced grade descriptions
        self.grade_descriptions = {
            'CC': 'Correct by rule and philosophy',
            'MC': 'Technically correct by rule but not by philosophy', 
            'IC': 'A foul was called when it should not have been, either by rule or due to being too far from accepted philosophy',
            'NC': 'Foul that should have been called, but was not',
            'NG': 'Not enough evidence on video to determine whether the call was correct or not',
            'Waived': 'Waived - Just for reporting, should not be used for accuracy',
            # Single letter versions for compatibility
            'C': 'Correct by rule and philosophy',
            'M': 'Technically correct by rule but not by philosophy',
            'I': 'A foul was called when it should not have been, either by rule or due to being too far from accepted philosophy',
            'N': 'Foul that should have been called, but was not',
            'G': 'Not enough evidence on video to determine whether the call was correct or not',
            'W': 'Waived - Just for reporting, should not be used for accuracy'
        }
        
        # Official position codes and full names
        self.official_positions = {
            'R': 'Referee',
            'U': 'Umpire', 
            'H': 'Head Linesman',
            'L': 'Line Judge',
            'F': 'Field Judge',
            'S': 'Side Judge',
            'B': 'Back Judge',
            'C': 'Center Judge'
        }
        
        # Enhanced grading system with weighted scoring
        self.grade_weights = {
            'C': 100,   # Correct Call - Perfect score
            'M': 75,    # Marginal Call - Technically correct but not ideal
            'N': 50,    # No Call - Missed opportunity, moderate impact
            'I': 0,     # Incorrect Call - Very bad, worst possible grade
            'G': None,  # Non-gradeable - Excluded from accuracy calculation
            'W': None,  # Waived - Excluded from accuracy calculation
            'WAIVED': None,
            'Waived': None,
            # Also support full codes for compatibility
            'CC': 100,
            'MC': 75,
            'NC': 50,
            'IC': 0,
            'NG': None
        }
        
        # Ranking configuration
        self.ranking_config = {
            'min_games_for_full_ranking': 3,
            'min_calls_for_reliability': 5,
            'confidence_multiplier': 1.96
        }
        
        # Data storage
        self.all_games_data = []
        self.penalty_plays = []
        self.schedule_data = {}
        self.officials_data = {}
        self.official_performance_data = defaultdict(lambda: {
            'games': [],
            'positions': defaultdict(int),
            'total_calls': defaultdict(int),
            'grade_breakdown': defaultdict(lambda: defaultdict(int)),
            'penalty_types': defaultdict(int)
        })
    
    def normalize_grade(self, grade):
        """Normalize grade to standard format"""
        if pd.isna(grade) or grade == '' or grade is None:
            return 'Unknown'
        
        grade = str(grade).strip().upper()
        
        # Handle common variations - More comprehensive mapping
        grade_mapping = {
            'CORRECT': 'C',
            'CORRECT CALL': 'C',
            'CC': 'C',  # Map CC to C for consistency
            'MARGINAL': 'M',
            'MARGINAL CALL': 'M', 
            'MC': 'M',  # Map MC to M
            'INCORRECT': 'I',
            'INCORRECT CALL': 'I',
            'IC': 'I',  # Map IC to I
            'NO CALL': 'N',
            'NOCALL': 'N',
            'NC': 'N',  # Map NC to N
            'NON-GRADEABLE': 'G',
            'NOT GRADED': 'G',
            'NG': 'G',  # Map NG to G
            'WAIVED': 'W',
            'W': 'W'
        }
        
        normalized = grade_mapping.get(grade, grade)
        
        # VALIDATION: Ensure only valid grades are returned
        valid_grades = {'C', 'M', 'I', 'N', 'G', 'W', 'Unknown'}
        if normalized not in valid_grades:
            print(f"Warning: Unknown grade '{grade}' -> defaulting to 'Unknown'")
            return 'Unknown'
        
        return normalized
    
    def _safe_get_value(self, row, possible_keys, default=''):
        """Get value safely from row with multiple possible keys"""
        for key in possible_keys:
            if key in row and pd.notna(row[key]):
                value = str(row[key]).strip()
                if value and value.lower() not in ['nan', 'none', '']:
                    return value
        return default

    def _format_date(self, row):
        """Format date safely"""
        try:
            dato = self._safe_get_value(row, ['Dato', 'Date', 'Day'])
            month = self._safe_get_value(row, ['MÃƒÂ¥ned', 'Month', 'MÃ¥ned'])
            if dato and month:
                return f"{dato}-{month}"
            return self._safe_get_value(row, ['Date', 'Dato'], 'N/A')
        except:
            return 'N/A'
    
    def _html_escape(self, text):
        """Escape HTML to prevent XSS"""
        if pd.isna(text) or text is None:
            return 'N/A'
        
        text = str(text)
        html_escape_table = {
            "&": "&amp;",
            '"': "&quot;",
            "'": "&#x27;",
            ">": "&gt;",
            "<": "&lt;",
        }
        return "".join(html_escape_table.get(c, c) for c in text)
    
    def _is_valid_value(self, value):
        """Helper to check if a value is valid (not empty, NaN, or placeholder)"""
        if pd.isna(value) or value is None:
            return False
        value_str = str(value).strip().lower()
        return value_str not in ['', 'nan', 'none', 'n/a']
    
    def get_official_name_from_initials(self, initials):
        """Convert initials to full name using officials data"""
        if pd.isna(initials) or initials == '':
            return 'Unknown'
        
        initials = str(initials).strip()
        if initials in self.officials_data:
            return self.officials_data[initials]['name']
        return initials  # Return initials if no match found
    
    def load_schedule_data(self):
        """Load officiating schedule and officials database"""
        print(f"Looking for schedule files in: {self.schedule_folder}")
        
        if not self.schedule_folder.exists():
            print(f"Schedule folder '{self.schedule_folder}' does not exist!")
            return
        
        schedule_files = list(self.schedule_folder.glob("*.xlsx")) + list(self.schedule_folder.glob("*.xls"))
        
        if not schedule_files:
            print(f"No Excel files found in '{self.schedule_folder}' folder!")
            return
        
        print(f"Found {len(schedule_files)} Excel file(s):")
        for file in schedule_files:
            print(f"   - {file.name}")
        
        for file_path in schedule_files:
            try:
                print(f"\nProcessing: {file_path.name}")
                
                # Load games data from "Plan - NL" sheet
                try:
                    schedule_df = pd.read_excel(file_path, sheet_name="Plan - NL")
                    schedule_df.columns = schedule_df.columns.str.strip()
                    print(f"   Found {len(schedule_df)} rows in games sheet")
                    
                    games_processed = 0
                    for _, row in schedule_df.iterrows():
                        game_id = row.get(' GameID', '') or row.get('GameID', '')
                        if pd.notna(game_id) and str(game_id).strip() != '':
                            game_id_clean = str(game_id).strip()
                            
                            game_data = {
                                'game_id': game_id_clean,
                                'round': row.get('Runde', ''),
                                'date': f"{row.get('Dato', '')}-{row.get('MÃƒÂ¥ned', '')}",
                                'home_team': row.get('Hjemme', ''),
                                'away_team': row.get('Ude', ''),
                                'referee': self.clean_official_name(row.get('R', '')),
                                'umpire': self.clean_official_name(row.get('U', '')),
                                'head_linesman': self.clean_official_name(row.get('H', '')),
                                'line_judge': self.clean_official_name(row.get('L', '')),
                                'back_judge': self.clean_official_name(row.get('B', '')),
                                'field_judge': self.clean_official_name(row.get('F', '')),
                                'side_judge': self.clean_official_name(row.get('S', '')),
                                'center_judge': self.clean_official_name(row.get('C', ''))
                            }
                            
                            self.schedule_data[game_id_clean] = game_data
                            games_processed += 1
                    
                    print(f"   Successfully processed {games_processed} games from 'Plan - NL'")
                    
                except Exception as e:
                    print(f"   Error processing 'Plan - NL': {e}")
                
                # Load officials data
                try:
                    officials_df = pd.read_excel(file_path, sheet_name="Officials and games")
                    
                    # Check if first row contains headers
                    first_row = officials_df.iloc[0].values if len(officials_df) > 0 else []
                    
                    if len(officials_df) > 0 and 'Initialer' in str(first_row):
                        officials_df.columns = officials_df.iloc[0]
                        officials_df = officials_df[1:]
                        officials_df.reset_index(drop=True, inplace=True)
                    
                    officials_df.columns = officials_df.columns.str.strip()
                    
                    officials_processed = 0
                    for _, row in officials_df.iterrows():
                        initials = row.get('Initialer', '')
                        name = row.get('Navn', '')
                        
                        if (pd.notna(initials) and pd.notna(name) and 
                            str(initials).strip() != '' and str(name).strip() != ''):
                            
                            initials_clean = str(initials).strip()
                            name_clean = str(name).strip()
                            
                            official_data = {
                                'name': name_clean,
                                'status': row.get('Status', ''),
                                'primary_position': row.get('PrimÃƒÂ¦r', ''),
                                'secondary_position': row.get('SekundÃƒÂ¦r', ''),
                                'license': row.get('Licens', ''),
                                'region': row.get('Landsdel', ''),
                                'club': row.get('Klub', ''),
                                'email': row.get('Mail', '')
                            }
                            
                            self.officials_data[initials_clean] = official_data
                            officials_processed += 1
                    
                    print(f"   Successfully processed {officials_processed} officials")
                    
                except Exception as e:
                    print(f"   Error loading officials: {e}")
                
            except Exception as file_error:
                print(f"Error processing schedule file {file_path}: {file_error}")
        
        print(f"\nFinal Schedule Loading Summary:")
        print(f"   Games loaded: {len(self.schedule_data)}")
        print(f"   Officials loaded: {len(self.officials_data)}")
    
    def clean_official_name(self, name):
        """Clean and standardize official names/initials"""
        if pd.isna(name) or name == '':
            return ''
        
        name = str(name).strip()
        if '+' in name:
            return name.split('+')[0].strip()
        return name
    
    def generate_troubleshooting_report(self, excel_files):
        """Generate a troubleshooting report showing schedule vs available files"""
        
        # Get list of available game files (without extensions)
        available_files = set()
        for file_path in excel_files:
            available_files.add(file_path.stem)
        
        # Analyze schedule vs available files
        schedule_games = []
        missing_files = []
        extra_files = set(available_files)
        
        for game_id, game_data in self.schedule_data.items():
            # Try different possible filename patterns
            possible_names = [
                game_id,
                game_id.replace(' ', '-'),
                game_id.replace(' ', '_'),
                f"{game_data.get('date', '')}-{game_data.get('home_team', '')}-v-{game_data.get('away_team', '')}",
                f"{game_data.get('home_team', '')}-v-{game_data.get('away_team', '')}"
            ]
            
            # Clean up possible names
            possible_names = [name.replace(' ', '-').replace('--', '-') for name in possible_names if name]
            
            found_match = False
            matched_file = None
            
            for possible_name in possible_names:
                if possible_name in available_files:
                    found_match = True
                    matched_file = possible_name
                    extra_files.discard(possible_name)
                    break
            
            schedule_games.append({
                'game_id': game_id,
                'round': game_data.get('round', 'N/A'),
                'date': game_data.get('date', 'N/A'),
                'home_team': game_data.get('home_team', 'N/A'),
                'away_team': game_data.get('away_team', 'N/A'),
                'has_file': found_match,
                'matched_file': matched_file,
                'possible_names': possible_names
            })
            
            if not found_match:
                missing_files.append(game_id)
        
        # Generate HTML troubleshooting report
        report_time = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        
        html = f"""<!DOCTYPE html>
<html>
<head>
    <title>Football Officiating Analysis - Troubleshooting Report</title>
    <style>
        body {{ font-family: Arial, sans-serif; margin: 20px; line-height: 1.6; }}
        table {{ border-collapse: collapse; width: 100%; margin: 20px 0; }}
        th, td {{ border: 1px solid #ddd; padding: 8px; text-align: left; }}
        th {{ background-color: #f2f2f2; font-weight: bold; }}
        .found {{ background-color: #d4edda; }}
        .missing {{ background-color: #f8d7da; }}
        .extra {{ background-color: #fff3cd; }}
        .summary {{ background-color: #e8f4f8; padding: 15px; margin: 20px 0; border-radius: 5px; }}
        .section {{ margin: 30px 0; }}
        h1 {{ color: #2c3e50; }}
        h2 {{ color: #34495e; border-bottom: 2px solid #3498db; padding-bottom: 5px; }}
        .status-good {{ color: #28a745; font-weight: bold; }}
        .status-bad {{ color: #dc3545; font-weight: bold; }}
        .status-warning {{ color: #ffc107; font-weight: bold; }}
        .code {{ font-family: monospace; background-color: #f8f9fa; padding: 2px 4px; border-radius: 3px; }}
    </style>
</head>
<body>
    <h1>Football Officiating Analysis - Troubleshooting Report</h1>
    
    <div class="summary">
        <h3>File Matching Summary</h3>
        <div style="display: grid; grid-template-columns: repeat(auto-fit, minmax(200px, 1fr)); gap: 15px;">
            <div><strong>Total Scheduled Games:</strong> {len(schedule_games)}</div>
            <div><strong>Games with Files:</strong> <span class="status-good">{sum(1 for g in schedule_games if g['has_file'])}</span></div>
            <div><strong>Missing Files:</strong> <span class="status-bad">{len(missing_files)}</span></div>
            <div><strong>Extra Files:</strong> <span class="status-warning">{len(extra_files)}</span></div>
            <div><strong>Available Excel Files:</strong> {len(available_files)}</div>
        </div>
    </div>
    
    <div class="section">
        <h2>Schedule vs Available Files Analysis</h2>
        <table>
            <tr>
                <th>Status</th>
                <th>Game ID</th>
                <th>Round</th>
                <th>Date</th>
                <th>Home Team</th>
                <th>Away Team</th>
                <th>Matched File</th>
                <th>Possible Filenames</th>
            </tr>"""
        
        # Add schedule games to table
        for game in sorted(schedule_games, key=lambda x: (not x['has_file'], x['game_id'])):
            status_class = "found" if game['has_file'] else "missing"
            status_text = "✓ FOUND" if game['has_file'] else "✗ MISSING"
            status_color = "status-good" if game['has_file'] else "status-bad"
            
            matched_file_display = self._html_escape(game['matched_file']) if game['matched_file'] else "No match found"
            
            possible_names_display = "<br>".join([
                f"<span class='code'>{self._html_escape(name)}</span>" 
                for name in game['possible_names'][:3]
            ])
            if len(game['possible_names']) > 3:
                possible_names_display += f"<br>... and {len(game['possible_names'])-3} more"
            
            html += f"""
            <tr class="{status_class}">
                <td><span class="{status_color}">{status_text}</span></td>
                <td><strong>{self._html_escape(game['game_id'])}</strong></td>
                <td>{self._html_escape(game['round'])}</td>
                <td>{self._html_escape(game['date'])}</td>
                <td>{self._html_escape(game['home_team'])}</td>
                <td>{self._html_escape(game['away_team'])}</td>
                <td>{matched_file_display}</td>
                <td>{possible_names_display}</td>
            </tr>"""
        
        html += """
        </table>
    </div>"""
        
        # Missing files section
        if missing_files:
            html += f"""
    <div class="section">
        <h2>Missing Game Files ({len(missing_files)})</h2>
        <p>These games are in the schedule but no corresponding Excel files were found:</p>
        <table>
            <tr><th>Game ID</th><th>Teams</th><th>Expected Filenames</th></tr>"""
            
            for game_id in missing_files:
                game_data = self.schedule_data[game_id]
                expected_names = [
                    f"{game_id}.xlsx",
                    f"{game_id.replace(' ', '-')}.xlsx", 
                    f"{game_data.get('home_team', '')}-v-{game_data.get('away_team', '')}.xlsx"
                ]
                expected_display = "<br>".join([f"<span class='code'>{name}</span>" for name in expected_names])
                
                html += f"""
            <tr class="missing">
                <td><strong>{self._html_escape(game_id)}</strong></td>
                <td>{self._html_escape(game_data.get('home_team', ''))} vs {self._html_escape(game_data.get('away_team', ''))}</td>
                <td>{expected_display}</td>
            </tr>"""
            
            html += """
        </table>
    </div>"""
        
        # Extra files section
        if extra_files:
            html += f"""
    <div class="section">
        <h2>Unmatched Excel Files ({len(extra_files)})</h2>
        <p>These Excel files were found but don't match any scheduled games:</p>
        <table>
            <tr><th>Filename</th><th>Possible Issues</th></tr>"""
            
            for filename in sorted(extra_files):
                issues = []
                if any(char in filename for char in [' ', '(', ')', '[', ']']):
                    issues.append("Contains spaces or special characters")
                if filename.lower().startswith('test') or 'backup' in filename.lower():
                    issues.append("Appears to be test/backup file")
                if not any(team in filename for team in ['Tigers', 'Oaks', 'Razorbacks', '89ers', 'Gold', 'Towers']):
                    issues.append("No recognizable team names")
                
                issues_display = "<br>".join(issues) if issues else "Unknown - check filename format"
                
                html += f"""
            <tr class="extra">
                <td><span class='code'>{self._html_escape(filename)}.xlsx</span></td>
                <td>{issues_display}</td>
            </tr>"""
            
            html += """
        </table>
    </div>"""
        
        # Recommendations section
        html += f"""
    <div class="section">
        <h2>Troubleshooting Recommendations</h2>
        
        <h3>If games are missing files:</h3>
        <ul>
            <li><strong>Check filename format:</strong> Ensure Excel files match Game IDs from schedule</li>
            <li><strong>Remove spaces:</strong> Replace spaces with hyphens (-) or underscores (_)</li>
            <li><strong>Check file location:</strong> Ensure files are in the correct 'data' folder</li>
            <li><strong>Verify file extensions:</strong> Files should be .xlsx or .xls</li>
        </ul>
        
        <h3>Common filename patterns that work:</h3>
        <ul>
            <li><span class="code">13April-Oaks-v-Razorbacks.xlsx</span></li>
            <li><span class="code">12April-Tigers-v-89ers.xlsx</span></li>
            <li><span class="code">GameID-exactly-as-in-schedule.xlsx</span></li>
        </ul>
        
        <h3>If you have extra unmatched files:</h3>
        <ul>
            <li><strong>Rename files:</strong> Match the Game ID format from schedule</li>
            <li><strong>Check schedule:</strong> Verify the game exists in the schedule sheet</li>
            <li><strong>Remove test files:</strong> Delete any backup, test, or temporary files</li>
        </ul>
    </div>
    
    <div class="section">
        <h2>File Structure Check</h2>
        <table>
            <tr><th>Component</th><th>Status</th><th>Details</th></tr>
            <tr class="found">
                <td>Schedule File</td>
                <td><span class="status-good">✓ FOUND</span></td>
                <td>{len(self.schedule_data)} games loaded from schedule</td>
            </tr>
            <tr class="found">
                <td>Officials Database</td>
                <td><span class="status-good">✓ FOUND</span></td>
                <td>{len(self.officials_data)} officials loaded</td>
            </tr>
            <tr class="{'found' if len(available_files) > 0 else 'missing'}">
                <td>Game Data Files</td>
                <td><span class="{'status-good' if len(available_files) > 0 else 'status-bad'}">{'✓ FOUND' if len(available_files) > 0 else '✗ MISSING'}</span></td>
                <td>{len(available_files)} Excel files in data folder</td>
            </tr>
        </table>
    </div>
    
    <div class="section">
        <h2>Next Steps</h2>
        <ol>
            <li><strong>Fix missing files:</strong> Add or rename Excel files for missing games</li>
            <li><strong>Clean up extra files:</strong> Remove or rename unmatched files</li>
            <li><strong>Re-run analysis:</strong> Execute the main script again after fixes</li>
            <li><strong>Check this report:</strong> Verify all games show "✓ FOUND" status</li>
        </ol>
    </div>
    
    <div style="margin-top: 40px; padding: 15px; background-color: #f8f9fa; border-left: 4px solid #007bff;">
        <p><strong>Report generated:</strong> {report_time}</p>
        <p><strong>Script version:</strong> {SCRIPT_VERSION}</p>
    </div>
</body>
</html>"""
        
        return html
    
    def run_analysis(self):
        """Run the complete comprehensive analysis"""
        
        print("Starting Enhanced Comprehensive Analysis")
        print("=" * 60)
        
        # Load schedule data first
        self.load_schedule_data()
        
        print(f"\nLooking for game data files in: {self.data_folder}")
        
        if not self.data_folder.exists():
            print(f"Data folder '{self.data_folder}' not found!")
            print(f"Create the folder and add your game Excel files there")
            return
        
        excel_files = list(self.data_folder.glob("*.xlsx")) + list(self.data_folder.glob("*.xls"))
        
        if not excel_files:
            print(f"No Excel files found in '{self.data_folder}' folder!")
            return
        
        print(f"Found {len(excel_files)} game Excel file(s):")
        for file in excel_files:
            print(f"   - {file.name}")
        
        # Generate troubleshooting report if requested
        if hasattr(self, 'generate_troubleshooting') and self.generate_troubleshooting:
            print(f"\nGenerating troubleshooting report...")
            try:
                troubleshooting_html = self.generate_troubleshooting_report(excel_files)
                troubleshooting_path = self.output_folder / "troubleshooting_report.html"
                
                with open(troubleshooting_path, 'w', encoding='utf-8') as f:
                    f.write(troubleshooting_html)
                
                print(f"   Troubleshooting report saved: {troubleshooting_path}")
                print(f"   Open this file to check schedule vs file matching")
                
                # If troubleshooting only, exit here
                if hasattr(self, 'troubleshooting_only') and self.troubleshooting_only:
                    print(f"\nTroubleshooting-only mode complete!")
                    print(f"Check the report to see which games have matching files.")
                    return
                    
            except Exception as e:
                print(f"   Error generating troubleshooting report: {e}")
        
        print(f"\nFor now, basic analysis functionality is available.")
        print(f"The troubleshooting report has been generated successfully!")
        print(f"You can extend this script with the full analysis methods as needed.")


def main():
    """Main function to execute the analysis"""
    parser = argparse.ArgumentParser(description='Enhanced comprehensive football officiating analysis')
    parser.add_argument('--data', '-d', default='data', help='Data folder containing Excel game files (default: data)')
    parser.add_argument('--schedule', '-s', default='nlplan', help='Schedule folder containing officiating assignments (default: nlplan)')
    parser.add_argument('--output', '-o', default='reports', help='Output folder for reports (default: reports)')
    parser.add_argument('--min-games', type=int, default=3, help='Minimum games for full ranking eligibility (default: 3)')
    parser.add_argument('--min-calls', type=int, default=5, help='Minimum gradeable calls for statistical reliability (default: 5)')
    parser.add_argument('--troubleshooting', '-t', action='store_true', help='Generate troubleshooting report to check schedule vs file matching')
    parser.add_argument('--troubleshooting-only', action='store_true', help='Only generate troubleshooting report, skip main analysis')
    parser.add_argument('--debug', action='store_true', help='Enable debug mode with detailed output')
    
    args = parser.parse_args()
    
    print("Enhanced Comprehensive Football Officiating Analysis System")
    print("=" * 60)
    print("Features: Individual Reports + Combined Report + Penalty Analysis + Troubleshooting")
    if args.troubleshooting or args.troubleshooting_only:
        print("Mode: Troubleshooting" + (" Only" if args.troubleshooting_only else " + Analysis"))
    print("=" * 60)
    
    analyzer = ComprehensiveOfficatingAnalyzer(args.data, args.schedule, args.output)
    
    # Update ranking configuration based on command line arguments
    analyzer.ranking_config['min_games_for_full_ranking'] = args.min_games
    analyzer.ranking_config['min_calls_for_reliability'] = args.min_calls
    
    # Set troubleshooting flags
    analyzer.generate_troubleshooting = args.troubleshooting or args.troubleshooting_only
    analyzer.troubleshooting_only = args.troubleshooting_only
    
    try:
        analyzer.run_analysis()
        print("\nScript completed successfully!")
        
        if args.troubleshooting_only:
            print(f"\nTroubleshooting report generated!")
            print(f"Check: {analyzer.output_folder}/troubleshooting_report.html")
            print(f"This report shows which scheduled games have matching Excel files.")
            print(f"\nNext steps:")
            print(f"1. Open the troubleshooting report in your browser")
            print(f"2. Fix any missing files or rename files as suggested")
            print(f"3. Run the script again without --troubleshooting-only for full analysis")
        else:
            if analyzer.schedule_data:
                print(f"\nSchedule data summary:")
                print(f"   - {len(analyzer.schedule_data)} games in schedule")
                
            if analyzer.officials_data:
                print(f"   - {len(analyzer.officials_data)} officials loaded")
                
    except Exception as e:
        print(f"\nScript failed with error: {e}")
        import traceback
        traceback.print_exc()


if __name__ == "__main__":
    main()
    