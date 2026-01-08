#!/usr/bin/env python3
"""
Enhanced Comprehensive Football Officiating Performance Analysis System
Complete version with individual and combined reports, penalty-by-penalty analysis
FIXED VERSION: Correct game counting and position-based rankings
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
SCRIPT_VERSION = "1.5.0"
SCRIPT_CREATED = "2024-12-28"
SCRIPT_UPDATED = "2025-01-13"
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
        
        # FIXED: Enhanced official performance tracking
        self.official_performance_data = defaultdict(lambda: {
            'games': [],
            'scheduled_games': set(),  # NEW: Track ALL scheduled games
            'positions': defaultdict(int),
            'total_calls': defaultdict(int),
            'grade_breakdown': defaultdict(lambda: defaultdict(int)),
            'penalty_types': defaultdict(int),
            'position_performance': defaultdict(lambda: {  # NEW: Position-specific tracking
                'games': set(),
                'total_calls': defaultdict(int),
                'grade_breakdown': defaultdict(int),
                'penalty_types': defaultdict(int)
            })
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
            month = self._safe_get_value(row, ['Måned', 'Month', 'Måned'])
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
                                'date': f"{row.get('Dato', '')}-{row.get('Måned', '')}",
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
                            
                            # FIXED: Track ALL scheduled games for each official
                            position_mapping = {
                                'R': 'referee',
                                'U': 'umpire',
                                'H': 'head_linesman',
                                'L': 'line_judge',
                                'B': 'back_judge',
                                'F': 'field_judge',
                                'S': 'side_judge',
                                'C': 'center_judge'
                            }
                            
                            for pos_code, pos_field in position_mapping.items():
                                initials = game_data.get(pos_field, '')
                                if initials and initials.strip():
                                    full_name = self.get_official_name_from_initials(initials)
                                    if full_name != 'Unknown':
                                        # Track scheduled game
                                        self.official_performance_data[full_name]['scheduled_games'].add(game_id_clean)
                                        # Track position assignment
                                        self.official_performance_data[full_name]['positions'][self.official_positions[pos_code]] += 1
                                        # Track position-specific scheduled games
                                        self.official_performance_data[full_name]['position_performance'][pos_code]['games'].add(game_id_clean)
                            
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
                                'primary_position': row.get('Primær', ''),
                                'secondary_position': row.get('Sekundær', ''),
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
        
        # Debug: Print scheduled games count for officials
        print(f"\nScheduled games tracking:")
        for official_name, data in list(self.official_performance_data.items())[:5]:  # Show first 5
            if not official_name.startswith('Unknown'):
                scheduled_count = len(data['scheduled_games'])
                print(f"   {official_name}: {scheduled_count} scheduled games")
    
    def clean_official_name(self, name):
        """Clean and standardize official names/initials"""
        if pd.isna(name) or name == '':
            return ''
        
        name = str(name).strip()
        if '+' in name:
            return name.split('+')[0].strip()
        return name
    
    def get_official_for_position(self, game_id, position_code):
        """Get the official assigned to a specific position for a game - returns FULL NAME"""
        if game_id not in self.schedule_data:
            return f'Unknown (Game {game_id} not in schedule)'
        
        game_data = self.schedule_data[game_id]
        position_mapping = {
            'R': 'referee',
            'U': 'umpire',
            'H': 'head_linesman',
            'L': 'line_judge',
            'B': 'back_judge',
            'F': 'field_judge',
            'S': 'side_judge',
            'C': 'center_judge'
        }
        
        position_field = position_mapping.get(position_code, '')
        if position_field:
            initials = game_data.get(position_field, 'Unknown')
            if initials and initials != 'Unknown':
                full_name = self.get_official_name_from_initials(initials)
                return full_name if full_name != initials else f'{initials} (Name not found)'
            return f'No {position_field.replace("_", " ").title()} assigned'
        return f'Unknown position: {position_code}'

    def parse_grade_official(self, grade_official_str):
        """Parse the grade-official string and match with scheduled officials"""
        if pd.isna(grade_official_str) or not grade_official_str:
            return []
        
        grade_official_str = str(grade_official_str).upper().strip()
        results = []
        
        # Handle WAIVED specially
        if 'WAIVED' in grade_official_str or grade_official_str == 'W':
            for pos_code in self.official_positions.keys():
                pattern = f"{pos_code}(WAIVED|W)"
                if re.search(pattern, grade_official_str):
                    results.append({
                        'position_code': pos_code,
                        'position': self.official_positions[pos_code],
                        'grade': 'W'
                    })
            return results
        
        # Parse character by character for normal grades
        i = 0
        while i < len(grade_official_str):
            if grade_official_str[i] in self.official_positions:
                position_code = grade_official_str[i]
                
                grade = 'Unknown'
                next_pos = i + 1
                
                if next_pos < len(grade_official_str):
                    # Check for two-letter grades first
                    if next_pos + 1 < len(grade_official_str):
                        two_char = grade_official_str[next_pos:next_pos + 2]
                        if two_char in ['CC', 'MC', 'IC', 'NC', 'NG']:
                            grade = self.normalize_grade(two_char)
                            i = next_pos + 2
                        elif grade_official_str[next_pos] in ['C', 'M', 'I', 'N', 'G', 'W']:
                            grade = self.normalize_grade(grade_official_str[next_pos])
                            i = next_pos + 1
                        else:
                            i = next_pos + 1
                    elif grade_official_str[next_pos] in ['C', 'M', 'I', 'N', 'G', 'W']:
                        grade = self.normalize_grade(grade_official_str[next_pos])
                        i = next_pos + 1
                    else:
                        i = next_pos + 1
                else:
                    i += 1
                
                results.append({
                    'position_code': position_code,
                    'position': self.official_positions[position_code],
                    'grade': grade
                })
            else:
                i += 1
        
        return results

    def load_game_data(self, file_path):
        """Load data from a single game Excel file"""
        try:
            print(f"   Loading game data from: {file_path.name}")
            
            # Try different engines if one fails
            engines = ['openpyxl', 'xlrd', None]
            df = None
            
            for engine in engines:
                try:
                    if engine:
                        df = pd.read_excel(file_path, engine=engine)
                    else:
                        df = pd.read_excel(file_path)
                    break
                except Exception:
                    continue
            
            if df is None:
                print(f"   Could not load file with any engine")
                return [], ""
            
            game_name = file_path.stem
            
            if df.empty:
                print(f"   File {file_path.name} is empty")
                return [], ""
                
            df.columns = df.columns.astype(str).str.strip()
            print(f"   Loaded {len(df)} rows from {game_name}")
            
            penalties = []
            penalty_count = 0
            
            for _, row in df.iterrows():
                play_num = row.get('PLAY #')
                if not self._is_valid_value(play_num) or str(play_num).strip() == 'PLAY #':
                    continue
                
                play_data = {
                    'game': game_name,
                    'play_num': play_num,
                    'quarter': row.get('QTR', ''),
                    'penalty_cat_1': row.get('PENALTY-CAT 1', ''),
                    'flag_1': row.get('FLAG 1', ''),
                    'grade_official_1': row.get('GRADE OFFICIAL 1', ''),
                    'penalty_cat_2': row.get('PENALTY CAT 2', ''),
                    'flag_2': row.get('FLAG 2', ''),
                    'grade_official_2': row.get('GRADE OFFICIAL 2', ''),
                    'group': row.get('GROUP', ''),
                    'category': row.get('CATEGORY', ''),
                    'pos_foul': row.get('POS. FOUL', ''),
                    'viewers_notes': row.get('VIEVERS NOTES', ''),
                    'penalty': row.get('PENALTY', '')
                }
                
                has_penalty = any([
                    self._is_valid_value(play_data['penalty_cat_1']),
                    self._is_valid_value(play_data['flag_1']),
                    self._is_valid_value(play_data['grade_official_1'])
                ])
                
                if has_penalty:
                    self.penalty_plays.append(play_data)
                    penalty_count += 1
                
                penalties.append(play_data)
            
            print(f"   Loaded {penalty_count} penalty plays from {game_name}")
            return penalties, game_name
            
        except Exception as e:
            print(f"Error loading {file_path}: {str(e)}")
            return [], ""
    
    def analyze_game(self, game_data, game_name):
        """Analyze a single game's officiating performance"""
        
        penalty_counts = defaultdict(int)
        grade_counts = defaultdict(int)
        official_performance = defaultdict(lambda: defaultdict(int))
        penalty_details = []
        
        for play in game_data:
            # Process penalty 1
            if pd.notna(play['penalty_cat_1']) and play['penalty_cat_1'] != '':
                penalty_counts[play['penalty_cat_1']] += 1
                
                if pd.notna(play['flag_1']) and play['flag_1'] != '':
                    normalized_flag = self.normalize_grade(play['flag_1'])
                    grade_counts[normalized_flag] += 1
                
                if pd.notna(play['grade_official_1']) and play['grade_official_1'] != '':
                    official_grades = self.parse_grade_official(play['grade_official_1'])
                    for og in official_grades:
                        official_name = self.get_official_for_position(game_name, og['position_code'])
                        
                        # FIXED: Track penalty data games (different from scheduled games)
                        self.official_performance_data[official_name]['games'].append(game_name)
                        self.official_performance_data[official_name]['total_calls'][og['grade']] += 1
                        self.official_performance_data[official_name]['grade_breakdown'][og['position']][og['grade']] += 1
                        self.official_performance_data[official_name]['penalty_types'][play['penalty_cat_1']] += 1
                        
                        # FIXED: Track position-specific performance
                        pos_data = self.official_performance_data[official_name]['position_performance'][og['position_code']]
                        pos_data['total_calls'][og['grade']] += 1
                        pos_data['penalty_types'][play['penalty_cat_1']] += 1
                        
                        official_performance[official_name][og['grade']] += 1
                
                penalty_details.append({
                    'play': play['play_num'],
                    'quarter': play['quarter'],
                    'penalty': play['penalty_cat_1'],
                    'flag': play['flag_1'],
                    'officials': play['grade_official_1']
                })
            
            # Process penalty 2 if exists
            if pd.notna(play['penalty_cat_2']) and play['penalty_cat_2'] != '':
                penalty_counts[play['penalty_cat_2']] += 1
                
                if pd.notna(play['flag_2']) and play['flag_2'] != '':
                    normalized_flag = self.normalize_grade(play['flag_2'])
                    grade_counts[normalized_flag] += 1
                
                if pd.notna(play['grade_official_2']) and play['grade_official_2'] != '':
                    official_grades = self.parse_grade_official(play['grade_official_2'])
                    for og in official_grades:
                        official_name = self.get_official_for_position(game_name, og['position_code'])
                        
                        # FIXED: Track penalty data games (different from scheduled games)
                        self.official_performance_data[official_name]['games'].append(game_name)
                        self.official_performance_data[official_name]['total_calls'][og['grade']] += 1
                        self.official_performance_data[official_name]['grade_breakdown'][og['position']][og['grade']] += 1
                        self.official_performance_data[official_name]['penalty_types'][play['penalty_cat_2']] += 1
                        
                        # FIXED: Track position-specific performance
                        pos_data = self.official_performance_data[official_name]['position_performance'][og['position_code']]
                        pos_data['total_calls'][og['grade']] += 1
                        pos_data['penalty_types'][play['penalty_cat_2']] += 1
                        
                        official_performance[official_name][og['grade']] += 1
                
                penalty_details.append({
                    'play': play['play_num'],
                    'quarter': play['quarter'],
                    'penalty': play['penalty_cat_2'],
                    'flag': play['flag_2'],
                    'officials': play['grade_official_2']
                })
        
        return {
            'penalty_counts': dict(penalty_counts),
            'grade_counts': dict(grade_counts),
            'official_performance': dict(official_performance),
            'penalty_details': penalty_details,
            'total_penalties': sum(penalty_counts.values())
        }
    
    def generate_metadata_html(self):
        """Generate metadata section for reports"""
        report_generated = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        
        return f"""
        <div style="background-color: #f8f9fa; border: 1px solid #dee2e6; padding: 10px; margin-bottom: 20px; font-size: 0.9em;">
            <h4 style="margin: 0 0 10px 0; color: #495057;">Report Metadata</h4>
            <div style="display: grid; grid-template-columns: repeat(auto-fit, minmax(250px, 1fr)); gap: 10px;">
                <div><strong>Script Version:</strong> {SCRIPT_VERSION}</div>
                <div><strong>Script Created:</strong> {SCRIPT_CREATED}</div>
                <div><strong>Last Updated:</strong> {SCRIPT_UPDATED}</div>
                <div><strong>Report Generated:</strong> {report_generated}</div>
                <div><strong>Generated By:</strong> {SCRIPT_AUTHOR}</div>
            </div>
        </div>
        """
    
    def calculate_accuracy_score(self, grade_counts):
        """Calculate weighted accuracy score based on grading philosophy"""
        
        if not grade_counts or not isinstance(grade_counts, dict):
            return 0, 0, {}
        
        total_weighted_score = 0
        total_gradeable_calls = 0
        
        for grade, count in grade_counts.items():
            if not isinstance(count, (int, float)) or count < 0:
                continue
                
            normalized_grade = self.normalize_grade(grade)
            if normalized_grade in self.grade_weights and self.grade_weights[normalized_grade] is not None:
                weight = self.grade_weights[normalized_grade]
                total_weighted_score += weight * count
                total_gradeable_calls += count
        
        if total_gradeable_calls == 0:
            return 0, 0, {}
        
        accuracy_score = total_weighted_score / total_gradeable_calls
        accuracy_score = max(0, min(100, accuracy_score))
        
        # Calculate breakdown
        breakdown = {}
        for grade in ['C', 'M', 'N', 'I']:
            count = 0
            for original_grade, original_count in grade_counts.items():
                if self.normalize_grade(original_grade) == grade:
                    count += original_count
                    
            weight = self.grade_weights.get(grade, 0)
            breakdown[grade] = {
                'count': count,
                'weight': weight,
                'contribution': (weight * count) if count > 0 else 0,
                'percentage': (count / total_gradeable_calls * 100) if total_gradeable_calls > 0 else 0
            }
        
        return accuracy_score, total_gradeable_calls, breakdown
    
    def generate_individual_official_report_html(self, official_name):
        """Generate detailed HTML report for an individual official"""
        
        if official_name not in self.official_performance_data:
            return f"""<html><body>
            <h1>No data found for official: {self._html_escape(official_name)}</h1>
            </body></html>"""
        
        data = self.official_performance_data[official_name]
        
        if not data or not any(data.get('total_calls', {}).values()):
            return f"""<html><body>
            <h1>No penalty call data found for official: {self._html_escape(

official_name)}</h1>
            </body></html>"""
        
        # Get official details from database
        official_details = {}
        for initials, details in self.officials_data.items():
            if details['name'] == official_name or initials == official_name:
                official_details = details
                break
        
        # FIXED: Calculate statistics using correct game counts
        penalty_data_games = list(set(data['games']))  # Games with penalty data
        total_scheduled_games = len(data['scheduled_games'])  # All scheduled games
        total_calls = sum(data['total_calls'].values())
        accuracy_score, gradeable_calls, accuracy_breakdown = self.calculate_accuracy_score(data['total_calls'])
        
        # Better handling of grade counts with normalization
        cc_calls = sum(count for grade, count in data['total_calls'].items() if self.normalize_grade(grade) == 'C')
        mc_calls = sum(count for grade, count in data['total_calls'].items() if self.normalize_grade(grade) == 'M')
        ic_calls = sum(count for grade, count in data['total_calls'].items() if self.normalize_grade(grade) == 'I')
        nc_calls = sum(count for grade, count in data['total_calls'].items() if self.normalize_grade(grade) == 'N')
        ng_calls = sum(count for grade, count in data['total_calls'].items() if self.normalize_grade(grade) == 'G')
        waived_calls = sum(count for grade, count in data['total_calls'].items() if self.normalize_grade(grade) == 'W')

        # Collect penalty-by-penalty analysis for this official
        penalty_analysis = defaultdict(lambda: {'grades': defaultdict(int), 'total': 0})
        
        def process_penalty_for_official(play, penalty_field, grade_field):
            """Helper function to process penalty data safely"""
            if not play.get(grade_field):
                return
                
            try:
                official_grades = self.parse_grade_official(play[grade_field])
                for og in official_grades:
                    found_official = self.get_official_for_position(play['game'], og['position_code'])
                    if found_official == official_name:
                        penalty_type = play.get(penalty_field, '').strip()
                        if penalty_type and penalty_type not in ['', 'nan', 'None']:
                            normalized_grade = self.normalize_grade(og['grade'])
                            penalty_analysis[penalty_type]['grades'][normalized_grade] += 1
                            penalty_analysis[penalty_type]['total'] += 1
            except Exception:
                pass
        
        for play in self.penalty_plays:
            process_penalty_for_official(play, 'penalty_cat_1', 'grade_official_1')
            process_penalty_for_official(play, 'penalty_cat_2', 'grade_official_2')

        html = f"""
        <!DOCTYPE html>
        <html>
        <head>
            <title>Official Report - {self._html_escape(official_name)}</title>
            <style>
                body {{ font-family: Arial, sans-serif; margin: 20px; }}
                table {{ border-collapse: collapse; width: 100%; margin: 20px 0; }}
                th, td {{ border: 1px solid #ddd; padding: 8px; text-align: left; }}
                th {{ background-color: #f2f2f2; }}
                .summary {{ background-color: #e8f4f8; padding: 15px; margin: 20px 0; }}
                .section {{ margin: 30px 0; }}
                h1 {{ color: #2c3e50; }}
                h2 {{ color: #34495e; border-bottom: 2px solid #3498db; padding-bottom: 5px; }}
                .excellent {{ background-color: #d4edda; }}
                .acceptable {{ background-color: #fff3cd; }}
                .attention {{ background-color: #ffeaa7; }}
                .critical {{ background-color: #f8d7da; }}
                .excluded {{ background-color: #e2e3e5; }}
                .grade-legend {{ background-color: #f8f9fa; border: 1px solid #dee2e6; padding: 15px; margin: 20px 0; }}
            </style>
        </head>
        <body>
            <h1>Official Performance Report</h1>
            <h2>Official: {self._html_escape(official_name)}</h2>
            
            {self.generate_metadata_html()}
            
            <div class="grade-legend">
                <h4>Grade Legend</h4>
                <table style="margin: 10px 0;">
                    <tr><th>Code</th><th>Explanation</th><th>Description</th></tr>
                    <tr><td><strong>CC</strong></td><td>Correct call</td><td>Correct by rule and philosophy</td></tr>
                    <tr><td><strong>MC</strong></td><td>Marginal call</td><td>Technically correct by rule but not by philosophy</td></tr>
                    <tr><td><strong>NC</strong></td><td>No-call</td><td>Foul that should have been called, but was not</td></tr>
                    <tr><td><strong>IC</strong></td><td>Incorrect call</td><td>A foul was called when it should not have been, either by rule or due to being too far from accepted philosophy</td></tr>
                    <tr class="excluded"><td><strong>NG</strong></td><td>Not graded</td><td>Not enough evidence on video to determine whether the call was correct or not</td></tr>
                    <tr class="excluded"><td><strong>W</strong></td><td>Waived</td><td>Waived - Just for reporting, should not be used for accuracy</td></tr>
                </table>
            </div>
            
            <div class="summary">
                <h3>Official Overview</h3>
                <p><strong>Full Name:</strong> {self._html_escape(official_details.get('name', 'N/A'))}</p>
                <p><strong>Status:</strong> {self._html_escape(official_details.get('status', 'N/A'))}</p>
                <p><strong>Primary Position:</strong> {self._html_escape(official_details.get('primary_position', 'N/A'))}</p>
                <p><strong>License Level:</strong> {self._html_escape(official_details.get('license', 'N/A'))}</p>
                <p><strong>Region:</strong> {self._html_escape(official_details.get('region', 'N/A'))}</p>
                <p><strong>Total Scheduled Games:</strong> {total_scheduled_games}</p>
                <p><strong>Games with Penalty Data:</strong> {len(penalty_data_games)}</p>
                <p><strong>Total Calls:</strong> {total_calls}</p>
                <p><strong>Gradeable Calls:</strong> {gradeable_calls}</p>
                <p><strong>Accuracy Score:</strong> {accuracy_score:.1f}%</p>
            </div>
            
            <div class="section">
                <h2>Performance Summary</h2>
                <table>
                    <tr><th>Grade</th><th>Points</th><th>Count</th><th>Weighted Score</th><th>% of Calls</th><th>Status</th></tr>
        """
        
        for grade in ['C', 'M', 'N', 'I']:
            grade_data = accuracy_breakdown.get(grade, {'count': 0, 'weight': 0, 'contribution': 0, 'percentage': 0})
            count = grade_data['count']
            weight = grade_data['weight']
            contribution = grade_data['contribution']
            percentage = grade_data['percentage']
            
            grade_display_map = {
                'C': 'CC - Correct Call', 
                'M': 'MC - Marginal Call', 
                'N': 'NC - No Call', 
                'I': 'IC - Incorrect Call'
            }
            grade_display = grade_display_map.get(grade, grade)
            
            status_map = {
                'C': "Excellent", 
                'M': "Acceptable", 
                'N': "Needs Attention", 
                'I': "Critical Issue"
            }
            status = status_map.get(grade, "Unknown")
            
            row_class_map = {
                'C': "excellent", 
                'M': "acceptable", 
                'N': "attention", 
                'I': "critical"
            }
            row_class = row_class_map.get(grade, "")
            
            html += f"""<tr class="{row_class}">
                <td><strong>{grade_display}</strong></td>
                <td>{weight}</td>
                <td>{count}</td>
                <td>{contribution:.0f}</td>
                <td>{percentage:.1f}%</td>
                <td>{status}</td>
            </tr>"""
        
        html += f"""
                    <tr class="excluded"><td>NG - Not Graded</td><td>-</td><td>{ng_calls}</td><td>Excluded</td><td>-</td><td>Not Counted</td></tr>
                    <tr class="excluded"><td>Waived</td><td>-</td><td>{waived_calls}</td><td>Excluded</td><td>-</td><td>Not Counted</td></tr>
                </table>
                <p><strong>Total Weighted Score:</strong> {sum(accuracy_breakdown[g]['contribution'] for g in accuracy_breakdown):.0f} points from {gradeable_calls} gradeable calls = {accuracy_score:.1f}% accuracy</p>
            </div>
            
            <div class="section">
                <h2>Penalty-by-Penalty Analysis</h2>
                <table>
                    <tr><th>Penalty Type</th><th>Times Called</th><th>C</th><th>M</th><th>N</th><th>I</th><th>G</th><th>W</th><th>Accuracy Score</th></tr>
        """
        
        # Sort penalties by frequency (most called first)
        sorted_penalties = sorted(penalty_analysis.items(), key=lambda x: x[1]['total'], reverse=True)
        
        for penalty_type, penalty_data in sorted_penalties:
            grades = penalty_data['grades']
            total_calls = penalty_data['total']
            
            cc = grades.get('C', 0)
            mc = grades.get('M', 0)
            nc = grades.get('N', 0)
            ic = grades.get('I', 0)
            ng = grades.get('G', 0)
            waived = grades.get('W', 0)
            
            # Calculate accuracy for this penalty type
            penalty_accuracy, gradeable_penalty_calls, _ = self.calculate_accuracy_score(grades)
            
            # Color code based on accuracy
            row_class = ""
            if gradeable_penalty_calls > 0:
                if penalty_accuracy >= 95:
                    row_class = "excellent"
                elif penalty_accuracy >= 85:
                    row_class = "acceptable"  
                elif penalty_accuracy >= 70:
                    row_class = "attention"
                else:
                    row_class = "critical"
            else:
                row_class = "excluded"
            
            accuracy_display = f"{penalty_accuracy:.1f}%" if gradeable_penalty_calls > 0 else "N/A"
            
            html += f"""<tr class="{row_class}">
                <td><strong>{self._html_escape(penalty_type)}</strong></td>
                <td>{total_calls}</td>
                <td>{cc}</td>
                <td>{mc}</td>
                <td>{nc}</td>
                <td>{ic}</td>
                <td>{ng}</td>
                <td>{waived}</td>
                <td><strong>{accuracy_display}</strong></td>
            </tr>"""
        
        html += """
                </table>
                <p><em>Note: Accuracy score calculated using weighted scoring system (C=100pts, M=75pts, N=50pts, I=0pts). G and W calls excluded from accuracy calculation.</em></p>
            </div>
            
            <div class="section">
                <h2>Games Officiated</h2>
                <table>
                    <tr><th>Game</th><th>Date</th><th>Teams</th><th>Position(s)</th><th>Penalty Data</th></tr>
        """
        
        # FIXED: Show ALL scheduled games, indicate which have penalty data
        all_scheduled_games = sorted(data['scheduled_games'])
        for game in all_scheduled_games:
            has_penalty_data = game in penalty_data_games
            penalty_data_status = "Yes" if has_penalty_data else "No"
            row_class = "" if has_penalty_data else "excluded"
            
            if game in self.schedule_data:
                game_info = self.schedule_data[game]
                positions = []
                for pos_code, pos_name in self.official_positions.items():
                    if self.get_official_for_position(game, pos_code) == official_name:
                        positions.append(pos_name)
                
                html += f"""<tr class="{row_class}">
                    <td>{self._html_escape(game)}</td>
                    <td>{self._html_escape(game_info.get('date', 'N/A'))}</td>
                    <td>{self._html_escape(game_info.get('home_team', ''))} vs {self._html_escape(game_info.get('away_team', ''))}</td>
                    <td>{self._html_escape(', '.join(positions) if positions else 'Unknown')}</td>
                    <td><strong>{penalty_data_status}</strong></td>
                </tr>"""
        
        html += """
                </table>
                <p><em>Note: Gray rows indicate scheduled games without penalty evaluation data.</em></p>
            </div>
        </body>
        </html>
        """
        
        return html
    
    def generate_position_rankings_html(self):
        """Generate position-specific rankings tables"""
        
        position_stats = defaultdict(list)  # pos_code -> [official_stats]
        
        # Collect position-specific statistics for each official
        for official_name, data in self.official_performance_data.items():
            if official_name.startswith('No ') or official_name.startswith('Unknown'):
                continue
                
            # Get official details
            official_details = {}
            for initials, details in self.officials_data.items():
                if details['name'] == official_name or initials == official_name:
                    official_details = details
                    break
            
            # Process each position this official worked
            for pos_code, pos_performance in data['position_performance'].items():
                if not pos_performance['total_calls']:  # Skip positions with no calls
                    continue
                    
                # Calculate position-specific metrics
                accuracy_score, gradeable_calls, breakdown = self.calculate_accuracy_score(pos_performance['total_calls'])
                total_calls = sum(pos_performance['total_calls'].values())
                scheduled_games_at_position = len(pos_performance['games'])
                
                if gradeable_calls == 0:
                    continue
                
                # Calculate confidence interval
                if gradeable_calls >= 5:
                    error_margin = 1.96 * math.sqrt((accuracy_score/100 * (1-accuracy_score/100)) / gradeable_calls) * 100
                    ci_lower = max(0, accuracy_score - error_margin)
                    ci_upper = min(100, accuracy_score + error_margin)
                    ci_width = ci_upper - ci_lower
                    reliability = "High" if scheduled_games_at_position >= 3 and gradeable_calls >= 5 else "Limited Data"
                else:
                    ci_lower = 0
                    ci_upper = 100
                    ci_width = 100
                    reliability = "Limited Data"
                
                # Extract grade counts safely
                correct_count = breakdown.get('C', {})
                if isinstance(correct_count, dict):
                    correct_count = correct_count.get('count', 0)
                else:
                    correct_count = 0
                    
                incorrect_count = breakdown.get('I', {})
                if isinstance(incorrect_count, dict):
                    incorrect_count = incorrect_count.get('count', 0)
                else:
                    incorrect_count = 0
                    
                no_call_count = breakdown.get('N', {})
                if isinstance(no_call_count, dict):
                    no_call_count = no_call_count.get('count', 0)
                else:
                    no_call_count = 0
                
                correct_pct = (correct_count / gradeable_calls * 100) if gradeable_calls > 0 else 0
                error_pct = ((incorrect_count + no_call_count) / gradeable_calls * 100) if gradeable_calls > 0 else 0
                calls_per_game = total_calls / scheduled_games_at_position if scheduled_games_at_position > 0 else 0
                
                position_stats[pos_code].append({
                    'name': official_name,
                    'position_name': self.official_positions[pos_code],
                    'accuracy': accuracy_score,
                    'games': scheduled_games_at_position,
                    'gradeable_calls': gradeable_calls,
                    'total_calls': total_calls,
                    'ci_lower': ci_lower,
                    'ci_upper': ci_upper,
                    'ci_width': ci_width,
                    'reliability': reliability,
                    'correct_pct': correct_pct,
                    'error_pct': error_pct,
                    'calls_per_game': calls_per_game,
                    'primary_position': official_details.get('primary_position', 'N/A'),
                    'license': official_details.get('license', 'N/A'),
                    'status': official_details.get('status', 'N/A'),
                    'breakdown': breakdown,
                    'c_count': correct_count,
                    'm_count': breakdown.get('M', {}).get('count', 0) if isinstance(breakdown.get('M', {}), dict) else 0,
                    'i_count': incorrect_count,
                    'n_count': no_call_count
                })
        
        # Sort each position by accuracy
        for pos_code in position_stats:
            position_stats[pos_code].sort(key=lambda x: x['accuracy'], reverse=True)
        
        # Generate HTML for position rankings
        html = """
            <div class="section">
                <h2>Position-Specific Rankings</h2>
                <div style="background-color: #d1ecf1; padding: 15px; border-left: 4px solid #17a2b8; margin-bottom: 20px;">
                    <h4>Position-Specific Analysis</h4>
                    <p>Officials are ranked separately for each position they worked, allowing for position-specific performance evaluation. An official may appear in multiple tables if they worked different positions.</p>
                    <ul>
                        <li><strong>Games:</strong> Number of scheduled games at this specific position</li>
                        <li><strong>Calls:</strong> Total penalty evaluations received while working this position</li>
                        <li><strong>Accuracy:</strong> Position-specific weighted accuracy score</li>
                        <li><strong>Primary Pos:</strong> Official's designated primary position</li>
                    </ul>
                </div>
        """
        
        # Generate table for each position
        position_order = ['R', 'U', 'H', 'L', 'F', 'S', 'B', 'C']  # Standard order
        
        for pos_code in position_order:
            if pos_code not in position_stats or not position_stats[pos_code]:
                continue
                
            position_name = self.official_positions[pos_code]
            officials_at_position = position_stats[pos_code]
            
            html += f"""
                <h3>{position_name} ({pos_code}) Rankings</h3>
                <p><em>{len(officials_at_position)} officials with penalty evaluation data at this position</em></p>
                <table>
                    <tr>
                        <th>Rank</th>
                        <th>Official</th>
                        <th>Accuracy</th>
                        <th>95% CI</th>
                        <th>Games</th>
                        <th>Calls</th>
                        <th>C</th>
                        <th>M</th>
                        <th>I</th>
                        <th>N</th>
                        <th>Primary Pos</th>
                        <th>License</th>
                        <th>Reliability</th>
                    </tr>
            """
            
            for i, official in enumerate(officials_at_position):
                rank = i + 1
                row_class = ""
                
                if official['reliability'] == "Limited Data":
                    row_class = "excluded"
                elif official['accuracy'] >= 90:
                    row_class = "excellent"
                elif official['accuracy'] >= 80:
                    row_class = "acceptable"
                elif official['accuracy'] >= 70:
                    row_class = "attention"
                else:
                    row_class = "critical"
                
                # Add special highlighting for top 3 with sufficient data
                if official['reliability'] == "High":
                    if rank == 1:
                        row_class = "rank1"
                    elif rank == 2:
                        row_class = "rank2"
                    elif rank == 3:
                        row_class = "rank3"
                
                # Highlight if working their primary position
                primary_indicator = ""
                if official['primary_position'] == position_name:
                    primary_indicator = " ★"
                
                html += f"""<tr class="{row_class}">
                    <td><strong>{rank}</strong></td>
                    <td><strong>{self._html_escape(official['name'])}{primary_indicator}</strong></td>
                    <td>{official['accuracy']:.1f}%</td>
                    <td>{official['ci_lower']:.1f}%-{official['ci_upper']:.1f}%</td>
                    <td>{official['games']}</td>
                    <td>{official['gradeable_calls']}</td>
                    <td>{official['c_count']}</td>
                    <td>{official['m_count']}</td>
                    <td>{official['i_count']}</td>
                    <td>{official['n_count']}</td>
                    <td>{self._html_escape(official['primary_position'])}</td>
                    <td>{self._html_escape(official['license'])}</td>
                    <td>{official['reliability']}</td>
                </tr>"""
            
            # Add position summary
            high_performers = len([o for o in officials_at_position if o['accuracy'] >= 90 and o['reliability'] == "High"])
            avg_accuracy = sum(o['accuracy'] for o in officials_at_position) / len(officials_at_position) if officials_at_position else 0
            experienced_officials = len([o for o in officials_at_position if o['reliability'] == "High"])
            
            html += f"""
                </table>
                <div style="background-color: #f8f9fa; padding: 10px; margin: 10px 0; font-size: 0.9em;">
                    <strong>Position Summary:</strong> 
                    Average Accuracy: {avg_accuracy:.1f}% | 
                    High Performers (90%+): {high_performers} | 
                    Experienced Officials: {experienced_officials} | 
                    ★ = Working Primary Position
                </div>
            """
        
        html += """
            </div>
        """
        
        return html
    
    def generate_combined_report_html(self, all_analyses):
        """Generate comprehensive combined report for all games and officials"""
        
        # Aggregate data across all games
        total_penalty_counts = defaultdict(int)
        total_grade_counts = defaultdict(int)
        
        for game_name, analysis in all_analyses.items():
            for penalty, count in analysis['penalty_counts'].items():
                total_penalty_counts[penalty] += count
            
            for grade, count in analysis['grade_counts'].items():
                total_grade_counts[grade] += count
        
        total_penalties = sum(total_penalty_counts.values())
        
        # Calculate official rankings and statistics
        official_stats = []
        meaningful_officials = []
        
        for official_name, data in self.official_performance_data.items():
            if official_name.startswith('No ') or official_name.startswith('Unknown'):
                continue
                
            total_calls = sum(data['total_calls'].values())
            if total_calls == 0:
                continue
                
            accuracy_score, gradeable_calls, breakdown = self.calculate_accuracy_score(data['total_calls'])
            
            # FIXED: Use correct game counts
            penalty_data_games = len(set(data['games']))  # Games with penalty data
            total_scheduled_games = len(data['scheduled_games'])  # All scheduled games
            
            # Calculate confidence interval (simplified)
            if gradeable_calls >= 5:
                error_margin = 1.96 * math.sqrt((accuracy_score/100 * (1-accuracy_score/100)) / gradeable_calls) * 100
                ci_lower = max(0, accuracy_score - error_margin)
                ci_upper = min(100, accuracy_score + error_margin)
                ci_width = ci_upper - ci_lower
                reliability = "High" if total_scheduled_games >= 3 and gradeable_calls >= 5 else "Limited Data"
            else:
                ci_lower = 0
                ci_upper = 100
                ci_width = 100
                reliability = "Limited Data"
            
            # Calculate performance metrics - FIX: Handle case where breakdown values might be integers
            correct_count = breakdown.get('C', {})
            if isinstance(correct_count, dict):
                correct_count = correct_count.get('count', 0)
            else:
                correct_count = 0
                
            incorrect_count = breakdown.get('I', {})
            if isinstance(incorrect_count, dict):
                incorrect_count = incorrect_count.get('count', 0)
            else:
                incorrect_count = 0
                
            no_call_count = breakdown.get('N', {})
            if isinstance(no_call_count, dict):
                no_call_count = no_call_count.get('count', 0)
            else:
                no_call_count = 0
            
            correct_pct = (correct_count / gradeable_calls * 100) if gradeable_calls > 0 else 0
            error_pct = ((incorrect_count + no_call_count) / gradeable_calls * 100) if gradeable_calls > 0 else 0
            
            # FIXED: Calculate calls per game using ALL scheduled games
            calls_per_game = total_calls / total_scheduled_games if total_scheduled_games > 0 else 0
            
            # Get official details
            official_details = {}
            for initials, details in self.officials_data.items():
                if details['name'] == official_name or initials == official_name:
                    official_details = details
                    break
            
            # Get position summary - FIX: Handle case where positions might not be a dictionary
            position_summary = []
            positions_data = data.get('positions', {})
            if isinstance(positions_data, dict):
                for pos, count in positions_data.items():
                    position_summary.append(f"{pos} (n={count})")
            else:
                position_summary = ['Unknown']
            
            official_stats.append({
                'name': official_name,
                'accuracy': accuracy_score,
                'scheduled_games': total_scheduled_games,  # FIXED: Show scheduled games
                'penalty_data_games': penalty_data_games,  # FIXED: Show games with data
                'gradeable_calls': gradeable_calls,
                'total_calls': total_calls,
                'ci_lower': ci_lower,
                'ci_upper': ci_upper,
                'ci_width': ci_width,
                'reliability': reliability,
                'correct_pct': correct_pct,
                'error_pct': error_pct,
                'calls_per_game': calls_per_game,
                'primary_position': official_details.get('primary_position', 'N/A'),
                'license': official_details.get('license', 'N/A'),
                'status': official_details.get('status', 'N/A'),
                'positions': ', '.join(position_summary) if position_summary else 'Unknown',
                'breakdown': breakdown
            })
            
            if reliability == "High":
                meaningful_officials.append(official_stats[-1])
        
        # Sort officials by accuracy
        official_stats.sort(key=lambda x: x['accuracy'], reverse=True)
        meaningful_officials.sort(key=lambda x: x['accuracy'], reverse=True)
        
        # Calculate overall league metrics
        total_officials = len(official_stats)
        experienced_officials = len(meaningful_officials)
        avg_accuracy = sum(stat['accuracy'] for stat in meaningful_officials) / len(meaningful_officials) if meaningful_officials else 0
        high_performers = len([stat for stat in meaningful_officials if stat['accuracy'] >= 90])
        
        # Calculate overall league accuracy
        total_gradeable = sum(total_grade_counts.get(grade, 0) for grade in ['C', 'M', 'I', 'N'])
        if total_gradeable > 0:
            league_accuracy = sum([
                total_grade_counts.get('C', 0) * 100,
                total_grade_counts.get('M', 0) * 75,
                total_grade_counts.get('N', 0) * 50,
                total_grade_counts.get('I', 0) * 0
            ]) / total_gradeable
        else:
            league_accuracy = 0

        html = f"""
        <!DOCTYPE html>
        <html>
        <head>
            <title>Combined Officiating Report - Season Overview with Rankings</title>
            <style>
                body {{ font-family: Arial, sans-serif; margin: 20px; }}
                table {{ border-collapse: collapse; width: 100%; margin: 20px 0; }}
                th, td {{ border: 1px solid #ddd; padding: 8px; text-align: left; }}
                th {{ background-color: #f2f2f2; }}
                .summary {{ background-color: #e8f4f8; padding: 15px; margin: 20px 0; }}
                .section {{ margin: 30px 0; }}
                h1 {{ color: #2c3e50; }}
                h2 {{ color: #34495e; border-bottom: 2px solid #3498db; padding-bottom: 5px; }}
                .excellent {{ background-color: #d4edda; }}
                .acceptable {{ background-color: #fff3cd; }}
                .attention {{ background-color: #ffeaa7; }}
                .critical {{ background-color: #f8d7da; }}
                .excluded {{ background-color: #e2e3e5; }}
                .rank1 {{ background-color: #ffd700; font-weight: bold; }}
                .rank2 {{ background-color: #c0c0c0; font-weight: bold; }}
                .rank3 {{ backgroun
d-color: #cd7f32; font-weight: bold; }}
            </style>
        </head>
        <body>
            <h1>Comprehensive Football Officiating Analysis Report</h1>
            <h2>Season Overview - All Games with Official Rankings</h2>
            
            {self.generate_metadata_html()}
            
            <div class="summary">
                <h3>Season Summary</h3>
                <div style="display: grid; grid-template-columns: repeat(auto-fit, minmax(200px, 1fr)); gap: 15px;">
                    <div><strong>Total Games Analyzed:</strong> {len(all_analyses)}</div>
                    <div><strong>Total Officials Tracked:</strong> {total_officials}</div>
                    <div><strong>Total Penalty Calls:</strong> {total_penalties}</div>
                    <div><strong>Average Penalties per Game:</strong> {total_penalties / len(all_analyses):.1f}</div>
                    <div><strong>Officials with 3+ Games:</strong> {experienced_officials}</div>
                    <div><strong>Average Accuracy Score:</strong> {avg_accuracy:.1f}% (n={len(official_stats)})</div>
                </div>
            </div>
            
            <div class="section">
                <h2>Official Rankings & Performance Analysis</h2>
                
                <div style="background-color: #f8f9fa; padding: 15px; border-left: 4px solid #007bff; margin-bottom: 20px;">
                    <h4>Ranking Methodology</h4>
                    <ul>
                        <li><strong>Primary Metric:</strong> Weighted Accuracy Score (0-100 points)</li>
                        <li><strong>Minimum Reliability Threshold:</strong> 3+ games and 5+ gradeable calls</li>
                        <li><strong>Confidence Intervals:</strong> 95% confidence based on sample size</li>
                        <li><strong>Error Rate:</strong> Percentage of Incorrect Calls + No Calls</li>
                        <li><strong>Perfect Rate:</strong> Percentage of Correct Calls</li>
                        <li><strong>FIXED:</strong> Game counts now show ALL scheduled games vs games with penalty data</li>
                    </ul>
                </div>

                <h3>Top Performing Officials (Experienced)</h3>
                <p><em>Officials with 3+ games and 5+ gradeable calls</em></p>
                <table>
                    <tr>
                        <th>Rank</th>
                        <th>Official</th>
                        <th>Accuracy</th>
                        <th>95% CI</th>
                        <th>Scheduled Games</th>
                        <th>Data Games</th>
                        <th>Calls</th>
                        <th>Perfect%</th>
                        <th>Error%</th>
                        <th>Primary Position</th>
                        <th>License</th>
                        <th>Calls/Game</th>
                    </tr>
        """
        
        # Top performing officials table
        for i, official in enumerate(meaningful_officials):
            rank = i + 1
            row_class = ""
            if rank == 1:
                row_class = "rank1"
            elif rank == 2:
                row_class = "rank2"
            elif rank == 3:
                row_class = "rank3"
            elif official['accuracy'] >= 90:
                row_class = "excellent"
            elif official['accuracy'] >= 80:
                row_class = "acceptable"
            elif official['accuracy'] >= 70:
                row_class = "attention"
            else:
                row_class = "critical"
            
            html += f"""<tr class="{row_class}">
                <td><strong>{rank}</strong></td>
                <td><strong>{self._html_escape(official['name'])}</strong></td>
                <td>{official['accuracy']:.1f}%</td>
                <td>{official['ci_lower']:.1f}%-{official['ci_upper']:.1f}%</td>
                <td>{official['scheduled_games']}</td>
                <td>{official['penalty_data_games']}</td>
                <td>{official['gradeable_calls']}</td>
                <td>{official['correct_pct']:.1f}%</td>
                <td>{official['error_pct']:.1f}%</td>
                <td>{self._html_escape(official['primary_position'])}</td>
                <td>{self._html_escape(official['license'])}</td>
                <td>{official['calls_per_game']:.1f}</td>
            </tr>"""
        
        html += """
                </table>

                <h3>Complete Official Rankings</h3>
                <p><em>All officials sorted by accuracy score (including those with limited data)</em></p>
                <table>
                    <tr>
                        <th>Rank</th>
                        <th>Official</th>
                        <th>Accuracy</th>
                        <th>Reliability</th>
                        <th>95% CI Width</th>
                        <th>Scheduled Games</th>
                        <th>Data Games</th>
                        <th>Gradeable Calls</th>
                        <th>C</th>
                        <th>M</th>
                        <th>I</th>
                        <th>N</th>
                        <th>Position</th>
                        <th>Status</th>
                    </tr>
        """
        
        # Complete rankings table
        for i, official in enumerate(official_stats):
            rank = i + 1
            row_class = ""
            if official['reliability'] == "Limited Data":
                row_class = "excluded"
            elif official['accuracy'] >= 90:
                row_class = "excellent"
            elif official['accuracy'] >= 80:
                row_class = "acceptable"
            elif official['accuracy'] >= 70:
                row_class = "attention"
            else:
                row_class = "critical"
            
            # Safely extract breakdown counts
            breakdown = official['breakdown']
            c_count = breakdown.get('C', {}).get('count', 0) if isinstance(breakdown.get('C', {}), dict) else 0
            m_count = breakdown.get('M', {}).get('count', 0) if isinstance(breakdown.get('M', {}), dict) else 0
            i_count = breakdown.get('I', {}).get('count', 0) if isinstance(breakdown.get('I', {}), dict) else 0
            n_count = breakdown.get('N', {}).get('count', 0) if isinstance(breakdown.get('N', {}), dict) else 0
            
            html += f"""<tr class="{row_class}">
                <td>{rank}</td>
                <td><strong>{self._html_escape(official['name'])}</strong></td>
                <td>{official['accuracy']:.1f}%</td>
                <td>{official['reliability']}</td>
                <td>{official['ci_width']:.1f}%</td>
                <td>{official['scheduled_games']}</td>
                <td>{official['penalty_data_games']}</td>
                <td>{official['gradeable_calls']}</td>
                <td>{c_count}</td>
                <td>{m_count}</td>
                <td>{i_count}</td>
                <td>{n_count}</td>
                <td>{self._html_escape(official['positions'])}</td>
                <td>{self._html_escape(official['status'])}</td>
            </tr>"""
        
        # Performance insights
        officials_needing_support = len([stat for stat in meaningful_officials if stat['accuracy'] < 75])
        limited_data_officials = len([stat for stat in official_stats if stat['reliability'] == "Limited Data"])
        high_error_rate = len([stat for stat in meaningful_officials if stat['error_pct'] > 25])
        
        html += f"""
                </table>
                
                <div style="background-color: #fff3cd; padding: 10px; margin-top: 20px; border-left: 4px solid #ffc107;">
                    <h4>Understanding the Rankings</h4>
                    <ul>
                        <li><strong>Scheduled Games:</strong> Total games assigned to official in schedule</li>
                        <li><strong>Data Games:</strong> Games where official had penalty evaluations</li>
                        <li><strong>Confidence Interval Width:</strong> Narrower intervals indicate more reliable accuracy estimates</li>
                        <li><strong>Reliability Status:</strong> "High" means sufficient data for reliable comparison</li>
                        <li><strong>Limited Data:</strong> Officials may rank high/low due to small sample sizes</li>
                        <li><strong>Grade Legend:</strong> C=Correct, M=Marginal, I=Incorrect, N=No Call</li>
                        <li><strong>Calls/Game:</strong> Based on ALL scheduled games (not just games with penalty data)</li>
                    </ul>
                </div>
            </div>
            
            {self.generate_position_rankings_html()}
            
            <div class="section">
                <h2>Overall Call Grades Distribution</h2>
                <table>
                    <tr><th>Grade</th><th>Description</th><th>Count</th><th>Percentage</th></tr>
        """
        
        # Grade distribution
        total_grades = sum(total_grade_counts.values())
        grade_data = [
            ('C', 'CC - Correct Call', 'excellent'),
            ('M', 'MC - Marginal Call', ''),
            ('I', 'IC - Incorrect Call', 'critical'),
            ('N', 'NC - No Call', ''),
            ('G', 'NG - Non-gradeable', ''),
            ('W', 'W - Waived', '')
        ]
        
        for grade, description, css_class in grade_data:
            count = total_grade_counts.get(grade, 0)
            percentage = (count / total_grades * 100) if total_grades > 0 else 0
            html += f"""<tr class='{css_class}'><td>{grade}</td><td>{description}</td><td>{count}</td><td>{percentage:.1f}%</td></tr>"""
        
        html += f"""
                </table>
                <p><strong>Overall League Accuracy:</strong> {league_accuracy:.1f}%</p>
            </div>
            
            <div class="section">
                <h2>Performance Insights & Recommendations</h2>
                
                <div style="display: grid; grid-template-columns: repeat(auto-fit, minmax(300px, 1fr)); gap: 20px;">
                    <div style="background-color: #d4edda; padding: 15px; border-radius: 5px;">
                        <h4>Excellence Metrics</h4>
                        <ul>
                            <li><strong>High Performers (90%+):</strong> {high_performers} officials</li>
                            <li><strong>Average Accuracy (Experienced):</strong> {avg_accuracy:.1f}%</li>
                            <li><strong>Officials with 5+ Games:</strong> {len([s for s in meaningful_officials if s['scheduled_games'] >= 5])}</li>
                        </ul>
                    </div>
                    
                    <div style="background-color: #f8d7da; padding: 15px; border-radius: 5px;">
                        <h4>Areas for Improvement</h4>
                        <ul>
                            <li><strong>Officials Needing Support:</strong> {officials_needing_support} (&lt;75% accuracy)</li>
                            <li><strong>Limited Data Officials:</strong> {limited_data_officials}</li>
                            <li><strong>Error Rate Concern:</strong> {high_error_rate} officials &gt;25% errors</li>
                        </ul>
                    </div>
                    
                    <div style="background-color: #d1ecf1; padding: 15px; border-radius: 5px;">
                        <h4>Development Recommendations</h4>
                        <ul>
                            <li><strong>Mentorship Program:</strong> Pair top performers with developing officials</li>
                            <li><strong>Position Training:</strong> Focus on positions with lower average scores</li>
                            <li><strong>Consistency Work:</strong> Target officials with high error rates</li>
                            <li><strong>Game Assignment:</strong> Gradually increase games for promising officials</li>
                        </ul>
                    </div>
                </div>
                
                <div style="background-color: #fff3cd; padding: 15px; margin-top: 20px; border-left: 4px solid #ffc107;">
                    <h4>Key Performance Indicators</h4>
                    <div style="display: grid; grid-template-columns: repeat(auto-fit, minmax(250px, 1fr)); gap: 15px;">
                        <div><strong>League Standard:</strong> 80%+ accuracy target</div>
                        <div><strong>Excellence Threshold:</strong> 90%+ accuracy</div>
                        <div><strong>Minimum Experience:</strong> 3 games for ranking</div>
                        <div><strong>Statistical Reliability:</strong> 5+ gradeable calls</div>
                    </div>
                </div>
            </div>
            
            <div class="section">
                <h2>Most Common Penalties</h2>
                <table>
                    <tr>
                        <th>Penalty Type</th>
                        <th>Total</th>
        """
        
        # Add game headers for penalty breakdown
        sorted_games = sorted(all_analyses.keys())
        for game in sorted_games:
            html += f"<th style='writing-mode: vertical-rl; text-orientation: mixed; padding: 4px; font-size: 0.8em;'>{self._html_escape(game)}</th>"
        
        html += "</tr>"
        
        # Show top 15 most common penalties with game-by-game breakdown
        sorted_penalties = sorted(total_penalty_counts.items(), key=lambda x: x[1], reverse=True)
        
        for penalty, total_count in sorted_penalties[:15]:
            html += f"<tr><td><strong>{self._html_escape(penalty)}</strong></td><td><strong>{total_count}</strong></td>"
            
            for game in sorted_games:
                game_analysis = all_analyses[game]
                game_count = game_analysis['penalty_counts'].get(penalty, 0)
                if game_count > 0:
                    html += f"<td style='text-align: center; background-color: #e8f4f8;'><strong>{game_count}</strong></td>"
                else:
                    html += "<td style='text-align: center; color: #ccc;'>-</td>"
            
            html += "</tr>"
        
        html += """
                </table>
                <p><em>Note: Numbers show penalty counts per game. Games are sorted alphabetically. Empty cells (-) indicate no penalties of that type in the game.</em></p>
            </div>
            
            <div class="section">
                <h2>Navigation</h2>
                <p>Individual official reports are available as separate HTML files:</p>
                <ul>
        """
        
        # Navigation links to individual reports
        for official_name in sorted([stat['name'] for stat in official_stats]):
            safe_name = official_name.replace(' ', '_').replace('/', '_').replace('(', '').replace(')', '')
            html += f'<li><a href="{safe_name}_official_report.html">{self._html_escape(official_name)}</a></li>'
        
        html += """
                </ul>
            </div>
        </body>
        </html>
        """
        
        return html

    def run_analysis(self):
        """Run the complete comprehensive analysis"""
        
        print("Starting Enhanced Comprehensive Analysis (FIXED VERSION)")
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
        
        all_analyses = {}
        
        # Process each game file
        for file_path in excel_files:
            print(f"\nProcessing game file: {file_path.name}")
            
            game_data, game_name = self.load_game_data(file_path)
            if game_data:
                print(f"   Loaded {len(game_data)} plays from {game_name}")
                analysis = self.analyze_game(game_data, game_name)
                all_analyses[game_name] = analysis
                
                print(f"   Analysis complete for: {game_name}")
                print(f"      - Total penalties: {analysis['total_penalties']}")
                
                self.all_games_data.extend(game_data)
            else:
                print(f"   No data loaded from: {file_path.name}")
        
        print(f"\nGame Analysis Complete!")
        print("=" * 60)
        print(f"Total games analyzed: {len(all_analyses)}")
        print(f"Total officials tracked: {len(self.official_performance_data)}")
        print(f"Total penalty plays found: {len(self.penalty_plays)}")
        
        # FIXED: Show game count verification for a few officials
        print(f"\nFIXED - Game Count Verification:")
        sample_officials = list(self.official_performance_data.items())[:3]
        for official_name, data in sample_officials:
            if not official_name.startswith('Unknown'):
                scheduled = len(data['scheduled_games'])
                penalty_data = len(set(data['games']))
                print(f"   {official_name}: {scheduled} scheduled games, {penalty_data} games with penalty data")
        
        if len(all_analyses) == 0:
            print(f"\nWARNING: No games were successfully analyzed!")
            return
        
        # Check if we have meaningful official data
        meaningful_officials = [name for name in self.official_performance_data.keys() 
                              if not name.startswith('No ') and not name.startswith('Unknown')]
        
        # Generate reports
        print(f"\nGenerating reports...")
        reports_generated = 0
        
        # Generate individual official reports
        print(f"   Generating individual official reports...")
        for official_name in meaningful_officials:
            try:
                print(f"      Creating report for: {official_name}")
                official_html = self.generate_individual_official_report_html(official_name)
                safe_name = official_name.replace(' ', '_').replace('/', '_').replace('(', '').replace(')', '')
                official_report_path = self.output_folder / f"{safe_name}_official_report.html"
                
                with open(official_report_path, 'w', encoding='utf-8') as f:
                    f.write(official_html)
                
                print(f"      Generated: {official_report_path}")
                reports_generated += 1
                
            except Exception as e:
                print(f"      Error generating report for {official_name}: {e}")
        
        # Generate combined report
        if all_analyses:
            try:
                print(f"   Generating combined report with position rankings...")
                combined_html = self.generate_combined_report_html(all_analyses)
                combined_report_path = self.output_folder / "comprehensive_combined_report.html"
                
                with open(combined_report_path, 'w', encoding='utf-8') as f:
                    f.write(combined_html)
                
                print(f"   Generated combined report: {combined_report_path}")
                reports_generated += 1
                
            except Exception as e:
                print(f"   Error generating combined report: {e}")
        
        # Save detailed penalty plays data
        if self.penalty_plays:
            try:
                print(f"   Generating penalty data CSV...")
                penalty_df = pd.DataFrame(self.penalty_plays)
                penalty_csv_path = self.output_folder / "comprehensive_penalty_analysis.csv"
                penalty_df.to_csv(penalty_csv_path, index=False)
                print(f"   Penalty data saved to: {penalty_csv_path}")
                reports_generated += 1
                
            except Exception as e:
                print(f"   Error generating penalty CSV: {e}")
        
        print(f"\nAnalysis Complete!")
        print(f"Reports generated: {reports_generated}")
        print(f"Reports saved in: {self.output_folder}")
        
        if reports_generated > 0:
            print(f"\nFiles created:")
            for report_file in self.output_folder.glob("*"):
                print(f"   - {report_file.name}")


def main():
    """Main function to execute the analysis"""
    parser = argparse.ArgumentParser(description='Enhanced comprehensive football officiating analysis')
    parser.add_argument('--data', '-d', default='data', help='Data folder containing Excel game files (default: data)')
    parser.add_argument('--schedule', '-s', default='nlplan', help='Schedule folder containing officiating assignments (default: nlplan)')
    parser.add_argument('--output', '-o', default='reports', help='Output folder for reports (default: reports)')
    parser.add_argument('--min-games', type=int, default=3, help='Minimum games for full ranking eligibility (default: 3)')
    parser.add_argument('--min-calls', type=int, default=5, help='Minimum gradeable calls for statistical reliability (default: 5)')
    parser.add_argument('--debug', action='store_true', help='Enable debug mode with detailed output')
    
    args = parser.parse_args()
    
    print("Enhanced Comprehensive Football Officiating Analysis System")
    print("=" * 60)
    print("FIXED VERSION: Correct game counting + Position-based rankings")
    print("Features: Individual Reports + Combined Report + Position Rankings + Penalty Analysis")
    print("=" * 60)
    
    analyzer = ComprehensiveOfficatingAnalyzer(args.data, args.schedule, args.output)
    
    # Update ranking configuration based on command line arguments
    analyzer.ranking_config['min_games_for_full_ranking'] = args.min_games
    analyzer.ranking_config['min_calls_for_reliability'] = args.min_calls
    
    try:
        analyzer.run_analysis()
        print("\nScript completed successfully!")
        
        print(f"\nKey Fixes Applied:")
        print(f"   ✓ Game counting now shows ALL scheduled games vs games with penalty data")
        print(f"   ✓ Position-specific rankings added to combined report")
        print(f"   ✓ Officials appear in multiple position tables if they worked different positions")
        print(f"   ✓ Individual reports show complete scheduled game history")
        
        if analyzer.schedule_data:
            print(f"\nSchedule data summary:")
            print(f"   - {len(analyzer.schedule_data)} games in schedule")
            
        if analyzer.official_performance_data:
            print(f"   - {len(analyzer.official_performance_data)} officials have performance data")
                
    except Exception as e:
        print(f"\nScript failed with error: {e}")
        import traceback
        traceback.print_exc()


if __name__ == "__main__":
    main()