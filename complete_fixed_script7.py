#!/usr/bin/env python3
"""
FIXED VERSION: Enhanced Comprehensive Football Officiating Performance Analysis System
KEY FIXES:
1. Better game ID matching between schedule and data files
2. Improved official name resolution and tracking
3. Fallback logic for when officials have penalty data but no scheduled games
4. Better debugging and logging
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
SCRIPT_VERSION = "1.6.0"
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
        self.game_filename_to_id = {}  # NEW: Map filenames to game IDs
        
        # FIXED: Enhanced official performance tracking with better name resolution
        self.official_performance_data = defaultdict(lambda: {
            'games': [],
            'scheduled_games': set(),  # Track ALL scheduled games
            'positions': defaultdict(int),
            'total_calls': defaultdict(int),
            'grade_breakdown': defaultdict(lambda: defaultdict(int)),
            'penalty_types': defaultdict(int),
            'position_performance': defaultdict(lambda: {
                'games': set(),
                'total_calls': defaultdict(int),
                'grade_breakdown': defaultdict(int),
                'penalty_types': defaultdict(int)
            }),
            'alternative_names': set()  # NEW: Track different name variations
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
            month = self._safe_get_value(row, ['MÃ¥ned', 'Month', 'MÃ¥ned'])
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
    
    def create_game_mapping(self, filename):
        """FIXED: Create better mapping between filename and potential game IDs"""
        # Extract components from filename
        filename_clean = filename.replace('.xlsx', '').replace('.xls', '')
        
        # Try to extract date and teams
        patterns = [
            r'(\d+)(\w+)([A-Z][a-z_]+)v([A-Z][a-z_]+)',  # Like "31AugustTowersvRazorbacks"
            r'(\d+)([A-Z][a-z]+)-([A-Z][a-z_]+)-v-([A-Z][a-z_]+)',  # Alternative formats
        ]
        
        possible_ids = set()
        possible_ids.add(filename_clean)  # Always include exact filename
        
        for pattern in patterns:
            match = re.match(pattern, filename_clean)
            if match:
                day, month, team1, team2 = match.groups()
                
                # Create various possible game ID formats
                possible_ids.add(f"{day}{month}-{team1}-v-{team2}")
                possible_ids.add(f"{day}-{month}-{team1}-v-{team2}")
                possible_ids.add(f"{team1}v{team2}-{day}{month}")
                possible_ids.add(f"{team1} vs {team2}")
                
                # Handle underscores
                team1_clean = team1.replace('_', ' ')
                team2_clean = team2.replace('_', ' ')
                possible_ids.add(f"{team1_clean} vs {team2_clean}")
        
        return possible_ids
    
    def get_official_name_from_initials(self, initials):
        """Convert initials to full name using officials data"""
        if pd.isna(initials) or initials == '':
            return 'Unknown'
        
        initials = str(initials).strip()
        if initials in self.officials_data:
            return self.officials_data[initials]['name']
        return initials  # Return initials if no match found
    
    def normalize_official_name(self, name):
        """FIXED: Normalize official names for consistent tracking"""
        if pd.isna(name) or not name or str(name).strip() == '':
            return None
            
        name = str(name).strip()
        
        # Skip obviously invalid names
        if name.startswith('No ') or name.startswith('Unknown'):
            return None
            
        # Handle name variations
        if '(Name not found)' in name:
            # Extract the initials part
            base_name = name.replace('(Name not found)', '').strip()
            if base_name:
                return base_name
        
        return name
    
    def merge_official_data(self, primary_name, secondary_name):
        """FIXED: Merge data when we discover name variations refer to same official"""
        if not primary_name or not secondary_name or primary_name == secondary_name:
            return
            
        primary_data = self.official_performance_data[primary_name]
        secondary_data = self.official_performance_data[secondary_name]
        
        # Merge all the data
        primary_data['games'].extend(secondary_data['games'])
        primary_data['scheduled_games'].update(secondary_data['scheduled_games'])
        primary_data['alternative_names'].add(secondary_name)
        
        for key, value in secondary_data['total_calls'].items():
            primary_data['total_calls'][key] += value
            
        for key, value in secondary_data['penalty_types'].items():
            primary_data['penalty_types'][key] += value
            
        for pos, value in secondary_data['positions'].items():
            primary_data['positions'][pos] += value
        
        # Merge position performance
        for pos_code, pos_data in secondary_data['position_performance'].items():
            primary_pos_data = primary_data['position_performance'][pos_code]
            primary_pos_data['games'].update(pos_data['games'])
            for key, value in pos_data['total_calls'].items():
                primary_pos_data['total_calls'][key] += value
            for key, value in pos_data['penalty_types'].items():
                primary_pos_data['penalty_types'][key] += value
        
        # Remove the secondary entry
        if secondary_name in self.official_performance_data:
            del self.official_performance_data[secondary_name]
        
        print(f"   Merged official data: '{secondary_name}' -> '{primary_name}'")
    
    def load_schedule_data(self):
        """FIXED: Load officiating schedule and officials database with better game mapping"""
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
                
                # Load officials data first (we need this for name resolution)
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
                                'primary_position': row.get('PrimÃ¦r', ''),
                                'secondary_position': row.get('SekundÃ¦r', ''),
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
                
                # Now load games data from "Plan - NL" sheet
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
                                'date': f"{row.get('Dato', '')}-{row.get('MÃ¥ned', '')}",
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
                            
                            # FIXED: Track ALL scheduled games for each official with proper name resolution
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
                                    # Get full name from initials
                                    full_name = self.get_official_name_from_initials(initials)
                                    normalized_name = self.normalize_official_name(full_name)
                                    
                                    if normalized_name:
                                        # Track scheduled game
                                        self.official_performance_data[normalized_name]['scheduled_games'].add(game_id_clean)
                                        # Track position assignment
                                        self.official_performance_data[normalized_name]['positions'][self.official_positions[pos_code]] += 1
                                        # Track position-specific scheduled games
                                        self.official_performance_data[normalized_name]['position_performance'][pos_code]['games'].add(game_id_clean)
                                        # Track alternative names
                                        self.official_performance_data[normalized_name]['alternative_names'].add(initials)
                                        if full_name != normalized_name:
                                            self.official_performance_data[normalized_name]['alternative_names'].add(full_name)
                            
                            games_processed += 1
                    
                    print(f"   Successfully processed {games_processed} games from 'Plan - NL'")
                    
                except Exception as e:
                    print(f"   Error processing 'Plan - NL': {e}")
                
            except Exception as file_error:
                print(f"Error processing schedule file {file_path}: {file_error}")
        
        print(f"\nSchedule Loading Summary:")
        print(f"   Games loaded: {len(self.schedule_data)}")
        print(f"   Officials loaded: {len(self.officials_data)}")
        print(f"   Officials with scheduled games: {len([o for o in self.official_performance_data if self.official_performance_data[o]['scheduled_games']])}")
    
    def clean_official_name(self, name):
        """Clean and standardize official names/initials"""
        if pd.isna(name) or name == '':
            return ''
        
        name = str(name).strip()
        if '+' in name:
            return name.split('+')[0].strip()
        return name
    
    def find_game_id_for_filename(self, filename):
        """FIXED: Try to find matching game ID for a filename"""
        possible_ids = self.create_game_mapping(filename)
        
        # Check if any of the possible IDs match scheduled games
        for possible_id in possible_ids:
            if possible_id in self.schedule_data:
                return possible_id
        
        # If no exact match, try fuzzy matching based on teams
        filename_lower = filename.lower()
        for game_id, game_data in self.schedule_data.items():
            home_team = str(game_data.get('home_team', '')).lower().replace(' ', '')
            away_team = str(game_data.get('away_team', '')).lower().replace(' ', '')
            
            if (home_team and away_team and 
                home_team in filename_lower and away_team in filename_lower):
                print(f"   Fuzzy matched: {filename} -> {game_id} ({game_data.get('home_team', '')} vs {game_data.get('away_team', '')})")
                return game_id
        
        return filename  # Return filename as game ID if no match found
    
    def get_official_for_position(self, game_identifier, position_code):
        """FIXED: Get the official assigned to a specific position for a game - returns FULL NAME"""
        # First try to find the actual game ID
        game_id = self.find_game_id_for_filename(game_identifier)
        
        if game_id not in self.schedule_data:
            return f'Unknown (Game {game_identifier} not in schedule)'
        
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
                normalized_name = self.normalize_official_name(full_name)
                return normalized_name if normalized_name else f'{initials} (Name not found)'
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
            