"""
Configuration module for timetable generation
Contains all constants, colors, and configuration loading functions
"""

import json
from datetime import time

# Load duration constants from config
def load_config():
    try:
        with open('config.json', 'r') as f:
            config = json.load(f)
            return config['duration_constants']
    except:
        # Return defaults if config file not found
        return {
            'hour_slots': 2,
            'lecture_duration': 3,
            'lab_duration': 4,
            'tutorial_duration': 2,
            'self_study_duration': 2, 
            'break_duration': 1
        }

# Initialize duration constants
durations = load_config()
HOUR_SLOTS = durations['hour_slots']
LECTURE_DURATION = durations['lecture_duration']
LAB_DURATION = durations['lab_duration']
TUTORIAL_DURATION = durations['tutorial_duration']
SELF_STUDY_DURATION = durations['self_study_duration']
BREAK_DURATION = durations['break_duration']

# Time Constants
DAYS = ['Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday']
START_TIME = time(9, 0)
END_TIME = time(18, 30)

# Lunch break parameters
LUNCH_WINDOW_START = time(12, 30)
LUNCH_WINDOW_END = time(14, 0)
LUNCH_DURATION = 60

# Color palette for subjects
SUBJECT_COLORS = [
    "FFB6C1", "98FB98", "87CEFA", "DDA0DD", "F0E68C", 
    "E6E6FA", "FFDAB9", "B0E0E6", "FFA07A", "D8BFD8",
    "AFEEEE", "F08080", "90EE90", "ADD8E6", "FFB6C1"
]

# Specific colors for basket groups
BASKET_GROUP_COLORS = {
    'B1': "FF9999",  # Light red
    'B2': "99FF99",  # Light green  
    'B3': "9999FF",  # Light blue
    'B4': "FFFF99",  # Light yellow
    'B5': "FF99FF",  # Light magenta
    'B6': "99FFFF",  # Light cyan
    'B7': "FFB366",  # Light orange
    'B8': "B366FF",  # Light purple
    'B9': "66FFB3"   # Light mint
}

# Excel formatting colors
HEADER_FILL_COLOR = "FFD700"
LEC_FILL_COLOR = "E6E6FA"
LAB_FILL_COLOR = "98FB98"
TUT_FILL_COLOR = "FFE4E1"
SS_FILL_COLOR = "ADD8E6"
BREAK_FILL_COLOR = "D3D3D3"
UNSCHEDULED_FILL_COLOR = "FFE0E0"
LEGEND_FILL_COLOR = "F0F0F0"

# Global variables
TIME_SLOTS = []
lunch_breaks = {}