"""
Utility functions for timetable generation
Contains helper functions for data loading, room allocation, and scheduling checks
"""

import pandas as pd
import csv
import os
from datetime import datetime, time, timedelta
import config

def generate_time_slots():
    """Generate all time slots for the day"""
    slots = []
    current_time = datetime.combine(datetime.today(), config.START_TIME)
    end_time = datetime.combine(datetime.today(), config.END_TIME)
    
    while current_time < end_time:
        current = current_time.time()
        next_time = current_time + timedelta(minutes=30)
        slots.append((current, next_time.time()))
        current_time = next_time
    
    return slots

def initialize_time_slots():
    """Initialize global time slots"""
    config.TIME_SLOTS = generate_time_slots()

def calculate_lunch_breaks(semesters):
    """Dynamically calculate staggered lunch breaks for semesters"""
    lunch_breaks = {}
    total_semesters = len(semesters)
    
    if total_semesters == 0:
        return lunch_breaks
        
    total_window_minutes = (
        config.LUNCH_WINDOW_END.hour * 60 + config.LUNCH_WINDOW_END.minute -
        config.LUNCH_WINDOW_START.hour * 60 - config.LUNCH_WINDOW_START.minute
    )
    stagger_interval = (total_window_minutes - config.LUNCH_DURATION) / (total_semesters - 1) if total_semesters > 1 else 0
    
    sorted_semesters = sorted(semesters)
    
    for i, semester in enumerate(sorted_semesters):
        start_minutes = (config.LUNCH_WINDOW_START.hour * 60 + config.LUNCH_WINDOW_START.minute + 
                        int(i * stagger_interval))
        start_hour = start_minutes // 60
        start_min = start_minutes % 60
        
        end_minutes = start_minutes + config.LUNCH_DURATION
        end_hour = end_minutes // 60
        end_min = end_minutes % 60
        
        lunch_breaks[semester] = (
            time(start_hour, start_min),
            time(end_hour, end_min)
        )
    
    return lunch_breaks

def load_rooms():
    """Load room information from CSV"""
    rooms = {}
    try:
        with open('rooms.csv', 'r') as f:
            reader = csv.DictReader(f)
            for row in reader:
                rooms[row['id']] = {
                    'capacity': int(row['capacity']),
                    'type': row['type'],
                    'roomNumber': row['roomNumber'],
                    'schedule': {day: set() for day in range(len(config.DAYS))}
                }
    except FileNotFoundError:
        print("Warning: rooms.csv not found, using default room allocation")
        return None
    return rooms

def load_batch_data():
    """Load batch information from combined.csv"""
    batch_info = {}
    
    try:
        df = pd.read_csv('combined.csv')
        grouped = df.groupby(['Department', 'Semester'])
        
        for (dept, sem), group in grouped:
            if 'total_students' in group.columns and not group['total_students'].isna().all():
                total_students = int(group['total_students'].max())
                max_batch_size = 70
                num_sections = (total_students + max_batch_size - 1) // max_batch_size
                section_size = (total_students + num_sections - 1) // num_sections

                batch_info[(dept, sem)] = {
                    'total': total_students,
                    'num_sections': num_sections,
                    'section_size': section_size
                }
        
        # Process basket courses
        basket_courses = df[df['Course Code'].astype(str).str.contains('^B[0-9]')]
        for _, course in basket_courses.iterrows():
            code = str(course['Course Code'])
            if 'total_students' in df.columns and pd.notna(course['total_students']):
                total_students = int(course['total_students'])
            else:
                total_students = 35
                
            batch_info[('ELECTIVE', code)] = {
                'total': total_students,
                'num_sections': 1,
                'section_size': total_students
            }
    except Exception as e:
        print(f"Warning: Error processing batch sizes: {e}")
        
    return batch_info

def load_reserved_slots():
    """Load reserved time slots from CSV"""
    try:
        reserved_slots_path = os.path.join('tt data', 'reserved_slots.csv')
        if not os.path.exists(reserved_slots_path):
            print("Warning: reserved_slots.csv not found")
            return {day: {} for day in config.DAYS}
            
        df = pd.read_csv(reserved_slots_path)
        reserved = {day: {} for day in config.DAYS}
        
        for _, row in df.iterrows():
            day = row['Day']
            start = datetime.strptime(row['Start Time'], '%H:%M').time()
            end = datetime.strptime(row['End Time'], '%H:%M').time()
            department = str(row['Department'])
            
            semesters = []
            for s in str(row['Semester']).split(';'):
                s = s.strip()
                if s.isdigit():
                    base_sem = int(s)  
                    semesters.extend([f"{base_sem}A", f"{base_sem}B", str(base_sem)])
                else:
                    semesters.append(s)
            
            key = (department, tuple(semesters))
            if day not in reserved:
                reserved[day] = {}
            if key not in reserved[day]:
                reserved[day][key] = []
                
            reserved[day][key].append((start, end))
        return reserved
    except Exception as e:
        print(f"Warning: Error loading reserved slots: {e}")
        return {day: {} for day in config.DAYS}

def load_faculty_preferences():
    """Load faculty scheduling preferences"""
    preferences = {}
    try:
        with open('tt data/FACULTY.csv', 'r') as f:
            reader = csv.DictReader(f)
            for row in reader:
                preferred_days = [d.strip() for d in row['Preferred Days'].split(';')] if row['Preferred Days'] else []
                preferred_times = []
                if row['Preferred Times']:
                    time_ranges = row['Preferred Times'].split(';')
                    for time_range in time_ranges:
                        start, end = time_range.split('-')
                        start_time = datetime.strptime(start.strip(), '%H:%M').time()
                        end_time = datetime.strptime(end.strip(), '%H:%M').time()
                        preferred_times.append((start_time, end_time))
                
                preferences[row['Name']] = {
                    'preferred_days': preferred_days,
                    'preferred_times': preferred_times
                }
    except FileNotFoundError:
        print("Warning: FACULTY.csv not found")
        return {}
    return preferences

def is_basket_course(code):
    """Check if course is part of a basket"""
    return code.startswith('B') and '-' in code

def get_basket_group(code):
    """Get basket group (B1, B2, etc) from course code"""
    if is_basket_course(code):
        return code.split('-')[0]
    return None

def get_basket_group_slots(timetable, day, basket_group):
    """Find existing slots with courses from same basket group"""
    basket_slots = []
    for slot_idx, slot in timetable[day].items():
        code = slot.get('code', '')
        if code and get_basket_group(code) == basket_group:
            basket_slots.append(slot_idx)
    return basket_slots

def is_break_time(slot, semester=None):
    """Check if a time slot falls within break times"""
    start, end = slot
    
    # Morning break: 10:30-11:00
    morning_break = (time(10, 30) <= start < time(11, 0))
    
    # Staggered lunch breaks
    lunch_break = False
    if semester:
        base_sem = int(str(semester)[0])
        if base_sem in config.lunch_breaks:
            lunch_start, lunch_end = config.lunch_breaks[base_sem]
            lunch_break = (lunch_start <= start < lunch_end)
    else:
        lunch_break = any(lunch_start <= start < lunch_end 
                         for lunch_start, lunch_end in config.lunch_breaks.values())
    
    return morning_break or lunch_break

def is_slot_reserved(slot, day, semester, department, reserved_slots):
    """Check if a time slot is reserved"""
    if day not in reserved_slots:
        return False
        
    slot_start, slot_end = slot
    
    for (dept, semesters), slots in reserved_slots[day].items():
        if dept == 'ALL' or dept == department:
            if str(semester) in semesters or any(str(semester).startswith(s) for s in semesters):
                for reserved_start, reserved_end in slots:
                    if (slot_start >= reserved_start and slot_start < reserved_end) or \
                       (slot_end > reserved_start and slot_end <= reserved_end):
                        return True
    return False

def is_preferred_slot(faculty, day, time_slot, faculty_preferences):
    """Check if a time slot is within faculty's preferences"""
    if faculty not in faculty_preferences:
        return True
        
    prefs = faculty_preferences[faculty]
    
    if prefs['preferred_days'] and config.DAYS[day] not in prefs['preferred_days']:
        return False
        
    if prefs['preferred_times']:
        slot_start, slot_end = time_slot
        for pref_start, pref_end in prefs['preferred_times']:
            if (slot_start >= pref_start and slot_end <= pref_end):
                return True
        return False
        
    return True

def select_faculty(faculty_str):
    """Select a faculty from potentially multiple options"""
    if '/' in faculty_str:
        faculty_options = [f.strip() for f in faculty_str.split('/')]
        return faculty_options[0]
    return faculty_str

def calculate_required_slots(course):
    """Calculate required slots based on L, T, P, S values"""
    l = float(course['L']) if pd.notna(course['L']) else 0
    t = int(course['T']) if pd.notna(course['T']) else 0
    p = int(course['P']) if pd.notna(course['P']) else 0
    s = int(course['S']) if pd.notna(course['S']) else 0
    
    if s > 0 and l == 0 and t == 0 and p == 0:
        return 0, 0, 0, 0
    
    lecture_sessions = 0
    if l > 0:
        lecture_sessions = max(1, round(l * 2/3))
    
    tutorial_sessions = t  
    lab_sessions = p // 2
    self_study_sessions = s // 4 if (l > 0 or t > 0 or p > 0) else 0
    
    return lecture_sessions, tutorial_sessions, lab_sessions, self_study_sessions

def get_required_room_type(course):
    """Determine required room type based on course attributes"""
    if pd.notna(course['P']) and course['P'] > 0:
        course_code = str(course['Course Code']).upper()
        if 'CS' in course_code or 'DS' in course_code:
            return 'COMPUTER_LAB'
        elif 'EC' in course_code:
            return 'HARDWARE_LAB'
        return 'COMPUTER_LAB'
    else:
        return 'LECTURE_ROOM'

def get_course_priority(course):
    """Calculate course scheduling priority"""
    priority = 0
    code = str(course['Course Code'])
    
    if pd.notna(course['P']) and course['P'] > 0 and not is_basket_course(code):
        priority += 10
        if 'CS' in code or 'EC' in code:
            priority += 2
    elif is_basket_course(code):
        priority += 1
    elif pd.notna(course['L']) and course['L'] > 2:
        priority += 3
    elif pd.notna(course['T']) and course['T'] > 0:
        priority += 2
    return priority

class UnscheduledComponent:
    """Class to track unscheduled course components"""
    def __init__(self, department, semester, code, name, faculty, component_type, sessions, section='', reason=''):
        self.department = department
        self.semester = semester
        self.code = code
        self.name = name
        self.faculty = faculty 
        self.component_type = component_type
        self.sessions = sessions
        self.section = section
        self.reason = reason
        
    def __eq__(self, other):
        if not isinstance(other, UnscheduledComponent):
            return False
        return (self.department == other.department and
                self.semester == other.semester and
                self.code == other.code and
                self.component_type == other.component_type and
                self.section == other.section)
    
    def __hash__(self):
        return hash((self.department, self.semester, self.code, self.component_type, self.section))