"""
Scheduler module for timetable generation
Contains room allocation and scheduling logic
"""

import random
import pandas as pd
from datetime import time
import config
from utils import (is_basket_course, get_basket_group, get_basket_group_slots,
                   is_break_time, is_slot_reserved, is_preferred_slot)

def find_adjacent_lab_room(room_id, rooms):
    """Find an adjacent lab room based on room numbering"""
    if not room_id:
        return None
    
    current_num = int(''.join(filter(str.isdigit, rooms[room_id]['roomNumber'])))
    current_floor = current_num // 100
    
    for rid, room in rooms.items():
        if rid != room_id and room['type'] == rooms[room_id]['type']:
            room_num = int(''.join(filter(str.isdigit, room['roomNumber'])))
            if room_num // 100 == current_floor and abs(room_num - current_num) == 1:
                return rid
    return None

def try_room_allocation(rooms, course_type, required_capacity, day, start_slot, duration, used_room_ids):
    """Helper function to try allocating rooms of a certain type"""
    for room_id, room in rooms.items():
        if room_id in used_room_ids or room['type'].upper() == 'LIBRARY':
            continue
            
        if course_type in ['LEC', 'TUT', 'SS']:
            if not ('LECTURE_ROOM' in room['type'].upper() or 'SEATER' in room['type'].upper()):
                continue
        elif course_type == 'COMPUTER_LAB' and room['type'].upper() != 'COMPUTER_LAB':
            continue
        elif course_type == 'HARDWARE_LAB' and room['type'].upper() != 'HARDWARE_LAB':
            continue
            
        if course_type not in ['COMPUTER_LAB', 'HARDWARE_LAB'] and room['capacity'] < required_capacity:
            continue

        slots_free = True
        for i in range(duration):
            if start_slot + i in room['schedule'][day]:
                slots_free = False
                break
                
        if slots_free:
            for i in range(duration):
                room['schedule'][day].add(start_slot + i)
            return room_id
                
    return None

def find_suitable_room(course_type, department, semester, day, start_slot, duration, rooms, batch_info, timetable, course_code="", used_rooms=None):
    """Find suitable room(s) considering student numbers"""
    if not rooms:
        return "DEFAULT_ROOM"
    
    required_capacity = 60
    is_basket = is_basket_course(course_code)
    total_students = None
    
    try:
        df = pd.read_csv('combined.csv')
        
        if course_code and not is_basket:
            course_row = df[df['Course Code'] == course_code]
            if not course_row.empty and 'total_students' in course_row.columns:
                try:
                    val = course_row['total_students'].iloc[0]
                    if pd.notna(val) and str(val).isdigit():
                        total_students = int(val)
                except (ValueError, TypeError):
                    pass
        elif is_basket:
            course_row = df[df['Course Code'] == course_code]
            if not course_row.empty and 'total_students' in course_row.columns:
                try:
                    val = course_row['total_students'].iloc[0]
                    if pd.notna(val) and str(val).isdigit():
                        total_students = int(val)
                    else:
                        elective_info = batch_info.get(('ELECTIVE', course_code))
                        if elective_info:
                            total_students = elective_info['section_size']
                except (ValueError, TypeError):
                    elective_info = batch_info.get(('ELECTIVE', course_code))
                    if elective_info:
                        total_students = elective_info['section_size']
        else:
            dept_info = batch_info.get((department, semester))
            if dept_info:
                total_students = dept_info['section_size']
    except Exception as e:
        print(f"Warning: Error getting total_students: {e}")
    
    if total_students:
        required_capacity = total_students
    elif batch_info:
        if is_basket:
            elective_info = batch_info.get(('ELECTIVE', course_code))
            if elective_info:
                required_capacity = elective_info['section_size']
        else:
            dept_info = batch_info.get((department, semester))
            if dept_info:
                required_capacity = dept_info['section_size']

    used_room_ids = set() if used_rooms is None else used_rooms

    # Handle large classes
    if course_type in ['LEC', 'TUT', 'SS'] and required_capacity > 70:
        seater_120_rooms = {rid: room for rid, room in rooms.items() 
                           if 'SEATER_120' in room['type'].upper()}
        
        if required_capacity > 120:
            seater_240_rooms = {rid: room for rid, room in rooms.items() 
                              if 'SEATER_240' in room['type'].upper()}
            
            room_id = try_room_allocation(seater_240_rooms, 'LEC', required_capacity,
                                        day, start_slot, duration, used_room_ids)
            if room_id:
                return room_id
                
        room_id = try_room_allocation(seater_120_rooms, 'LEC', required_capacity,
                                    day, start_slot, duration, used_room_ids)
        if room_id:
            return room_id

    # Handle lab room allocation
    if course_type in ['COMPUTER_LAB', 'HARDWARE_LAB']:
        dept_info = batch_info.get((department, semester))
        if dept_info and dept_info['total'] > 35:
            for room_id, room in rooms.items():
                if room_id in used_room_ids or room['type'].upper() != course_type:
                    continue
                    
                slots_free = True
                for i in range(duration):
                    if start_slot + i in room['schedule'][day]:
                        slots_free = False
                        break
                
                if slots_free:
                    adjacent_room = find_adjacent_lab_room(room_id, rooms)
                    if adjacent_room and adjacent_room not in used_room_ids:
                        adjacent_free = True
                        for i in range(duration):
                            if start_slot + i in rooms[adjacent_room]['schedule'][day]:
                                adjacent_free = False
                                break
                        
                        if adjacent_free:
                            for i in range(duration):
                                room['schedule'][day].add(start_slot + i)
                                rooms[adjacent_room]['schedule'][day].add(start_slot + i)
                            return f"{room_id},{adjacent_room}"
                            
        return try_room_allocation(rooms, course_type, required_capacity, day, start_slot, duration, used_room_ids)

    # Handle lecture and basket course allocation
    if course_type in ['LEC', 'TUT', 'SS'] or is_basket:
        lecture_rooms = {rid: room for rid, room in rooms.items() 
                        if 'LECTURE_ROOM' in room['type'].upper()}
        
        seater_rooms = {rid: room for rid, room in rooms.items()
                       if 'SEATER' in room['type'].upper()}
        
        if is_basket:
            basket_group = get_basket_group(course_code)
            basket_used_rooms = set()
            basket_group_rooms = {}
            
            room_usage = {rid: sum(len(room['schedule'][d]) for d in range(len(config.DAYS))) 
                         for rid, room in rooms.items()}
            
            sorted_lecture_rooms = dict(sorted(lecture_rooms.items(), 
                                             key=lambda x: room_usage[x[0]]))
            sorted_seater_rooms = dict(sorted(seater_rooms.items(),
                                            key=lambda x: room_usage[x[0]]))
            
            for room_dict in [sorted_lecture_rooms, sorted_seater_rooms]:
                for room_id, room in room_dict.items():
                    is_used = False
                    for slot in range(start_slot, start_slot + duration):
                        if slot in rooms[room_id]['schedule'][day]:
                            if slot in timetable[day]:
                                slot_data = timetable[day][slot]
                                if (slot_data['classroom'] == room_id and 
                                    slot_data['type'] is not None):
                                    slot_code = slot_data.get('code', '')
                                    if get_basket_group(slot_code) == basket_group:
                                        basket_group_rooms[slot_code] = room_id
                                    else:
                                        basket_used_rooms.add(room_id)
                            is_used = True
                            break
                    
                    if not is_used and room_id not in basket_used_rooms:
                        if 'capacity' in room and room['capacity'] >= required_capacity:
                            for i in range(duration):
                                room['schedule'][day].add(start_slot + i)
                            return room_id
            
            if course_code in basket_group_rooms:
                return basket_group_rooms[course_code]
            
            room_id = try_room_allocation(lecture_rooms, 'LEC', required_capacity,
                                        day, start_slot, duration, basket_used_rooms)
            
            if not room_id:
                room_id = try_room_allocation(seater_rooms, 'LEC', required_capacity,
                                            day, start_slot, duration, basket_used_rooms)
            
            if room_id:
                basket_group_rooms[course_code] = room_id
            
            return room_id

        room_id = try_room_allocation(lecture_rooms, 'LEC', required_capacity,
                                    day, start_slot, duration, used_room_ids)
        if not room_id:
            room_id = try_room_allocation(seater_rooms, 'LEC', required_capacity,
                                        day, start_slot, duration, used_room_ids)
        return room_id
    
    return try_room_allocation(rooms, course_type, required_capacity,
                             day, start_slot, duration, used_room_ids)

def is_lecture_scheduled(timetable, day, start_slot, end_slot):
    """Check if there's a lecture scheduled in the given time range"""
    for slot in range(start_slot, end_slot):
        if (slot < len(timetable[day]) and 
            timetable[day][slot]['type'] and 
            timetable[day][slot]['type'] in ['LEC', 'LAB', 'TUT']):
            return True
    return False

def check_faculty_daily_components(professor_schedule, faculty, day, department, semester, section, timetable, course_code=None, activity_type=None):
    """Check faculty/course scheduling constraints for the day"""
    component_count = 0
    faculty_courses = set()
    
    for slot in timetable[day].values():
        if slot['faculty'] == faculty and slot['type'] in ['LEC', 'LAB', 'TUT']:
            slot_code = slot.get('code', '')
            if slot_code:
                if not is_basket_course(slot_code):
                    component_count += 1
                elif slot_code not in faculty_courses:
                    component_count += 1
                    faculty_courses.add(slot_code)
    
    if course_code and is_basket_course(course_code):
        basket_group = get_basket_group(course_code)
        existing_slots = get_basket_group_slots(timetable, day, basket_group)
        if existing_slots:
            return component_count < 3
    
    return component_count < 2

def check_faculty_course_gap(professor_schedule, timetable, faculty, course_code, day, start_slot):
    """Check if there is sufficient gap (3 hours) between sessions"""
    min_gap_hours = 3
    slots_per_hour = 2
    required_gap = min_gap_hours * slots_per_hour
    
    for i in range(max(0, start_slot - required_gap), start_slot):
        if i in professor_schedule[faculty][day]:
            slot_data = timetable[day][i]
            if slot_data['code'] == course_code and slot_data['type'] in ['LEC', 'TUT']:
                return False
                
    for i in range(start_slot + 1, min(len(config.TIME_SLOTS), start_slot + required_gap)):
        if i in professor_schedule[faculty][day]:
            slot_data = timetable[day][i]
            if slot_data['code'] == course_code and slot_data['type'] in ['LEC', 'TUT']:
                return False
    
    return True

def get_best_slots(timetable, professor_schedule, faculty, day, duration, reserved_slots, semester, department, faculty_preferences):
    """Find best available consecutive slots in a day"""
    best_slots = []
    preferred_slots = []
    
    for start_slot in range(len(config.TIME_SLOTS) - duration + 1):
        slots_free = True
        for i in range(duration):
            current_slot = start_slot + i
            if duration == config.LAB_DURATION:
                if (current_slot in professor_schedule[faculty][day] or
                    timetable[day][current_slot]['type'] is not None or
                    is_break_time(config.TIME_SLOTS[current_slot], semester) or
                    is_slot_reserved(config.TIME_SLOTS[current_slot], config.DAYS[day], semester, department, reserved_slots)):
                    slots_free = False
                    break
            else:
                if (current_slot in professor_schedule[faculty][day] or
                    (timetable[day][current_slot]['type'] is not None and
                     not is_basket_course(timetable[day][current_slot].get('code', ''))) or
                    is_break_time(config.TIME_SLOTS[current_slot], semester) or 
                    is_slot_reserved(config.TIME_SLOTS[current_slot], config.DAYS[day], semester, department, reserved_slots)):
                    slots_free = False
                    break

        if slots_free:
            if duration == config.LAB_DURATION:
                slot_time = config.TIME_SLOTS[start_slot][0]
                if slot_time < time(12, 30):
                    preferred_slots.append(start_slot)
                else:
                    best_slots.append(start_slot)
            else:
                if is_preferred_slot(faculty, day, config.TIME_SLOTS[start_slot], faculty_preferences):
                    preferred_slots.append(start_slot)
                else:
                    best_slots.append(start_slot)
    
    return preferred_slots + best_slots