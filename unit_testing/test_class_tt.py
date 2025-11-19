import importlib.util
import sys
from types import ModuleType
import pandas as pd
import pytest
from datetime import time
import os

SRC_PATH = os.path.join(os.path.dirname(__file__), '..', 'src', 'Class_TT.py')


def load_class_tt_module(monkeypatch):
    """Load src/Class_TT.py as a module while stubbing pandas.read_csv to return
    small in-memory DataFrames for combined.csv and rooms.csv.
    """
    # Stub for combined.csv
    combined_df = pd.DataFrame([
        {
            'Course Code': 'CS101',
            'Course Name': 'Intro CS',
            'Department': 'CSE',
            'Semester': 1,
            'L': 2,
            'T': 0,
            'P': 0,
            'total_students': 60,
            'Faculty': 'Dr A/Dr B'
        }
    ])

    # Stub for rooms.csv
    rooms_df = pd.DataFrame([
        {'roomNumber': 'R1', 'type': 'LECTURE_ROOM', 'capacity': 80},
        {'roomNumber': 'Lab1', 'type': 'COMPUTER_LAB', 'capacity': 30},
        {'roomNumber': 'C004', 'type': 'SEATER_120', 'capacity': 240}
    ])

    def fake_read_csv(path, *args, **kwargs):
        path = str(path)
        if path.endswith('combined.csv'):
            return combined_df
        if path.endswith('rooms.csv'):
            return rooms_df
        # fallback to real read_csv for other uses
        return pd.read_csv(path, *args, **kwargs)

    monkeypatch.setattr(pd, 'read_csv', fake_read_csv)

    spec = importlib.util.spec_from_file_location('Class_TT', SRC_PATH)
    module = importlib.util.module_from_spec(spec)
    # Ensure it is importable by name if needed
    sys.modules['Class_TT'] = module
    spec.loader.exec_module(module)
    return module


def test_slot_minutes_and_overlaps(monkeypatch):
    m = load_class_tt_module(monkeypatch)
    assert m.slot_minutes((time(9, 0), time(10, 30))) == 90
    assert m.slot_minutes((time(23, 30), time(0, 30))) == 60  # wrap-around
    assert m.overlaps(time(9, 0), time(10, 0), time(9, 30), time(9, 45)) is True
    assert m.overlaps(time(9, 0), time(10, 0), time(10, 0), time(11, 0)) is False


def test_break_and_minor_slots(monkeypatch):
    m = load_class_tt_module(monkeypatch)
    # Lunch break slot exists in TIME_SLOTS
    lunch_slot = next(s for s in m.TIME_SLOTS if s[0] == m.LUNCH_BREAK_START)
    assert m.is_break_time_slot(lunch_slot) is True
    # Early minor (07:30) should be minor
    early = (time(7, 30), time(9, 0))
    assert m.is_minor_slot(early) is True
    # Evening minor (18:30) should be minor
    assert m.is_minor_slot((time(18, 30), time(20, 0))) is True


def test_select_and_split_faculty(monkeypatch):
    m = load_class_tt_module(monkeypatch)
    assert m.select_faculty('Dr X/Dr Y') == 'Dr X'
    assert m.select_faculty('') == 'TBD'

    parts = m.split_faculty_names('Dr X/Dr Y')
    assert parts == ['Dr X', 'Dr Y']
    assert m.split_faculty_names(None) == []


def test_elective_helpers(monkeypatch):
    m = load_class_tt_module(monkeypatch)
    assert m.extract_elective_basket('B1-MA161') == 'B1'
    assert m.extract_elective_basket('MA161') is None
    assert m.get_base_course_code('B2-MA161-C004') == 'MA161-C004'
    assert m.get_base_course_code('MA161') == 'MA161'


def test_parse_cell_for_course():
    # Load module via the helper (keeps tests isolated from package imports)
    m = load_class_tt_module(pytest.MonkeyPatch())
    parse_cell_for_course = m.parse_cell_for_course

    txt = "MA161\nLEC\nRoom: C004\nDr Alice"
    code, typ, room, faculty = parse_cell_for_course(txt)
    assert 'MA161' in code
    assert typ in ('LEC', '') or typ == 'LEC'
    assert room == 'C004'
    assert 'Dr Alice' in faculty

    # Empty or None
    assert parse_cell_for_course(None) == ('', '', '', '')


def test_find_suitable_room_for_slot(monkeypatch):
    m = load_class_tt_module(monkeypatch)
    room_schedule = {}
    course_room_mapping = {}
    # Try to book lecture room for 60 students
    slot_idx = [0]
    room = m.find_suitable_room_for_slot('CS101', 'LECTURE_ROOM', 0, slot_idx, room_schedule, course_room_mapping, 'LEC', 60)
    assert room in ('R1', 'C004')

    # If require huge capacity > any room except C004
    room_schedule2 = {}
    crm2 = {}
    room_big = m.find_suitable_room_for_slot('BIG101', 'LECTURE_ROOM', 0, slot_idx, room_schedule2, crm2, 'LEC', 200)
    assert room_big == 'C004'

    # Laboratory pairing: require 50 in lab, but Lab1 is 30, so no single; function may return combined or None
    room_schedule3 = {}
    crm3 = {}
    lab_room = m.find_suitable_room_for_slot('CSLAB', 'COMPUTER_LAB', 0, slot_idx, room_schedule3, crm3, 'LAB', 50)
    # Either a combined room name (contains +) or None depending on available labs
    assert (lab_room is None) or ('+' in lab_room) or (lab_room in m.ROOM_DATA)
