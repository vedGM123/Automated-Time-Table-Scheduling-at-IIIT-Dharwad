import importlib.util
import os
from pathlib import Path
import re


def load_exam_module():
    project_root = Path(__file__).resolve().parents[1]
    module_path = project_root / "src" / "Exam_TT.py"
    spec = importlib.util.spec_from_file_location("exam_tt", str(module_path))
    mod = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(mod)
    return mod


def _extract_course_code(seat_str: str):
    if not seat_str:
        return None
    if "EMPTY" in seat_str:
        return "EMPTY"
    m = re.search(r"\(([^)]+)\)", str(seat_str))
    return m.group(1) if m else None


def test_generate_seating_two_groups_no_same_bench():
    mod = load_exam_module()
    groups = [
        {"course_code": "CSE101", "student_ids": ["S1", "S2", "S3"]},
        {"course_code": "DSAI201", "student_ids": ["T1", "T2"]},
    ]
    df = mod.generate_seating_matrix("R1", capacity=8, student_groups=groups)

    assert not df.empty

    # For each bench (Left bench seats and Right bench seats) ensure two non-empty
    # students on the same bench are from different courses
    for _, row in df.iterrows():
        left_a = _extract_course_code(row["Left Bench - Seat A"])
        left_b = _extract_course_code(row["Left Bench - Seat B"])
        right_a = _extract_course_code(row["Right Bench - Seat A"])
        right_b = _extract_course_code(row["Right Bench - Seat B"])

        if left_a != "EMPTY" and left_b != "EMPTY":
            assert left_a != left_b
        if right_a != "EMPTY" and right_b != "EMPTY":
            assert right_a != right_b


def test_generate_seating_single_group_has_empty_and_seats_all_students():
    mod = load_exam_module()
    groups = [{"course_code": "CSE101", "student_ids": ["S1", "S2", "S3"]}]
    df = mod.generate_seating_matrix("R2", capacity=8, student_groups=groups)

    # There should be EMPTY placeholders (anti-cheat) when only one group exists
    # Ignore the 'Row' label column when counting seats
    seat_values = df.loc[:, df.columns != "Row"].values.flatten().astype(str)
    empty_count = sum(1 for v in seat_values if "EMPTY" in v)
    assert empty_count >= 1

    # Count actual student seats (non-EMPTY entries)
    non_empty_count = sum(1 for v in seat_values if "EMPTY" not in v)
    assert non_empty_count == 3


def test_generate_seating_over_capacity_some_unseated():
    mod = load_exam_module()
    groups = [
        {"course_code": "CSE101", "student_ids": ["S1", "S2", "S3"]},
        {"course_code": "DSAI201", "student_ids": ["T1", "T2", "T3"]},
    ]
    total_students = 6
    df = mod.generate_seating_matrix("R3", capacity=4, student_groups=groups)

    seat_values = df.loc[:, df.columns != "Row"].values.flatten().astype(str)
    non_empty_count = sum(1 for v in seat_values if "EMPTY" not in v)

    # At most capacity seats can be occupied
    assert non_empty_count <= 4
    # And since total_students > capacity, some students should be unseated
    assert non_empty_count < total_students
