import pytest
from generate_timetable1 import time_to_minutes, has_conflict, get_session_counts, schedule_course, generate_color

# -----------------------------
# TEST time_to_minutes
# -----------------------------
@pytest.mark.parametrize("time_str, expected", [
    ("09:30", 570),
    ("14:00", 840),
    ("00:00", 0),
    ("23:59", 1439)
])
def test_time_to_minutes(time_str, expected):
    assert time_to_minutes(time_str) == expected
    print(f"✅ test_time_to_minutes('{time_str}') works correctly")

# -----------------------------
# TEST has_conflict
# -----------------------------
def test_has_conflict_overlap():
    existing = [{'Day':'Mon', 'Start-Time':'09:00', 'End-Time':'10:30','Room':'R1','Instructor':'A','Semester':'3'}]
    new_class = {'Day':'Mon', 'Start-Time':'09:30', 'End-Time':'11:00','Room':'R1','Instructor':'A','Semester':'3'}
    assert has_conflict(existing, new_class) is True
    print("✅ test_has_conflict_overlap works correctly")

def test_has_conflict_no_overlap():
    existing = [{'Day':'Mon', 'Start-Time':'09:00', 'End-Time':'10:30','Room':'R1','Instructor':'A','Semester':'3'}]
    new_class = {'Day':'Tue', 'Start-Time':'09:30', 'End-Time':'11:00','Room':'R2','Instructor':'B','Semester':'3'}
    assert has_conflict(existing, new_class) is False
    print("✅ test_has_conflict_no_overlap works correctly")

# -----------------------------
# TEST get_session_counts
# -----------------------------
def test_get_session_counts():
    course = {'L-T-P-S-C':'3-1-2-0-4'}
    lectures, tutorials, labs = get_session_counts(course)
    assert (lectures, tutorials, labs) == (2, 1, 1)
    print("✅ test_get_session_counts works correctly")

# -----------------------------
# TEST schedule_course
# -----------------------------
def test_schedule_course_basic():
    course = {'Course Code':'CS101','Course-Name':'Data Structures','Instructor':'Prof A','Semester':'3','Branch':'CSE'}
    timetable = []
    day = schedule_course(course, [('09:00','10:30')], " (Lecture)", timetable, type_name="Lecture")
    assert day in ["Mon", "Tue", "Wed", "Thu", "Fri"]
    assert len(timetable) == 1
    print("✅ test_schedule_course_basic works correctly")

# -----------------------------
# TEST generate_color
# -----------------------------
def test_generate_color():
    color = generate_color("CS101")
    assert isinstance(color, str) and len(color) == 6
    print("✅ test_generate_color works correctly")

# -----------------------------
# RUN TESTS
# -----------------------------
if __name__ == "__main__":
    pytest.main(["-v", __file__])
