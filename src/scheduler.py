import pandas as pd
import random
from datetime import timedelta, datetime

def parse_ltpsc(ltpsc):
    """Parse the L-T-P-S-C structure (ignore S for timetable)"""
    l, t, p, s, c = map(int, ltpsc.split('-'))
    return {"L": l, "T": t, "P": p, "C": c}

def generate_time_slots(start_time, end_time, duration_hrs):
    """Generate all valid time slots between start and end times for a given duration"""
    slots = []
    current = start_time
    while current + timedelta(hours=duration_hrs) <= end_time:
        end_slot = current + timedelta(hours=duration_hrs)
        slots.append(f"{current.strftime('%H:%M')}-{end_slot.strftime('%H:%M')}")
        current += timedelta(minutes=30)  # 30-min gap between possible slots
    return slots

def generate_timetable(course_data, room_data):
    days = ["Monday", "Tuesday", "Wednesday", "Thursday", "Friday"]

    # Define allowed hours for each type
    lecture_times = generate_time_slots(datetime.strptime("09:00", "%H:%M"),
                                        datetime.strptime("18:00", "%H:%M"), 1.5)
    tutorial_times = generate_time_slots(datetime.strptime("09:00", "%H:%M"),
                                         datetime.strptime("18:00", "%H:%M"), 1.0)
    lab_times = generate_time_slots(datetime.strptime("09:00", "%H:%M"),
                                    datetime.strptime("18:00", "%H:%M"), 2.0)
    minor_times = generate_time_slots(datetime.strptime("07:30", "%H:%M"),
                                      datetime.strptime("09:00", "%H:%M"), 1.5) + \
                  generate_time_slots(datetime.strptime("18:30", "%H:%M"),
                                      datetime.strptime("20:00", "%H:%M"), 1.5)

    timetable = []

    for _, course in course_data.iterrows():
        ltpsc = parse_ltpsc(course["LTPSC"])
        instructor = course["Instructor"]
        course_name = course["Course Name"]
        credits = ltpsc["C"]

        # Determine if minor or core course
        is_minor = credits <= 2

        available_rooms = room_data["Room Number"].tolist()
        available_times = minor_times if is_minor else lecture_times

        # Lecture Sessions
        for _ in range(ltpsc["L"]):
            timetable.append({
                "Course": course_name,
                "Type": "Lecture",
                "Instructor": instructor,
                "Day": random.choice(days),
                "Time": random.choice(available_times),
                "Room": random.choice(available_rooms)
            })

        # Tutorial Sessions
        for _ in range(ltpsc["T"]):
            timetable.append({
                "Course": course_name,
                "Type": "Tutorial",
                "Instructor": instructor,
                "Day": random.choice(days),
                "Time": random.choice(tutorial_times),
                "Room": random.choice(available_rooms)
            })

        # Lab Sessions
        for _ in range(ltpsc["P"]):
            timetable.append({
                "Course": course_name,
                "Type": "Lab",
                "Instructor": instructor,
                "Day": random.choice(days),
                "Time": random.choice(lab_times),
                "Room": random.choice(available_rooms)
            })

    df = pd.DataFrame(timetable)
    df.sort_values(["Day", "Time"], inplace=True)
    return df
