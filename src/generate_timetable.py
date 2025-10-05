import pandas as pd
import random

def generate_timetable():
    # Load course data
    courses = pd.read_csv("data/course_data.csv")

    # Define days
    days = ["Mon", "Tue", "Wed", "Thu", "Fri"]

    # Define precise time slots
    time_slots = [
        ("09:00", "10:00"), ("10:00", "10:30"), ("10:30", "10:45"),
        ("10:45", "11:00"), ("11:00", "11:30"), ("11:30", "12:00"),
        ("12:00", "12:15"), ("12:15", "12:30"), ("12:30", "13:15"),
        ("13:15", "13:30"), ("13:30", "14:00"), ("14:00", "14:30"),
        ("14:30", "15:30"), ("15:30", "15:40"), ("15:40", "16:00"),
        ("16:00", "16:30"), ("16:30", "17:10"), ("17:10", "17:30"),
        ("17:30", "18:30"), ("18:30", "19:00")
    ]

    # Example rooms
    rooms = ["R1", "R2", "R3", "R4", "R5"]

    timetable = []

    # Conflict detection
    def has_conflict(existing_schedule, new_class):
        for cls in existing_schedule:
            if cls['Day'] == new_class['Day']:
                overlap = not (new_class['End-Time'] <= cls['Start-Time'] or new_class['Start-Time'] >= cls['End-Time'])
                if overlap:
                    if cls['Room'] == new_class['Room']:
                        return f"Room conflict with {cls['Course Code']}"
                    if cls['Instructor'] == new_class['Instructor']:
                        return f"Instructor conflict with {cls['Course Code']}"
        return None

    # Schedule courses
    for _, course in courses.iterrows():
        scheduled = False
        random.shuffle(days)
        random.shuffle(time_slots)
        random.shuffle(rooms)

        for day in days:
            for start_time, end_time in time_slots:
                for room in rooms:
                    new_class = {
                        'Course Code': course['Course Code'],
                        'Course-Name': course['Course-Name'],
                        'Instructor': course['Instructor'],
                        'Room': room,
                        'Day': day,
                        'Start-Time': start_time,
                        'End-Time': end_time
                    }
                    if has_conflict(timetable, new_class) is None:
                        timetable.append(new_class)
                        scheduled = True
                        break
                if scheduled:
                    break
            if scheduled:
                break
        if not scheduled:
            print(f"Could not schedule {course['Course Code']} due to conflicts.")

    # Convert to DataFrame
    timetable_df = pd.DataFrame(timetable)

    # Optional: remove Room column if you don’t want it
    # timetable_df = timetable_df.drop(columns=['Room'])

    # Save timetable
    timetable_df.to_csv("generated_timetable.csv", index=False)
    print("Timetable generated successfully!")
    return timetable_df
