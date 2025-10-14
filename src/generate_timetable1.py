import pandas as pd
import random
import math
import os
from openpyxl import load_workbook
from openpyxl.styles import PatternFill, Alignment, Font

# ============================================================
# Configuration
# ============================================================
COURSE_FILE = "data/course_data.csv"
OUTPUT_CSV = "output/flat_timetable.csv"
OUTPUT_EXCEL = "output/structured_timetable.xlsx"

days = ["Mon", "Tue", "Wed", "Thu", "Fri"]

LECTURE_SLOTS = [("09:00", "10:30"), ("10:30", "12:00"), ("13:00", "14:30"), ("14:30", "16:00"), ("16:00", "17:30")]
LAB_SLOTS = [("09:00", "11:00"), ("11:00", "13:00"), ("13:00", "15:00"), ("15:00", "17:00")]
TUTORIAL_SLOTS = [("09:00", "10:00"), ("10:00", "11:00"), ("11:00", "12:00"), ("13:00", "14:00"), ("14:00", "15:00"), ("15:00", "16:00"), ("16:00", "17:00")]
ROOMS = ["R1", "R2", "R3", "R4", "Lab1", "Lab2"]

# ============================================================
# Utility Functions
# ============================================================

def parse_ltp(ltp):
    L, T, P, *_ = map(int, ltp.split("-"))
    lectures = math.ceil(L / 1.5) if L > 0 else 0
    tutorials = T
    labs = math.ceil(P / 2) if P > 0 else 0
    return lectures, tutorials, labs

def generate_color(course_code):
    random.seed(hash(course_code) % (2**32))
    r, g, b = [random.randint(100, 220) for _ in range(3)]
    return f"{r:02X}{g:02X}{b:02X}"

def has_conflict(existing, new_class, branch, sem):
    """Prevents overlapping times for same branch/semester, room, or instructor."""
    for c in existing:
        if c['Day'] == new_class['Day']:
            overlap = not (new_class['End-Time'] <= c['Start-Time'] or new_class['Start-Time'] >= c['End-Time'])
            if overlap:
                if c['Room'] == new_class['Room']:
                    return True
                if c['Instructor'] == new_class['Instructor']:
                    return True
                if (c.get('Branch') == branch and c.get('Semester') == sem):
                    return True
    return False

# ============================================================
# Scheduling Logic
# ============================================================

def schedule_course(course, slots, suffix, timetable, valid_rooms):
    random.shuffle(days)
    random.shuffle(slots)
    random.shuffle(valid_rooms)
    for d in days:
        for s, e in slots:
            for room in valid_rooms:
                new_class = {
                    'Course Code': course['Course Code'],
                    'Course-Name': course['Course-Name'] + f" ({suffix})",
                    'Instructor': course['Instructor'],
                    'Room': room,
                    'Day': d,
                    'Start-Time': s,
                    'End-Time': e,
                    'Branch': course['Branch'],
                    'Semester': course['Semester']
                }
                if not has_conflict(timetable, new_class, course['Branch'], course['Semester']):
                    timetable.append(new_class)
                    return True
    return False

def generate_timetable(courses):
    timetable = []
    for _, course in courses.iterrows():
        L, T, P = parse_ltp(course["L-T-P-S-C"])

        if L > 0:
            for _ in range(L):
                schedule_course(course, LECTURE_SLOTS, "Lecture", timetable, [r for r in ROOMS if "Lab" not in r])
        if T > 0:
            for _ in range(T):
                schedule_course(course, TUTORIAL_SLOTS, "Tutorial", timetable, [r for r in ROOMS if "Lab" not in r])
        if P > 0:
            for _ in range(P):
                schedule_course(course, LAB_SLOTS, "Lab", timetable, [r for r in ROOMS if "Lab" in r])

    return timetable

# ============================================================
# Structured Excel Export (Course Code + Name only)
# ============================================================

def export_structured_excel(timetable):
    df = pd.DataFrame(timetable)
    df['Slot'] = df['Start-Time'] + " - " + df['End-Time']
    df['Display'] = df['Course Code'] + "\n" + df['Course-Name']

    structured_df = df.pivot_table(index='Day', columns='Slot', values='Display', aggfunc=lambda x: "\n---\n".join(x))
    structured_df = structured_df.reindex(days)
    structured_df = structured_df[sorted(structured_df.columns)]

    os.makedirs(os.path.dirname(OUTPUT_EXCEL), exist_ok=True)
    structured_df.to_excel(OUTPUT_EXCEL)
    wb = load_workbook(OUTPUT_EXCEL)
    ws = wb.active

    # Apply colors
    colors = {}
    for i in range(2, ws.max_row + 1):
        for j in range(2, ws.max_column + 1):
            cell = ws.cell(row=i, column=j)
            if cell.value:
                cell.alignment = Alignment(wrap_text=True, vertical="center", horizontal="center")
                code = cell.value.split("\n")[0].strip()
                if code not in colors:
                    colors[code] = generate_color(code)
                color = colors[code]
                cell.fill = PatternFill(start_color=color, end_color=color, fill_type="solid")
                cell.font = Font(size=11, bold=True)

    # Adjust row heights
    for i in range(2, ws.max_row + 1):
        ws.row_dimensions[i].height = 60

    wb.save(OUTPUT_EXCEL)
    print(f"✅ Structured Excel saved to {OUTPUT_EXCEL}")

# ============================================================
# Main Script
# ============================================================

if __name__ == "__main__":
    os.makedirs("output", exist_ok=True)
    courses = pd.read_csv(COURSE_FILE)

    timetable = generate_timetable(courses)
    pd.DataFrame(timetable).to_csv(OUTPUT_CSV, index=False)
    print(f"✅ Flat timetable saved to {OUTPUT_CSV}")

    export_structured_excel(timetable)
