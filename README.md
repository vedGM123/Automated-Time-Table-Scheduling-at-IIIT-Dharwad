# Automated Time-Table Scheduling at IIIT Dharwad

Generate clash-free class timetables (lectures, tutorials, labs, and elective baskets) using course, room, and faculty data. The output is an Excel workbook for each section plus teacher schedules and a list of unscheduled courses.

**What this tool does**
- Schedules lectures, tutorials, and labs with faculty and room constraints
- Supports elective baskets and shared cross-department courses
- Enforces room capacity and lab/lecture room types
- Produces color-coded Excel timetables, teacher timetables, and an unscheduled list

## Quick Start

**Requirements**
- Python 3.10+ (3.11+ recommended)
- Windows/macOS/Linux

**Install**
```bash
git clone <your-repo-url>
cd Automated-Time-Table-Scheduling-At-IIIT-Dharwad
python -m venv .venv
.venv\Scripts\activate  # Windows
# source .venv/bin/activate  # macOS/Linux
pip install -r requirements.txt
```

**Run**
```bash
python src/Class_TT.py
```

Outputs are written to `output/`:
- `timetable_all_departments.xlsx`
- `teacher_timetables.xlsx`
- `unscheduled_courses.xlsx`

## Data Inputs

All inputs live in `data/`:
- `combined.csv`
  Course list with L/T/P values, faculty, semester, and student counts.
- `rooms.csv`
  Room number, room type (e.g., `LECTURE_ROOM`, `COMPUTER_LAB`), and capacity.
- `config.json`
  Scheduling settings (days, slot durations, etc.).

## Configuration

Edit `data/config.json` to tune the scheduler. Example keys:
- `days` (list of weekdays)
- `LECTURE_MIN`, `TUTORIAL_MIN`, `LAB_MIN` (slot duration in minutes)
- `SELF_STUDY_MIN` (if used)

## Project Structure

```
Automated-Time-Table-Scheduling-At-IIIT-Dharwad/
|-- README.md
|-- requirements.txt
|-- src/
|   |-- Class_TT.py
|-- data/
|   |-- combined.csv
|   |-- rooms.csv
|   |-- config.json
|-- output/
```

## Troubleshooting

- `combined.csv` not found: Ensure `data/combined.csv` exists.
- Rooms not loading: Check `data/rooms.csv` headers and values.
- Many unscheduled courses:
  - Add more rooms or increase room capacities.
  - Relax constraints in `data/config.json`.
  - Extend the available time slots.

## Team

Ved Chandorikar ? 24BCS161  
Sharanprakash R Kasbag ? 24BCS136  
Rangineni Srihith ? 24BCS116  
Shubham Ramesh Vaddar ? 24BCS143  
Guide: Vivekraj V K, Assistant Professor, IIIT Dharwad

## References

1. Asli N Goktug et al., ?A timetable organizer for the planning and implementation of screenings in manual or semi-automation mode,? Journal of Biomolecular Screening, 18:938?942, 2013.
2. Vamsi Krishna Yepuri et al., ?Examination management automation system,? Int. Res. J. Eng Technol, 5:2773?2779, 2018.
