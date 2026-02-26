<div align="center">

# Automated TimeTable Generator — IIIT Dharwad

**A constraint-aware scheduling engine that automatically generates conflict-free timetables for departments, faculty, and rooms.**

[![Python](https://img.shields.io/badge/Python-3.8%2B-3776AB?logo=python&logoColor=white)](https://www.python.org/)
[![Pandas](https://img.shields.io/badge/Pandas-Data%20Processing-150458?logo=pandas&logoColor=white)](https://pandas.pydata.org/)
[![OpenPyXL](https://img.shields.io/badge/OpenPyXL-Excel%20Export-217346?logo=microsoftexcel&logoColor=white)](https://openpyxl.readthedocs.io/)
[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](LICENSE)

</div>

---

## Table of Contents

- [Overview](#overview)
- [Features](#features)
- [Project Structure](#project-structure)
- [Prerequisites](#prerequisites)
- [Installation](#installation)
- [Input Format](#input-format)
- [Usage](#usage)
- [Output](#output)
- [Scheduling Algorithm](#scheduling-algorithm)
- [Configuration](#configuration)
- [Contributing](#contributing)

---

## Overview

Managing timetables manually for a multi-department institution is tedious and error-prone. **Automated TimeTable Generator** reads course and room data from Excel spreadsheets, applies a constraint-satisfaction scheduling pipeline, and exports polished, color-coded Excel timetables — all in a single command.

Built specifically for **IIIT Dharwad**, the system handles real-world complexities like shared courses across departments, basket electives, lab-type matching, faculty conflict avoidance, lunch-break enforcement, and multi-retry recovery for hard-to-place blocks.

---

## Features

| Category | Details |
|---|---|
| **Multi-view Timetables** | Generates separate timetables per department/semester, per faculty member, and per room |
| **Shared Course Handling** | Detects mutually shared courses across departments and co-schedules them in the same slot |
| **Basket Electives** | Groups basket courses together and schedules them as unified bundles |
| **Constraint Solver** | Respects room capacity, room type (classroom vs. lab), faculty availability, and cohort conflicts |
| **Lunch Break Enforcement** | Guarantees a minimum 1-hour free window within the 12:00–14:00 lunch period for every entity |
| **Back-to-Back Prevention** | Ensures spacing between consecutive sessions of the same course on a given day |
| **Multi-Lab Combination** | Automatically combines multiple smaller labs when a single lab cannot fit a large practical section |
| **Retry & Recovery** | Runs multiple retry passes (with optional shuffling) to maximize placement of hard-to-schedule blocks |
| **Unscheduled Report** | Exports a detailed Excel report with reasons for any blocks that could not be placed |
| **Color-Coded Excel Output** | Lectures (blue), Tutorials (green), Practicals (yellow), Lunch (grey) — with merged cells for multi-slot sessions |
| **Reserved Windows** | Supports institute-level reserved time blocks (e.g., club hours, meetings) |
| **Configurable Grid** | 5-day week, 09:00–18:30, 15-minute slot granularity — all easily adjustable |

---

## Project Structure

```
Automated-TimeTable-IIITDWD/
├── main.py                          # Core scheduling engine (single-file)
├── input/
│   ├── courses.xlsx                 # Course catalog with L/T/P hours, faculty, sharing info
│   └── rooms.xlsx                   # Room definitions with capacity and type
├── output/
│   ├── department_timetables/       # Per department-semester Excel files
│   │   ├── CSEA_1.xlsx
│   │   ├── DSAI_3.xlsx
│   │   └── ...
│   ├── faculty_timetables/          # Per faculty member Excel files
│   │   ├── Dr__Animesh_Chaturvedi.xlsx
│   │   └── ...
│   ├── room_timetables/             # Per room Excel files
│   │   ├── C101.xlsx
│   │   ├── L105.xlsx
│   │   └── ...
│   └── unscheduled_blocks.xlsx      # Report of blocks that couldn't be placed
└── README.md
```

---

## Prerequisites

- **Python 3.8+**
- pip (Python package manager)

---

## Installation

1. **Clone the repository**

   ```bash
   git clone https://github.com/Srr28/Automated-TimeTable-IIITDWD.git
   cd Automated-TimeTable-IIITDWD
   ```

2. **Install dependencies**

   ```bash
   pip install pandas openpyxl
   ```

---

## Input Format

### `input/courses.xlsx`

| Column | Description | Example |
|---|---|---|
| `Course Code` | Unique course identifier | `CS201` |
| `Course Name` | Full name of the course | `Data Structures` |
| `Department` | Offering department | `CSEA` |
| `Semester` | Target semester | `3` |
| `Faculty` | Assigned instructor | `Dr. Animesh Chaturvedi` |
| `Lab Assistant` | Optional lab assistant name | `Mr. Kumar` |
| `L` | Lecture durations (comma-separated hours) | `1.5, 1.5` |
| `T` | Tutorial durations (comma-separated hours) | `1` |
| `P` | Practical duration (single value in hours) | `2` |
| `Number of Students` | Enrollment count | `60` |
| `Shared With` | Departments sharing this course | `CSEB, DSAI` |
| `Lab Type` | Required lab type for practicals | `computer lab` |
| `Basket` | Basket elective group code | `BSK1` |
| `Schedule or Not` | Whether to include in scheduling | `Yes` |

### `input/rooms.xlsx`

| Column | Description | Example |
|---|---|---|
| `Room Code` | Room identifier | `C101` |
| `Room Capacity` | Maximum seating capacity | `120` |
| `Room Type` | Type of room | `classroom` / `computer lab` |

---

## Usage

```bash
python main.py
```

The program will:

1. Load course and room data from `input/`
2. Preprocess courses into schedulable Lecture, Tutorial, and Practical events
3. Build logical blocks (honoring cross-department shared courses)
4. Extract and prioritize basket elective bundles
5. Schedule baskets first, then remaining blocks
6. Run up to 3 retry passes for any unscheduled blocks
7. Export all timetables to `output/`

### Sample Console Output

```
Loading data...
Preprocessing courses...
Generated 142 event blocks.
Building logical blocks...
Prepared 128 blocks for scheduling.
Detected 8 basket bundles and 120 standard blocks.
Scheduling remaining blocks with break buffers...
Exporting scheduled timetables...
Exported timetables to per-entity Excel files inside dedicated folders.
All blocks scheduled. No unscheduled blocks to report.
Done.
```

---

## Output

### Department Timetables (`output/department_timetables/`)

Each file (e.g., `CSEA_3.xlsx`) contains:
- **Timetable sheet** — A weekly grid (Mon–Fri, 09:00–18:30) with color-coded, merged cells showing course code, name, faculty, and room
- **Baskets sheet** — A summary of basket elective groupings with course and room details

### Faculty Timetables (`output/faculty_timetables/`)

One file per instructor showing their complete weekly teaching schedule across all departments.

### Room Timetables (`output/room_timetables/`)

One file per room/lab showing occupancy across the week — useful for room utilization analysis.

### Unscheduled Report (`output/unscheduled_blocks.xlsx`)

If any blocks couldn't be placed, this file lists them with:
- Block ID, type, duration, student count
- Course and faculty details
- **Reason** — e.g., *"no suitable room/lab available; faculty busy"*

---

## Scheduling Algorithm

```
┌─────────────────────────────────┐
│     Load courses.xlsx &         │
│        rooms.xlsx               │
└──────────────┬──────────────────┘
               ▼
┌─────────────────────────────────┐
│  Preprocess → L/T/P events      │
│  (parse durations, shared info) │
└──────────────┬──────────────────┘
               ▼
┌─────────────────────────────────┐
│  Build blocks (mutual sharing   │
│  via connected components)      │
└──────────────┬──────────────────┘
               ▼
┌─────────────────────────────────┐
│  Extract basket bundles         │
│  (grouped electives)            │
└──────────────┬──────────────────┘
               ▼
┌─────────────────────────────────┐
│  Schedule baskets first         │
│  (priority placement)           │
└──────────────┬──────────────────┘
               ▼
┌─────────────────────────────────┐
│  Schedule remaining blocks      │
│  (sorted by constraint weight)  │
└──────────────┬──────────────────┘
               ▼
┌─────────────────────────────────┐
│  Retry passes (up to 3x)       │
│  with relaxed constraints       │
└──────────────┬──────────────────┘
               ▼
┌─────────────────────────────────┐
│  Export Excel timetables &      │
│  unscheduled report             │
└─────────────────────────────────┘
```

**Key heuristics:**
- Blocks are sorted by constraint weight (number of courses, student count, duration) — most constrained first
- Day assignment uses round-robin rotation to spread load evenly across the week
- Room selection uses best-fit (smallest sufficient capacity) to minimize waste
- Practicals fall back to multi-lab combination when no single lab is large enough

---

## Configuration

All grid parameters are defined at the top of `main.py` and are easy to adjust:

```python
DAYS = ['Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday']
START_TIME = 9.0        # Day starts at 09:00
END_TIME = 18.5         # Day ends at 18:30
TIME_STEP = 0.25        # 15-minute slot granularity
LUNCH_START = 12.0      # Lunch window start
LUNCH_END = 14.0        # Lunch window end
MIN_LUNCH_BREAK = 1.0   # Minimum free time during lunch (hours)

# Institute-level reserved windows (e.g., club hours)
RESERVED_WINDOWS = {
    # 'Monday': [(9.0, 10.0)],
    # 'Friday': [(15.0, 16.0)],
}
```

---


---

<div align="center">

**Built for IIIT Dharwad** &nbsp;|&nbsp; Made with Python

</div>
