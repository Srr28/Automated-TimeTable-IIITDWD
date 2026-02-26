import random

import pandas as pd
import openpyxl
from openpyxl.styles import PatternFill, Alignment
from collections import defaultdict
from pathlib import Path

"""High-level scheduling pipeline.

1. Load raw Excel sheets for courses and rooms.
2. Normalize each course row into schedulable events.
3. Group events into logical blocks, keeping mutually shared courses together.
4. Attempt to place each block on the timetable grid while respecting rooms, staff, and cohorts.
5. Export successful placements and report anything unscheduled.
"""

BASE_DIR = Path(__file__).resolve().parent
INPUT_DIR = BASE_DIR / 'input'
OUTPUT_DIR = BASE_DIR / 'outputs'

# Grid configuration covering a five-day work week with 15-minute slots.
DAYS = ['Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday']
START_TIME = 9.0
END_TIME = 18.5
TIME_STEP = 0.25  # 15 minutes
BREAK_LENGTH = 0.08# 15 minutes
LUNCH_START = 12.0
LUNCH_END = 14.0
MIN_LUNCH_BREAK = 1.0  # hours of free time required within lunch window

RESERVED_WINDOWS = {
	# 'Monday': [(9.0, 10.0)],
	# 'Tuesday': [(15.0, 16.0)],
	# 'Wednesday': [(11.0, 12.0)],
	# 'Thursday': [(14.0, 15.0)],
	# 'Friday': [(10.0, 11.0)]

}


def _rotated_days(offset=0):
	"""Yield days starting from the given offset to spread placements round-robin."""
	for idx in range(len(DAYS)):
		yield DAYS[(offset + idx) % len(DAYS)]



def load_data():
	"""Load and normalize input spreadsheets."""
	try:
		courses_path = INPUT_DIR / 'courses.xlsx'
		rooms_path = INPUT_DIR / 'rooms.xlsx'
		courses_df = pd.read_excel(courses_path)
		rooms_df = pd.read_excel(rooms_path)
	except FileNotFoundError as exc:
		print(f"Error: Missing input file. {exc}")
		raise
	except Exception as exc:
		print(f"Error loading data: {exc}")
		raise

	courses_df.columns = courses_df.columns.str.strip().str.lower()
	rooms_df.columns = rooms_df.columns.str.strip().str.lower()
	return courses_df, rooms_df


# Expand course rows into schedulable L/T/P events.
def preprocess_courses(courses_df):
	"""Convert course rows into schedulable event dictionaries."""
	events = []
	event_counter = 1

	for _, row in courses_df.iterrows():
		# Skip any course rows that are not meant to appear in the timetable.
		if str(row.get('schedule or not', '')).strip().lower() != 'yes':
			continue

		course_code = str(row['course code']).strip()
		course_name = str(row['course name']).strip()
		department = str(row['department']).strip()
		semester = str(row['semester']).strip()
		faculty = str(row['faculty']).strip()
		lab_assistant = str(row.get('lab assistant', '')).strip()
		shared_with = str(row.get('shared with', '')).strip()
		shared_targets = _parse_shared_targets(shared_with)
		lab_type = str(row.get('lab type', '')).strip().lower()
		basket_raw = row.get('basket', '')
		basket_code = str(basket_raw).strip()
		if basket_code.lower() == 'nan':
			basket_code = ''
		else:
			basket_code = basket_code.upper()

		num_students_raw = row.get('number of students', row.get('number or students', 0))
		num_students = int(num_students_raw if pd.notna(num_students_raw) else 0)

		lecture_durations = _parse_duration_list(row.get('l'))
		tutorial_durations = _parse_duration_list(row.get('t'))
		practical_hours = _safe_float(row.get('p'))

		# Emit lecture events using provided durations (e.g., "1.5, 1.5" => two slots).
		for duration in lecture_durations:
			events.append(_build_event(
				event_counter,
				'L',
				duration,
				course_code,
				course_name,
				faculty,
				lab_assistant,
				department,
				semester,
				num_students,
				shared_with,
				shared_targets,
				'',
				basket_code
			))
			event_counter += 1

		# Tutorials now also respect per-session durations from the T column.
		for duration in tutorial_durations:
			events.append(_build_event(
				event_counter,
				'T',
				duration,
				course_code,
				course_name,
				faculty,
				lab_assistant,
				department,
				semester,
				num_students,
				shared_with,
				shared_targets,
				'',
				basket_code
			))
			event_counter += 1

		# Practicals become a single block using the provided hour count.
		if practical_hours > 0:
			events.append(_build_event(
				event_counter,
				'P',
				practical_hours,
				course_code,
				course_name,
				faculty,
				lab_assistant,
				department,
				semester,
				num_students,
				shared_with,
				shared_targets,
				lab_type,
				basket_code
			))
			event_counter += 1

	return events


# Coerce optional numeric into float with zero default.
def _safe_float(value):
	"""Return a numeric value as float, defaulting to 0.0 for blanks or NaN."""
	if pd.notna(value) and value != '':
		return float(value)
	return 0.0


# Coerce optional numeric into int with zero default.
def _safe_int(value):
	"""Return a numeric value as int, defaulting to 0 for blanks or NaN."""
	if pd.notna(value) and value != '':
		return int(value)
	return 0

# Parse comma/semicolon separated durations into float list.
def _parse_duration_list(value):
	"""Parse comma/semicolon-separated durations; fallback to a single numeric."""
	if value is None or value == '' or (isinstance(value, float) and pd.isna(value)):
		return []
	if isinstance(value, str):
		cleaned = value.replace(';', ',')
		parts = [p.strip() for p in cleaned.split(',') if p.strip()]
		durations = []
		for part in parts:
			try:
				dur = float(part)
				if dur > 0:
					durations.append(dur)
			except ValueError:
				continue
		if durations:
			return durations
		try:
			num = float(value)
			return [num] if num > 0 else []
		except ValueError:
			return []
	try:
		num = float(value)
		return [num] if num > 0 else []
	except (TypeError, ValueError):
		return []


# Build normalized event dict for a single session.
def _build_event(counter, event_type, duration, course_code, course_name, faculty,
			 lab_assistant, department, semester, num_students, shared_with, shared_targets, lab_type, basket_code):
	"""Create a normalized event dictionary for a single L/T/P occurrence."""
	return {
		'id': f"{course_code}_{event_type}_{counter}",
		'course_code': course_code,
		'course_name': course_name,
		'faculty': faculty,
		'lab_assistant': lab_assistant,
		'department': department,
		'semester': semester,
		'num_students': num_students,
		'type': event_type,
		'duration_hours': duration,
		'lab_type': lab_type,
		'shared_with': shared_with,
		'shared_targets': shared_targets,
		'basket_code': basket_code
	}


# Normalize shared-with string into tokens.
def _parse_shared_targets(shared_with_value):
	"""Normalize "shared with" text into a list of target department tokens."""
	if not shared_with_value:
		return []
	cleaned = shared_with_value.replace(';', ',')
	tokens = [token.strip() for token in cleaned.split(',') if token.strip()]
	return tokens


# Simplify labels for comparison.
def _normalize_label(value):
	"""Lowercase a label and strip non-alphanumeric characters for comparison."""
	return ''.join(ch for ch in value.lower().strip() if ch.isalnum())


# Apply label normalization across token list.
def _normalize_tokens(values):
	"""Apply label normalization to a list of raw tokens."""
	return [_normalize_label(value) for value in values if value.strip()]


# Pull next unassigned event from a list.
def _next_unassigned_event(event_list, start_index, assigned):
	"""Return the next event from the list that has not yet been assigned to a block."""
	idx = start_index
	while idx < len(event_list):
		event = event_list[idx]
		idx += 1
		if event['id'] not in assigned:
			return event, idx
	return None, idx


# Check whether any event mentions a target department.
def _department_mentions_target(events, target_label):
	"""Check if any event's shared targets mention the supplied department label."""
	normalized_target = _normalize_label(target_label)
	if not normalized_target:
		return False
	for event in events:
		targets = set(_normalize_tokens(event.get('shared_targets', [])))
		if normalized_target in targets:
			return True
	return False


# Detect mutual shared-with relationship between departments.
def _has_mutual_share(events_a, dept_a, events_b, dept_b):
	"""Return True when both departments mutually reference each other in sharing data."""
	return _department_mentions_target(events_a, dept_b) and _department_mentions_target(events_b, dept_a)


# Build connected components of mutually sharing departments.
def _find_shared_components(dept_events):
	"""Build connected components of departments that mutually share a course."""
	departments = list(dept_events.keys())
	adjacency = {dept: set() for dept in departments}
	for i in range(len(departments)):
		for j in range(i + 1, len(departments)):
			dept_a = departments[i]
			dept_b = departments[j]
			if _has_mutual_share(dept_events[dept_a], dept_a, dept_events[dept_b], dept_b):
				adjacency[dept_a].add(dept_b)
				adjacency[dept_b].add(dept_a)

	components = []
	visited = set()
	for dept in departments:
		if dept in visited:
			continue
		stack = [dept]
		component = []
		while stack:
			current = stack.pop()
			if current in visited:
				continue
			visited.add(current)
			component.append(current)
			stack.extend(adjacency[current] - visited)
		components.append(component)
	return components


# Group events into schedulable blocks, honoring mutual sharing.
def build_blocks(events):
	"""Group events into blocks, honoring cross-scheduled pairs."""
	assigned = set()
	blocks = []
	block_counter = 1

	# Group events by (course, type) so we can keep like sessions together.
	events_by_key = defaultdict(list)
	for event in events:
		events_by_key[(event['course_code'], event['type'])].append(event)

	for (course_code, event_type), course_events in events_by_key.items():
		dept_events = defaultdict(list)
		for evt in course_events:
			dept_events[evt['department']].append(evt)

		for dept_list in dept_events.values():
			dept_list.sort(key=lambda e: e['id'])

		# Connected components identify which departments must share time.
		components = _find_shared_components(dept_events)
		for component in components:
			if len(component) <= 1:
				continue
			# Keeps every department mentioned in a mutual "shared with" cluster on the same timeline.
			pointers = {dept: 0 for dept in component}
			while True:
				block_courses = []
				for dept in component:
					event, next_idx = _next_unassigned_event(dept_events[dept], pointers[dept], assigned)
					if not event:
						block_courses = []
						break
					block_courses.append(event)
					pointers[dept] = next_idx
				if not block_courses:
					break
				for evt in block_courses:
					assigned.add(evt['id'])
				blocks.append(_build_block(block_counter, block_courses))
				block_counter += 1

	# Any remaining singletons become their own block.
	for event in events:
		if event['id'] in assigned:
			continue
		blocks.append(_build_block(block_counter, [event]))
		assigned.add(event['id'])
		block_counter += 1

	return blocks


# Split blocks into basket bundles scheduled first and remaining blocks.
def extract_basket_blocks(blocks):
	"""Split blocks into basket bundles (scheduled first) and remaining singles."""
	basket_groups = defaultdict(list)
	remaining = []
	for block in blocks:
		basket_code = (block.get('basket_code') or '').strip()
		if not basket_code or block['type'] not in {'L', 'T'}:
			remaining.append(block)
			continue
		basket_groups[(basket_code, block['type'])].append(block)

	basket_blocks = []
	for (basket_code, block_type), blocks_list in basket_groups.items():
		blocks_list.sort(key=lambda blk: blk['id'])
		pointers = defaultdict(int)  # keyed by course-set signature
		bundle_counter = 1
		aborted = False
		while True:
			members = []
			durations = set()
			lab_types = set()
			# pick next block from each distinct course-set signature
			signatures = set()
			for blk in blocks_list:
				sig = tuple(sorted((c['course_code'], c['department'], c['semester']) for c in blk['courses']))
				signatures.add(sig)
			for sig in sorted(signatures):
				candidates = [blk for blk in blocks_list if tuple(sorted((c['course_code'], c['department'], c['semester']) for c in blk['courses'])) == sig]
				idx = pointers[sig]
				if idx >= len(candidates):
					members = []
					break
				block = candidates[idx]
				members.append(block)
				durations.add(block['duration'])
				lab_types.add(block['lab_type'])
			if not members:
				break
			if len(durations) > 1 or len(lab_types) > 1:
				remaining.extend(blocks_list)
				aborted = True
				break
			basket_blocks.append({
				'id': f"BASKET_{basket_code}_{block_type}_{bundle_counter}",
				'basket_code': basket_code,
				'type': block_type,
				'duration': members[0]['duration'],
				'members': members
			})
			bundle_counter += 1
			for sig in signatures:
				pointers[sig] += 1
		if not aborted:
			for sig in signatures:
				candidates = [blk for blk in blocks_list if tuple(sorted((c['course_code'], c['department'], c['semester']) for c in blk['courses'])) == sig]
				while pointers[sig] < len(candidates):
					remaining.append(candidates[pointers[sig]])
					pointers[sig] += 1

	return basket_blocks, remaining


# Flatten basket bundle for reporting or unscheduled output.
def _summarize_basket_block(bundle):
	"""Flatten a basket bundle into a block-shaped dict for reporting."""
	courses = []
	total_students = 0
	for member in bundle.get('members', []):
		courses.extend(member.get('courses', []))
		total_students += member.get('total_students', 0)
	return {
		'id': bundle.get('id', 'UNPLACED_BASKET'),
		'courses': courses,
		'type': bundle.get('type', 'L'),
		'duration': bundle.get('duration', 0),
		'lab_type': '',
		'total_students': total_students,
		'basket_code': bundle.get('basket_code', '')
	}


# Aggregate one or more events into a single block record.
def _build_block(block_id, courses):
	"""Aggregate one or more events into a schedulable block entry."""
	duration = courses[0]['duration_hours']
	block_type = courses[0]['type']
	lab_type = courses[0]['lab_type']
	total_students = sum(course['num_students'] for course in courses)
	basket_codes = {course.get('basket_code', '').strip() for course in courses if course.get('basket_code')}
	basket_code = basket_codes.pop() if len(basket_codes) == 1 else ''

	return {
		'id': f"BLOCK_{block_id}",
		'courses': courses,
		'type': block_type,
		'duration': duration,
		'lab_type': lab_type,
		'total_students': total_students,
		'basket_code': basket_code
	}


# Clean room definitions and discard unusable rows.
def _normalize_rooms(rooms_df):
	"""Normalize room definitions from the spreadsheet, skipping unusable entries."""
	rooms = []
	for _, row in rooms_df.iterrows():
		code = str(row.get('room code', '')).strip()
		capacity_raw = row.get('room capacity')
		if not code:
			continue
		if pd.isna(capacity_raw) or capacity_raw == '':
			print(f"Skipping room '{code}' due to missing capacity")
			continue
		try:
			capacity = int(float(capacity_raw))
		except (TypeError, ValueError):
			print(f"Skipping room '{code}' due to invalid capacity '{capacity_raw}'")
			continue
		rooms.append({
			'code': code,
			'capacity': capacity,
			'type': str(row.get('room type', '')).strip().lower()
		})
	return rooms

# Initialize or reuse shared busy maps for scheduling.
def _init_schedule_state(state=None):
	"""Return shared busy maps, creating them when state is missing."""
	if state is None:
		return {
			'room_busy': defaultdict(list),
			'faculty_busy': defaultdict(list),
			'group_busy': defaultdict(list),
			'course_day_usage': defaultdict(set),
			'course_slots_by_day': defaultdict(list),  # course_key -> [(day, start, end)]
			'faculty_slots_by_day': defaultdict(list),  # faculty -> [(day, start, end)]
			'group_slots_by_day': defaultdict(list)  # dept_sem -> [(day, start, end)]
		}
	return state


def _has_course_spacing(course_slots_by_day, course_key, day, start, end, min_gap=TIME_STEP):
	"""Ensure the course has a gap before another same-day slot; disallow back-to-back."""
	for c_day, c_start, c_end in course_slots_by_day.get(course_key, []):
		if c_day != day:
			continue
		# overlap already handled elsewhere; here prevent immediate adjacency
		if start < c_end + min_gap and start >= c_end:
			return False
		if c_start < end + min_gap and c_start >= end:
			return False
	return True


def _has_lunch_break_generic(slots_by_day, key, day, start, end, min_break=MIN_LUNCH_BREAK):
	"""Ensure at least min_break hours free within the lunch window after adding a slot for any entity key."""
	intervals = []
	for c_day, c_start, c_end in slots_by_day.get(key, []):
		if c_day != day:
			continue
		intervals.append((c_start, c_end))
	intervals.append((start, end))

	# Collect overlaps with lunch window.
	lunch_intervals = []
	for s, e in intervals:
		if e <= LUNCH_START or s >= LUNCH_END:
			continue
		lunch_intervals.append((max(s, LUNCH_START), min(e, LUNCH_END)))

	if not lunch_intervals:
		return (LUNCH_END - LUNCH_START) >= min_break

	# Merge overlaps and check gaps.
	lunch_intervals.sort()
	merged = []
	for s, e in lunch_intervals:
		if not merged or s > merged[-1][1]:
			merged.append([s, e])
		else:
			merged[-1][1] = max(merged[-1][1], e)

	prev_end = LUNCH_START
	for s, e in merged:
		if s - prev_end >= min_break:
			return True
		prev_end = max(prev_end, e)

	return (LUNCH_END - prev_end) >= min_break


def _has_lunch_break(course_slots_by_day, course_key, day, start, end, min_break=MIN_LUNCH_BREAK):
	"""Course-specific lunch gap helper (wraps generic)."""
	return _has_lunch_break_generic(course_slots_by_day, course_key, day, start, end, min_break=min_break)


	# Produce feasible start times for a given duration.
def generate_start_times(duration):
	"""Produce every feasible start time for a block with the given duration."""
	times = []
	current = START_TIME
	while current + duration <= END_TIME:
		times.append(current)
		current += TIME_STEP
	return times


# Check if a resource has no overlapping busy slots.
def _is_free(busy_map, key, day, start, end):
	"""Return True if the resource has no busy slot overlapping the requested window."""
	for b_day, b_start, b_end in busy_map[key]:
		if b_day != day:
			continue
		if not (end <= b_start or b_end <= start):
			return False
	return True


# Mark a resource busy and optionally add post-slot break.
def _mark_busy(busy_map, key, day, start, end, include_break=False):
	"""Reserve a time window for a resource and optionally append a post-slot break."""
	busy_map[key].append((day, start, end))
	if include_break and end < END_TIME:
		break_end = min(end + BREAK_LENGTH, END_TIME)
		busy_map[key].append((day, end, break_end))


# Determine if a slot overlaps the fixed lunch window.
def _conflicts_with_lunch(start, end):
	"""Lunch overlap is allowed; spacing handled elsewhere."""
	return False


# Determine if a slot overlaps any reserved window.
def _conflicts_with_reserved_window(day, start, end):
	"""Return True if the slot intersects any institute-reserved window for the day."""
	for win_start, win_end in RESERVED_WINDOWS.get(day, []):
		if not (end <= win_start or start >= win_end):
			return True
	return False


# Pick the smallest-capacity room that satisfies capacity/type and is free.
def _select_single_room(block, rooms, day, start, end, room_busy, exclude_codes=None):
	"""Pick the tightest-fitting single room for the block (min capacity >= students).

	Policy now: lectures and tutorials never use labs; practicals unchanged.
	"""

	def _is_lab(room_type):
		return 'lab' in (room_type or '')

	eligible = []
	for room in rooms:
		if exclude_codes and room['code'] in exclude_codes:
			continue
		if room['capacity'] < block['total_students']:
			continue

		room_type = room['type']
		room_is_lab = _is_lab(room_type)

		if block['type'] == 'P':
			# Practicals: must be in a lab. If lab_type given, enforce match; otherwise any lab.
			if not room_is_lab:
				continue
			if block['lab_type'] and room_type != block['lab_type']:
				continue
		else:
			# L/T: avoid labs and non-classroom types
			if room_is_lab:
				continue
			if room_type not in ('classroom', '', None):
				continue

		if not _is_free(room_busy, room['code'], day, start, end):
			continue
		eligible.append(room)

	if not eligible:
		return None

	eligible.sort(key=lambda r: (r['capacity'], r['code']))
	return [eligible[0]]


# Combine multiple labs to satisfy a practical block.
def _select_lab_rooms(block, rooms, day, start, end, room_busy):
	"""Group multiple labs so combined capacity can host the block."""
	eligible = []
	for room in rooms:
		room_type = room['type']
		if 'lab' not in (room_type or ''):
			continue
		if block['lab_type'] and room_type != block['lab_type']:
			continue
		if not _is_free(room_busy, room['code'], day, start, end):
			continue
		eligible.append(room)

	if not eligible:
		return None

	eligible.sort(key=lambda r: r['capacity'], reverse=True)
	selection = []
	remaining = block['total_students']
	for room in eligible:
		selection.append(room)
		remaining -= room['capacity']
		if remaining <= 0:
			break

	if remaining > 0:
		return None
	return selection


def schedule_basket_blocks(basket_blocks, rooms_df, state=None):
	"""Place basket bundles before the general scheduler runs."""
	if not basket_blocks:
		return [], [], [], _init_schedule_state(state)

	rooms = _normalize_rooms(rooms_df)
	state = _init_schedule_state(state)
	room_busy = state['room_busy']
	faculty_busy = state['faculty_busy']
	group_busy = state['group_busy']
	course_day_usage = state['course_day_usage']
	course_slots_by_day = state['course_slots_by_day']
	faculty_slots_by_day = state['faculty_slots_by_day']
	group_slots_by_day = state['group_slots_by_day']
	faculty_slots_by_day = state['faculty_slots_by_day']
	group_slots_by_day = state['group_slots_by_day']
	faculty_slots_by_day = state['faculty_slots_by_day']
	group_slots_by_day = state['group_slots_by_day']
	faculty_slots_by_day = state['faculty_slots_by_day']
	group_slots_by_day = state['group_slots_by_day']

	assignments = []
	course_slots = []
	unscheduled = []
	blocks_sorted = sorted(
		basket_blocks,
		key=lambda b: (len(b['members']), b['duration']),
		reverse=True
	)

	for bundle_idx, bundle in enumerate(blocks_sorted):
		day_order = DAYS[bundle_idx % len(DAYS):] + DAYS[:bundle_idx % len(DAYS)]
		members = bundle['members']
		start_times = generate_start_times(bundle['duration'])
		placed = False
		day_available = False
		time_available = False
		rooms_issue = False
		faculty_issue = False
		group_issue = False
		course_spacing_issue = False
		lunch_gap_issue = False

		for day in day_order:
			if any(
				day in course_day_usage[(course['course_code'], course['department'], course['semester'])]
				for member in members
				for course in member['courses']
			):
				continue
			day_available = True
			for start in start_times:
				end = start + bundle['duration']
				if _conflicts_with_lunch(start, end) or _conflicts_with_reserved_window(day, start, end):
					continue
				time_available = True
				used_rooms = set()
				allocations = {}
				allocation_failed = False
				for member in members:
					allocation = _select_single_room(member, rooms, day, start, end, room_busy, exclude_codes=used_rooms)
					if allocation is None:
						rooms_issue = True
						allocation_failed = True
						break
					room_codes = [room['code'] for room in allocation]
					if used_rooms.intersection(room_codes):
						allocation_failed = True
						rooms_issue = True
						break
					allocations[member['id']] = allocation
					used_rooms.update(room_codes)
				if allocation_failed:
					continue

				faculty_ok = all(
					_is_free(faculty_busy, course['faculty'], day, start, end)
					for member in members
					for course in member['courses']
				)
				if not faculty_ok:
					faculty_issue = True
					continue

				groups_ok = all(
					_is_free(group_busy, f"{course['department']}_{course['semester']}", day, start, end)
					for member in members
					for course in member['courses']
				)
				if not groups_ok:
					group_issue = True
					continue

				# spacing and lunch-gap checks for all courses in members
				spacing_ok = True
				lunch_ok = True
				for member in members:
					for course in member['courses']:
						course_key = (course['course_code'], course['department'], course['semester'])
						group_key = f"{course['department']}_{course['semester']}"
						if not _has_course_spacing(course_slots_by_day, course_key, day, start, end):
							spacing_ok = False
							break
						if not _has_lunch_break_generic(course_slots_by_day, course_key, day, start, end, min_break=MIN_LUNCH_BREAK):
							lunch_ok = False
							break
						if not _has_lunch_break_generic(faculty_slots_by_day, course['faculty'], day, start, end, min_break=MIN_LUNCH_BREAK):
							lunch_ok = False
							break
						if not _has_lunch_break_generic(group_slots_by_day, group_key, day, start, end, min_break=MIN_LUNCH_BREAK):
							lunch_ok = False
							break
					if not spacing_ok or not lunch_ok:
						break
				if not spacing_ok:
					course_spacing_issue = True
					continue
				if not lunch_ok:
					lunch_gap_issue = True
					continue

				for member in members:
					rooms_for_member = [room['code'] for room in allocations[member['id']]]
					assignments.append({
						'block_id': member['id'],
						'day': day,
						'start': start,
						'end': end,
						'rooms': rooms_for_member,
						'type': member['type'],
						'courses': member['courses'],
						'basket_code': bundle.get('basket_code', '')
					})
					for room in allocations[member['id']]:
						_mark_busy(room_busy, room['code'], day, start, end, include_break=True)
					for course in member['courses']:
						_mark_busy(faculty_busy, course['faculty'], day, start, end, include_break=True)
						group_key = f"{course['department']}_{course['semester']}"
						_mark_busy(group_busy, group_key, day, start, end, include_break=True)
						course_day_usage[(course['course_code'], course['department'], course['semester'])].add(day)
						course_slots_by_day[(course['course_code'], course['department'], course['semester'])].append((day, start, end))
						faculty_slots_by_day[course['faculty']].append((day, start, end))
						group_slots_by_day[group_key].append((day, start, end))
						course_slots.append({
							'course_code': course['course_code'],
							'course_name': course['course_name'],
							'faculty': course['faculty'],
							'department': course['department'],
							'semester': course['semester'],
							'rooms': rooms_for_member,
							'day': day,
							'start': start,
							'end': end,
							'type': member['type'],
							'basket_code': bundle.get('basket_code', '')
						})

				placed = True
				break
			if placed:
				break

		if not placed:
			if not day_available:
				reason = "Basket members already occupy every day"
			elif not time_available:
				reason = "No valid day/time window for basket"
			else:
				conflicts = []
				if rooms_issue:
					conflicts.append("insufficient distinct rooms")
				if faculty_issue:
					conflicts.append("faculty busy")
				if group_issue:
					conflicts.append("department/semester busy")
				if course_spacing_issue:
					conflicts.append("requires gap between same-course slots on a day")
				if lunch_gap_issue:
					conflicts.append("requires 1h free during lunch window")
				reason = '; '.join(conflicts) if conflicts else "Requirement conflicts prevented scheduling"
			unscheduled.append({'block': _summarize_basket_block(bundle), 'reason': reason})

	return assignments, course_slots, unscheduled, state
def schedule_blocks(blocks, rooms_df, state=None, allow_same_day_repeat=False):
	"""Assign blocks to days/rooms while tracking busy maps and unscheduled causes."""
	rooms = _normalize_rooms(rooms_df)

	# Schedule the most constrained blocks first (more courses, students, longer duration).
	blocks_sorted = sorted(
		blocks,
		key=lambda b: (len(b['courses']), b['total_students'], b['duration']),
		reverse=True
	)

	state = _init_schedule_state(state)
	room_busy = state['room_busy']
	faculty_busy = state['faculty_busy']
	group_busy = state['group_busy']
	course_day_usage = state['course_day_usage']
	course_slots_by_day = state['course_slots_by_day']
	faculty_slots_by_day = state['faculty_slots_by_day']
	group_slots_by_day = state['group_slots_by_day']

	assignments = []
	course_slots = []
	unscheduled = []

	# Iterate through blocks in priority order and try to place each one.
	for block_idx, block in enumerate(blocks_sorted):
		day_order = DAYS[block_idx % len(DAYS):] + DAYS[:block_idx % len(DAYS)]
		placed = False
		start_times = generate_start_times(block['duration'])
		day_available = False
		time_available = False
		rooms_issue = False
		faculty_issue = False
		group_issue = False
		course_spacing_issue = False
		lunch_gap_issue = False

		# Try each day/time window until we find a slot that satisfies rooms, faculty, and cohorts.
		for day in day_order:
			if not allow_same_day_repeat and any(
				day in course_day_usage[(course['course_code'], course['department'], course['semester'])]
				for course in block['courses']
			):
				continue
			day_available = True
			if placed:
				break
			for start in start_times:
				end = start + block['duration']
				if _conflicts_with_lunch(start, end):
					continue
				if _conflicts_with_reserved_window(day, start, end):
					continue
				time_available = True
				if placed:
					break
				allocation = None
				if block['type'] == 'P':
					# Practical blocks first attempt to reserve one large lab, then fall back to combining smaller labs.
					allocation = _select_single_room(block, rooms, day, start, end, room_busy)
					if allocation is None:
						allocation = _select_lab_rooms(block, rooms, day, start, end, room_busy)
				else:
					allocation = _select_single_room(block, rooms, day, start, end, room_busy)

				if allocation is None:
					rooms_issue = True
					continue

				faculty_ok = all(
					_is_free(faculty_busy, course['faculty'], day, start, end)
					for course in block['courses']
				)
				if not faculty_ok:
					faculty_issue = True
					continue

				groups_ok = all(
					_is_free(group_busy, f"{course['department']}_{course['semester']}", day, start, end)
					for course in block['courses']
				)
				if not groups_ok:
					group_issue = True
					continue

				# Prevent back-to-back same-course slots and ensure 1h lunch gap.
				spacing_ok = True
				lunch_ok = True
				for course in block['courses']:
					course_key = (course['course_code'], course['department'], course['semester'])
					group_key = f"{course['department']}_{course['semester']}"
					if not _has_course_spacing(course_slots_by_day, course_key, day, start, end):
						spacing_ok = False
						break
					if not _has_lunch_break_generic(course_slots_by_day, course_key, day, start, end, min_break=MIN_LUNCH_BREAK):
						lunch_ok = False
						break
					if not _has_lunch_break_generic(faculty_slots_by_day, course['faculty'], day, start, end, min_break=MIN_LUNCH_BREAK):
						lunch_ok = False
						break
					if not _has_lunch_break_generic(group_slots_by_day, group_key, day, start, end, min_break=MIN_LUNCH_BREAK):
						lunch_ok = False
						break
				if not spacing_ok:
					course_spacing_issue = True
					continue
				if not lunch_ok:
					lunch_gap_issue = True
					continue

				room_codes = [room['code'] for room in allocation]
				assignments.append({
					'block_id': block['id'],
					'day': day,
					'start': start,
					'end': end,
					'rooms': room_codes,
					'type': block['type'],
					'courses': block['courses'],
					'basket_code': block.get('basket_code', '')
				})

				for room in allocation:
					_mark_busy(room_busy, room['code'], day, start, end, include_break=True)

				for course in block['courses']:
					_mark_busy(faculty_busy, course['faculty'], day, start, end, include_break=True)
					group_key = f"{course['department']}_{course['semester']}"
					_mark_busy(group_busy, group_key, day, start, end, include_break=True)
					course_day_usage[(course['course_code'], course['department'], course['semester'])].add(day)
					course_slots_by_day[(course['course_code'], course['department'], course['semester'])].append((day, start, end))
					faculty_slots_by_day[course['faculty']].append((day, start, end))
					group_slots_by_day[group_key].append((day, start, end))
					course_slots.append({
						'course_code': course['course_code'],
						'course_name': course['course_name'],
						'faculty': course['faculty'],
						'department': course['department'],
						'semester': course['semester'],
						'rooms': room_codes,
						'day': day,
						'start': start,
						'end': end,
						'type': block['type'],
						'basket_code': block.get('basket_code', '')
					})

				placed = True
				break

		if not placed:
			if not day_available:
				reason = "Course already has a slot on every day"
			elif not time_available:
				reason = "No valid day/time window after lunch or reserved blocks"
			else:
				conflicts = []
				if rooms_issue:
					conflicts.append("no suitable room/lab available")
				if faculty_issue:
					conflicts.append("faculty busy")
				if group_issue:
					conflicts.append("department/semester busy")
				if course_spacing_issue:
					conflicts.append("requires gap between same-course slots on a day")
				if lunch_gap_issue:
					conflicts.append("requires 1h free during lunch window")
				reason = "; ".join(conflicts) if conflicts else "Requirement conflicts prevented scheduling"
			unscheduled.append({'block': block, 'reason': reason})

	state = {
		'room_busy': room_busy,
		'faculty_busy': faculty_busy,
		'group_busy': group_busy,
		'course_day_usage': course_day_usage,
		'course_slots_by_day': course_slots_by_day,
		'faculty_slots_by_day': faculty_slots_by_day,
		'group_slots_by_day': group_slots_by_day
	}

	return assignments, course_slots, unscheduled, state


def _retry_unscheduled(unscheduled_entries, rooms_df, state, allow_same_day_repeat, label, randomize_order=False):
	"""Attempt to place unscheduled blocks again, optionally shuffling order."""
	if not unscheduled_entries:
		return [], [], [], state
	retry_blocks = [entry['block'] for entry in unscheduled_entries]
	if randomize_order:
		random.shuffle(retry_blocks)
	print(f"{label}: attempting to place {len(retry_blocks)} pending blocks...")
	retry_assignments, retry_slots, retry_unscheduled, state = schedule_blocks(
		retry_blocks,
		rooms_df,
		state=state,
		allow_same_day_repeat=allow_same_day_repeat
	)
	if retry_assignments:
		print(f"{label}: recovered {len(retry_assignments)} blocks.")
	if retry_unscheduled:
		print(f"{label}: {len(retry_unscheduled)} blocks remain unscheduled.")
	else:
		print(f"{label}: all pending blocks scheduled.")
	return retry_assignments, retry_slots, retry_unscheduled, state


def time_to_str(time_float):
	"""Format a decimal hour (e.g., 9.5) as an HH:MM string."""
	hours = int(time_float)
	minutes = int(round((time_float - hours) * 60))
	return f"{hours:02d}:{minutes:02d}"


def get_time_columns():
	"""Generate the ordered list of timetable column labels."""
	columns = []
	current = START_TIME
	while current < END_TIME:
		columns.append(time_to_str(current))
		current += TIME_STEP
	return columns


def export_timetables(course_slots):
	"""Create per-department, faculty, and room Excel timetables from scheduled slots."""
	if not course_slots:
		print("No scheduled slots to export.")
		return

	OUTPUT_DIR.mkdir(parents=True, exist_ok=True)

	time_cols = get_time_columns()
	colors = {
		'L': PatternFill(start_color='ADD8E6', end_color='ADD8E6', fill_type='solid'),
		'T': PatternFill(start_color='90EE90', end_color='90EE90', fill_type='solid'),
		'P': PatternFill(start_color='FFFFE0', end_color='FFFFE0', fill_type='solid'),
		'LUNCH': PatternFill(start_color='D3D3D3', end_color='D3D3D3', fill_type='solid')
	}

	dept_sched = defaultdict(lambda: {day: {col: None for col in time_cols} for day in DAYS})
	faculty_sched = defaultdict(lambda: {day: {col: None for col in time_cols} for day in DAYS})
	room_sched = defaultdict(lambda: {day: {col: None for col in time_cols} for day in DAYS})
	dept_baskets = defaultdict(lambda: defaultdict(set))  # dept_key -> basket_code -> {(course display, rooms)}

	for slot in course_slots:
		room_codes = slot.get('rooms') or ([slot['room']] if slot.get('room') else [])
		room_display = ', '.join(room_codes)
		basket_code = slot.get('basket_code', '') or ''
		basket_label = f"BASKET {basket_code}" if basket_code else ""
		course_label = f"{slot['course_code']} | {slot['course_name']} | {slot['faculty']} | {room_display}"
		# For room sheets, include both basket code and course details so room timetables stay informative.
		room_label = f"{basket_label} | {course_label}" if basket_code else course_label
		dept_room_label = basket_label if basket_code else course_label
		faculty_label = f"{basket_label} | {course_label}" if basket_code else course_label
		day = slot['day']
		dept_key = f"{slot['department']}_{slot['semester']}"
		if basket_code:
			course_label = f"{slot['course_code']} - {slot['course_name']} ({slot['faculty']})"
			dept_baskets[dept_key][basket_code].add((course_label, room_display))
		span_slots = max(1, int(round((slot['end'] - slot['start']) / TIME_STEP)))
		start_col = time_to_str(slot['start'])
		dept_sched[dept_key][day][start_col] = (dept_room_label, slot['type'], span_slots)
		faculty_sched[slot['faculty']][day][start_col] = (faculty_label, slot['type'], span_slots)
		for room_code in room_codes:
			room_sched[room_code][day][start_col] = (room_label, slot['type'], span_slots)

	# Add lunch break blocks across all generated schedules without overwriting existing entries.
	# Mark free lunch segments as merged cells; do not overwrite existing classes.
	def _mark_lunch(schedule_map):
		for schedule in schedule_map.values():
			for day in DAYS:
				current = LUNCH_START
				while current < LUNCH_END:
					col = time_to_str(current)
					if schedule[day].get(col):
						current += TIME_STEP
						continue
					# start of a free lunch stretch
					start_free = current
					while current < LUNCH_END:
						col_inner = time_to_str(current)
						if schedule[day].get(col_inner):
							break
						current += TIME_STEP
					end_free = current
					span_slots = max(1, int(round((end_free - start_free) / TIME_STEP)))
					start_col = time_to_str(start_free)
					# mark all slots in the stretch so we don't double-insert adjacent blocks
					tmp = start_free
					while tmp < end_free:
						schedule[day][time_to_str(tmp)] = ('LUNCH BREAK', 'LUNCH', 0)
						tmp += TIME_STEP
					# set merged entry at the start
					schedule[day][start_col] = ('LUNCH BREAK', 'LUNCH', span_slots)

	_mark_lunch(dept_sched)
	_mark_lunch(faculty_sched)
	_mark_lunch(room_sched)

	_write_department_workbooks(dept_sched, dept_baskets, time_cols, colors, 'department_timetables')
	_write_individual_workbooks(faculty_sched, time_cols, colors, 'faculty_timetables')
	_write_individual_workbooks(room_sched, time_cols, colors, 'room_timetables')
	print("Exported timetables to per-entity Excel files inside dedicated folders.")


def _write_department_workbooks(schedule_map, basket_map, time_cols, colors, folder_name):
	"""Write department workbooks with timetable + baskets sheet."""
	folder_path = OUTPUT_DIR / folder_name
	folder_path.mkdir(parents=True, exist_ok=True)

	for key, schedule in schedule_map.items():
		wb = openpyxl.Workbook()
		ws = wb.active
		ws.title = 'Timetable'
		ws.append(['Day'] + time_cols)
		for day in DAYS:
			row = [day]
			for col in time_cols:
				cell = schedule[day][col]
				row.append(cell[0] if cell else '')
			ws.append(row)
			row_idx = ws.max_row
			ws.cell(row=row_idx, column=1).alignment = Alignment(wrap_text=True)
			for idx, col in enumerate(time_cols, start=2):
				cell_val = schedule[day][col]
				if not cell_val:
					continue
				cell_obj = ws.cell(row=row_idx, column=idx)
				if isinstance(cell_obj, openpyxl.cell.cell.MergedCell):
					# Skip cells already covered by a prior merge.
					continue
				label_text, slot_type, span_slots = cell_val if len(cell_val) == 3 else (cell_val[0], cell_val[1], 1)
				cell_obj.value = label_text
				cell_obj.alignment = Alignment(wrap_text=True, horizontal='center', vertical='center')
				cell_obj.fill = colors.get(slot_type, PatternFill())
				if span_slots > 1:
					end_col = min(idx + span_slots - 1, len(time_cols) + 1)
					ws.merge_cells(start_row=row_idx, start_column=idx, end_row=row_idx, end_column=end_col)

		# Baskets sheet with course and room listing
		bs = wb.create_sheet('Baskets')
		bs.append(['Basket Code', 'Course (Faculty)', 'Rooms'])
		for basket_code, entries in sorted(basket_map.get(key, {}).items()):
			entries_sorted = sorted(entries)
			start_row = bs.max_row + 1
			for idx, (course_display, rooms) in enumerate(entries_sorted):
				code_val = basket_code if idx == 0 else ''
				bs.append([code_val, course_display, rooms])
				row_idx = bs.max_row
				bs.cell(row=row_idx, column=2).alignment = Alignment(wrap_text=True)
				bs.cell(row=row_idx, column=3).alignment = Alignment(wrap_text=True)
				bs.cell(row=row_idx, column=1).alignment = Alignment(wrap_text=True, horizontal='center', vertical='center')
			if len(entries_sorted) > 1:
				end_row = start_row + len(entries_sorted) - 1
				bs.merge_cells(start_row=start_row, start_column=1, end_row=end_row, end_column=1)
				cell_obj = bs.cell(row=start_row, column=1)
				cell_obj.alignment = Alignment(wrap_text=True, horizontal='center', vertical='center')

		filename = folder_path / f"{_safe_filename(key)}.xlsx"
		wb.save(filename)


def _write_individual_workbooks(schedule_map, time_cols, colors, folder_name):
	"""Write per-entity timetables (faculty or room) without basket sheet."""
	folder_path = OUTPUT_DIR / folder_name
	folder_path.mkdir(parents=True, exist_ok=True)

	for key, schedule in schedule_map.items():
		wb = openpyxl.Workbook()
		ws = wb.active
		ws.title = 'Timetable'
		ws.append(['Day'] + time_cols)
		for day in DAYS:
			row = [day]
			for col in time_cols:
				cell = schedule[day][col]
				row.append(cell[0] if cell else '')
			ws.append(row)
			row_idx = ws.max_row
			ws.cell(row=row_idx, column=1).alignment = Alignment(wrap_text=True)
			for idx, col in enumerate(time_cols, start=2):
				cell_val = schedule[day][col]
				if not cell_val:
					continue
				cell_obj = ws.cell(row=row_idx, column=idx)
				if isinstance(cell_obj, openpyxl.cell.cell.MergedCell):
					continue
				label_text, slot_type, span_slots = cell_val if len(cell_val) == 3 else (cell_val[0], cell_val[1], 1)
				cell_obj.value = label_text
				cell_obj.alignment = Alignment(wrap_text=True, horizontal='center', vertical='center')
				cell_obj.fill = colors.get(slot_type, PatternFill())
				if span_slots > 1:
					end_col = min(idx + span_slots - 1, len(time_cols) + 1)
					ws.merge_cells(start_row=row_idx, start_column=idx, end_row=row_idx, end_column=end_col)

		filename = folder_path / f"{_safe_filename(key)}.xlsx"
		wb.save(filename)


def _safe_filename(name):
	"""Sanitize a string so it can safely be used as a filename."""
	sanitized = ''.join(ch if ch.isalnum() or ch in (' ', '_', '-') else '_' for ch in name)
	sanitized = sanitized.strip().replace(' ', '_')
	return sanitized or 'timetable'



def export_unscheduled(entries):
	"""Write a spreadsheet summarizing every block that could not be scheduled."""
	if not entries:
		print("All blocks scheduled successfully.")
		return

	OUTPUT_DIR.mkdir(parents=True, exist_ok=True)

	# Flatten each failed block into a row for the Excel summary.
	rows = []
	for entry in entries:
		block = entry['block']
		reason = entry.get('reason', '')
		course_details = []
		faculties = []
		departments = set()
		semesters = set()
		for course in block['courses']:
			course_details.append(f"{course['course_code']} ({course['course_name']})")
			faculties.append(course['faculty'])
			departments.add(course['department'])
			semesters.add(str(course['semester']))
		rows.append({
			'block_id': block['id'],
			'type': block['type'],
			'duration_hours': block['duration'],
			'total_students': block['total_students'],
			'lab_type': block['lab_type'],
			'courses': '; '.join(course_details),
			'faculties': '; '.join(faculties),
			'departments': '; '.join(sorted(departments)),
			'semesters': '; '.join(sorted(semesters)),
			'reason': reason
		})

	output_file = OUTPUT_DIR / 'unscheduled_blocks.xlsx'
	pd.DataFrame(rows).to_excel(output_file, index=False)
	print(f"Saved {len(entries)} unscheduled blocks to {output_file}")


def main():
	"""Orchestrate the end-to-end scheduling pipeline and handle retries/exports."""
	try:
		print("Loading data...")
		courses_df, rooms_df = load_data()

		print("Preprocessing courses...")
		events = preprocess_courses(courses_df)
		print(f"Generated {len(events)} event blocks.")

		print("Building logical blocks...")
		blocks = build_blocks(events)
		print(f"Prepared {len(blocks)} blocks for scheduling.")
		basket_blocks, remaining_blocks = extract_basket_blocks(blocks)
		print(f"Detected {len(basket_blocks)} basket bundles and {len(remaining_blocks)} standard blocks.")

		assignments = []
		course_slots = []
		final_unscheduled = []

		basket_assignments, basket_slots, basket_unscheduled, state = schedule_basket_blocks(basket_blocks, rooms_df)
		assignments.extend(basket_assignments)
		course_slots.extend(basket_slots)
		final_unscheduled.extend(basket_unscheduled)

		print("Scheduling remaining blocks with break buffers...")
		regular_assignments, regular_slots, unscheduled, state = schedule_blocks(remaining_blocks, rooms_df, state=state)
		assignments.extend(regular_assignments)
		course_slots.extend(regular_slots)

		if unscheduled:
			print(f"Initial pass left {len(unscheduled)} standard blocks unscheduled. Running retries...")
			# Give the leftovers a few extra passes, relaxing constraints if needed.
			for attempt in range(3):
				if not unscheduled:
					break
				label = f"Retry pass #{attempt + 1}"
				randomize_order = (attempt == 2)
				retry_assignments, retry_slots, unscheduled, state = _retry_unscheduled(
					unscheduled,
					rooms_df,
					state,
					allow_same_day_repeat=True,
					label=label,
					randomize_order=randomize_order
				)
				assignments.extend(retry_assignments)
				course_slots.extend(retry_slots)
				if not retry_assignments and unscheduled:
					print(f"{label}: no additional placements possible.")

		final_unscheduled.extend(unscheduled)

		print("Exporting scheduled timetables...")
		export_timetables(course_slots)

		if final_unscheduled:
			export_unscheduled(final_unscheduled)
		else:
			print("All blocks scheduled. No unscheduled blocks to report.")

		print("Done.")
	except Exception as exc:
		print(f"An error occurred: {exc}")
		raise


if __name__ == '__main__':
	main()
