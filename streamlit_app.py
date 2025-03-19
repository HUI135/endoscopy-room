import pandas as pd
import openpyxl
import streamlit as st
from io import BytesIO
from collections import defaultdict
import random
import time
import os
from datetime import datetime
from collections import Counter
import numpy as np
import re
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.comments import Comment

# 고유한 시드 생성
random.seed(time.time_ns() ^ int.from_bytes(os.urandom(4), 'big'))

st.header("내시경 스케쥴 방배정 도구", divider='rainbow')
st.write(" ")

# 최대 한계값 입력 UI
st.sidebar.header("최대 배정 한계 설정")
MAX_DUTY = st.sidebar.number_input("1. 최대 당직 합계", min_value=1, value=3, step=1)
MAX_EARLY = st.sidebar.number_input("2. 최대 이른방 합계", min_value=1, value=6, step=1)
MAX_LATE = st.sidebar.number_input("3. 최대 늦은방 합계", min_value=1, value=6, step=1)
MAX_ROOM = st.sidebar.number_input("4. 최대 방별 합계", min_value=1, value=3, step=1)

uploaded_file = st.file_uploader("엑셀 파일을 업로드하세요 (Sheet1과 Sheet2 포함)", type=["xlsx"])

if uploaded_file is not None:
    wb = openpyxl.load_workbook(uploaded_file)
    Sheet1 = wb['Sheet1']
    Sheet2 = wb['Sheet2']

    def extract_data(sheet):
        data = {}
        headers = [cell.value for cell in sheet[1]]
        for row in sheet.iter_rows(min_row=2):
            date_cell = row[0]
            if date_cell.value:
                date = date_cell.value
                personnel = []
                day_of_week = row[1].value if row[1].value else ""
                memo_dict = {}
                for cell in row[2:]:
                    if cell.value and cell.value not in ['월요일', '화요일', '수요일', '목요일', '금요일', '토요일', '일요일']:
                        personnel.append(cell.value)
                    if cell.comment and cell.value:
                        memo_dict[cell.value] = cell.comment.text.strip()
                personnel_with_suffix = []
                name_counts = Counter()
                for name in personnel:
                    name_counts[name] += 1
                    suffix = f"_{name_counts[name]}" if name_counts[name] > 1 else ""
                    personnel_with_suffix.append(f"{name}{suffix}")
                data[date.strftime('%Y-%m-%d')] = {'personnel': personnel_with_suffix, 'original_personnel': personnel, 'day': day_of_week, 'memos': memo_dict, 'headers': headers}
        return data

    Sheet1_data = extract_data(Sheet1)
    Sheet2_data = extract_data(Sheet2)

    def apply_memo_rules(assignment, personnel, memos, fixed_personnel, slots, assigned_counts, personnel_counts, time_groups, assigned_by_time, total_early, total_late, total_duty, total_rooms, ignore_memos=None):
        if ignore_memos is None:
            ignore_memos = set()
        prioritized = []
        all_slots = set(slots)
        for person in personnel:
            original_name = person.split('_')[0]
            if original_name in memos and person not in fixed_personnel and original_name not in ignore_memos:
                rule = memos[original_name]
                if rule in memo_rules:
                    if rule in ['당직 안됨', '오전 당직 안됨', '오후 당직 안됨']:
                        forbidden_slots = memo_rules[rule]
                        allowed_slots = list(all_slots - set(forbidden_slots))
                        prioritized.append((person, allowed_slots))
                    else:
                        prioritized.append((person, memo_rules[rule]))
        remaining_slots = [i for i, x in enumerate(assignment) if x is None]
        memo_assignments = {}
        for person, allowed_slots in prioritized:
            original_name = person.split('_')[0]
            valid_slots = [
                i for i in remaining_slots 
                if slots[i] in allowed_slots 
                and assigned_counts[person] < personnel_counts[person]
                and original_name not in assigned_by_time.get(next(t for t, g in time_groups.items() if slots[i] in g), set())
                and total_early[original_name] < MAX_EARLY
                and total_late[original_name] < MAX_LATE
                and total_duty[original_name] < MAX_DUTY
            ]
            if valid_slots:
                slot_idx = random.choice(valid_slots)
                assignment[slot_idx] = person
                assigned_counts[person] += 1
                memo_assignments.setdefault(slots[slot_idx], Counter())[person] += 1
                remaining_slots.remove(slot_idx)
                for time_group, group in time_groups.items():
                    if slots[slot_idx] in group:
                        assigned_by_time[time_group].add(original_name)
                if slots[slot_idx] in {'8:30(1)_당직', '8:30(2)', '8:30(4)', '8:30(7)'}:
                    total_early[original_name] += 1
                if slots[slot_idx] in {'10:00(9)', '10:00(3)'}:
                    total_late[original_name] += 1
                if slots[slot_idx] in {'8:30(1)_당직', '13:30(3)_당직'}:
                    total_duty[original_name] += 1
                room_num = re.search(r'\((\d+)\)', slots[slot_idx])
                if room_num:
                    total_rooms[room_num.group(1)][original_name] += 1
        return assignment, memo_assignments

    def calculate_stats(assignment, slots, day_of_week):
        early_slots = {'8:30(1)_당직', '8:30(2)', '8:30(4)', '8:30(7)'}
        late_slots = {'10:00(9)', '10:00(3)'}
        duty_slots = {'8:30(1)_당직', '13:30(3)_당직'}
        slot_counts = {slot.replace('_당직', ''): Counter() for slot in time_slots.keys() if slot != '온콜'}
        
        stats = Counter()
        early_count = Counter()
        late_count = Counter()
        duty_count = Counter()
        
        for slot, person in zip(slots, assignment):
            if person:
                original_name = person.split('_')[0]
                stats[original_name] += 1
                if day_of_week != '토요일':
                    if slot in early_slots:
                        early_count[original_name] += 1
                    if slot in late_slots:
                        late_count[original_name] += 1
                    if slot in duty_slots:
                        duty_count[original_name] += 1
                    if slot != '온콜':
                        slot_counts[slot.replace('_당직', '')][original_name] += 1
        
        return stats, early_count, late_count, duty_count, slot_counts

    def count_violations(total_early, total_late, total_duty, total_slots):
        violations = 0
        all_personnel = set(total_early.keys()) | set(total_late.keys()) | set(total_duty.keys()) | set().union(*[total_slots[slot].keys() for slot in total_slots])
        for person in all_personnel:
            if total_early.get(person, 0) > MAX_EARLY:
                violations += total_early.get(person, 0) - MAX_EARLY
            if total_late.get(person, 0) > MAX_LATE:
                violations += total_late.get(person, 0) - MAX_LATE
            if total_duty.get(person, 0) > MAX_DUTY:
                violations += total_duty.get(person, 0) - MAX_DUTY
            for slot in total_slots:
                if total_slots[slot].get(person, 0) > MAX_ROOM:
                    violations += total_slots[slot].get(person, 0) - MAX_ROOM
        return violations

    def random_assign(personnel, slots, fixed_assignments, memos, day_of_week, time_groups, total_stats):
        random.seed(time.time_ns() ^ int.from_bytes(os.urandom(4), 'big'))
        
        max_attempts = 100
        duty_slots = ['8:30(1)_당직', '13:30(3)_당직']
        early_slots = ['8:30(1)_당직', '8:30(2)', '8:30(4)', '8:30(7)']
        late_slots = ['10:00(9)', '10:00(3)']
        
        best_assignment = None
        best_fixed_assignments_record = None
        best_memo_assignments = None
        min_violations = float('inf')
        best_total_early = total_stats['early'].copy()
        best_total_late = total_stats['late'].copy()
        best_total_duty = total_stats['duty'].copy()
        best_total_slots = {slot: total_stats['slots'][slot].copy() for slot in total_stats['slots']}
        best_total_stats = total_stats['total'].copy()

        for attempt in range(max_attempts):
            assignment = [None] * len(slots)
            fixed_personnel = set()
            assigned_counts = Counter()
            personnel_counts = Counter(personnel)
            assigned_by_time = {time_group: set() for time_group in time_groups.keys()}
            fixed_assignments_record = {}
            memo_assignments = {}
            
            total_early = total_stats['early'].copy()
            total_late = total_stats['late'].copy()
            total_duty = total_stats['duty'].copy()
            total_rooms = {str(i): total_stats['rooms'][str(i)].copy() for i in range(1, 13)}
            
            for date, assignments in fixed_assignments.items():
                if date == current_date:
                    for person, fixed_slot in assignments.items():
                        if fixed_slot in slots:
                            slot_idx = slots.index(fixed_slot)
                            original_name = person.split('_')[0]
                            time_group = next(t for t, g in time_groups.items() if fixed_slot in g)
                            if original_name in assigned_by_time[time_group]:
                                st.error(f"{current_date}: {original_name}이(가) {time_group} 시간대에 이미 배정되어 {fixed_slot}에 중복 배치될 수 없습니다.")
                                return assignment, {}, {}
                            assignment[slot_idx] = person
                            fixed_personnel.add(person)
                            assigned_counts[person] += 1
                            fixed_assignments_record.setdefault(fixed_slot, Counter())[person] += 1
                            assigned_by_time[time_group].add(original_name)
                            if fixed_slot in early_slots and day_of_week != '토요일':
                                total_early[original_name] += 1
                            if fixed_slot in late_slots and day_of_week != '토요일':
                                total_late[original_name] += 1
                            if fixed_slot in duty_slots and day_of_week != '토요일':
                                total_duty[original_name] += 1
                            room_num = re.search(r'\((\d+)\)', fixed_slot)
                            if room_num and day_of_week != '토요일':
                                total_rooms[room_num.group(1)][original_name] += 1
            
            all_slots = set(slots)
            prioritized = []
            for person in personnel:
                original_name = person.split('_')[0]
                if original_name in memos and person not in fixed_personnel:
                    rule = memos[original_name]
                    if rule in memo_rules:
                        if rule in ['당직 안됨', '오전 당직 안됨', '오후 당직 안됨']:
                            forbidden_slots = memo_rules[rule]
                            allowed_slots = list(all_slots - set(forbidden_slots))
                            prioritized.append((person, allowed_slots))
                        else:
                            prioritized.append((person, memo_rules[rule]))
            
            remaining_slots = [i for i, x in enumerate(assignment) if x is None]
            for person, allowed_slots in prioritized:
                original_name = person.split('_')[0]
                valid_slots = [
                    i for i in remaining_slots 
                    if slots[i] in allowed_slots 
                    and assigned_counts[person] < personnel_counts[person]
                    and original_name not in assigned_by_time.get(next(t for t, g in time_groups.items() if slots[i] in g), set())
                    and total_early[original_name] < MAX_EARLY
                    and total_late[original_name] < MAX_LATE
                    and total_duty[original_name] < MAX_DUTY
                ]
                if valid_slots:
                    slot_idx = random.choice(valid_slots)
                    assignment[slot_idx] = person
                    assigned_counts[person] += 1
                    memo_assignments.setdefault(slots[slot_idx], Counter())[person] += 1
                    remaining_slots.remove(slot_idx)
                    for time_group, group in time_groups.items():
                        if slots[slot_idx] in group:
                            assigned_by_time[time_group].add(original_name)
                    if slots[slot_idx] in early_slots and day_of_week != '토요일':
                        total_early[original_name] += 1
                    if slots[slot_idx] in late_slots and day_of_week != '토요일':
                        total_late[original_name] += 1
                    if slots[slot_idx] in duty_slots and day_of_week != '토요일':
                        total_duty[original_name] += 1
                    room_num = re.search(r'\((\d+)\)', slots[slot_idx])
                    if room_num and day_of_week != '토요일':
                        total_rooms[room_num.group(1)][original_name] += 1

            available_slots = [i for i, slot in enumerate(slots) if assignment[i] is None]
            personnel_list = [p for p in personnel if assigned_counts[p] < personnel_counts[p]]
            duty_indices = [i for i in available_slots if slots[i] in duty_slots]
            personnel_list = sorted(personnel_list, key=lambda p: total_duty[p.split('_')[0]])
            for slot_idx in duty_indices:
                for person in personnel_list:
                    original_name = person.split('_')[0]
                    time_group = next(t for t, g in time_groups.items() if slots[slot_idx] in g)
                    if (original_name not in assigned_by_time[time_group] and 
                        assigned_counts[person] < personnel_counts[person] and
                        total_duty[original_name] < MAX_DUTY):
                        assignment[slot_idx] = person
                        assigned_counts[person] += 1
                        assigned_by_time[time_group].add(original_name)
                        if day_of_week != '토요일':
                            total_duty[original_name] += 1
                            total_early[original_name] += 1
                        room_num = re.search(r'\((\d+)\)', slots[slot_idx])
                        if room_num and day_of_week != '토요일':
                            total_rooms[room_num.group(1)][original_name] += 1
                        available_slots.remove(slot_idx)
                        personnel_list = [p for p in personnel_list if assigned_counts[p] < personnel_counts[p]]
                        break
            
            early_indices = [i for i in available_slots if slots[i] in early_slots]
            personnel_list = sorted(personnel_list, key=lambda p: total_early[p.split('_')[0]])
            for slot_idx in early_indices:
                for person in personnel_list:
                    original_name = person.split('_')[0]
                    time_group = next(t for t, g in time_groups.items() if slots[slot_idx] in g)
                    if (original_name not in assigned_by_time[time_group] and 
                        assigned_counts[person] < personnel_counts[person] and
                        total_early[original_name] < MAX_EARLY):
                        assignment[slot_idx] = person
                        assigned_counts[person] += 1
                        assigned_by_time[time_group].add(original_name)
                        if day_of_week != '토요일':
                            total_early[original_name] += 1
                        room_num = re.search(r'\((\d+)\)', slots[slot_idx])
                        if room_num and day_of_week != '토요일':
                            total_rooms[room_num.group(1)][original_name] += 1
                        available_slots.remove(slot_idx)
                        personnel_list = [p for p in personnel_list if assigned_counts[p] < personnel_counts[p]]
                        break
            
            late_indices = [i for i in available_slots if slots[i] in late_slots]
            personnel_list = sorted(personnel_list, key=lambda p: total_late[p.split('_')[0]])
            for slot_idx in late_indices:
                for person in personnel_list:
                    original_name = person.split('_')[0]
                    time_group = next(t for t, g in time_groups.items() if slots[slot_idx] in g)
                    if (original_name not in assigned_by_time[time_group] and 
                        assigned_counts[person] < personnel_counts[person] and
                        total_late[original_name] < MAX_LATE):
                        assignment[slot_idx] = person
                        assigned_counts[person] += 1
                        assigned_by_time[time_group].add(original_name)
                        if day_of_week != '토요일':
                            total_late[original_name] += 1
                        room_num = re.search(r'\((\d+)\)', slots[slot_idx])
                        if room_num and day_of_week != '토요일':
                            total_rooms[room_num.group(1)][original_name] += 1
                        available_slots.remove(slot_idx)
                        personnel_list = [p for p in personnel_list if assigned_counts[p] < personnel_counts[p]]
                        break
            
            available_slots = [i for i, slot in enumerate(slots) if assignment[i] is None]
            personnel_list = [p for p in personnel if assigned_counts[p] < personnel_counts[p]]
            random.shuffle(personnel_list)
            assignment, available_slots = assign_remaining(assignment, personnel_list, available_slots, slots, assigned_counts, personnel_counts, time_groups, assigned_by_time, total_early, total_late, total_duty, total_rooms, MAX_EARLY, MAX_LATE, MAX_DUTY, MAX_ROOM, day_of_week)
            
            if available_slots:
                personnel_list = sorted(
                    personnel_list,
                    key=lambda p: (total_duty[p.split('_')[0]], total_early[p.split('_')[0]], total_late[p.split('_')[0]], sum(total_rooms[r][p.split('_')[0]] for r in total_rooms))
                )
                for slot_idx in available_slots:
                    for person in personnel_list:
                        original_name = person.split('_')[0]
                        time_group = next(t for t, g in time_groups.items() if slots[slot_idx] in g)
                        if (assigned_counts[person] < personnel_counts[person] and 
                            original_name not in assigned_by_time[time_group]):
                            assignment[slot_idx] = person
                            assigned_counts[person] += 1
                            assigned_by_time[time_group].add(original_name)
                            if slots[slot_idx] in early_slots and day_of_week != '토요일':
                                total_early[original_name] += 1
                            if slots[slot_idx] in late_slots and day_of_week != '토요일':
                                total_late[original_name] += 1
                            if slots[slot_idx] in duty_slots and day_of_week != '토요일':
                                total_duty[original_name] += 1
                            room_num = re.search(r'\((\d+)\)', slots[slot_idx])
                            if room_num and day_of_week != '토요일':
                                total_rooms[room_num.group(1)][original_name] += 1
                            available_slots.remove(slot_idx)
                            break
            
            if None not in assignment:
                stats, early_count, late_count, duty_count, slot_counts = calculate_stats(assignment, slots, day_of_week)
                temp_total_early = total_stats['early'].copy()
                temp_total_late = total_stats['late'].copy()
                temp_total_duty = total_stats['duty'].copy()
                temp_total_slots = {slot: total_stats['slots'][slot].copy() for slot in total_stats['slots']}
                temp_total_stats = total_stats['total'].copy()

                temp_total_early.update(early_count)
                temp_total_late.update(late_count)
                temp_total_duty.update(duty_count)
                for slot in slot_counts:
                    temp_total_slots[slot].update(slot_counts[slot])
                temp_total_stats.update(stats)

                violations = count_violations(temp_total_early, temp_total_late, temp_total_duty, temp_total_slots)

                if violations < min_violations:
                    min_violations = violations
                    best_assignment = assignment.copy()
                    best_fixed_assignments_record = fixed_assignments_record.copy()
                    best_memo_assignments = memo_assignments.copy()
                    best_total_early = temp_total_early.copy()
                    best_total_late = temp_total_late.copy()
                    best_total_duty = temp_total_duty.copy()
                    best_total_slots = {slot: temp_total_slots[slot].copy() for slot in temp_total_slots}
                    best_total_stats = temp_total_stats.copy()
                    if min_violations == 0:  # 위반이 0이면 더 이상 시도하지 않음
                        break

        if best_assignment is not None:
            total_stats['early'] = best_total_early
            total_stats['late'] = best_total_late
            total_stats['duty'] = best_total_duty
            total_stats['slots'] = best_total_slots
            total_stats['total'] = best_total_stats
            return best_assignment, best_fixed_assignments_record, best_memo_assignments
        
        stats, early_count, late_count, duty_count, slot_counts = calculate_stats(assignment, slots, day_of_week)
        total_stats['early'].update(early_count)
        total_stats['late'].update(late_count)
        total_stats['duty'].update(duty_count)
        for slot in slot_counts:
            total_stats['slots'][slot].update(slot_counts[slot])
        total_stats['total'].update(stats)
        return assignment, fixed_assignments_record, memo_assignments

    def assign_remaining(assignment, personnel_list, available_slots, slots, assigned_counts, personnel_counts, time_groups, assigned_by_time, total_early, total_late, total_duty, total_rooms, MAX_EARLY, MAX_LATE, MAX_DUTY, MAX_ROOM, day_of_week):
        random.shuffle(personnel_list)
        for person in personnel_list:
            if available_slots:
                original_name = person.split('_')[0]
                possible_slots = []
                
                for slot_idx in available_slots:
                    slot = slots[slot_idx]
                    time_group = next(t for t, g in time_groups.items() if slot in g)
                    if original_name not in assigned_by_time[time_group]:
                        early_ok = (total_early[original_name] + (1 if slot in {'8:30(1)_당직', '8:30(2)', '8:30(4)', '8:30(7)'} else 0)) <= MAX_EARLY or day_of_week == '토요일'
                        late_ok = (total_late[original_name] + (1 if slot in {'10:00(9)', '10:00(3)'} else 0)) <= MAX_LATE or day_of_week == '토요일'
                        duty_ok = (total_duty[original_name] + (1 if slot in {'8:30(1)_당직', '13:30(3)_당직'} else 0)) <= MAX_DUTY or day_of_week == '토요일'
                        room_num = re.search(r'\((\d+)\)', slot)
                        room_ok = True
                        if room_num:
                            room = room_num.group(1)
                            room_ok = (total_rooms[room][original_name] + 1) <= MAX_ROOM or day_of_week == '토요일'
                            if not room_ok and (day_of_week != '토요일' or time_group != '9:00'):
                                group = next(g for t, g in time_groups.items() if slot in g)
                                alt_slots = [s for s in group if s != slot and s in slots and slots.index(s) in available_slots]
                                for alt_slot in alt_slots:
                                    alt_idx = slots.index(alt_slot)
                                    alt_room_num = re.search(r'\((\d+)\)', alt_slot)
                                    if alt_room_num and (total_rooms[alt_room_num.group(1)][original_name] < MAX_ROOM or day_of_week == '토요일'):
                                        possible_slots.append(alt_idx)
                                        break
                                continue
                        
                        if early_ok and late_ok and duty_ok and room_ok:
                            possible_slots.append(slot_idx)
                
                if possible_slots:
                    slot_idx = random.choice(possible_slots)
                    assignment[slot_idx] = person
                    assigned_counts[person] += 1
                    for time_group, group in time_groups.items():
                        if slots[slot_idx] in group:
                            assigned_by_time[time_group].add(original_name)
                    if slots[slot_idx] in {'8:30(1)_당직', '8:30(2)', '8:30(4)', '8:30(7)'} and day_of_week != '토요일':
                        total_early[original_name] += 1
                    if slots[slot_idx] in {'10:00(9)', '10:00(3)'} and day_of_week != '토요일':
                        total_late[original_name] += 1
                    if slots[slot_idx] in {'8:30(1)_당직', '13:30(3)_당직'} and day_of_week != '토요일':
                        total_duty[original_name] += 1
                    room_num = re.search(r'\((\d+)\)', slots[slot_idx])
                    if room_num and day_of_week != '토요일':
                        total_rooms[room_num.group(1)][original_name] += 1
                    available_slots.remove(slot_idx)
        return assignment, available_slots

    time_slots = {
        '8:30(1)_당직': 0, '8:30(2)': 1, '8:30(4)': 2, '8:30(7)': 3,
        '9:00(10)': 4, '9:00(11)': 5, '9:00(12)': 6,
        '9:30(8)': 7, '9:30(5)': 8, '9:30(6)': 9,
        '10:00(9)': 10, '10:00(3)': 11,
        '온콜': 12,
        '13:30(3)_당직': 13, '13:30(4)': 14, '13:30(9)': 15, '13:30(2)': 16
    }

    time_groups = {
        '8:30': ['8:30(1)_당직', '8:30(2)', '8:30(4)', '8:30(7)'],
        '9:00': ['9:00(10)', '9:00(11)', '9:00(12)'],
        '9:30': ['9:30(8)', '9:30(5)', '9:30(6)'],
        '10:00': ['10:00(9)', '10:00(3)'],
        '13:30': ['13:30(3)_당직', '13:30(4)', '13:30(9)', '13:30(2)'],
        '온콜': ['온콜']
    }

    weekday_slots = list(time_slots.keys())
    saturday_slots = ['8:30(1)_당직', '8:30(2)', '8:30(4)', '8:30(7)', '9:00(10)', '9:30(8)', '9:30(5)', '9:30(6)', '10:00(9)', '10:00(3)']

    memo_rules = {
        '당직 안됨': ['8:30(1)_당직', '13:30(3)_당직'],
        '오전 당직 안됨': ['8:30(1)_당직'],
        '오후 당직 안됨': ['13:30(3)_당직'],
        '당직 아닌 이른방': ['8:30(2)', '8:30(4)', '8:30(7)'],
        '8:30': ['8:30(2)', '8:30(4)', '8:30(7)'],
        '9:00': ['9:00(10)', '9:00(11)', '9:00(12)'],
        '9:30': ['9:30(8)', '9:30(5)', '9:30(6)'],
        '10:00': ['10:00(9)', '10:00(3)'],
        '이른방': ['8:30(1)_당직', '8:30(2)', '8:30(4)', '8:30(7)'],
        '오후 당직': ['13:30(3)_당직'],
        '오전 당직': ['8:30(1)_당직']
    }

    # 파일 변경 감지 및 세션 초기화
    file_hash = hash(uploaded_file.getvalue())
    if 'last_file_hash' not in st.session_state or st.session_state.last_file_hash != file_hash:
        st.session_state.clear()
        st.session_state.last_file_hash = file_hash

    # 세션 상태 초기화 및 배정 계산
    if 'assignments' not in st.session_state:
        assignments = {}
        slot_mappings = {}
        total_stats = {
            'total': Counter(), 
            'early': Counter(), 
            'late': Counter(), 
            'duty': Counter(), 
            'rooms': {str(i): Counter() for i in range(1, 13)},
            'slots': {slot.replace('_당직', ''): Counter() for slot in time_slots.keys() if slot != '온콜'}
        }
        total_fixed_stats = {slot: Counter() for slot in time_slots.keys()}
        total_memo_stats = {slot: Counter() for slot in time_slots.keys()}

        for date, data in Sheet1_data.items():
            personnel = data['personnel']
            original_personnel = data['original_personnel']
            memos = data['memos']
            day_of_week = data['day']

            fixed_assignments = {}
            current_date = date
            for row in Sheet2.iter_rows(min_row=2):
                sheet2_date = row[0].value
                if sheet2_date:
                    date_str = sheet2_date.strftime('%Y-%m-%d')
                    fixed_assignments[date_str] = {}
                    headers = Sheet2_data[date_str]['headers']
                    for col_idx, cell in enumerate(row[2:], 2):
                        if cell.value:
                            slot_key = headers[col_idx]
                            fixed_assignments[date_str][cell.value] = slot_key

            if day_of_week == '토요일':
                slots = saturday_slots.copy()
            elif day_of_week in ['월요일', '화요일', '수요일', '목요일', '금요일']:
                slots = weekday_slots.copy()
            else:
                slots = weekday_slots.copy()

            if personnel:
                assignment, fixed_assignments_record, memo_assignments = random_assign(personnel, slots, fixed_assignments, memos, day_of_week, time_groups, total_stats)
                assignments[date] = assignment
                slot_mappings[date] = slots
                
                for slot in fixed_assignments_record:
                    total_fixed_stats[slot].update(fixed_assignments_record[slot])
                for slot in memo_assignments:
                    total_memo_stats[slot].update(memo_assignments[slot])
            else:
                assignments[date] = [None] * len(slots)
                slot_mappings[date] = slots

        st.session_state.assignments = assignments
        st.session_state.slot_mappings = slot_mappings
        st.session_state.total_stats = total_stats
        st.session_state.total_fixed_stats = total_fixed_stats
        st.session_state.total_memo_stats = total_memo_stats
    else:
        assignments = st.session_state.assignments
        slot_mappings = st.session_state.slot_mappings
        total_stats = st.session_state.total_stats
        total_fixed_stats = st.session_state.total_fixed_stats
        total_memo_stats = st.session_state.total_memo_stats

    # 통합 배치 결과 DataFrame
    result_data = []
    all_columns = ['날짜', '요일'] + list(time_slots.keys())
    memo_mapping = {}

    for date in Sheet1_data.keys():
        assigned_slots = slot_mappings.get(date, weekday_slots)
        slot_to_person = {slot: None for slot in time_slots.keys()}
        assign = assignments.get(date, [None] * len(assigned_slots))
        memos = Sheet1_data[date]['memos']
        
        memo_mapping[date] = {}
        for slot, person in zip(assigned_slots, assign):
            if person:
                original_name = person.split('_')[0] if '_' in person else person
                slot_to_person[slot] = original_name
                if original_name in memos:
                    memo_mapping[date][(original_name, slot)] = memos[original_name]

        row = [date, Sheet1_data[date]['day']] + [slot_to_person[slot] for slot in time_slots.keys()]
        result_data.append(row)

    result_df = pd.DataFrame(result_data, columns=all_columns)

    # 인원별 전체 통계 DataFrame
    all_personnel = set(total_stats['total'].keys())
    stats_data = []
    slot_columns = [slot.replace('_당직', '') for slot in time_slots.keys() if slot != '온콜']
    for person in all_personnel:
        row = {
            '인원': person,
            '전체 합계': total_stats['total'].get(person, 0),
            '이른방 합계': total_stats['early'].get(person, 0),
            '늦은방 합계': total_stats['late'].get(person, 0),
            '당직 합계': total_stats['duty'].get(person, 0)
        }
        for slot in slot_columns:
            row[f'{slot} 합계'] = total_stats['slots'][slot].get(person, 0)
        stats_data.append(row)
    
    stats_df = pd.DataFrame(stats_data)
    stats_df = stats_df.sort_values(by='인원')
    stats_df = stats_df.reset_index(drop=True)

    # 정보 출력
    person_info = {}
    max_assignments = {
        '이른방 합계': MAX_EARLY, '늦은방 합계': MAX_LATE, '당직 합계': MAX_DUTY,
        '8:30(1) 합계': MAX_ROOM, '8:30(2) 합계': MAX_ROOM, '8:30(4) 합계': MAX_ROOM, '8:30(7) 합계': MAX_ROOM,
        '9:00(10) 합계': MAX_ROOM, '9:00(11) 합계': MAX_ROOM, '9:00(12) 합계': MAX_ROOM,
        '9:30(8) 합계': MAX_ROOM, '9:30(5) 합계': MAX_ROOM, '9:30(6) 합계': MAX_ROOM,
        '10:00(9) 합계': MAX_ROOM, '10:00(3) 합계': MAX_ROOM,
        '13:30(3) 합계': MAX_ROOM, '13:30(4) 합계': MAX_ROOM, '13:30(9) 합계': MAX_ROOM, '13:30(2) 합계': MAX_ROOM
    }

    for slot in total_fixed_stats:
        for person, count in total_fixed_stats[slot].items():
            if count > 0:
                if person not in person_info:
                    person_info[person] = {'fixed': {}, 'priority': {}, 'sums': {}}
                person_info[person]['fixed'][slot] = count

    for slot in total_memo_stats:
        for person, count in total_memo_stats[slot].items():
            if count > 0:
                if person not in person_info:
                    person_info[person] = {'fixed': {}, 'priority': {}, 'sums': {}}
                person_info[person]['priority'][slot] = count

    for idx, row in stats_df.iterrows():
        person = row['인원']
        if person not in person_info:
            person_info[person] = {'fixed': {}, 'priority': {}, 'sums': {}}
        for col in stats_df.columns[1:]:
            person_info[person]['sums'][col] = row[col]

    st.divider()
    st.write("### 👥 인원별 우선(고정)배정 정보")

    html_content = ""
    sorted_names = sorted(person_info.keys())

    merged_info = defaultdict(lambda: {"fixed": [], "priority": []})

    for person, info in person_info.items():
        base_name = re.sub(r'_\d+$', '', person)
        for slot, count in info['fixed'].items():
            merged_info[base_name]["fixed"].append(f"{slot} {count}번 고정 배치")
        for slot, count in info['priority'].items():
            merged_info[base_name]["priority"].append(f"{slot} {count}번 우선배치")

    html_content = ""
    sorted_names = sorted(merged_info.keys())

    for person in sorted_names:
        info = merged_info[person]
        output = [f"<span class='person'>{person}: </span>"]
        fixed_str = " / ".join(info["fixed"])
        priority_str = " / ".join(info["priority"])
        if fixed_str or priority_str:
            if fixed_str:
                output.append(fixed_str)
            if priority_str:
                output.append(f" / {priority_str}" if fixed_str else priority_str)
            html_content += f"<p>{''.join(output)}</p>"

    st.markdown(
        f"""
        <style>
        .custom-callout {{
            background-color: #f0f8ff;
            padding: 8px;
            border-radius: 6px;
            border-left: 4px solid #4682b4;
            box-shadow: 0px 2px 4px rgba(0, 0, 0, 0.1);
            margin-bottom: 4px;
            font-size: 14px;
            color: #2C3E50;
            line-height: 1.3;
        }}
        .custom-callout p {{
            margin: 0;
            padding: 2px 0;
            text-align: left;
        }}
        .person {{
            font-weight: bold;
            color: #2C3E50;
        }}
        </style>
        <div class="custom-callout">{html_content}</div>
        """,
        unsafe_allow_html=True
    )

    st.divider()
    st.write("### ⚠️ 최대 배정 한계 초과 경고")

    warnings = []
    for person in sorted_names:
        info = person_info[person]
        for slot_sum, count in info['sums'].items():
            max_count = max_assignments.get(slot_sum, float('inf'))
            if count > max_count:
                warnings.append(f"<span class='person'>{person}: </span>{slot_sum} = {count} (최대 {max_count}번 초과)")

    if warnings:
        warning_text = "".join([f"<p>{w}</p>" for w in warnings])
        html_content = f"""
        <div class="custom-callout warning-callout">
            {warning_text}
        </div>
        """
    else:
        html_content = """
        <div class="custom-callout warning-callout">
            <p>모든 배정이 적절한 한계 내에 있습니다.</p>
        </div>
        """

    st.markdown(
        f"""
        <style>
        .custom-callout {{
            padding: 8px;
            border-radius: 6px;
            box-shadow: 0px 2px 4px rgba(0, 0, 0, 0.1);
            margin-bottom: 4px;
            font-size: 14px;
            color: #2C3E50;
            line-height: 1.3;
        }}
        .custom-callout p {{
            margin: 0;
            padding: 2px 0;
            text-align: left;
        }}
        .person {{
            font-weight: bold;
            color: #2C3E50;
        }}
        .warning-callout {{
            background-color: #fff3cd;
            border-left: 4px solid #ffa500;
        }}
        </style>
        {html_content}
        """,
        unsafe_allow_html=True
    )

    st.divider()
    st.write("### 통합 배치 결과")
    st.dataframe(result_df)

    # "재랜덤화" 버튼 (result_df 아래)
    if st.button("재랜덤화"):
        st.session_state.clear()
        st.session_state.last_file_hash = file_hash
        st.rerun()  # Streamlit 재실행으로 새 배정 반영

    st.divider()
    st.write("### 인원별 전체 통계")
    st.dataframe(stats_df)

    # 엑셀 워크북 생성
    output_wb = openpyxl.Workbook()
    schedule_sheet = output_wb.active
    schedule_sheet.title = "Schedule"

    default_font = Font(name="맑은 고딕", size=9)
    bold_font = Font(name="맑은 고딕", size=9, bold=True)
    magenta_bold_font = Font(name="맑은 고딕", size=9, bold=True, color="FF00FF")
    alignment_center = Alignment(horizontal='center', vertical='center')
    border = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))

    date_fill = PatternFill(start_color="808080", end_color="808080", fill_type="solid")
    empty_row_fill = PatternFill(start_color="808080", end_color="808080", fill_type="solid")
    weekday_fill = PatternFill(start_color="FFF2CC", end_color="FFF2CC", fill_type="solid")
    saturday_with_person_fill = PatternFill(start_color="BFBFBF", end_color="BFBFBF", fill_type="solid")

    schedule_header_colors = {
        '8:30(1)_당직': PatternFill(start_color="FFE699", end_color="FFE699", fill_type="solid"),
        '8:30(2)': PatternFill(start_color="FFE699", end_color="FFE699", fill_type="solid"),
        '8:30(4)': PatternFill(start_color="FFE699", end_color="FFE699", fill_type="solid"),
        '8:30(7)': PatternFill(start_color="FFE699", end_color="FFE699", fill_type="solid"),
        '9:00(10)': PatternFill(start_color="F8CBAD", end_color="F8CBAD", fill_type="solid"),
        '9:00(11)': PatternFill(start_color="F8CBAD", end_color="F8CBAD", fill_type="solid"),
        '9:00(12)': PatternFill(start_color="F8CBAD", end_color="F8CBAD", fill_type="solid"),
        '9:30(8)': PatternFill(start_color="B4C6E7", end_color="B4C6E7", fill_type="solid"),
        '9:30(5)': PatternFill(start_color="B4C6E7", end_color="B4C6E7", fill_type="solid"),
        '9:30(6)': PatternFill(start_color="B4C6E7", end_color="B4C6E7", fill_type="solid"),
        '10:00(9)': PatternFill(start_color="C6E0B4", end_color="C6E0B4", fill_type="solid"),
        '10:00(3)': PatternFill(start_color="C6E0B4", end_color="C6E0B4", fill_type="solid"),
        '13:30(3)_당직': PatternFill(start_color="CC99FF", end_color="CC99FF", fill_type="solid"),
        '13:30(4)': PatternFill(start_color="CC99FF", end_color="CC99FF", fill_type="solid"),
        '13:30(9)': PatternFill(start_color="CC99FF", end_color="CC99FF", fill_type="solid"),
        '13:30(2)': PatternFill(start_color="CC99FF", end_color="CC99FF", fill_type="solid"),
        '온콜': PatternFill(start_color="E0E0E0", end_color="E0E0E0", fill_type="solid")
    }

    date_header_cell = schedule_sheet['A1']
    date_header_cell.value = '날짜'
    date_header_cell.font = bold_font
    date_header_cell.alignment = alignment_center
    date_header_cell.border = border

    day_header_cell = schedule_sheet['B1']
    day_header_cell.value = '요일'
    day_header_cell.font = bold_font
    day_header_cell.alignment = alignment_center
    day_header_cell.border = border

    for i, slot in enumerate(time_slots.keys(), 2):
        cell = schedule_sheet.cell(row=1, column=i+1, value=slot)
        cell.fill = schedule_header_colors.get(slot, PatternFill())
        cell.font = bold_font
        cell.alignment = alignment_center
        cell.border = border

    schedule_sheet.column_dimensions['A'].width = 12
    schedule_sheet.column_dimensions['B'].width = 8
    for col in range(3, len(time_slots.keys()) + 3):
        schedule_sheet.column_dimensions[openpyxl.utils.get_column_letter(col)].width = 10

    for i, row in enumerate(result_data, 2):
        date = row[0]
        date_obj = datetime.strptime(date, '%Y-%m-%d')
        formatted_date = date_obj.strftime('%m월 %d일')
        
        has_person = any(x is not None and x != '' for x in row[2:])

        date_cell = schedule_sheet.cell(row=i, column=1, value=formatted_date)
        date_cell.fill = date_fill
        date_cell.font = bold_font
        date_cell.alignment = alignment_center
        date_cell.border = border

        day_of_week = row[1]
        day_cell = schedule_sheet.cell(row=i, column=2, value=day_of_week)
        if not has_person:
            day_cell.fill = empty_row_fill
        elif day_of_week == '토요일':
            day_cell.fill = saturday_with_person_fill
        elif day_of_week in ['월요일', '화요일', '수요일', '목요일']:
            day_cell.fill = weekday_fill
        day_cell.font = default_font
        day_cell.alignment = alignment_center
        day_cell.border = border

        for j, value in enumerate(row[2:], 2):
            cell = schedule_sheet.cell(row=i, column=j+1, value=value)
            slot = list(time_slots.keys())[j-2]
            if slot in ['8:30(1)_당직', '13:30(3)_당직', '온콜']:
                cell.font = magenta_bold_font
            else:
                cell.font = default_font
            if not has_person:
                cell.fill = empty_row_fill
            cell.alignment = alignment_center
            cell.border = border
            memo_key = (value, slot) if value else None
            if value and date in memo_mapping and memo_key in memo_mapping[date]:
                memo = memo_mapping[date][memo_key]
                cell.comment = Comment(memo, "Memo")

    stats_sheet = output_wb.create_sheet(title="Personnel_Stats")

    personnel_fill = PatternFill(start_color="D0CECE", end_color="D0CECE", fill_type="solid")
    total_sum_fill = PatternFill(start_color="E0E0E0", end_color="E0E0E0", fill_type="solid")
    early_sum_fill = PatternFill(start_color="FFE699", end_color="FFE699", fill_type="solid")
    late_sum_fill = PatternFill(start_color="C6E0B4", end_color="C6E0B4", fill_type="solid")
    duty_sum_fill = PatternFill(start_color="FF00FF", end_color="FF00FF", fill_type="solid")
    slot_830_fill = PatternFill(start_color="FFE699", end_color="FFE699", fill_type="solid")
    slot_900_fill = PatternFill(start_color="F8CBAD", end_color="F8CBAD", fill_type="solid")
    slot_930_fill = PatternFill(start_color="B4C6E7", end_color="B4C6E7", fill_type="solid")
    slot_1000_fill = PatternFill(start_color="C6E0B4", end_color="C6E0B4", fill_type="solid")
    slot_1330_fill = PatternFill(start_color="CC99FF", end_color="CC99FF", fill_type="solid")

    headers = [
        '인원', '전체 합계', '이른방 합계', '늦은방 합계', '당직 합계',
        '8:30(1) 합계', '8:30(2) 합계', '8:30(4) 합계', '8:30(7) 합계',
        '9:00(10) 합계', '9:00(11) 합계', '9:00(12) 합계',
        '9:30(8) 합계', '9:30(5) 합계', '9:30(6) 합계',
        '10:00(9) 합계', '10:00(3) 합계',
        '13:30(3) 합계', '13:30(4) 합계', '13:30(9) 합계', '13:30(2) 합계'
    ]

    for col, header in enumerate(headers, 1):
        cell = stats_sheet.cell(row=1, column=col, value=header)
        cell.font = bold_font
        cell.alignment = alignment_center
        cell.border = border
        if header == '인원':
            cell.fill = personnel_fill
        elif header == '전체 합계':
            cell.fill = total_sum_fill
        elif header == '이른방 합계':
            cell.fill = early_sum_fill
        elif header == '늦은방 합계':
            cell.fill = late_sum_fill
        elif header == '당직 합계':
            cell.fill = duty_sum_fill
        elif header in ['8:30(1) 합계', '8:30(2) 합계', '8:30(4) 합계', '8:30(7) 합계']:
            cell.fill = slot_830_fill
        elif header in ['9:00(10) 합계', '9:00(11) 합계', '9:00(12) 합계']:
            cell.fill = slot_900_fill
        elif header in ['9:30(8) 합계', '9:30(5) 합계', '9:30(6) 합계']:
            cell.fill = slot_930_fill
        elif header in ['10:00(9) 합계', '10:00(3) 합계']:
            cell.fill = slot_1000_fill
        elif header in ['13:30(3) 합계', '13:30(4) 합계', '13:30(9) 합계', '13:30(2) 합계']:
            cell.fill = slot_1330_fill

    for row_idx, row in enumerate(stats_df.values, 2):
        for col_idx, value in enumerate(row, 1):
            cell = stats_sheet.cell(row=row_idx, column=col_idx, value=value)
            cell.font = default_font
            cell.alignment = alignment_center
            cell.border = border
            header = headers[col_idx - 1]
            if header == '인원':
                cell.font = bold_font
                cell.fill = personnel_fill

    stats_sheet.column_dimensions['A'].width = 8
    for col in range(2, len(headers) + 1):
        stats_sheet.column_dimensions[openpyxl.utils.get_column_letter(col)].width = 10

    output_stream = BytesIO()
    output_wb.save(output_stream)
    output_stream.seek(0)

    today = datetime.today().strftime("%Y-%m-%d")
    st.divider()
    st.write("### 결과 다운로드")
    st.write("- 통합 배치 결과, 인원별 전체 통계 엑셀 파일을 다운로드합니다.")
    st.download_button(
        label="다운로드",
        data=output_stream,
        file_name = f"{today}_내시경실배정.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
