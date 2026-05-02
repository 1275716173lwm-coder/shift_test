from __future__ import annotations

import csv
from collections import defaultdict
from pathlib import Path

from openpyxl import Workbook

from scheduler_app.core.models import Assignment, Employee, POSITION_LABELS


def export_csv(path: Path, assignments: list[Assignment], employees: list[Employee]) -> None:
    by_id = {e.id: e for e in employees}
    with path.open("w", newline="", encoding="utf-8-sig") as f:
        w = csv.writer(f)
        w.writerow(["日期", "岗位", "姓名", "大队", "中队", "类别", "职责", "人工"])
        for a in sorted(assignments, key=lambda x: (x.work_date, x.position)):
            e = by_id[a.employee_id]
            w.writerow([a.work_date.isoformat(), POSITION_LABELS.get(a.position, a.position), e.name, e.team, e.squad, e.duty_group, e.role, int(a.manual)])


def export_excel(path: Path, assignments: list[Assignment], employees: list[Employee]) -> None:
    by_id = {e.id: e for e in employees}
    wb = Workbook()

    ws = wb.active
    ws.title = "Schedule"
    ws.append(["日期", "岗位", "姓名", "大队", "中队", "类别", "职责", "人工"])
    for a in sorted(assignments, key=lambda x: (x.work_date, x.position)):
        e = by_id[a.employee_id]
        ws.append([a.work_date.isoformat(), POSITION_LABELS.get(a.position, a.position), e.name, e.team, e.squad, e.duty_group, e.role, int(a.manual)])

    person_month = wb.create_sheet("Monthly_Person")
    person_month.append(["姓名", "大队", "中队", "月度值班总数"])
    p_counter = defaultdict(int)
    for a in assignments:
        p_counter[a.employee_id] += 1
    for e in employees:
        person_month.append([e.name, e.team, e.squad, p_counter[e.id]])

    team_month = wb.create_sheet("Monthly_Team")
    team_month.append(["大队", "月度值班总数"])
    t_counter = defaultdict(int)
    for a in assignments:
        t_counter[by_id[a.employee_id].team] += 1
    for team in sorted(t_counter):
        team_month.append([team, t_counter[team]])

    wb.save(path)
