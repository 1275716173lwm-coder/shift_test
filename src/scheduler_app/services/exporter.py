from __future__ import annotations

import csv
from collections import defaultdict
from datetime import timedelta
from pathlib import Path

from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side

from scheduler_app.core.models import Assignment, Employee, POSITION_LABELS


def export_csv(path: Path, assignments: list[Assignment], employees: list[Employee]) -> None:
    by_id = {e.id: e for e in employees}
    with path.open("w", newline="", encoding="utf-8-sig") as f:
        w = csv.writer(f)
        w.writerow(["日期", "岗位", "姓名", "大队", "中队", "类别", "职责", "人工"])
        for a in sorted(assignments, key=lambda x: (x.work_date, x.position)):
            e = by_id[a.employee_id]
            w.writerow([
                a.work_date.isoformat(),
                POSITION_LABELS.get(a.position, a.position),
                e.name,
                e.team,
                e.squad,
                e.duty_group,
                e.role,
                int(a.manual),
            ])


def export_excel(path: Path, assignments: list[Assignment], employees: list[Employee], day_tags: dict | None = None) -> None:
    by_id = {e.id: e for e in employees}
    day_tags = day_tags or {}
    wb = Workbook()

    head_fill = PatternFill("solid", fgColor="1F4E78")
    head_font = Font(color="FFFFFF", bold=True, size=11)
    title_font = Font(bold=True, size=14, color="1F2937")
    center = Alignment(horizontal="center", vertical="center")
    thin = Side(style="thin", color="D9D9D9")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)

    ws = wb.active
    ws.title = "排班结果"
    ws.merge_cells("A1:K1")
    ws["A1"] = "排班结果明细"
    ws["A1"].font = title_font
    ws["A1"].alignment = center

    ws.append([])
    ws.append(["日期", "星期", "是否周末/节假日", "双流主班", "双流副班（空勤）", "双流副班（地勤）", "天府主班", "天府副班", "机队总负责", "双流主班大队", "天府主班大队"])

    weekday_map = {0: "星期一", 1: "星期二", 2: "星期三", 3: "星期四", 4: "星期五", 5: "星期六", 6: "星期日"}
    by_date_pos: dict = {}
    for a in assignments:
        by_date_pos[(a.work_date, a.position)] = a

    all_dates = sorted({a.work_date for a in assignments})
    if all_dates:
        d = all_dates[0]
        end = all_dates[-1]
        continuous_dates = []
        while d <= end:
            continuous_dates.append(d)
            d += timedelta(days=1)
    else:
        continuous_dates = []

    def is_special(d):
        if d in day_tags:
            return day_tags[d] == "SPECIAL"
        return d.weekday() >= 5

    def name_of(d, pos):
        a = by_date_pos.get((d, pos))
        if not a:
            return ""
        return by_id[a.employee_id].name

    def team_of(d, pos):
        a = by_date_pos.get((d, pos))
        if not a:
            return ""
        return by_id[a.employee_id].team

    for d in continuous_dates:
        ws.append([
            d.isoformat(),
            weekday_map[d.weekday()],
            "是" if is_special(d) else "否",
            name_of(d, "SL_MAIN"),
            name_of(d, "SL_AIR"),
            name_of(d, "SL_GROUND"),
            name_of(d, "TF_MAIN"),
            name_of(d, "TF_GROUND"),
            name_of(d, "FLEET_LEAD"),
            team_of(d, "SL_MAIN"),
            team_of(d, "TF_MAIN"),
        ])

    for c, w in zip("ABCDEFGHIJK", [14, 10, 14, 16, 18, 18, 16, 14, 16, 14, 14]):
        ws.column_dimensions[c].width = w
    for row in ws.iter_rows(min_row=3, max_row=ws.max_row, min_col=1, max_col=11):
        for cell in row:
            cell.border = border
            cell.alignment = center
            if cell.row == 3:
                cell.fill = head_fill
                cell.font = head_font

    person_month = wb.create_sheet("人员月度统计")
    person_month.merge_cells("A1:D1")
    person_month["A1"] = "人员月度统计"
    person_month["A1"].font = title_font
    person_month["A1"].alignment = center
    person_month.append([])
    person_month.append(["姓名", "大队", "中队", "月度值班总数"])
    p_counter = defaultdict(int)
    for a in assignments:
        p_counter[a.employee_id] += 1
    for e in employees:
        person_month.append([e.name, e.team, e.squad, p_counter[e.id]])

    for c, w in zip("ABCD", [14, 10, 12, 14]):
        person_month.column_dimensions[c].width = w
    for row in person_month.iter_rows(min_row=3, max_row=person_month.max_row, min_col=1, max_col=4):
        for cell in row:
            cell.border = border
            cell.alignment = center
            if cell.row == 3:
                cell.fill = head_fill
                cell.font = head_font

    team_month = wb.create_sheet("大队月度统计")
    team_month.merge_cells("A1:B1")
    team_month["A1"] = "大队月度统计"
    team_month["A1"].font = title_font
    team_month["A1"].alignment = center
    team_month.append([])
    team_month.append(["大队", "月度值班总数"])
    t_counter = defaultdict(int)
    for a in assignments:
        t_counter[by_id[a.employee_id].team] += 1
    for team in sorted(t_counter):
        team_month.append([team, t_counter[team]])

    for c, w in zip("AB", [12, 16]):
        team_month.column_dimensions[c].width = w
    for row in team_month.iter_rows(min_row=3, max_row=team_month.max_row, min_col=1, max_col=2):
        for cell in row:
            cell.border = border
            cell.alignment = center
            if cell.row == 3:
                cell.fill = head_fill
                cell.font = head_font

    wb.save(path)
