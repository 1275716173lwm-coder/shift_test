from __future__ import annotations

from collections import defaultdict
from datetime import date, timedelta

from scheduler_app.core.models import Assignment, Employee, LEADER_ROLES, POSITION_LABELS, POSITIONS_SPECIAL, POSITIONS_WORKDAY, ScheduleResult


class SchedulerEngine:
    def solve(
        self,
        employees: list[Employee],
        start: date,
        end: date,
        leaves: set[tuple[int, date]],
        day_tags: dict[date, str],
        manual_special_sl_main: dict[date, int],
    ) -> ScheduleResult:
        logs: list[str] = []
        em_by_id = {e.id: e for e in employees}
        assignments: list[Assignment] = []

        team_counts = defaultdict(lambda: defaultdict(int))
        person_days = defaultdict(list)
        fleet_counts = defaultdict(int)
        last_sl_main_team = None
        last_tf_main_team = None

        def is_special(d: date) -> bool:
            if d in day_tags:
                return day_tags[d] == "SPECIAL"
            return d.weekday() >= 5

        def available(e: Employee, d: date, auto=True) -> bool:
            if not e.participate:
                return False
            if (e.id, d) in leaves:
                return False
            for old in person_days[e.id]:
                if abs((d - old).days) < 3:
                    return False
            if auto and e.role in LEADER_ROLES:
                return False
            return True

        def pick_person(cands: list[Employee], pos: str) -> Employee | None:
            if not cands:
                return None
            cands.sort(key=lambda e: (team_counts[e.team][pos], len(person_days[e.id]), e.id))
            return cands[0]

        def pick_team(cands: set[str], pos: str, last_team: str | None) -> str | None:
            if not cands:
                return None
            ranked = sorted(cands, key=lambda t: (team_counts[t][pos], t == last_team, t))
            if last_team and len(cands) > 1 and ranked[0] == last_team:
                return ranked[1]
            return ranked[0]

        d = start
        while d <= end:
            special = is_special(d)
            day_asg: list[Assignment] = []

            if special:
                if d not in manual_special_sl_main:
                    logs.append(f"{d} 特殊日期缺少人工指定的{POSITION_LABELS['SL_MAIN']}")
                    d += timedelta(days=1)
                    continue
                eid = manual_special_sl_main[d]
                e = em_by_id.get(eid)
                if not e or e.role not in LEADER_ROLES or (eid, d) in leaves:
                    logs.append(f"{d} 特殊日期人工指定的{POSITION_LABELS['SL_MAIN']}无效")
                    d += timedelta(days=1)
                    continue
                day_asg.append(Assignment(d, "SL_MAIN", eid, manual=True))
                sl_main_team = e.team
            else:
                teams = {e.team for e in employees if available(e, d, auto=True)}
                sl_main_team = pick_team(teams, "SL_MAIN", last_sl_main_team)
                if not sl_main_team:
                    logs.append(f"{d} 无可用大队安排{POSITION_LABELS['SL_MAIN']}")
                    d += timedelta(days=1)
                    continue
                cands = [e for e in employees if e.team == sl_main_team and e.duty_group == "空勤" and available(e, d, auto=True)]
                chosen = pick_person(cands, "SL_MAIN")
                if not chosen:
                    logs.append(f"{d} 大队{sl_main_team}无可用人员安排{POSITION_LABELS['SL_MAIN']}")
                    d += timedelta(days=1)
                    continue
                day_asg.append(Assignment(d, "SL_MAIN", chosen.id))

            teams_tf = {e.team for e in employees if available(e, d, auto=True)} - {sl_main_team}
            tf_team = pick_team(teams_tf, "TF_MAIN", last_tf_main_team)
            if not tf_team:
                logs.append(f"{d} 无可用大队安排{POSITION_LABELS['TF_MAIN']}")
                d += timedelta(days=1)
                continue
            cands_tf = [e for e in employees if e.team == tf_team and e.duty_group == "空勤" and available(e, d, auto=True)]
            chosen_tf = pick_person(cands_tf, "TF_MAIN")
            if not chosen_tf:
                logs.append(f"{d} 大队{tf_team}无可用人员安排{POSITION_LABELS['TF_MAIN']}")
                d += timedelta(days=1)
                continue
            day_asg.append(Assignment(d, "TF_MAIN", chosen_tf.id))

            used_teams = {sl_main_team, tf_team}
            teams_air = {e.team for e in employees if available(e, d, auto=True)} - used_teams
            sl_air_team = pick_team(teams_air, "SL_AIR", None)
            if not sl_air_team:
                logs.append(f"{d} 无可用大队安排{POSITION_LABELS['SL_AIR']}")
                d += timedelta(days=1)
                continue
            cands_air = [e for e in employees if e.team == sl_air_team and e.duty_group == "空勤" and available(e, d, auto=True)]
            chosen_air = pick_person(cands_air, "SL_AIR")
            if not chosen_air:
                logs.append(f"{d} 大队{sl_air_team}无可用人员安排{POSITION_LABELS['SL_AIR']}")
                d += timedelta(days=1)
                continue
            day_asg.append(Assignment(d, "SL_AIR", chosen_air.id))

            cands_slg = [e for e in employees if e.team == tf_team and e.duty_group == "地勤" and available(e, d, auto=True)]
            chosen_slg = pick_person(cands_slg, "SL_GROUND")
            if not chosen_slg:
                logs.append(f"{d} 大队{tf_team}无可用人员安排{POSITION_LABELS['SL_GROUND']}")
                d += timedelta(days=1)
                continue
            day_asg.append(Assignment(d, "SL_GROUND", chosen_slg.id))

            if special:
                cands_tfg = [e for e in employees if e.team == sl_main_team and e.duty_group == "地勤" and available(e, d, auto=True)]
                chosen_tfg = pick_person(cands_tfg, "TF_GROUND")
                if not chosen_tfg:
                    logs.append(f"{d} 大队{sl_main_team}无可用人员安排{POSITION_LABELS['TF_GROUND']}")
                    d += timedelta(days=1)
                    continue
                day_asg.append(Assignment(d, "TF_GROUND", chosen_tfg.id))

            sl_main_emp = em_by_id[next(a.employee_id for a in day_asg if a.position == "SL_MAIN")]
            if sl_main_emp.role in LEADER_ROLES:
                fleet_emp = sl_main_emp
            else:
                leaders = [
                    e for e in employees
                    if e.team == sl_main_emp.team and e.role in LEADER_ROLES and (e.id, d) not in leaves
                ]
                leaders.sort(key=lambda e: (fleet_counts[e.id], e.id))
                fleet_emp = leaders[0] if leaders else None
            if not fleet_emp:
                logs.append(f"{d} 无可用人员安排{POSITION_LABELS['FLEET_LEAD']}")
                d += timedelta(days=1)
                continue
            if person_days[fleet_emp.id] and (d - person_days[fleet_emp.id][-1]).days < 2:
                logs.append(f"{d} {POSITION_LABELS['FLEET_LEAD']}连续值班提醒: {fleet_emp.name}")
            day_asg.append(Assignment(d, "FLEET_LEAD", fleet_emp.id))

            for a in day_asg:
                assignments.append(a)
                person_days[a.employee_id].append(d)
                team_counts[em_by_id[a.employee_id].team][a.position] += 1
                if a.position == "FLEET_LEAD":
                    fleet_counts[a.employee_id] += 1
            last_sl_main_team = sl_main_team
            last_tf_main_team = tf_team
            d += timedelta(days=1)

        return ScheduleResult(assignments=assignments, logs=logs)
