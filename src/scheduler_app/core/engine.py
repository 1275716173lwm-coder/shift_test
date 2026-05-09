from __future__ import annotations

from collections import defaultdict
from datetime import date, timedelta

from scheduler_app.core.models import Assignment, Employee, LEADER_ROLES, POSITION_LABELS, ScheduleResult, SPECIAL_SL_MAIN_ROLES


class SchedulerEngine:
    def solve(
        self,
        employees: list[Employee],
        start: date,
        end: date,
        leaves: set[tuple[int, date]],
        day_tags: dict[date, str],
        manual_special_sl_main: dict[date, int],
        history_tail: list[dict] | None = None,
        manual_overrides: dict[tuple[date, str], int] | None = None,
    ) -> ScheduleResult:
        logs: list[str] = []
        em_by_id = {e.id: e for e in employees}
        assignments: list[Assignment] = []
        manual_overrides = manual_overrides or {}
        history_tail = history_tail or []

        team_counts = defaultdict(lambda: defaultdict(int))
        person_days = defaultdict(list)
        fleet_counts = defaultdict(int)
        last_sl_main_team = None
        last_tf_main_team = None

        for row in history_tail:
            d = date.fromisoformat(str(row["work_date"]))
            pos = str(row["position"])
            eid = int(row["employee_id"])
            e = em_by_id.get(eid)
            if not e:
                continue
            person_days[eid].append(d)
            team_counts[e.team][pos] += 1
            if pos == "FLEET_LEAD":
                fleet_counts[eid] += 1
            if pos == "SL_MAIN":
                last_sl_main_team = e.team
            if pos == "TF_MAIN":
                last_tf_main_team = e.team

        def is_special(d: date) -> bool:
            if d in day_tags:
                return day_tags[d] == "SPECIAL"
            return d.weekday() >= 5

        def is_available(e: Employee, d: date, block_leaders: bool = False) -> bool:
            if not e.participate:
                return False
            if (e.id, d) in leaves:
                return False
            for old in person_days[e.id]:
                if abs((d - old).days) < 3:
                    return False
            if block_leaders and e.role in LEADER_ROLES:
                return False
            return True

        def pick_person(cands: list[Employee], pos: str) -> Employee | None:
            if not cands:
                return None
            cands.sort(key=lambda x: (team_counts[x.team][pos], len(person_days[x.id]), x.id))
            return cands[0]

        def pick_team(cands: set[str], pos: str, last_team: str | None) -> str | None:
            if not cands:
                return None
            ranked = sorted(cands, key=lambda t: (team_counts[t][pos], t == last_team, t))
            if last_team and len(cands) > 1 and ranked[0] == last_team:
                return ranked[1]
            return ranked[0]

        def put_assignment(day_asg: dict[str, Assignment], d: date, pos: str, eid: int, manual: bool = False) -> bool:
            e = em_by_id.get(eid)
            if not e:
                logs.append(f"{d} {POSITION_LABELS.get(pos, pos)} 人员不存在")
                return False
            if (eid, d) in leaves:
                logs.append(f"{d} {POSITION_LABELS.get(pos, pos)} 指定人员请假")
                return False
            if pos != "FLEET_LEAD":
                for a in day_asg.values():
                    if a.employee_id == eid:
                        logs.append(f"{d} {e.name} 同日重复岗位")
                        return False
            if not manual and not is_available(e, d, block_leaders=(pos != "FLEET_LEAD")):
                logs.append(f"{d} {POSITION_LABELS.get(pos, pos)} 无可用人员")
                return False
            day_asg[pos] = Assignment(d, pos, eid, manual=manual)
            return True

        d = start
        while d <= end:
            special = is_special(d)
            day_asg: dict[str, Assignment] = {}

            # 1) plan page manual for special SL_MAIN
            if special and d in manual_special_sl_main:
                eid = manual_special_sl_main[d]
                e = em_by_id.get(eid)
                if not e or e.role not in SPECIAL_SL_MAIN_ROLES:
                    logs.append(f"{d} 特殊日期人工指定的{POSITION_LABELS['SL_MAIN']}无效")
                    d += timedelta(days=1)
                    continue
                if not put_assignment(day_asg, d, "SL_MAIN", eid, manual=True):
                    d += timedelta(days=1)
                    continue
            elif special:
                logs.append(f"{d} 特殊日期缺少人工指定的{POSITION_LABELS['SL_MAIN']}")
                d += timedelta(days=1)
                continue

            # 2) result page manual overrides
            required_positions = ["SL_MAIN", "TF_MAIN", "SL_AIR", "SL_GROUND", "FLEET_LEAD"]
            if special:
                required_positions.append("TF_GROUND")
            for pos in required_positions:
                eid = manual_overrides.get((d, pos))
                if eid is None:
                    continue
                if pos == "SL_MAIN" and special and d in manual_special_sl_main and manual_special_sl_main[d] != eid:
                    logs.append(f"{d} {POSITION_LABELS['SL_MAIN']}与值班计划人工指定冲突，保留值班计划")
                    continue
                if not put_assignment(day_asg, d, pos, int(eid), manual=True):
                    day_asg = {}
                    break
            if not day_asg and any((d, p) in manual_overrides for p in required_positions):
                d += timedelta(days=1)
                continue

            # 3) auto fill remaining positions
            if "SL_MAIN" not in day_asg:
                teams = {e.team for e in employees if is_available(e, d, block_leaders=True)}
                sl_main_team = pick_team(teams, "SL_MAIN", last_sl_main_team)
                if not sl_main_team:
                    logs.append(f"{d} 无可用大队安排{POSITION_LABELS['SL_MAIN']}")
                    d += timedelta(days=1)
                    continue
                cands = [e for e in employees if e.team == sl_main_team and e.duty_group == "空勤" and is_available(e, d, block_leaders=True)]
                chosen = pick_person(cands, "SL_MAIN")
                if not chosen or not put_assignment(day_asg, d, "SL_MAIN", chosen.id):
                    logs.append(f"{d} 大队{sl_main_team}无可用人员安排{POSITION_LABELS['SL_MAIN']}")
                    d += timedelta(days=1)
                    continue
            sl_main_team = em_by_id[day_asg["SL_MAIN"].employee_id].team

            if "TF_MAIN" not in day_asg:
                teams_tf = {e.team for e in employees if is_available(e, d, block_leaders=True)} - {sl_main_team}
                tf_team = pick_team(teams_tf, "TF_MAIN", last_tf_main_team)
                if not tf_team:
                    logs.append(f"{d} 无可用大队安排{POSITION_LABELS['TF_MAIN']}")
                    d += timedelta(days=1)
                    continue
                cands_tf = [e for e in employees if e.team == tf_team and e.duty_group == "空勤" and is_available(e, d, block_leaders=True)]
                chosen_tf = pick_person(cands_tf, "TF_MAIN")
                if not chosen_tf or not put_assignment(day_asg, d, "TF_MAIN", chosen_tf.id):
                    logs.append(f"{d} 大队{tf_team}无可用人员安排{POSITION_LABELS['TF_MAIN']}")
                    d += timedelta(days=1)
                    continue
            tf_team = em_by_id[day_asg["TF_MAIN"].employee_id].team

            if "SL_AIR" not in day_asg:
                used_teams = {sl_main_team, tf_team}
                teams_air = {e.team for e in employees if is_available(e, d, block_leaders=True)} - used_teams
                sl_air_team = pick_team(teams_air, "SL_AIR", None)
                if not sl_air_team:
                    logs.append(f"{d} 无可用大队安排{POSITION_LABELS['SL_AIR']}")
                    d += timedelta(days=1)
                    continue
                cands_air = [e for e in employees if e.team == sl_air_team and e.duty_group == "空勤" and is_available(e, d, block_leaders=True)]
                chosen_air = pick_person(cands_air, "SL_AIR")
                if not chosen_air or not put_assignment(day_asg, d, "SL_AIR", chosen_air.id):
                    logs.append(f"{d} 大队{sl_air_team}无可用人员安排{POSITION_LABELS['SL_AIR']}")
                    d += timedelta(days=1)
                    continue

            if "SL_GROUND" not in day_asg:
                cands_slg = [e for e in employees if e.team == tf_team and e.duty_group == "地勤" and is_available(e, d)]
                chosen_slg = pick_person(cands_slg, "SL_GROUND")
                if not chosen_slg or not put_assignment(day_asg, d, "SL_GROUND", chosen_slg.id):
                    logs.append(f"{d} 大队{tf_team}无可用人员安排{POSITION_LABELS['SL_GROUND']}")
                    d += timedelta(days=1)
                    continue

            if special and "TF_GROUND" not in day_asg:
                cands_tfg = [e for e in employees if e.team == sl_main_team and e.duty_group == "地勤" and is_available(e, d)]
                chosen_tfg = pick_person(cands_tfg, "TF_GROUND")
                if not chosen_tfg or not put_assignment(day_asg, d, "TF_GROUND", chosen_tfg.id):
                    logs.append(f"{d} 大队{sl_main_team}无可用人员安排{POSITION_LABELS['TF_GROUND']}")
                    d += timedelta(days=1)
                    continue

            if "FLEET_LEAD" not in day_asg:
                sl_main_emp = em_by_id[day_asg["SL_MAIN"].employee_id]
                if sl_main_emp.role in LEADER_ROLES and is_available(sl_main_emp, d):
                    fleet_emp = sl_main_emp
                else:
                    leaders = [e for e in employees if e.team == sl_main_emp.team and e.role in LEADER_ROLES and (e.id, d) not in leaves]
                    leaders.sort(key=lambda x: (fleet_counts[x.id], x.id))
                    fleet_emp = leaders[0] if leaders else None
                if not fleet_emp or not put_assignment(day_asg, d, "FLEET_LEAD", fleet_emp.id):
                    logs.append(f"{d} 无可用人员安排{POSITION_LABELS['FLEET_LEAD']}")
                    d += timedelta(days=1)
                    continue

            for a in day_asg.values():
                assignments.append(a)
                person_days[a.employee_id].append(d)
                team_counts[em_by_id[a.employee_id].team][a.position] += 1
                if a.position == "FLEET_LEAD":
                    fleet_counts[a.employee_id] += 1
            last_sl_main_team = sl_main_team
            last_tf_main_team = tf_team
            d += timedelta(days=1)

        assignments.sort(key=lambda x: (x.work_date, x.position))
        return ScheduleResult(assignments=assignments, logs=logs)
