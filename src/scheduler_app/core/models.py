from dataclasses import dataclass
from datetime import date


ROLE_LEADER = "大队长"
ROLE_DEPUTY_LEADER_AIR = "副大队长_空勤"
ROLE_DEPUTY_LEADER_GROUND = "副大队长_地勤"
ROLE_PARTY_SECRETARY = "党总支书记"
ROLE_SQUAD_LEADER = "中队长"
ROLE_SQUAD_DEPUTY = "副中队长"
ROLE_SQUAD_SECRETARY = "中队书记"

LEADER_ROLES = {ROLE_LEADER, ROLE_DEPUTY_LEADER_AIR, ROLE_DEPUTY_LEADER_GROUND}


@dataclass(slots=True)
class Employee:
    id: int
    name: str
    team: str
    squad: str
    duty_group: str  # 空勤 | 地勤
    role: str
    participate: bool = True


@dataclass(slots=True)
class Assignment:
    work_date: date
    position: str
    employee_id: int
    manual: bool = False


@dataclass(slots=True)
class ScheduleResult:
    assignments: list[Assignment]
    logs: list[str]


POSITIONS_WORKDAY = ["SL_MAIN", "SL_AIR", "SL_GROUND", "TF_MAIN"]
POSITIONS_SPECIAL = ["SL_MAIN", "SL_AIR", "SL_GROUND", "TF_MAIN", "TF_GROUND"]

POSITION_LABELS = {
    "SL_MAIN": "双流主班",
    "SL_AIR": "双流副班(空勤)",
    "SL_GROUND": "双流副班(地勤)",
    "TF_MAIN": "天府主班",
    "TF_GROUND": "天府副班(地勤)",
    "FLEET_LEAD": "机队总负责",
}
