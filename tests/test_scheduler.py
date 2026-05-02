from datetime import date
from pathlib import Path
import tempfile

from scheduler_app.core.engine import SchedulerEngine
from scheduler_app.core.models import Employee
from scheduler_app.data.repository import SchedulerRepository


def test_repo_seed_has_three_teams():
    with tempfile.TemporaryDirectory() as td:
        repo = SchedulerRepository(Path(td) / "a.db")
        repo.seed_if_empty()
        teams = {e.team for e in repo.load_employees()}
        assert teams == {"二", "三", "四"}


def test_engine_generates_assignments():
    employees = [
        Employee(1, "二空1", "二", "一中队", "空勤", "中队长", True),
        Employee(2, "二地1", "二", "二中队", "地勤", "中队书记", True),
        Employee(3, "二队长", "二", "一中队", "空勤", "大队长", True),
        Employee(4, "二副队", "二", "一中队", "空勤", "副大队长_空勤", True),
        Employee(5, "三空1", "三", "一中队", "空勤", "中队长", True),
        Employee(6, "三地1", "三", "二中队", "地勤", "中队书记", True),
        Employee(7, "三队长", "三", "一中队", "空勤", "大队长", True),
        Employee(8, "三副队", "三", "一中队", "空勤", "副大队长_空勤", True),
        Employee(9, "四空1", "四", "一中队", "空勤", "中队长", True),
        Employee(10, "四地1", "四", "二中队", "地勤", "中队书记", True),
        Employee(11, "四队长", "四", "一中队", "空勤", "大队长", True),
        Employee(12, "四副队", "四", "一中队", "空勤", "副大队长_空勤", True),
    ]
    engine = SchedulerEngine()
    res = engine.solve(
        employees=employees,
        start=date(2026, 5, 4),
        end=date(2026, 5, 4),
        leaves=set(),
        day_tags={date(2026, 5, 4): "WORKDAY"},
        manual_special_sl_main={},
    )
    assert len(res.assignments) >= 5
