from datetime import date
from pathlib import Path
import tempfile

from openpyxl import Workbook

from scheduler_app.core.engine import SchedulerEngine
from scheduler_app.core.models import Assignment, Employee
from scheduler_app.data.repository import SchedulerRepository
from scheduler_app.ui.main_window import MainWindow


def test_repo_seed_has_three_teams():
    with tempfile.TemporaryDirectory() as td:
        repo = SchedulerRepository(Path(td) / "a.db")
        repo.seed_if_empty()
        teams = {e.team for e in repo.load_employees()}
        assert teams == {"二", "三", "四"}


def test_default_admin_login_available():
    with tempfile.TemporaryDirectory() as td:
        repo = SchedulerRepository(Path(td) / "a.db")
        account = repo.verify_login("admin111", "admin111")
        assert account is not None
        assert account["is_admin"] is True


def test_cannot_remove_last_admin():
    with tempfile.TemporaryDirectory() as td:
        repo = SchedulerRepository(Path(td) / "a.db")
        admin = repo.verify_login("admin111", "admin111")
        assert admin is not None
        try:
            repo.delete_account(admin["id"])
            assert False, "expected ValueError"
        except ValueError as exc:
            assert "管理员" in str(exc)


def test_audit_logs_can_be_filtered_and_survive_account_delete():
    with tempfile.TemporaryDirectory() as td:
        repo = SchedulerRepository(Path(td) / "a.db")
        repo.create_account("user1", "pw1", False)
        account = repo.verify_login("user1", "pw1")
        assert account is not None
        repo.add_audit_log(account["id"], account["username"], "login_success", "登录成功", "成功登录排班系统")
        logs = repo.load_audit_logs(account["id"])
        assert len(logs) == 1
        assert logs[0]["username_snapshot"] == "user1"
        repo.delete_account(account["id"])
        all_logs = repo.load_audit_logs()
        assert len(all_logs) == 1
        audit_accounts = repo.list_audit_accounts()
        assert audit_accounts[0]["username_snapshot"] == "user1"


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


def test_replace_employees_from_xlsx_reads_squad_column():
    with tempfile.TemporaryDirectory() as td:
        repo = SchedulerRepository(Path(td) / "a.db")
        xlsx_path = Path(td) / "people.xlsx"
        wb = Workbook()
        ws = wb.active
        ws.append(["人员导入表", None, None, None, None, None])
        ws.append(["姓名", "所属大队", "所属中队", "类别", "职务", "是否参与排班"])
        ws.append(["张三", "二大队", "直属", "空勤", "中队长", "是"])
        ws.append(["李四", "三大队", "2中队", "地勤", "中队书记", "是"])
        wb.save(xlsx_path)

        count = repo.replace_employees_from_xlsx(xlsx_path)

        employees = repo.load_employees()
        assert count == 2
        assert [(e.name, e.team, e.squad) for e in employees] == [
            ("张三", "二", "直属"),
            ("李四", "三", "二中队"),
        ]


def test_replace_employees_from_xlsx_reads_short_squad_values():
    with tempfile.TemporaryDirectory() as td:
        repo = SchedulerRepository(Path(td) / "a.db")
        xlsx_path = Path(td) / "people_short_squad.xlsx"
        wb = Workbook()
        ws = wb.active
        ws.append(["大队", "中队", "姓名", "类别", "职责"])
        ws.append(["二", "一", "赵甲", "空勤", "中队长"])
        ws.append(["三", "二", "钱乙", "地勤", "中队书记"])
        ws.append(["四", "直属", "孙丙", "空勤", "副大队长_空勤"])
        wb.save(xlsx_path)

        repo.replace_employees_from_xlsx(xlsx_path)

        employees = repo.load_employees()
        assert [(e.name, e.team, e.squad) for e in employees] == [
            ("赵甲", "二", "一中队"),
            ("钱乙", "三", "二中队"),
            ("孙丙", "四", "直属"),
        ]

def test_save_schedule_and_stats_counts_same_day_dual_roles_once():
    with tempfile.TemporaryDirectory() as td:
        repo = SchedulerRepository(Path(td) / "a.db")
        employees = [
            Employee(1, "张三", "二大队", "一中队", "空勤", "中队长", True),
            Employee(2, "李四", "三大队", "一中队", "空勤", "中队长", True),
        ]
        assignments = [
            Assignment(date(2026, 5, 1), "SL_MAIN", 1),
            Assignment(date(2026, 5, 1), "FLEET_LEAD", 1),
            Assignment(date(2026, 5, 1), "TF_MAIN", 2),
        ]

        repo.save_schedule_and_stats(assignments, employees, "2026-05")

        assert repo.monthly_person_totals("2026-05") == {1: 1, 2: 1}
        assert repo.yearly_totals("2026") == {1: 1, 2: 1}


def test_current_year_totals_do_not_double_count_saved_current_month():
    with tempfile.TemporaryDirectory() as td:
        repo = SchedulerRepository(Path(td) / "a.db")
        employees = [
            Employee(1, "张三", "二大队", "一中队", "空勤", "中队长", True),
            Employee(2, "李四", "三大队", "一中队", "空勤", "中队长", True),
        ]
        assignments = [
            Assignment(date(2026, 5, 1), "SL_MAIN", 1),
            Assignment(date(2026, 5, 1), "FLEET_LEAD", 1),
            Assignment(date(2026, 5, 2), "TF_MAIN", 2),
        ]
        repo.save_schedule_and_stats(assignments, employees, "2026-05")

        fake_window = MainWindow.__new__(MainWindow)
        fake_window.current_assignments = assignments
        fake_window.repo = repo
        fake_window.history_person_year_totals = {}
        fake_window.history_team_year_totals = {}
        fake_window.employees = employees
        fake_window._month_start = lambda: date(2026, 5, 1)

        person_year_totals, team_year_totals = MainWindow._current_year_totals(fake_window)

        assert person_year_totals == {1: 1, 2: 1}
        assert team_year_totals == {"二大队": 1, "三大队": 1}


def test_export_allows_employees_with_zero_month_assignments(tmp_path):
    from scheduler_app.services.exporter import export_csv, export_excel

    employees = [
        Employee(1, "张三", "二大队", "一中队", "空勤", "中队长", True),
        Employee(7, "赵六", "三大队", "一中队", "空勤", "中队长", True),
    ]
    assignments = [
        Assignment(date(2026, 6, 1), "SL_MAIN", 1),
        Assignment(date(2026, 6, 1), "FLEET_LEAD", 1),
    ]

    export_csv(tmp_path / "out.csv", assignments, employees, person_year_totals={1: 1}, team_year_totals={"二大队": 1, "三大队": 0})
    export_excel(tmp_path / "out.xlsx", assignments, employees, {}, person_year_totals={1: 1}, team_year_totals={"二大队": 1, "三大队": 0})

    assert (tmp_path / "out.csv").exists()
    assert (tmp_path / "out.xlsx").exists()
