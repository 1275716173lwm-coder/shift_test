from __future__ import annotations

import sqlite3
from datetime import date, datetime
from pathlib import Path
from openpyxl import load_workbook

from scheduler_app.core.models import Assignment, Employee


ROLE_MAP = {
    "leader": "大队长",
    "deputy_leader": "副大队长_空勤",
    "ground_admin": "中队书记",
    "member": "中队长",
    "squad_leader": "中队长",
    "squad_deputy": "副中队长",
}
GROUP_MAP = {"air": "空勤", "ground": "地勤"}
TEAM_MAP = {"A": "二", "B": "三", "C": "四"}


class SchedulerRepository:
    def __init__(self, db_path: Path) -> None:
        self.db_path = db_path
        self.db_path.parent.mkdir(parents=True, exist_ok=True)
        self._init_schema()

    def _connect(self) -> sqlite3.Connection:
        conn = sqlite3.connect(self.db_path)
        conn.row_factory = sqlite3.Row
        return conn

    def _init_schema(self) -> None:
        with self._connect() as conn:
            conn.executescript(
                """
                CREATE TABLE IF NOT EXISTS employees(
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    name TEXT NOT NULL,
                    team TEXT NOT NULL,
                    squad TEXT NOT NULL DEFAULT '一中队',
                    duty_group TEXT NOT NULL,
                    role TEXT NOT NULL,
                    participate INTEGER NOT NULL DEFAULT 1,
                    order_index INTEGER NOT NULL DEFAULT 0
                );
                CREATE TABLE IF NOT EXISTS leaves(
                    employee_id INTEGER NOT NULL,
                    work_date TEXT NOT NULL,
                    PRIMARY KEY(employee_id, work_date)
                );
                CREATE TABLE IF NOT EXISTS day_tags(
                    work_date TEXT PRIMARY KEY,
                    tag TEXT NOT NULL
                );
                CREATE TABLE IF NOT EXISTS manual_assignments(
                    work_date TEXT NOT NULL,
                    position TEXT NOT NULL,
                    employee_id INTEGER NOT NULL,
                    PRIMARY KEY(work_date, position)
                );
                CREATE TABLE IF NOT EXISTS snapshots(
                    id TEXT PRIMARY KEY,
                    created_at TEXT NOT NULL,
                    note TEXT NOT NULL,
                    assignments_json TEXT NOT NULL,
                    logs_text TEXT NOT NULL
                );
                """
            )
            cols = {r["name"] for r in conn.execute("PRAGMA table_info(employees)").fetchall()}
            if "squad" not in cols:
                conn.execute("ALTER TABLE employees ADD COLUMN squad TEXT NOT NULL DEFAULT '一中队'")
            if "order_index" not in cols:
                conn.execute("ALTER TABLE employees ADD COLUMN order_index INTEGER NOT NULL DEFAULT 0")

            rows = conn.execute("SELECT id, team, duty_group, role FROM employees").fetchall()
            for r in rows:
                team = TEAM_MAP.get(r["team"], r["team"])
                duty = GROUP_MAP.get(r["duty_group"], r["duty_group"])
                role = ROLE_MAP.get(r["role"], r["role"])
                conn.execute(
                    "UPDATE employees SET team=?, duty_group=?, role=? WHERE id=?",
                    (team, duty, role, r["id"]),
                )
            # Ensure every record has a stable order.
            ordered = conn.execute("SELECT id FROM employees ORDER BY order_index, id").fetchall()
            for idx, r in enumerate(ordered):
                conn.execute("UPDATE employees SET order_index=? WHERE id=?", (idx, r["id"]))

    def seed_if_empty(self) -> None:
        with self._connect() as conn:
            n = conn.execute("SELECT COUNT(*) n FROM employees").fetchone()["n"]
            if n:
                return
            seed = [
                ("二大队长", "二", "一中队", "空勤", "大队长", 1),
                ("二副大队长空", "二", "一中队", "空勤", "副大队长_空勤", 1),
                ("二副大队长地", "二", "二中队", "地勤", "副大队长_地勤", 1),
                ("二党总支书记", "二", "二中队", "地勤", "党总支书记", 1),
                ("二中队长", "二", "三中队", "空勤", "中队长", 1),
                ("二副中队长", "二", "三中队", "空勤", "副中队长", 1),
                ("二中队书记", "二", "四中队", "地勤", "中队书记", 1),
                ("三大队长", "三", "一中队", "空勤", "大队长", 1),
                ("三副大队长空", "三", "一中队", "空勤", "副大队长_空勤", 1),
                ("三副大队长地", "三", "二中队", "地勤", "副大队长_地勤", 1),
                ("三党总支书记", "三", "二中队", "地勤", "党总支书记", 1),
                ("三中队长", "三", "三中队", "空勤", "中队长", 1),
                ("三副中队长", "三", "三中队", "空勤", "副中队长", 1),
                ("三中队书记", "三", "四中队", "地勤", "中队书记", 1),
                ("四大队长", "四", "一中队", "空勤", "大队长", 1),
                ("四副大队长空", "四", "一中队", "空勤", "副大队长_空勤", 1),
                ("四副大队长地", "四", "二中队", "地勤", "副大队长_地勤", 1),
                ("四党总支书记", "四", "二中队", "地勤", "党总支书记", 1),
                ("四中队长", "四", "三中队", "空勤", "中队长", 1),
                ("四副中队长", "四", "三中队", "空勤", "副中队长", 1),
                ("四中队书记", "四", "四中队", "地勤", "中队书记", 1),
            ]
            for idx, row in enumerate(seed):
                conn.execute(
                    "INSERT INTO employees(name,team,squad,duty_group,role,participate,order_index) VALUES(?,?,?,?,?,?,?)",
                    (*row, idx),
                )

    def load_employees(self) -> list[Employee]:
        with self._connect() as conn:
            rows = conn.execute("SELECT * FROM employees ORDER BY order_index, id").fetchall()
            return [Employee(id=r["id"], name=r["name"], team=r["team"], squad=r["squad"], duty_group=r["duty_group"], role=r["role"], participate=bool(r["participate"])) for r in rows]

    def add_employee(self, name: str, team: str, squad: str, duty_group: str, role: str, participate: bool) -> None:
        with self._connect() as conn:
            next_order = conn.execute("SELECT COALESCE(MAX(order_index), -1) + 1 AS n FROM employees").fetchone()["n"]
            conn.execute(
                "INSERT INTO employees(name,team,squad,duty_group,role,participate,order_index) VALUES(?,?,?,?,?,?,?)",
                (name, team, squad, duty_group, role, 1 if participate else 0, next_order),
            )

    def replace_employees_from_xlsx(self, xlsx_path: Path) -> int:
        wb = load_workbook(xlsx_path, data_only=True)
        ws = wb.active

        headers = [str(ws.cell(1, c).value or "").strip() for c in range(1, ws.max_column + 1)]
        idx = {h: i + 1 for i, h in enumerate(headers)}

        # Required columns (Chinese)
        required = ["姓名", "大队", "中队", "类别", "职责"]
        missing = [c for c in required if c not in idx]
        if missing:
            raise ValueError(f"Excel缺少必要列: {', '.join(missing)}")

        # Optional: 参与排班
        participate_col = idx.get("参与排班")

        rows_to_insert: list[tuple[str, str, str, str, str, int, int]] = []
        order_index = 0
        for r in range(2, ws.max_row + 1):
            name = str(ws.cell(r, idx["姓名"]).value or "").strip()
            team = str(ws.cell(r, idx["大队"]).value or "").strip()
            squad = str(ws.cell(r, idx["中队"]).value or "").strip()
            duty_group = str(ws.cell(r, idx["类别"]).value or "").strip()
            role = str(ws.cell(r, idx["职责"]).value or "").strip()
            if not name:
                continue
            if not team:
                team = "二"
            if not squad:
                squad = "直属"
            if not duty_group:
                duty_group = "空勤"
            if not role:
                role = "中队长"

            participate = 1
            if participate_col:
                raw = str(ws.cell(r, participate_col).value or "").strip()
                if raw in {"0", "否", "false", "False", "不参与"}:
                    participate = 0

            rows_to_insert.append((name, team, squad, duty_group, role, participate, order_index))
            order_index += 1

        with self._connect() as conn:
            conn.execute("DELETE FROM manual_assignments")
            conn.execute("DELETE FROM leaves")
            conn.execute("DELETE FROM employees")
            conn.execute("DELETE FROM sqlite_sequence WHERE name='employees'")
            for row in rows_to_insert:
                conn.execute(
                    "INSERT INTO employees(name,team,squad,duty_group,role,participate,order_index) VALUES(?,?,?,?,?,?,?)",
                    row,
                )
        return len(rows_to_insert)

    def update_employee_meta(self, employee_id: int, team: str, squad: str, duty_group: str, role: str) -> None:
        with self._connect() as conn:
            conn.execute(
                "UPDATE employees SET team=?, squad=?, duty_group=?, role=? WHERE id=?",
                (team, squad, duty_group, role, employee_id),
            )

    def delete_employee(self, employee_id: int) -> None:
        with self._connect() as conn:
            conn.execute("DELETE FROM leaves WHERE employee_id=?", (employee_id,))
            conn.execute("DELETE FROM employees WHERE id=?", (employee_id,))

    def set_participate(self, employee_id: int, participate: bool) -> None:
        with self._connect() as conn:
            conn.execute("UPDATE employees SET participate=? WHERE id=?", (1 if participate else 0, employee_id))

    def reorder_employees(self, employee_ids_in_order: list[int]) -> None:
        with self._connect() as conn:
            for idx, eid in enumerate(employee_ids_in_order):
                conn.execute("UPDATE employees SET order_index=? WHERE id=?", (idx, eid))

    def load_leaves(self) -> set[tuple[int, date]]:
        with self._connect() as conn:
            return {(r["employee_id"], date.fromisoformat(r["work_date"])) for r in conn.execute("SELECT * FROM leaves").fetchall()}

    def toggle_leave(self, employee_id: int, d: date) -> bool:
        with self._connect() as conn:
            row = conn.execute("SELECT 1 FROM leaves WHERE employee_id=? AND work_date=?", (employee_id, d.isoformat())).fetchone()
            if row:
                conn.execute("DELETE FROM leaves WHERE employee_id=? AND work_date=?", (employee_id, d.isoformat())); return False
            conn.execute("INSERT INTO leaves(employee_id,work_date) VALUES(?,?)", (employee_id, d.isoformat())); return True

    def remove_leave(self, employee_id: int, d: date) -> None:
        with self._connect() as conn:
            conn.execute("DELETE FROM leaves WHERE employee_id=? AND work_date=?", (employee_id, d.isoformat()))

    def clear_leaves_for_employee(self, employee_id: int) -> None:
        with self._connect() as conn:
            conn.execute("DELETE FROM leaves WHERE employee_id=?", (employee_id,))

    def clear_all_leaves(self) -> None:
        with self._connect() as conn:
            conn.execute("DELETE FROM leaves")

    def set_day_tag(self, d: date, tag: str) -> None:
        with self._connect() as conn:
            conn.execute("INSERT OR REPLACE INTO day_tags(work_date,tag) VALUES(?,?)", (d.isoformat(), tag))

    def load_day_tags(self) -> dict[date, str]:
        with self._connect() as conn:
            return {date.fromisoformat(r["work_date"]): r["tag"] for r in conn.execute("SELECT * FROM day_tags").fetchall()}

    def set_manual_assignment(self, d: date, position: str, employee_id: int) -> None:
        with self._connect() as conn:
            conn.execute("INSERT OR REPLACE INTO manual_assignments(work_date,position,employee_id) VALUES(?,?,?)", (d.isoformat(), position, employee_id))

    def clear_manual_assignment(self, d: date, position: str) -> None:
        with self._connect() as conn:
            conn.execute(
                "DELETE FROM manual_assignments WHERE work_date=? AND position=?",
                (d.isoformat(), position),
            )

    def load_manual_assignments(self, position: str) -> dict[date, int]:
        with self._connect() as conn:
            return {date.fromisoformat(r["work_date"]): r["employee_id"] for r in conn.execute("SELECT * FROM manual_assignments WHERE position=?", (position,)).fetchall()}

    def save_snapshot(self, snapshot_id: str, note: str, assignments: list[Assignment], logs: list[str]) -> None:
        import json
        payload = [{"work_date": a.work_date.isoformat(), "position": a.position, "employee_id": a.employee_id, "manual": a.manual} for a in assignments]
        with self._connect() as conn:
            conn.execute("INSERT INTO snapshots(id,created_at,note,assignments_json,logs_text) VALUES(?,?,?,?,?)", (snapshot_id, datetime.now().isoformat(timespec="seconds"), note, json.dumps(payload, ensure_ascii=False), "\n".join(logs)))

    def list_snapshots(self) -> list[dict]:
        with self._connect() as conn:
            return [dict(r) for r in conn.execute("SELECT id,created_at,note FROM snapshots ORDER BY created_at DESC").fetchall()]
