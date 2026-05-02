from __future__ import annotations

import sys
from collections import defaultdict
from datetime import date, timedelta
from pathlib import Path
from typing import Callable

from PySide6.QtCore import QDate, Qt
from PySide6.QtWidgets import (
    QAbstractItemView,
    QApplication,
    QCalendarWidget,
    QCheckBox,
    QComboBox,
    QDateEdit,
    QFileDialog,
    QHBoxLayout,
    QHeaderView,
    QLabel,
    QLineEdit,
    QListWidget,
    QMainWindow,
    QMessageBox,
    QPushButton,
    QPlainTextEdit,
    QTableWidget,
    QTableWidgetItem,
    QTabWidget,
    QVBoxLayout,
    QWidget,
)
from PySide6.QtGui import QColor, QTextCharFormat, QPainter

from scheduler_app.core.engine import SchedulerEngine
from scheduler_app.core.models import Assignment, LEADER_ROLES, POSITION_LABELS
from scheduler_app.data.repository import SchedulerRepository
from scheduler_app.services.exporter import export_csv, export_excel

try:
    import chinese_calendar
except Exception:
    chinese_calendar = None


class PeopleTableWidget(QTableWidget):
    def __init__(self, on_reordered: Callable[[], None], *args, **kwargs):
        super().__init__(*args, **kwargs)
        self._on_reordered = on_reordered
        self._drag_allowed = False

    def mousePressEvent(self, event):
        idx = self.indexAt(event.pos())
        self._drag_allowed = idx.isValid() and idx.column() == 0
        super().mousePressEvent(event)

    def startDrag(self, supportedActions):
        if not self._drag_allowed:
            return
        super().startDrag(supportedActions)

    def dropEvent(self, event):
        super().dropEvent(event)
        self._on_reordered()


class PlanCalendarWidget(QCalendarWidget):
    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)
        self._manual_name_map: dict[date, str] = {}

    def set_manual_name_map(self, manual_name_map: dict[date, str]):
        self._manual_name_map = manual_name_map
        self.updateCells()

    def paintCell(self, painter: QPainter, rect, qdate):
        super().paintCell(painter, rect, qdate)
        name = self._manual_name_map.get(qdate.toPython())
        if not name:
            return
        painter.save()
        painter.setPen(QColor("#2b6de0"))
        f = painter.font()
        f.setPointSize(max(7, f.pointSize() - 1))
        painter.setFont(f)
        text_rect = rect.adjusted(1, rect.height() // 2 + 3, -1, -2)
        painter.drawText(text_rect, Qt.AlignHCenter | Qt.AlignBottom, name)
        painter.restore()


class MainWindow(QMainWindow):
    def __init__(self, db_path: Path):
        super().__init__()
        self.setWindowTitle("排班系统")
        self.resize(1480, 920)
        self.repo = SchedulerRepository(db_path)
        self.repo.seed_if_empty()

        self.current_assignments: list[Assignment] = []
        self.employees = self.repo.load_employees()
        self.current_plan_date: date | None = None

        self._build_ui()
        self.refresh_people()
        self.refresh_leave_list()
        self.load_day_plan_table()

    def _build_ui(self):
        root = QWidget()
        self.setCentralWidget(root)
        layout = QVBoxLayout(root)

        top = QHBoxLayout()
        self.start_date = QDateEdit()
        self.start_date.setCalendarPopup(True)
        self.start_date.setDate(date.today().replace(day=1))
        self.end_date = QDateEdit()
        self.end_date.setCalendarPopup(True)
        m = self.start_date.date().toPython()
        next_month = (m.replace(day=28) + timedelta(days=4)).replace(day=1)
        self.end_date.setDate(next_month - timedelta(days=1))
        for t, w in [("开始", self.start_date), ("结束", self.end_date)]:
            top.addWidget(QLabel(t))
            top.addWidget(w)

        self.btn_reload_days = QPushButton("重建计划日历")
        self.btn_reload_days.clicked.connect(self.load_day_plan_table)
        self.btn_generate = QPushButton("自动排班")
        self.btn_generate.clicked.connect(self.generate_schedule)
        self.btn_export_csv = QPushButton("导出CSV")
        self.btn_export_csv.clicked.connect(self.on_export_csv)
        self.btn_export_xlsx = QPushButton("导出Excel")
        self.btn_export_xlsx.clicked.connect(self.on_export_xlsx)
        for b in [self.btn_reload_days, self.btn_generate, self.btn_export_csv, self.btn_export_xlsx]:
            top.addWidget(b)
        top.addStretch(1)
        layout.addLayout(top)

        self.tabs = QTabWidget()
        layout.addWidget(self.tabs, 1)

        self.tab_plan = QWidget()
        self.tab_people = QWidget()
        self.tab_leave = QWidget()
        self.tab_summary = QWidget()
        self.tabs.addTab(self.tab_plan, "值班计划")
        self.tabs.addTab(self.tab_people, "人员管理")
        self.tabs.addTab(self.tab_leave, "请假管理")
        self.tabs.addTab(self.tab_summary, "统计")

        self._build_plan_tab()
        self._build_people_tab()
        self._build_leave_tab()
        self._build_summary_tab()

    def _build_plan_tab(self):
        layout = QVBoxLayout(self.tab_plan)

        top = QHBoxLayout()
        left = QVBoxLayout()
        right = QVBoxLayout()

        self.plan_calendar = PlanCalendarWidget()
        self.plan_calendar.clicked.connect(self.on_plan_calendar_clicked)
        self.plan_calendar.setMinimumSize(760, 520)
        left.addWidget(QLabel("值班计划日历（点击日期进行配置）"))
        left.addWidget(self.plan_calendar)

        self.plan_date_label = QLabel("当前日期: -")
        self.plan_default_label = QLabel("默认类型: -")
        right.addWidget(self.plan_date_label)
        right.addWidget(self.plan_default_label)

        right.addWidget(QLabel("日期标签"))
        self.plan_tag_combo = QComboBox()
        self.plan_tag_combo.addItems(["工作日", "特殊日期"])
        self.plan_tag_combo.currentTextChanged.connect(self.save_selected_plan_day)
        right.addWidget(self.plan_tag_combo)

        right.addWidget(QLabel("特殊日期双流主班（人工指定）"))
        self.plan_leader_combo = QComboBox()
        self.plan_leader_combo.currentIndexChanged.connect(self.save_selected_plan_day)
        right.addWidget(self.plan_leader_combo)

        self.plan_day_tip = QLabel("提示: 特殊日期需指定双流主班（大队长/副大队长）")
        right.addWidget(self.plan_day_tip)
        right.addStretch(1)

        top.addLayout(left, 3)
        top.addLayout(right, 1)
        layout.addLayout(top, 2)

        self.day_plan_table = QTableWidget(0, 4)
        self.day_plan_table.setHorizontalHeaderLabels(["日期", "默认", "标签", "特殊日双流主班(人工)"])
        self.day_plan_table.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        layout.addWidget(self.day_plan_table, 2)

        self.schedule_table = QTableWidget(0, 6)
        self.schedule_table.setHorizontalHeaderLabels(["日期", "岗位", "姓名", "大队", "中队", "人工"])
        self.schedule_table.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        layout.addWidget(self.schedule_table, 3)

        self.log_box = QPlainTextEdit()
        self.log_box.setReadOnly(True)
        layout.addWidget(self.log_box, 1)

    def _build_people_tab(self):
        layout = QVBoxLayout(self.tab_people)
        form = QHBoxLayout()
        self.in_name = QLineEdit()
        self.in_team = QComboBox(); self.in_team.addItems(["二", "三", "四"])
        self.in_squad = QComboBox(); self.in_squad.addItems(["一中队", "二中队", "三中队", "四中队"])
        self.in_group = QComboBox(); self.in_group.addItems(["空勤", "地勤"])
        self.in_role = QComboBox(); self.in_role.addItems(["大队长", "副大队长_空勤", "副大队长_地勤", "党总支书记", "中队长", "副中队长", "中队书记"])
        self.in_participate = QCheckBox("参与排班"); self.in_participate.setChecked(True)
        self.btn_add_person = QPushButton("添加人员")
        self.btn_add_person.clicked.connect(self.add_person)
        for t, w in [("姓名", self.in_name), ("大队", self.in_team), ("中队", self.in_squad), ("类别", self.in_group), ("职责", self.in_role)]:
            form.addWidget(QLabel(t)); form.addWidget(w)
        form.addWidget(self.in_participate)
        form.addWidget(self.btn_add_person)
        layout.addLayout(form)

        self.people_table = PeopleTableWidget(self.persist_people_order, 0, 9)
        self.people_table.setHorizontalHeaderLabels(["拖动", "ID", "姓名", "大队", "中队", "类别", "职责", "参与排班", "操作"])
        self.people_table.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        self.people_table.horizontalHeader().setSectionResizeMode(0, QHeaderView.ResizeToContents)
        self.people_table.horizontalHeader().setSectionResizeMode(1, QHeaderView.ResizeToContents)
        self.people_table.setEditTriggers(QAbstractItemView.NoEditTriggers)
        self.people_table.setSelectionBehavior(QAbstractItemView.SelectRows)
        self.people_table.setDragDropMode(QAbstractItemView.InternalMove)
        self.people_table.setDragEnabled(True)
        self.people_table.setDropIndicatorShown(True)
        self.people_table.setDragDropOverwriteMode(False)
        self.people_table.setDefaultDropAction(Qt.MoveAction)
        layout.addWidget(self.people_table)

    def _build_leave_tab(self):
        layout = QVBoxLayout(self.tab_leave)
        top = QHBoxLayout()
        self.leave_person = QComboBox()
        self.leave_person.currentIndexChanged.connect(self.refresh_leave_list)
        top.addWidget(QLabel("选择人员"))
        top.addWidget(self.leave_person)
        top.addWidget(QLabel("日历点选切换请假/取消请假"))
        self.btn_remove_selected_leave = QPushButton("删除选中请假")
        self.btn_remove_selected_leave.clicked.connect(self.remove_selected_leave)
        self.btn_clear_all_leave = QPushButton("删除该人员全部请假")
        self.btn_clear_all_leave.clicked.connect(self.clear_all_leave_for_person)
        top.addWidget(self.btn_remove_selected_leave)
        top.addWidget(self.btn_clear_all_leave)
        top.addStretch(1)
        layout.addLayout(top)

        self.calendar_leave = QCalendarWidget()
        self.calendar_leave.clicked.connect(self.toggle_leave_on_calendar)
        layout.addWidget(self.calendar_leave, 3)

        self.leave_list = QListWidget()
        layout.addWidget(self.leave_list, 2)

        self.leave_message_box = QPlainTextEdit()
        self.leave_message_box.setReadOnly(True)
        self.leave_message_box.setPlaceholderText("请假操作提示会显示在这里")
        self.leave_message_box.setStyleSheet("QPlainTextEdit { font-size: 16px; }")
        layout.addWidget(self.leave_message_box, 1)

    def _build_summary_tab(self):
        layout = QVBoxLayout(self.tab_summary)
        self.summary_table = QTableWidget(0, 5)
        self.summary_table.setHorizontalHeaderLabels(["姓名", "大队", "中队", "月度次数", "年度次数(当前库)"])
        self.summary_table.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        layout.addWidget(self.summary_table)

    def _default_tag(self, d: date) -> str:
        if chinese_calendar is not None:
            try:
                if chinese_calendar.is_holiday(d) or d.weekday() >= 5:
                    return "SPECIAL"
                return "WORKDAY"
            except Exception:
                pass
        return "SPECIAL" if d.weekday() >= 5 else "WORKDAY"

    def _paint_plan_calendar(self):
        start = self.start_date.date().toPython()
        end = self.end_date.date().toPython()
        manual_tags = self.repo.load_day_tags()

        fmt_reset = QTextCharFormat()
        fmt_holiday = QTextCharFormat()
        fmt_holiday.setForeground(QColor("#ff0000"))  # 法定节假日红字
        fmt_weekend = QTextCharFormat()
        fmt_weekend.setForeground(QColor("#ff0000"))  # 周末红字
        fmt_manual_special = QTextCharFormat()
        fmt_manual_special.setForeground(QColor("#ff0000"))  # 人工改为特殊日期红字
        fmt_manual_workday = QTextCharFormat()
        fmt_manual_workday.setForeground(QColor("#000000"))  # 人工改为工作日黑字
        fmt_selected = QTextCharFormat()
        fmt_selected.setForeground(QColor("#ff0000"))
        fmt_selected.setFontWeight(700)

        d = start
        while d <= end:
            qd = QDate.fromString(d.isoformat(), "yyyy-MM-dd")
            self.plan_calendar.setDateTextFormat(qd, fmt_reset)
            d += timedelta(days=1)

        d = start
        while d <= end:
            default_tag = self._default_tag(d)
            tagged = manual_tags.get(d)
            qd = QDate.fromString(d.isoformat(), "yyyy-MM-dd")
            if tagged is not None and tagged != default_tag:
                if tagged == "SPECIAL":
                    self.plan_calendar.setDateTextFormat(qd, fmt_manual_special)
                else:
                    self.plan_calendar.setDateTextFormat(qd, fmt_manual_workday)
            else:
                if chinese_calendar is not None:
                    try:
                        if chinese_calendar.is_holiday(d):
                            self.plan_calendar.setDateTextFormat(qd, fmt_holiday)
                        elif d.weekday() >= 5:
                            self.plan_calendar.setDateTextFormat(qd, fmt_weekend)
                    except Exception:
                        if d.weekday() >= 5:
                            self.plan_calendar.setDateTextFormat(qd, fmt_weekend)
                elif d.weekday() >= 5:
                    self.plan_calendar.setDateTextFormat(qd, fmt_weekend)
            d += timedelta(days=1)
        selected_qd = self.plan_calendar.selectedDate()
        if selected_qd.isValid():
            self.plan_calendar.setDateTextFormat(selected_qd, fmt_selected)

    def refresh_people(self):
        self.employees = self.repo.load_employees()
        self.people_table.setRowCount(len(self.employees))
        self.leave_person.clear()
        self.people_table.blockSignals(True)
        for i, e in enumerate(self.employees):
            drag_item = QTableWidgetItem("☰")
            drag_item.setTextAlignment(Qt.AlignCenter)
            drag_item.setFlags(Qt.ItemIsEnabled | Qt.ItemIsSelectable | Qt.ItemIsDragEnabled)
            self.people_table.setItem(i, 0, drag_item)
            id_item = QTableWidgetItem(str(e.id))
            id_item.setFlags(Qt.ItemIsEnabled | Qt.ItemIsSelectable)
            self.people_table.setItem(i, 1, id_item)
            self.people_table.setItem(i, 2, QTableWidgetItem(e.name))

            cb_team = QComboBox(); cb_team.addItems(["二", "三", "四"]); cb_team.setCurrentText(e.team)
            cb_squad = QComboBox(); cb_squad.addItems(["一中队", "二中队", "三中队", "四中队"]); cb_squad.setCurrentText(e.squad)
            cb_group = QComboBox(); cb_group.addItems(["空勤", "地勤"]); cb_group.setCurrentText(e.duty_group)
            cb_role = QComboBox(); cb_role.addItems(["大队长", "副大队长_空勤", "副大队长_地勤", "党总支书记", "中队长", "副中队长", "中队书记"]); cb_role.setCurrentText(e.role)
            cb_team.currentTextChanged.connect(lambda _, eid=e.id, a=cb_team, b=cb_squad, c=cb_group, d=cb_role: self.update_person_meta(eid, a, b, c, d))
            cb_squad.currentTextChanged.connect(lambda _, eid=e.id, a=cb_team, b=cb_squad, c=cb_group, d=cb_role: self.update_person_meta(eid, a, b, c, d))
            cb_group.currentTextChanged.connect(lambda _, eid=e.id, a=cb_team, b=cb_squad, c=cb_group, d=cb_role: self.update_person_meta(eid, a, b, c, d))
            cb_role.currentTextChanged.connect(lambda _, eid=e.id, a=cb_team, b=cb_squad, c=cb_group, d=cb_role: self.update_person_meta(eid, a, b, c, d))
            self.people_table.setCellWidget(i, 3, cb_team)
            self.people_table.setCellWidget(i, 4, cb_squad)
            self.people_table.setCellWidget(i, 5, cb_group)
            self.people_table.setCellWidget(i, 6, cb_role)

            ck = QCheckBox(); ck.setChecked(e.participate)
            ck.stateChanged.connect(lambda state, eid=e.id: self.repo.set_participate(eid, bool(state)))
            self.people_table.setCellWidget(i, 7, ck)
            btn = QPushButton("删除")
            btn.clicked.connect(lambda _, eid=e.id: self.delete_person(eid))
            self.people_table.setCellWidget(i, 8, btn)
            self.leave_person.addItem(f"{e.id}-{e.name}-{e.team}-{e.squad}", e.id)
        self.people_table.blockSignals(False)

        self._reload_plan_day_leaders()

    def update_person_meta(self, employee_id: int, team_box: QComboBox, squad_box: QComboBox, group_box: QComboBox, role_box: QComboBox):
        self.repo.update_employee_meta(
            employee_id,
            team_box.currentText(),
            squad_box.currentText(),
            group_box.currentText(),
            role_box.currentText(),
        )
        self.employees = self.repo.load_employees()
        self._reload_plan_day_leaders()

    def persist_people_order(self):
        ordered_ids = []
        for i in range(self.people_table.rowCount()):
            item = self.people_table.item(i, 1)
            if item is not None:
                ordered_ids.append(int(item.text()))
        if ordered_ids:
            self.repo.reorder_employees(ordered_ids)

    def _reload_plan_day_leaders(self):
        self.plan_leader_combo.blockSignals(True)
        self.plan_leader_combo.clear()
        self.plan_leader_combo.addItem("(空)", None)
        for e in [x for x in self.employees if x.role in LEADER_ROLES]:
            self.plan_leader_combo.addItem(f"{e.name}-{e.team}-{e.squad}-{e.role}", e.id)
        self.plan_leader_combo.blockSignals(False)

    def add_person(self):
        name = self.in_name.text().strip()
        if not name:
            QMessageBox.warning(self, "提示", "请输入姓名")
            return
        self.repo.add_employee(name, self.in_team.currentText(), self.in_squad.currentText(), self.in_group.currentText(), self.in_role.currentText(), self.in_participate.isChecked())
        self.in_name.clear()
        self.refresh_people()
        self.load_day_plan_table()

    def delete_person(self, eid: int):
        self.repo.delete_employee(eid)
        self.refresh_people()
        self.load_day_plan_table()

    def toggle_leave_on_calendar(self):
        eid = self.leave_person.currentData()
        if eid is None:
            return
        d = self.calendar_leave.selectedDate().toPython()
        is_leave = self.repo.toggle_leave(eid, d)
        self.refresh_leave_list()
        person_label = self.leave_person.currentText().split("-", 2)[1] if "-" in self.leave_person.currentText() else self.leave_person.currentText()
        action = "新增请假" if is_leave else "取消请假"
        self.leave_message_box.appendPlainText(f"{action}: {person_label} - {d.isoformat()}")

    def refresh_leave_list(self):
        self.leave_list.clear()
        eid = self.leave_person.currentData()
        if eid is None:
            return
        for d in sorted(dd for pid, dd in self.repo.load_leaves() if pid == eid):
            self.leave_list.addItem(d.isoformat())

    def remove_selected_leave(self):
        eid = self.leave_person.currentData()
        item = self.leave_list.currentItem()
        if eid is None or item is None:
            return
        d = date.fromisoformat(item.text())
        self.repo.remove_leave(eid, d)
        self.refresh_leave_list()
        person_label = self.leave_person.currentText().split("-", 2)[1] if "-" in self.leave_person.currentText() else self.leave_person.currentText()
        self.leave_message_box.appendPlainText(f"删除请假: {person_label} - {d.isoformat()}")

    def clear_all_leave_for_person(self):
        eid = self.leave_person.currentData()
        if eid is None:
            return
        person_label = self.leave_person.currentText().split("-", 2)[1] if "-" in self.leave_person.currentText() else self.leave_person.currentText()
        self.repo.clear_leaves_for_employee(eid)
        self.refresh_leave_list()
        self.leave_message_box.appendPlainText(f"删除全部请假: {person_label}")

    def on_plan_calendar_clicked(self):
        self.current_plan_date = self.plan_calendar.selectedDate().toPython()
        self.load_plan_day_editor()

    def load_plan_day_editor(self):
        if self.current_plan_date is None:
            return
        d = self.current_plan_date
        tag = self.repo.load_day_tags().get(d, self._default_tag(d))
        manual = self.repo.load_manual_assignments("SL_MAIN")

        self.plan_date_label.setText(f"当前日期: {d.isoformat()}")
        default_label = "特殊日期" if self._default_tag(d) == "SPECIAL" else "工作日"
        self.plan_default_label.setText(f"默认类型: {default_label}")

        self.plan_tag_combo.blockSignals(True)
        self.plan_tag_combo.setCurrentText("特殊日期" if tag == "SPECIAL" else "工作日")
        self.plan_tag_combo.blockSignals(False)

        self.plan_leader_combo.blockSignals(True)
        idx = self.plan_leader_combo.findData(manual.get(d))
        self.plan_leader_combo.setCurrentIndex(idx if idx >= 0 else 0)
        self.plan_leader_combo.blockSignals(False)

    def save_selected_plan_day(self):
        if self.current_plan_date is None:
            return
        d = self.current_plan_date
        tag = "SPECIAL" if self.plan_tag_combo.currentText() == "特殊日期" else "WORKDAY"
        self.repo.set_day_tag(d, tag)
        eid = self.plan_leader_combo.currentData()
        if eid is not None:
            self.repo.set_manual_assignment(d, "SL_MAIN", int(eid))
        self.load_day_plan_table()
        self._paint_plan_calendar()

    def load_day_plan_table(self):
        start = self.start_date.date().toPython()
        end = self.end_date.date().toPython()
        tags = self.repo.load_day_tags()
        manual = self.repo.load_manual_assignments("SL_MAIN")

        days = []
        d = start
        while d <= end:
            days.append(d)
            d += timedelta(days=1)

        by_id = {e.id: e for e in self.employees}
        self.day_plan_table.setRowCount(len(days))
        manual_name_map: dict[date, str] = {}
        for i, d in enumerate(days):
            default_tag = self._default_tag(d)
            tag = tags.get(d, default_tag)
            default_tag_text = "特殊日期" if default_tag == "SPECIAL" else "工作日"
            tag_text = "特殊日期" if tag == "SPECIAL" else "工作日"
            leader_text = ""
            if d in manual and manual[d] in by_id:
                e = by_id[manual[d]]
                leader_text = f"{e.name}({e.team}-{e.squad})"
                manual_name_map[d] = e.name

            self.day_plan_table.setItem(i, 0, QTableWidgetItem(d.isoformat()))
            self.day_plan_table.setItem(i, 1, QTableWidgetItem(default_tag_text))
            self.day_plan_table.setItem(i, 2, QTableWidgetItem(tag_text))
            self.day_plan_table.setItem(i, 3, QTableWidgetItem(leader_text))

        if self.current_plan_date is None:
            self.current_plan_date = start
        self.load_plan_day_editor()
        self._paint_plan_calendar()
        self.plan_calendar.set_manual_name_map(manual_name_map)

    def generate_schedule(self):
        self.employees = self.repo.load_employees()
        result = SchedulerEngine().solve(
            self.employees,
            self.start_date.date().toPython(),
            self.end_date.date().toPython(),
            self.repo.load_leaves(),
            self.repo.load_day_tags(),
            self.repo.load_manual_assignments("SL_MAIN"),
        )
        self.current_assignments = result.assignments
        self.log_box.setPlainText("\n".join(result.logs) if result.logs else "排班完成")
        self.render_assignments()
        self.render_summary()

    def render_assignments(self):
        by_id = {e.id: e for e in self.employees}
        rows = sorted(self.current_assignments, key=lambda x: (x.work_date, x.position))
        self.schedule_table.setRowCount(len(rows))
        for i, a in enumerate(rows):
            e = by_id[a.employee_id]
            self.schedule_table.setItem(i, 0, QTableWidgetItem(a.work_date.isoformat()))
            self.schedule_table.setItem(i, 1, QTableWidgetItem(POSITION_LABELS.get(a.position, a.position)))
            self.schedule_table.setItem(i, 2, QTableWidgetItem(e.name))
            self.schedule_table.setItem(i, 3, QTableWidgetItem(e.team))
            self.schedule_table.setItem(i, 4, QTableWidgetItem(e.squad))
            self.schedule_table.setItem(i, 5, QTableWidgetItem("是" if a.manual else "否"))

    def render_summary(self):
        month_cnt = defaultdict(int)
        year_cnt = defaultdict(int)
        for a in self.current_assignments:
            month_cnt[a.employee_id] += 1
            year_cnt[a.employee_id] += 1

        self.summary_table.setRowCount(len(self.employees))
        for i, e in enumerate(self.employees):
            self.summary_table.setItem(i, 0, QTableWidgetItem(e.name))
            self.summary_table.setItem(i, 1, QTableWidgetItem(e.team))
            self.summary_table.setItem(i, 2, QTableWidgetItem(e.squad))
            self.summary_table.setItem(i, 3, QTableWidgetItem(str(month_cnt[e.id])))
            self.summary_table.setItem(i, 4, QTableWidgetItem(str(year_cnt[e.id])))

    def on_export_csv(self):
        if not self.current_assignments:
            return
        path, _ = QFileDialog.getSaveFileName(self, "导出CSV", "schedule.csv", "CSV (*.csv)")
        if path:
            export_csv(Path(path), self.current_assignments, self.employees)

    def on_export_xlsx(self):
        if not self.current_assignments:
            return
        path, _ = QFileDialog.getSaveFileName(self, "导出Excel", "schedule.xlsx", "Excel (*.xlsx)")
        if path:
            export_excel(Path(path), self.current_assignments, self.employees)


def run_app() -> None:
    app = QApplication(sys.argv)
    app.setApplicationName("Scheduler App")
    db_path = Path.home() / "AppData" / "Local" / "SchedulerApp" / "scheduler.db"
    window = MainWindow(db_path)
    window.show()
    sys.exit(app.exec())
