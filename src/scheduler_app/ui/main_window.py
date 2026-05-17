from __future__ import annotations

import json
import sys
from collections import defaultdict
from datetime import date, timedelta
from pathlib import Path
from typing import Callable

from openpyxl import load_workbook
from PySide6.QtCore import QDate, Qt
from PySide6.QtGui import QColor, QPainter, QTextCharFormat
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

from scheduler_app.core.engine import SchedulerEngine
from scheduler_app.core.models import Assignment, POSITION_LABELS, SPECIAL_SL_MAIN_ROLES
from scheduler_app.data.repository import SchedulerRepository
from scheduler_app.services.exporter import export_csv, export_excel

try:
    import chinese_calendar
except Exception:
    chinese_calendar = None


RESULT_POSITIONS = ["SL_MAIN", "SL_AIR", "SL_GROUND", "TF_MAIN", "TF_GROUND", "FLEET_LEAD"]
RESULT_COL_POS = {3: "SL_MAIN", 4: "SL_AIR", 5: "SL_GROUND", 6: "TF_MAIN", 7: "TF_GROUND", 8: "FLEET_LEAD"}
WEEKDAY_MAP = {0: "周一", 1: "周二", 2: "周三", 3: "周四", 4: "周五", 5: "周六", 6: "周日"}


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
        self.resize(1520, 940)
        self.repo = SchedulerRepository(db_path)
        self.repo.seed_if_empty()
        self.current_assignments: list[Assignment] = []
        self.current_logs: list[str] = []
        self.employees = self.repo.load_employees()
        self.current_plan_date: date | None = None
        self.history_tail: list[dict] = []
        self.history_person_year_totals: dict[int, int] = {}
        self.history_team_year_totals: dict[str, int] = {}
        self.rerun_seed: int = 0

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
        self.tab_result = QWidget()
        self.tab_summary = QWidget()
        self.tabs.addTab(self.tab_plan, "值班计划")
        self.tabs.addTab(self.tab_people, "人员管理")
        self.tabs.addTab(self.tab_leave, "请假管理")
        self.tabs.addTab(self.tab_result, "排班结果")
        self.tabs.addTab(self.tab_summary, "统计")
        self._build_plan_tab()
        self._build_people_tab()
        self._build_leave_tab()
        self._build_result_tab()
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
        btns = QHBoxLayout()
        self.btn_import_history = QPushButton("导入历史排班")
        self.btn_import_history.clicked.connect(self.import_history_from_file)
        self.btn_use_saved_history = QPushButton("引用排班数据库")
        self.btn_use_saved_history.clicked.connect(self.use_history_from_db)
        self.btn_clear_saved_history = QPushButton("清除历史数据")
        self.btn_clear_saved_history.clicked.connect(self.clear_saved_history_data)
        btns.addWidget(self.btn_import_history)
        btns.addWidget(self.btn_use_saved_history)
        btns.addWidget(self.btn_clear_saved_history)
        left.addLayout(btns)
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
        self.day_plan_table.setMinimumHeight(360)
        layout.addWidget(self.day_plan_table, 4)
        self.log_box = QPlainTextEdit()
        self.log_box.setReadOnly(True)
        layout.addWidget(self.log_box, 1)

    def _build_people_tab(self):
        layout = QVBoxLayout(self.tab_people)
        form = QHBoxLayout()
        self.in_name = QLineEdit()
        self.in_team = QComboBox(); self.in_team.addItems(["二", "三", "四"])
        self.in_squad = QComboBox(); self.in_squad.addItems(["直属", "一中队", "二中队", "三中队", "四中队"])
        self.in_group = QComboBox(); self.in_group.addItems(["空勤", "地勤"])
        self.in_role = QComboBox(); self.in_role.addItems(["大队长", "副大队长_空勤", "副大队长_地勤", "党总支书记", "中队长", "副中队长", "中队书记"])
        self.in_participate = QCheckBox("参与排班"); self.in_participate.setChecked(True)
        self.btn_add_person = QPushButton("添加人员")
        self.btn_add_person.clicked.connect(self.add_person)
        self.btn_import_people = QPushButton("从Excel替换人员")
        self.btn_import_people.clicked.connect(self.import_people_from_excel)
        for t, w in [("姓名", self.in_name), ("大队", self.in_team), ("中队", self.in_squad), ("类别", self.in_group), ("职责", self.in_role)]:
            form.addWidget(QLabel(t)); form.addWidget(w)
        form.addWidget(self.in_participate)
        form.addWidget(self.btn_add_person)
        form.addWidget(self.btn_import_people)
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
        self.btn_clear_everyone_leave = QPushButton("删除所有人员所有请假")
        self.btn_clear_everyone_leave.clicked.connect(self.clear_all_leave_for_everyone)
        top.addWidget(self.btn_remove_selected_leave)
        top.addWidget(self.btn_clear_all_leave)
        top.addWidget(self.btn_clear_everyone_leave)
        top.addStretch(1)
        layout.addLayout(top)
        self.calendar_leave = QCalendarWidget()
        self.calendar_leave.clicked.connect(self.toggle_leave_on_calendar)
        layout.addWidget(self.calendar_leave, 3)
        self.leave_table = QTableWidget(0, 2)
        self.leave_table.setHorizontalHeaderLabels(["请假信息", "操作"])
        self.leave_table.horizontalHeader().setSectionResizeMode(0, QHeaderView.Stretch)
        self.leave_table.horizontalHeader().setSectionResizeMode(1, QHeaderView.ResizeToContents)
        self.leave_table.setSelectionBehavior(QAbstractItemView.SelectRows)
        self.leave_table.setEditTriggers(QAbstractItemView.NoEditTriggers)
        layout.addWidget(self.leave_table, 2)
        self.leave_message_box = QPlainTextEdit()
        self.leave_message_box.setReadOnly(True)
        self.leave_message_box.setPlaceholderText("请假操作提示会显示在这里")
        layout.addWidget(self.leave_message_box, 1)

    def _build_result_tab(self):
        layout = QVBoxLayout(self.tab_result)
        self.result_status_label = QLabel("结果状态：未生成")
        layout.addWidget(self.result_status_label)
        ops = QHBoxLayout()
        self.btn_rerun_with_mod = QPushButton("根据修改重新排班")
        self.btn_rerun_with_mod.clicked.connect(self.regenerate_with_overrides)
        self.btn_rerun_plan_only = QPushButton("全部重排")
        self.btn_rerun_plan_only.clicked.connect(self.regenerate_plan_only)
        self.btn_save_result = QPushButton("保存数据")
        self.btn_save_result.clicked.connect(self.save_current_schedule_data)
        ops.addWidget(self.btn_rerun_with_mod)
        ops.addWidget(self.btn_rerun_plan_only)
        ops.addWidget(self.btn_save_result)
        ops.addStretch(1)
        layout.addLayout(ops)
        self.result_table = QTableWidget(0, 11)
        self.result_table.setHorizontalHeaderLabels(["日期", "星期", "是否周末/节假日", "双流主班", "双流副班(空勤)", "双流副班(地勤)", "天府主班", "天府副班", "机队总负责", "双流主班大队", "天府主班大队"])
        self.result_table.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        self.result_table.setEditTriggers(QAbstractItemView.NoEditTriggers)
        layout.addWidget(self.result_table, 1)

    def _build_summary_tab(self):
        layout = QVBoxLayout(self.tab_summary)
        self.summary_table = QTableWidget(0, 5)
        self.summary_table.setHorizontalHeaderLabels(["姓名", "大队", "中队", "月度次数", "年度次数(累计)"])
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
        fmt_weekend = QTextCharFormat()
        fmt_weekend.setForeground(QColor("#ff0000"))
        fmt_manual_special = QTextCharFormat()
        fmt_manual_special.setForeground(QColor("#ff0000"))
        fmt_manual_workday = QTextCharFormat()
        fmt_manual_workday.setForeground(QColor("#000000"))
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
                self.plan_calendar.setDateTextFormat(qd, fmt_manual_special if tagged == "SPECIAL" else fmt_manual_workday)
            elif d.weekday() >= 5:
                self.plan_calendar.setDateTextFormat(qd, fmt_weekend)
            d += timedelta(days=1)

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
            cb_squad = QComboBox(); cb_squad.addItems(["直属", "一中队", "二中队", "三中队", "四中队"]); cb_squad.setCurrentText(e.squad)
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
        self.repo.update_employee_meta(employee_id, team_box.currentText(), squad_box.currentText(), group_box.currentText(), role_box.currentText())
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
        for e in [x for x in self.employees if x.role in SPECIAL_SL_MAIN_ROLES]:
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

    def import_people_from_excel(self):
        file_path, _ = QFileDialog.getOpenFileName(self, "选择人员Excel文件", "", "Excel (*.xlsx *.xlsm)")
        if not file_path:
            return
        try:
            count = self.repo.replace_employees_from_xlsx(Path(file_path))
            self.refresh_people()
            self.refresh_leave_list()
            self.load_day_plan_table()
            QMessageBox.information(self, "导入完成", f"已替换人员，共导入 {count} 条。")
        except Exception as e:
            QMessageBox.critical(self, "导入失败", str(e))

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
        self.leave_message_box.appendPlainText(f"{person_label}，{d.year}/{d.month}/{d.day}（{action}）")

    def refresh_leave_list(self):
        self.leave_table.setRowCount(0)
        eid = self.leave_person.currentData()
        if eid is None:
            return
        person_label = self.leave_person.currentText().split("-", 2)[1] if "-" in self.leave_person.currentText() else self.leave_person.currentText()
        dates = sorted(dd for pid, dd in self.repo.load_leaves() if pid == eid)
        self.leave_table.setRowCount(len(dates))
        for i, d in enumerate(dates):
            item = QTableWidgetItem(f"{person_label}，{d.year}/{d.month}/{d.day}")
            item.setData(Qt.UserRole, d.isoformat())
            self.leave_table.setItem(i, 0, item)
            btn = QPushButton("删除")
            btn.clicked.connect(lambda _, dd=d, ee=eid: self.remove_leave_item(ee, dd))
            self.leave_table.setCellWidget(i, 1, btn)

    def remove_selected_leave(self):
        eid = self.leave_person.currentData()
        row = self.leave_table.currentRow()
        if eid is None or row < 0:
            return
        d = date.fromisoformat(self.leave_table.item(row, 0).data(Qt.UserRole))
        self.repo.remove_leave(eid, d)
        self.refresh_leave_list()

    def remove_leave_item(self, eid: int, d: date):
        self.repo.remove_leave(eid, d)
        self.refresh_leave_list()

    def clear_all_leave_for_person(self):
        eid = self.leave_person.currentData()
        if eid is None:
            return
        self.repo.clear_leaves_for_employee(eid)
        self.refresh_leave_list()

    def clear_all_leave_for_everyone(self):
        self.repo.clear_all_leaves()
        self.refresh_leave_list()

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
        else:
            self.repo.clear_manual_assignment(d, "SL_MAIN")
        self.load_day_plan_table()

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
            leader_text = ""
            if d in manual and manual[d] in by_id:
                e = by_id[manual[d]]
                leader_text = f"{e.name}({e.team}-{e.squad})"
                manual_name_map[d] = e.name
            self.day_plan_table.setItem(i, 0, QTableWidgetItem(d.isoformat()))
            self.day_plan_table.setItem(i, 1, QTableWidgetItem("特殊日期" if default_tag == "SPECIAL" else "工作日"))
            self.day_plan_table.setItem(i, 2, QTableWidgetItem("特殊日期" if tag == "SPECIAL" else "工作日"))
            self.day_plan_table.setItem(i, 3, QTableWidgetItem(leader_text))
        if self.current_plan_date is None:
            self.current_plan_date = start
        self.load_plan_day_editor()
        self._paint_plan_calendar()
        self.plan_calendar.set_manual_name_map(manual_name_map)

    def _collect_history_tail_from_xlsx(self, path: Path) -> list[dict]:
        wb = load_workbook(path, data_only=True)
        ws = wb["排班表"] if "排班表" in wb.sheetnames else wb.active
        header = [str(ws.cell(1, c).value or "").strip() for c in range(1, ws.max_column + 1)]
        idx = {h: i + 1 for i, h in enumerate(header)}
        required = ["日期", "双流主班", "天府主班", "机队总负责"]
        if any(h not in idx for h in required):
            raise ValueError("历史排班Excel缺少必要列")
        name_to_id = {e.name: e.id for e in self.employees}
        rows = []
        for r in range(2, ws.max_row + 1):
            d_raw = ws.cell(r, idx["日期"]).value
            if d_raw is None:
                continue
            d_val = d_raw.date() if hasattr(d_raw, "date") else date.fromisoformat(str(d_raw)[:10])
            rows.append((d_val, str(ws.cell(r, idx["双流主班"]).value or "").strip(), str(ws.cell(r, idx["天府主班"]).value or "").strip(), str(ws.cell(r, idx["机队总负责"]).value or "").strip()))
        rows.sort(key=lambda x: x[0])
        tail = rows[-3:]
        out = []
        for d_val, sl_name, tf_name, fleet_name in tail:
            if sl_name in name_to_id:
                out.append({"work_date": d_val.isoformat(), "position": "SL_MAIN", "employee_id": name_to_id[sl_name]})
            if tf_name in name_to_id:
                out.append({"work_date": d_val.isoformat(), "position": "TF_MAIN", "employee_id": name_to_id[tf_name]})
            if fleet_name in name_to_id:
                out.append({"work_date": d_val.isoformat(), "position": "FLEET_LEAD", "employee_id": name_to_id[fleet_name]})
        return out

    def _collect_year_stats_from_xlsx(self, path: Path) -> tuple[dict[int, int], dict[str, int]]:
        wb = load_workbook(path, data_only=True)
        person_year_totals: dict[int, int] = {}
        team_year_totals: dict[str, int] = {}
        employees_by_name = {e.name: e for e in self.employees}

        if "人员年度统计" in wb.sheetnames:
            ws = wb["人员年度统计"]
            for r in range(4, ws.max_row + 1):
                name = str(ws.cell(r, 1).value or "").strip()
                total = ws.cell(r, 4).value
                if name in employees_by_name and total not in (None, ""):
                    person_year_totals[employees_by_name[name].id] = int(total)

        if "大队年度统计" in wb.sheetnames:
            ws = wb["大队年度统计"]
            for r in range(4, ws.max_row + 1):
                team = str(ws.cell(r, 1).value or "").strip()
                total = ws.cell(r, 2).value
                if team and total not in (None, ""):
                    team_year_totals[team] = int(total)

        return person_year_totals, team_year_totals

    def import_history_from_file(self):
        file_path, _ = QFileDialog.getOpenFileName(self, "导入历史排班", "", "Excel (*.xlsx *.xlsm)")
        if not file_path:
            return
        try:
            self.history_tail = self._collect_history_tail_from_xlsx(Path(file_path))
            self.history_person_year_totals, self.history_team_year_totals = self._collect_year_stats_from_xlsx(Path(file_path))
            self.repo.save_history_import_cache("file", file_path, json.dumps(self.history_tail, ensure_ascii=False))
            QMessageBox.information(self, "成功", f"已导入历史尾部记录 {len(self.history_tail)} 条")
        except Exception as e:
            QMessageBox.critical(self, "导入失败", str(e))

    def _previous_month_key(self) -> str:
        s = self.start_date.date().toPython().replace(day=1)
        prev_end = s - timedelta(days=1)
        return f"{prev_end.year:04d}-{prev_end.month:02d}"

    def use_history_from_db(self):
        tail = self.repo.latest_saved_tail(self._previous_month_key())
        if not tail:
            QMessageBox.warning(self, "提示", "数据库中没有可引用的上月排班记录")
            return
        self.history_tail = tail
        year_key = self.start_date.date().toPython().strftime("%Y")
        self.history_person_year_totals = self.repo.yearly_totals(year_key)
        team_totals = defaultdict(int)
        by_id = {e.id: e for e in self.employees}
        for eid, total in self.history_person_year_totals.items():
            e = by_id.get(eid)
            if e is not None:
                team_totals[e.team] += total
        self.history_team_year_totals = dict(team_totals)
        self.repo.save_history_import_cache("db", self._previous_month_key(), json.dumps(self.history_tail, ensure_ascii=False))
        QMessageBox.information(self, "成功", f"已从数据库引用历史尾部记录 {len(self.history_tail)} 条")

    def clear_saved_history_data(self):
        reply = QMessageBox.question(
            self,
            "确认清除",
            "这会清除数据库中的排班结果、月度统计、年度统计和历史引用缓存，用于避免重复叠加。是否继续？",
            QMessageBox.Yes | QMessageBox.No,
            QMessageBox.No,
        )
        if reply != QMessageBox.Yes:
            return
        self.repo.clear_saved_schedule_stats()
        self.history_tail = []
        self.history_person_year_totals = {}
        self.history_team_year_totals = {}
        self.render_summary()
        QMessageBox.information(self, "成功", "已清除排班统计历史数据")

    def _run_schedule(self, use_result_overrides: bool, seed: int):
        self.current_assignments = []
        self.current_logs = []
        self.log_box.clear()
        self.employees = self.repo.load_employees()
        overrides = self.repo.load_result_overrides() if use_result_overrides else {}
        result = SchedulerEngine().solve(
            self.employees,
            self.start_date.date().toPython(),
            self.end_date.date().toPython(),
            self.repo.load_leaves(),
            self.repo.load_day_tags(),
            self.repo.load_manual_assignments("SL_MAIN"),
            history_tail=self.history_tail,
            manual_overrides=overrides,
            rerun_seed=seed,
            person_year_baseline=self.history_person_year_totals,
            team_year_baseline=self.history_team_year_totals,
        )
        self.current_assignments = result.assignments
        self.current_logs = result.logs
        self.log_box.setPlainText("\n".join(result.logs) if result.logs else "排班完成")
        self.render_result_table()
        self.render_summary()
        self.result_status_label.setText(
            f"结果状态：已刷新（共 {len(self.current_assignments)} 条岗位安排，日志 {len(self.current_logs)} 条）"
        )
        # 每次重排后切换到“排班结果”页，确保用户立即看到最新结果。
        self.tabs.setCurrentWidget(self.tab_result)
        self.result_table.viewport().update()
        self.summary_table.viewport().update()

    def generate_schedule(self):
        self.rerun_seed = 0
        self._run_schedule(use_result_overrides=True, seed=self.rerun_seed)

    def regenerate_with_overrides(self):
        self.rerun_seed += 1
        self._run_schedule(use_result_overrides=True, seed=self.rerun_seed)

    def regenerate_plan_only(self):
        self.repo.clear_all_result_overrides()
        self.rerun_seed += 1
        self._run_schedule(use_result_overrides=False, seed=self.rerun_seed)

    def _position_options(self, pos: str, d: date) -> list[tuple[str, int]]:
        leaves = self.repo.load_leaves()
        opts = []
        for e in self.employees:
            if not e.participate or (e.id, d) in leaves:
                continue
            if pos in ("SL_GROUND", "TF_GROUND") and e.duty_group != "地勤":
                continue
            if pos in ("SL_MAIN", "SL_AIR", "TF_MAIN") and e.duty_group != "空勤":
                continue
            opts.append((f"{e.name}({e.team}-{e.squad})", e.id))
        return opts

    def on_result_assignment_changed(self, d: date, pos: str, combo: QComboBox):
        eid = combo.currentData()
        if eid is None:
            self.repo.clear_result_override(d, pos)
        else:
            self.repo.save_result_override(d, pos, int(eid))

    def render_result_table(self):
        by_id = {e.id: e for e in self.employees}
        rows_by_date = defaultdict(dict)
        for a in self.current_assignments:
            rows_by_date[a.work_date][a.position] = a
        all_dates = sorted(rows_by_date)
        overrides = self.repo.load_result_overrides()
        self.result_table.setRowCount(len(all_dates))
        for i, d in enumerate(all_dates):
            self.result_table.setItem(i, 0, QTableWidgetItem(d.isoformat()))
            self.result_table.setItem(i, 1, QTableWidgetItem(WEEKDAY_MAP[d.weekday()]))
            special = self.repo.load_day_tags().get(d, self._default_tag(d)) == "SPECIAL"
            self.result_table.setItem(i, 2, QTableWidgetItem("是" if special else "否"))
            for c, pos in RESULT_COL_POS.items():
                current = rows_by_date[d].get(pos)
                combo = QComboBox()
                combo.addItem("(空)", None)
                for label, eid in self._position_options(pos, d):
                    combo.addItem(label, eid)
                selected = overrides.get((d, pos), current.employee_id if current else None)
                idx = combo.findData(selected)
                combo.setCurrentIndex(idx if idx >= 0 else 0)
                combo.currentIndexChanged.connect(lambda _, dd=d, pp=pos, cb=combo: self.on_result_assignment_changed(dd, pp, cb))
                self.result_table.setCellWidget(i, c, combo)
            sl = rows_by_date[d].get("SL_MAIN")
            tf = rows_by_date[d].get("TF_MAIN")
            self.result_table.setItem(i, 9, QTableWidgetItem(by_id[sl.employee_id].team if sl and sl.employee_id in by_id else ""))
            self.result_table.setItem(i, 10, QTableWidgetItem(by_id[tf.employee_id].team if tf and tf.employee_id in by_id else ""))

    def save_current_schedule_data(self):
        if not self.current_assignments:
            QMessageBox.warning(self, "提示", "请先生成排班结果")
            return
        month_key = self.start_date.date().toPython().strftime("%Y-%m")
        self.repo.save_schedule_and_stats(self.current_assignments, self.employees, month_key)
        QMessageBox.information(self, "成功", "排班结果与统计数据已保存")
        self.render_summary()

    def render_summary(self):
        month_cnt = defaultdict(int)
        for a in self.current_assignments:
            month_cnt[a.employee_id] += 1
        year_key = self.start_date.date().toPython().strftime("%Y")
        year_cnt = self.repo.yearly_totals(year_key)
        for eid, c in month_cnt.items():
            year_cnt[eid] = year_cnt.get(eid, 0) + c
        self.summary_table.setRowCount(len(self.employees))
        for i, e in enumerate(self.employees):
            self.summary_table.setItem(i, 0, QTableWidgetItem(e.name))
            self.summary_table.setItem(i, 1, QTableWidgetItem(e.team))
            self.summary_table.setItem(i, 2, QTableWidgetItem(e.squad))
            self.summary_table.setItem(i, 3, QTableWidgetItem(str(month_cnt.get(e.id, 0))))
            self.summary_table.setItem(i, 4, QTableWidgetItem(str(year_cnt.get(e.id, 0))))

    def _export_year_totals(self) -> tuple[dict[int, int], dict[str, int]]:
        month_cnt = defaultdict(int)
        for a in self.current_assignments:
            month_cnt[a.employee_id] += 1
        year_key = self.start_date.date().toPython().strftime("%Y")
        person_year_totals = self.repo.yearly_totals(year_key)
        for eid, cnt in month_cnt.items():
            person_year_totals[eid] = person_year_totals.get(eid, 0) + cnt
        team_year_totals = defaultdict(int)
        by_id = {e.id: e for e in self.employees}
        for eid, cnt in person_year_totals.items():
            e = by_id.get(eid)
            if e is not None:
                team_year_totals[e.team] += cnt
        return person_year_totals, dict(team_year_totals)

    def on_export_csv(self):
        if not self.current_assignments:
            return
        path, _ = QFileDialog.getSaveFileName(self, "导出CSV", "schedule.csv", "CSV (*.csv)")
        if path:
            person_year_totals, team_year_totals = self._export_year_totals()
            export_csv(
                Path(path),
                self.current_assignments,
                self.employees,
                person_year_totals=person_year_totals,
                team_year_totals=team_year_totals,
            )

    def on_export_xlsx(self):
        if not self.current_assignments:
            return
        path, _ = QFileDialog.getSaveFileName(self, "导出Excel", "schedule.xlsx", "Excel (*.xlsx)")
        if path:
            person_year_totals, team_year_totals = self._export_year_totals()
            export_excel(
                Path(path),
                self.current_assignments,
                self.employees,
                self.repo.load_day_tags(),
                person_year_totals=person_year_totals,
                team_year_totals=team_year_totals,
            )


def run_app() -> None:
    app = QApplication(sys.argv)
    app.setApplicationName("Scheduler App")
    db_path = Path.home() / "AppData" / "Local" / "SchedulerApp" / "scheduler.db"
    window = MainWindow(db_path)
    window.show()
    sys.exit(app.exec())
