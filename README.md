# 排班系统（Windows 离线版）

## 已实现功能（按 AGENTS.md 首版落地）
- 岗位：`SL_MAIN` / `SL_AIR` / `SL_GROUND` / `TF_MAIN` / `TF_GROUND(特殊日)` / `FLEET_LEAD`
- 人员管理：新增/删除人员，维护大队、空勤/地勤、角色、是否参与排班
- 请假模块：人员+日历点击即可切换请假（请假日不参与排班）
- 日期标注：将日期标注为 `WORKDAY` 或 `SPECIAL`
- 特殊日双流主班：支持人工指定（仅大队长/副大队长）
- 自动排班：按规则顺序生成岗位，并记录冲突日志
- 导出：CSV / Excel
- 统计：个人月度统计、大队月度统计（年度位先用当前库累计占位）

## 运行
```powershell
python -m venv .venv
.\.venv\Scripts\Activate.ps1
pip install -r requirements.txt -c constraints.txt
python -m scheduler_app
```

## 打包 exe
```powershell
python scripts\build_exe.py
```
输出：`dist\SchedulerApp\SchedulerApp.exe`

## 说明
- 当前为 `PySide6` 桌面版（不是浏览器网页）。
- 数据库路径：`%USERPROFILE%\AppData\Local\SchedulerApp\scheduler.db`
- 如需改成网页形态（前后端分离）可在此规则引擎基础上迁移。
