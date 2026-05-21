from pathlib import Path

from PyInstaller.__main__ import run


if __name__ == "__main__":
    project_root = Path(__file__).resolve().parents[1]
    work_path = project_root / "build_exe_tmp"
    dist_path = project_root / "dist_exe_tmp"
    run(
        [
            "--noconfirm",
            "--clean",
            "--onedir",
            "--name",
            "SchedulerApp",
            "--windowed",
            "--collect-submodules",
            "PySide6",
            "--hidden-import",
            "openpyxl",
            "--workpath",
            str(work_path),
            "--distpath",
            str(dist_path),
            str(project_root / "src" / "scheduler_app" / "__main__.py"),
        ]
    )
