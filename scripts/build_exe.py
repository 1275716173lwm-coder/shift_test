from pathlib import Path

from PyInstaller.__main__ import run


if __name__ == "__main__":
    project_root = Path(__file__).resolve().parents[1]
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
            str(project_root / "src" / "scheduler_app" / "__main__.py"),
        ]
    )
