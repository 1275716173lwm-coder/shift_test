from __future__ import annotations

import os
import shutil
import stat
from pathlib import Path

from PyInstaller.__main__ import run


def _clear_readonly(func, path, exc_info) -> None:
    try:
        os.chmod(path, stat.S_IWRITE)
        func(path)
    except Exception as exc:  # pragma: no cover - surfaces a clearer build error for local runs
        raise RuntimeError(f"无法清理旧的打包目录：{path}\n原始错误：{exc}") from exc


def _safe_rmtree(path: Path) -> None:
    if not path.exists():
        return
    shutil.rmtree(path, onexc=_clear_readonly)


if __name__ == "__main__":
    project_root = Path(__file__).resolve().parents[1]
    work_path = project_root / "build_exe_tmp"
    dist_path = project_root / "dist_exe_tmp"

    # 先清理旧的工作目录，避免 PyInstaller 在 --clean 阶段删除只读目录失败。
    _safe_rmtree(work_path / "SchedulerApp")

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
