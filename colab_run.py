"""
Google Colab entry point for the FSC Matrix pipeline.

Usage in a Colab cell:

    exec(open(
        "/content/drive/Shareddrives/FA Ops Europe: Rate Maintenance Team "
        "/Documents/AI Adoption RMT/RMT_Daikin/RMT_FSC/Daikin-FSC/colab_run.py"
    ).read())
"""

from __future__ import annotations

import sys
from pathlib import Path

COLAB_DRIVE_ROOT = Path(
    "/content/drive/Shareddrives/FA Ops Europe: Rate Maintenance Team "
    "/Documents/AI Adoption RMT/RMT_Daikin/RMT_FSC"
)
CODE_DIR = COLAB_DRIVE_ROOT / "Daikin-FSC"

sys.path.insert(0, str(CODE_DIR))

from config import ensure_workspace_dirs, is_colab, setup_paths


def mount_drive_if_needed() -> None:
    if not is_colab():
        return

    from google.colab import drive

    drive.mount("/content/drive")


def install_dependencies() -> None:
    try:
        import openpyxl  # noqa: F401
        import pandas  # noqa: F401
    except ImportError:
        import subprocess

        subprocess.check_call(
            ["pip", "install", "-q", "pandas>=2.0.0", "openpyxl>=3.1.0"]
        )


mount_drive_if_needed()
setup_paths(CODE_DIR)
install_dependencies()
ensure_workspace_dirs()

from pipeline import run_pipeline

print(f"Workspace root: {COLAB_DRIVE_ROOT}")
print(f"Code folder:    {CODE_DIR}")
print()

run_pipeline()
