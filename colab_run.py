"""
Google Colab entry point for the FSC Matrix pipeline.

Recommended usage in a Colab cell:

    exec(open("/content/Daikin-FSC/colab_run.py").read())

Or from Google Drive:

    exec(open(
        "/content/drive/Shareddrives/FA Ops Europe: Rate Maintenance Team "
        "/Documents/AI Adoption RMT/RMT_Daikin/RMT_FSC/Daikin-FSC/colab_run.py"
    ).read())

You can also run the pipeline directly after mounting Drive:

    exec(open("/content/Daikin-FSC/pipeline.py").read())
"""

from __future__ import annotations

import sys
from pathlib import Path

LOCAL_COLAB_CODE_DIR = Path("/content/Daikin-FSC")
COLAB_DRIVE_ROOT = Path(
    "/content/drive/Shareddrives/FA Ops Europe: Rate Maintenance Team "
    "/Documents/AI Adoption RMT/RMT_Daikin/RMT_FSC"
)


def find_code_dir() -> Path:
    candidates = [
        LOCAL_COLAB_CODE_DIR,
        COLAB_DRIVE_ROOT / "Daikin-FSC",
        Path.cwd(),
        Path.cwd() / "Daikin-FSC",
    ]

    seen: set[str] = set()
    for candidate in candidates:
        candidate_str = str(candidate)
        if candidate_str in seen:
            continue
        seen.add(candidate_str)

        if (candidate / "config.py").exists():
            return candidate

    raise ModuleNotFoundError(
        "Could not locate Daikin-FSC. Expected config.py in /content/Daikin-FSC "
        "or on Google Drive under RMT_FSC/Daikin-FSC."
    )


def mount_drive_if_needed() -> None:
    try:
        import google.colab  # noqa: F401
    except ImportError:
        return

    from google.colab import drive

    if not COLAB_DRIVE_ROOT.parent.exists():
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
CODE_DIR = find_code_dir()
sys.path.insert(0, str(CODE_DIR))

from config import WORKSPACE_ROOT, ensure_workspace_dirs

try:
    from config import bootstrap_sys_path
except ImportError:
    from config import setup_paths as bootstrap_sys_path

bootstrap_sys_path(CODE_DIR)
install_dependencies()
ensure_workspace_dirs()

from pipeline import run_pipeline

print(f"Workspace root: {WORKSPACE_ROOT}")
print(f"Code folder:    {CODE_DIR}")
print()

run_pipeline()
