from __future__ import annotations

import sys
from pathlib import Path

COLAB_DRIVE_ROOT = Path(
    "/content/drive/Shareddrives/FA Ops Europe: Rate Maintenance Team "
    "/Documents/AI Adoption RMT/RMT_Daikin/RMT_FSC"
)
CODE_FOLDER_NAME = "Daikin-FSC"


def is_colab() -> bool:
    try:
        import google.colab  # noqa: F401

        return True
    except ImportError:
        return False


def get_code_dir() -> Path:
    if is_colab():
        return COLAB_DRIVE_ROOT / CODE_FOLDER_NAME

    return Path(__file__).resolve().parent


def get_workspace_root() -> Path:
    if is_colab():
        return COLAB_DRIVE_ROOT

    return Path(__file__).resolve().parent


CODE_DIR = get_code_dir()
WORKSPACE_ROOT = get_workspace_root()

INPUT_DIR = WORKSPACE_ROOT / "input" / "fsc file"
PROCESSING_DIR = WORKSPACE_ROOT / "processing"
OUTPUT_DIR = WORKSPACE_ROOT / "output"
CALCULATION_BASIS_DIR = WORKSPACE_ROOT / "input" / "rules" / "rules"
ORIGINAL_FILE_DIR = WORKSPACE_ROOT / "input" / "rules" / "original file"


def setup_paths(code_dir: Path | None = None) -> Path:
    """Add the code folder to sys.path for Colab and local runs."""
    folder = code_dir or CODE_DIR
    folder_str = str(folder)
    if folder_str not in sys.path:
        sys.path.insert(0, folder_str)
    return folder


def ensure_workspace_dirs() -> None:
    """Create standard workspace folders if they do not exist."""
    for folder in (
        INPUT_DIR,
        PROCESSING_DIR,
        OUTPUT_DIR,
        CALCULATION_BASIS_DIR,
        ORIGINAL_FILE_DIR,
    ):
        folder.mkdir(parents=True, exist_ok=True)
