from __future__ import annotations

import sys
from pathlib import Path

COLAB_DRIVE_ROOT = Path(
    "/content/drive/Shareddrives/FA Ops Europe: Rate Maintenance Team "
    "/Documents/AI Adoption RMT/RMT_Daikin/RMT_FSC"
)
CODE_FOLDER_NAME = "Daikin-FSC"
LOCAL_COLAB_CODE_DIR = Path("/content/Daikin-FSC")


def is_colab() -> bool:
    try:
        import google.colab  # noqa: F401

        return True
    except ImportError:
        return False


def is_notebook_environment() -> bool:
    """Return True when running inside Jupyter, Colab, or IPython."""
    if is_colab():
        return True

    try:
        from IPython import get_ipython

        return get_ipython() is not None
    except ImportError:
        return False


def is_jupyter_kernel_argv(argv: list[str] | None = None) -> bool:
    """Detect Jupyter/Colab kernel arguments injected into sys.argv."""
    args = argv if argv is not None else sys.argv[1:]
    return len(args) >= 2 and args[0] in {"-f", "-F"} and args[1].endswith(".json")


def find_code_dir() -> Path:
    """Locate the folder that contains the Daikin-FSC Python modules."""
    candidates: list[Path] = []

    try:
        candidates.append(Path(__file__).resolve().parent)
    except NameError:
        pass

    candidates.extend(
        [
            LOCAL_COLAB_CODE_DIR,
            COLAB_DRIVE_ROOT / CODE_FOLDER_NAME,
            Path.cwd(),
            Path.cwd() / CODE_FOLDER_NAME,
        ]
    )

    seen: set[str] = set()
    for candidate in candidates:
        candidate_str = str(candidate)
        if candidate_str in seen:
            continue
        seen.add(candidate_str)

        if (candidate / "config.py").exists():
            return candidate

    raise ModuleNotFoundError(
        "Could not locate the Daikin-FSC code folder. "
        "Expected config.py in /content/Daikin-FSC or on Google Drive under RMT_FSC/Daikin-FSC."
    )


def get_workspace_root() -> Path:
    if is_colab():
        if COLAB_DRIVE_ROOT.exists():
            return COLAB_DRIVE_ROOT

        code_dir = find_code_dir()
        if code_dir.name == CODE_FOLDER_NAME and (code_dir.parent / "input").exists():
            return code_dir.parent

        return code_dir

    return find_code_dir()


CODE_DIR = find_code_dir()
WORKSPACE_ROOT = get_workspace_root()

INPUT_DIR = WORKSPACE_ROOT / "input" / "fsc file"
PROCESSING_DIR = WORKSPACE_ROOT / "processing"
OUTPUT_DIR = WORKSPACE_ROOT / "output"
CALCULATION_BASIS_DIR = WORKSPACE_ROOT / "input" / "rules" / "rules"
ORIGINAL_FILE_DIR = WORKSPACE_ROOT / "input" / "rules" / "original file"


def bootstrap_sys_path(code_dir: Path | None = None) -> Path:
    """Add the code folder to sys.path before importing local modules."""
    folder = code_dir or find_code_dir()
    folder_str = str(folder)
    if folder_str not in sys.path:
        sys.path.insert(0, folder_str)
    return folder


def setup_paths(code_dir: Path | None = None) -> Path:
    """Backwards-compatible alias for bootstrap_sys_path."""
    return bootstrap_sys_path(code_dir)


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
