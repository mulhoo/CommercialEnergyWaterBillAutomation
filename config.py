"""
Configuration constants for Water Bill PDF Processor
"""

from pathlib import Path
from calendar import month_name
from datetime import datetime

# Get the directory where the executable is located
def get_base_dir():
    """Get the base directory for the application"""
    import sys
    import os

    if getattr(sys, 'frozen', False):
        # Running as PyInstaller bundle
        app_dir = Path(sys.executable).parent
    else:
        # Running from source
        app_dir = Path(__file__).parent

    return app_dir

BASE_DIR = get_base_dir() / "Bills"
REPORTS_ROOT = get_base_dir() / "Reports"
BILLS_ROOT = BASE_DIR

REPORTS_DIRS = {
    "North Marin": REPORTS_ROOT / "North Marin",
    "Marin Municipal": REPORTS_ROOT / "Marin Municipal",
}

BILLS_DIRS = {
    "North Marin": BILLS_ROOT / "North Marin",
    "Marin Municipal": BILLS_ROOT / "Marin Municipal",
}

TEMPLATES = {
    "North Marin": "BioMarin Pharmaceutical Inc. Account Allocation - North Marin Water - Template.xlsx",
    "Marin Municipal": "BioMarin Pharmaceutical Inc. Account Allocation - Marin Municipal Water District - Template.xlsx",
}

DISTRICT_CONFIG = {
    "North Marin": {
        "vendor_id": "300011",
        "supplier_name": "North Marin Water District"
    },
    "Marin Municipal": {
        "vendor_id": "309438",
        "supplier_name": "Marin Municipal Water District"
    },
}

EXCEL_LAYOUT = {
    "start_row": 9,
    "last_col": 10,
    "account_col": 8,  # H
}

def month_year_folder(bill_date_str: str) -> str:
    """
    Convert bill date string to month/year folder format.
    bill_date_str is expected as %m/%d/%Y (e.g., 09/15/2025).
    Falls back to current month/year if parsing fails.
    """
    try:
        dt = datetime.strptime(bill_date_str, "%m/%d/%Y")
    except Exception:
        dt = datetime.now()
    return f"{month_name[dt.month]} {dt.year}"

def ensure_directories():
    """Create directories if they don't exist - call this when needed, not on import"""
    try:
        BASE_DIR.mkdir(exist_ok=True)
        REPORTS_ROOT.mkdir(exist_ok=True)
        for dir_path in REPORTS_DIRS.values():
            dir_path.mkdir(parents=True, exist_ok=True)
        for dir_path in BILLS_DIRS.values():
            dir_path.mkdir(parents=True, exist_ok=True)
        return True
    except Exception as e:
        print(f"Warning: Could not create directories: {e}")
        return False