"""
Configuration constants for Water Bill PDF Processor
"""

from pathlib import Path
from calendar import month_name
from datetime import datetime

DESKTOP = Path.home() / "Desktop"
BASE_DIR = DESKTOP / "Reports & Bills"
REPORTS_ROOT = BASE_DIR / "Reports"
BILLS_ROOT = BASE_DIR / "Bills"

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