"""
Configuration constants for Water Bill PDF Processor
"""

import os
import sys
from pathlib import Path
from calendar import month_name
from datetime import datetime

# Get the user's Desktop directory
def get_desktop_dir():
    """Get the user's Desktop directory - prioritize regular Desktop over OneDrive"""
    import os
    
    # Try regular Desktop first (this is what you want)
    regular_desktop = Path.home() / "Desktop"
    if regular_desktop.exists():
        print(f"Using regular Desktop: {regular_desktop}")
        return regular_desktop
    
    # Only fall back to OneDrive if regular Desktop doesn't exist
    onedrive_desktop = Path.home() / "OneDrive" / "Desktop"
    if onedrive_desktop.exists():
        print(f"Using OneDrive Desktop: {onedrive_desktop}")
        return onedrive_desktop
    
    # Other fallbacks
    userprofile_desktop = Path(os.environ.get('USERPROFILE', '')) / "Desktop" if os.environ.get('USERPROFILE') else None
    if userprofile_desktop and userprofile_desktop.exists():
        print(f"Using USERPROFILE Desktop: {userprofile_desktop}")
        return userprofile_desktop
    
    # Last resort - use home directory
    print(f"Desktop not found, using home: {Path.home()}")
    return Path.home()

def get_base_dir():
    """Get the directory where the executable/script is located"""
    if getattr(sys, 'frozen', False):
        # Running as PyInstaller bundle
        return Path(sys.executable).parent
    else:
        # Running from source
        return Path(__file__).parent

DESKTOP = get_desktop_dir()
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

# Template paths - handle both bundled and development environments
def get_template_path(filename):
    """Get the full path to a template file"""
    base_dir = get_base_dir()
    template_path = base_dir / filename
    
    if template_path.exists():
        print(f"Found template: {template_path}")
        return str(template_path)
    else:
        print(f"Template not found: {template_path}")
        return filename  # Fallback to relative path

TEMPLATES = {
    "North Marin": get_template_path("BioMarin Pharmaceutical Inc. Account Allocation - North Marin Water - Template.xlsx"),
    "Marin Municipal": get_template_path("BioMarin Pharmaceutical Inc. Account Allocation - Marin Municipal Water District - Template.xlsx"),
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