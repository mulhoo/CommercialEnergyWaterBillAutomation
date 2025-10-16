"""
Configuration constants for Water Bill PDF Processor - Network Drive Version with Desktop Fallback
"""
import os
import sys
from pathlib import Path
from calendar import month_name
from datetime import datetime

# Helper function to get the correct base path for bundled files
def get_base_path():
    """Get base path for resources, handling PyInstaller bundled app"""
    if getattr(sys, 'frozen', False):
        # Running as compiled executable
        return Path(sys._MEIPASS)
    else:
        # Running as script
        return Path(__file__).parent

# Base path for bundled resources (templates)
RESOURCE_BASE = get_base_path()

# Network drive base paths
NETWORK_BASE = Path("X:/Sales/Customers/Current Clients/B/BioMarin Pharmaceuticals/BioMarin Billing agreement/Payment Schedules")

# Check if network drive is accessible
def check_network_access():
    """Check if the X: drive is accessible"""
    try:
        if NETWORK_BASE.exists():
            print(f"Network drive accessible: {NETWORK_BASE}")
            return True
        else:
            print(f"Warning: Network drive not accessible: {NETWORK_BASE}")
            return False
    except Exception as e:
        print(f"Error accessing network drive: {e}")
        return False

# Determine if we should use network or desktop fallback
USE_NETWORK = check_network_access()

if USE_NETWORK:
    # Use network drive
    BIOMARIN_BASE = NETWORK_BASE
    BILLS_ROOT = BIOMARIN_BASE / "Utility Bills"
    REPORTS_ROOT = BIOMARIN_BASE / "Pending Invoice"
else:
    # Fall back to Desktop
    print("Using Desktop fallback location")
    # Try multiple Desktop locations (handles OneDrive)
    possible_desktops = [
        Path.home() / "Desktop",
        Path.home() / "OneDrive" / "Desktop",
        Path(os.environ.get('USERPROFILE', '')) / "Desktop" if os.environ.get('USERPROFILE') else None,
    ]
    
    DESKTOP = None
    for desktop in possible_desktops:
        if desktop and desktop.exists():
            DESKTOP = desktop
            break
    
    if not DESKTOP:
        DESKTOP = Path.home()
    
    BIOMARIN_BASE = DESKTOP / "WaterBills"
    BILLS_ROOT = BIOMARIN_BASE / "Bills"
    REPORTS_ROOT = BIOMARIN_BASE / "Reports"
    print(f"Using fallback location: {BIOMARIN_BASE}")

BILLS_DIRS = {
    "North Marin": BILLS_ROOT / "North Marin Water District",
    "Marin Municipal": BILLS_ROOT / "Marin Water",
}

REPORTS_DIRS = {
    "North Marin": REPORTS_ROOT,
    "Marin Municipal": REPORTS_ROOT,
}

# Templates are bundled with the application
TEMPLATES = {
    "North Marin": RESOURCE_BASE / "BioMarin Pharmaceutical Inc. Account Allocation - North Marin Water - Template.xlsx",
    "Marin Municipal": RESOURCE_BASE / "BioMarin Pharmaceutical Inc. Account Allocation - Marin Municipal Water District - Template.xlsx",
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
        BIOMARIN_BASE.mkdir(parents=True, exist_ok=True)
        BILLS_ROOT.mkdir(parents=True, exist_ok=True)
        REPORTS_ROOT.mkdir(parents=True, exist_ok=True)
        
        for dir_path in BILLS_DIRS.values():
            dir_path.mkdir(parents=True, exist_ok=True)
            print(f"Created/verified bills directory: {dir_path}")

        for dir_path in REPORTS_DIRS.values():
            dir_path.mkdir(parents=True, exist_ok=True)
            print(f"Created/verified reports directory: {dir_path}")

        return True
    except PermissionError as e:
        print(f"Permission error: Cannot create directories: {e}")
        return False
    except Exception as e:
        print(f"Error: Could not create directories: {e}")
        return False

def get_fallback_dirs():
    """Get local fallback directories if network drive is unavailable"""
    desktop = Path.home() / "Desktop" / "BioMarin_Backup"

    fallback_bills = {
        "North Marin": desktop / "Bills" / "North Marin Water District",
        "Marin Municipal": desktop / "Bills" / "Marin Water",
    }

    fallback_reports = {
        "North Marin": desktop / "Reports",
        "Marin Municipal": desktop / "Reports",
    }

    return fallback_bills, fallback_reports