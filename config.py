"""
Configuration constants for Water Bill PDF Processor
"""
import os
from pathlib import Path
from calendar import month_name
from datetime import datetime

# Network drive base paths
BIOMARIN_BASE = Path("X:/Sales/Customers/Current Clients/B/BioMarin Pharmaceuticals/BioMarin Billing agreement/Payment Schedules")
BILLS_ROOT = BIOMARIN_BASE / "Utility Bills"
REPORTS_ROOT = BIOMARIN_BASE / "Pending Invoice"

BILLS_DIRS = {
    "North Marin": BILLS_ROOT / "North Marin Water District",
    "Marin Municipal": BILLS_ROOT / "Marin Water",
}

REPORTS_DIRS = {
    "North Marin": REPORTS_ROOT,
    "Marin Municipal": REPORTS_ROOT,
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

def check_network_access():
    """Check if the X: drive is accessible"""
    try:
        if BIOMARIN_BASE.exists():
            print(f"Network drive accessible: {BIOMARIN_BASE}")
            return True
        else:
            print(f"Warning: Network drive not accessible: {BIOMARIN_BASE}")
            return False
    except Exception as e:
        print(f"Error accessing network drive: {e}")
        return False

def ensure_directories():
    """Create directories if they don't exist - call this when needed, not on import"""
    try:
        if not check_network_access():
            print("Cannot create directories - network drive not accessible")
            return False

        BILLS_ROOT.mkdir(parents=True, exist_ok=True)
        for dir_path in BILLS_DIRS.values():
            dir_path.mkdir(parents=True, exist_ok=True)
            print(f"Created/verified bills directory: {dir_path}")

        # Create the reports directory
        REPORTS_ROOT.mkdir(parents=True, exist_ok=True)
        print(f"Created/verified reports directory: {REPORTS_ROOT}")

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