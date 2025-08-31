"""
Data models for Water Bill PDF Processor
"""

from dataclasses import dataclass
from datetime import datetime
import re

@dataclass
class BillData:
    """Data structure for bill information"""
    account_number: str
    bill_date: str
    due_date: str
    total_due: float
    service_address: str
    current_usage_gallons: int
    service_period: str
    district: str
    original_filename: str
    bill_start_date: str = ""
    bill_end_date: str = ""

def normalize_mmddyyyy(date_str: str) -> str:
    """Return date string as MM/DD/YYYY (pads zeros, handles 2-digit years)."""
    date_str = date_str.strip()

    for fmt in ("%m/%d/%Y", "%m/%d/%y", "%-m/%-d/%Y", "%-m/%-d/%y"):
        try:
            dt = datetime.strptime(date_str, fmt)
            return dt.strftime("%m/%d/%Y")
        except Exception:
            continue

    parts = date_str.split("/")
    if len(parts) == 3 and len(parts[2]) == 2:
        parts[2] = "20" + parts[2]

    try:
        dt = datetime.strptime("/".join(parts), "%m/%d/%Y")
        return dt.strftime("%m/%d/%Y")
    except Exception:
        return date_str

def extract_period_dates(text: str) -> tuple[str, str]:
    """
    Extract (start_date, end_date) in MM/DD/YYYY from various bill wordings:
      - 'MM/DD/YYYY - MM/DD/YYYY'
      - 'FROM MM/DD/YYYY TO MM/DD/YYYY'
      - 'Meter Read Date: MM/DD/YYYY - MM/DD/YYYY'
      - 'Service Period: MM/DD/YYYY to MM/DD/YYYY'
    """
    patterns = [
        r'(\d{1,2}/\d{1,2}/\d{2,4})\s*[-–]\s*(\d{1,2}/\d{1,2}/\d{2,4})',
        r'FROM\s+(\d{1,2}/\d{1,2}/\d{2,4})\s+TO\s+(\d{1,2}/\d{1,2}/\d{2,4})',
        r'(?:Meter\s+Read\s+Date|Service\s+Period)[:\s]*'
        r'(\d{1,2}/\d{1,2}/\d{2,4})\s*(?:to|[-–])\s*(\d{1,2}/\d{1,2}/\d{2,4})',
    ]

    for pattern in patterns:
        match = re.search(pattern, text, flags=re.IGNORECASE)
        if match:
            start, end = match.group(1), match.group(2)
            return normalize_mmddyyyy(start), normalize_mmddyyyy(end)

    return "", ""