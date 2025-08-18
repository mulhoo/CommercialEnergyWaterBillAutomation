from __future__ import annotations
import sys

if sys.version_info < (3, 10):
    raise RuntimeError("This script requires Python 3.10 or higher")

# NOW your regular imports start
import os
import re
import csv
import logging
from datetime import datetime
from pathlib import Path
from typing import Dict, List, Optional
from dataclasses import dataclass

# Third-party imports
try:
    from selenium import webdriver
    from selenium.webdriver.common.by import By
    # ... rest of imports
except ImportError as e:
    print(f"Missing required package: {e}")
    print("Install with: pip install selenium pdfplumber pandas")

# Then your classes and functions
@dataclass
class BillData:
