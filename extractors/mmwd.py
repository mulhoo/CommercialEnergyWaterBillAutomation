"""
Marin Municipal Water District PDF extractor
"""

import os
import re
from typing import Optional

try:
    import pdfplumber
except ImportError as e:
    raise SystemExit(f"Missing PDF dependency: {e}")

from extractors.base import BaseExtractor
from models.bill_data import BillData, normalize_mmddyyyy

class MMWDExtractor(BaseExtractor):
    def extract_data(self, pdf_path: str) -> Optional[BillData]:
        try:
            with pdfplumber.open(pdf_path) as pdf:
                text = pdf.pages[0].extract_text()

                if not self._is_mmwd_bill(text):
                    text = self._ocr_extract(pdf_path)
                    if not self._is_mmwd_bill(text):
                        return None

                if not text:
                    return None

                account_number = self._extract_pattern(text, r'Customer Number:?\s*(\d+)')
                bill_date = self._extract_pattern(text, r'Billing Date:?\s*(\d{2}/\d{2}/\d{4})')
                due_date = (
                    self._extract_pattern(text, r'Current Charges Due By:?\s*(\d{2}/\d{2}/\d{4})')
                    or "Upon Receipt"
                )
                total_due = self._extract_currency(text, r'TOTAL DUE:?\s*\$?([\d,]+\.?\d*)')
                service_address = self._extract_pattern(text, r'Service Address:?\s+(.+?)(?=\n)')

                current_units = 0

                units_match = re.search(r'Water Use\s+Units\*\s+(\d+)', text, re.IGNORECASE)
                if units_match:
                    current_units = int(units_match.group(1))
                    print(f"DEBUG MMWD: Found units via pattern 1: {current_units}")
                else:
                    table_match = re.search(r'(\d+)\s+(\d+(?:\s*1/2)?\")\s+(\d+)\s+(\d+)\s+(\d+)', text)
                    if table_match:
                        current_units = int(table_match.group(5))
                        print(f"DEBUG MMWD: Found units via pattern 2: {current_units}")
                    else:
                        lines = text.split('\n')
                        for i, line in enumerate(lines):
                            if 'Water Use' in line and i + 2 < len(lines):
                                if 'Units*' in lines[i + 1]:
                                    for j in range(i + 2, min(i + 5, len(lines))):
                                        number_match = re.search(r'^\s*(\d+)\s*$', lines[j].strip())
                                        if number_match:
                                            current_units = int(number_match.group(1))
                                            print(f"DEBUG MMWD: Found units via pattern 3: {current_units}")
                                            break
                                    break

                current_usage_gallons = current_units * 748
                print(f"DEBUG MMWD: Final usage - units: {current_units}, gallons: {current_usage_gallons}")

                start_date, end_date = self._extract_mmwd_meter_read_dates(text)
                service_period = f"{start_date} - {end_date}" if start_date and end_date else ""

                if not account_number or total_due is None:
                    return None

                return BillData(
                    account_number=account_number,
                    bill_date=bill_date or '',
                    due_date=due_date,
                    total_due=total_due,
                    service_address=service_address or '',
                    current_usage_gallons=current_usage_gallons,
                    service_period=service_period,
                    bill_start_date=start_date,
                    bill_end_date=end_date,
                    district="Marin Municipal",
                    original_filename=os.path.basename(pdf_path)
                )

        except Exception as e:
            self.logger.error(f"Failed to extract MMWD data from {pdf_path}: {e}")
            return None

    def _is_mmwd_bill(self, text: str) -> bool:
        """Check if this is actually a Marin Municipal bill"""
        if not text:
            return False

        text_upper = text.upper()

        strong_mmwd_indicators = [
            "MARIN MUNICIPAL",
            "220 NELLEN AVENUE",
            "CORTE MADERA",
            "MARINWATER.ORG"
        ]

        strong_nmwd_indicators = [
            "NORTH MARIN WATER DISTRICT",
            "NORTH MARIN",
        ]

        has_strong_mmwd = any(indicator in text_upper for indicator in strong_mmwd_indicators)
        has_strong_nmwd = any(indicator in text_upper for indicator in strong_nmwd_indicators)

        if has_strong_mmwd and has_strong_nmwd:
            return "MARIN MUNICIPAL" in text_upper

        return has_strong_mmwd and not has_strong_nmwd

    def _extract_mmwd_meter_read_dates(self, text: str) -> tuple[str, str]:
        """
        Extract start/end from the 'Meter Read Date' line on Marin Municipal bills.
        Handles:
          - same line:   'Meter Read Date: 06/11/2025 - 08/11/2025'
          - wrapped:     'Meter Read Date:\n06/11/2025 - 08/11/2025'
          - 'to' instead of '-' and Unicode dashes.
        """
        normalized_text = text.replace("\u2012", "-").replace("\u2013", "-").replace("\u2014", "-").replace("\u2212", "-")

        pattern = re.compile(
            r'Meter\s*Read\s*Date\s*[:\-]?\s*'
            r'(?:\n|\r|\s)*'
            r'(\d{1,2}/\d{1,2}/\d{2,4})'
            r'\s*(?:to|-)\s*'
            r'(\d{1,2}/\d{1,2}/\d{2,4})',
            flags=re.IGNORECASE
        )

        match = pattern.search(normalized_text)
        if not match:
            for line in normalized_text.splitlines():
                if "METER" in line.upper() and "READ" in line.upper() and "DATE" in line.upper():
                    fallback_match = re.search(
                        r'(\d{1,2}/\d{1,2}/\d{2,4})\s*(?:to|-)\s*(\d{1,2}/\d{1,2}/\d{2,4})',
                        line,
                        re.IGNORECASE
                    )
                    if fallback_match:
                        match = fallback_match
                        break

        if match:
            start, end = match.group(1), match.group(2)
            return normalize_mmddyyyy(start), normalize_mmddyyyy(end)

        return "", ""