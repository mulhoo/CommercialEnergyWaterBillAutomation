"""
North Marin Water District PDF extractor - Fixed date extraction and large number handling
"""

import os
import re
from typing import Optional

try:
    import pdfplumber
except ImportError as e:
    raise SystemExit(f"Missing PDF dependency: {e}")

from extractors.base import BaseExtractor
from models.bill_data import BillData

class NMWDExtractor(BaseExtractor):
    """Extract data from North Marin Water District bills"""

    def extract_data(self, pdf_path: str) -> Optional[BillData]:
        """Extract data from North Marin Water District bill"""
        try:
            with pdfplumber.open(pdf_path) as pdf:
                text = pdf.pages[0].extract_text()

                if not self._is_nmwd_bill(text):
                    # Only try OCR if text extraction failed AND OCR is available
                    if not text and hasattr(self, '_ocr_extract'):
                        text = self._ocr_extract(pdf_path)
                        if not self._is_nmwd_bill(text):
                            return None
                    else:
                        return None

                if not text:
                    return None

                print(f"DEBUG NMWD: Extracting from text length {len(text)}")

                account_number = (
                    self._extract_pattern(text, r'ACCOUNT(?:/CUSTOMER)? NUMBER[:\s]*([A-Z0-9\-]{6,})') or
                    self._extract_pattern(text, r'Customer Number[:\s]*([A-Z0-9\-]{6,})')
                )

                bill_date = self._extract_pattern(text, r'(\d{2}/\d{2}/\d{4})')

                due_date = "Upon Receipt" if "Upon Receipt" in text else \
                          self._extract_pattern(text, r'DUE DATE[^$]*(\d{2}/\d{2}/\d{4})')

                total_due = self._extract_nmwd_total_due(text)

                service_address = self._extract_pattern(text, r'SERVICE ADDRESS.*?(\d+[^,\n]*)')

                # FIXED: Allow multiple comma groups for large numbers like 3,864,065
                current_usage = (
                    self._extract_number(text, r'CURRENT PERIOD:?\s*(\d{1,3}(?:,\d{3})*)') or
                    self._extract_number(text, r'(\d{1,3}(?:,\d{3})*)\s+GAL') or 0
                )

                # Updated date extraction for NMWD
                start_date, end_date = self._extract_nmwd_period_dates(text)
                service_period = f"{start_date} - {end_date}" if start_date and end_date else ""

                print(f"DEBUG NMWD: Extracted dates - start: {start_date}, end: {end_date}")
                print(f"DEBUG NMWD: Extracted usage: {current_usage:,} gallons")

                if not account_number or total_due is None:
                    return None

                return BillData(
                    account_number=account_number,
                    bill_date=bill_date or '',
                    due_date=due_date or "Upon Receipt",
                    total_due=total_due,
                    service_address=service_address or '',
                    bill_start_date=start_date,
                    bill_end_date=end_date,
                    current_usage_gallons=current_usage,
                    service_period=service_period,
                    district="North Marin",
                    original_filename=os.path.basename(pdf_path)
                )

        except Exception as e:
            self.logger.error(f"Failed to extract NMWD data from {pdf_path}: {e}")
            return None

    def _extract_nmwd_period_dates(self, text: str) -> tuple[str, str]:
        """
        Extract billing period dates specifically for NMWD bills.
        Look for the service period dates, not due dates or other dates.
        """
        print(f"DEBUG NMWD: Looking for period dates in text...")

        # Normalize different dash types
        normalized_text = text.replace("\u2012", "-").replace("\u2013", "-").replace("\u2014", "-").replace("\u2212", "-")

        # NMWD-specific patterns - look for service period or billing period
        patterns = [
            # Pattern 1: "SERVICE PERIOD: MM/DD/YYYY - MM/DD/YYYY"
            r'SERVICE\s+PERIOD[:\s]*(\d{1,2}/\d{1,2}/\d{2,4})\s*[-–]\s*(\d{1,2}/\d{1,2}/\d{2,4})',

            # Pattern 2: "BILLING PERIOD: MM/DD/YYYY - MM/DD/YYYY"
            r'BILLING\s+PERIOD[:\s]*(\d{1,2}/\d{1,2}/\d{2,4})\s*[-–]\s*(\d{1,2}/\d{1,2}/\d{2,4})',

            # Pattern 3: "FROM MM/DD/YYYY TO MM/DD/YYYY" (but only in service context)
            r'(?:SERVICE|BILLING|PERIOD).*?FROM\s+(\d{1,2}/\d{1,2}/\d{2,4})\s+TO\s+(\d{1,2}/\d{1,2}/\d{2,4})',

            # Pattern 4: Look for dates near "CURRENT PERIOD" text
            r'CURRENT\s+PERIOD.*?(\d{1,2}/\d{1,2}/\d{2,4})\s*[-–]\s*(\d{1,2}/\d{1,2}/\d{2,4})',

            # Pattern 5: Look in a table structure for service dates
            r'(?:Service|Billing).*?(\d{1,2}/\d{1,2}/\d{2,4})\s*[-–]\s*(\d{1,2}/\d{1,2}/\d{2,4})',
        ]

        for i, pattern in enumerate(patterns):
            match = re.search(pattern, normalized_text, re.IGNORECASE | re.DOTALL)
            if match:
                start_date = self._normalize_date(match.group(1))
                end_date = self._normalize_date(match.group(2))
                print(f"DEBUG NMWD: Found dates with pattern {i+1}: {start_date} - {end_date}")
                return start_date, end_date

        # Fallback: Look for any two dates that might be service period
        # but be more careful about which ones we pick
        lines = normalized_text.split('\n')
        for line in lines:
            # Skip lines that clearly contain due dates or bill dates
            if any(keyword in line.upper() for keyword in ['DUE', 'PAYMENT', 'BILL DATE', 'INVOICE']):
                continue

            # Look for two dates in lines that might contain service period info
            if any(keyword in line.upper() for keyword in ['PERIOD', 'SERVICE', 'USAGE', 'CURRENT']):
                date_matches = re.findall(r'(\d{1,2}/\d{1,2}/\d{2,4})', line)
                if len(date_matches) >= 2:
                    start_date = self._normalize_date(date_matches[0])
                    end_date = self._normalize_date(date_matches[1])
                    print(f"DEBUG NMWD: Found dates in service line: {start_date} - {end_date}")
                    return start_date, end_date

        print(f"DEBUG NMWD: No service period dates found")
        return "", ""

    def _normalize_date(self, date_str: str) -> str:
        """Normalize date to MM/DD/YYYY format"""
        try:
            from datetime import datetime
            # Handle 2-digit years
            if date_str.count('/') == 2:
                parts = date_str.split('/')
                if len(parts[2]) == 2:
                    parts[2] = "20" + parts[2]
                    date_str = "/".join(parts)

            # Parse and reformat
            dt = datetime.strptime(date_str, "%m/%d/%Y")
            return dt.strftime("%m/%d/%Y")
        except:
            return date_str

    def _is_nmwd_bill(self, text: str) -> bool:
        """Check if this is actually a North Marin bill"""
        if not text:
            return False

        text_upper = text.upper()

        strong_nmwd_indicators = [
            "NORTH MARIN WATER DISTRICT",
            "NORTH MARIN",
        ]

        strong_mmwd_indicators = [
            "MARIN MUNICIPAL",
            "220 NELLEN AVENUE",
            "CORTE MADERA",
            "MARINWATER.ORG"
        ]

        has_strong_nmwd = any(indicator in text_upper for indicator in strong_nmwd_indicators)
        has_strong_mmwd = any(indicator in text_upper for indicator in strong_mmwd_indicators)

        if has_strong_nmwd and has_strong_mmwd:
            return "NORTH MARIN" in text_upper

        return has_strong_nmwd and not has_strong_mmwd

    def _extract_nmwd_total_due(self, text: str) -> Optional[float]:
        """
        North Marin PDFs sometimes cause the generic regex to latch onto a stray '7'.
        Strategy:
          - Look for lines that contain BOTH 'TOTAL' and 'DUE'
          - On that line, take the LAST currency-looking value (right-aligned on bills)
          - Fall back to a strict label→amount pattern if needed
        """
        money_pat = re.compile(r'\$?\s*\(?(?:\d{1,3}(?:,\d{3})*|\d+)(?:\.\d{2})?\)?')
        best: Optional[float] = None

        for raw_line in text.splitlines():
            line = raw_line.strip()
            upper_line = line.upper()
            if "TOTAL" in upper_line and "DUE" in upper_line:
                amounts = list(money_pat.finditer(line))
                if amounts:
                    amount_str = amounts[-1].group(0)
                    amount_str = amount_str.replace('$', '').replace(',', '').strip()
                    is_negative = amount_str.startswith('(') and amount_str.endswith(')')
                    if is_negative:
                        amount_str = amount_str[1:-1]
                    try:
                        value = float(amount_str)
                        best = -value if is_negative else value
                        break
                    except ValueError:
                        pass

        if best is not None:
            return best

        strict_pattern = (
            r'(?:^|\n)\s*(?:TOTAL\s+(?:AMOUNT\s+)?DUE(?:\s+NOW)?)\s*[:\-]?\s*\$?\s*'
            r'(\(?(?:\d{1,3}(?:,\d{3})*|\d+)(?:\.\d{2})?\)?)'
        )
        strict_match = re.search(strict_pattern, text, flags=re.IGNORECASE | re.MULTILINE)
        if strict_match:
            amount_str = strict_match.group(1).replace(',', '').replace('$', '').strip()
            is_negative = amount_str.startswith('(') and amount_str.endswith(')')
            if is_negative:
                amount_str = amount_str[1:-1]
            try:
                value = float(amount_str)
                return -value if is_negative else value
            except ValueError:
                return None

        return None