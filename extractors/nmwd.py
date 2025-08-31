"""
North Marin Water District PDF extractor
"""

import os
import re
from typing import Optional

try:
    import pdfplumber
except ImportError as e:
    raise SystemExit(f"Missing PDF dependency: {e}")

from extractors.base import BaseExtractor
from models.bill_data import BillData, extract_period_dates

class NMWDExtractor(BaseExtractor):
    """Extract data from North Marin Water District bills"""

    def extract_data(self, pdf_path: str) -> Optional[BillData]:
        """Extract data from North Marin Water District bill"""
        try:
            with pdfplumber.open(pdf_path) as pdf:
                text = pdf.pages[0].extract_text()

                if not self._is_nmwd_bill(text):
                    text = self._ocr_extract(pdf_path)
                    if not self._is_nmwd_bill(text):
                        return None

                if not text:
                    return None

                account_number = (
                    self._extract_pattern(text, r'ACCOUNT(?:/CUSTOMER)? NUMBER[:\s]*([A-Z0-9\-]{6,})') or
                    self._extract_pattern(text, r'Customer Number[:\s]*([A-Z0-9\-]{6,})')
                )

                bill_date = self._extract_pattern(text, r'(\d{2}/\d{2}/\d{4})')

                due_date = "Upon Receipt" if "Upon Receipt" in text else \
                          self._extract_pattern(text, r'DUE DATE[^$]*(\d{2}/\d{2}/\d{4})')

                total_due = self._extract_nmwd_total_due(text)

                service_address = self._extract_pattern(text, r'SERVICE ADDRESS.*?(\d+[^,\n]*)')

                current_usage = (
                    self._extract_number(text, r'CURRENT PERIOD:?\s*(\d+(?:,\d+)?)') or
                    self._extract_number(text, r'(\d+)\s+GAL') or 0
                )

                start_date, end_date = extract_period_dates(text)
                service_period = f"{start_date} - {end_date}" if start_date and end_date else ""

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