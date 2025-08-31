"""
Excel template processing functionality
"""

import os
import re
from typing import List, Optional
from datetime import datetime
from pathlib import Path

try:
    from openpyxl import load_workbook
    from openpyxl.styles import numbers
except ImportError as e:
    raise SystemExit(f"Missing Excel dependency: {e}")

from models.bill_data import BillData
from config import TEMPLATES, DISTRICT_CONFIG, EXCEL_LAYOUT, REPORTS_DIRS

class ExcelProcessor:
    """Process Excel templates and populate with bill data"""

    @staticmethod
    def _norm_acct(value) -> str:
        """Normalize account strings/numbers to digits-only for comparison."""
        if value is None:
            return ""
        return re.sub(r"\D", "", str(value))

    @staticmethod
    def _is_blank(cell) -> bool:
        """Check if a cell is blank or contains only whitespace"""
        cell_value = cell.value
        return cell_value is None or (isinstance(cell_value, str) and cell_value.strip() == "")

    def generate_excel_report(self, bills: List[BillData], district: str) -> Optional[str]:
        """Fill existing template rows by matching Account Number (col H). Do NOT add/delete rows."""
        if not bills:
            return None

        self.last_unmatched = []

        try:
            template_path = TEMPLATES[district]
            if not os.path.exists(template_path):
                print(f"Template not found: {template_path}")
                return None

            workbook = load_workbook(template_path)
            worksheet = workbook.active
            config = DISTRICT_CONFIG[district]

            start_row = EXCEL_LAYOUT["start_row"]
            account_col = EXCEL_LAYOUT["account_col"]

            acct_to_rows: dict[str, list[int]] = {}
            for row_num in range(start_row, worksheet.max_row + 1):
                acct_norm = self._norm_acct(worksheet.cell(row=row_num, column=account_col).value)
                if acct_norm:
                    acct_to_rows.setdefault(acct_norm, []).append(row_num)

            for bill in bills:
                acct_norm = self._norm_acct(bill.account_number)
                rows = acct_to_rows.get(acct_norm, [])

                if not rows:
                    self.last_unmatched.append((bill.account_number, bill.original_filename))
                    continue

                target_row = None
                for row_num in rows:
                    if self._is_blank(worksheet.cell(row=row_num, column=9)):
                        target_row = row_num
                        break

                if target_row is None:
                    target_row = rows[0]

                self._populate_row(worksheet, target_row, bill, config)

            output_path = self._generate_output_path(district)
            workbook.save(output_path)

            if self.last_unmatched:
                print("WARNING: No matching account rows for these bills:")
                for acct, filename in self.last_unmatched:
                    print(f"  - Account {acct} ({filename})")

            return str(output_path)

        except Exception as e:
            print(f"Error generating Excel report: {e}")
            return None

    def _populate_row(self, worksheet, row_num: int, bill: BillData, config: dict):
        """Populate a single row with bill data"""
        date_cell = worksheet.cell(row=row_num, column=1)
        if self._is_blank(date_cell):
            try:
                invoice_date = datetime.strptime(bill.bill_date, "%m/%d/%Y")
                date_cell.value = invoice_date
                date_cell.number_format = numbers.FORMAT_DATE_YYYYMMDD2
            except Exception:
                date_cell.value = bill.bill_date

        address_cell = worksheet.cell(row=row_num, column=2)
        if self._is_blank(address_cell):
            address_cell.value = bill.service_address

        period_cell = worksheet.cell(row=row_num, column=3)
        if self._is_blank(period_cell):
            period_cell.value = bill.service_period

        commodity_cell = worksheet.cell(row=row_num, column=4)
        if self._is_blank(commodity_cell):
            commodity_cell.value = "Water"

        gl_cell = worksheet.cell(row=row_num, column=5)
        if self._is_blank(gl_cell):
            gl_cell.value = "105-000-60035-803-0000"

        vendor_cell = worksheet.cell(row=row_num, column=6)
        if self._is_blank(vendor_cell):
            vendor_cell.value = config["vendor_id"]

        supplier_cell = worksheet.cell(row=row_num, column=7)
        if self._is_blank(supplier_cell):
            supplier_cell.value = config["supplier_name"]

        account_cell = worksheet.cell(row=row_num, column=8)
        if self._is_blank(account_cell):
            account_cell.value = bill.account_number

        charges_cell = worksheet.cell(row=row_num, column=9)
        if self._is_blank(charges_cell):
            charges_cell.value = bill.total_due
            charges_cell.number_format = numbers.FORMAT_CURRENCY_USD_SIMPLE

        usage_cell = worksheet.cell(row=row_num, column=10)
        if self._is_blank(usage_cell):
            try:
                gallons = int(bill.current_usage_gallons)
            except Exception:
                gallons = bill.current_usage_gallons
            usage_cell.value = gallons
            usage_cell.number_format = '#,##0'

    def _generate_output_path(self, district: str) -> Path:
        """Generate output path for the Excel report"""
        timestamp = datetime.now().strftime("%Y%m%d_%H%M")
        district_short = "NMWD" if district == "North Marin" else "MMWD"
        output_filename = f"BioMarin_{district_short}_Report_{timestamp}.xlsx"

        reports_dir = REPORTS_DIRS[district]
        reports_dir.mkdir(parents=True, exist_ok=True)

        return reports_dir / output_filename