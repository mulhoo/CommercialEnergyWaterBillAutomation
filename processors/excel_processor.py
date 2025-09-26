"""
Excel template processing functionality - DEBUG VERSION
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
    def _is_account_match(bill_account: str, excel_account: str) -> bool:
        """Check if bill account matches excel account (handles partial matches)"""
        if not bill_account or not excel_account:
            return False

        # Normalize both accounts (remove non-digits)
        bill_norm = re.sub(r"\D", "", str(bill_account))
        excel_norm = re.sub(r"\D", "", str(excel_account))

        # Check exact match first
        if bill_norm == excel_norm:
            return True

        # Check if bill account is contained in excel account (for cases like 495805 in 495805-61362)
        if bill_norm in excel_norm:
            return True

        # Check if excel account is contained in bill account (reverse case)
        if excel_norm in bill_norm:
            return True

        return False

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
                return None

            workbook = load_workbook(template_path)
            worksheet = workbook.active
            config = DISTRICT_CONFIG[district]

            start_row = EXCEL_LAYOUT["start_row"]
            account_col = EXCEL_LAYOUT["account_col"]

            excel_accounts = {}

            for row_num in range(start_row, min(start_row + 50, worksheet.max_row + 1)):
                cell_value = worksheet.cell(row=row_num, column=account_col).value
                if cell_value:
                    excel_account = str(cell_value).strip()
                    excel_accounts[row_num] = excel_account

            if not excel_accounts:
                for r in range(max(1, start_row-2), start_row+5):
                    for c in range(max(1, account_col-2), account_col+3):
                        cell_val = worksheet.cell(row=r, column=c).value

            for bill in bills:
                target_row = None

                for row_num, excel_account in excel_accounts.items():
                    if self._is_account_match(bill.account_number, excel_account):
                       if self._is_blank(worksheet.cell(row=row_num, column=9)):
                            target_row = row_num
                            break


                if target_row is None:
                    for row_num, excel_account in excel_accounts.items():
                        if self._is_account_match(bill.account_number, excel_account):
                            target_row = row_num
                            break

                if target_row:
                    self._populate_row(worksheet, target_row, bill, config)
                else:
                    self.last_unmatched.append((bill.account_number, bill.original_filename))

            output_path = self._generate_output_path(district, bills)
            workbook.save(output_path)

            return str(output_path)

        except Exception as e:
            import traceback
            traceback.print_exc()
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

    def _generate_output_path(self, district: str, bills: List[BillData]) -> Path:
        """Generate output path for the Excel report, replacing existing reports from the same date"""
        reports_dir = REPORTS_DIRS[district]
        reports_dir.mkdir(parents=True, exist_ok=True)

        if district == "North Marin":
            output_filename = "BioMarin Pharmaceutical Inc. Account Allocation - North Marin Water.xlsx"
        else:  # Marin Municipal
            output_filename = "BioMarin Pharmaceutical Inc. Account Allocation - Marin Municipal Water District.xlsx"
        output_path = reports_dir / output_filename

        return output_path