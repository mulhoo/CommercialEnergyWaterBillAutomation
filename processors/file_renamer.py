"""
File renaming functionality for water bill PDFs
"""

import shutil
from pathlib import Path
from datetime import datetime

from models.bill_data import BillData

class FileRenamer:
    """Handle PDF file renaming according to specifications"""

    def generate_filename(self, bill_data: BillData) -> str:
        """Generate new filename based on district and data"""
        try:
            date_obj = datetime.strptime(bill_data.bill_date, "%m/%d/%Y")
            date_str = date_obj.strftime("%y%m%d")
        except:
            date_str = "000000"

        if bill_data.district == "North Marin":
            filename = f"{date_str} Account #{bill_data.account_number}.pdf"
        elif bill_data.district == "Marin Municipal":
            filename = f"{date_str} MMWD {bill_data.account_number}.pdf"
        else:
            filename = f"{date_str} {bill_data.account_number}.pdf"

        filename = "".join(c for c in filename if c.isalnum() or c in " #.-_")
        return filename

    def rename_file(self, original_path: str, new_filename: str, output_dir: str) -> str:
        """Rename and move file to output directory"""
        output_path = Path(output_dir) / new_filename
        shutil.copy2(original_path, output_path)
        return str(output_path)