"""
File renaming functionality for water bill PDFs
"""
import shutil
from pathlib import Path
from datetime import datetime
from models.bill_data import BillData
from config import BILLS_DIRS, month_year_folder

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

    def get_output_directory(self, bill_data: BillData) -> Path:
        """Get the correct output directory based on district and bill date"""
        base_dir = BILLS_DIRS[bill_data.district]

        month_year = month_year_folder(bill_data.bill_date)
        output_dir = base_dir / month_year

        return output_dir

    def rename_file(self, original_path: str, bill_data: BillData) -> str:
        """Rename and move file to correct network location"""
        try:
            new_filename = self.generate_filename(bill_data)
            output_dir = self.get_output_directory(bill_data)

            try:
                output_dir.mkdir(parents=True, exist_ok=True)
                print(f"Created/verified directory: {output_dir}")
            except Exception as e:
                print(f"Error creating directory {output_dir}: {e}")
                return None

            output_path = output_dir / new_filename

            try:
                shutil.copy2(original_path, output_path)
                print(f"File copied to: {output_path}")
                return str(output_path)
            except PermissionError as e:
                print(f"Permission error copying file: {e}")
                return None
            except Exception as e:
                print(f"Error copying file: {e}")
                return None

        except Exception as e:
            print(f"Error in rename_file: {e}")
            return None

    def check_network_access(self, district: str) -> bool:
        """Check if the network drive is accessible for the given district"""
        try:
            base_dir = BILLS_DIRS[district]
            if base_dir.parent.exists():
                return True
            else:
                print(f"Warning: Network path not accessible: {base_dir}")
                return False
        except Exception as e:
            print(f"Error accessing network drive: {e}")
            return False