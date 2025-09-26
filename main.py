#!/usr/bin/env python3
"""
Water Bill PDF Processor - Main Entry Point
Processes batch PDFs, renames files, and generates Excel reports
"""
import os
import sys
import tkinter as tk
from pathlib import Path

try:
    from tkinterdnd2 import TkinterDnD, DND_FILES
    DND_OK = True
except Exception:
    DND_OK = False
    TkinterDnD = tk
    DND_FILES = None

from gui.main_window import WaterBillProcessorGUI

def setup_bundled_dependencies():
    """Setup paths for bundled Tesseract and Poppler"""
    if getattr(sys, 'frozen', False):
        bundle_dir = Path(sys._MEIPASS)

        tesseract_path = bundle_dir / 'tesseract' / 'tesseract.exe'
        if tesseract_path.exists():
            try:
                import pytesseract
                pytesseract.pytesseract.tesseract_cmd = str(tesseract_path)
                print(f"Using bundled Tesseract: {tesseract_path}")
            except ImportError:
                pass

        poppler_path = bundle_dir / 'poppler'
        if poppler_path.exists():
            current_path = os.environ.get('PATH', '')
            os.environ['PATH'] = str(poppler_path) + os.pathsep + current_path
            print(f"Using bundled Poppler: {poppler_path}")

    else:
        print("Running from source - using system dependencies")

def check_dependencies():
    """Check if dependencies are available (after setup)"""
    missing = []

    # Test Tesseract
    try:
        import pytesseract
        # Try to get version to test if it works
        pytesseract.get_tesseract_version()
        print("Tesseract: Available")
    except Exception as e:
        missing.append(f"Tesseract OCR: {str(e)}")

    # Test Poppler
    try:
        from pdf2image import convert_from_path
        # Try to import - this will fail if poppler isn't found
        print("Poppler: Available")
    except Exception as e:
        missing.append(f"Poppler: {str(e)}")

    return missing

def main():
    """Run the application"""
    setup_bundled_dependencies()

    missing_deps = check_dependencies()

    if missing_deps:
        print("Missing dependencies:")
        for dep in missing_deps:
            print(f"  - {dep}")
        print("Some features may not work correctly.")

    root = TkinterDnD.Tk() if DND_OK else tk.Tk()
    app = WaterBillProcessorGUI(root)

    if missing_deps and hasattr(app, 'warnings_listbox'):
        for dep in missing_deps:
            app.warnings_listbox.insert(tk.END, f"Missing: {dep}")
        if hasattr(app, 'warnings_frame'):
            app.warnings_frame.grid(row=4, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=(0, 10))

    root.mainloop()
    return 0

if __name__ == "__main__":
    sys.exit(main())