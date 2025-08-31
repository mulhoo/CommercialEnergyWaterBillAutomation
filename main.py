#!/usr/bin/env python3
"""
Water Bill PDF Processor - Main Entry Point
Processes batch PDFs, renames files, and generates Excel reports
"""
import tkinter as tk
from shutil import which

try:
    from tkinterdnd2 import TkinterDnD, DND_FILES
    DND_OK = True
except Exception:
    DND_OK = False
    TkinterDnD = tk
    DND_FILES = None

from gui.main_window import WaterBillProcessorGUI

def check_system_dependencies():
    """Check for required system dependencies"""
    missing = []
    if which("tesseract") is None:
        missing.append("Tesseract OCR (brew install tesseract / choco install tesseract)")
    if which("pdftoppm") is None:
        missing.append("Poppler (brew install poppler / choco install poppler)")
    if missing:
        raise RuntimeError("Missing system dependencies:\n- " + "\n- ".join(missing))

def main():
    """Run the application"""
    try:
        check_system_dependencies()
    except RuntimeError as e:
        print(f"Error: {e}")
        return 1

    root = TkinterDnD.Tk() if DND_OK else tk.Tk()
    app = WaterBillProcessorGUI(root)
    root.mainloop()
    return 0

if __name__ == "__main__":
    import sys
    sys.exit(main())