#!/usr/bin/env python3
"""
Water Bill PDF Processor - Main Entry Point
Processes batch PDFs, renames files, and generates Excel reports
"""
import os
import sys
import tkinter as tk
from pathlib import Path
import logging
from datetime import datetime

# Setup logging to file FIRST (before any other imports)
def setup_logging():
    """Setup logging to a file - handles OneDrive Desktop"""
    # Try multiple possible Desktop locations
    possible_desktops = [
        Path.home() / "Desktop",
        Path.home() / "OneDrive" / "Desktop",
        Path(os.environ.get('USERPROFILE', '')) / "Desktop" if os.environ.get('USERPROFILE') else None,
        Path.home()  # Fallback to home directory
    ]
    
    log_dir = None
    for desktop in possible_desktops:
        if desktop and desktop.exists():
            log_dir = desktop / "WaterBillProcessor_Logs"
            try:
                log_dir.mkdir(exist_ok=True)
                break
            except:
                continue
    
    # If no Desktop found, use temp directory
    if not log_dir or not log_dir.exists():
        import tempfile
        log_dir = Path(tempfile.gettempdir()) / "WaterBillProcessor_Logs"
        log_dir.mkdir(exist_ok=True)
    
    log_file = log_dir / f"debug_{datetime.now().strftime('%Y%m%d_%H%M%S')}.log"
    
    # Configure logging - FILE ONLY, don't redirect stdout
    logging.basicConfig(
        level=logging.DEBUG,
        format='%(asctime)s - %(name)s - %(levelname)s - %(message)s',
        handlers=[
            logging.FileHandler(log_file, encoding='utf-8')
        ]
    )
    
    # Log startup info
    logging.info("="*60)
    logging.info("Water Bill Processor Started")
    logging.info(f"Log file: {log_file}")
    logging.info(f"Python version: {sys.version}")
    logging.info(f"Current directory: {os.getcwd()}")
    logging.info("="*60)
    
    return log_file

# Setup logging immediately
LOG_FILE = setup_logging()

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
                logging.info(f"Using bundled Tesseract: {tesseract_path}")
            except ImportError:
                pass

        poppler_path = bundle_dir / 'poppler'
        if poppler_path.exists():
            current_path = os.environ.get('PATH', '')
            os.environ['PATH'] = str(poppler_path) + os.pathsep + current_path
            logging.info(f"Using bundled Poppler: {poppler_path}")

    else:
        logging.info("Running from source - using system dependencies")

def check_dependencies():
    """Check if dependencies are available (after setup)"""
    missing = []

    # Test Tesseract
    try:
        import pytesseract
        pytesseract.get_tesseract_version()
        logging.info("Tesseract: Available")
    except Exception as e:
        missing.append(f"Tesseract OCR: {str(e)}")

    # Test Poppler
    try:
        from pdf2image import convert_from_path
        logging.info("Poppler: Available")
    except Exception as e:
        missing.append(f"Poppler: {str(e)}")

    return missing

def show_log_location(root):
    """Show a message box with the log file location"""
    # Find where logs are actually saved
    possible_desktops = [
        Path.home() / "Desktop" / "WaterBillProcessor_Logs",
        Path.home() / "OneDrive" / "Desktop" / "WaterBillProcessor_Logs",
        Path(os.environ.get('USERPROFILE', '')) / "Desktop" / "WaterBillProcessor_Logs" if os.environ.get('USERPROFILE') else None,
    ]
    
    log_dir = None
    for loc in possible_desktops:
        if loc and loc.exists():
            log_dir = loc
            break
    
    if not log_dir:
        import tempfile
        log_dir = Path(tempfile.gettempdir()) / "WaterBillProcessor_Logs"
    
    msg = f"Debug logs are being saved to:\n\n{log_dir}\n\nIf you encounter any issues, please send the latest log file to support."
    
    from tkinter import messagebox
    messagebox.showinfo("Debug Logging Enabled", msg)

def main():
    """Run the application"""
    try:
        setup_bundled_dependencies()

        missing_deps = check_dependencies()

        if missing_deps:
            logging.warning("Missing dependencies:")
            for dep in missing_deps:
                logging.warning(f"  - {dep}")
            logging.warning("Some features may not work correctly.")

        root = TkinterDnD.Tk() if DND_OK else tk.Tk()
        app = WaterBillProcessorGUI(root)

        if missing_deps and hasattr(app, 'warnings_listbox'):
            for dep in missing_deps:
                app.warnings_listbox.insert(tk.END, f"Missing: {dep}")
            if hasattr(app, 'warnings_frame'):
                app.warnings_frame.grid(row=4, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=(0, 10))

        # Show log location on startup
        root.after(1000, lambda: show_log_location(root))

        logging.info("Application GUI initialized successfully")
        root.mainloop()
        
        logging.info("Application closed normally")
        return 0
        
    except Exception as e:
        logging.exception("FATAL ERROR in main()")
        return 1

if __name__ == "__main__":
    sys.exit(main())