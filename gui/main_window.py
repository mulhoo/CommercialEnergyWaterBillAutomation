"""
Main GUI window for Water Bill PDF Processor - Windows Optimized
"""

import os
import logging
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from pathlib import Path
from urllib.parse import urlparse, unquote

logger = logging.getLogger(__name__)

try:
    from tkinterdnd2 import DND_FILES
    DND_OK = True
except Exception:
    DND_OK = False
    DND_FILES = None

from extractors.nmwd import NMWDExtractor
from extractors.mmwd import MMWDExtractor
from processors.file_renamer import FileRenamer
from processors.excel_processor import ExcelProcessor
from config import BILLS_DIRS, REPORTS_ROOT, TEMPLATES, month_year_folder, ensure_directories

class WaterBillProcessorGUI:
    """Main GUI application for water bill processing"""

    def __init__(self, root):
        self.root = root
        self.root.title("Commercial Energy Water Bill PDF Processor")

        # Configure window
        sw, sh = self.root.winfo_screenwidth(), self.root.winfo_screenheight()
        w = max(1200, int(sw * 0.75))
        h = max(800, int(sh * 0.80))
        self.root.geometry(f"{w}x{h}")
        self.root.minsize(1200, 800)
        
        # Set window colors for better Windows appearance
        self.root.configure(bg="#f0f0f0")

        # Initialize processors
        self.nmwd_extractor = NMWDExtractor()
        self.mmwd_extractor = MMWDExtractor()
        self.renamer = FileRenamer()
        self.excel_processor = ExcelProcessor()

        # Initialize state
        self.selected_files = []
        self._processing = False
        self._dialog_open = False
        self._last_dir = None

        # Create directories when needed
        try:
            ensure_directories()
        except Exception as e:
            logger.warning(f"Could not create directories: {e}")

        self.setup_gui()

    def setup_gui(self):
        """Setup the GUI components with Windows styling"""
        # Configure styling for Windows
        style = ttk.Style()
        style.theme_use('vista')
        
        # Configure custom styles
        style.configure("Title.TLabel", font=("Arial", 18, "bold"), foreground="#2c3e50")
        style.configure("Status.TLabel", font=("Arial", 11, "bold"), background="#ffffff")
        
        main_frame = ttk.Frame(self.root, padding="15")  # Reduced from 20
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))

        # Title
        title_label = ttk.Label(
            main_frame, 
            text="Water Bill PDF Processor",
            style="Title.TLabel"
        )
        title_label.grid(row=0, column=0, columnspan=3, pady=(0, 15))  # Reduced from 25

        # District selection
        district_frame = ttk.LabelFrame(main_frame, text="Select District", padding="10")  # Reduced from 15
        district_frame.grid(row=1, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=(0, 10))  # Reduced from 15

        self.district_var = tk.StringVar(value="North Marin")
        ttk.Radiobutton(
            district_frame, text="North Marin Water District",
            variable=self.district_var, value="North Marin"
        ).grid(row=0, column=0, sticky=tk.W, padx=(0, 60))
        ttk.Radiobutton(
            district_frame, text="Marin Municipal Water District",
            variable=self.district_var, value="Marin Municipal"
        ).grid(row=0, column=1, sticky=tk.W)

        # File Processing frame
        file_frame = ttk.LabelFrame(main_frame, text="File Processing", padding="10")  # Reduced from 15
        file_frame.grid(row=2, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=(0, 10))  # Reduced from 15
        file_frame.columnconfigure(0, weight=1)

        # Drop Zone with improved Windows styling - made more compact
        self.drop_zone = tk.Label(
            file_frame,
            text="⬇  Drop PDF files here (or click to select)\nFor Outlook: Save attachments first, then drag from folder",
            bd=2, 
            relief="groove",
            anchor="center",
            font=("Arial", 11, "bold"),
            height=3,  # Reduced from 6 to 3
            cursor="hand2",
            fg="#555555",
            bg="#f8f9fa",
            highlightbackground="#2196F3",
            highlightcolor="#1976D2",
            highlightthickness=1
        )
        self.drop_zone.grid(row=0, column=0, sticky=(tk.W, tk.E), pady=(10, 10))  # Reduced from (15, 15)
        self.drop_zone.bind("<Button-1>", lambda e: self.select_files())

        # Hover effects for drop zone
        def on_enter(e):
            self.drop_zone.config(bg="#e3f2fd", highlightthickness=2)
        def on_leave(e):
            self.drop_zone.config(bg="#f8f9fa", highlightthickness=1)
        
        self.drop_zone.bind("<Enter>", on_enter)
        self.drop_zone.bind("<Leave>", on_leave)

        # Setup drag and drop
        if DND_OK:
            try:
                self.drop_zone.drop_target_register(DND_FILES)
                self.drop_zone.dnd_bind("<<Drop>>", self._on_drop)
            except Exception as e:
                logger.warning(f"DnD setup failed: {e}")

        # Status banner and Process button
        status_frame = ttk.Frame(main_frame)
        status_frame.grid(row=3, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=(10, 10))

        self.status_var = tk.StringVar(value="No files selected")
        
        # Status label with better styling
        self.status_label = tk.Label(
            status_frame,
            textvariable=self.status_var,
            font=("Arial", 11, "bold"),
            anchor=tk.W,
            bg="#ffffff",
            fg="#333333",
            relief="sunken",
            bd=1,
            padx=10,
            pady=8
        )
        self.status_label.grid(row=0, column=0, sticky=(tk.W, tk.E), padx=(0, 10))

        # Process button with better styling
        self.process_btn = tk.Button(
            status_frame, 
            text="Process Files",
            font=("Arial", 11, "bold"),
            bg="#4CAF50",
            fg="white",
            relief="raised",
            bd=2,
            padx=20,
            pady=8,
            cursor="hand2",
            command=self.process_files
        )
        self.process_btn.grid(row=0, column=1)

        status_frame.columnconfigure(0, weight=1)

        # Warnings frame
        self.warnings_frame = ttk.LabelFrame(main_frame, text="Processing Warnings", padding="10")
        self.warnings_listbox = tk.Listbox(
            self.warnings_frame, height=3,
            fg="#ff6600", selectmode=tk.SINGLE,
            font=("Arial", 9)
        )
        warnings_scroll = ttk.Scrollbar(self.warnings_frame, orient=tk.VERTICAL,
                                      command=self.warnings_listbox.yview)
        self.warnings_listbox.configure(yscrollcommand=warnings_scroll.set)

        self.warnings_listbox.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        warnings_scroll.grid(row=0, column=1, sticky=(tk.N, tk.S))

        self.warnings_frame.columnconfigure(0, weight=1)
        self.warnings_frame.rowconfigure(0, weight=1)

        # Selected files frame (initially hidden) - made more compact
        self.selected_frame = ttk.LabelFrame(main_frame, text="Selected Files", padding="8")
        self.files_listbox = tk.Listbox(
            self.selected_frame, height=4, activestyle="dotbox",  # Reduced from 6 to 4
            exportselection=False, selectmode=tk.EXTENDED,
            font=("Arial", 9)
        )
        
        if DND_OK and hasattr(self.files_listbox, "drop_target_register"):
            try:
                self.files_listbox.drop_target_register(DND_FILES)
                self.files_listbox.dnd_bind("<<Drop>>", self._on_drop)
            except:
                pass

        files_scroll = ttk.Scrollbar(self.selected_frame, orient=tk.VERTICAL, command=self.files_listbox.yview)
        self.files_listbox.configure(yscrollcommand=files_scroll.set)

        self.files_listbox.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        files_scroll.grid(row=0, column=1, sticky=(tk.N, tk.S))

        # Remove button with better styling
        self.remove_btn = tk.Button(
            self.selected_frame, 
            text="✕",
            font=("Arial", 10, "bold"),
            bg="#f44336",
            fg="white",
            width=3,
            relief="raised",
            bd=1,
            cursor="hand2",
            command=self.remove_selected_files
        )
        self.remove_btn.grid(row=0, column=2, padx=(6, 0), sticky=tk.N)

        self.selected_frame.columnconfigure(0, weight=1)
        self.selected_frame.rowconfigure(0, weight=1)

        # Bind events
        self.files_listbox.bind("<Double-1>", self._on_file_double_click)
        self.files_listbox.bind("<Delete>", lambda e: self.remove_selected_files())
        self.files_listbox.bind("<BackSpace>", lambda e: self.remove_selected_files())

        # Results frame
        self.results_frame = ttk.LabelFrame(main_frame, text="Processing Results", padding="10")
        self.results_frame.grid(row=5, column=0, columnspan=3, sticky=(tk.W, tk.E, tk.N, tk.S), pady=(0, 10))

        # Clear button with proper Windows styling - simplified positioning
        self.clear_btn = tk.Button(
            self.results_frame, 
            text="Clear",
            bg="#d32f2f",
            fg="white",
            font=("Arial", 10, "bold"),
            relief="raised",
            borderwidth=2,
            padx=15,
            pady=5,
            cursor="hand2",
            command=self.clear_round
        )
        self.clear_btn.grid(row=0, column=1, sticky=tk.E, pady=(0, 10), padx=(0, 5))

        # Treeview with guaranteed scrollbar - simplified approach
        columns = (
            "Original File", "Renamed File", "Account", "Statement Date",
            "Bill Start", "Bill End", "Usage (gal)", "Amount", "Status",
        )
        
        # Set fixed height to ensure scrollbar appears when needed
        self.results_tree = ttk.Treeview(self.results_frame, columns=columns, show="headings", height=8)
        
        # Configure column headings and widths
        for col in columns:
            self.results_tree.heading(col, text=col)

        self.results_tree.column("Original File", width=180, minwidth=120)
        self.results_tree.column("Renamed File", width=180, minwidth=120)
        self.results_tree.column("Account", width=100, minwidth=80)
        self.results_tree.column("Statement Date", width=110, minwidth=90)
        self.results_tree.column("Bill Start", width=90, minwidth=70)
        self.results_tree.column("Bill End", width=90, minwidth=70)
        self.results_tree.column("Usage (gal)", width=100, minwidth=80)
        self.results_tree.column("Amount", width=90, minwidth=70)
        self.results_tree.column("Status", width=80, minwidth=60)

        # Configure treeview styling
        style.configure("Treeview", 
                       font=("Arial", 9),
                       rowheight=25)
        style.configure("Treeview.Heading", 
                       font=("Arial", 10, "bold"),
                       background="#e0e0e0")

        # Create and configure scrollbar
        tree_scrollbar = ttk.Scrollbar(self.results_frame, orient=tk.VERTICAL, command=self.results_tree.yview)
        self.results_tree.configure(yscrollcommand=tree_scrollbar.set)

        # Grid the treeview and scrollbar
        self.results_tree.grid(row=1, column=0, columnspan=2, sticky=(tk.W, tk.E, tk.N, tk.S), padx=(0, 5))
        tree_scrollbar.grid(row=1, column=2, sticky=(tk.N, tk.S))

        # Configure layout weights properly
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)

        main_frame.columnconfigure(0, weight=1)
        main_frame.rowconfigure(5, weight=1)

        self.results_frame.columnconfigure(0, weight=1)
        self.results_frame.columnconfigure(1, weight=0)
        self.results_frame.columnconfigure(2, weight=0)
        self.results_frame.rowconfigure(1, weight=1)

    def select_files(self):
        """Append newly chosen PDFs to the current selection"""
        if getattr(self, "_processing", False) or getattr(self, "_dialog_open", False):
            return

        self._dialog_open = True
        try:
            self.root.lift()
            self.root.attributes('-topmost', True)
            self.root.update_idletasks()
            self.root.after(50, lambda: self.root.attributes('-topmost', False))

            start_dir = (
                self._last_dir
                if self._last_dir and os.path.isdir(self._last_dir)
                else str(Path.home() / "Downloads")
            )

            chosen = filedialog.askopenfilenames(
                parent=self.root,
                title="Select Water Bill PDFs",
                initialdir=start_dir,
                filetypes=[("PDF files", "*.pdf")]
            )

            if chosen:
                # Add all chosen files without deduplication
                added = [os.path.abspath(p) for p in chosen]
                if added:
                    self.selected_files.extend(added)
                    self._last_dir = os.path.dirname(added[0])

                if hasattr(self, "selected_frame") and not self.selected_frame.winfo_ismapped():
                    self.selected_frame.grid(
                        row=4, column=0, columnspan=3,
                        sticky=(tk.W, tk.E, tk.N, tk.S), pady=(0, 10)
                    )

                if hasattr(self, "files_listbox"):
                    self.files_listbox.delete(0, tk.END)
                    for p in self.selected_files:
                        self.files_listbox.insert(tk.END, os.path.basename(p))

            self._update_selected_status()
        finally:
            self._dialog_open = False

    def _on_drop(self, event):
        """Accept dropped PDFs and optional folders - enhanced for Outlook compatibility"""
        paths = []
        try:
            # Try to split the data - handles regular file drops
            paths = list(self.root.tk.splitlist(event.data))
        except Exception:
            # Fallback for single items or special formats
            paths = [event.data]

        to_add = []
        outlook_temp_files = []
        
        for p in paths:
            # Handle file:// URLs
            if p.startswith("file://"):
                p = unquote(urlparse(p).path)
                # Fix Windows path issues from file:// URLs
                if p.startswith('/') and ':' in p:
                    p = p[1:]  # Remove leading slash for Windows paths like /C:/...

            p = os.path.abspath(p)

            # Check if it's a directory
            if os.path.isdir(p):
                for name in os.listdir(p):
                    if name.lower().endswith(".pdf"):
                        to_add.append(os.path.abspath(os.path.join(p, name)))
            elif p.lower().endswith(".pdf"):
                # Check if this might be an Outlook temp file
                if "outlook" in p.lower() or "tmp" in p.lower() or "temp" in p.lower():
                    # Copy Outlook temp files to a permanent location
                    try:
                        import shutil
                        import tempfile
                        
                        # Create a temp directory for Outlook files
                        temp_dir = tempfile.mkdtemp(prefix="outlook_pdfs_")
                        original_name = os.path.basename(p)
                        if not original_name or original_name.startswith('~'):
                            # Generate a name if Outlook gives us a temp name
                            original_name = f"outlook_attachment_{len(outlook_temp_files)+1}.pdf"
                        
                        permanent_path = os.path.join(temp_dir, original_name)
                        shutil.copy2(p, permanent_path)
                        to_add.append(permanent_path)
                        outlook_temp_files.append(permanent_path)
                        logger.info(f"Copied Outlook attachment: {p} -> {permanent_path}")
                    except Exception as e:
                        logger.warning(f"Failed to copy Outlook attachment: {e}")
                        # Try to use the original path anyway
                        to_add.append(p)
                else:
                    to_add.append(p)

        if not to_add:
            if paths:
                # Show a helpful message if drag and drop failed
                messagebox.showinfo("Drag and Drop", 
                    "No PDF files detected. If dragging from Outlook:\n\n"
                    "1. Save attachments to a folder first\n"
                    "2. Then drag from that folder\n\n"
                    "Or use the 'click to select' option instead.")
            return

        self.selected_files.extend(to_add)
        if to_add:
            self._last_dir = os.path.dirname(to_add[0])

        # Store outlook temp files for cleanup later
        if outlook_temp_files:
            if not hasattr(self, '_outlook_temp_files'):
                self._outlook_temp_files = []
            self._outlook_temp_files.extend(outlook_temp_files)

        if hasattr(self, "selected_frame") and not self.selected_frame.winfo_ismapped():
            self.selected_frame.grid(row=4, column=0, columnspan=3,
                                   sticky=(tk.W, tk.E, tk.N, tk.S), pady=(0, 10))

        if hasattr(self, "files_listbox"):
            self.files_listbox.delete(0, tk.END)
            for p in self.selected_files:
                self.files_listbox.insert(tk.END, os.path.basename(p))

        self._update_selected_status()

    def set_buttons_enabled(self, enabled: bool):
        """Enable/disable buttons during processing with proper styling"""
        state = tk.NORMAL if enabled else tk.DISABLED
        
        self.process_btn.configure(state=state)
        self.clear_btn.configure(state=state)
        
        # Update colors when disabled
        if enabled:
            self.process_btn.config(bg="#4CAF50", fg="white")
            self.clear_btn.config(bg="#d32f2f", fg="white")
        else:
            self.process_btn.config(bg="#cccccc", fg="#666666")
            self.clear_btn.config(bg="#cccccc", fg="#666666")

    def clear_round(self):
        """Clear all results and selected files"""
        self.results_tree.delete(*self.results_tree.get_children())
        self.selected_files = []
        if hasattr(self, 'files_listbox'):
            self.files_listbox.delete(0, tk.END)
        self.status_var.set("Ready to process files")
        if hasattr(self, "selected_frame") and self.selected_frame.winfo_ismapped():
            self.selected_frame.grid_remove()

    def _update_selected_status(self):
        """Update status bar with selected file count"""
        self.status_var.set(f"{len(self.selected_files)} file(s) selected for {self.district_var.get()}")

    def remove_selected_files(self):
        """Remove highlighted entries using listbox indices"""
        if not hasattr(self, "files_listbox"):
            return

        selections = list(self.files_listbox.curselection())
        if not selections:
            return

        selections.sort(reverse=True)
        for idx in selections:
            self.files_listbox.delete(idx)
            if 0 <= idx < len(self.selected_files):
                self.selected_files.pop(idx)

        if not self.selected_files and hasattr(self, "selected_frame") and self.selected_frame.winfo_ismapped():
            self.selected_frame.grid_remove()

        self._update_selected_status()

    def _on_file_double_click(self, event):
        """Show full path of double-clicked file in status bar"""
        if not hasattr(self, "files_listbox"):
            return

        selections = self.files_listbox.curselection()
        if not selections:
            return

        basename = self.files_listbox.get(selections[0])
        for p in self.selected_files:
            if os.path.basename(p) == basename:
                self.status_var.set(p)
                break

    def process_files(self):
        """Process the selected PDF files"""
        if not self.selected_files:
            messagebox.showwarning("No Files", "Please select PDF files first.")
            return

        if self._processing:
            return

        self._processing = True
        self.set_buttons_enabled(False)

        try:
            selected_district = self.district_var.get()

            self.warnings_listbox.delete(0, tk.END)
            warnings = []

            self.results_tree.delete(*self.results_tree.get_children())
            successful_bills = []

            for file_path in self.selected_files:
                self.status_var.set(f"Processing {os.path.basename(file_path)}...")
                self.root.update()

                logger.info(f"\n=== Processing File: {os.path.basename(file_path)} ===")

                try:
                    import pdfplumber
                    with pdfplumber.open(file_path) as pdf:
                        text = pdf.pages[0].extract_text()
                        logger.debug(f"Raw text length: {len(text) if text else 0}")
                        if text:
                            logger.debug(f"First 200 chars: {repr(text[:200])}")
                except Exception as e:
                    logger.error(f"Debug extraction failed: {e}")

                bill_data = None
                actual_district = None

                nmwd_data = self.nmwd_extractor.extract_data(file_path)
                if nmwd_data:
                    if hasattr(nmwd_data, 'district') and nmwd_data.district == "North Marin":
                        bill_data = nmwd_data
                        actual_district = "North Marin"

                if not bill_data:
                    mmwd_data = self.mmwd_extractor.extract_data(file_path)
                    if mmwd_data:
                        if hasattr(mmwd_data, 'district') and mmwd_data.district == "Marin Municipal":
                            bill_data = mmwd_data
                            actual_district = "Marin Municipal"

                if bill_data and actual_district != selected_district:
                    warning_msg = f"{os.path.basename(file_path)}: Bill is from {actual_district}, skipping (expected {selected_district})"
                    warnings.append(warning_msg)
                    self.warnings_listbox.insert(tk.END, warning_msg)

                    self.results_tree.insert("", "end", values=(
                        os.path.basename(file_path),
                        "—",
                        bill_data.account_number if bill_data else "—",
                        bill_data.bill_date if bill_data else "—",
                        "—", "—", "—", "—",
                        "Skipped - Wrong District"
                    ))
                    continue

                if bill_data:
                    try:
                        new_filename = self.renamer.generate_filename(bill_data)

                        month_folder = month_year_folder(bill_data.bill_date)
                        district_bills_dir = BILLS_DIRS[selected_district] / month_folder
                        district_bills_dir.mkdir(parents=True, exist_ok=True)

                        new_path = self.renamer.rename_file(file_path, bill_data)

                        self.results_tree.insert(
                            "", "end",
                            values=(
                                bill_data.original_filename,
                                new_filename,
                                bill_data.account_number,
                                bill_data.bill_date,
                                bill_data.bill_start_date,
                                bill_data.bill_end_date,
                                f"{bill_data.current_usage_gallons:,}",
                                f"${bill_data.total_due:,.2f}",
                                "Success",
                            ),
                        )

                        successful_bills.append(bill_data)
                    except Exception as e:
                        logger.error(f"Error processing bill: {e}", exc_info=True)
                        self.results_tree.insert("", "end", values=(
                            os.path.basename(file_path),
                            "Error",
                            "—", "—", "—", "—", "—",
                            f"Rename failed: {str(e)[:30]}",
                            "Failed"
                        ))
                else:
                    self.results_tree.insert("", "end", values=(
                        os.path.basename(file_path),
                        "—", "—", "—", "—", "—", "—", "—",
                        "Unable to extract data"
                    ))

            if successful_bills:
                logger.info(f"\n=== Generating Excel Report for {len(successful_bills)} bills ===")
                excel_path = self.excel_processor.generate_excel_report(successful_bills, selected_district)

                if hasattr(self.excel_processor, 'last_unmatched'):
                    for acct, filename in self.excel_processor.last_unmatched:
                        warning_msg = f"{filename}: Account {acct} not found in Excel template"
                        warnings.append(warning_msg)
                        self.warnings_listbox.insert(tk.END, warning_msg)

                if excel_path:
                    month_folder = month_year_folder(successful_bills[0].bill_date)
                    self.status_var.set(
                        f"Processed {len(successful_bills)} files. Excel report: {os.path.basename(excel_path)}"
                    )

                    success_message = "Processing complete!\n\n"
                    success_message += f"Excel report: {os.path.basename(excel_path)}\n"
                    success_message += f"Location: {excel_path}\n"
                    success_message += f"Renamed PDFs in: {BILLS_DIRS[selected_district] / month_folder}"
                    if warnings:
                        success_message += f"\n\n⚠ {len(warnings)} warning(s) - see warnings panel below"

                    messagebox.showinfo("Success", success_message)
                else:
                    # Excel generation failed - show detailed error
                    logger.error("Excel generation failed")
                    self.status_var.set(f"Processed {len(successful_bills)} PDFs, but Excel generation FAILED.")
                    
                    # Check for common issues
                    error_details = "Excel generation failed. Possible causes:\n\n"
                    
                    # Check if template exists
                    template_path = TEMPLATES[selected_district]
                    if not os.path.exists(template_path):
                        error_details += f"❌ Template file not found:\n   {template_path}\n\n"
                        logger.error(f"Template file not found: {template_path}")
                    else:
                        error_details += f"✓ Template file exists\n\n"
                    
                    # Check if output directory is accessible
                    if not REPORTS_ROOT.exists():
                        error_details += f"❌ Output directory not accessible:\n   {REPORTS_ROOT}\n\n"
                        logger.error(f"Output directory not accessible: {REPORTS_ROOT}")
                    else:
                        error_details += f"✓ Output directory accessible\n\n"
                    
                    # Check log file location
                    possible_desktops = [
                        Path.home() / "Desktop" / "WaterBillProcessor_Logs",
                        Path.home() / "OneDrive" / "Desktop" / "WaterBillProcessor_Logs",
                    ]
                    
                    log_dir = None
                    for loc in possible_desktops:
                        if loc and loc.exists():
                            log_dir = loc
                            break
                    
                    if not log_dir:
                        import tempfile
                        log_dir = Path(tempfile.gettempdir()) / "WaterBillProcessor_Logs"
                    
                    error_details += f"📋 Check the log file for details:\n   {log_dir}\n\n"
                    error_details += "Common fixes:\n"
                    error_details += "• Close any open Excel files\n"
                    error_details += "• Check network drive (X:) is connected\n"
                    error_details += "• Verify write permissions to the folder"
                    
                    messagebox.showerror("Excel Generation Failed", error_details)
            else:
                self.status_var.set("No files processed successfully.")

            if warnings:
                self.warnings_frame.grid(row=4, column=0, columnspan=3,
                                      sticky=(tk.W, tk.E), pady=(0, 10))
                if hasattr(self, "selected_frame") and self.selected_frame.winfo_ismapped():
                    self.selected_frame.grid_configure(row=5)
                self.results_frame.grid_configure(row=6)
            else:
                self.warnings_frame.grid_remove()
                self.results_frame.grid_configure(row=5)

        except Exception as e:
            logger.exception("Fatal error in process_files")
            messagebox.showerror("Error", f"An unexpected error occurred:\n\n{str(e)}")
        finally:
            self._processing = False
            self.set_buttons_enabled(True)

    def _is_wrong_district(self, bill_data, selected_district):
        """Check if extracted bill is from wrong district"""
        return bill_data.district != selected_district