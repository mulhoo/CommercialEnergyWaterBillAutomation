"""
Main GUI window for Water Bill PDF Processor
"""

import os
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from pathlib import Path
from urllib.parse import urlparse, unquote

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
from config import BASE_DIR, BILLS_DIRS, month_year_folder, ensure_directories

class WaterBillProcessorGUI:
    """Main GUI application for water bill processing"""

    def __init__(self, root):
            self.root = root
            self.root.title("Commercial Energy Water Bill PDF Processor")

            sw, sh = self.root.winfo_screenwidth(), self.root.winfo_screenheight()
            w = max(1000, int(sw * 0.65))
            h = max(720, int(sh * 0.70))
            self.root.geometry(f"{w}x{h}")
            self.root.minsize(1000, 720)

            self.nmwd_extractor = NMWDExtractor()
            self.mmwd_extractor = MMWDExtractor()
            self.renamer = FileRenamer()
            self.excel_processor = ExcelProcessor()

            self.selected_files = []
            self._processing = False
            self._dialog_open = False
            self._last_dir = None

            # Create directories when needed, not on import
            from config import ensure_directories
            try:
                ensure_directories()
            except Exception as e:
                print(f"Warning: Could not create directories: {e}")

            self.setup_gui()

    def setup_gui(self):
        """Setup the GUI components"""
        main_frame = ttk.Frame(self.root, padding="10")
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))

        # Title
        title_label = ttk.Label(main_frame, text="Water Bill PDF Processor",
                               font=("Arial", 16, "bold"))
        title_label.grid(row=0, column=0, columnspan=3, pady=(0, 20))

        # District selection
        district_frame = ttk.LabelFrame(main_frame, text="Select District", padding="10")
        district_frame.grid(row=1, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=(0, 10))

        self.district_var = tk.StringVar(value="North Marin")
        ttk.Radiobutton(
            district_frame, text="North Marin Water District",
            variable=self.district_var, value="North Marin"
        ).grid(row=0, column=0, sticky=tk.W, padx=(0, 50))
        ttk.Radiobutton(
            district_frame, text="Marin Municipal Water District",
            variable=self.district_var, value="Marin Municipal"
        ).grid(row=0, column=1, sticky=tk.W)

        # File Processing frame
        file_frame = ttk.LabelFrame(main_frame, text="File Processing", padding="10")
        file_frame.grid(row=2, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=(0, 10))
        file_frame.columnconfigure(0, weight=1)

        # Drop Zone
        self.drop_zone = tk.Label(
            file_frame,
            text="⬇️  Drop PDF files here\n(or click to select)",
            bd=1, relief="solid", anchor="center",
            font=("Arial", 13), height=6, cursor="hand2",
            fg="#666666"
        )
        self.drop_zone.grid(row=0, column=0, sticky=(tk.W, tk.E), pady=(15, 15))
        self.drop_zone.bind("<Button-1>", lambda e: self.select_files())

        if DND_OK:
            try:
                self.drop_zone.drop_target_register(DND_FILES)
                self.drop_zone.dnd_bind("<<Drop>>", self._on_drop)
            except Exception as e:
                print(f"DnD setup failed: {e}")

        if DND_OK and hasattr(self.drop_zone, "drop_target_register"):
            self.drop_zone.drop_target_register(DND_FILES)
            self.drop_zone.dnd_bind("<<Drop>>", self._on_drop)

        # Status banner and Process button
        status_frame = ttk.Frame(main_frame)
        status_frame.grid(row=3, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=(6, 6))

        self.status_var = tk.StringVar(value="No files selected")
        status_style = ttk.Style()
        status_style.configure("Status.TLabel", font=("Arial", 11, "bold"))
        status_bar = ttk.Label(status_frame, textvariable=self.status_var, style="Status.TLabel",
                              anchor=tk.W, padding=(8, 6, 8, 6))
        status_bar.grid(row=0, column=0, sticky=(tk.W, tk.E))

        self.process_btn = ttk.Button(status_frame, text="Process Files", command=self.process_files)
        self.process_btn.grid(row=0, column=1, padx=(10, 0))

        status_frame.columnconfigure(0, weight=1)

        # Warnings
        warnings_frame = ttk.LabelFrame(main_frame, text="Processing Warnings", padding="10")
        self.warnings_frame = warnings_frame

        # Warning listbox
        self.warnings_listbox = tk.Listbox(
            warnings_frame, height=3,
            fg="orange", selectmode=tk.SINGLE
        )
        warnings_scroll = ttk.Scrollbar(warnings_frame, orient=tk.VERTICAL,
                                      command=self.warnings_listbox.yview)
        self.warnings_listbox.configure(yscrollcommand=warnings_scroll.set)

        self.warnings_listbox.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        warnings_scroll.grid(row=0, column=1, sticky=(tk.N, tk.S))

        warnings_frame.columnconfigure(0, weight=1)
        warnings_frame.rowconfigure(0, weight=1)

        # Selected files (initially hidden)
        self.selected_frame = ttk.LabelFrame(main_frame, text="Selected Files", padding="10")
        self.files_listbox = tk.Listbox(
            self.selected_frame, height=6, activestyle="dotbox",
            exportselection=False, selectmode=tk.EXTENDED
        )
        if DND_OK and hasattr(self.files_listbox, "drop_target_register"):
            self.files_listbox.drop_target_register(DND_FILES)
            self.files_listbox.dnd_bind("<<Drop>>", self._on_drop)

        files_scroll = ttk.Scrollbar(self.selected_frame, orient=tk.VERTICAL, command=self.files_listbox.yview)
        self.files_listbox.configure(yscrollcommand=files_scroll.set)

        self.files_listbox.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        files_scroll.grid(row=0, column=1, sticky=(tk.N, tk.S))

        remove_btn = ttk.Button(self.selected_frame, text="✕", width=3, command=self.remove_selected_files)
        remove_btn.grid(row=0, column=2, padx=(6, 0), sticky=tk.N)

        self.selected_frame.columnconfigure(0, weight=1)
        self.selected_frame.rowconfigure(0, weight=1)

        self.files_listbox.bind("<Double-1>", self._on_file_double_click)
        self.files_listbox.bind("<Delete>", lambda e: self.remove_selected_files())
        self.files_listbox.bind("<BackSpace>", lambda e: self.remove_selected_files())

        # Results
        results_frame = ttk.LabelFrame(main_frame, text="Processing Results", padding="10")
        results_frame.grid(row=5, column=0, columnspan=3, sticky=(tk.W, tk.E, tk.N, tk.S), pady=(0, 10))

        # Clear button at top-right
        style = ttk.Style()
        style.configure("Clear.TButton", padding=6, relief="flat", borderwidth=0,
                        background="#e57373", foreground="white")
        self.clear_btn = ttk.Button(results_frame, text="Clear",
                                    style="Clear.TButton", command=self.clear_round)
        self.clear_btn.grid(row=0, column=1, sticky=tk.E, pady=(0, 10))

        # Treeview and scrollbar
        columns = (
            "Original File", "Renamed File", "Account", "Statement Date",
            "Bill Start", "Bill End", "Usage (gal)", "Amount", "Status",
        )
        self.results_tree = ttk.Treeview(results_frame, columns=columns, show="headings", height=15)
        for col in columns:
            self.results_tree.heading(col, text=col)

        self.results_tree.column("Original File", width=180)
        self.results_tree.column("Renamed File", width=180)
        self.results_tree.column("Account", width=110)
        self.results_tree.column("Statement Date", width=110)
        self.results_tree.column("Bill Start", width=100)
        self.results_tree.column("Bill End", width=100)
        self.results_tree.column("Usage (gal)", width=110)
        self.results_tree.column("Amount", width=100)
        self.results_tree.column("Status", width=90)

        scrollbar = ttk.Scrollbar(results_frame, orient=tk.VERTICAL, command=self.results_tree.yview)
        self.results_tree.configure(yscrollcommand=scrollbar.set)

        self.results_tree.grid(row=1, column=0, columnspan=2, sticky=(tk.W, tk.E, tk.N, tk.S))
        scrollbar.grid(row=1, column=2, sticky=(tk.N, tk.S))

        # Layout weights
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)

        main_frame.columnconfigure(0, weight=1)
        main_frame.rowconfigure(5, weight=1)

        results_frame.columnconfigure(0, weight=1, minsize=800)
        results_frame.columnconfigure(1, weight=0)
        results_frame.columnconfigure(2, weight=0)
        results_frame.rowconfigure(1, weight=1, minsize=280)

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
                existing = set(map(os.path.abspath, self.selected_files))
                added = [os.path.abspath(p) for p in chosen if os.path.abspath(p) not in existing]
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
        """Accept dropped PDFs and optional folders"""
        paths = []
        try:
            paths = list(self.root.tk.splitlist(event.data))
        except Exception:
            paths = [event.data]

        to_add = []
        for p in paths:
            if p.startswith("file://"):
                p = unquote(urlparse(p).path)

            p = os.path.abspath(p)

            if os.path.isdir(p):
                for name in os.listdir(p):
                    if name.lower().endswith(".pdf"):
                        to_add.append(os.path.abspath(os.path.join(p, name)))
            else:
                if p.lower().endswith(".pdf"):
                    to_add.append(p)

        existing = set(map(os.path.abspath, self.selected_files))
        added = [p for p in to_add if os.path.abspath(p) not in existing]
        if not added:
            return

        self.selected_files.extend(added)
        self._last_dir = os.path.dirname(added[0])

        if hasattr(self, "selected_frame") and not self.selected_frame.winfo_ismapped():
            self.selected_frame.grid(row=4, column=0, columnspan=3,
                                   sticky=(tk.W, tk.E, tk.N, tk.S), pady=(0, 10))

        if hasattr(self, "files_listbox"):
            self.files_listbox.delete(0, tk.END)
            for p in self.selected_files:
                self.files_listbox.insert(tk.END, os.path.basename(p))

        self._update_selected_status()

    def set_buttons_enabled(self, enabled: bool):
        """Enable/disable buttons during processing"""
        state = tk.NORMAL if enabled else tk.DISABLED
        self.process_btn.configure(state=state)
        self.clear_btn.configure(state=state)

    def clear_round(self):
        """Clear all results and selected files"""
        self.results_tree.delete(*self.results_tree.get_children())
        self.selected_files = []
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

                print(f"\n=== DEBUG - File: {os.path.basename(file_path)} ===")

                try:
                    import pdfplumber
                    with pdfplumber.open(file_path) as pdf:
                        text = pdf.pages[0].extract_text()
                        print(f"Raw text length: {len(text) if text else 0}")
                        if text:
                            print(f"First 200 chars: {repr(text[:200])}")
                            text_upper = text.upper()

                            nmwd_indicators = [
                                "NORTH MARIN WATER DISTRICT",
                                "NORTH MARIN",
                                "999 RUSH CREEK",
                                "NOVATO",
                                "NMWD.COM"
                            ]

                            found_nmwd = [ind for ind in nmwd_indicators if ind in text_upper]
                            print(f"NMWD indicators found: {found_nmwd}")

                            mmwd_indicators = [
                                "MARIN WATER",
                                "MARIN MUNICIPAL",
                                "220 NELLEN AVENUE",
                                "CORTE MADERA",
                                "MARINWATER.ORG"
                            ]

                            found_mmwd = [ind for ind in mmwd_indicators if ind in text_upper]
                            print(f"MMWD indicators found: {found_mmwd}")
                except Exception as e:
                    print(f"Debug extraction failed: {e}")

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

                        new_path = self.renamer.rename_file(file_path, new_filename, str(district_bills_dir))

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
                    success_message += f"Renamed PDFs in: {BASE_DIR / 'Bills' / selected_district / month_folder}"
                    if warnings:
                        success_message += f"\n\n⚠️ {len(warnings)} warning(s) - see warnings panel below"

                    messagebox.showinfo("Success", success_message)
                else:
                    self.status_var.set(f"Processed {len(successful_bills)} files. Excel generation failed.")
            else:
                self.status_var.set("No files processed successfully.")

            if warnings:
                self.warnings_frame.grid(row=4, column=0, columnspan=3,
                                      sticky=(tk.W, tk.E), pady=(0, 10))
                if hasattr(self, "selected_frame") and self.selected_frame.winfo_ismapped():
                    self.selected_frame.grid_configure(row=5)
                results_frame = self.results_tree.master
                results_frame.grid_configure(row=6)
            else:
                self.warnings_frame.grid_remove()
                results_frame = self.results_tree.master
                results_frame.grid_configure(row=5)

        finally:
            self._processing = False
            self.set_buttons_enabled(True)

    def _is_wrong_district(self, bill_data, selected_district):
        """Check if extracted bill is from wrong district"""
        return bill_data.district != selected_district