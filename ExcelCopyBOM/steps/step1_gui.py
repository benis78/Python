"""
TRIN 1: Brugergrænseflade
Håndterer brugergrænsefladen for ExcelCopyBOM programmet
"""

import logging
import os
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from datetime import datetime
from tkcalendar import DateEntry

class ExcelCopyBOMGUI:
    def __init__(self, process_callback):
        self.logger = logging.getLogger('ExcelCopyBOM.GUI')
        self.process_callback = process_callback
        
        # Opret hovedvindue
        self.root = tk.Tk()
        self.root.title("ExcelCopyBOM")
        self.root.geometry("600x400")
        
        # Tilføj padding og styling
        style = ttk.Style()
        style.configure('TFrame', padding=10)
        style.configure('TButton', padding=5)
        style.configure('TLabel', padding=5)
        
        # Opret hovedramme
        self.main_frame = ttk.Frame(self.root)
        self.main_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        # Opret GUI elementer
        self._create_file_selector()
        self._create_options()
        self._create_progress()
        self._create_start_button()
        
    def _create_file_selector(self):
        """Opretter fil-vælger sektion"""
        # BOM fil vælger
        bom_frame = ttk.Frame(self.main_frame)
        bom_frame.pack(fill=tk.X, pady=5)
        
        ttk.Label(bom_frame, text="BOM fil:").pack(side=tk.LEFT)
        self.bom_path = tk.StringVar()
        ttk.Entry(bom_frame, textvariable=self.bom_path, width=50).pack(side=tk.LEFT, padx=5)
        ttk.Button(bom_frame, text="Vælg...", command=self._browse_bom).pack(side=tk.LEFT)
        
        # Gammel BOM fil vælger (valgfri)
        old_bom_frame = ttk.Frame(self.main_frame)
        old_bom_frame.pack(fill=tk.X, pady=5)
        
        ttk.Label(old_bom_frame, text="Gammel BOM (valgfri):").pack(side=tk.LEFT)
        self.old_bom_path = tk.StringVar()
        ttk.Entry(old_bom_frame, textvariable=self.old_bom_path, width=50).pack(side=tk.LEFT, padx=5)
        ttk.Button(old_bom_frame, text="Vælg...", command=self._browse_old_bom).pack(side=tk.LEFT)
        
    def _create_options(self):
        """Opretter valgmuligheder"""
        options_frame = ttk.LabelFrame(self.main_frame, text="Indstillinger")
        options_frame.pack(fill=tk.X, pady=10)
        
        # Checkboxes
        self.include_equipment = tk.BooleanVar(value=True)
        ttk.Checkbutton(
            options_frame, 
            text="Inkluder udstyr",
            variable=self.include_equipment
        ).pack(anchor=tk.W, padx=5)
        
        self.include_datasheet = tk.BooleanVar(value=False)
        ttk.Checkbutton(
            options_frame,
            text="Inkluder datablad",
            variable=self.include_datasheet
        ).pack(anchor=tk.W, padx=5)
        
        # Dato vælger
        date_frame = ttk.Frame(options_frame)
        date_frame.pack(fill=tk.X, pady=5)
        
        self.use_date = tk.BooleanVar(value=False)
        ttk.Checkbutton(
            date_frame,
            text="Find revisioner før:",
            variable=self.use_date,
            command=self._toggle_date
        ).pack(side=tk.LEFT, padx=5)
        
        self.date_picker = DateEntry(
            date_frame,
            width=12,
            background='darkblue',
            foreground='white',
            borderwidth=2,
            state='disabled'
        )
        self.date_picker.pack(side=tk.LEFT, padx=5)
        
    def _create_progress(self):
        """Opretter fremskridtsindikator"""
        progress_frame = ttk.Frame(self.main_frame)
        progress_frame.pack(fill=tk.X, pady=10)
        
        self.progress = ttk.Progressbar(
            progress_frame,
            orient=tk.HORIZONTAL,
            length=300,
            mode='determinate'
        )
        self.progress.pack(fill=tk.X)
        
        self.status_label = ttk.Label(progress_frame, text="")
        self.status_label.pack(fill=tk.X)
        
    def _create_start_button(self):
        """Opretter start knap"""
        button_frame = ttk.Frame(self.main_frame)
        button_frame.pack(fill=tk.X, pady=10)
        
        self.start_button = ttk.Button(
            button_frame,
            text="Start behandling",
            command=self._start_processing
        )
        self.start_button.pack(side=tk.RIGHT)
        
    def _browse_bom(self):
        """Åbner fil-dialog for BOM fil"""
        filename = filedialog.askopenfilename(
            title="Vælg BOM fil",
            filetypes=[("Excel filer", "*.xlsx *.xls")],
            initialdir="C:\\Working Folder\\Designs\\5-Projects"
        )
        if filename:
            self.bom_path.set(filename)
            
    def _browse_old_bom(self):
        """Åbner fil-dialog for gammel BOM fil"""
        filename = filedialog.askopenfilename(
            title="Vælg gammel BOM fil",
            filetypes=[("Excel filer", "*.xlsx *.xls")],
            initialdir="C:\\Working Folder\\Designs\\5-Projects"
        )
        if filename:
            self.old_bom_path.set(filename)
            
    def _toggle_date(self):
        """Aktiverer/deaktiverer dato-vælger"""
        if self.use_date.get():
            self.date_picker.configure(state='normal')
        else:
            self.date_picker.configure(state='disabled')
            
    def _start_processing(self):
        """Starter behandling af BOM fil"""
        if not self.bom_path.get():
            messagebox.showerror("Fejl", "Vælg venligst en BOM fil")
            return
            
        try:
            self.start_button.configure(state='disabled')
            self.progress['value'] = 0
            self.status_label['text'] = "Starter behandling..."
            
            # Kald process callback med valgte indstillinger
            success = self.process_callback(
                self.bom_path.get(),
                self.old_bom_path.get() if self.old_bom_path.get() else None,
                self.include_equipment.get(),
                self.include_datasheet.get(),
                datetime.combine(self.date_picker.get_date(), datetime.min.time()) if self.use_date.get() else None,
                self  # Send GUI reference for progress updates
            )
            
            if success:
                messagebox.showinfo("Succes", "BOM behandling gennemført!")
            else:
                messagebox.showerror("Fejl", "Der opstod en fejl under behandlingen")
                
        except Exception as e:
            self.logger.error(f"Fejl under behandling: {str(e)}", exc_info=True)
            messagebox.showerror("Fejl", f"Uventet fejl: {str(e)}")
        finally:
            self.start_button.configure(state='normal')
            
    def update_progress(self, value: int, status: str = None):
        """Opdaterer fremskridtsindikator og status"""
        self.progress['value'] = value
        if status:
            self.status_label['text'] = status
        self.root.update_idletasks()
        
    def run(self):
        """Starter GUI"""
        self.root.mainloop() 