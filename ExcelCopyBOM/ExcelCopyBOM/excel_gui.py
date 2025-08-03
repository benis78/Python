"""
GUI modul til ExcelCopyBOM (TRIN 1)
Håndterer brugergrænsefladen og input validering
"""

import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import os
from datetime import datetime
import logging

class ExcelCopyBOMGUI:
    def __init__(self, process_callback):
        """Initialiserer GUI vinduet"""
        self.window = tk.Tk()
        self.window.title("ExcelCopyBOM")
        self.window.geometry("600x400")
        
        # Centrer vinduet på skærmen
        screen_width = self.window.winfo_screenwidth()
        screen_height = self.window.winfo_screenheight()
        x = (screen_width - 600) // 2
        y = (screen_height - 400) // 2
        self.window.geometry(f"600x400+{x}+{y}")
        
        # Gem callback
        self.process_callback = process_callback
        
        # Opret hovedramme med padding
        self.main_frame = ttk.Frame(self.window, padding="10")
        self.main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # Opret input felter
        self._create_file_inputs()
        self._create_checkboxes()
        self._create_date_picker()
        self._create_progress()
        self._create_buttons()
        
        # Konfigurer grid weights
        self.window.columnconfigure(0, weight=1)
        self.window.rowconfigure(0, weight=1)
        self.main_frame.columnconfigure(1, weight=1)
        
        # Variable til at holde styr på valgte filer
        self.bom_file = ""
        self.prev_bom_file = ""
        
        # Logger
        self.logger = logging.getLogger('ExcelCopyBOM.GUI')
        
    def _create_file_inputs(self):
        """Opretter fil input felter"""
        # BOM fil vælger
        ttk.Label(self.main_frame, text="Excel BOM List:").grid(row=0, column=0, sticky=tk.W, pady=5)
        self.bom_entry = ttk.Entry(self.main_frame)
        self.bom_entry.grid(row=0, column=1, sticky=(tk.W, tk.E), padx=5)
        ttk.Button(self.main_frame, text="Browse", command=self._browse_bom).grid(row=0, column=2)
        
        # Previous BOM fil vælger
        ttk.Label(self.main_frame, text="Previous Drawing Package:").grid(row=1, column=0, sticky=tk.W, pady=5)
        self.prev_bom_entry = ttk.Entry(self.main_frame)
        self.prev_bom_entry.grid(row=1, column=1, sticky=(tk.W, tk.E), padx=5)
        ttk.Button(self.main_frame, text="Browse", command=self._browse_prev_bom).grid(row=1, column=2)
        
    def _create_checkboxes(self):
        """Opretter checkboxes"""
        # Equipment checkbox
        self.include_equipment = tk.BooleanVar()
        self.equipment_check = ttk.Checkbutton(
            self.main_frame, 
            text="Include Equipment, Valve, Instrument",
            variable=self.include_equipment,
            command=self._toggle_datasheet
        )
        self.equipment_check.grid(row=2, column=0, columnspan=2, sticky=tk.W, pady=5)
        
        # Datasheet checkbox
        self.include_datasheet = tk.BooleanVar()
        self.datasheet_check = ttk.Checkbutton(
            self.main_frame,
            text="Include Data Sheet",
            variable=self.include_datasheet,
            state=tk.DISABLED
        )
        self.datasheet_check.grid(row=3, column=0, columnspan=2, sticky=tk.W, pady=5)
        
    def _create_date_picker(self):
        """Opretter dato vælger"""
        # Find REV før dato
        self.use_date = tk.BooleanVar()
        self.date_check = ttk.Checkbutton(
            self.main_frame,
            text="Find REV files before date:",
            variable=self.use_date,
            command=self._toggle_date
        )
        self.date_check.grid(row=4, column=0, sticky=tk.W, pady=5)
        
        # Dato input felter
        date_frame = ttk.Frame(self.main_frame)
        date_frame.grid(row=4, column=1, sticky=tk.W)
        
        # Dag, måned, år input
        self.day_var = tk.StringVar(value=datetime.now().strftime("%d"))
        self.month_var = tk.StringVar(value=datetime.now().strftime("%m"))
        self.year_var = tk.StringVar(value=datetime.now().strftime("%Y"))
        
        self.day_entry = ttk.Entry(date_frame, width=3, textvariable=self.day_var, state=tk.DISABLED)
        self.month_entry = ttk.Entry(date_frame, width=3, textvariable=self.month_var, state=tk.DISABLED)
        self.year_entry = ttk.Entry(date_frame, width=5, textvariable=self.year_var, state=tk.DISABLED)
        
        self.day_entry.pack(side=tk.LEFT, padx=2)
        ttk.Label(date_frame, text="/").pack(side=tk.LEFT)
        self.month_entry.pack(side=tk.LEFT, padx=2)
        ttk.Label(date_frame, text="/").pack(side=tk.LEFT)
        self.year_entry.pack(side=tk.LEFT, padx=2)
        
    def _create_progress(self):
        """Opretter progress bar og status label"""
        # Progress frame
        progress_frame = ttk.Frame(self.main_frame)
        progress_frame.grid(row=5, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=20)
        
        # Progress bar
        self.progress = ttk.Progressbar(progress_frame, length=400, mode='determinate')
        self.progress.pack(fill=tk.X, padx=5)
        
        # Status label
        self.status_label = ttk.Label(progress_frame, text="")
        self.status_label.pack(pady=5)
        
    def _create_buttons(self):
        """Opretter knapper"""
        button_frame = ttk.Frame(self.main_frame)
        button_frame.grid(row=6, column=0, columnspan=3, pady=10)
        
        ttk.Button(button_frame, text="Start", command=self._start_process).pack(side=tk.LEFT, padx=5)
        ttk.Button(button_frame, text="Cancel", command=self.window.quit).pack(side=tk.LEFT, padx=5)
        
    def _browse_bom(self):
        """Åbner fil dialog for BOM fil"""
        filename = filedialog.askopenfilename(
            title="Vælg Excel BOM fil",
            filetypes=[("Excel files", "*.xlsx *.xls")]
        )
        if filename:
            self.bom_file = filename
            self.bom_entry.delete(0, tk.END)
            self.bom_entry.insert(0, filename)
            
    def _browse_prev_bom(self):
        """Åbner fil dialog for previous BOM fil"""
        filename = filedialog.askopenfilename(
            title="Vælg tidligere BOM fil",
            filetypes=[("Excel files", "*.xlsx *.xls")]
        )
        if filename:
            self.prev_bom_file = filename
            self.prev_bom_entry.delete(0, tk.END)
            self.prev_bom_entry.insert(0, filename)
            
    def _toggle_datasheet(self):
        """Aktiverer/deaktiverer datasheet checkbox baseret på equipment valg"""
        if self.include_equipment.get():
            self.datasheet_check.config(state=tk.NORMAL)
        else:
            self.include_datasheet.set(False)
            self.datasheet_check.config(state=tk.DISABLED)
            
    def _toggle_date(self):
        """Aktiverer/deaktiverer dato input felter"""
        state = tk.NORMAL if self.use_date.get() else tk.DISABLED
        self.day_entry.config(state=state)
        self.month_entry.config(state=state)
        self.year_entry.config(state=state)
        
    def _validate_inputs(self):
        """Validerer bruger input"""
        if not self.bom_file:
            messagebox.showerror("Fejl", "Vælg venligst en Excel BOM fil")
            return False
            
        if not os.path.exists(self.bom_file):
            messagebox.showerror("Fejl", "Den valgte BOM fil findes ikke")
            return False
            
        if self.prev_bom_file and not os.path.exists(self.prev_bom_file):
            messagebox.showerror("Fejl", "Den valgte tidligere BOM fil findes ikke")
            return False
            
        if self.use_date.get():
            try:
                datetime(
                    int(self.year_var.get()),
                    int(self.month_var.get()),
                    int(self.day_var.get())
                )
            except ValueError:
                messagebox.showerror("Fejl", "Ugyldig dato")
                return False
                
        return True
        
    def _start_process(self):
        """Starter behandlingen hvis input er validt"""
        if not self._validate_inputs():
            return
            
        # Saml parametre
        params = {
            'bom_file': self.bom_file,
            'old_bom_file': self.prev_bom_file if self.prev_bom_file else None,
            'include_equipment': self.include_equipment.get(),
            'include_datasheet': self.include_datasheet.get(),
            'find_rev_before': None,
            'gui': self
        }
        
        # Tilføj dato hvis valgt
        if self.use_date.get():
            params['find_rev_before'] = datetime(
                int(self.year_var.get()),
                int(self.month_var.get()),
                int(self.day_var.get())
            )
            
        # Start behandling
        try:
            self.process_callback(**params)
        except Exception as e:
            self.logger.error(f"Fejl under behandling: {str(e)}")
            messagebox.showerror("Fejl", f"Der opstod en fejl under behandlingen:\n{str(e)}")
            
    def update_progress(self, value: int, status: str = ""):
        """Opdaterer progress bar og status"""
        self.progress['value'] = value
        self.status_label['text'] = status
        self.window.update_idletasks()
        
    def run(self):
        """Starter GUI'en"""
        self.window.mainloop()

if __name__ == "__main__":
    # Test kode
    def test_callback(*args):
        print("Process called with:", args)
        return True
        
    gui = ExcelCopyBOMGUI(test_callback)
    gui.run() 