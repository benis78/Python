"""
GUI modul til ExcelCopyBOM
Håndterer brugergrænsefladen med fil/mappe vælgere og options
"""

import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from datetime import datetime
import os

class ExcelCopyBOMGUI:
    def __init__(self, process_callback):
        """
        Initialiserer GUI
        :param process_callback: Callback funktion der kaldes når Start knappen trykkes
        """
        self.process_callback = process_callback
        self.root = tk.Tk()
        self.root.title("ExcelCopyBOM")
        self.root.geometry("600x500")
        
        # Variable til at holde stier og options
        self.bom_file = tk.StringVar()
        self.old_bom_file = tk.StringVar()
        self.source_dir = tk.StringVar()
        self.include_equipment = tk.BooleanVar()
        self.find_rev_before = tk.BooleanVar()
        self.include_datasheet = tk.BooleanVar()
        
        # Opret hovedramme med padding
        self.main_frame = ttk.Frame(self.root, padding="10")
        self.main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # Fil vælgere
        self.create_file_selectors()
        
        # Options
        self.create_options()
        
        # Progress bar og status
        self.create_progress_section()
        
        # Start knap
        self.start_button = ttk.Button(
            self.main_frame,
            text="Start",
            command=self.start_process
        )
        self.start_button.grid(row=8, column=0, columnspan=3, pady=20)
        
    def create_file_selectors(self):
        """Opretter fil/mappe vælgere"""
        # BOM fil vælger
        ttk.Label(self.main_frame, text="BOM Excel fil:").grid(
            row=0, column=0, sticky=tk.W, pady=5
        )
        self.bom_entry = ttk.Entry(self.main_frame, width=50)
        self.bom_entry.grid(row=0, column=1, pady=5)
        ttk.Button(
            self.main_frame,
            text="Vælg fil",
            command=lambda: self.select_file(self.bom_entry, "Excel fil", ".xlsx")
        ).grid(row=0, column=2, padx=5, pady=5)
        
        # Gammel BOM fil vælger
        ttk.Label(self.main_frame, text="Tidligere BOM fil (valgfri):").grid(
            row=1, column=0, sticky=tk.W, pady=5
        )
        self.old_bom_entry = ttk.Entry(self.main_frame, width=50)
        self.old_bom_entry.grid(row=1, column=1, pady=5)
        ttk.Button(
            self.main_frame,
            text="Vælg fil",
            command=lambda: self.select_file(self.old_bom_entry, "Excel fil", ".xlsx")
        ).grid(row=1, column=2, padx=5, pady=5)
        
        # PDF/DWG mappe vælger
        ttk.Label(self.main_frame, text="PDF/DWG mappe:").grid(
            row=2, column=0, sticky=tk.W, pady=5
        )
        self.source_entry = ttk.Entry(self.main_frame, width=50)
        self.source_entry.grid(row=2, column=1, pady=5)
        ttk.Button(
            self.main_frame,
            text="Vælg mappe",
            command=self.select_directory
        ).grid(row=2, column=2, padx=5, pady=5)
        
    def create_options(self):
        """Opretter options sektion"""
        # Options frame
        options_frame = ttk.LabelFrame(self.main_frame, text="Options", padding="5")
        options_frame.grid(row=3, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=10)
        
        # Equipment checkbox
        self.equipment_check = ttk.Checkbutton(
            options_frame,
            text="Inkluder equipment (0000-7xx)",
            variable=self.include_equipment
        )
        self.equipment_check.grid(row=0, column=0, sticky=tk.W, pady=5)
        
        # REV dato frame
        rev_frame = ttk.Frame(options_frame)
        rev_frame.grid(row=1, column=0, sticky=(tk.W, tk.E), pady=5)
        
        self.rev_check = ttk.Checkbutton(
            rev_frame,
            text="Find REV før dato:",
            variable=self.find_rev_before,
            command=self.toggle_date_entry
        )
        self.rev_check.grid(row=0, column=0, sticky=tk.W)
        
        # Dato entry felter
        self.date_entries = []
        for i, (label, width) in enumerate([("År", 6), ("Måned", 4), ("Dag", 4)]):
            ttk.Label(rev_frame, text=label).grid(row=0, column=i*2+1, padx=5)
            entry = ttk.Entry(rev_frame, width=width, state="disabled")
            entry.grid(row=0, column=i*2+2, padx=2)
            self.date_entries.append(entry)
        
        # Datasheet checkbox
        self.datasheet_check = ttk.Checkbutton(
            options_frame,
            text="Inkluder datasheets",
            variable=self.include_datasheet
        )
        self.datasheet_check.grid(row=2, column=0, sticky=tk.W, pady=5)
        
    def create_progress_section(self):
        """Opretter progress bar og status label"""
        # Progress frame
        progress_frame = ttk.LabelFrame(
            self.main_frame,
            text="Fremskridt",
            padding="5"
        )
        progress_frame.grid(
            row=7, column=0, columnspan=3,
            sticky=(tk.W, tk.E),
            pady=10
        )
        
        # Progress bar
        self.progress_bar = ttk.Progressbar(
            progress_frame,
            orient="horizontal",
            length=400,
            mode="determinate"
        )
        self.progress_bar.grid(row=0, column=0, pady=5)
        
        # Status label
        self.status_label = ttk.Label(progress_frame, text="Klar")
        self.status_label.grid(row=1, column=0, pady=5)
        
    def select_file(self, entry_widget, file_type, file_ext):
        """
        Åbner fil vælger dialog
        :param entry_widget: Entry widget der skal opdateres
        :param file_type: Beskrivelse af filtype
        :param file_ext: Filendelse (f.eks. .xlsx)
        """
        filename = filedialog.askopenfilename(
            title=f"Vælg {file_type}",
            filetypes=[(file_type, f"*{file_ext}")]
        )
        if filename:
            entry_widget.delete(0, tk.END)
            entry_widget.insert(0, filename)
    
    def select_directory(self):
        """Åbner mappe vælger dialog"""
        directory = filedialog.askdirectory(title="Vælg PDF/DWG mappe")
        if directory:
            self.source_entry.delete(0, tk.END)
            self.source_entry.insert(0, directory)
    
    def toggle_date_entry(self):
        """Aktiverer/deaktiverer dato felterne"""
        state = "normal" if self.find_rev_before.get() else "disabled"
        for entry in self.date_entries:
            entry.config(state=state)
    
    def get_date_value(self) -> datetime:
        """
        Henter dato fra dato felterne
        :return: datetime objekt eller None hvis dato ikke er valid
        """
        if not self.find_rev_before.get():
            return None
            
        try:
            year = int(self.date_entries[0].get())
            month = int(self.date_entries[1].get())
            day = int(self.date_entries[2].get())
            return datetime(year, month, day)
        except:
            return None
    
    def update_progress(self, value: int, status: str = None):
        """
        Opdaterer progress bar og status
        :param value: Progress værdi (0-100)
        :param status: Status tekst (optional)
        """
        self.progress_bar["value"] = value
        if status:
            self.status_label["text"] = status
        self.root.update_idletasks()
    
    def start_process(self):
        """Starter behandlingen når der trykkes på Start"""
        # Valider input
        bom_file = self.bom_entry.get()
        if not bom_file:
            messagebox.showerror("Fejl", "Vælg venligst en BOM Excel fil")
            return
            
        if not os.path.isfile(bom_file):
            messagebox.showerror("Fejl", "BOM filen findes ikke")
            return
        
        # Hent options
        old_bom_file = self.old_bom_entry.get()
        if old_bom_file and not os.path.isfile(old_bom_file):
            messagebox.showerror("Fejl", "Den valgte tidligere BOM fil findes ikke")
            return
            
        source_dir = self.source_entry.get()
        if source_dir and not os.path.isdir(source_dir):
            messagebox.showerror("Fejl", "Den valgte PDF/DWG mappe findes ikke")
            return
            
        # Valider dato hvis valgt
        rev_date = None
        if self.find_rev_before.get():
            rev_date = self.get_date_value()
            if not rev_date:
                messagebox.showerror("Fejl", "Ugyldig dato")
                return
        
        # Start behandlingen
        try:
            self.start_button["state"] = "disabled"
            success = self.process_callback(
                bom_file,
                old_bom_file if old_bom_file else None,
                source_dir if source_dir else None,
                self.include_equipment.get(),
                rev_date,
                self.include_datasheet.get()
            )
            
            if not success:
                messagebox.showerror(
                    "Fejl",
                    "Der opstod en fejl under behandlingen.\n\n"
                    "Se log filen for flere detaljer."
                )
                
        finally:
            self.start_button["state"] = "normal"
    
    def run(self):
        """Starter GUI'en"""
        self.root.mainloop()

if __name__ == "__main__":
    # Test kode
    def test_callback(*args):
        print("Process called with:", args)
        return True
        
    gui = ExcelCopyBOMGUI(test_callback)
    gui.run() 