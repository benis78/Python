"""
GUI modul til ExcelCopyBOM (TRIN 1)
Håndterer alle GUI relaterede funktioner og brugerinteraktion
"""

import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import os
from datetime import datetime

class ExcelCopyBOMGUI:
    def __init__(self):
        self.root = tk.Tk()
        self.root.attributes('-topmost', True)
        self.root.title("ExcelCopyBOM")
        
        # Variables
        self.bom_path = tk.StringVar()
        self.prev_path = tk.StringVar()
        self.pdf_source = tk.StringVar(value=r'C:\Coding\Python\ExcelCopyBOM\Files')
        self.include_equipment = tk.BooleanVar()
        self.find_rev = tk.BooleanVar()
        self.include_data = tk.BooleanVar()
        self.date_var = tk.StringVar()
        self.status_var = tk.StringVar(value="Ready")
        
        self._create_gui()
        
    def _create_gui(self):
        """Opretter alle GUI elementer"""
        # Main frame
        main_frame = ttk.Frame(self.root, padding="10")
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # File selection
        ttk.Label(main_frame, text="Open Excel BOM List:").grid(row=0, column=0, sticky=tk.W)
        ttk.Entry(main_frame, textvariable=self.bom_path, width=50).grid(row=0, column=1, padx=5)
        ttk.Button(main_frame, text="Browse", command=self._browse_bom).grid(row=0, column=2)
        
        ttk.Label(main_frame, text="Previous Drawing Package BOM List:").grid(row=1, column=0, sticky=tk.W)
        ttk.Entry(main_frame, textvariable=self.prev_path, width=50).grid(row=1, column=1, padx=5)
        ttk.Button(main_frame, text="Browse", command=self._browse_prev).grid(row=1, column=2)
        
        # PDF source directory
        ttk.Label(main_frame, text="PDF/DWG Source Directory:").grid(row=2, column=0, sticky=tk.W)
        ttk.Entry(main_frame, textvariable=self.pdf_source, width=50).grid(row=2, column=1, padx=5)
        ttk.Button(main_frame, text="Browse", command=self._browse_pdf_dir).grid(row=2, column=2)
        
        # Checkboxes
        self.data_check = ttk.Checkbutton(main_frame, text="Include Equipment, Valve, Instrument", 
                                         variable=self.include_equipment,
                                         command=self._on_equipment_change)
        self.data_check.grid(row=3, column=0, columnspan=3, sticky=tk.W)
        
        ttk.Checkbutton(main_frame, text="Find REV files before date", 
                       variable=self.find_rev,
                       command=self._on_rev_change).grid(row=4, column=0, columnspan=3, sticky=tk.W)
        
        self.include_data_check = ttk.Checkbutton(main_frame, text="Include Data Sheet", 
                                                 variable=self.include_data)
        self.include_data_check.grid(row=5, column=0, columnspan=3, sticky=tk.W)
        
        # Date picker
        self.date_entry = ttk.Entry(main_frame, textvariable=self.date_var, width=20)
        self.date_entry.grid(row=6, column=0, columnspan=3, sticky=tk.W)
        self.date_entry.state(['disabled'])
        
        # Progress bar
        self.progress = ttk.Progressbar(main_frame, mode='determinate', maximum=100)
        self.progress.grid(row=7, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=10)
        
        # Status label
        self.status_label = ttk.Label(main_frame, textvariable=self.status_var)
        self.status_label.grid(row=8, column=0, columnspan=3)
        
        # Start button
        ttk.Button(main_frame, text="Start", 
                   command=self._start_processing).grid(row=9, column=0, columnspan=3, pady=10)
    
    def _browse_bom(self):
        """Håndterer valg af BOM fil"""
        filename = filedialog.askopenfilename(
            title="Select Excel BOM file",
            filetypes=[("Excel files", "*.xlsx *.xls")]
        )
        if filename:
            self.bom_path.set(filename)
    
    def _browse_prev(self):
        """Håndterer valg af tidligere BOM fil"""
        filename = filedialog.askopenfilename(
            title="Select Previous Excel BOM file",
            filetypes=[("Excel files", "*.xlsx *.xls")]
        )
        if filename:
            self.prev_path.set(filename)
    
    def _browse_pdf_dir(self):
        """Håndterer valg af PDF/DWG mappe"""
        directory = filedialog.askdirectory(
            title="Select PDF/DWG source directory"
        )
        if directory:
            self.pdf_source.set(directory)
    
    def _on_equipment_change(self):
        """Håndterer ændring af equipment checkbox"""
        if not self.include_equipment.get():
            self.include_data.set(False)
            self.include_data_check.state(['disabled'])
        else:
            self.include_data_check.state(['!disabled'])
    
    def _on_rev_change(self):
        """Håndterer ændring af REV checkbox"""
        self.date_entry.state(['!disabled'] if self.find_rev.get() else ['disabled'])
    
    def _start_processing(self):
        """Starter behandling af filer"""
        if not self.bom_path.get():
            messagebox.showerror("Error", "Please select an Excel BOM file")
            return
        
        if not self.pdf_source.get():
            messagebox.showerror("Error", "Please select PDF/DWG source directory")
            return
        
        # Her skal vi kalde på de andre moduler i rækkefølge
        # Dette implementeres når de andre moduler er klar
        pass
    
    def update_progress(self, value, status_text):
        """Opdaterer progress bar og status tekst"""
        self.progress['value'] = value
        self.status_var.set(status_text)
        self.root.update_idletasks()
    
    def show_completion_dialog(self, log_text, dest_path):
        """Viser afslutningsdialog med log data"""
        dialog = tk.Toplevel(self.root)
        dialog.title("Processing Complete")
        
        # Text widget med scrollbar
        frame = ttk.Frame(dialog)
        frame.pack(padx=10, pady=10, fill=tk.BOTH, expand=True)
        
        scrollbar = ttk.Scrollbar(frame)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        text = tk.Text(frame, height=20, width=60, yscrollcommand=scrollbar.set)
        text.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.config(command=text.yview)
        
        text.insert('1.0', log_text)
        text.configure(state='disabled')
        
        # Knapper
        button_frame = ttk.Frame(dialog)
        button_frame.pack(pady=5)
        
        ttk.Button(button_frame, text="Save Log", 
                   command=lambda: self._save_log(log_text, dest_path)).pack(side=tk.LEFT, padx=5)
        ttk.Button(button_frame, text="Open Folder", 
                   command=lambda: self._open_folder(dest_path)).pack(side=tk.LEFT, padx=5)
        ttk.Button(button_frame, text="OK", 
                   command=dialog.destroy).pack(side=tk.LEFT, padx=5)
        
        # Centrér dialogen
        dialog.transient(self.root)
        dialog.grab_set()
        dialog.update_idletasks()
        x = (dialog.winfo_screenwidth() - dialog.winfo_width()) // 2
        y = (dialog.winfo_screenheight() - dialog.winfo_height()) // 2
        dialog.geometry(f'+{x}+{y}')
    
    def _save_log(self, log_text, dest_path):
        """Gemmer log til fil"""
        try:
            log_file = os.path.join(os.path.dirname(dest_path), "process_log.txt")
            with open(log_file, 'w', encoding='utf-8') as f:
                f.write(log_text)
            messagebox.showinfo("Success", f"Log saved to: {log_file}")
        except Exception as e:
            messagebox.showerror("Error", f"Could not save log: {str(e)}")
    
    def _open_folder(self, file_path):
        """Åbner Windows Stifinder i den angivne mappe"""
        folder_path = os.path.dirname(file_path)
        os.startfile(folder_path)
    
    def run(self):
        """Starter GUI"""
        self.root.mainloop()

if __name__ == "__main__":
    app = ExcelCopyBOMGUI()
    app.run() 