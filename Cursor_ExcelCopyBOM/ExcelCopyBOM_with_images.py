import os
import shutil
import re
import time
import openpyxl
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from win32com.client import Dispatch
from concurrent.futures import ThreadPoolExecutor
from pandas import read_excel
import logging
import win32com.client
import pythoncom
from typing import Dict, List, Optional, Tuple
import tempfile
from PIL import ImageGrab
import win32clipboard
from pathlib import Path

# Konfigurer logging med absolut sti
log_file = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'ExcelCopyBOM.log')
logging.basicConfig(
    filename=log_file,
    level=logging.DEBUG,
    format='%(asctime)s - %(levelname)s - %(message)s',
    filemode='w'
)

# Tilføj initial log besked
logging.info(f"Log file location: {log_file}")

class ExcelImageHandler:
    def __init__(self):
        self._excel = None
        self._workbook = None
        self._worksheet = None
        self._temp_dir = Path(tempfile.gettempdir()) / "ExcelCopyBOM"
        self._temp_dir.mkdir(exist_ok=True)
        self._temp_files = []

    def _init_excel_com(self):
        """Initialiserer Excel COM objektet"""
        logging.info("Initialiserer Excel COM objekt")
        pythoncom.CoInitialize()
        self._excel = win32com.client.Dispatch("Excel.Application")
        self._excel.Visible = False
        self._excel.DisplayAlerts = False

    def _cleanup_excel_com(self):
        """Lukker Excel ned og frigiver COM objektet"""
        if self._excel:
            logging.info("Lukker Excel ned")
            self._excel.Quit()
            pythoncom.CoUninitialize()

    def _cleanup_temp_files(self):
        """Ryd op i midlertidige filer"""
        for temp_file in self._temp_files:
            try:
                temp_file.unlink()
            except Exception as e:
                logging.warning(f"Kunne ikke slette temp fil {temp_file}: {str(e)}")

    def extract_images(self, input_file: str) -> Dict[str, str]:
        """
        Udtræk billeder fra Excel fil.
        Args:
            input_file: Sti til input Excel fil
        Returns:
            Dict med part number -> billede sti
        """
        image_mapping = {}
        
        try:
            # Initialiser Excel
            self._init_excel_com()
            
            # Åbn workbook
            logging.info(f"Åbner workbook for billede udtrækning: {input_file}")
            workbook = self._excel.Workbooks.Open(input_file)
            sheet = workbook.Sheets('BOM')  # Ændret til 'BOM' sheet
            
            # Find Thumbnail og Part Number kolonner
            thumbnail_col = None
            part_number_col = None
            for col in range(1, sheet.UsedRange.Columns.Count + 1):
                if sheet.Cells(1, col).Value == "Thumbnail":
                    thumbnail_col = col
                elif sheet.Cells(1, col).Value == "Part Number":
                    part_number_col = col
                    
            if not thumbnail_col or not part_number_col:
                raise ValueError("Kunne ikke finde Thumbnail eller Part Number kolonne")
                
            # Gennemgå hver række
            for row in range(2, sheet.UsedRange.Rows.Count + 1):
                part_number = str(sheet.Cells(row, part_number_col).Value)
                if part_number:
                    # Check om der er et billede i cellen
                    cell = sheet.Cells(row, thumbnail_col)
                    shape = None
                    
                    # Find shape i cellen
                    for s in sheet.Shapes:
                        if (s.Left >= cell.Left and 
                            s.Left <= cell.Left + cell.Width and
                            s.Top >= cell.Top and 
                            s.Top <= cell.Top + cell.Height):
                            shape = s
                            break
                    
                    if shape:
                        try:
                            # Gem billede til temp fil
                            temp_file = self._temp_dir / f"{part_number}.png"
                            self._temp_files.append(temp_file)
                            
                            # Kopier til clipboard
                            shape.Copy()
                            
                            # Gem fra clipboard som PNG
                            image = ImageGrab.grabclipboard()
                            if image:
                                image.save(temp_file, 'PNG')
                                image_mapping[part_number] = str(temp_file)
                                logging.info(f"Gemt billede for {part_number} til {temp_file}")
                            
                        except Exception as e:
                            logging.warning(f"Kunne ikke gemme billede for {part_number}: {str(e)}")
                            continue
            
            return image_mapping
            
        except Exception as e:
            logging.error(f"Fejl ved udtrækning af billeder: {str(e)}")
            raise
            
        finally:
            if workbook:
                workbook.Close(SaveChanges=False)
            self._cleanup_excel_com()

    def process_images(self, output_file: str, image_mapping: Dict[str, str]) -> None:
        """
        Indsæt billeder i Excel fil.
        Args:
            output_file: Sti til output Excel fil
            image_mapping: Dictionary med part number -> billede sti
        """
        try:
            # Initialiser Excel
            self._init_excel_com()
            
            # Åbn workbook
            logging.info(f"Åbner workbook for billede indsættelse: {output_file}")
            workbook = self._excel.Workbooks.Open(output_file)
            
            # Gennemgå alle sheets
            for sheet in workbook.Sheets:
                logging.info(f"Behandler sheet: {sheet.Name}")
                
                # Find Thumbnail og Part Number kolonner
                thumbnail_col = None
                part_number_col = None
                for col in range(1, sheet.UsedRange.Columns.Count + 1):
                    if sheet.Cells(1, col).Value == "Thumbnail":
                        thumbnail_col = col
                    elif sheet.Cells(1, col).Value == "Part Number":
                        part_number_col = col
                        
                if not thumbnail_col or not part_number_col:
                    logging.warning(f"Kunne ikke finde Thumbnail eller Part Number kolonne i sheet {sheet.Name}")
                    continue
                
                # Gennemgå hver række
                for row in range(2, sheet.UsedRange.Rows.Count + 1):
                    part_number = str(sheet.Cells(row, part_number_col).Value)
                    if part_number in image_mapping:
                        # Slet eksisterende billeder i cellen
                        cell = sheet.Cells(row, thumbnail_col)
                        for shape in sheet.Shapes:
                            if (shape.Left >= cell.Left and 
                                shape.Left <= cell.Left + cell.Width and
                                shape.Top >= cell.Top and 
                                shape.Top <= cell.Top + cell.Height):
                                shape.Delete()
                        
                        # Indsæt nyt billede
                        image_path = image_mapping[part_number]
                        if Path(image_path).exists():
                            # Indsæt billede
                            picture = sheet.Shapes.AddPicture(
                                image_path,
                                LinkToFile=False,
                                SaveWithDocument=True,
                                Left=cell.Left,
                                Top=cell.Top,
                                Width=0,  # 0 = auto width
                                Height=0   # 0 = auto height
                            )
                            
                            # Sæt billede højde til 2.38 cm (konverter fra cm til points)
                            target_height = 2.38 * 28.3465  # 1 cm = 28.3465 points
                            scale_factor = target_height / picture.Height
                            
                            picture.Height = target_height
                            picture.Width = picture.Width * scale_factor
                            
                            # Centrer i cellen
                            picture.Left = cell.Left + (cell.Width - picture.Width) / 2
                            picture.Top = cell.Top + (cell.Height - picture.Height) / 2
                            
                            logging.info(f"Indsat billede for {part_number} i række {row} i sheet {sheet.Name}")
                            
            # Gem ændringer
            workbook.Save()
            
        except Exception as e:
            logging.error(f"Fejl ved indsættelse af billeder: {str(e)}")
            raise
            
        finally:
            if workbook:
                workbook.Close(SaveChanges=True)
            self._cleanup_excel_com()
            self._cleanup_temp_files()

class ExcelCopyBOMGUI:
    def __init__(self):
        self.root = tk.Tk()
        self.root.title("Excel Copy BOM")
        self.root.attributes('-topmost', True)
        
        # Centrér vinduet
        window_width = 500
        window_height = 200
        screen_width = self.root.winfo_screenwidth()
        screen_height = self.root.winfo_screenheight()
        x = (screen_width - window_width) // 2
        y = (screen_height - window_height) // 2
        self.root.geometry(f'{window_width}x{window_height}+{x}+{y}')
        
        # Excel fil vælger
        frame = ttk.Frame(self.root, padding="10")
        frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        ttk.Label(frame, text="Excel BOM File:").grid(row=0, column=0, sticky=tk.W)
        self.file_path_var = tk.StringVar()
        self.file_entry = ttk.Entry(frame, textvariable=self.file_path_var, width=50)
        self.file_entry.grid(row=0, column=1, padx=5)
        
        ttk.Button(frame, text="Browse", command=self.browse_file).grid(row=0, column=2)
        
        # DWG Checkbox
        self.include_dwg = tk.BooleanVar()
        ttk.Checkbutton(frame, text="Include DWG files", variable=self.include_dwg).grid(row=1, column=0, columnspan=3, pady=5)
        
        # Procesbar
        self.progress_var = tk.DoubleVar()
        self.progress = ttk.Progressbar(frame, length=400, mode='determinate', variable=self.progress_var)
        self.progress.grid(row=2, column=0, columnspan=3, pady=20)
        
        # Status label
        self.status_var = tk.StringVar(value="Ready to start...")
        self.status_label = ttk.Label(frame, textvariable=self.status_var)
        self.status_label.grid(row=3, column=0, columnspan=3)
        
        # Start knap
        self.start_button = ttk.Button(frame, text="Start Processing", command=self.start_processing)
        self.start_button.grid(row=4, column=0, columnspan=3, pady=10)

    def browse_file(self):
        file_path = filedialog.askopenfilename(
            title="OPEN EXCEL BOM FILE",
            filetypes=[('Excel files', '.xlsx .xls')],
            initialdir='C:\\Working Folder\\Designs\\5-Projects'
        )
        if file_path:
            self.file_path_var.set(file_path)

    def update_progress(self, value, message):
        self.progress_var.set(value)
        self.status_var.set(message)
        self.root.update()

    def start_processing(self):
        file_path = self.file_path_var.get()
        if not file_path:
            messagebox.showerror("Error", "Please select an Excel BOM file first.")
            return
            
        self.start_button.config(state='disabled')
        self.update_progress(0, "Starting process...")
        
        try:
            self.process_bom(file_path)
        except Exception as e:
            logging.error(f"Error during processing: {str(e)}", exc_info=True)
            messagebox.showerror("Error", f"An error occurred: {str(e)}")
            self.start_button.config(state='normal')

    def process_bom(self, file_path):
        # Eksisterende variabler
        efile = os.path.basename(file_path)
        ePath = os.path.dirname(file_path)
        eBOM = os.path.abspath(file_path)
        
        # Lav en backup af den originale fil
        backup_file = os.path.join(ePath, 'original_' + efile)
        shutil.copy2(file_path, backup_file)
        
        self.update_progress(10, "Checking file permissions...")
        
        # Tjek om kildefilen er skrivebeskyttet
        if not os.access(file_path, os.W_OK):
            messagebox.showerror("ERROR", f"The Excel file '{efile}' is not writeable check the file out of the vault and try again.")
            return

        self.update_progress(20, "Creating destination folder...")
        
        # Find positionerne af de første 3 '-' i filnavnet
        dash_positions = [pos for pos, char in enumerate(efile) if char == '-']

        # Opret mappenavn ved at tage alt indtil det tredje '-' plus det første tegn efter det tredje '-'
        if len(dash_positions) >= 3:
            # Find alt indtil det tredje '-'
            base_name = efile[:dash_positions[2]]  # Fjernet +1 for at undgå ekstra '-'
            
            # Tag det første tegn efter det tredje '-'
            if len(efile) > dash_positions[2] + 1:
                base_name += efile[dash_positions[2]:dash_positions[2] + 2]  # Inkluder kun ét tegn efter '-'

            # Rens mappenavnet
            new_folder_name = base_name.strip().rstrip('-')  # Fjern ekstra mellemrum og '-'
        else:
            # Hvis der ikke er nok '-', brug hele filnavnet
            new_folder_name = os.path.splitext(efile)[0]  # Brug filnavnet uden udvidelse

        # Definer destinationen for den nye mappe
        dest_path = os.path.join(ePath, new_folder_name)

        # Sikr at mappen eksisterer
        os.makedirs(dest_path, exist_ok=True)

        destPath=(os.path.abspath(dest_path))
        destDir=(os.path.basename)
        if dest_path == '':
            os._exit(1)    

        source = (r'C:\Coding\Python\ExcelCopyBOM\Files')

        start_time = time.time()

        # Initialiser billede handler
        image_handler = ExcelImageHandler()
        
        # Udtræk billeder fra original fil
        self.update_progress(25, "Udtrækker billeder fra original fil...")
        image_mapping = image_handler.extract_images(file_path)

        # Resten af den eksisterende kode...

        # Efter at have gemt den nye fil, indsæt billederne
        self.update_progress(95, "Indsætter billeder i den nye fil...")
        image_handler.process_images(output_path, image_mapping)
        
        # Kopier den originale fil tilbage
        shutil.copy2(backup_file, file_path)
        # Slet backup filen
        os.remove(backup_file)
        
        duration=(time.time() - start_time)

        messagebox.showinfo(title='ExcelCopyBOM', message='Done\nTime: '+"{0:.2f}".format(duration)+' seconds')

        # Når processen er færdig
        self.update_progress(100, "Process completed!")
        
        # Gem reference til destPath før vi lukker vinduet
        final_path = destPath
        
        # Deaktiver start-knappen permanent
        self.start_button.config(state='disabled')
        
        # Vis afslutningsbesked og spørg om at åbne mappe
        if messagebox.askyesno("Done", "Processing completed successfully!\nWould you like to open the output folder?"):
            try:
                os.startfile(final_path)
            except Exception as e:
                logging.error(f"Error opening output folder: {str(e)}", exc_info=True)
        
        # Luk vinduet til sidst
        try:
            self.root.quit()
            self.root.destroy()
        except Exception as e:
            logging.error(f"Error closing window: {str(e)}")

def main():
    gui = ExcelCopyBOMGUI()
    gui.root.mainloop()

if __name__ == "__main__":
    main() 