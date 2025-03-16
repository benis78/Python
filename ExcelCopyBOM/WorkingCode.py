import os
import shutil
import pandas as pd
import openpyxl
import time
import sys
import subprocess
import json
import re
from openpyxl.utils import get_column_letter
from openpyxl.styles import PatternFill, Font
import tkinter as tk
from tkinter import filedialog, messagebox, ttk, Toplevel
from win32com.client import Dispatch
from concurrent.futures import ThreadPoolExecutor
import pythoncom
import threading
import glob

# Opsæt logging til debug.txt
debug_file = os.path.join(os.path.dirname(__file__), "debug.txt")
sys.stdout = open(debug_file, 'w', encoding='utf-8')

# Cache for kategorier og piping matches
category_cache = {}
piping_cache = {}

# Global kategori konfiguration
category_config = None

# Global variabler
#NETWORK_PATH = r'\\192.168.170.18\drawings'
NETWORK_PATH = r'C:\Coding\Python\ExcelCopyBOM\Files'

class ProgressWindow:
    def __init__(self, parent):
        self.window = tk.Toplevel(parent)
        self.window.title("Progress")
        self.window.attributes('-topmost', 1)
        
        # Centrér vinduet
        window_width = 400
        window_height = 150
        screen_width = self.window.winfo_screenwidth()
        screen_height = self.window.winfo_screenheight()
        x = (screen_width - window_width) // 2
        y = (screen_height - window_height) // 2
        self.window.geometry(f"{window_width}x{window_height}+{x}+{y}")
        
        # Opret main frame
        main_frame = ttk.Frame(self.window, padding="10")
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # Status label
        self.status_label = ttk.Label(main_frame, text="Forbereder...")
        self.status_label.grid(row=0, column=0, sticky=tk.W, pady=5)
        
        # Progress bar
        self.progress_var = tk.DoubleVar()
        self.progress_bar = ttk.Progressbar(main_frame, length=300, mode='determinate', 
                                          variable=self.progress_var)
        self.progress_bar.grid(row=1, column=0, sticky=(tk.W, tk.E), pady=5)
        
        # Cancel button
        self.cancel_button = ttk.Button(main_frame, text="Annuller", command=self.cancel)
        self.cancel_button.grid(row=2, column=0, pady=10)
        
        # Flag for cancellation
        self.cancelled = False
        
        # Opdater vinduet
        self.window.update()
    
    def update(self, message, progress=None):
        """Opdater status besked og progress bar."""
        if self.cancelled:
            return
            
        def _update():
            if hasattr(self, 'status_label'):
                self.status_label.config(text=message)
            if progress is not None and hasattr(self, 'progress_var'):
                self.progress_var.set(progress)
            if hasattr(self, 'window'):
                self.window.update()
        
        # Kør opdatering i hovedtråden
        if threading.current_thread() is threading.main_thread():
            _update()
        else:
            self.window.after(0, _update)
    
    def cancel(self):
        """Håndter annullering af processen."""
        self.cancelled = True
        
        def _cancel():
            if hasattr(self, 'status_label'):
                self.status_label.config(text="Annullerer...")
            if hasattr(self, 'cancel_button'):
                self.cancel_button.state(['disabled'])
            if hasattr(self, 'window'):
                self.window.update()
        
        # Kør annullering i hovedtråden
        if threading.current_thread() is threading.main_thread():
            _cancel()
        else:
            self.window.after(0, _cancel)
    
    def close(self):
        """Luk progress vinduet."""
        def _close():
            if hasattr(self, 'window'):
                self.window.destroy()
                delattr(self, 'window')
        
        # Kør lukning i hovedtråden
        if threading.current_thread() is threading.main_thread():
            _close()
        else:
            self.window.after(0, _close)

class MainWindow:
    def __init__(self):
        self.root = tk.Tk()
        self.root.title("Excel Copy BOM")
        self.root.attributes('-topmost', 1)
        
        # Centrér vinduet
        window_width = 600
        window_height = 400
        screen_width = self.root.winfo_screenwidth()
        screen_height = self.root.winfo_screenheight()
        x = (screen_width - window_width) // 2
        y = (screen_height - window_height) // 2
        self.root.geometry(f"{window_width}x{window_height}+{x}+{y}")
        
        # Opret main frame
        main_frame = ttk.Frame(self.root, padding="10")
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # Excel BOM List
        ttk.Label(main_frame, text="Open Excel BOM List:").grid(row=0, column=0, sticky=tk.W, pady=5)
        self.excel_path = tk.StringVar()
        ttk.Entry(main_frame, textvariable=self.excel_path, width=50).grid(row=0, column=1, padx=5)
        ttk.Button(main_frame, text="Browse", command=self.browse_excel).grid(row=0, column=2)
        
        # Previous Drawing Package BOM List
        ttk.Label(main_frame, text="Previous Drawing Package BOM List:").grid(row=1, column=0, sticky=tk.W, pady=5)
        self.prev_excel_path = tk.StringVar()
        ttk.Entry(main_frame, textvariable=self.prev_excel_path, width=50).grid(row=1, column=1, padx=5)
        ttk.Button(main_frame, text="Browse", command=self.browse_prev_excel).grid(row=1, column=2)
        
        # Checkboxes
        self.include_equipment = tk.BooleanVar(value=True)
        self.equipment_checkbox = ttk.Checkbutton(main_frame, text="Include Equipment, Valve, Instrument", 
                       variable=self.include_equipment, command=self.toggle_data_sheet)
        self.equipment_checkbox.grid(row=2, column=0, columnspan=3, sticky=tk.W, pady=5)
        
        self.find_rev_before_date = tk.BooleanVar(value=False)
        ttk.Checkbutton(main_frame, text="Find 'REV' files before date", 
                       variable=self.find_rev_before_date, 
                       command=self.toggle_date_frame).grid(row=3, column=0, columnspan=3, sticky=tk.W, pady=5)
        
        # Date selection frame
        self.date_frame = ttk.Frame(main_frame)
        self.date_frame.grid(row=4, column=0, columnspan=3, sticky=tk.W, pady=5)
        self.date_frame.grid_remove()  # Skjult som standard
        
        ttk.Label(self.date_frame, text="Date:").grid(row=0, column=0, padx=5)
        self.date_var = tk.StringVar(value=time.strftime("%Y-%m-%d"))
        self.date_entry = ttk.Entry(self.date_frame, textvariable=self.date_var, width=10)
        self.date_entry.grid(row=0, column=1)
        
        self.include_data_sheet = tk.BooleanVar(value=False)
        self.data_sheet_checkbox = ttk.Checkbutton(main_frame, text="Include Data Sheet", 
                       variable=self.include_data_sheet)
        self.data_sheet_checkbox.grid(row=5, column=0, columnspan=3, sticky=tk.W, pady=5)
        
        # Start button
        self.start_button = ttk.Button(main_frame, text="Start", command=self.start_processing)
        self.start_button.grid(row=6, column=0, columnspan=3, pady=20)
        
        # Progress window (initially None)
        self.progress_window = None
        
        # Initialiser checkbox states
        self.toggle_data_sheet()
    
    def toggle_date_frame(self, *args):
        """Vis/skjul dato vælger baseret på checkbox status."""
        if self.find_rev_before_date.get():
            self.date_frame.grid()
        else:
            self.date_frame.grid_remove()
    
    def toggle_data_sheet(self, *args):
        """Aktiver/deaktiver Data Sheet checkbox baseret på Equipment checkbox."""
        if not self.include_equipment.get():
            self.include_data_sheet.set(False)
            self.data_sheet_checkbox.state(['disabled'])
        else:
            self.data_sheet_checkbox.state(['!disabled'])
    
    def browse_excel(self):
        """Åbn filvælger for Excel BOM List."""
        file_path = filedialog.askopenfilename(
            title="Vælg Excel BOM fil",
            initialdir=r'C:\Working Folder\Designs\5-Projects',
            filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")]
        )
        if file_path:
            self.excel_path.set(file_path)
    
    def browse_prev_excel(self):
        """Åbn filvælger for Previous Drawing Package BOM List."""
        file_path = filedialog.askopenfilename(
            title="Vælg Previous Drawing Package BOM fil",
            initialdir=r'C:\Working Folder\Designs\5-Projects',
            filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")]
        )
        if file_path:
            self.prev_excel_path.set(file_path)
    
    def start_processing(self):
        """Start behandlingen af filerne."""
        if not self.excel_path.get():
            messagebox.showerror("Error", "Vælg venligst en Excel BOM fil")
            return
        
        # Deaktiver start knappen
        self.start_button.state(['disabled'])
        
        # Opret progress window i hovedtråden
        self.root.after(0, self.create_progress_window)
    
    def create_progress_window(self):
        """Opret progress window i hovedtråden."""
        self.progress_window = ProgressWindow(self.root)
        
        # Start behandling i en separat tråd
        thread = threading.Thread(target=self.process_files)
        thread.daemon = True
        thread.start()
    
    def process_files(self):
        """Hovedfunktion for filbehandling."""
        try:
            # Initialiser COM
            pythoncom.CoInitialize()
            
            # Start behandling
            main(self.excel_path.get(), self.prev_excel_path.get(), 
                 self.include_equipment.get(), self.find_rev_before_date.get(),
                 self.date_var.get(), self.include_data_sheet.get(),
                 self.progress_window)
            
        except Exception as e:
            # Gem fejlbeskeden
            error_msg = str(e)
            # Vis fejl i hovedtråden
            self.root.after(0, lambda: messagebox.showerror("Error", error_msg))
        finally:
            # Ryd op
            if self.progress_window:
                # Luk progress window i hovedtråden
                self.root.after(0, self.cleanup_progress_window)
            pythoncom.CoUninitialize()
    
    def cleanup_progress_window(self):
        """Ryd op efter behandling i hovedtråden."""
        if self.progress_window:
            self.progress_window.close()
            self.progress_window = None
        self.start_button.state(['!disabled'])
    
    def run(self):
        """Start GUI event loop."""
        self.root.mainloop()

def show_success_message(message, parent=None):
    """Viser en success besked i et Toplevel vindue der altid er øverst."""
    if parent is None:
        parent = tk.Tk()
        parent.withdraw()  # Skjul hovedvinduet hvis der ikke er et parent vindue
    
    top = tk.Toplevel(parent)
    top.title("Success")
    top.attributes('-topmost', 1)
    
    # Centrér vinduet
    window_width = 400
    window_height = 150
    screen_width = top.winfo_screenwidth()
    screen_height = top.winfo_screenheight()
    x = (screen_width - window_width) // 2
    y = (screen_height - window_height) // 2
    top.geometry(f"{window_width}x{window_height}+{x}+{y}")
    
    # Tilføj besked og OK knap
    ttk.Label(top, text=message, wraplength=350, justify='center').pack(pady=20)
    
    def on_ok():
        top.destroy()
        if parent.winfo_name() == '.':  # Hvis det er et midlertidigt hovedvindue
            parent.destroy()
    
    ttk.Button(top, text="OK", command=on_ok).pack(pady=10)

def choose_file():
    """Lader brugeren vælge en Excel-fil og returnerer dens sti."""
    file_path = "C:\\Coding\\Python\\ExcelCopyBOM\\4003-615-A01-E - BOM.xlsx"
    # file_path = filedialog.askopenfilename(
    #     title="Vælg Excel BOM fil",
    #     initialdir=r'C:\Working Folder\Designs\5-Projects',
    #     filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")]
    # )
    return file_path

def check_network_path(network_path):
    """Tjekker hurtigt om stien er tilgængelig."""
    return os.path.exists(network_path)

def extract_revision_from_partnumber(part_number):
    """Udtrækker og fjerner revisionsbogstav fra part number hvis det findes."""
    parts = part_number.split('-')
    
    if len(parts) >= 3:
        last_part = parts[-1]
        if len(last_part) == 1 and last_part.isalpha():
            revision = last_part
            clean_partnumber = '-'.join(parts[:-1])
            return clean_partnumber, revision
    
    return part_number, None

def get_drawing_status(part_number, files, source_part_number=None, rev=None):
    """Bestemmer drawing status for et part number."""
    # Hvis der er angivet source_part_number og rev, returner special besked
    if source_part_number and rev:
        return f"(See {source_part_number}-{rev} BOM)"
    
    # Tjek først for "For Review" filer
    has_review = any(f.endswith('.pdf') and 'for review' in os.path.basename(f).lower() and os.path.basename(f).startswith(part_number) for f in files)
    if has_review:
        return "PDF For Review"
    
    # Tjek for normale filer (ignorer "For Review" filer)
    has_pdf = any(f.endswith('.pdf') and 'for review' not in os.path.basename(f).lower() and os.path.basename(f).startswith(part_number) for f in files)
    has_dwg = any(f.endswith('.dwg') and os.path.basename(f).startswith(part_number) for f in files)
    
    if has_pdf and has_dwg:
        return "PDF_DWG"
    elif has_pdf:
        return "PDF"
    elif has_dwg:
        return "DWG"
    else:
        return "NA"

def has_drawings_for_category(category_data, files):
    """Tjekker om der findes PDF eller DWG filer for en given kategori."""
    for _, row in category_data.iterrows():
        part_number = str(row.iloc[1]).strip()  # Kolonne 2 (Part Number)
        if any(os.path.basename(f).startswith(part_number) for f in files):
            return True
    return False

def should_include_row(row, df):
    """Afgør om en række skal inkluderes baseret på BOM Structure."""
    item_number = str(row.iloc[0])
    
    # Find BOM Structure kolonnen
    bom_structure_col = None
    for col in df.columns:
        if "BOM STRUCTURE" in str(col).upper():
            bom_structure_col = col
            break
    
    if bom_structure_col is None:
        return True  # Hvis kolonnen ikke findes, inkluder alle rækker
    
    # Tjek om denne række er markeret som "Inseparable"
    current_structure = str(row[bom_structure_col]).strip().upper()
    if current_structure == "INSEPARABLE":
        return True
    
    # Tjek om denne række er under en "Inseparable" række
    if '.' in item_number:
        parts = item_number.split('.')
        current_parent = ""
        for part in parts[:-1]:
            if current_parent:
                current_parent += "."
            current_parent += part
            
            # Find parent rækken
            parent_mask = df.iloc[:, 0] == current_parent
            if parent_mask.any():
                parent_row = df[parent_mask].iloc[0]
                parent_structure = str(parent_row[bom_structure_col]).strip().upper()
                if parent_structure == "INSEPARABLE":
                    return False
    
    return True

def load_category_config():
    """Indlæser kategori konfigurationen fra JSON-filen."""
    global category_config
    config_path = os.path.join(os.path.dirname(__file__), "categories.json")
    print(f"Forsøger at indlæse kategori konfiguration fra: {config_path}")
    try:
        if not os.path.exists(config_path):
            raise FileNotFoundError(f"Kunne ikke finde filen: {config_path}")
            
        with open(config_path, 'r', encoding='utf-8') as f:
            try:
                file_content = f.read()
                print(f"Fil indhold: {file_content[:200]}...")  # Vis de første 200 tegn
                try:
                    category_config = json.loads(file_content)
                    print("JSON indlæst succesfuldt")
                    print(f"Kategorier fundet: {list(category_config.get('categories', {}).keys())}")
                    
                    # Validér struktur
                    if not isinstance(category_config, dict):
                        raise ValueError("Root objekt er ikke et dictionary")
                    if 'categories' not in category_config:
                        raise ValueError("Mangler 'categories' key i root objekt")
                    if not isinstance(category_config['categories'], dict):
                        raise ValueError("'categories' er ikke et dictionary")
                        
                    print("JSON struktur valideret succesfuldt")
                    return category_config
                except json.JSONDecodeError as je:
                    print(f"JSON parsing fejl på position {je.pos}: {je.msg}")
                    print(f"Problematisk linje: {je.doc.splitlines()[je.lineno-1]}")
                    raise
            except Exception as e:
                print(f"Generel fejl ved parsing af JSON: {str(e)}")
                raise
    except Exception as e:
        print(f"Fejl ved indlæsning af kategori konfiguration: {str(e)}")
        messagebox.showerror("ERROR", f"Kunne ikke indlæse kategori konfiguration: {str(e)}")
        return None

def find_category(part_number):
    """Find kategori for et part number ved hjælp af regex mønstre."""
    if not part_number or not isinstance(part_number, str):
        return category_config['categories']['default_category']
    
    # Tjek cache først
    if part_number in category_cache:
        return category_cache[part_number]
    
    # Split part number i grupper
    groups = part_number.split('-')
    if not groups:
        return category_config['categories']['default_category']
    
    group1 = groups[0]
    group2 = groups[1] if len(groups) > 1 else None
    group3 = groups[2] if len(groups) > 2 else None
    
    # Tjek for 0000 (Basic Components og Suppliers Parts)
    if group1 == "0000" and group2 and group3:
        for pattern in category_config['categories']['0000']['patterns']:
            if (re.match(pattern['group2'], group2) and 
                re.match(pattern['group3'], group3)):
                result = pattern['category']
                category_cache[part_number] = result
                return result
    
    # Tjek for projekt numre (2000-9999)
    try:
        project_num = int(group1)
        if 2000 <= project_num <= 9999:
            if group2:
                for pattern in category_config['categories']['project_numbers']['patterns']:
                    if re.match(pattern['group2'], group2):
                        result = pattern['category']
                        category_cache[part_number] = result
                        return result
    except ValueError:
        pass
    
    # Tjek for Area Drawings kategorier
    if group2 and re.match(r'^[A-Za-z0-9]{2,4}$', group2):
        if group3:
            # Tjek for piping mønstre
            for pattern, category in category_config['categories']['area_drawings']['piping_patterns'].items():
                if re.match(pattern, group3):
                    result = category
                    category_cache[part_number] = result
                    return result
            
            # Tjek for andre Area Drawings mønstre
            for pattern, category in category_config['categories']['area_drawings']['group3_patterns'].items():
                if re.match(pattern, group3):
                    result = category
                    category_cache[part_number] = result
                    return result
    
    # Hvis ingen match fundet, returner default kategori
    result = category_config['categories']['default_category']
    category_cache[part_number] = result
    return result

def categorize_data(sheet, file_path):
    """Optimeret kategorisering af data med regex-baseret matching."""
    try:
        print("\n=== KATEGORISERING START ===")
        
        # Konverter Excel data til pandas DataFrame
        data = []
        headers = []
        
        # Læs headers
        for col in range(1, sheet.UsedRange.Columns.Count + 1):
            header = str(sheet.Cells(1, col).Value).strip().upper()
            headers.append(header)
        
        # Læs data
        for row in range(2, sheet.UsedRange.Rows.Count + 1):
            row_data = []
            for col in range(1, sheet.UsedRange.Columns.Count + 1):
                value = sheet.Cells(row, col).Value
                row_data.append(str(value) if value is not None else "")
            data.append(row_data)
        
        # Opret DataFrame
        df = pd.DataFrame(data, columns=headers)
        
        # Verificer nødvendige kolonner
        required_cols = ['ITEM', 'PART NUMBER']
        missing_cols = [col for col in required_cols if col not in df.columns]
        if missing_cols:
            raise ValueError(f"Manglende kolonner: {', '.join(missing_cols)}")
        
        # Opret base_number kolonne for gruppering
        df['BASE_NUMBER'] = df['ITEM'].astype(str).str.split('.').str[0]
        
        # Kategoriser data
        categorized_data = {}
        for base_number, group in df.groupby('BASE_NUMBER'):
            try:
                # Find parent række
                parent_mask = group['ITEM'].astype(str) == str(base_number)
                parent_row = group[parent_mask].iloc[0] if parent_mask.any() else group.iloc[0]
                part_number = str(parent_row['PART NUMBER'])
                
                # Find kategori
                category = find_category(part_number)
                
                # Tilføj til kategori
                if category not in categorized_data:
                    categorized_data[category] = pd.DataFrame(columns=df.columns)
                categorized_data[category] = pd.concat([categorized_data[category], group], ignore_index=True)
                
            except Exception as e:
                print(f"Fejl ved kategorisering af {base_number}: {str(e)}")
                continue
        
        # Log resultater
        print("\n=== KATEGORISERING AFSLUTTET ===")
        for category, data in categorized_data.items():
            if len(data) > 0:
                print(f"{category}: {len(data)} items")
        
        return categorized_data
        
    except Exception as e:
        print(f"Fejl under kategorisering: {str(e)}")
        raise

def create_category_sheets(wb, categorized_data):
    """Opretter faner for hver kategori i Excel."""
    try:
        # Opret faner for hver kategori
        for category, data in categorized_data.items():
            if data.empty:
                continue
                
            # Opret ny fane
            new_sheet = wb.Worksheets.Add(After=wb.Worksheets(wb.Worksheets.Count))
            new_sheet.Name = category
            
            # Kopier header fra BOM (Raw)
            wb.Sheets("BOM (Raw)").Range("1:1").Copy(new_sheet.Range("1:1"))
            
            # Kopier data
            for idx, row in data.iterrows():
                source_row = idx + 2  # +2 fordi vi starter fra række 2 og har header
                target_row = idx + 2
                wb.Sheets("BOM (Raw)").Range(f"{source_row}:{source_row}").Copy(
                    new_sheet.Range(f"{target_row}:{target_row}")
                )
            
            # Gruppér rækker hvis det er en piping kategori
            if "Piping" in category:
                group_by_parent_items(new_sheet)
            
            # Formater arket
            format_excel_sheet(new_sheet)
        
        return True
        
    except Exception as e:
        print(f"Fejl ved oprettelse af kategori faner: {str(e)}")
        raise

def should_include_excel_row(item_number, sheet):
    """Excel version af should_include_row funktionen."""
    if not item_number or '.' not in str(item_number):
        return True
    
    # Find den aktuelle række
    current_range = sheet.Range("A:A").Find(item_number)
    if current_range:
        current_structure = str(sheet.Cells(current_range.Row, 5).Value).strip().upper()  # Kolonne E (BOM Structure)
        if current_structure == "INSEPARABLE":
            return True
    
    # Split item number i dele
    parts = item_number.split('.')
    current_parent = ""
    for part in parts[:-1]:
        if current_parent:
            current_parent += "."
        current_parent += part
        
        # Find parent row
        parent_range = sheet.Range("A:A").Find(current_parent)
        if parent_range:
            parent_row = parent_range.Row
            parent_structure = str(sheet.Cells(parent_row, 5).Value).strip().upper()
            if parent_structure == "INSEPARABLE":
                return False
    
    return True

def process_excel(file_path, progress_window):
    """Behandler Excel-filen og returnerer kategoriseret data."""
    try:
        # Initialiser Excel
        excel = Dispatch("Excel.Application")
        excel.Visible = False
        excel.DisplayAlerts = False
        
        # Åbn kildefilen
        wb = excel.Workbooks.Open(file_path)
        sheet = wb.Sheets(1)
        
        # Opret "Bom (Raw)" fane
        progress_window.update("Opretter Bom (Raw) fane...")
        if sheet.Name != "Bom (Raw)":
            sheet.Name = "Bom (Raw)"
        
        # Identificer kolonne numre
        progress_window.update("Identificerer kolonner...")
        headers = {}
        for i in range(1, sheet.UsedRange.Columns.Count + 1):
            header = str(sheet.Cells(1, i).Value).strip().upper()
            if header in ["ITEM", "PART NUMBER", "REV", "BOM STRUCTURE", "DESCRIPTION", "QTY", "D", "T", "L"]:
                headers[header] = i
        
        # Udtræk information fra filnavn
        file_name = os.path.basename(file_path)
        first_4_digits = file_name[:4]
        rev_from_filename = extract_revision_from_partnumber(file_name)
        description = "Arrangement Drawing" if "A" in file_name else "Basic Equipment Drawing"
        
        # Indsæt linje i række 2
        progress_window.update("Indsætter arrangement linje...")
        sheet.Cells(2, headers["ITEM"]).Value = 0
        sheet.Cells(2, headers["PART NUMBER"]).Value = first_4_digits
        sheet.Cells(2, headers["REV"]).Value = rev_from_filename
        sheet.Cells(2, headers["BOM STRUCTURE"]).Value = "Inseparable"
        sheet.Cells(2, headers["DESCRIPTION"]).Value = description
        sheet.Cells(2, headers["QTY"]).Value = 1
        sheet.Cells(2, headers["D"]).Value = 1
        sheet.Cells(2, headers["T"]).Value = 1
        sheet.Cells(2, headers["L"]).Value = 1
        
        # Håndter "Part Number" kolonnen
        progress_window.update("Håndterer Part Number kolonnen...")
        rows_to_delete = []
        for row in range(2, sheet.UsedRange.Rows.Count + 1):
            part_number = str(sheet.Cells(row, headers["PART NUMBER"]).Value).strip()
            if part_number.startswith(("0000-700", "0000-701", "0000-702")):
                rows_to_delete.append(row)
            elif part_number.isalpha():
                rows_to_delete.append(row)
        
        # Slet markerede rækker
        for row in reversed(rows_to_delete):
            sheet.Rows(row).Delete()
        
        # Håndter "BOM Structure" kolonnen
        progress_window.update("Håndterer BOM Structure...")
        rows_to_delete = []
        for row in range(2, sheet.UsedRange.Rows.Count + 1):
            bom_structure = str(sheet.Cells(row, headers["BOM STRUCTURE"]).Value).strip()
            part_number = str(sheet.Cells(row, headers["PART NUMBER"]).Value).strip()
            
            if bom_structure == "Phantom":
                rows_to_delete.append(row)
            elif bom_structure == "Inseparable" or part_number.startswith("0000-3"):
                # Find og slet alle child rækker
                parent_item = str(sheet.Cells(row, headers["ITEM"]).Value).strip()
                for child_row in range(row + 1, sheet.UsedRange.Rows.Count + 1):
                    child_item = str(sheet.Cells(child_row, headers["ITEM"]).Value).strip()
                    if not child_item.startswith(parent_item):
                        break
                    rows_to_delete.append(child_row)
        
        # Slet markerede rækker
        for row in reversed(rows_to_delete):
            sheet.Rows(row).Delete()
    
    # Beregn Total QTY
        progress_window.update("Beregner Total QTY...")
        total_qty_col = headers["QTY"] + 1
        sheet.Cells(1, total_qty_col).Value = "Total QTY"
        
        for row in range(2, sheet.UsedRange.Rows.Count + 1):
            item = str(sheet.Cells(row, headers["ITEM"]).Value).strip()
            qty = float(sheet.Cells(row, headers["QTY"]).Value or 0)
            
            # Find parent QTY
            parent_qty = 1
            if '.' in item:
                parent_item = '.'.join(item.split('.')[:-1])
                for parent_row in range(2, row):
                    if str(sheet.Cells(parent_row, headers["ITEM"]).Value).strip() == parent_item:
                        parent_qty = float(sheet.Cells(parent_row, total_qty_col).Value or 0)
                    break
            
            sheet.Cells(row, total_qty_col).Value = qty * parent_qty
        
        # Gem ændringer
        wb.Save()
        
        return sheet, headers, wb
        
    except Exception as e:
        print(f"Fejl under Excel behandling: {str(e)}")
        raise
    finally:
        if 'excel' in locals():
            excel.Quit()

def set_cell_color(cell, status):
    """Sætter farve på cellen baseret på status."""
    if status == "PDF_DWG":
        cell.Font.Color = rgb_to_int(0, 128, 0)  # Grøn
    elif status in ["PDF", "DWG"]:
        cell.Font.Color = rgb_to_int(255, 192, 0)  # Gul
    elif status in ["NA", "PDF For Review"]:
        cell.Font.Color = rgb_to_int(255, 0, 0)  # Rød

def rgb_to_int(red, green, blue):
    """Konverterer RGB værdier til Excel farve integer."""
    return red + (green * 256) + (blue * 256 * 256)

def newRev(name, files):
    """Finder den seneste revision ud fra filnavnet."""
    latest_files = {}
    latest_revs = {}
    
    if not files:
        return latest_files, latest_revs
    
    for ext in [".pdf", ".dwg"]:
        relevant_files = [f for f in files if f.endswith(ext) and 'for review' not in os.path.basename(f).lower()]
        if relevant_files:
            latest_rev = relevant_files[0]
            rev_char = '-'
            
            for file in relevant_files:
                rev_pos = file.find(str(name)) + len(str(name)) + 1
                if rev_pos < len(file):
                    current_rev = file[rev_pos]
                    if current_rev > rev_char:
                        latest_rev = file
                        rev_char = current_rev
            
            latest_files[ext] = latest_rev
            latest_revs[ext] = rev_char
    
    return latest_files, latest_revs

def scan_directory_concurrent(pdf_source):
    """Scanner hele PDF-kataloget for PDF og DWG filer."""
    # Brug global cache til at undgå gentagne scanninger
    global _file_cache
    if '_file_cache' not in globals():
        _file_cache = {}
    
    # Tjek om vi allerede har scannet denne mappe
    if pdf_source in _file_cache:
        print(f"\n=== BRUGER CACHED FILSCAN FOR: {pdf_source} ===")
        return _file_cache[pdf_source]
    
    print(f"\n=== SCANNER MAPPE: {pdf_source} ===")
    print("Dette kan tage et øjeblik...")
    
    def scan_directory_task(entry):
        if entry.is_dir(follow_symlinks=False):
            file_paths = []
            for sub_entry in os.scandir(entry.path):
                file_paths.extend(scan_directory_task(sub_entry))
            return file_paths
        elif entry.is_file(follow_symlinks=False) and (entry.name.endswith(".pdf") or entry.name.endswith(".dwg")):
            return [entry.path]
        return []

    file_paths = []
    with ThreadPoolExecutor() as executor:
        futures = [executor.submit(scan_directory_task, entry) for entry in os.scandir(pdf_source)]
        for future in futures:
            file_paths.extend(future.result())
    
    # Gem resultatet i cachen
    _file_cache[pdf_source] = file_paths
    print(f"Fandt {len(file_paths)} filer")
    return file_paths

def get_source_folder(default_network_path):
    """Håndterer valg af kildemappe med fejlhåndtering for netværksdrev."""
    if check_network_path(default_network_path):
        return default_network_path
    
    response = messagebox.askquestion(
        "Netværksdrev utilgængeligt",
        f"Kan ikke få adgang til netværksdrevet:\n{default_network_path}\n\n"
        "Vil du vælge en lokal mappe i stedet?",
        icon='warning'
    )
    
    if response == 'yes':
        local_path = filedialog.askdirectory(
            title="Vælg mappe med PDF/DWG filer",
            initialdir=r"C:\Working Folder\Designs\5-Projects"
        )
        if local_path:
            return local_path
        else:
            raise Exception("Ingen mappe valgt")
    else:
        raise Exception("Kan ikke fortsætte uden adgang til PDF/DWG filer")

def process_excel_file(excel_path, files, progress_window, categorized_data):
    """Håndterer al Excel-relateret behandling."""
    excel = None
    wb = None
    sheet = None
    headers = {}  # Initialiser headers dictionary
    rows_to_delete = []  # Initialiser liste over rækker der skal slettes
    
    try:
        excel = Dispatch("Excel.Application")
        excel.DisplayAlerts = False
        excel.Visible = False
        
        wb = excel.Workbooks.Open(excel_path)
        sheet = wb.Sheets(1)
        sheet.Name = "BOM (Raw)"
        
        # Find nødvendige kolonner
        for i in range(1, sheet.UsedRange.Columns.Count + 1):
            header = str(sheet.Cells(1, i).Value).strip().upper()
            if header in ["ITEM", "PART NUMBER", "REV", "BOM STRUCTURE", "QTY", "DRAWINGS"]:
                headers[header] = i
        
        # Tilføj manglende kolonner
        if "DRAWINGS" not in headers:
            headers["DRAWINGS"] = sheet.UsedRange.Columns.Count + 1
            sheet.Cells(1, headers["DRAWINGS"]).Value = "DRAWINGS"
        
        if "QTY" in headers:
            headers["TOTAL QTY"] = headers["QTY"] + 1
            sheet.Columns(headers["TOTAL QTY"]).Insert()
            sheet.Cells(1, headers["TOTAL QTY"]).Value = "Total QTY"
        
        # Opret liste over rækker der skal slettes
        for row in range(2, sheet.UsedRange.Rows.Count + 1):
            item_number = str(sheet.Cells(row, 1).Value)
            if not should_include_excel_row(item_number, sheet):
                rows_to_delete.append(row)
                continue
            
            # Opdater række information hvis headers er korrekt initialiseret
            if headers:
                update_row_information(sheet, row, headers, files)
        
        # Slet uønskede rækker (fra bunden)
        for row in sorted(rows_to_delete, reverse=True):
            sheet.Rows(row).Delete()
        
        # Opret kategori faner
        print("\n=== OPRETTER KATEGORI FANER ===")
        for category, data in categorized_data.items():
            if not data.empty:
                print(f"\nBehandler {category}...")
                target_sheet = None
                
                # Find eller opret fanen
                try:
                    target_sheet = wb.Sheets(category)
                except:
                    target_sheet = wb.Worksheets.Add(After=wb.Worksheets(wb.Worksheets.Count))
                    target_sheet.Name = category
                    sheet.Range("1:1").Copy(target_sheet.Range("1:1"))
                
                if target_sheet:
                    # Kopier data
                    for idx, (item_number, part_number, source_row) in enumerate(data, start=2):
                        sheet.Range(f"{source_row}:{source_row}").Copy(target_sheet.Range(f"{idx}:{idx}"))
                    
                    # Gruppér rækker hvis det er en piping kategori
                    if "Piping" in category:
                        last_row = target_sheet.UsedRange.Rows.Count
                        if last_row > 1:  # Kun hvis der er data
                            print(f"Grupperer rækker 1-{last_row} i {category}")
                            target_sheet.Range(f"1:{last_row}").Rows.Group()
                            target_sheet.Outline.ShowLevels(RowLevels=8)
        
        # Flyt BOM (Raw) forrest og gem
        if sheet:
            sheet.Move(Before=wb.Sheets(1))
            wb.Save()
        
        return sheet
        
    except Exception as e:
        print(f"Fejl i process_excel_file: {str(e)}")
        raise
        
    finally:
        if wb:
            try:
                wb.Close(SaveChanges=False)
            except:
                pass
        if excel:
            try:
                excel.Quit()
            except:
                pass

def update_row_information(sheet, row, headers, files):
    """Opdaterer information for en enkelt række i Excel."""
    try:
        # Tjek input parametre
        if not sheet or not headers or not files:
            return
            
        part_number = str(sheet.Cells(row, 2).Value)  # Kolonne B (Part Number)
        if not part_number or not part_number.strip():
            return
        
        # Initialiser variabler
        clean_partnumber = part_number
        extracted_rev = None
        matching_files = []
        drawing_status = "NA"  # Standard værdi hvis ingen filer findes
        
        # Håndter revision
        clean_partnumber, extracted_rev = extract_revision_from_partnumber(part_number)
        if extracted_rev and "REV" in headers:
            sheet.Cells(row, 2).Value = clean_partnumber
            sheet.Cells(row, headers["REV"]).Value = extracted_rev
        
        # Find matchende filer
        matching_files = [f for f in files if os.path.basename(f).startswith(clean_partnumber if extracted_rev else part_number)]
        
        # Opdater revision hvis nødvendigt
        if matching_files and "REV" in headers and not extracted_rev:
            latest_files, latest_revs = newRev(part_number, matching_files)
            if latest_revs:
                rev = next(iter(latest_revs.values()))
                if rev != '-':  # Kun opdater hvis der faktisk er fundet en revision
                    sheet.Cells(row, headers["REV"]).Value = rev
        
        # Opdater drawing status
        if "DRAWINGS" in headers:
            drawing_status = get_drawing_status(part_number, matching_files)
            cell = sheet.Cells(row, headers["DRAWINGS"])
            if cell:
                cell.Value = drawing_status
                set_cell_color(cell, drawing_status)
                
    except Exception as e:
        print(f"Fejl i update_row_information for række {row}: {str(e)}")
        # Fortsæt med næste række i stedet for at stoppe hele processen

def find_latest_files(part_number, files):
    """Finder de seneste filer for et part number."""
    latest_files = {}
    
    # Find alle matchende filer
    matching_files = [f for f in files if os.path.basename(f).startswith(part_number)]
    if not matching_files:
        return latest_files
    
    # Gruppér filer efter type (PDF/DWG)
    for file in matching_files:
        ext = os.path.splitext(file)[1].lower()
        if ext in ['.pdf', '.dwg']:
            if ext not in latest_files or os.path.getmtime(file) > os.path.getmtime(latest_files[ext]):
                latest_files[ext] = file
    
    return latest_files

def copy_files_to_destination(dest_path, categorized_data):
    """Kopierer filer til deres respektive mapper."""
    try:
        os.makedirs(dest_path, exist_ok=True)
        for category, data in categorized_data.items():
            if data.empty:
                continue
            category_path = os.path.join(dest_path, category)
            os.makedirs(category_path, exist_ok=True)
            for _, row in data.iterrows():
                part_number = str(row['PART NUMBER']).strip()
                if not part_number:
                    continue
                latest_files = find_latest_files(part_number, files)
                for file in latest_files.values():
                    try:
                        dest_file = os.path.join(category_path, os.path.basename(file))
                        shutil.copy2(file, dest_file)
                    except Exception as e:
                        print(f"Fejl ved kopiering af {file}: {str(e)}")
        return True
    except Exception as e:
        print(f"Fejl under kopiering af filer: {str(e)}")
        raise

def kill_excel(progress_window):
    """Lukker alle Excel processer med advarsel."""
    dialog = Toplevel(progress_window.window)
    dialog.title("Luk Excel")
    dialog.attributes('-topmost', 1)
    
    # Centrér vinduet
    window_width = 400
    window_height = 150
    screen_width = dialog.winfo_screenwidth()
    screen_height = dialog.winfo_screenheight()
    x = (screen_width - window_width) // 2
    y = (screen_height - window_height) // 2
    dialog.geometry(f"{window_width}x{window_height}+{x}+{y}")
    
    # Tilføj besked og knapper
    ttk.Label(dialog, text="Programmet skal lukke alle åbne Excel vinduer.\n\n"
              "Gem venligst dine åbne Excel dokumenter og klik 'Ja' for at fortsætte.",
              wraplength=350, justify='center').pack(pady=20)
    
    result = [False]  # Brug en liste for at kunne ændre værdien inde i funktionen
    
    def on_yes():
        result[0] = True
        dialog.destroy()
    
    def on_no():
        result[0] = False
        dialog.destroy()
    
    button_frame = ttk.Frame(dialog)
    button_frame.pack(pady=10)
    ttk.Button(button_frame, text="Ja", command=on_yes).pack(side='left', padx=10)
    ttk.Button(button_frame, text="Nej", command=on_no).pack(side='left', padx=10)
    
    # Gør vinduet modalt
    dialog.transient(progress_window.window)
    dialog.grab_set()
    dialog.wait_window()
    
    if result[0]:
        try:
            subprocess.run(['taskkill', '/F', '/IM', 'excel.exe'], 
                          stdout=subprocess.DEVNULL, 
                          stderr=subprocess.DEVNULL)
            time.sleep(2)  # Vent på at processerne lukkes
            return True
        except:
            pass
    return False

def process_excel_in_order(file_path, progress_window, files):
    """Behandler Excel-filen i den specificerede rækkefølge."""
    excel = None
    wb = None
    try:
        # Sikr at alle Excel processer er lukket
        if not kill_excel(progress_window):
            raise Exception("Processen blev annulleret af brugeren")
        
        # 1. Opret kopi af Excel filen
        progress_window.update("Opretter kopi af Excel fil...")
        base_name = os.path.basename(file_path)
        if " - BOM" in base_name:
            base_name = base_name.replace(" - BOM.xlsx", "")
        dest_folder = os.path.join(os.path.dirname(file_path), base_name)
        os.makedirs(dest_folder, exist_ok=True)
        dest_path = os.path.join(dest_folder, base_name + ".xlsx")
        
        # Kopier filen manuelt først
        shutil.copy2(file_path, dest_path)
        time.sleep(1)  # Vent et sekund
        
        # Initialiser Excel
        try:
            excel = Dispatch("Excel.Application")
            excel.DisplayAlerts = False
            excel.Visible = False
            time.sleep(1)  # Vent på Excel initialization
        except Exception as e:
            raise Exception(f"Kunne ikke starte Excel: {str(e)}")
        
        # Åbn workbook
        try:
            wb = excel.Workbooks.Open(dest_path)
            time.sleep(1)  # Vent på at filen åbnes
        except Exception as e:
            raise Exception(f"Kunne ikke åbne Excel-filen: {str(e)}")

        sheet = wb.Sheets(1)
        sheet.Name = "BOM (Raw)"  # Omdøb første ark
        
        # 2. Identificer kolonne numre
        progress_window.update("Identificerer kolonner...")
        headers = {}
        for i in range(1, sheet.UsedRange.Columns.Count + 1):
            header = str(sheet.Cells(1, i).Value).strip().upper()
            if header in ["ITEM", "PART NUMBER", "REV", "BOM STRUCTURE", "QTY", "DRAWINGS"]:
                headers[header] = i
            last_column = i
        
        if not all(key in headers for key in ["ITEM", "PART NUMBER", "BOM STRUCTURE", "QTY"]):
            raise Exception("Kunne ikke finde alle nødvendige kolonner")
        
        # Tilføj manglende kolonner
        if "REV" not in headers:
            headers["REV"] = last_column + 1
            sheet.Columns(headers["REV"]).Insert()
            sheet.Cells(1, headers["REV"]).Value = "REV"
            last_column += 1
        
        if "DRAWINGS" not in headers:
            headers["DRAWINGS"] = last_column + 1
            sheet.Columns(headers["DRAWINGS"]).Insert()
            sheet.Cells(1, headers["DRAWINGS"]).Value = "DRAWINGS"
            last_column += 1
        
        # 3. Håndter "Inseparable" rækker
        progress_window.update("Håndterer Inseparable rækker...")
        rows_to_delete = []
        for row in range(2, sheet.UsedRange.Rows.Count + 1):
            item_number = str(sheet.Cells(row, headers["ITEM"]).Value)
            bom_structure = str(sheet.Cells(row, headers["BOM STRUCTURE"]).Value).strip().upper()
            
            if bom_structure == "INSEPARABLE":
                # Find og marker children til sletning
                for child_row in range(2, sheet.UsedRange.Rows.Count + 1):
                    child_item = str(sheet.Cells(child_row, headers["ITEM"]).Value)
                    if child_item.startswith(item_number + '.'):
                        rows_to_delete.append(child_row)
        
        # Slet markerede rækker (fra bunden)
        for row in sorted(rows_to_delete, reverse=True):
            sheet.Rows(row).Delete()
        
        # 4. Håndter revisioner og opdater drawing status
        progress_window.update("Håndterer revisioner...")
        # Brug de allerede scannede filer
        for row in range(2, sheet.UsedRange.Rows.Count + 1):
            part_number = str(sheet.Cells(row, headers["PART NUMBER"]).Value)
            if part_number.strip():
                # Håndter revision
                clean_partnumber, extracted_rev = extract_revision_from_partnumber(part_number)
                if extracted_rev:
                    sheet.Cells(row, headers["PART NUMBER"]).Value = clean_partnumber
                    sheet.Cells(row, headers["REV"]).Value = extracted_rev
                
                # Find matchende filer og opdater revision hvis nødvendigt
                matching_files = [f for f in files if os.path.basename(f).startswith(clean_partnumber if extracted_rev else part_number)]
                if matching_files and not extracted_rev:
                    latest_files, latest_revs = newRev(part_number, matching_files)
                    if latest_revs:
                        rev = next(iter(latest_revs.values()))
                        if rev != '-':  # Kun opdater hvis der faktisk er fundet en revision
                            sheet.Cells(row, headers["REV"]).Value = rev
                
                # Opdater drawing status
                drawing_status = get_drawing_status(part_number, matching_files)
                cell = sheet.Cells(row, headers["DRAWINGS"])
                cell.Value = drawing_status
                set_cell_color(cell, drawing_status)
        
        # 5. Beregn Total QTY
        progress_window.update("Beregner Total QTY...")
        # Indsæt Total QTY kolonne
        total_qty_col = headers["QTY"] + 1
        sheet.Columns(total_qty_col).Insert()
        sheet.Cells(1, total_qty_col).Value = "Total QTY"
        
        items = []
        qtys = []
        for row in range(2, sheet.UsedRange.Rows.Count + 1):
            items.append(str(sheet.Cells(row, headers["ITEM"]).Value))
            qtys.append(float(sheet.Cells(row, headers["QTY"]).Value or 0))
        
        total_qtys = []
        for idx, item in enumerate(items):
            if '.' not in str(item):
                total_qtys.append(qtys[idx])
            else:
                parent_item = '.'.join(str(item).split('.')[:-1])
                parent_idx = None
                for i, prev_item in enumerate(items[:idx]):
                    if str(prev_item) == parent_item:
                        parent_idx = i
                        break
                if parent_idx is not None:
                    total_qtys.append(qtys[idx] * total_qtys[parent_idx])
                else:
                    total_qtys.append(qtys[idx])
        
        # Indsæt Total QTY værdier
        for idx, total_qty in enumerate(total_qtys):
            sheet.Cells(idx + 2, total_qty_col).Value = total_qty
        
        # 6. Opret kategori faner
        progress_window.update("Opretter kategori faner...")
        # Opret en liste af part numbers fra den redigerede fil
        part_numbers = []
        for row in range(2, sheet.UsedRange.Rows.Count + 1):
            part_number = str(sheet.Cells(row, headers["PART NUMBER"]).Value).strip()
            if part_number:
                part_numbers.append(part_number)
        
        # Kategoriser baseret på de tilbageværende part numbers
        categorized_parts = {}
        
        # Først find piping grupper
        piping_groups = {}  # base_number -> piping_category
        for part_number in part_numbers:
            parts = part_number.split('-')
            if len(parts) >= 3:
                for part in parts:
                    if part.startswith('BM') or part[:2] in piping_categories:
                        prefix = 'BM' if part.startswith('BM') else part[:2]
                        target_category = piping_categories[prefix]
                        # Find base number fra item number
                        for row in range(2, sheet.UsedRange.Rows.Count + 1):
                            if str(sheet.Cells(row, headers["PART NUMBER"]).Value).strip() == part_number:
                                item_number = str(sheet.Cells(row, headers["ITEM"]).Value)
                                base_number = item_number.split('.')[0]
                                piping_groups[base_number] = target_category
                                break
        
        # Kategoriser items baseret på piping grupper og normale kategorier
        categorized_rows = {}  # category -> [(item_number, part_number, row)]
        for row in range(2, sheet.UsedRange.Rows.Count + 1):
            part_number = str(sheet.Cells(row, headers["PART NUMBER"]).Value).strip()
            item_number = str(sheet.Cells(row, headers["ITEM"]).Value)
            base_number = item_number.split('.')[0]
            
            # Tjek om denne item er del af en piping gruppe
            if base_number in piping_groups:
                target_category = piping_groups[base_number]
                if target_category not in categorized_rows:
                    categorized_rows[target_category] = []
                categorized_rows[target_category].append((item_number, part_number, row))
            
            # Hvis ikke piping, kategoriser normalt
            categorized = False
            for prefix, cat_name in categories.items():
                if part_number.startswith(prefix):
                    if cat_name not in categorized_rows:
                        categorized_rows[cat_name] = []
                    categorized_rows[cat_name].append((item_number, part_number, row))
                    categorized = True
                    break
            
            if not categorized:
                if "Other Items" not in categorized_rows:
                    categorized_rows["Other Items"] = []
                categorized_rows["Other Items"].append((item_number, part_number, row))
        
        # Opret faner for hver kategori
        for sheet_name, rows_to_copy in categorized_rows.items():
            if not rows_to_copy:  # Spring over tomme kategorier
                continue
                
            new_sheet = wb.Worksheets.Add(After=wb.Worksheets(wb.Worksheets.Count))
            new_sheet.Name = sheet_name
            
            # Kopier header
            sheet.Range("1:1").Copy(new_sheet.Range("1:1"))
            
            # Sorter rækker efter item nummer
            rows_to_copy.sort(key=lambda x: [float(n) if n.replace('.','').isdigit() else n 
                           for n in x[0].split('.')])
            
            # Kopier rækker i sorteret rækkefølge
            row_idx = 2
            for _, _, row_num in rows_to_copy:
                sheet.Range(f"{row_num}:{row_num}").Copy(
                    new_sheet.Range(f"{row_idx}:{row_idx}")
                )
                row_idx += 1
            
            # Formater det nye ark
            format_excel_sheet(new_sheet)
            
            # Gruppér rækker hvis det er en piping kategori
            if "Piping" in sheet_name:
                group_by_parent_items(new_sheet)
        
        # Formater BOM (Raw) arket (uden gruppering)
        format_excel_sheet(sheet)
        
        # Flyt BOM (Raw) forrest
        sheet.Move(Before=wb.Sheets(1))
        
        # Tjek og korriger formatering af række 2 i alle faner
        for sheet in wb.Sheets:
            if sheet.UsedRange.Rows.Count > 1:
                row2 = sheet.Range("2:2")
                row2.RowHeight = 91  # Sikrer at højden er 91 pixels
                row2.Font.Bold = False  # Sikrer at teksten ikke er fed
        
        wb.Save()
        return sheet, headers["DRAWINGS"], dest_path
        
    except Exception as e:
        # Luk Excel helt ned ved fejl
        if excel:
            try:
                excel.Quit()
            except:
                pass
        raise e
    
    finally:
        # Sikr at Excel lukkes ordenligt
        if wb:
            try:
                wb.Close(SaveChanges=False)
            except:
                pass
        if excel:
            try:
                excel.Application.Quit()
            except:
                pass
            finally:
                excel = None

def format_excel_sheet(sheet):
    """Formaterer et Excel ark med de specificerede indstillinger."""
    # Indstil kolonnebredder (konverter pixels til Excel enheder, ca. 7 pixels per enhed)
    column_widths = {
        "A": 52,   # Item
        "B": 152,  # Part Number
        "C": 47,   # Rev
        "D": 111,  # Description 1
        "E": 93,   # Description 2
        "F": 115,  # BOM Structure
        "G": 423,  # Description
        "H": 135,  # Material
        "I": 173,  # Standard/PED
        "J": 48,   # QTY
        "K": 82,   # Total QTY
        "L": 39,   # Weight
        "M": 39,   # Surface Area
        "N": 39,   # Volume
        "O": 200,  # Comment
        "P": 94,   # Drawings
    }
    
    for col_letter, width in column_widths.items():
        col = sheet.Range(f"{col_letter}:{col_letter}")
        col.ColumnWidth = width / 7.0  # Excel bruger ca. 7 pixels per enhed
    
    # Indstil rækkehøjder og formatering
    header_row = sheet.Range("1:1")
    header_row.RowHeight = 20  # Header højde: 20 pixels
    header_row.Font.Bold = True  # Gør kun header fed
    
    # Indstil højde for række 2 og resten af rækkerne
    if sheet.UsedRange.Rows.Count > 1:
        row2 = sheet.Range("2:2")
        row2.RowHeight = 91  # Række 2 højde: 91 pixels
        row2.Font.Bold = False  # Sikrer at række 2 ikke er fed
        
        if sheet.UsedRange.Rows.Count > 2:
            data_rows = sheet.Range(f"3:{sheet.UsedRange.Rows.Count}")
            data_rows.RowHeight = 91  # Data række højde: 91 pixels
    
    # Tilføj filter
    sheet.Range("1:1").AutoFilter()

def group_by_parent_items(sheet):
    """Grupperer rækker baseret på parent items med hierarkisk struktur."""
    last_row = sheet.UsedRange.Rows.Count
    if last_row <= 1:
        return
    
    # Opbyg hierarkisk struktur
    hierarchy = {}  # item_number -> [start_row, end_row, level]
    current_items = {}  # level -> item_number
    
    for row in range(2, last_row + 1):
        item_number = str(sheet.Cells(row, 1).Value)
        if not item_number:
            continue
        
        # Bestem niveau baseret på antal punktummer
        level = len(item_number.split('.'))
        
        # Find parent
        if level > 1:
            parent_number = '.'.join(item_number.split('.')[:-1])
            if parent_number in hierarchy:
                # Opdater parent's end_row
                hierarchy[parent_number][1] = row
        
        # Gem dette item
        hierarchy[item_number] = [row, row, level]
        current_items[level] = item_number
    
    # Opret grupper for hvert niveau, startende med det dybeste
    max_level = max(item[2] for item in hierarchy.values()) if hierarchy else 0
    
    for level in range(max_level, 1, -1):
        for item_number, (start_row, end_row, item_level) in hierarchy.items():
            if item_level == level - 1:  # Parent niveau
                # Find alle direkte children
                children = [child for child in hierarchy.keys() 
                          if child.startswith(item_number + '.') and 
                          len(child.split('.')) == level]
                
                if children:
                    # Find start og slut række for denne gruppe
                    group_start = min(hierarchy[child][0] for child in children)
                    group_end = max(hierarchy[child][1] for child in children)
                    
                    # Opret gruppe
                    range_to_group = sheet.Range(f"{group_start}:{group_end}")
                    range_to_group.Rows.Group()
                    
                    # Opdater parent's end_row hvis nødvendigt
                    hierarchy[item_number][1] = max(hierarchy[item_number][1], group_end)
    
    # Vis alle niveauer som standard
    sheet.Outline.ShowLevels(RowLevels=max_level)

def move_area_drawings_from_other_items(excel_copy_path, dest_path):
    """Flytter A-tegninger fra Other Items til Area Drawings."""
    if not os.path.exists(os.path.join(dest_path, "Other Items")):
        return
        
    # Find alle filer i Other Items mappen der matcher *-*-A* mønstret
    other_items_path = os.path.join(dest_path, "Other Items")
    area_drawings_path = os.path.join(dest_path, "Area Drawings")
    os.makedirs(area_drawings_path, exist_ok=True)
    
    # Find project og sub-project nummer fra filnavnet
    filename = os.path.basename(excel_copy_path)
    project_parts = filename.split('-')
    if len(project_parts) >= 2:
        project_number = project_parts[0]
        sub_project = project_parts[1]
        pattern = f"{project_number}-{sub_project}-A"
        # Opret den indskudte række med tomme værdier for alle kolonner
        inserted_row_data = [""] * 20  # Nok tomme værdier til alle kolonner
        # Sæt kun de specifikke værdier vi kender
        inserted_row_data[0] = "0"  # Item
        inserted_row_data[1] = f"{project_number}-{sub_project}-A01"  # Part Number
        inserted_row_data[6] = "Area Layout Drawing"  # Description (kolonne G)
        inserted_row_data[5] = "Inseparable"  # BOM Structure (kolonne F)
        inserted_row_data[9] = "1"  # QTY (kolonne J)
        inserted_row_data[10] = "1"  # Total QTY (kolonne K)
    else:
        return
    
    # Find og flyt filer
    matching_files = []
    for file in os.listdir(other_items_path):
        if file.startswith(pattern):
            source_file = os.path.join(other_items_path, file)
            dest_file = os.path.join(area_drawings_path, file)
            shutil.move(source_file, dest_file)
            matching_files.append(dest_file)
            print(f"Flytter {file} fra Other Items til Area Drawings")
    
    # Opdater Excel filen
    excel = None
    wb = None
    try:
        excel = Dispatch("Excel.Application")
        excel.DisplayAlerts = False
        excel.Visible = False
        wb = excel.Workbooks.Open(excel_copy_path)
        
        # Find Other Items og Area Drawings fanerne
        other_items_sheet = None
        area_drawings_sheet = None
        raw_sheet = None
        drawing_status = None  # Initialiser drawing_status variablen
        
        for sheet in wb.Sheets:
            if sheet.Name == "Other Items":
                other_items_sheet = sheet
            elif sheet.Name == "Area Drawings":
                area_drawings_sheet = sheet
            elif sheet.Name == "BOM (Raw)":
                raw_sheet = sheet
        
        if not other_items_sheet or not raw_sheet:
            return
            
        # Opret Area Drawings fanen hvis den ikke findes
        if not area_drawings_sheet:
            area_drawings_sheet = wb.Worksheets.Add(After=wb.Worksheets(wb.Worksheets.Count))
            area_drawings_sheet.Name = "Area Drawings"
            raw_sheet.Range("1:1").Copy(area_drawings_sheet.Range("1:1"))
        
        # Find rækker der skal flyttes
        rows_to_move = []
        for row in range(2, other_items_sheet.UsedRange.Rows.Count + 1):
            part_number = str(other_items_sheet.Cells(row, 2).Value)
            if part_number.startswith(pattern):
                rows_to_move.append(row)
        
        # Indsæt den indskudte række i Area Drawings
        area_drawings_sheet.Range("2:2").Insert(Shift=-4121)  # xlShiftDown
        for col, value in enumerate(inserted_row_data, start=1):
            if value:  # Kun sæt værdier der ikke er tomme
                area_drawings_sheet.Cells(2, col).Value = value
        
        # Opdater drawing status for den indskudte række hvis der er fundet filer
        if matching_files:
            drawing_status = get_drawing_status(inserted_row_data[1], matching_files)
            drawings_col = None
            for i in range(1, area_drawings_sheet.UsedRange.Columns.Count + 1):
                if str(area_drawings_sheet.Cells(1, i).Value).strip().upper() == "DRAWINGS":
                    drawings_col = i
                    break
            if drawings_col:
                cell = area_drawings_sheet.Cells(2, drawings_col)
                cell.Value = drawing_status
                set_cell_color(cell, drawing_status)
        
        # Kopier rækker til Area Drawings
        target_row = 3  # Start fra række 3, da række 2 er den indskudte række
        for row in rows_to_move:
            # Kopier til Area Drawings
            other_items_sheet.Range(f"{row}:{row}").Copy(area_drawings_sheet.Range(f"{target_row}:{target_row}"))
            target_row += 1
        
        # Slet rækker fra Other Items (fra bunden)
        for row in sorted(rows_to_move, reverse=True):
            other_items_sheet.Range(f"{row}:{row}").Delete()
        
        # Indsæt den indskudte række i BOM (Raw)
        raw_sheet.Range("2:2").Insert(Shift=-4121)  # xlShiftDown
        for col, value in enumerate(inserted_row_data, start=1):
            if value:  # Kun sæt værdier der ikke er tomme
                raw_sheet.Cells(2, col).Value = value
        
        # Opdater drawing status for den indskudte række i BOM (Raw)
        if matching_files and drawing_status:  # Tjek at både matching_files og drawing_status er sat
            drawings_col = None
            for i in range(1, raw_sheet.UsedRange.Columns.Count + 1):
                if str(raw_sheet.Cells(1, i).Value).strip().upper() == "DRAWINGS":
                    drawings_col = i
                    break
            if drawings_col:
                cell = raw_sheet.Cells(2, drawings_col)
                cell.Value = drawing_status
                set_cell_color(cell, drawing_status)
        
        # Opdater BOM (Raw) fanen
        raw_rows_to_move = []
        for row in range(3, raw_sheet.UsedRange.Rows.Count + 1):  # Start fra række 3
            part_number = str(raw_sheet.Cells(row, 2).Value)
            if part_number.startswith(pattern):
                raw_rows_to_move.append(row)
        
        # Flyt rækker til efter den indskudte række i BOM (Raw)
        if raw_rows_to_move:
            # Gem rækkerne midlertidigt
            temp_range = raw_sheet.Range(f"{raw_rows_to_move[0]}:{raw_rows_to_move[-1]}")
            temp_range.Copy()
            
            # Indsæt efter den indskudte række
            raw_sheet.Range("3:3").Insert(Shift=-4121)  # xlShiftDown
            
            # Slet de originale rækker
            for row in sorted(raw_rows_to_move, reverse=True):
                raw_sheet.Range(f"{row+1}:{row+1}").Delete()
        
        # Tjek og korriger formatering af række 2 i alle faner
        for sheet in wb.Sheets:
            if sheet.UsedRange.Rows.Count > 1:
                row2 = sheet.Range("2:2")
                row2.RowHeight = 91  # Sikrer at højden er 91 pixels
                row2.Font.Bold = False  # Sikrer at teksten ikke er fed
        
        wb.Save()
        
    finally:
        if wb:
            try:
                wb.Close(SaveChanges=False)
            except:
                pass
        if excel:
            try:
                excel.Quit()
            except:
                pass

def main(excel_path, prev_excel_path=None, include_equipment=True, 
         find_rev_before_date=False, rev_date=None, include_data_sheet=False,
         progress_window=None):
    """Hovedfunktion der implementerer den nye rækkefølge."""
    try:
        # TRIN 1: GUI opstart
        if progress_window:
            progress_window.update("Initialiserer program...")
        
        # TRIN 2: Data Indlæsning og Validering
        if progress_window:
            progress_window.update("Indlæser og validerer data...")
        
        # Behandl Excel filen
        sheet, headers, wb = process_excel(excel_path, progress_window)
        
        # TRIN 3: Kategorisere "Part number"
        if progress_window:
            progress_window.update("Kategoriserer part numbers...")
        
        # Indlæs kategori konfiguration og kategoriser data
        category_config = load_category_config()
        categorized_data = categorize_data(sheet, excel_path)
        
        # TRIN 4: Oprettelse Partlist
        if progress_window:
            progress_window.update("Opretter partlist...")
        
        # Opret partlist fane
        create_partlist_sheet(wb, sheet, headers)
        
        # TRIN 5: Kopiere filer
        if progress_window:
            progress_window.update("Kopierer filer...")
        
        # Opret destination og kopier filer
        base_dest_dir = os.path.join(os.path.dirname(excel_path), "Drawing Package")
        os.makedirs(base_dest_dir, exist_ok=True)
        copy_files_to_destination(base_dest_dir, categorized_data)
        
        # TRIN 6: Compare
        if prev_excel_path:
            if progress_window:
                progress_window.update("Sammenligner med tidligere tegning-pakke...")
            
            # Sammenlign og gem rapport
            compare_and_save_report(excel_path, prev_excel_path, base_dest_dir)
        
        # Vis success besked
        parent = progress_window.window.master if progress_window else None
        show_success_message(f"Tegning-pakke er blevet oprettet i:\n{base_dest_dir}", parent)
        
    except Exception as e:
        print(f"Fejl under behandling: {str(e)}")
        raise
    finally:
        if 'wb' in locals():
            wb.Close(SaveChanges=False)
        if 'excel' in locals():
            excel.Quit()

def has_pdf(part_number):
    """Tjekker om der findes en PDF fil for et part number i netværksmappen."""
    try:
        # Scan netværksmappen for filer
        files = scan_directory_concurrent(NETWORK_PATH)
        
        # Find matchende PDF filer
        matching_files = [f for f in files if os.path.basename(f).startswith(part_number) 
                         and f.endswith('.pdf')]
        return len(matching_files) > 0
    except Exception as e:
        print(f"Fejl ved søgning efter PDF for {part_number}: {str(e)}")
        return False

def has_dwg(part_number):
    """Tjekker om der findes en DWG fil for et part number i netværksmappen."""
    try:
        # Scan netværksmappen for filer
        files = scan_directory_concurrent(NETWORK_PATH)
        
        # Find matchende DWG filer
        matching_files = [f for f in files if os.path.basename(f).startswith(part_number) 
                         and f.endswith('.dwg')]
        return len(matching_files) > 0
    except Exception as e:
        print(f"Fejl ved søgning efter DWG for {part_number}: {str(e)}")
        return False

def find_latest_revision(part_number, current_rev):
    """Finder den seneste revision for et part number."""
    return current_rev  # TODO: Implementer revision tjek

def find_files_for_part(part_number, rev):
    """Finder filer der matcher et part number og revision."""
    return []  # TODO: Implementer fil søgning

def compare_drawing_packages(current_df, prev_df):
    """Sammenligner to tegning-pakker og returnerer en liste af ændringer."""
    changes = []
    
    # Konverter til pandas DataFrames hvis de ikke allerede er det
    if not isinstance(current_df, pd.DataFrame):
        current_df = pd.read_excel(current_df)
    if not isinstance(prev_df, pd.DataFrame):
        prev_df = pd.read_excel(prev_df)
    
    # Find fælles kolonner
    common_cols = list(set(current_df.columns) & set(prev_df.columns))
    
    # Sammenlign hver række i den nuværende fil med den tidligere
    for _, current_row in current_df.iterrows():
        part_number = str(current_row['PART NUMBER']).strip()
        
        # Find tilsvarende række i den tidligere fil
        prev_row = prev_df[prev_df['PART NUMBER'] == part_number]
        
        if prev_row.empty:
            # Ny del - tilføjet
            changes.append(f"NY DEL: {part_number} - {current_row['DESCRIPTION']}")
            continue
        
        # Sammenlign værdier
        prev_row = prev_row.iloc[0]
        for col in common_cols:
            if col in ['PART NUMBER', 'DESCRIPTION']:  # Fokusér på vigtige kolonner
                current_val = str(current_row[col]).strip()
                prev_val = str(prev_row[col]).strip()
                if current_val != prev_val:
                    changes.append(f"ÆNDRET {col} for {part_number}:")
                    changes.append(f"  Gammel: {prev_val}")
                    changes.append(f"  Ny: {current_val}")
    
    # Find slettede dele
    for _, prev_row in prev_df.iterrows():
        part_number = str(prev_row['PART NUMBER']).strip()
        if part_number not in current_df['PART NUMBER'].values:
            changes.append(f"SLETTET DEL: {part_number} - {prev_row['DESCRIPTION']}")
    
    return changes

def compare_and_save_report(current_excel, prev_excel, dest_dir):
    """Sammenligner to tegning-pakker og gemmer en rapport."""
    try:
        # Sammenlign tegning-pakkerne
        changes = compare_drawing_packages(current_excel, prev_excel)
        
        # Gem rapport
        report_path = os.path.join(dest_dir, "changes_report.txt")
        with open(report_path, 'w', encoding='utf-8') as f:
            f.write("Ændringer i tegning-pakke:\n\n")
            if changes:
                for change in changes:
                    f.write(f"{change}\n")
            else:
                f.write("Ingen ændringer fundet.\n")
        
        return report_path
    
    except Exception as e:
        print(f"Fejl under sammenligning: {str(e)}")
        raise

if __name__ == "__main__":
    # Start GUI
    app = MainWindow()
    app.run()
