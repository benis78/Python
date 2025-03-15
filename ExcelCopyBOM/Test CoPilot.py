import os
import shutil
import pandas as pd
import openpyxl
import time
import sys
import subprocess
from openpyxl.utils import get_column_letter
from openpyxl.styles import PatternFill, Font
from tkinter import Tk, filedialog, messagebox, ttk, Toplevel
from win32com.client import Dispatch
from concurrent.futures import ThreadPoolExecutor
import pythoncom
import concurrent.futures
from pathlib import Path

# Opsæt logging til debug.txt
debug_file = os.path.join(os.path.dirname(__file__), "debug.txt")
sys.stdout = open(debug_file, 'w', encoding='utf-8')

# Cache for kategorier og piping matches
category_cache = {}
piping_cache = {}

# Global piping kategorier
piping_categories = {}

class ProgressWindow:
    def __init__(self, title="Behandler filer..."):
        self.root = Toplevel()
        self.root.title(title)
        self.root.attributes('-topmost', 1)
        
        # Centrér vinduet
        window_width = 300
        window_height = 100
        screen_width = self.root.winfo_screenwidth()
        screen_height = self.root.winfo_screenheight()
        x = (screen_width - window_width) // 2
        y = (screen_height - window_height) // 2
        self.root.geometry(f"{window_width}x{window_height}+{x}+{y}")
        
        # Tilføj labels og progressbar
        self.label = ttk.Label(self.root, text="Starter...")
        self.label.pack(pady=10)
        
        self.progress = ttk.Progressbar(self.root, length=250, mode='determinate')
        self.progress.pack(pady=10)
        
        self.root.protocol("WM_DELETE_WINDOW", lambda: None)  # Deaktiver luk-knappen
        self.root.update()
    
    def update_progress(self, value, text=None):
        if text:
            self.label.config(text=text)
        self.progress['value'] = value
        self.root.update()
    
    def close(self):
        self.root.destroy()

def show_success_message(message):
    """Viser en success besked i et Toplevel vindue der altid er øverst."""
    top = Toplevel()
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
    
    ttk.Label(top, text=message, wraplength=350, justify='center').pack(pady=20)
    ttk.Button(top, text="OK", command=top.destroy).pack(pady=10)

def choose_file():
    """Lader brugeren vælge en Excel-fil og returnerer dens sti."""
    file_path = "C:\\Coding\\Python\\ExcelCopyBOM\\4003-615-A01-E - BOM.xlsx"
    #file_path = filedialog.askopenfilename(
    #    title="Vælg Excel BOM fil",
    #    initialdir=r'C:\Working Folder\Designs\5-Projects',
    #    filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")]
    #)
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

def load_categories(file_path):
    """Indlæser kategorier fra en ekstern .txt-fil og erstatter * med projekt nummeret."""
    categories = {}
    categories_file = os.path.join(os.path.dirname(__file__), "categories.txt")
    if not os.path.exists(categories_file):
        messagebox.showerror("ERROR", f"Categories file not found: {categories_file}")
        return categories
    
    # Hent projekt nummeret og sub-projekt nummeret fra Excel-filens navn
    filename = os.path.basename(file_path)
    project_parts = filename.split('-')
    if len(project_parts) >= 2:
        project_number = project_parts[0]  # f.eks. "4003"
        sub_project = project_parts[1]     # f.eks. "02.1"
    else:
        project_number = filename[:4]
        sub_project = ""
    
    with open(categories_file, "r") as file:
        for line in file:
            if '=' not in line:
                continue
                
            parts = line.strip().split("=")
            if len(parts) == 2:
                key = parts[0].strip()
                category_name = parts[1].strip()
                
                # Håndter forskellige wildcard mønstre
                if "*-*-" in key:
                    # For mønstre som "*-*-BM" eller "*-*-A"
                    if sub_project:
                        base_key = key.replace("*-*-", f"{project_number}-{sub_project}-")
                    else:
                        base_key = key.replace("*-*-", f"{project_number}-")
                    categories[base_key] = category_name
                elif key.startswith("*-"):
                    # For mønstre som "*-610"
                    base_key = key.replace("*-", f"{project_number}-")
                    categories[base_key] = category_name
                else:
                    # For almindelige mønstre som "0000-703"
                    categories[key] = category_name
    
    return categories

def load_piping_categories():
    """Indlæser piping kategorier fra piping_categories.txt."""
    piping_categories = {}
    piping_file = os.path.join(os.path.dirname(__file__), "piping_categories.txt")
    
    print("\n=== DEBUG: INDLÆSER PIPING KATEGORIER ===")
    if not os.path.exists(piping_file):
        print(f"ADVARSEL: Piping categories fil ikke fundet: {piping_file}")
        return piping_categories
    
    print(f"Læser piping kategorier fra: {piping_file}")
    with open(piping_file, "r") as file:
        for line in file:
            if '=' not in line:
                continue
            prefix, category = line.strip().split("=")
            prefix = prefix.strip()
            category = category.strip()
            piping_categories[prefix] = category
            print(f"Tilføjet piping kategori: {prefix} -> {category}")
    
    print(f"Indlæst {len(piping_categories)} piping kategorier")
    return piping_categories

def initialize_categories():
    """Initialiserer globale kategori variabler."""
    global piping_categories
    piping_categories = load_piping_categories()

def is_piping_item(part_number):
    """Hurtig tjek om et part number er en piping item."""
    global piping_categories
    
    # Tjek cache først
    if part_number in piping_cache:
        return piping_cache[part_number]
    
    # Hurtig pre-check for kendte ikke-piping prefixes
    if str(part_number).startswith('0000-'):
        result = (False, None)
        piping_cache[part_number] = result
        return result
    
    # Tjek for piping kategorier med nyt format (f.eks. BM001 eller BMxx1)
    parts = str(part_number).split('-')
    for part in parts:
        # Find alle to-bogstavs koder i piping_categories
        for prefix in piping_categories.keys():
            # Tjek om delen starter med prefix og følges af enten 'xx' eller tal
            if (part.startswith(prefix) and 
                len(part) >= 5 and  # Minimum længde for f.eks. BMxx1
                (part[2:4] == 'xx' or part[2:4].isdigit()) and
                part[4].isdigit()):
                
                result = (True, piping_categories[prefix])
                piping_cache[part_number] = result
                return result
    
    result = (False, None)
    piping_cache[part_number] = result
    return result

def find_category(part_number, categories):
    """Find den bedste matchende kategori for et part number."""
    try:
        # Tjek først for BioMix
        if any(pattern in part_number for pattern in ["-BM-", "-BM", "630"]):
            return "BioMix"
        
        # Find den længste matchende prefix
        matching_category = None
        matching_prefix = None
        
        for pattern, category in categories.items():
            prefix = pattern.replace("*-*-", "").replace("*-", "")
            if part_number.startswith(prefix):
                if matching_prefix is None or len(prefix) > len(matching_prefix):
                    matching_prefix = prefix
                    matching_category = category
        
        return matching_category if matching_category else "Other Items"
        
    except Exception as e:
        print(f"Fejl under kategorisøgning for {part_number}: {str(e)}")
        return "Other Items"

def categorize_data(df, file_path):
    """Optimeret kategorisering af data med piping først, derefter resten."""
    try:
        print("\n=== DEBUG: KATEGORISERING START ===")
        print(f"Antal rækker i dataframe: {len(df)}")
        print(f"DataFrame kolonner: {', '.join(df.columns)}")
        
        # Opret case-insensitive mapping af kolonnenavne
        column_mapping = {}
        for col in df.columns:
            if col.upper() == 'ITEM':
                column_mapping[col] = 'ITEM'
            elif col.upper() == 'PART NUMBER':
                column_mapping[col] = 'PART NUMBER'
        
        # Omdøb kolonner til standardiserede navne
        df = df.rename(columns=column_mapping)
        print(f"Kolonner efter omdøbning: {', '.join(df.columns)}")
        
        # Verificer at nødvendige kolonner findes
        required_cols = ['ITEM', 'PART NUMBER']
        missing_cols = [col for col in required_cols if col not in df.columns]
        if missing_cols:
            raise ValueError(f"Manglende påkrævede kolonner: {', '.join(missing_cols)}")
        
        # Indlæs kategorier
        piping_categories = load_piping_categories()
        print(f"Indlæste piping kategorier: {', '.join(piping_categories.values())}")
        
        regular_categories = load_categories(file_path)
        print(f"Indlæste almindelige kategorier: {', '.join(regular_categories.values())}")
        
        # Opret base_number kolonne for gruppering
        print("Opretter BASE_NUMBER kolonne...")
        df['BASE_NUMBER'] = df['ITEM'].astype(str).str.split('.').str[0]
        print(f"Unikke base numbers: {len(df['BASE_NUMBER'].unique())}")
        
        categorized_data = {}
        processed_part_numbers = set()  # Hold styr på behandlede part numbers
        
        # TRIN 1: Kategoriser piping først
        print("\n=== KATEGORISERER PIPING ===")
        for base_number, group in df.groupby('BASE_NUMBER'):
            try:
                parent_mask = group['ITEM'].astype(str) == str(base_number)
                parent_row = group[parent_mask].iloc[0] if parent_mask.any() else group.iloc[0]
                part_number = str(parent_row['PART NUMBER'])
                
                is_piping, piping_category = is_piping_item(part_number)
                if is_piping and piping_category:
                    if piping_category not in categorized_data:
                        categorized_data[piping_category] = pd.DataFrame(columns=df.columns)
                    categorized_data[piping_category] = pd.concat([categorized_data[piping_category], group], ignore_index=True)
                    processed_part_numbers.update(group['PART NUMBER'].astype(str).tolist())
                    print(f"Piping: {part_number} -> {piping_category}")
            except Exception as e:
                print(f"Fejl under behandling af piping base_number {base_number}: {str(e)}")
                continue
        
        # TRIN 2: Kategoriser resten (undtagen allerede kategoriserede rør)
        print("\n=== KATEGORISERER ØVRIGE KOMPONENTER ===")
        remaining_mask = ~df['PART NUMBER'].astype(str).isin(processed_part_numbers)
        remaining_df = df[remaining_mask]
        
        for base_number, group in remaining_df.groupby('BASE_NUMBER'):
            try:
                parent_mask = group['ITEM'].astype(str) == str(base_number)
                parent_row = group[parent_mask].iloc[0] if parent_mask.any() else group.iloc[0]
                part_number = str(parent_row['PART NUMBER'])
                
                # Find normal kategori
                category = find_category(part_number, regular_categories)
                if category not in categorized_data:
                    categorized_data[category] = pd.DataFrame(columns=df.columns)
                categorized_data[category] = pd.concat([categorized_data[category], group], ignore_index=True)
                print(f"Normal: {part_number} -> {category}")
            except Exception as e:
                print(f"Fejl under behandling af normal base_number {base_number}: {str(e)}")
                continue
        
        # Log resultater
        print("\n=== KATEGORISERING AFSLUTTET ===")
        for category, data in categorized_data.items():
            if len(data) > 0:
                print(f"{category}: {len(data)} items")
        
        return categorized_data
        
    except Exception as e:
        print(f"Fejl under kategorisering: {str(e)}")
        print(f"DataFrame info:")
        print(df.info())
        raise

def should_include_excel_row(item_number, sheet):
    """Excel version af should_include_row funktionen."""
    if not item_number or '.' not in str(item_number):
        return True
    
    # Find den aktuelle række
    current_range = sheet.Range("A:A").Find(item_number)
    if current_range:
        current_row = current_range.Row
        current_structure = str(sheet.Cells(current_row, 5).Value).strip().upper()  # Kolonne E (BOM Structure)
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
    progress_window.update_progress(10, "Læser data fra Excel...")
    
    try:
        print("\n=== DEBUG: LÆSER EXCEL FIL ===")
        print(f"Fil: {file_path}")
        
        # Læs først kolonne navne for at identificere de korrekte kolonner
        df_headers = pd.read_excel(file_path, nrows=0)
        print(f"Fundne kolonner: {', '.join(df_headers.columns)}")
        
        # Læs data med optimerede datatyper
        print("Læser Excel data med optimerede datatyper...")
        df = pd.read_excel(
            file_path,
            dtype={
                'Item': str,
                'Part Number': str,
                'REV': 'category',
                'BOM Structure': 'category',
                'Description': str,
                'QTY': float
            }
        )
        print(f"Data indlæst: {len(df)} rækker")
        
        # Kategoriser data
        print("Starter kategorisering...")
        categorized_data = categorize_data(df, file_path)
        
        return categorized_data, df['TOTAL QTY'].tolist(), df
        
    except Exception as e:
        print(f"Fejl under Excel behandling: {str(e)}")
        print("DataFrame info:")
        if 'df' in locals():
            print(df.info())
        raise

def calculate_total_qty(row, df):
    """Beregner Total QTY for en række ved at bruge vektoriserede operationer."""
    item = str(row['ITEM'])
    if '.' not in item:
        return float(row['QTY'])
    
    parent_item = '.'.join(item.split('.')[:-1])
    parent_row = df[df['ITEM'] == parent_item]
    
    if not parent_row.empty:
        return float(row['QTY']) * float(parent_row.iloc[0]['TOTAL QTY'])
    return float(row['QTY'])

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
    if not files:
        return {}, {}
    
    latest_files = {}
    latest_revs = {}
    
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
    try:
        excel = Dispatch("Excel.Application")
        excel.DisplayAlerts = False
        excel.Visible = False
        
        wb = excel.Workbooks.Open(excel_path)
        sheet = wb.Sheets(1)
        sheet.Name = "BOM (Raw)"
        
        # Find nødvendige kolonner
        headers = {}
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
        rows_to_delete = []
        for row in range(2, sheet.UsedRange.Rows.Count + 1):
            item_number = str(sheet.Cells(row, 1).Value)
            if not should_include_excel_row(item_number, sheet):
                rows_to_delete.append(row)
                continue
            
            # Opdater række information
            update_row_information(sheet, row, headers, files)
        
        # Slet uønskede rækker (fra bunden)
        for row in sorted(rows_to_delete, reverse=True):
            sheet.Rows(row).Delete()
        
        # Opret kategori faner
        print("\n=== OPRETTER KATEGORI FANER ===")
        for sheet_name, data in categorized_data.items():
            if len(data) == 0:  # Spring over tomme kategorier
                print(f"Springer over tom kategori: {sheet_name}")
                continue
                
            print(f"\nOpretter fane: {sheet_name} med {len(data)} rækker")
            
            # Find eller opret fanen
            target_sheet = None
            for s in wb.Sheets:
                if s.Name == sheet_name:
                    target_sheet = s
                    print(f"Fane {sheet_name} findes allerede")
                    break
            
            if target_sheet is None:
                target_sheet = wb.Worksheets.Add(After=wb.Worksheets(wb.Worksheets.Count))
                target_sheet.Name = sheet_name
                print(f"Oprettet ny fane: {sheet_name}")
                # Kopier header
                sheet.Range("1:1").Copy(target_sheet.Range("1:1"))
            
            # Indsæt den indskudte linje hvis det er en piping kategori
            if "Piping" in sheet_name or sheet_name == "BioMix":
                # Indsæt den indskudte linje
                target_sheet.Range("2:2").Insert(Shift=-4121)  # xlShiftDown
                # Kopier første række fra data som den indskudte linje
                if len(data) > 0:
                    first_row = data.iloc[0]
                    target_sheet.Cells(2, headers["ITEM"]).Value = first_row["ITEM"].split('.')[0]
                    target_sheet.Cells(2, headers["PART NUMBER"]).Value = first_row["PART NUMBER"]
                    if "Description" in data.columns:
                        target_sheet.Cells(2, headers["Description"]).Value = f"{sheet_name} Assembly"
                    if "BOM STRUCTURE" in headers:
                        target_sheet.Cells(2, headers["BOM STRUCTURE"]).Value = "Inseparable"
                    if "QTY" in headers:
                        target_sheet.Cells(2, headers["QTY"]).Value = 1
                    if "TOTAL QTY" in headers:
                        target_sheet.Cells(2, headers["TOTAL QTY"]).Value = 1
            
            print("Kopierer rækker:")
            start_row = 3 if ("Piping" in sheet_name or sheet_name == "BioMix") else 2
            for idx, row in data.iterrows():
                part_number = str(row["PART NUMBER"])
                if part_number and part_number.strip():
                    found_range = sheet.Range("B:B").Find(part_number)
                    if found_range:
                        row_num = found_range.Row
                        print(f"  Kopierer {part_number} fra række {row_num} til {idx+start_row}")
                        sheet.Range(f"{row_num}:{row_num}").Copy(
                            target_sheet.Range(f"{idx+start_row}:{idx+start_row}")
                        )
            
            # Formater og gruppér
            format_excel_sheet(target_sheet)
            if "Piping" in sheet_name or sheet_name == "BioMix":
                group_by_parent_items(target_sheet)
        
        # Flyt BOM (Raw) forrest og gem
        sheet.Move(Before=wb.Sheets(1))
        wb.Save()
        return sheet
    
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
    part_number = str(sheet.Cells(row, 2).Value)  # Kolonne B (Part Number)
    if not part_number.strip():
        return
    
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
            sheet.Cells(row, headers["REV"]).Value = rev
    
    # Opdater drawing status
    if "DRAWINGS" in headers:
        drawing_status = get_drawing_status(part_number, matching_files)
        cell = sheet.Cells(row, headers["DRAWINGS"])
        cell.Value = drawing_status
        set_cell_color(cell, drawing_status)

def scan_files_parallel(source_path):
    """Scanner filer parallelt med concurrent.futures."""
    files = []
    with concurrent.futures.ThreadPoolExecutor() as executor:
        futures = []
        for entry in os.scandir(source_path):
            if entry.is_dir():
                futures.append(executor.submit(scan_directory_task, entry))
            elif entry.is_file() and (entry.name.endswith('.pdf') or entry.name.endswith('.dwg')):
                files.append(entry.path)
        
        for future in concurrent.futures.as_completed(futures):
            try:
                result = future.result()
                files.extend(result)
            except Exception as e:
                print(f"Fejl under filscanning: {str(e)}")
    
    return files

def copy_files_to_destination(dest_path, files, top_assembly, categories_with_files):
    """Kopierer filer til destinationsmapper med parallel processing."""
    try:
        # Korriger destinationsstien til at være mappen og ikke Excel-filen
        base_dest_path = os.path.dirname(dest_path)
        print(f"\nKopierer filer til: {base_dest_path}")
        print(f"Top assembly number: {top_assembly}")
        
        # Opret destinationsmapper
        for category in categories_with_files.keys():
            category_path = os.path.join(base_dest_path, category)
            os.makedirs(category_path, exist_ok=True)
            print(f"Oprettet mappe: {category_path}")
        
        # Gruppér filer efter deres kategori
        files_by_category = {}
        for category, data in categories_with_files.items():
            files_by_category[category] = []
            for _, row in data.iterrows():
                part_number = str(row['PART NUMBER'])
                if not part_number:
                    continue
                
                matching_files = [f for f in files if part_number in os.path.basename(f)]
                files_by_category[category].extend(matching_files)
        
        # Kopier filer parallelt
        with concurrent.futures.ThreadPoolExecutor() as executor:
            futures = []
            for category, file_list in files_by_category.items():
                category_path = os.path.join(base_dest_path, category)
                for file_path in file_list:
                    dest_file = os.path.join(category_path, os.path.basename(file_path))
                    futures.append(
                        executor.submit(safe_copy_file, file_path, dest_file)
                    )
            
            # Vent på at alle kopieringer er færdige
            for future in concurrent.futures.as_completed(futures):
                try:
                    future.result()
                except Exception as e:
                    print(f"Fejl under filkopiering: {str(e)}")
        
        print("Filkopiering afsluttet")
        
    except Exception as e:
        print(f"Fejl i copy_files_to_destination: {str(e)}")
        raise

def safe_copy_file(src, dst):
    """Sikker kopiering af en enkelt fil med fejlhåndtering."""
    try:
        import shutil
        shutil.copy2(src, dst)
        print(f"Kopieret: {os.path.basename(src)} -> {os.path.dirname(dst)}")
    except Exception as e:
        print(f"Kunne ikke kopiere {src}: {str(e)}")
        raise

def kill_excel(progress_window):
    """Lukker alle Excel processer med advarsel."""
    dialog = Toplevel(progress_window.root)
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
    dialog.transient(progress_window.root)
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
    start_time = time.time()
    dest_path = None
    
    try:
        # Sikr at alle Excel processer er lukket
        if not kill_excel(progress_window):
            raise Exception("Processen blev annulleret af brugeren")
        
        # 1. Opret kopi af Excel filen
        progress_window.update_progress(5, "Opretter kopi af Excel fil...")
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
        
        # 2. Indsæt linje i række 2
        progress_window.update_progress(10, "Indsætter linje i række 2...")
        sheet.Rows(2).Insert()
        sheet.Cells(2, 1).Value = "0"
        sheet.Cells(2, 2).Value = f"{base_name}-A01"
        sheet.Cells(2, 6).Value = "Inseparable"
        sheet.Cells(2, 7).Value = "Area Layout Drawing"
        sheet.Cells(2, 10).Value = 1
        sheet.Cells(2, 11).Value = 1
        
        # 3. Identificer kolonne numre
        progress_window.update_progress(15, "Identificerer kolonner...")
        headers = {}
        for i in range(1, sheet.UsedRange.Columns.Count + 1):
            header = str(sheet.Cells(1, i).Value).strip().upper()
            if header in ["ITEM", "PART NUMBER", "REV", "BOM STRUCTURE", "QTY", "DRAWINGS", "TOTAL QTY", "DESCRIPTION"]:
                headers[header] = i
            last_column = i
        
        if not all(key in headers for key in ["ITEM", "PART NUMBER", "BOM STRUCTURE", "QTY", "DESCRIPTION"]):
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
        
        # 4. Håndter "Inseparable" rækker
        progress_window.update_progress(20, "Håndterer Inseparable rækker...")
        rows_to_delete = []
        for row in range(2, sheet.UsedRange.Rows.Count + 1):
            item_number = str(sheet.Cells(row, headers["ITEM"]).Value)
            bom_structure = str(sheet.Cells(row, headers["BOM STRUCTURE"]).Value).strip().upper()
            
            if bom_structure == "INSEPARABLE":
                rows_to_delete.append(row)
        
        # Slet markerede rækker (fra bunden)
        for row in sorted(rows_to_delete, reverse=True):
            sheet.Rows(row).Delete()
        
        # 5. Håndter revisioner og opdater drawing status
        progress_window.update_progress(25, "Håndterer revisioner...")
        for row in range(2, sheet.UsedRange.Rows.Count + 1):
            part_number = str(sheet.Cells(row, headers["PART NUMBER"]).Value)
            if part_number.strip():
                update_row_information(sheet, row, headers, files)
        
        # 6. Beregn Total QTY
        progress_window.update_progress(30, "Beregner Total QTY...")
        # Indsæt Total QTY kolonne hvis den ikke findes
        if "TOTAL QTY" not in headers:
            total_qty_col = headers["QTY"] + 1
            sheet.Columns(total_qty_col).Insert()
            sheet.Cells(1, total_qty_col).Value = "Total QTY"
            headers["TOTAL QTY"] = total_qty_col
        
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
                    if prev_item == parent_item:
                        parent_idx = i
                        break
                if parent_idx is not None:
                    total_qtys.append(qtys[idx] * total_qtys[parent_idx])
                else:
                    total_qtys.append(qtys[idx])
        
        # Indsæt Total QTY værdier
        for idx, total_qty in enumerate(total_qtys):
            sheet.Cells(idx + 2, headers["TOTAL QTY"]).Value = total_qty
        
        # 7. Fjern Inseparable children linjer
        progress_window.update_progress(35, "Fjerner Inseparable children linjer...")
        rows_to_delete = []
        for row in range(2, sheet.UsedRange.Rows.Count + 1):
            item_number = str(sheet.Cells(row, headers["ITEM"]).Value)
            if item_number.startswith("0."):
                rows_to_delete.append(row)
        
        for row in sorted(rows_to_delete, reverse=True):
            sheet.Rows(row).Delete()
        
        # 8. Scanner for piping items
        progress_window.update_progress(40, "Scanner for piping items...")
        
        # Load Excel data into pandas DataFrame
        df = pd.read_excel(dest_path)
        
        # Ensure all columns are strings
        for col in df.columns:
            df[col] = df[col].astype(str)
        
        # Define categorize_part function
        def categorize_part(row):
            part_number = row['PART NUMBER']
            
            # Check for Piping Categories
            if re.match(r'BM\d{2}', part_number):
                return 'Biomass Piping'
            elif part_number.startswith('630'):
                return 'BioMix Piping'
            
            # Check for other categories
            for category_code, category_name in categories.items():
                if part_number.startswith(category_code):
                    return category_name
            
            return 'Other Items'
        
        # Load categories from categories.txt
        categories = {}
        with open(os.path.join(os.path.dirname(__file__), 'categories.txt'), 'r', encoding='utf-8') as f:
            for line in f:
                if '=' in line and not line.startswith('#'):
                    code, name = line.strip().split('=', 1)
                    categories[code] = name
        
        # Apply categorization
        df['Category'] = df.apply(categorize_part, axis=1)
        
        # Group data by category
        categorized_data = df.groupby('Category')
        
        # 9. Scanner for andre part numbers (ikke nødvendigt her, da det er en del af kategoriseringen)
        progress_window.update_progress(45, "Scanner for andre part numbers...")
        
        # 10. Formater Excel ark
        progress_window.update_progress(50, "Formaterer Excel ark...")
        format_excel_sheet(sheet)
        group_by_parent_items(sheet)
        
        # 11. Opret kategori faner
        progress_window.update_progress(55, "Opretter kategori faner...")
        
        # Create category sheets
        for category, data in categorized_data:
            if category == "nan":
                print(f"Skipping category 'nan'")
                continue
            
            print(f"Creating sheet for category: {category}")
            
            # Create new sheet
            try:
                new_sheet = wb.Worksheets.Add(After=wb.Worksheets(wb.Worksheets.Count))
                new_sheet.Name = category
            except Exception as e:
                print(f"Error creating sheet {category}: {e}")
                continue
            
            # Copy headers
            sheet.Range("1:1").Copy(new_sheet.Range("1:1"))
            
            # Write data
            start_row = 2
            for _, row_data in data.iterrows():
                try:
                    for col, value in enumerate(row_data.values, start=1):
                        new_sheet.Cells(start_row, col).Value = value
                    start_row += 1
                except Exception as e:
                    print(f"Error writing data to sheet {category}: {e}")
                    continue
            
            # Format the new sheet
            format_excel_sheet(new_sheet)
            group_by_parent_items(new_sheet)
        
        # 12. Opret mapper
        progress_window.update_progress(60, "Opretter mapper...")
        
        # Create directories for each category
        categories_with_files = {}
        for category in df['Category'].unique():
            if category == "nan":
                continue
            
            category_path = os.path.join(dest_folder, category)
            os.makedirs(category_path, exist_ok=True)
            categories_with_files[category] = category_path
        
        # 13. Kopier filer efter kategorisering
        progress_window.update_progress(65, "Kopierer filer efter kategorisering...")
        
        # Copy files to destination folders
        for row in range(2, sheet.UsedRange.Rows.Count + 1):
            part_number = str(sheet.Cells(row, headers["PART NUMBER"]).Value)
            drawings = str(sheet.Cells(row, headers["DRAWINGS"]).Value)
            
            # Find category for this row
            category = df[df['PART NUMBER'] == part_number]['Category'].iloc[0]
            
            if category == "nan":
                continue
            
            # Get destination folder
            dest_folder_path = categories_with_files.get(category)
            
            if not dest_folder_path:
                print(f"No destination folder found for category: {category}")
                continue
            
            # Copy files
            if drawings:
                for drawing in drawings.split(';'):
                    drawing = drawing.strip()
                    
                    # Find the file
                    source_file = None
                    for file in files:
                        if drawing in file:
                            source_file = file
                            break
                    
                    if source_file:
                        try:
                            dest_file = os.path.join(dest_folder_path, os.path.basename(source_file))
                            shutil.copy2(source_file, dest_file)
                        except Exception as e:
                            print(f"Error copying file {source_file} to {dest_file}: {e}")
        
        # Flyt A-tegninger fra Other Items til Area Drawings hvis nødvendigt
        progress_window.update_progress(70, "Kontrollerer for fejlplacerede Area Drawings...")
        move_area_drawings_from_other_items(dest_path, dest_path)
        
        # Gem ændringer
        wb.Save()
        
        duration = time.time() - start_time
        progress_window.close()
        show_success_message(f"BOM processing complete!\nFiles saved in {dest_path}\nTime taken: {duration:.2f} seconds")
        
    except Exception as e:
        if progress_window:
            progress_window.close()
        messagebox.showerror("Error", str(e))
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
    
    return dest_path

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
    
    # Sorter rækker først baseret på item number
    data = []
    for row in range(2, last_row + 1):
        item_number = str(sheet.Cells(row, 1).Value)
        if item_number:
            # Split item number i dele og konverter til numeriske værdier hvor muligt
            parts = []
            for part in str(item_number).split('.'):
                try:
                    # Fjern eventuelle mellemrum og konverter til numerisk værdi
                    cleaned_part = part.strip()
                    if cleaned_part.isdigit():
                        parts.append(int(cleaned_part))
                    else:
                        parts.append(cleaned_part)
                except (ValueError, TypeError):
                    parts.append(part)
            data.append((parts, row, item_number))
    
    # Sorter data baseret på de konverterede item numbers
    def sort_key(x):
        return [str(p) if isinstance(p, str) else format(p, '05d') for p in x[0]]
    
    data.sort(key=sort_key)
    
    # Flyt rækker til deres nye positioner
    if data:
        temp_sheet = sheet.Parent.Worksheets.Add()
        for i, (_, old_row, _) in enumerate(data, start=2):
            sheet.Range(f"{old_row}:{old_row}").Copy(temp_sheet.Range(f"{i}:{i}"))
        
        for i, (_, _, _) in enumerate(data, start=2):
            temp_sheet.Range(f"{i}:{i}").Copy(sheet.Range(f"{i}:{i}"))
        
        temp_sheet.Delete()
    
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
    
    # Collapse alle grupper som standard
    sheet.Outline.ShowLevels(RowLevels=1)

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
            part_number = str(other_items_sheet.Cells(row, 2).Value)  # Kolonne B (Part Number)
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
        if matching_files:
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
            part_number = str(raw_sheet.Cells(row, 2).Value)  # Kolonne B (Part Number)
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

def main():
    """Hovedfunktion der håndterer programflow og fejlhåndtering."""
    # Gem original stdout
    original_stdout = sys.stdout
    debug_file = None
    
    try:
        # Opsæt debug logging
        debug_path = os.path.join(os.path.dirname(__file__), "debug.txt")
        debug_file = open(debug_path, 'w', encoding='utf-8')
        sys.stdout = debug_file
        
        print("=== Program Start ===")
        print(f"Debug fil oprettet: {debug_path}")
        start_time = time.time()
        
        # Initialiser COM
        pythoncom.CoInitialize()
        print("COM initialiseret")
        
        # Opret root vindue
        root = Tk()
        root.withdraw()
        print("Tkinter root vindue oprettet")
        
        # Initialiser kategorier
        initialize_categories()
        print("Kategorier initialiseret")
        
        # Vælg Excel fil
        file_path = choose_file()
        if not file_path:
            print("Ingen fil valgt - afslutter")
            return
        print(f"Valgt fil: {file_path}")
        
        # Opret progress window
        progress_window = ProgressWindow()
        print("Progress window oprettet")
        
        try:
            # Definer netværkssti
            network_path = r'C:\Coding\Python\ExcelCopyBOM\Files'
            print(f"Netværkssti: {network_path}")
            
            progress_window.update_progress(5, "Tjekker netværksforbindelse...")
            if not check_network_path(network_path):
                print(f"Kunne ikke få adgang til netværkssti: {network_path}")
                progress_window.close()
                messagebox.showerror(
                    "Netværksfejl",
                    f"Kan ikke få adgang til:\n{network_path}\n\nProgrammet afsluttes."
                )
                return
            print("Netværksforbindelse OK")

            # Scan efter filer
            progress_window.update_progress(10, "Scanner efter PDF og DWG filer...")
            files = scan_directory_concurrent(network_path)
            print(f"Fandt {len(files)} filer")
            
            if not files:
                print("Ingen filer fundet")
                progress_window.close()
                messagebox.showwarning(
                    "Advarsel",
                    f"Ingen PDF eller DWG filer blev fundet i mappen:\n{network_path}"
                )
                return

            # Behandl Excel fil
            print("Starter Excel behandling")
            dest_path = process_excel_in_order(file_path, progress_window, files)
            print(f"Excel behandling afsluttet. Resultat gemt i: {dest_path}")
            
        except Exception as e:
            print(f"Fejl under filbehandling: {str(e)}")
            if progress_window:
                progress_window.close()
            messagebox.showerror("Error", str(e))
        finally:
            if 'progress_window' in locals():
                progress_window.close()
    
    except Exception as e:
        if debug_file:
            print(f"Kritisk fejl: {str(e)}")
        messagebox.showerror("Error", str(e))
    finally:
        # Luk vinduer og frigiv ressourcer
        if 'root' in locals():
            root.destroy()
        pythoncom.CoUninitialize()
        
        # Log afslutning og varighed
        if debug_file:
            duration = time.time() - start_time
            print(f"\n=== Program Afsluttet ===")
            print(f"Varighed: {duration:.2f} sekunder")
        
        # Gendan original stdout og luk debug fil
        sys.stdout = original_stdout
        if debug_file:
            debug_file.close()

if __name__ == "__main__":
    main()
