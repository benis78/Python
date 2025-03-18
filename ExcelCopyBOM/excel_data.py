"""
Data processor modul til ExcelCopyBOM (TRIN 2)
Håndterer indlæsning og validering af Excel data
"""

import pandas as pd
import openpyxl
from openpyxl.utils import get_column_letter
import os
from datetime import datetime
import re
from tkinter import messagebox
import win32com.client
import pythoncom

class ExcelDataProcessor:
    def __init__(self, file_path):
        self.file_path = file_path
        self.workbook = None
        self.excel_app = None
        self.part_number = None
        self.rev = None
        self.columns = {}
        
    def initialize_excel(self):
        """Initialiserer Excel med COM objekter"""
        pythoncom.CoInitialize()
        self.excel_app = win32com.client.Dispatch("Excel.Application")
        self.excel_app.Visible = False
        self.excel_app.DisplayAlerts = False
        
    def cleanup_excel(self):
        """Lukker Excel og frigør ressourcer"""
        if self.workbook:
            self.workbook.Close(SaveChanges=True)
        if self.excel_app:
            self.excel_app.Quit()
        pythoncom.CoUninitialize()
        
    def extract_part_number_info(self):
        """Udtrækker part number og revision fra filnavn"""
        file_name = os.path.basename(self.file_path)
        match = re.match(r"(\d{4}-[\d.]+[A-Za-z0-9-]+).*?(?:--|-).*", file_name)
        if match:
            self.part_number = match.group(1)
            # Find revision hvis den er i part number
            rev_match = re.search(r"-([A-Z])$", self.part_number)
            if rev_match:
                self.rev = rev_match.group(1)
                self.part_number = self.part_number[:-2]  # Fjern revision fra part number
        
    def identify_columns(self, sheet):
        """Identificerer kolonner ud fra header række"""
        required_columns = [
            "Item", "Part Number", "REV", "BOM Structure", 
            "Description", "QTY", "D", "t", "L"
        ]
        
        for col in range(1, sheet.MaxColumns + 1):
            header = sheet.Cells(1, col).Value
            if header in required_columns:
                self.columns[header] = col
                
        # Verificer at alle påkrævede kolonner er fundet
        missing = [col for col in required_columns if col not in self.columns]
        if missing:
            raise ValueError(f"Manglende påkrævede kolonner: {', '.join(missing)}")
    
    def insert_arrangement_row(self, sheet):
        """Indsætter arrangement række i position 2"""
        sheet.Rows(2).Insert()
        
        # Indsæt værdier
        sheet.Cells(2, self.columns["Item"]).Value = 0
        sheet.Cells(2, self.columns["Part Number"]).Value = self.part_number
        sheet.Cells(2, self.columns["REV"]).Value = self.rev or ""
        sheet.Cells(2, self.columns["BOM Structure"]).Value = "Inseparable"
        
        # Bestem beskrivelse baseret på part number
        if self.part_number.startswith("0000-"):
            description = "Basic Equipment Drawing"
        else:
            description = "Arrangement Drawing"
        sheet.Cells(2, self.columns["Description"]).Value = description
        
        # Indsæt standard værdier
        for col in ["QTY", "D", "t", "L"]:
            if col in self.columns:
                sheet.Cells(2, self.columns[col]).Value = 1
    
    def process_equipment_rows(self, sheet, include_equipment):
        """Håndterer equipment rækker baseret på checkbox"""
        if not include_equipment:
            equipment_patterns = ["0000-700", "0000-701", "0000-702"]
            rows_to_delete = []
            
            # Find rækker der skal slettes
            for row in range(sheet.UsedRange.Rows.Count, 1, -1):
                part_number = str(sheet.Cells(row, self.columns["Part Number"]).Value)
                if any(part_number.startswith(pattern) for pattern in equipment_patterns):
                    rows_to_delete.append(row)
            
            # Slet fundne rækker
            for row in rows_to_delete:
                sheet.Rows(row).Delete()
    
    def process_file(self, include_equipment=False):
        """Hovedfunktion der behandler Excel filen"""
        try:
            self.initialize_excel()
            self.extract_part_number_info()
            
            # Åbn workbook
            self.workbook = self.excel_app.Workbooks.Open(self.file_path)
            sheet = self.workbook.Sheets(1)
            sheet.Name = "BOM (Raw)"
            
            # Identificer kolonner
            self.identify_columns(sheet)
            
            # Indsæt arrangement række
            self.insert_arrangement_row(sheet)
            
            # Håndter equipment rækker
            self.process_equipment_rows(sheet, include_equipment)
            
            # Gem ændringer
            self.workbook.Save()
            
            return True
            
        except Exception as e:
            messagebox.showerror("Error", f"Fejl under behandling af Excel fil:\n{str(e)}")
            return False
            
        finally:
            self.cleanup_excel()

if __name__ == "__main__":
    # Test kode
    processor = ExcelDataProcessor("test.xlsx")
    processor.process_file() 