"""
Kategoriserings modul til ExcelCopyBOM (TRIN 3)
Håndterer kategorisering af part numbers og oprettelse af kategori-faner
"""

import win32com.client
import pythoncom
from typing import Dict, List, Optional
import re
from tkinter import messagebox

class ExcelCategoryProcessor:
    def __init__(self, workbook):
        """
        Initialiserer kategoriserings processoren
        :param workbook: Det åbne Excel workbook objekt
        """
        self.workbook = workbook
        self.categories = {
            "Piping": ["1", "2", "3"],
            "Valves": ["4"],
            "Instruments": ["5"],
            "Electrical": ["6"],
            "Steel": ["7"],
            "Equipment": ["0000-7"],
            "Insulation": ["8"],
            "Documentation": ["9"],
            "Misc": []  # Alt andet
        }
        self.columns = {}
        
    def identify_columns(self, sheet) -> None:
        """Identificerer kolonner i det aktive sheet"""
        required_columns = [
            "Item", "Part Number", "REV", "BOM Structure", 
            "Description", "QTY", "D", "t", "L"
        ]
        
        for col in range(1, sheet.UsedRange.Columns.Count + 1):
            header = sheet.Cells(1, col).Value
            if header in required_columns:
                self.columns[header] = col
    
    def get_category(self, part_number: str) -> str:
        """
        Bestemmer kategorien for et part number
        :param part_number: Part number der skal kategoriseres
        :return: Kategori navn
        """
        if not part_number:
            return "Misc"
            
        # Fjern eventuel revision fra part number
        part_number = re.sub(r"-[A-Z]$", "", part_number)
        
        # Find første ciffer i part number
        match = re.search(r"\d", part_number)
        if not match:
            return "Misc"
            
        first_digit = part_number[match.start()]
        
        # Check hver kategori
        for category, patterns in self.categories.items():
            for pattern in patterns:
                if part_number.startswith(pattern):
                    return category
                    
        # Special case for piping sub-kategorier
        if first_digit in self.categories["Piping"]:
            return f"Piping-{first_digit}xxx"
            
        return "Misc"
    
    def create_category_sheet(self, category: str) -> None:
        """
        Opretter et nyt sheet for en kategori
        :param category: Navn på kategorien
        """
        try:
            new_sheet = self.workbook.Sheets.Add(After=self.workbook.Sheets(self.workbook.Sheets.Count))
            new_sheet.Name = category
            
            # Kopier header række fra BOM (Raw)
            source_sheet = self.workbook.Sheets("BOM (Raw)")
            source_sheet.Rows(1).Copy(new_sheet.Rows(1))
            
            # Formatér header
            header_range = new_sheet.Range(new_sheet.Cells(1, 1), new_sheet.Cells(1, len(self.columns)))
            header_range.Font.Bold = True
            header_range.Interior.Color = 0xFFFF00  # Gul baggrund
            
        except Exception as e:
            messagebox.showerror("Error", f"Fejl under oprettelse af kategori sheet '{category}':\n{str(e)}")
    
    def copy_rows_to_category(self, source_sheet, category: str) -> None:
        """
        Kopierer relevante rækker til kategori sheet
        :param source_sheet: Kildesheet (BOM Raw)
        :param category: Målkategori
        """
        try:
            target_sheet = self.workbook.Sheets(category)
            next_row = target_sheet.UsedRange.Rows.Count + 1
            
            # Gennemgå alle rækker i kildesheetet
            for row in range(2, source_sheet.UsedRange.Rows.Count + 1):
                part_number = str(source_sheet.Cells(row, self.columns["Part Number"]).Value)
                if self.get_category(part_number) == category:
                    # Kopier hele rækken
                    source_sheet.Rows(row).Copy(target_sheet.Rows(next_row))
                    next_row += 1
                    
        except Exception as e:
            messagebox.showerror("Error", f"Fejl under kopiering af rækker til '{category}':\n{str(e)}")
    
    def process_categories(self) -> bool:
        """
        Hovedfunktion der håndterer kategorisering
        :return: True hvis succesfuld, ellers False
        """
        try:
            source_sheet = self.workbook.Sheets("BOM (Raw)")
            self.identify_columns(source_sheet)
            
            # Find alle unikke kategorier i brug
            used_categories = set()
            for row in range(2, source_sheet.UsedRange.Rows.Count + 1):
                part_number = str(source_sheet.Cells(row, self.columns["Part Number"]).Value)
                category = self.get_category(part_number)
                used_categories.add(category)
            
            # Opret sheets for hver kategori i brug
            for category in sorted(used_categories):
                self.create_category_sheet(category)
                self.copy_rows_to_category(source_sheet, category)
            
            # Gem ændringer
            self.workbook.Save()
            return True
            
        except Exception as e:
            messagebox.showerror("Error", f"Fejl under kategorisering:\n{str(e)}")
            return False

if __name__ == "__main__":
    # Test kode
    pythoncom.CoInitialize()
    excel = win32com.client.Dispatch("Excel.Application")
    wb = excel.Workbooks.Open("test.xlsx")
    
    processor = ExcelCategoryProcessor(wb)
    processor.process_categories()
    
    wb.Close(SaveChanges=True)
    excel.Quit()
    pythoncom.CoUninitialize() 