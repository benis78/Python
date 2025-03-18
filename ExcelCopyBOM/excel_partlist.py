"""
Partlist generator modul til ExcelCopyBOM (TRIN 4)
Håndterer generering af partliste med summerede mængder
"""

import win32com.client
import pythoncom
from typing import Dict, List, Tuple
from collections import defaultdict
import re
from tkinter import messagebox

class ExcelPartlistGenerator:
    def __init__(self, workbook):
        """
        Initialiserer partlist generatoren
        :param workbook: Det åbne Excel workbook objekt
        """
        self.workbook = workbook
        self.columns = {}
        self.part_list_data = defaultdict(lambda: {
            'description': '',
            'qty': 0,
            'rev': '',
            'structure': '',
            'dimensions': {'D': '', 't': '', 'L': ''}
        })
        
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
    
    def collect_part_data(self) -> None:
        """Indsamler data fra alle kategori sheets"""
        try:
            # Gennemgå alle sheets undtagen "BOM (Raw)" og "Part List"
            for sheet in self.workbook.Sheets:
                if sheet.Name not in ["BOM (Raw)", "Part List"]:
                    self.identify_columns(sheet)
                    
                    # Gennemgå alle rækker i sheetet
                    for row in range(2, sheet.UsedRange.Rows.Count + 1):
                        part_number = str(sheet.Cells(row, self.columns["Part Number"]).Value)
                        if not part_number:
                            continue
                            
                        # Opdater eller tilføj part data
                        qty = float(sheet.Cells(row, self.columns["QTY"]).Value or 0)
                        self.part_list_data[part_number]['qty'] += qty
                        
                        # Kun opdater andre felter hvis de ikke allerede er sat
                        if not self.part_list_data[part_number]['description']:
                            self.part_list_data[part_number].update({
                                'description': sheet.Cells(row, self.columns["Description"]).Value or '',
                                'rev': sheet.Cells(row, self.columns["REV"]).Value or '',
                                'structure': sheet.Cells(row, self.columns["BOM Structure"]).Value or '',
                                'dimensions': {
                                    'D': sheet.Cells(row, self.columns["D"]).Value or '',
                                    't': sheet.Cells(row, self.columns["t"]).Value or '',
                                    'L': sheet.Cells(row, self.columns["L"]).Value or ''
                                }
                            })
                            
        except Exception as e:
            messagebox.showerror("Error", f"Fejl under indsamling af part data:\n{str(e)}")
    
    def sort_part_numbers(self) -> List[str]:
        """
        Sorterer part numbers numerisk og alfabetisk
        :return: Liste af sorterede part numbers
        """
        def sort_key(part_number: str) -> Tuple:
            # Split part number i numeriske og alfabetiske dele
            parts = re.findall(r'\d+|\D+', part_number)
            # Konverter numeriske dele til tal for korrekt sortering
            return tuple(int(p) if p.isdigit() else p for p in parts)
            
        return sorted(self.part_list_data.keys(), key=sort_key)
    
    def create_part_list_sheet(self) -> None:
        """Opretter og udfylder Part List sheetet"""
        try:
            # Opret nyt sheet eller brug eksisterende
            try:
                part_list_sheet = self.workbook.Sheets("Part List")
            except:
                part_list_sheet = self.workbook.Sheets.Add(After=self.workbook.Sheets(self.workbook.Sheets.Count))
                part_list_sheet.Name = "Part List"
            
            # Opsæt header
            headers = ["Item", "Part Number", "REV", "Description", "QTY", "BOM Structure", "D", "t", "L"]
            for col, header in enumerate(headers, 1):
                part_list_sheet.Cells(1, col).Value = header
                part_list_sheet.Cells(1, col).Font.Bold = True
                part_list_sheet.Cells(1, col).Interior.Color = 0xFFFF00  # Gul baggrund
            
            # Udfyld data
            sorted_parts = self.sort_part_numbers()
            for row, part_number in enumerate(sorted_parts, 2):
                data = self.part_list_data[part_number]
                
                # Indsæt værdier
                part_list_sheet.Cells(row, 1).Value = row - 1  # Item number
                part_list_sheet.Cells(row, 2).Value = part_number
                part_list_sheet.Cells(row, 3).Value = data['rev']
                part_list_sheet.Cells(row, 4).Value = data['description']
                part_list_sheet.Cells(row, 5).Value = data['qty']
                part_list_sheet.Cells(row, 6).Value = data['structure']
                part_list_sheet.Cells(row, 7).Value = data['dimensions']['D']
                part_list_sheet.Cells(row, 8).Value = data['dimensions']['t']
                part_list_sheet.Cells(row, 9).Value = data['dimensions']['L']
            
            # Formatér kolonner
            part_list_sheet.Columns.AutoFit()
            
            # Tilføj filter
            header_range = part_list_sheet.Range(
                part_list_sheet.Cells(1, 1),
                part_list_sheet.Cells(1, len(headers))
            )
            header_range.AutoFilter()
            
        except Exception as e:
            messagebox.showerror("Error", f"Fejl under oprettelse af Part List:\n{str(e)}")
    
    def generate_part_list(self) -> bool:
        """
        Hovedfunktion der genererer partlisten
        :return: True hvis succesfuld, ellers False
        """
        try:
            self.collect_part_data()
            self.create_part_list_sheet()
            
            # Gem ændringer
            self.workbook.Save()
            return True
            
        except Exception as e:
            messagebox.showerror("Error", f"Fejl under generering af Part List:\n{str(e)}")
            return False

if __name__ == "__main__":
    # Test kode
    pythoncom.CoInitialize()
    excel = win32com.client.Dispatch("Excel.Application")
    wb = excel.Workbooks.Open("test.xlsx")
    
    generator = ExcelPartlistGenerator(wb)
    generator.generate_part_list()
    
    wb.Close(SaveChanges=True)
    excel.Quit()
    pythoncom.CoUninitialize() 