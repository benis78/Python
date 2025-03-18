"""
Sammenlignings modul til ExcelCopyBOM (TRIN 6)
Håndterer sammenligning af gamle og nye BOM lister samt farvekodning af ændringer
"""

import win32com.client
import pythoncom
from typing import Dict, List, Set, Tuple, Optional
from dataclasses import dataclass
from enum import Enum
import os
from tkinter import messagebox
import logging

class ChangeType(Enum):
    """Enum for forskellige typer af ændringer"""
    NEW = "Ny"
    DELETED = "Slettet"
    MODIFIED = "Ændret"
    UNCHANGED = "Uændret"

@dataclass
class PartData:
    """Dataklasse til at holde information om en part"""
    part_number: str
    rev: str
    description: str
    qty: float
    structure: str
    dimensions: Dict[str, str]

class ExcelCompareProcessor:
    def __init__(self, new_workbook, old_workbook_path: str):
        """
        Initialiserer sammenlignings processoren
        :param new_workbook: Det nye workbook objekt
        :param old_workbook_path: Sti til det gamle workbook
        """
        self.new_workbook = new_workbook
        self.old_workbook_path = old_workbook_path
        self.old_workbook = None
        self.columns = {}
        self.changes: Dict[str, Tuple[ChangeType, Optional[PartData], Optional[PartData]]] = {}
        
        # Opsæt logging
        self.logger = logging.getLogger('ExcelCompare')
        self.logger.setLevel(logging.DEBUG)
        
        if not self.logger.handlers:
            log_file = os.path.join(os.path.dirname(old_workbook_path), 'compare.log')
            handler = logging.FileHandler(log_file, encoding='utf-8')
            handler.setFormatter(logging.Formatter('%(asctime)s - %(levelname)s - %(message)s'))
            self.logger.addHandler(handler)
    
    def initialize_excel(self) -> None:
        """Initialiserer Excel og åbner det gamle workbook"""
        try:
            excel = win32com.client.GetObject(None, "Excel.Application")
            self.old_workbook = excel.Workbooks.Open(self.old_workbook_path)
        except:
            pythoncom.CoInitialize()
            excel = win32com.client.Dispatch("Excel.Application")
            excel.Visible = False
            excel.DisplayAlerts = False
            self.old_workbook = excel.Workbooks.Open(self.old_workbook_path)
    
    def cleanup_excel(self) -> None:
        """Lukker det gamle workbook"""
        if self.old_workbook:
            self.old_workbook.Close(SaveChanges=False)
    
    def identify_columns(self, sheet) -> None:
        """Identificerer kolonner i et sheet"""
        required_columns = [
            "Item", "Part Number", "REV", "BOM Structure", 
            "Description", "QTY", "D", "t", "L"
        ]
        
        for col in range(1, sheet.UsedRange.Columns.Count + 1):
            header = sheet.Cells(1, col).Value
            if header in required_columns:
                self.columns[header] = col
    
    def get_part_data(self, sheet, row: int) -> PartData:
        """
        Henter data for en part fra et sheet
        :param sheet: Excel sheet
        :param row: Række nummer
        :return: PartData objekt
        """
        return PartData(
            part_number=str(sheet.Cells(row, self.columns["Part Number"]).Value),
            rev=str(sheet.Cells(row, self.columns["REV"]).Value or ''),
            description=str(sheet.Cells(row, self.columns["Description"]).Value or ''),
            qty=float(sheet.Cells(row, self.columns["QTY"]).Value or 0),
            structure=str(sheet.Cells(row, self.columns["BOM Structure"]).Value or ''),
            dimensions={
                'D': str(sheet.Cells(row, self.columns["D"]).Value or ''),
                't': str(sheet.Cells(row, self.columns["t"]).Value or ''),
                'L': str(sheet.Cells(row, self.columns["L"]).Value or '')
            }
        )
    
    def compare_part_data(self, old_data: PartData, new_data: PartData) -> bool:
        """
        Sammenligner to PartData objekter
        :return: True hvis der er ændringer
        """
        return (
            old_data.rev != new_data.rev or
            old_data.description != new_data.description or
            abs(old_data.qty - new_data.qty) > 0.001 or
            old_data.structure != new_data.structure or
            old_data.dimensions != new_data.dimensions
        )
    
    def collect_changes(self) -> None:
        """Indsamler alle ændringer mellem gamle og nye data"""
        try:
            # Hent data fra begge workbooks
            new_sheet = self.new_workbook.Sheets("Part List")
            old_sheet = self.old_workbook.Sheets("Part List")
            
            self.identify_columns(new_sheet)
            
            # Indsaml alle part numbers og deres data
            new_parts: Dict[str, PartData] = {}
            old_parts: Dict[str, PartData] = {}
            
            # Indsaml nye parts
            for row in range(2, new_sheet.UsedRange.Rows.Count + 1):
                part_data = self.get_part_data(new_sheet, row)
                new_parts[part_data.part_number] = part_data
            
            # Indsaml gamle parts
            for row in range(2, old_sheet.UsedRange.Rows.Count + 1):
                part_data = self.get_part_data(old_sheet, row)
                old_parts[part_data.part_number] = part_data
            
            # Find ændringer
            all_parts = set(new_parts.keys()) | set(old_parts.keys())
            
            for part_number in all_parts:
                if part_number not in old_parts:
                    # Ny part
                    self.changes[part_number] = (ChangeType.NEW, None, new_parts[part_number])
                elif part_number not in new_parts:
                    # Slettet part
                    self.changes[part_number] = (ChangeType.DELETED, old_parts[part_number], None)
                else:
                    # Sammenlign eksisterende parts
                    old_data = old_parts[part_number]
                    new_data = new_parts[part_number]
                    
                    if self.compare_part_data(old_data, new_data):
                        self.changes[part_number] = (ChangeType.MODIFIED, old_data, new_data)
                    else:
                        self.changes[part_number] = (ChangeType.UNCHANGED, old_data, new_data)
            
        except Exception as e:
            self.logger.error(f"Fejl under indsamling af ændringer: {str(e)}")
            raise
    
    def create_compare_sheet(self) -> None:
        """Opretter et nyt sheet med sammenligningen"""
        try:
            # Opret nyt sheet
            compare_sheet = self.new_workbook.Sheets.Add(After=self.new_workbook.Sheets(self.new_workbook.Sheets.Count))
            compare_sheet.Name = "Compare"
            
            # Opsæt headers
            headers = [
                "Part Number", "Status", "Gammel REV", "Ny REV",
                "Gammel Beskrivelse", "Ny Beskrivelse",
                "Gammel QTY", "Ny QTY",
                "Gammel Structure", "Ny Structure",
                "Gamle Dimensioner", "Nye Dimensioner"
            ]
            
            for col, header in enumerate(headers, 1):
                compare_sheet.Cells(1, col).Value = header
                compare_sheet.Cells(1, col).Font.Bold = True
                compare_sheet.Cells(1, col).Interior.Color = 0xFFFF00  # Gul
            
            # Tilføj data
            row = 2
            for part_number, (change_type, old_data, new_data) in sorted(self.changes.items()):
                # Indsæt grunddata
                compare_sheet.Cells(row, 1).Value = part_number
                compare_sheet.Cells(row, 2).Value = change_type.value
                
                # Sæt baggrundsfarve baseret på ændring
                color = {
                    ChangeType.NEW: 0x92D050,      # Grøn
                    ChangeType.DELETED: 0xFF0000,  # Rød
                    ChangeType.MODIFIED: 0xFFFF00, # Gul
                    ChangeType.UNCHANGED: None     # Ingen farve
                }[change_type]
                
                if color:
                    compare_sheet.Range(
                        compare_sheet.Cells(row, 1),
                        compare_sheet.Cells(row, len(headers))
                    ).Interior.Color = color
                
                # Indsæt sammenligningsdata
                if old_data:
                    compare_sheet.Cells(row, 3).Value = old_data.rev
                    compare_sheet.Cells(row, 5).Value = old_data.description
                    compare_sheet.Cells(row, 7).Value = old_data.qty
                    compare_sheet.Cells(row, 9).Value = old_data.structure
                    compare_sheet.Cells(row, 11).Value = str(old_data.dimensions)
                
                if new_data:
                    compare_sheet.Cells(row, 4).Value = new_data.rev
                    compare_sheet.Cells(row, 6).Value = new_data.description
                    compare_sheet.Cells(row, 8).Value = new_data.qty
                    compare_sheet.Cells(row, 10).Value = new_data.structure
                    compare_sheet.Cells(row, 12).Value = str(new_data.dimensions)
                
                row += 1
            
            # Formatér sheet
            compare_sheet.Columns.AutoFit()
            compare_sheet.Range(
                compare_sheet.Cells(1, 1),
                compare_sheet.Cells(1, len(headers))
            ).AutoFilter()
            
        except Exception as e:
            self.logger.error(f"Fejl under oprettelse af sammenlignings-sheet: {str(e)}")
            raise
    
    def save_changes_report(self) -> None:
        """Gemmer en tekstrapport over ændringer"""
        try:
            report_path = os.path.join(
                os.path.dirname(self.old_workbook_path),
                'changes_report.txt'
            )
            
            with open(report_path, 'w', encoding='utf-8') as f:
                f.write("Rapport over ændringer i BOM\n")
                f.write("=" * 60 + "\n\n")
                
                # Gruppér efter ændringtype
                by_type: Dict[ChangeType, List[str]] = {t: [] for t in ChangeType}
                for part_number, (change_type, _, _) in self.changes.items():
                    by_type[change_type].append(part_number)
                
                # Skriv sammendrag
                f.write("SAMMENDRAG\n")
                f.write("-" * 40 + "\n")
                for change_type in ChangeType:
                    count = len(by_type[change_type])
                    if count > 0:
                        f.write(f"{change_type.value}: {count} parts\n")
                
                # Skriv detaljer
                f.write("\nDETALJER\n")
                f.write("-" * 40 + "\n")
                for change_type in ChangeType:
                    if by_type[change_type]:
                        f.write(f"\n{change_type.value}:\n")
                        for part_number in sorted(by_type[change_type]):
                            f.write(f"  - {part_number}\n")
                            
                            # Tilføj ændringsdetaljer for modificerede parts
                            if change_type == ChangeType.MODIFIED:
                                _, old_data, new_data = self.changes[part_number]
                                if old_data.rev != new_data.rev:
                                    f.write(f"    REV: {old_data.rev} -> {new_data.rev}\n")
                                if old_data.qty != new_data.qty:
                                    f.write(f"    QTY: {old_data.qty} -> {new_data.qty}\n")
                                if old_data.description != new_data.description:
                                    f.write(f"    Description ændret\n")
                                if old_data.dimensions != new_data.dimensions:
                                    f.write(f"    Dimensioner ændret\n")
            
            self.logger.info(f"Ændringsrapport gemt: {report_path}")
            
        except Exception as e:
            self.logger.error(f"Fejl under gemning af ændringsrapport: {str(e)}")
    
    def process_comparison(self) -> bool:
        """
        Hovedfunktion der håndterer sammenligningen
        :return: True hvis succesfuld
        """
        try:
            self.initialize_excel()
            self.collect_changes()
            self.create_compare_sheet()
            self.save_changes_report()
            
            # Gem ændringer
            self.new_workbook.Save()
            return True
            
        except Exception as e:
            self.logger.error(f"Fejl under sammenligning: {str(e)}")
            messagebox.showerror("Error", f"Fejl under sammenligning:\n{str(e)}")
            return False
            
        finally:
            self.cleanup_excel()

if __name__ == "__main__":
    # Test kode
    pythoncom.CoInitialize()
    excel = win32com.client.Dispatch("Excel.Application")
    new_wb = excel.Workbooks.Open(r"C:\Test\new.xlsx")
    
    processor = ExcelCompareProcessor(new_wb, r"C:\Test\old.xlsx")
    processor.process_comparison()
    
    new_wb.Close(SaveChanges=True)
    excel.Quit()
    pythoncom.CoUninitialize() 