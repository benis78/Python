"""
Hovedmodul til ExcelCopyBOM
Koordinerer alle andre moduler og implementerer den overordnede proces
"""

import os
import sys
from datetime import datetime
import logging
from typing import Set, Optional
import pythoncom
import win32com.client
from tkinter import messagebox

# Import af vores moduler
from excel_gui import ExcelCopyBOMGUI
from excel_data import ExcelDataProcessor
from excel_category import ExcelCategoryProcessor
from excel_partlist import ExcelPartlistGenerator
from excel_file_copier import ExcelFileCopier
from excel_compare import ExcelCompareProcessor

class ExcelCopyBOMMain:
    def __init__(self):
        """Initialiserer hovedprogrammet"""
        self.gui = ExcelCopyBOMGUI(self.process_files)
        self.logger = self.setup_logging()
        
    def setup_logging(self) -> logging.Logger:
        """Opsætter logging for hovedprogrammet"""
        logger = logging.getLogger('ExcelCopyBOM')
        logger.setLevel(logging.DEBUG)
        
        # Tilføj fil handler
        log_dir = os.path.join(os.path.dirname(__file__), 'logs')
        os.makedirs(log_dir, exist_ok=True)
        
        log_file = os.path.join(
            log_dir,
            f'process_{datetime.now().strftime("%Y%m%d_%H%M%S")}.log'
        )
        
        handler = logging.FileHandler(log_file, encoding='utf-8')
        handler.setFormatter(
            logging.Formatter('%(asctime)s - %(name)s - %(levelname)s - %(message)s')
        )
        logger.addHandler(handler)
        
        return logger
    
    def initialize_excel(self) -> tuple:
        """
        Initialiserer Excel og returnerer application objekt
        :return: Tuple med (excel_app, success)
        """
        try:
            # Prøv først at få fat i en eksisterende Excel instans
            excel = win32com.client.GetObject(None, "Excel.Application")
        except:
            # Hvis det fejler, opret en ny instans
            try:
                pythoncom.CoInitialize()
                excel = win32com.client.Dispatch("Excel.Application")
                excel.Visible = False
                excel.DisplayAlerts = False
            except Exception as e:
                self.logger.error(f"Fejl under initialisering af Excel: {str(e)}")
                messagebox.showerror(
                    "Excel Fejl",
                    "Kunne ikke initialisere Excel. Er Excel installeret korrekt?"
                )
                return None, False
        
        return excel, True
    
    def cleanup_excel(self, excel_app) -> None:
        """
        Lukker Excel ned
        :param excel_app: Excel application objekt
        """
        if excel_app:
            try:
                excel_app.Quit()
            except:
                pass
            finally:
                pythoncom.CoUninitialize()
    
    def extract_part_numbers(self, workbook) -> Set[str]:
        """
        Udtrækker alle unikke part numbers fra workbook
        :param workbook: Excel workbook objekt
        :return: Set af part numbers
        """
        part_numbers = set()
        try:
            sheet = workbook.Sheets("Part List")
            last_row = sheet.UsedRange.Rows.Count
            
            for row in range(2, last_row + 1):
                part_number = str(sheet.Cells(row, 2).Value)  # Part Number er i kolonne 2
                if part_number:
                    # Fjern eventuel revision fra part number
                    part_number = part_number.split('-')[0]
                    part_numbers.add(part_number)
                    
        except Exception as e:
            self.logger.error(f"Fejl under udtrækning af part numbers: {str(e)}")
            
        return part_numbers
    
    def process_files(
        self,
        bom_file: str,
        old_bom_file: str,
        source_dir: str,
        include_equipment: bool,
        find_rev_before: Optional[datetime],
        include_datasheet: bool
    ) -> bool:
        """
        Hovedfunktion der behandler filerne
        :param bom_file: Sti til BOM Excel fil
        :param old_bom_file: Sti til gammel BOM fil (kan være None)
        :param source_dir: Sti til kildemappe med PDF/DWG filer
        :param include_equipment: Om equipment skal inkluderes
        :param find_rev_before: Dato at finde revisioner før (kan være None)
        :param include_datasheet: Om datasheets skal inkluderes
        :return: True hvis succesfuld
        """
        excel_app = None
        workbook = None
        success = False
        
        try:
            # Initialiser Excel
            excel_app, excel_ok = self.initialize_excel()
            if not excel_ok:
                return False
            
            # Start processen
            self.logger.info(f"Starter behandling af {bom_file}")
            self.gui.update_progress(0, "Initialiserer...")
            
            # TRIN 1: Åbn og initialiser workbook
            workbook = excel_app.Workbooks.Open(bom_file)
            
            # TRIN 2: Behandl raw data
            self.gui.update_progress(20, "Behandler BOM data...")
            data_processor = ExcelDataProcessor(workbook)
            if not data_processor.process_file(include_equipment):
                raise Exception("Fejl under behandling af BOM data")
            
            # TRIN 3: Kategoriser data
            self.gui.update_progress(40, "Kategoriserer data...")
            category_processor = ExcelCategoryProcessor(workbook)
            if not category_processor.process_categories():
                raise Exception("Fejl under kategorisering")
            
            # TRIN 4: Generer partliste
            self.gui.update_progress(60, "Genererer partliste...")
            partlist_generator = ExcelPartlistGenerator(workbook)
            if not partlist_generator.generate_part_list():
                raise Exception("Fejl under generering af partliste")
            
            # TRIN 5: Kopier filer
            if source_dir:
                self.gui.update_progress(70, "Kopierer filer...")
                target_dir = os.path.splitext(bom_file)[0]  # Fjern .xlsx
                
                # Opret målmappe hvis den ikke findes
                os.makedirs(target_dir, exist_ok=True)
                
                # Udtræk part numbers og kopier filer
                part_numbers = self.extract_part_numbers(workbook)
                file_copier = ExcelFileCopier(
                    source_dir,
                    target_dir,
                    find_rev_before
                )
                if not file_copier.process_part_numbers(part_numbers):
                    raise Exception("Fejl under kopiering af filer")
            
            # TRIN 6: Sammenlign med gammel BOM hvis specificeret
            if old_bom_file:
                self.gui.update_progress(90, "Sammenligner med gammel BOM...")
                compare_processor = ExcelCompareProcessor(workbook, old_bom_file)
                if not compare_processor.process_comparison():
                    raise Exception("Fejl under sammenligning med gammel BOM")
            
            # Gem ændringer
            workbook.Save()
            
            self.gui.update_progress(100, "Færdig!")
            self.logger.info("Proces gennemført succesfuldt")
            success = True
            
            # Vis færdig besked
            messagebox.showinfo(
                "Færdig",
                "BOM behandling er gennemført succesfuldt!\n\n"
                f"Output er gemt i:\n{os.path.dirname(bom_file)}"
            )
            
        except Exception as e:
            self.logger.error(f"Fejl under behandling: {str(e)}")
            messagebox.showerror(
                "Fejl",
                f"Der opstod en fejl under behandlingen:\n{str(e)}\n\n"
                "Se log filen for flere detaljer."
            )
            
        finally:
            # Luk workbook og Excel
            if workbook:
                try:
                    workbook.Close(SaveChanges=True)
                except:
                    pass
            
            self.cleanup_excel(excel_app)
            
        return success

def main():
    """Entry point for programmet"""
    app = ExcelCopyBOMMain()
    app.gui.run()

if __name__ == "__main__":
    main() 