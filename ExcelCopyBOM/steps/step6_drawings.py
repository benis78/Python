"""
TRIN 6: Kopiering af tegninger
Håndterer søgning og kopiering af PDF og DWG filer til kategori-mapper
"""

import logging
import os
import shutil
from concurrent.futures import ThreadPoolExecutor
from datetime import datetime

class DrawingCopier:
    def __init__(self, workbook, source_path: str, dest_folder: str, old_bom=None):
        self.logger = logging.getLogger('ExcelCopyBOM.Drawings')
        self.workbook = workbook
        self.source_path = source_path
        self.dest_folder = dest_folder
        self.old_bom = old_bom
        self.raw_sheet = workbook.Sheets["BOM (Raw)"]
        
    def _scan_directory_task(self, entry) -> list:
        """Scanner en mappe for PDF og DWG filer"""
        file_paths = []
        try:
            if entry.is_dir(follow_symlinks=False):
                for sub_entry in os.scandir(entry.path):
                    file_paths.extend(self._scan_directory_task(sub_entry))
            elif entry.is_file(follow_symlinks=False):
                if entry.name.lower().endswith(('.pdf', '.dwg')):
                    file_paths.append(entry.path)
        except Exception as e:
            self.logger.error(f"Fejl under scanning af {entry.path}: {str(e)}")
            
        return file_paths
        
    def _scan_directory_concurrent(self) -> list:
        """Bruger multithreading til at scanne hele kildekataloget"""
        file_paths = []
        try:
            with ThreadPoolExecutor() as executor:
                futures = [
                    executor.submit(self._scan_directory_task, entry)
                    for entry in os.scandir(self.source_path)
                ]
                for future in futures:
                    file_paths.extend(future.result())
                    
            self.logger.info(f"Fundet {len(file_paths)} filer")
            return file_paths
            
        except Exception as e:
            self.logger.error(f"Fejl under concurrent scanning: {str(e)}")
            return []
            
    def _get_latest_revision(self, base_name: str, files: list) -> tuple:
        """Finder den seneste revision af PDF og DWG filer"""
        latest_files = {}
        revisions = {'pdf': None, 'dwg': None}
        
        # Gruppér filer efter deres type
        for ext in ['.pdf', '.dwg']:
            matching_files = [
                f for f in files 
                if f.lower().endswith(ext) and os.path.basename(f).startswith(base_name)
            ]
            
            if matching_files:
                # Find seneste revision ved at se på A01, A02, etc. i filnavnet
                latest = None
                latest_rev = None
                
                for f in matching_files:
                    filename = os.path.basename(f)
                    # Find revision i formatet A01, B02, etc.
                    parts = filename.split('-')
                    if len(parts) >= 3:
                        rev_part = parts[2]  # Tag den tredje del efter split på -
                        if rev_part.startswith('A'):  # Tjek om det starter med A for at undgå ACAD
                            rev = rev_part[0]  # Tag første bogstav som revision
                            if latest_rev is None or rev > latest_rev:
                                latest_rev = rev
                                latest = f
                
                if latest:
                    latest_files[ext] = latest
                    revisions[ext.lstrip('.')] = latest_rev
                
        # Find den højeste revision
        latest_rev = None
        if revisions['pdf'] and revisions['dwg']:
            latest_rev = max(revisions['pdf'], revisions['dwg'])
        elif revisions['pdf']:
            latest_rev = revisions['pdf']
        elif revisions['dwg']:
            latest_rev = revisions['dwg']
            
        self.logger.debug(f"Fundet revisioner for {base_name}: PDF={revisions['pdf']}, DWG={revisions['dwg']}, Latest={latest_rev}")
        return latest_files, revisions, latest_rev
        
    def _get_column_indices(self) -> dict:
        """Identificerer kolonnerne i Excel arket"""
        columns = {}
        last_col = self.raw_sheet.UsedRange.Columns.Count
        
        # Find kolonner
        for col in range(1, last_col + 1):
            header = str(self.raw_sheet.Cells(1, col).Value).strip()
            if header in ['Part Number', 'Type', 'Category']:
                columns[header] = col
                
        return columns
        
    def _should_copy_drawing(self, part_number: str) -> bool:
        """Tjekker om en tegning skal kopieres baseret på gammel BOM"""
        if not self.old_bom:
            return True
            
        # TODO: Implementer sammenligning med gammel BOM (step7)
        return True
        
    def _get_drawing_category(self, row: int, columns: dict) -> str:
        """Bestemmer hvilken mappe en tegning skal placeres i"""
        category = str(self.raw_sheet.Cells(row, columns['Category']).Value).strip()
        type_ = str(self.raw_sheet.Cells(row, columns['Type']).Value).strip()
        
        # Brug Type i stedet for Category for Piping
        if category == 'Piping':
            return type_
        return category
        
    def update_revisions(self) -> bool:
        """Opdaterer REV kolonnen baseret på seneste tegningsrevision"""
        try:
            self.logger.info("Opdaterer revisioner baseret på tegninger...")
            
            # Find kolonner
            columns = self._get_column_indices()
            if not all(col in columns for col in ['Part Number', 'REV']):
                self.logger.error("Mangler nødvendige kolonner")
                return False
                
            # Scan kildekataloget
            all_files = self._scan_directory_concurrent()
            if not all_files:
                self.logger.warning("Ingen filer fundet i kildekataloget")
                return False
                
            # For hver række i BOM (Raw)
            last_row = self.raw_sheet.UsedRange.Rows.Count
            updates = 0
            
            for row in range(2, last_row + 1):
                part_number = str(self.raw_sheet.Cells(row, columns['Part Number']).Value).strip()
                if not part_number:
                    continue
                    
                # Find seneste revision af tegninger
                _, revisions, latest_rev = self._get_latest_revision(part_number, all_files)
                if latest_rev:
                    current_rev = str(self.raw_sheet.Cells(row, columns['REV']).Value).strip()
                    if latest_rev != current_rev:
                        self.logger.debug(f"Opdaterer revision for {part_number} fra {current_rev} til {latest_rev}")
                        self.raw_sheet.Cells(row, columns['REV']).Value = latest_rev
                        updates += 1
                        
            self.logger.info(f"Opdateret {updates} revisioner baseret på tegninger")
            return True
            
        except Exception as e:
            self.logger.error(f"Fejl under opdatering af revisioner: {str(e)}", exc_info=True)
            return False
            
    def copy_drawings(self) -> bool:
        """Kopierer tegninger til de relevante kategori-mapper"""
        try:
            self.logger.info("Starter kopiering af tegninger")
            
            # Find kolonner
            columns = self._get_column_indices()
            if not all(col in columns for col in ['Part Number', 'Type', 'Category']):
                self.logger.error("Mangler nødvendige kolonner")
                return False
                
            # Scan kildekataloget
            all_files = self._scan_directory_concurrent()
            if not all_files:
                self.logger.warning("Ingen filer fundet i kildekataloget")
                return False
                
            # Dictionary til at holde styr på oprettede mapper
            created_folders = set()
            
            # For hver række i BOM (Raw)
            last_row = self.raw_sheet.UsedRange.Rows.Count
            for row in range(2, last_row + 1):
                part_number = str(self.raw_sheet.Cells(row, columns['Part Number']).Value).strip()
                if not part_number:
                    continue
                    
                # Find seneste revision af tegninger
                latest_files, _, _ = self._get_latest_revision(part_number, all_files)
                if not latest_files:
                    continue
                    
                # Tjek om tegningen skal kopieres
                if not self._should_copy_drawing(part_number):
                    continue
                    
                # Bestem kategori/type mappe
                category = self._get_drawing_category(row, columns)
                if not category:
                    continue
                    
                # Opret kategori-mappe hvis der er filer at kopiere
                category_path = os.path.join(self.dest_folder, category)
                if category not in created_folders:
                    os.makedirs(category_path, exist_ok=True)
                    created_folders.add(category)
                    
                # Kopier filer
                for ext, src_file in latest_files.items():
                    try:
                        dest_file = os.path.join(category_path, os.path.basename(src_file))
                        shutil.copy2(src_file, dest_file)
                        self.logger.debug(f"Kopieret {os.path.basename(src_file)} til {category}")
                    except Exception as e:
                        self.logger.error(f"Fejl under kopiering af {src_file}: {str(e)}")
                        
            self.logger.info("Kopiering af tegninger gennemført")
            return True
            
        except Exception as e:
            self.logger.error(f"Fejl under kopiering af tegninger: {str(e)}", exc_info=True)
            return False

    def add_drawing_column(self) -> bool:
        """Tilføjer Drawing kolonne med information om PDF og DWG filer"""
        try:
            self.logger.info("Tilføjer Drawing kolonne...")
            
            # Find Part Number kolonne
            columns = self._get_column_indices()
            if 'Part Number' not in columns:
                self.logger.error("Mangler Part Number kolonne")
                return False
                
            # Find sidste kolonne
            last_col = self.raw_sheet.UsedRange.Columns.Count
            
            # Indsæt ny kolonne efter sidste kolonne
            self.raw_sheet.Cells(1, last_col + 1).Value = "Drawing"
            
            # Scan kildekataloget
            all_files = self._scan_directory_concurrent()
            if not all_files:
                self.logger.warning("Ingen filer fundet i kildekataloget")
                return False
                
            # For hver række i BOM (Raw)
            last_row = self.raw_sheet.UsedRange.Rows.Count
            for row in range(2, last_row + 1):
                part_number = str(self.raw_sheet.Cells(row, columns['Part Number']).Value).strip()
                if not part_number:
                    continue
                    
                # Find filer og revisioner for dette part number
                latest_files, revisions, _ = self._get_latest_revision(part_number, all_files)
                
                # Bestem Drawing værdi
                has_pdf = '.pdf' in latest_files
                has_dwg = '.dwg' in latest_files
                
                if has_pdf and has_dwg:
                    # Tjek om revisionerne er ens
                    if revisions['pdf'] == revisions['dwg']:
                        value = "DWG_PDF"
                    else:
                        value = "DWG~PDF"
                elif has_dwg:
                    value = "DWG"
                elif has_pdf:
                    value = "PDF"
                else:
                    value = ""  # Tomt felt hvis ingen tegninger findes
                    
                # Indsæt værdi
                self.raw_sheet.Cells(row, last_col + 1).Value = value
                
            self.logger.info("Drawing kolonne tilføjet")
            return True
            
        except Exception as e:
            self.logger.error(f"Fejl under tilføjelse af Drawing kolonne: {str(e)}", exc_info=True)
            return False 