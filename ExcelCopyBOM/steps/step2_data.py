"""
TRIN 2: Data Indlæsning og Validering
Håndterer indlæsning og validering af Excel data
- Indlæs Excel BOM filen. 
- Fra filnavnet kan udtrækkes "Part Number" efter step3_categorize.py "REV" det sidste bogstav efter før " - " mellemrum bindestreg mellemrum.
- Kildefil må ikke ændres
- Identificere Kolonnenumre ud fra 1.række celleværdier (Item, Part Number, REV, BOM Structure, Description, QTY, D, t, L)
- Indskyd en linje i række 2 med følgende værdier. "Item"= 0, "Part Number"=Udtræk fra filnavn, REV=Udtræk fra filnavn, "BOM Structure"=None, "Description"="Arrangement Drawing" Eller "Basic Equipment Drawing", "QTY"=1, "D"=1, "t"=1, "L"=1 
- Find "Part Number" kolonnen. Hvis Check box i TRIN 1 er falsk, så skal alle rækker der starter med 0000-700, 0000-701, 0000-702 slettes. Hvis der er rækker som her kun består af bogstaver, så skal der poppe et tkinter vindue op med de rækker og brugeren skal promptes om man vil slette disse rækker.
- Find "Part Number" og "REV" kolonnen og hvis der indgår et revisionsnummer i "Part Number" skal det flyttes til "REV" Kolonnen efter extract_revision_from_partnumber funktionen
- Find "Item" kolonnen og indeksere alle rækker ud fra parent/child hierarkiet (Multi-Level Numbering System,Tree Structure)
- Find "BOM Structure" kolonnen. Hvis en række er "Inseparable" eller "Part Number" starter med 0000-3 skal alle child til den parent slettes. Hvis en række er "Phantom" skal denne række slettes.
- Find "QTY" kolonnen og indsæt en ny kolonne efter "Total QTY" hvor man ganger "QTY" tal med"QTY" i parent
- Kopier tegninger til kategori-mapper baseret på Type/Category (Piping bruger Type i stedet for Category)
- Find seneste tegningsrevision for hver part number og opdater "REV" kolonnen tilsvarende
- Find seneste "REV" med revNew eller hvis Find "REV" filer før dato er true så skal du finde seneste "REV" før den dato. Erstat alle "REV" værdier efter søgning
- Indsæt ny kolonne efter sidst brugte kolonne og navngiv den "Drawing" i 1.række. Indsæt her værdi efter "has_pdf" og "has_dwg" funktionen
- Det hele skal gøres i Fane 1 "Bom (Raw)"
"""

import logging
import re
import os
from datetime import datetime
import win32com.client
import tkinter as tk
from tkinter import messagebox
from .step3_categorize import PartNumberCategorizer
import shutil

class ExcelDataLoader:
    def __init__(self, workbook, logger):
        """Initialiserer klassen med en Excel workbook og logger"""
        self.workbook = workbook
        self.logger = logger
        self.sheet = workbook.ActiveSheet
        self.columns = self._get_column_indices()
        self.excel = workbook.Application
        self.categorizer = PartNumberCategorizer()
        self.include_suppliers = True
        
        # Gem oprindelige Excel indstillinger
        self._save_excel_settings()
        
    def _save_excel_settings(self):
        """Gemmer oprindelige Excel indstillinger"""
        self.original_calculation = self.excel.Calculation
        self.original_screen_updating = self.excel.ScreenUpdating
        self.original_display_alerts = self.excel.DisplayAlerts
        
    def _optimize_excel_settings(self):
        """Optimerer Excel indstillinger for bedre performance"""
        self.excel.Calculation = -4135  # xlCalculationManual
        self.excel.ScreenUpdating = False
        self.excel.DisplayAlerts = False
        
    def _restore_excel_settings(self):
        """Gendanner oprindelige Excel indstillinger"""
        self.excel.Calculation = self.original_calculation
        self.excel.ScreenUpdating = self.original_screen_updating
        self.excel.DisplayAlerts = self.original_display_alerts
        
    def _extract_file_info(self) -> dict:
        """Udtrækker information fra filnavnet"""
        filename = os.path.basename(self.workbook.FullName)
        self.logger.debug(f"Udtrækker information fra filnavn: {filename}")
        
        # Fjern .xlsx hvis det findes
        if filename.endswith('.xlsx'):
            filename = filename[:-5]
            
        # Fjern " - BOM" hvis det findes
        if " - BOM" in filename:
            filename = filename.split(" - BOM")[0]
            
        # Brug regex til at matche både -- og -X formater
        match = re.match(r'^(.*?)(-[A-Za-z0-9-]|--)$', filename)
        if match:
            part_number = match.group(1).strip()
            rev_part = match.group(2)
            rev = '-' if rev_part == '--' else rev_part[1:]
            
            # Brug categorizer til at bestemme type
            type_, category = self.categorizer.categorize(part_number)
            drawing_type = 'arrangement' if type_ == 'Area Drawing' else 'basic'
            
            return {
                'part_number': part_number,
                'rev': rev,
                'type': drawing_type
            }
        
        self.logger.warning(f"Kunne ikke udtrække information fra filnavn: {filename}")
        return None
        
    def _get_column_indices(self) -> dict:
        """Identificerer kolonnerne i Excel arket"""
        required_columns = {
            'Item': None,
            'Part Number': None,
            'REV': None,
            'BOM Structure': None,
            'Description': None,
            'QTY': None,
            'D': None,
            't': None,
            'L': None,
            'Type': None,  # Ny kolonne
            'Category': None  # Ny kolonne
        }
        
        # Find sidste kolonne
        last_col = self.sheet.UsedRange.Columns.Count
        
        # Scan første række
        for col in range(1, last_col + 1):
            cell_value = str(self.sheet.Cells(1, col).Value).strip()
            if cell_value in required_columns:
                required_columns[cell_value] = col
                
        # Tilføj manglende Type og Category kolonner
        if required_columns['Type'] is None:
            last_col += 1
            self.sheet.Cells(1, last_col).Value = 'Type'
            required_columns['Type'] = last_col
            
        if required_columns['Category'] is None:
            last_col += 1
            self.sheet.Cells(1, last_col).Value = 'Category'
            required_columns['Category'] = last_col
            
        # Tjek for manglende kolonner (undtagen Type og Category)
        missing = [col for col, pos in required_columns.items() 
                  if pos is None and col not in ['Type', 'Category']]
        if missing:
            self.logger.error(f"Manglende kolonner: {missing}")
            return None
            
        return {k: v for k, v in required_columns.items() if v is not None}
        
    def _validate_part_number(self, part_number: str) -> bool:
        """Validerer et part number baseret på kategoriseringsreglerne"""
        if not part_number:
            return False
            
        # Brug categorizer til at validere
        type_, category = self.categorizer.categorize(part_number)
        return type_ != "Unknown" and category != "Unknown"
        
    def _insert_arrangement_row(self, file_info: dict):
        """Indsætter arrangement række i række 2"""
        self.logger.debug("Indsætter arrangement række")
        
        # Indsæt ny række i position 2
        self.sheet.Rows(2).Insert()
        
        # Bestem beskrivelse baseret på type
        description = "Arrangement Drawing" if file_info['type'] == 'arrangement' else "Basic Equipment Drawing"
        
        # Udfyld data i række 2
        data = {
            'Item': '0',  # Sæt Item til 0
            'Part Number': file_info['part_number'],
            'REV': file_info['rev'],
            'BOM Structure': 'None',
            'Description': description,
            'QTY': '1',
            'D': '1',
            't': '1',
            'L': '1'
        }
        
        for col_name, value in data.items():
            if col_name in self.columns:
                self.sheet.Cells(2, self.columns[col_name]).Value = value
                
        self.logger.info(f"Indsat arrangement række med part number {file_info['part_number']} og beskrivelse {description}")
        
    def _categorize_all_rows(self):
        """Kategoriserer alle part numbers i arket"""
        last_row = self.sheet.UsedRange.Rows.Count
        part_number_col = self.columns['Part Number']
        type_col = self.columns['Type']
        category_col = self.columns['Category']
        
        for row in range(2, last_row + 1):
            part_number = str(self.sheet.Cells(row, part_number_col).Value).strip()
            if part_number:
                type_, category = self.categorizer.categorize(part_number)
                self.sheet.Cells(row, type_col).Value = type_
                self.sheet.Cells(row, category_col).Value = category

    def _handle_supplier_parts(self):
        """Håndterer supplier parts baseret på checkbox"""
        if not self.include_suppliers:
            self.logger.info("Fjerner supplier parts...")
            part_number_col = self.columns['Part Number']
            last_row = self.sheet.UsedRange.Rows.Count
            rows_to_delete = []
            
            # Find rækker der skal slettes
            for row in range(last_row, 1, -1):
                part_number = str(self.sheet.Cells(row, part_number_col).Value).strip()
                if part_number.startswith(('0000-700-', '0000-701-', '0000-702-')):
                    rows_to_delete.append(row)
                    
            # Slet rækker (baglæns for at undgå forskydning)
            for row in sorted(rows_to_delete, reverse=True):
                self.sheet.Rows(row).Delete()
                
            self.logger.info(f"Fjernet {len(rows_to_delete)} supplier parts rækker")

    def _handle_invalid_rows(self):
        """
        Håndterer ugyldige part numbers ved at flytte dem til Other Parts fanen
        """
        part_number_col = self.columns['Part Number']
        type_col = self.columns['Type']
        category_col = self.columns['Category']
        last_row = self.sheet.UsedRange.Rows.Count
        invalid_rows = []
        revision_updates = []
        
        # Find rækker med ugyldige part numbers og part numbers med revision
        for row in range(2, last_row + 1):
            part_number = str(self.sheet.Cells(row, part_number_col).Value).strip()
            if not part_number:
                continue
                
            # Tjek om part number er gyldigt ved at prøve at kategorisere det
            type_, category = self.categorizer.categorize(part_number)
            
            if type_ == "Unknown" or category == "Unknown":
                invalid_rows.append(row)
            else:
                # Brug categorizer til at parse part number
                parsed = self.categorizer._parse_part_number(part_number)
                if parsed and parsed.get('revision'):
                    # Gem original part number uden revision og revision separat
                    revision_updates.append((row, parsed['base'], parsed['revision']))
        
        # Håndter revision opdateringer først
        if revision_updates:
            for row, new_part_number, rev_letter in revision_updates:
                # Opdater part number uden revision
                self.sheet.Cells(row, part_number_col).Value = new_part_number
                # Opdater REV kolonne med revisionen
                self.sheet.Cells(row, rev_col).Value = rev_letter
            self.logger.info(f"Opdateret {len(revision_updates)} part numbers med revision information")
                
        # Marker ugyldige rækker som Other Parts
        for row in invalid_rows:
            self.sheet.Cells(row, type_col).Value = "Other Parts"
            self.sheet.Cells(row, category_col).Value = "Other Parts"
        
        if invalid_rows:
            self.logger.info(f"Kategoriseret {len(invalid_rows)} rækker som Other Parts")

    def _create_excel_groups(self):
        """
        Opretter Excel grupper baseret på hierarkiet i Item kolonnen.
        Konverteret fra VB til Python.
        """
        self.logger.info("Opretter Excel grupper...")
        
        item_col = self.columns['Item']
        last_row = self.sheet.UsedRange.Rows.Count
        
        # Opbyg gruppe map
        group_map = []  # Liste af tupler (celle_adresse, niveau)
        for row in range(2, last_row + 1):
            item_number = str(self.sheet.Cells(row, item_col).Value).strip()
            if item_number:
                # Bestem niveau baseret på antal punktummer
                level = len(item_number.split('.')) - 1
                cell = self.sheet.Cells(row, item_col)
                group_map.append((cell, level))
        
        # Find max niveau
        max_level = max(level for _, level in group_map) if group_map else 0
        
        # Fjern eksisterende grupper
        try:
            for _ in range(10):  # Antag max 10 niveauer
                self.sheet.Range(f"A2:A{last_row}").EntireRow.Ungroup()
        except:
            pass  # Ignorer fejl hvis ingen grupper findes
        
        # Opret grupper niveau for niveau (start med dybeste)
        for current_level in range(max_level, 0, -1):
            start_group = None
            last_group = None
            
            for cell, level in group_map:
                if level >= current_level:
                    # Udvid gruppe
                    if start_group is None:
                        start_group = cell
                    last_group = cell
                else:
                    # Afslut gruppe hvis vi har en
                    if start_group and last_group:
                        self.sheet.Range(start_group, last_group).EntireRow.Group()
                    start_group = None
                    last_group = None
            
            # Håndter sidste gruppe i niveau hvis nødvendigt
            if start_group and last_group:
                self.sheet.Range(start_group, last_group).EntireRow.Group()
        
        self.logger.info("Excel grupper oprettet")

    def _create_category_sheets(self):
        """
        Opretter faner for hver kategori og flytter rækker til de respektive faner.
        """
        self.logger.info("Opretter kategori-faner...")
        
        # Find nødvendige kolonner
        item_col = self.columns['Item']
        category_col = self.columns['Category']
        
        # Find sidste række
        last_row = self.sheet.UsedRange.Rows.Count
        
        # Opret dictionary til at holde styr på kategorier og deres rækker
        category_rows = {}
        
        # Find alle kategorier og deres rækker
        for row in range(2, last_row + 1):
            category = str(self.sheet.Cells(row, category_col).Value).strip()
            if not category:
                continue
                
            # Find alle child rækker hvis dette er en parent
            rows_to_move = {row}  # Start med current row
            
            # Find alle child rækker baseret på item number
            item_number = str(self.sheet.Cells(row, item_col).Value).strip()
            if item_number:
                parent_parts = item_number.split('.')
                for child_row in range(row + 1, last_row + 1):
                    child_item = str(self.sheet.Cells(child_row, item_col).Value).strip()
                    child_parts = child_item.split('.')
                    
                    # Hvis child_item starter med parent_item, er det en child
                    if (len(child_parts) > len(parent_parts) and 
                        '.'.join(child_parts[:len(parent_parts)]) == '.'.join(parent_parts)):
                        rows_to_move.add(child_row)
                    elif len(child_parts) <= len(parent_parts):
                        break
            
            # Tilføj rækker til kategori
            if category not in category_rows:
                category_rows[category] = set()
            category_rows[category].update(rows_to_move)
        
        # Opret faner og kopier rækker
        header_row = None
        for category, rows in category_rows.items():
            if not rows:
                continue
                
            # Opret ny fane
            try:
                sheet = self.workbook.Worksheets(category)
                sheet.Delete()  # Slet eksisterende fane hvis den findes
            except:
                pass
            
            new_sheet = self.workbook.Worksheets.Add()
            new_sheet.Name = category
            
            # Kopier header row hvis vi ikke har gemt den endnu
            if header_row is None:
                header_row = self.sheet.Rows(1).Copy()
                new_sheet.Paste(new_sheet.Range("A1"))
            else:
                self.sheet.Rows(1).Copy()
                new_sheet.Paste(new_sheet.Range("A1"))
            
            # Kopier rækker til ny fane
            current_row = 2
            for row in sorted(rows):
                self.sheet.Rows(row).Copy()
                new_sheet.Paste(new_sheet.Range(f"A{current_row}"))
                current_row += 1
            
            self.logger.info(f"Oprettet fane '{category}' med {len(rows)} rækker")
            
        self.logger.info("Kategori-faner oprettet")

    def _update_item_numbers(self):
        """Opdaterer item numre baseret på hierarkiet"""
        self.logger.info("Opdaterer item numre...")
        
        item_col = self.columns['Item']
        last_row = self.sheet.UsedRange.Rows.Count
        
        # Gennemgå alle rækker og behold det eksisterende hierarki
        for row in range(2, last_row + 1):
            cell = self.sheet.Cells(row, item_col)
            item_number = str(cell.Value).strip()
            
            # Konverter eventuelle kommaer til punktummer
            if ',' in item_number:
                item_number = item_number.replace(',', '.')
            
            # Fjern .0 fra alle niveauer
            parts = item_number.split('.')
            if parts[-1] == '0':
                parts = parts[:-1]
            item_number = '.'.join(parts)
            
            # Opdater celle og sæt NumberFormat til Text for at undgå auto-formatering
            cell.NumberFormat = "@"
            cell.Value = item_number
        
        # Opret Excel grupper baseret på hierarkiet
        self._create_excel_groups()
            
        self.logger.info("Item numre opdateret")

    def _handle_bom_structure(self):
        """
        Håndterer BOM Structure regler:
        - Sletter child rækker hvis parent er Inseparable eller part number starter med 0000-3
        - Sletter rækker markeret som Phantom
        """
        self.logger.info("Håndterer BOM Structure...")
        
        structure_col = self.columns['BOM Structure']
        part_number_col = self.columns['Part Number']
        item_col = self.columns['Item']
        last_row = self.sheet.UsedRange.Rows.Count
        rows_to_delete = set()  # Brug set for at undgå dubletter
        
        # Gennemgå rækker oppefra og ned
        for row in range(2, last_row + 1):
            structure = str(self.sheet.Cells(row, structure_col).Value).strip()
            part_number = str(self.sheet.Cells(row, part_number_col).Value).strip()
            item_number = str(self.sheet.Cells(row, item_col).Value).strip()
            
            # Tjek for Phantom
            if 'Phantom' in structure:
                rows_to_delete.add(row)
                # Find og slet alle child rækker til denne Phantom
                parent_parts = item_number.split('.')
                for next_row in range(row + 1, last_row + 1):
                    next_item = str(self.sheet.Cells(next_row, item_col).Value).strip()
                    next_parts = next_item.split('.')
                    if (len(next_parts) > len(parent_parts) and 
                        '.'.join(next_parts[:len(parent_parts)]) == '.'.join(parent_parts)):
                        rows_to_delete.add(next_row)
                    elif len(next_parts) <= len(parent_parts):
                        break
                continue
            
            # Tjek for Inseparable eller 0000-3
            if structure == 'Inseparable' or (part_number and part_number.startswith('0000-3')):  # Ændret til eksakt match
                # Find alle child rækker ved at sammenligne item numre
                parent_parts = item_number.split('.')
                for next_row in range(row + 1, last_row + 1):
                    next_item = str(self.sheet.Cells(next_row, item_col).Value).strip()
                    next_parts = next_item.split('.')
                    
                    # Hvis næste række har flere niveauer og starter med parent's nummer, er det en child
                    if (len(next_parts) > len(parent_parts) and 
                        '.'.join(next_parts[:len(parent_parts)]) == '.'.join(parent_parts)):
                        rows_to_delete.add(next_row)
                    elif len(next_parts) <= len(parent_parts):
                        # Vi har nået næste parent eller sibling, stop søgningen
                        break
        
        # Slet rækker (baglæns for at undgå forskydning)
        deleted_count = 0
        for row in sorted(rows_to_delete, reverse=True):
            self.sheet.Rows(row).Delete()
            deleted_count += 1
        
        self.logger.info(f"Slettet {deleted_count} rækker baseret på BOM Structure")

    def _calculate_total_qty(self):
        """
        Beregner Total QTY ved at gange parent QTY med child QTY.
        Indsætter en ny kolonne 'Total QTY' efter 'QTY' kolonnen.
        """
        self.logger.info("Beregner Total QTY...")
        
        # Find QTY kolonnen
        qty_col = self.columns['QTY']
        last_row = self.sheet.UsedRange.Rows.Count
        
        # Indsæt ny Total QTY kolonne efter QTY
        self.sheet.Columns(qty_col + 1).Insert()
        self.sheet.Cells(1, qty_col + 1).Value = "Total QTY"
        
        # Opbyg dictionary med parent QTY værdier
        parent_qty = {}  # Key: item_number, Value: total_qty
        item_col = self.columns['Item']
        
        # Første gennemløb: Gem alle QTY værdier
        for row in range(2, last_row + 1):
            item_number = str(self.sheet.Cells(row, item_col).Value).strip()
            qty = self.sheet.Cells(row, qty_col).Value
            if qty is None:
                qty = 1  # Default værdi hvis QTY er tom
            try:
                qty = float(qty)
            except (ValueError, TypeError):
                qty = 1  # Default værdi hvis QTY ikke er et tal
                
            parent_qty[item_number] = qty
            
        # Andet gennemløb: Beregn Total QTY
        for row in range(2, last_row + 1):
            item_number = str(self.sheet.Cells(row, item_col).Value).strip()
            if not item_number:
                continue
                
            # Split item number for at finde alle parents
            parts = item_number.split('.')
            total_qty = 1
            
            # Multiplicer med alle parent QTY værdier
            for i in range(len(parts)):
                parent = '.'.join(parts[:i+1])
                if parent in parent_qty:
                    total_qty *= parent_qty[parent]
                    
            # Indsæt Total QTY i den nye kolonne
            self.sheet.Cells(row, qty_col + 1).Value = total_qty
            
        self.logger.info("Total QTY beregnet og indsat")

    def _handle_drawings(self, source_path: str):
        """
        Håndterer tegninger:
        1. Finder og kopierer seneste PDF/DWG filer
        2. Opdaterer REV baseret på tegningsrevision
        3. Tilføjer Drawing kolonne med status (DWG_PDF hvis samme revision, DWG~PDF hvis forskellige revisioner)
        """
        self.logger.info("Håndterer tegninger...")
        
        # Find nødvendige kolonner
        part_number_col = self.columns['Part Number']
        rev_col = self.columns['REV']
        category_col = self.columns['Category']
        
        # Tilføj Drawing kolonne
        last_col = self.sheet.UsedRange.Columns.Count
        self.sheet.Cells(1, last_col + 1).Value = "Drawing"
        drawing_col = last_col + 1
        
        # Scan kildekataloget for PDF og DWG filer
        all_files = []
        for root, _, files in os.walk(source_path):
            for file in files:
                if file.lower().endswith(('.pdf', '.dwg')):
                    all_files.append(os.path.join(root, file))
        
        self.logger.info(f"Fundet {len(all_files)} tegninger")
        
        # Opret dictionary til at holde styr på oprettede mapper
        created_folders = set()
        
        # Gennemgå hver række i Excel
        last_row = self.sheet.UsedRange.Rows.Count
        for row in range(2, last_row + 1):
            part_number = str(self.sheet.Cells(row, part_number_col).Value).strip()
            if not part_number:
                continue
                
            # Find alle filer der matcher dette part number
            matching_files = {
                'pdf': [],
                'dwg': []
            }
            
            for file in all_files:
                filename = os.path.basename(file)
                if filename.startswith(part_number):
                    ext = os.path.splitext(filename)[1].lower()
                    if ext in ['.pdf', '.dwg']:
                        matching_files[ext[1:]].append(file)
            
            # Find seneste revision og filer
            latest_rev = None
            latest_files = {}
            latest_revs = {'pdf': None, 'dwg': None}
            
            for ext in ['pdf', 'dwg']:
                if matching_files[ext]:
                    # Find fil med seneste revision
                    latest = None
                    latest_rev_file = None
                    
                    for file in matching_files[ext]:
                        filename = os.path.basename(file)
                        match = re.match(r'^(.*)-(.)\s*-\s*', filename)
                        if match:
                            rev = match.group(2)  # Det sidste tegn (bindestreg eller bogstav)
                            if latest_rev_file is None or rev > latest_rev_file:
                                latest_rev_file = rev
                                latest = file
                                latest_revs[ext] = rev
                    
                    if latest:
                        latest_files[ext] = latest
                        if latest_rev is None or latest_rev_file > latest_rev:
                            latest_rev = latest_rev_file
            
            # Opdater REV hvis vi fandt en nyere revision
            if latest_rev:
                current_rev = str(self.sheet.Cells(row, rev_col).Value).strip()
                if latest_rev != current_rev:
                    self.sheet.Cells(row, rev_col).Value = latest_rev
            
            # Bestem Drawing kolonne værdi
            has_pdf = 'pdf' in latest_files
            has_dwg = 'dwg' in latest_files
            
            if has_pdf and has_dwg:
                # Sammenlign revisioner
                if latest_revs['pdf'] == latest_revs['dwg']:
                    value = "DWG_PDF"
                else:
                    value = "DWG~PDF"
            elif has_dwg:
                value = "DWG"
            elif has_pdf:
                value = "PDF"
            else:
                value = ""
            
            self.sheet.Cells(row, drawing_col).Value = value
            
            # Kopier filer til korrekt kategorimappe
            if latest_files:
                # Brug Category til mappenavn
                category = str(self.sheet.Cells(row, category_col).Value).strip()
                
                if category:
                    # Opret destinationsmappe
                    dest_folder = os.path.join(os.path.dirname(self.workbook.FullName), category)
                    if category not in created_folders:
                        os.makedirs(dest_folder, exist_ok=True)
                        created_folders.add(category)
                    
                    # Kopier filer
                    for ext, src_file in latest_files.items():
                        dest_file = os.path.join(dest_folder, os.path.basename(src_file))
                        shutil.copy2(src_file, dest_file)
                        self.logger.debug(f"Kopieret {os.path.basename(src_file)} til {category}")
        
        self.logger.info("Tegningshåndtering afsluttet")

    def process_file(self) -> bool:
        """
        Udfører TRIN 2: Data Indlæsning og Validering
        """
        try:
            self._optimize_excel_settings()
            
            # Omdøb aktivt sheet til "BOM (Raw)"
            self.sheet.Name = "BOM (Raw)"
            
            # Identificer kolonner
            self.columns = self._get_column_indices()
            if not self.columns:
                return False
                
            # Udtræk information fra filnavn og indsæt arrangement række
            file_info = self._extract_file_info()
            if not file_info:
                self.logger.error("Kunne ikke udtrække information fra filnavn")
                return False
                
            self._insert_arrangement_row(file_info)
            self.logger.info(f"Indsat arrangement række med part number {file_info['part_number']}")
                
            # Håndter supplier parts og ugyldige rækker
            self._handle_supplier_parts()
            self._handle_invalid_rows()
            
            # Opdater item numre for at sikre korrekt hierarki
            self._update_item_numbers()
            
            # Håndter BOM Structure regler EFTER item numre er opdateret
            self._handle_bom_structure()
            
            # Opdater item numre igen efter struktur ændringer
            self._update_item_numbers()
                
            # Kategoriser alle rækker
            self._categorize_all_rows()
            
            # Beregn Total QTY
            self._calculate_total_qty()
            
            # Opret kategori-faner og flyt rækker
            self._create_category_sheets()
            
            # Håndter tegninger med vores nye metode
            source_path = os.path.join(os.path.dirname(os.path.dirname(__file__)), "Files")  # Test sti
            try:
                self._handle_drawings(source_path)
            except Exception as e:
                self.logger.warning(f"Kunne ikke håndtere tegninger: {str(e)}")
                
            return True
            
        except Exception as e:
            self.logger.error(f"Fejl under indlæsning: {str(e)}", exc_info=True)
            return False
            
        finally:
            self._restore_excel_settings() 