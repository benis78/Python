"""
Excel data processor modul
Håndterer al Excel data behandling og validering
"""

import logging
import re
from datetime import datetime
import win32com.client
import os
import shutil

class ExcelDataProcessor:
    def __init__(self, workbook):
        """Initialiserer data processoren"""
        self.logger = logging.getLogger('ExcelCopyBOM.DataProcessor')
        self.logger.info("ExcelDataProcessor initialiseret")
        
        self.workbook = workbook
        self.sheet = workbook.ActiveSheet
        
        # Valider workbook reference
        self.logger.debug(f"Workbook reference gyldig: {self.workbook is not None}")
        
        # Gem oprindelige Excel indstillinger
        self.excel = workbook.Application
        self.original_calculation = self.excel.Calculation
        self.original_screen_updating = self.excel.ScreenUpdating
        self.original_display_alerts = self.excel.DisplayAlerts
        
        # Regex patterns for part numbers
        self.part_number_patterns = {
            'project': r'^[2-9]\d{3}',  # 2000-9999
            'sales': r'^[A-Za-z]{2,4}',  # Starter med 2-4 bogstaver
            'basic': r'^0000'  # Starter med 0000
        }
        
    def _optimize_excel_settings(self):
        """Optimerer Excel indstillinger for bedre performance"""
        self.excel.Calculation = -4135  # xlCalculationManual
        self.excel.ScreenUpdating = False
        self.excel.DisplayAlerts = False
        self.logger.debug("Excel indstillinger optimeret for performance")
        
    def _restore_excel_settings(self):
        """Gendanner oprindelige Excel indstillinger"""
        self.excel.Calculation = self.original_calculation
        self.excel.ScreenUpdating = self.original_screen_updating
        self.excel.DisplayAlerts = self.original_display_alerts
        self.logger.debug("Excel indstillinger gendannet")
        
    def _rename_active_sheet(self, new_name: str, max_attempts: int = 3):
        """Forsøger at omdøbe det aktive sheet"""
        self.logger.debug(f"Aktivt sheet: {self.sheet.Name}")
        
        for attempt in range(max_attempts):
            try:
                self.sheet.Name = new_name
                return True
            except Exception as e:
                self.logger.warning(f"Forsøg {attempt + 1} fejlede ved omdøbning af sheet: {str(e)}")
                if attempt < max_attempts - 1:
                    # Prøv med et nummer efter navnet
                    try:
                        self.sheet.Name = f"{new_name}_{attempt + 1}"
                        return True
                    except Exception as e2:
                        continue
                        
        self.logger.error(f"Kunne ikke omdøbe sheet efter {max_attempts} forsøg: {str(e)}")
        return False
        
    def _identify_columns(self):
        """Identificerer kolonnerne i Excel arket"""
        required_columns = {
            'Item': None,
            'Part Number': None,
            'REV': None,
            'Description': None,
            'QTY': None,
            'BOM Structure': None,
            'D': None,
            't': None,
            'L': None
        }
        
        # Find sidste kolonne med data
        last_col = self.sheet.UsedRange.Columns.Count
        self.logger.debug(f"Scanner {last_col} kolonner")
        
        # Scan første række for kolonnenavne
        for col in range(1, last_col + 1):
            cell_value = self.sheet.Cells(1, col).Value
            if cell_value:
                self.logger.debug(f"Kolonne {col}: '{cell_value}'")
                if cell_value in required_columns:
                    required_columns[cell_value] = col
                    
        # Check for manglende kolonner
        missing_columns = [col for col, pos in required_columns.items() if pos is None]
        if missing_columns:
            self.logger.warning(f"Manglende kolonner: {missing_columns}")
            
        # Fjern None værdier
        self.columns = {k: v for k, v in required_columns.items() if v is not None}
        self.logger.debug(f"Fandt følgende kolonner: {self.columns}")
        
        return self.columns
        
    def _format_worksheet(self, sheet=None):
        """Formaterer worksheet med standard indstillinger"""
        if sheet is None:
            sheet = self.sheet
            
        self.logger.debug(f"Starter formatering af worksheet: {sheet.Name}")
        
        # Indstil række højde
        self.logger.debug("Indstiller række højder")
        sheet.Rows(1).RowHeight = 20  # Header række (20 pixels)
        sheet.Rows.RowHeight = 91     # Alle andre rækker (91 pixels)
        
        # Definer præcise kolonnebredder
        self.logger.debug("Indstiller kolonnebredder")
        column_widths = {
            'Item': 52/7,           # 52 pixels
            'Part Number': 152/7,    # 152 pixels
            'REV': 47/7,            # 47 pixels
            'Description 1': 111/7,  # 111 pixels
            'Description 2': 93/7,   # 93 pixels
            'BOM Structure': 115/7,  # 115 pixels
            'Description': 423/7,    # 423 pixels
            'Material': 135/7,       # 135 pixels
            'Standard/PED': 173/7,   # 173 pixels
            'QTY': 48/7,            # 48 pixels
            'Total QTY': 82/7,      # 82 pixels
            'Weight': 39/7,         # 39 pixels
            'Surface Area': 39/7,    # 39 pixels
            'Volume': 39/7,         # 39 pixels
            'Comment': 200/7,       # 200 pixels
            'Drawings': 94/7        # 94 pixels
        }
        
        # Find kolonnepositioner i det aktuelle sheet
        col_positions = {}
        for i in range(1, sheet.UsedRange.Columns.Count + 1):
            header = str(sheet.Cells(1, i).Value).strip()
            if header in column_widths:
                col_positions[header] = i
                
        # Anvend bredder (divideret med 7 for at konvertere pixels til Excel enheder)
        for col_name, width in column_widths.items():
            if col_name in col_positions:
                sheet.Columns(col_positions[col_name]).ColumnWidth = width
                
        # Tilføj filter og frys første række
        self.logger.debug("Tilføjer filter og frys")
        sheet.Rows(1).AutoFilter()
        sheet.Rows(2).Select()
        self.excel.ActiveWindow.FreezePanes = True
        
        # Hvis det er et piping ark, gruppér efter parent/child hierarki
        if "Piping" in sheet.Name:
            self._group_piping_rows(sheet)
            
        self.logger.debug("Worksheet formatering gennemført")
        
    def _add_missing_columns(self):
        """Tilføjer manglende kolonner"""
        last_col = self.sheet.UsedRange.Columns.Count
        
        # Tilføj Drawings kolonne hvis den ikke findes
        if 'Drawings' not in self.columns:
            last_col += 1
            self.logger.debug(f"Tilføjer Drawings kolonne i position {last_col}")
            self.sheet.Cells(1, last_col).Value = "Drawings"
            self.columns['Drawings'] = last_col
            
        self.logger.debug("Manglende kolonner tilføjet")
        
    def _handle_equipment_rows(self, include_equipment: bool):
        """Håndterer equipment rækker baseret på bruger valg"""
        self.logger.debug("Starter håndtering af equipment rækker")
        
        if not include_equipment:
            rows_to_delete = []
            last_row = self.sheet.UsedRange.Rows.Count
            
            # Find rækker der starter med 0000-7
            for row in range(2, last_row + 1):
                part_number = str(self.sheet.Cells(row, self.columns['Part Number']).Value)
                if part_number and part_number.startswith("0000-7"):
                    rows_to_delete.append(row)
                    
            # Slet rækker (bagfra for at undgå forskydning)
            for row in sorted(rows_to_delete, reverse=True):
                self.sheet.Rows(row).Delete()
                
            self.logger.info(f"Håndtering af equipment rækker afsluttet, {len(rows_to_delete)} rækker fjernet")
            
        self.logger.debug("Equipment rækker håndteret")
        
    def _extract_file_info(self, filename: str) -> dict:
        """Udtrækker information fra filnavnet"""
        self.logger.debug(f"Udtrækker information fra filnavn: {filename}")
        
        # Fjern filendelse og eventuel sti
        base_name = os.path.splitext(os.path.basename(filename))[0]
        
        # Find part number og revision
        match = re.match(r'^(\d{4})-(\d{2}(?:\.\d)?)-([A-Za-z]\d{2})', base_name)
        if match:
            return {
                'part_number': f"{match.group(1)}-{match.group(2)}",
                'rev': match.group(3),
                'type': 'arrangement' if len(match.group(2).split('.')) > 1 else 'basic'
            }
        
        self.logger.warning(f"Kunne ikke udtrække information fra filnavn: {filename}")
        return None
        
    def _insert_arrangement_row(self, file_info: dict):
        """Indsætter arrangement række i toppen af BOM"""
        self.logger.debug("Indsætter arrangement række")
        
        # Indsæt ny række
        self.sheet.Rows(2).Insert()
        
        # Udfyld data
        col_data = {
            'Item': '0',
            'Part Number': file_info['part_number'],
            'REV': file_info['rev'],
            'BOM Structure': 'Inseparable',
            'Description': 'Arrangement Drawing' if file_info['type'] == 'arrangement' else 'Basic Equipment Drawing',
            'QTY': '1',
            'D': '1',
            't': '1',
            'L': '1'
        }
        
        for col_name, value in col_data.items():
            if col_name in self.columns:
                self.sheet.Cells(2, self.columns[col_name]).Value = value
                
    def _handle_special_rows(self):
        """Håndterer specielle rækker (kun bogstaver)"""
        self.logger.debug("Håndterer specielle rækker")
        
        special_rows = []
        last_row = self.sheet.UsedRange.Rows.Count
        
        # Find rækker med kun bogstaver i Part Number
        for row in range(2, last_row + 1):
            part_number = str(self.sheet.Cells(row, self.columns['Part Number']).Value)
            if part_number and part_number.isalpha():
                special_rows.append({
                    'row': row,
                    'part_number': part_number,
                    'description': str(self.sheet.Cells(row, self.columns['Description']).Value)
                })
                
        if special_rows:
            # TODO: Implementer GUI dialog til at håndtere disse rækker
            self.logger.warning(f"Fandt {len(special_rows)} rækker med kun bogstaver")
            
    def _index_rows(self):
        """Indekserer rækker efter parent/child hierarki"""
        self.logger.debug("Indekserer rækker")
        
        last_row = self.sheet.UsedRange.Rows.Count
        structure_col = self.columns['BOM Structure']
        item_col = self.columns['Item']
        
        # Initialiser variabler til at holde styr på hierarkiet
        current_level = 0
        index_stack = ['0']
        
        for row in range(2, last_row + 1):
            structure = str(self.sheet.Cells(row, structure_col).Value).lower()
            
            # Bestem niveau baseret på BOM Structure
            if structure == 'inseparable':
                current_level = 0
                index_stack = index_stack[:1]
            elif structure == 'phantom':
                current_level += 1
                if len(index_stack) > current_level:
                    index_stack = index_stack[:current_level]
                last_index = int(index_stack[-1].split('.')[-1])
                index_stack.append(f"{index_stack[-1]}.{last_index + 1}")
            else:
                if current_level > 0:
                    last_parts = index_stack[-1].split('.')
                    last_parts[-1] = str(int(last_parts[-1]) + 1)
                    index_stack[-1] = '.'.join(last_parts)
                else:
                    index_stack = [str(int(index_stack[0]) + 1)]
                    
            # Opdater Item nummer
            self.sheet.Cells(row, item_col).Value = index_stack[-1]
            
    def _handle_bom_structure(self):
        """Håndterer BOM Structure regler"""
        self.logger.debug("Håndterer BOM Structure")
        
        last_row = self.sheet.UsedRange.Rows.Count
        rows_to_delete = []
        
        for row in range(2, last_row + 1):
            structure = str(self.sheet.Cells(row, self.columns['BOM Structure']).Value).lower()
            part_number = str(self.sheet.Cells(row, self.columns['Part Number']).Value)
            
            # Slet phantom rækker
            if structure == 'phantom':
                rows_to_delete.append(row)
                continue
                
            # Slet children hvis Inseparable eller starter med 0000-3
            if structure == 'inseparable' or (part_number and part_number.startswith('0000-3')):
                # Find og marker alle children til denne parent
                current_item = str(self.sheet.Cells(row, self.columns['Item']).Value)
                for child_row in range(row + 1, last_row + 1):
                    child_item = str(self.sheet.Cells(child_row, self.columns['Item']).Value)
                    if child_item.startswith(current_item + '.'):
                        rows_to_delete.append(child_row)
                        
        # Slet markerede rækker (bagfra)
        for row in sorted(rows_to_delete, reverse=True):
            self.sheet.Rows(row).Delete()
            
    def _extract_revision(self, part_number: str) -> tuple:
        """Udtrækker revision fra part number"""
        match = re.search(r'(.+?)(-[A-Z]\d{2})?$', part_number)
        if match and match.group(2):
            return match.group(1), match.group(2)[1:]  # Returner base part number og revision
        return part_number, None
        
    def _move_revisions(self):
        """Flytter revisionsnumre fra Part Number til REV kolonne"""
        self.logger.debug("Flytter revisionsnumre")
        
        last_row = self.sheet.UsedRange.Rows.Count
        
        for row in range(2, last_row + 1):
            part_number = str(self.sheet.Cells(row, self.columns['Part Number']).Value)
            base_part, revision = self._extract_revision(part_number)
            
            if revision:
                # Opdater Part Number og REV
                self.sheet.Cells(row, self.columns['Part Number']).Value = base_part
                current_rev = self.sheet.Cells(row, self.columns['REV']).Value
                if not current_rev:  # Kun opdater hvis REV er tom
                    self.sheet.Cells(row, self.columns['REV']).Value = revision
                    
    def _calculate_total_qty(self):
        """Beregner Total QTY baseret på parent QTY"""
        self.logger.debug("Beregner Total QTY")
        
        # Tilføj Total QTY kolonne
        last_col = max(self.columns.values()) + 1
        self.sheet.Cells(1, last_col).Value = "Total QTY"
        self.columns['Total QTY'] = last_col
        
        last_row = self.sheet.UsedRange.Rows.Count
        
        for row in range(2, last_row + 1):
            item = str(self.sheet.Cells(row, self.columns['Item']).Value)
            qty = float(self.sheet.Cells(row, self.columns['QTY']).Value or 0)
            
            # Find parent QTY
            if '.' in item:
                parent_item = '.'.join(item.split('.')[:-1])
                for parent_row in range(2, last_row + 1):
                    if str(self.sheet.Cells(parent_row, self.columns['Item']).Value) == parent_item:
                        parent_qty = float(self.sheet.Cells(parent_row, self.columns['QTY']).Value or 0)
                        qty *= parent_qty
                        break
                        
            self.sheet.Cells(row, last_col).Value = qty
            
    def _load_categories(self):
        """Indlæser kategorier fra Categories.csv"""
        self.logger.debug("Indlæser kategorier fra CSV")
        
        categories = []
        csv_path = os.path.join(os.path.dirname(__file__), "Categories.csv")
        
        try:
            with open(csv_path, 'r', encoding='utf-8') as f:
                for line in f:
                    parts = line.strip().split(';')
                    if len(parts) >= 4:
                        pattern = {
                            'group1': parts[0].strip(),
                            'group2': parts[1].strip(),
                            'group3': parts[2].strip() if len(parts) > 2 else '',
                            'category': parts[3].strip(),
                            'type': parts[4].strip() if len(parts) > 4 else ''
                        }
                        categories.append(pattern)
                        
            self.logger.debug(f"Indlæst {len(categories)} kategorier")
            return categories
        except Exception as e:
            self.logger.error(f"Fejl ved indlæsning af kategorier: {str(e)}")
            return []

    def _match_pattern(self, part_number: str, pattern: dict) -> bool:
        """
        Tjekker om et part number matcher et mønster baseret på reglerne fra Kategori Rækkefølge.txt
        
        Et Part Number består af grupper adskilt af bindestreger:
        Gruppe 1: Altid 4 alfanumeriske tegn
        Gruppe 2: 2/4 alfanumeriske tegn (Area), 3 tal startende med 3 (Basic), 3 tal startende med 6 (Equipment)
        Gruppe 3: Varierer baseret på tidligere grupper
        """
        # Split part number i grupper
        groups = part_number.split('-')
        if len(groups) < 2:
            return False
            
        # Gruppe 1 (første 4 tegn)
        if len(groups[0]) != 4:  # Skal altid være 4 tegn
            return False
            
        if pattern['group1'] == '****':
            # Sales nummer (2-4 bogstaver først)
            if not (groups[0][:2].isalpha() or groups[0][:4].isalpha()):
                return False
        elif pattern['group1'] == '2000 - 9999':
            # Projekt nummer
            if not (groups[0].isdigit() and 2000 <= int(groups[0]) <= 9999):
                return False
        elif pattern['group1'] == '0000':
            # Suppliers Parts eller Basic Components
            if groups[0] != '0000':
                return False
                
        # Gruppe 2
        if len(groups) > 1:
            if pattern['group2'] == '****':
                # Area Drawings: 2 eller 4 alfanumeriske tegn
                base_group2 = groups[1].split('.')[0]  # Håndter decimaler (f.eks. 02.1)
                if not (len(base_group2) in [2, 4]):
                    return False
            elif pattern['group2'].startswith('3'):
                # Basic Components: 3 tal startende med 3
                if not (len(groups[1]) == 3 and groups[1].isdigit() and groups[1].startswith('3')):
                    return False
            elif pattern['group2'].startswith('6'):
                # Basic Equipment: 3 tal startende med 6
                if not (len(groups[1]) == 3 and groups[1].isdigit() and groups[1].startswith('6')):
                    return False
            else:
                # Specifikt nummer (f.eks. 615 for Primary Digester)
                if groups[1] != pattern['group2']:
                    return False
                    
        # Gruppe 3
        if len(groups) > 2:
            # Hvis Gruppe 1 er "0000"
            if groups[0] == '0000':
                if groups[1].startswith('3'):
                    return pattern['category'] == 'Basic Components'
                elif groups[1].startswith('7'):
                    return pattern['category'] == 'Suppliers Parts'
                    
            # Hvis det er Basic Equipment Drawing
            if groups[1].startswith('6') and groups[2][0].isalpha():
                return pattern['category'] == 'Basic Equipment Drawing'
                
            # Area Drawings specifikke prefixes
            area_prefixes = {
                'A': 'Arrangement Drawing',
                'E': 'Equipment Drawing',
                'P': 'Sub Terrain Piping Plan',
                'S': 'Sub Terrain Piping Plan',
                'F': 'Foundation',
                'PS': 'Project Specific Parts',
                'B': 'Building'
            }
            
            # Piping kategorier
            piping_prefixes = {
                'BM': 'Biomass Piping',
                'AF': 'Anti Foam Piping',
                'AA': 'Atmospheric Air Piping',
                'BG': 'Biogas Piping',
                'CD': 'Cable Ducts',
                'CS': 'Condensate Piping',
                'EL': 'Power Cable Piping',
                'CO': 'Cooling Water Piping',
                'EZ': 'Enzyme Piping',
                'HW': 'Hot Water Piping',
                'HO': 'Hydraulic Oil Piping',
                'IC': 'Iron Chloride Piping',
                'OA': 'Odour Piping',
                'NT': 'Nutrient Piping',
                'OG': 'Offgas Piping',
                'PW': 'Potable Water Piping',
                'PA': 'Pressurized Air Piping',
                'RW': 'Rain Water Piping',
                'SL': 'Sulphurous liquid Piping',
                'TD': 'Technical Drainage Piping',
                'TW': 'Technical Water Piping'
            }
            
            # Tjek Area Drawing prefixes
            for prefix, category in area_prefixes.items():
                if groups[2].startswith(prefix):
                    return pattern['category'] == category
                    
            # Tjek Piping prefixes
            for prefix, category in piping_prefixes.items():
                if groups[2].startswith(prefix):
                    return pattern['category'] == category
                    
        return True

    def _categorize_part_numbers(self):
        """Kategoriserer part numbers og opretter faner"""
        self.logger.debug("Starter kategorisering af part numbers")
        
        # Indlæs kategorier
        categories = self._load_categories()
        if not categories:
            self.logger.error("Ingen kategorier fundet")
            return
            
        # Find Part Number kolonne
        part_number_col = self.columns.get('Part Number')
        if not part_number_col:
            self.logger.error("Part Number kolonne ikke fundet")
            return
            
        # Opret dictionary til at holde styr på kategoriserede rækker
        categorized_rows = {}
        
        # Scan alle rækker
        last_row = self.sheet.UsedRange.Rows.Count
        for row in range(2, last_row + 1):
            part_number = str(self.sheet.Cells(row, part_number_col).Value)
            if not part_number:
                continue
                
            # Find matching kategori baseret på prioritet
            matched = False
            
            # 1. Basic Components (0000-3xx)
            if part_number.startswith('0000-3'):
                category = "Basic Components"
                if category not in categorized_rows:
                    categorized_rows[category] = []
                categorized_rows[category].append(row)
                matched = True
                
            # 2. Piping kategorier
            elif not matched:
                for pattern in [p for p in categories if any(x in p['category'] for x in ['Piping', 'Cable Ducts'])]:
                    if self._match_pattern(part_number, pattern):
                        if pattern['category'] not in categorized_rows:
                            categorized_rows[pattern['category']] = []
                        categorized_rows[pattern['category']].append(row)
                        matched = True
                        break
                        
            # 3. Area Drawings
            elif not matched:
                for pattern in [p for p in categories if 'Area Drawings' in p['category']]:
                    if self._match_pattern(part_number, pattern):
                        if pattern['category'] not in categorized_rows:
                            categorized_rows[pattern['category']] = []
                        categorized_rows[pattern['category']].append(row)
                        matched = True
                        break
                        
            # 4. Tank Drawings
            elif not matched:
                for pattern in [p for p in categories if 'Tank' in p['category']]:
                    if self._match_pattern(part_number, pattern):
                        if pattern['category'] not in categorized_rows:
                            categorized_rows[pattern['category']] = []
                        categorized_rows[pattern['category']].append(row)
                        matched = True
                        break
                        
            # 5. Andre kategorier
            elif not matched:
                for pattern in categories:
                    if self._match_pattern(part_number, pattern):
                        if pattern['category'] not in categorized_rows:
                            categorized_rows[pattern['category']] = []
                        categorized_rows[pattern['category']].append(row)
                        matched = True
                        break
                        
            # Hvis stadig ingen match, tilføj til Other Parts
            if not matched:
                if "Other Parts" not in categorized_rows:
                    categorized_rows["Other Parts"] = []
                categorized_rows["Other Parts"].append(row)
                
        # Opret faner i alfabetisk rækkefølge
        for category, rows in sorted(categorized_rows.items()):
            if not rows:
                continue
                
            # Opret nyt ark
            new_sheet = self.workbook.Sheets.Add(After=self.workbook.Sheets(self.workbook.Sheets.Count))
            new_sheet.Name = category
            
            # Kopier header
            self.sheet.Range("1:1").Copy(new_sheet.Range("1:1"))
            
            # Kopier rækker
            for idx, row in enumerate(sorted(rows), start=2):
                self.sheet.Range(f"{row}:{row}").Copy(new_sheet.Range(f"{idx}:{idx}"))
                
            # Formater det nye ark
            self._format_worksheet(new_sheet)
            
        self.logger.info(f"Kategorisering gennemført. Oprettet {len(categorized_rows)} faner")
        
    def _group_piping_rows(self, sheet):
        """Grupperer rækker i et piping ark baseret på hierarki"""
        self.logger.debug(f"Grupperer rækker i {sheet.Name}")
        
        try:
            # Find Item kolonne
            item_col = None
            for i in range(1, sheet.UsedRange.Columns.Count + 1):
                if str(sheet.Cells(1, i).Value).strip().upper() == "ITEM":
                    item_col = i
                    break
                    
            if not item_col:
                self.logger.error("Item kolonne ikke fundet")
                return
                
            # Find alle parent items og deres children
            last_row = sheet.UsedRange.Rows.Count
            current_parent = None
            group_start = None
            
            for row in range(2, last_row + 1):
                item = str(sheet.Cells(row, item_col).Value)
                
                # Hvis item ikke indeholder punktum, er det en parent
                if '.' not in item:
                    # Afslut forrige gruppe hvis der er en
                    if group_start and current_parent and group_start < row:
                        sheet.Range(f"{group_start}:{row-1}").Rows.Group()
                    
                    current_parent = item
                    group_start = row + 1
                    
            # Gruppér sidste gruppe hvis der er en
            if group_start and current_parent and group_start < last_row:
                sheet.Range(f"{group_start}:{last_row}").Rows.Group()
                
            # Kollaps alle grupper
            sheet.Outline.ShowLevels(RowLevels=1)
            
        except Exception as e:
            self.logger.error(f"Fejl under gruppering: {str(e)}")

    def _create_partlist(self):
        """
        Opretter Partlist fane med samlede mængder for hver unik Part Number/REV kombination
        """
        self.logger.debug("Opretter Partlist fane")
        
        try:
            # Opret ny fane
            partlist_sheet = self.workbook.Sheets.Add(After=self.sheet)
            partlist_sheet.Name = "Partlist"
            
            # Find relevante kolonner
            part_number_col = self.columns.get('Part Number')
            rev_col = self.columns.get('REV')
            qty_col = self.columns.get('Total QTY', self.columns.get('QTY'))
            
            if not all([part_number_col, rev_col, qty_col]):
                self.logger.error("Mangler nødvendige kolonner for Partlist")
                return False
                
            # Kopier header række (undtagen Item)
            header_range = self.sheet.Range(self.sheet.Cells(1, 1), self.sheet.Cells(1, self.sheet.UsedRange.Columns.Count))
            header_range.Copy(partlist_sheet.Range("A1"))
            
            # Slet Item kolonne hvis den findes
            item_col = None
            for i in range(1, partlist_sheet.UsedRange.Columns.Count + 1):
                if str(partlist_sheet.Cells(1, i).Value).strip().upper() == "ITEM":
                    item_col = i
                    break
            if item_col:
                partlist_sheet.Columns(item_col).Delete()
            
            # Opret dictionary til at holde unikke Part Number/REV kombinationer
            unique_parts = {}
            last_row = self.sheet.UsedRange.Rows.Count
            
            # Scan alle rækker
            for row in range(2, last_row + 1):
                part_number = str(self.sheet.Cells(row, part_number_col).Value).strip()
                rev = str(self.sheet.Cells(row, rev_col).Value or '').strip()
                qty = float(self.sheet.Cells(row, qty_col).Value or 0)
                
                if not part_number:  # Spring tomme rækker over
                    continue
                    
                # Opret nøgle og gem data
                key = (part_number, rev)
                if key not in unique_parts:
                    # Gem alle celleværdier fra rækken
                    row_data = {}
                    for col_name, col_idx in self.columns.items():
                        if col_name != 'Item':  # Skip Item kolonne
                            value = self.sheet.Cells(row, col_idx).Value
                            row_data[col_name] = value
                    row_data['QTY'] = qty  # Brug QTY i stedet for Total QTY
                    unique_parts[key] = row_data
                else:
                    # Opdater kun QTY
                    unique_parts[key]['QTY'] += qty
                    
            # Sorter part numbers numerisk
            sorted_parts = sorted(unique_parts.items(), 
                                key=lambda x: (self._get_numeric_sort_key(x[0][0]), x[0][1]))
            
            # Indsæt data i Partlist fanen
            for idx, ((part_number, rev), data) in enumerate(sorted_parts, start=2):
                col = 1
                for col_name, col_idx in self.columns.items():
                    if col_name != 'Item':  # Skip Item kolonne
                        value = data.get(col_name)
                        partlist_sheet.Cells(idx, col).Value = value
                        col += 1
                        
            # Formater fanen
            self._format_worksheet(partlist_sheet)
            
            self.logger.info(f"Partlist oprettet med {len(sorted_parts)} unikke dele")
            return True
            
        except Exception as e:
            self.logger.error(f"Fejl under oprettelse af Partlist: {str(e)}", exc_info=True)
            return False
            
    def _get_numeric_sort_key(self, part_number: str) -> tuple:
        """
        Genererer en nøgle til numerisk sortering af part numbers
        """
        parts = part_number.split('-')
        key_parts = []
        
        for part in parts:
            # Forsøg at konvertere til nummer hvis muligt
            try:
                # Håndter decimaler (f.eks. 02.1)
                if '.' in part:
                    main, decimal = part.split('.')
                    key_parts.extend([int(main), int(decimal)])
                else:
                    key_parts.append(int(part))
            except ValueError:
                # Hvis ikke et nummer, brug string
                key_parts.append(part)
                
        return tuple(key_parts)

    def process_file(self, include_equipment: bool = False) -> bool:
        """
        Behandler Excel filen
        :param include_equipment: Om equipment rækker skal inkluderes
        :return: True hvis behandlingen var succesfuld
        """
        self.logger.info("Starter process_file")
        
        try:
            # Optimer Excel indstillinger
            self._optimize_excel_settings()
            
            # Omdøb aktivt sheet til "BOM (Raw)"
            self._rename_active_sheet("BOM (Raw)")
            
            # Identificer kolonner
            self._identify_columns()
            
            # Udtræk information fra filnavn
            file_info = self._extract_file_info(self.workbook.FullName)
            if file_info:
                self._insert_arrangement_row(file_info)
                
            # Håndter specielle rækker
            self._handle_special_rows()
            
            # Håndter equipment rækker
            if not include_equipment:
                self._handle_equipment_rows(include_equipment)
                
            # Indekser rækker
            self._index_rows()
            
            # Håndter BOM Structure regler
            self._handle_bom_structure()
            
            # Flyt revisionsnumre
            self._move_revisions()
            
            # Beregn Total QTY
            self._calculate_total_qty()
            
            # Kategoriser part numbers og opret faner
            self._categorize_part_numbers()
            
            # Formater worksheet
            self._format_worksheet()
            
            # Tilføj manglende kolonner
            self._add_missing_columns()
            
            # Opret Partlist fane
            self._create_partlist()
            
            # Gendan Excel indstillinger
            self._restore_excel_settings()
            
            # Gem ændringer
            self.workbook.Save()
            self.logger.debug("Workbook gemt")
            
            self.logger.info("BOM data behandling gennemført succesfuldt")
            return True
            
        except Exception as e:
            self.logger.error(f"Fejl under behandling af BOM data: {str(e)}", exc_info=True)
            return False
            
        finally:
            self._restore_excel_settings()

if __name__ == "__main__":
    # Test kode
    processor = ExcelDataProcessor("test.xlsx")
    processor.process_file() 