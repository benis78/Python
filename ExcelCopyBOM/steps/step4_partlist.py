"""
TRIN 4: Oprettelse af stykliste
Håndterer oprettelse af stykliste med unikke part numbers og deres totale mængder
"""

import logging
from collections import defaultdict

class PartListCreator:
    def __init__(self, workbook):
        self.logger = logging.getLogger('ExcelCopyBOM.PartList')
        self.workbook = workbook
        self.raw_sheet = workbook.Sheets("BOM (Raw)")
        
    def _get_numeric_sort_key(self, part_number: str) -> tuple:
        """
        Genererer en sorteringsnøgle for part numbers der håndterer både tal og tekst
        F.eks.: 4003-02.1-E01 kommer før 4003-99-BM005-02
        """
        if not part_number:
            return (0, '')  # Default værdi for tomme part numbers
            
        parts = []
        for part in part_number.split('-'):
            if not part:  # Håndter tomme dele
                parts.extend([0, ''])
                continue
                
            # Håndter decimaler i area numbers (f.eks. 02.1)
            if '.' in part:
                num_parts = part.split('.')
                try:
                    # Konverter første del til heltal hvis muligt
                    parts.append(int(num_parts[0]) if num_parts[0].isdigit() else 0)
                    # Konverter decimal del hvis muligt
                    if len(num_parts) > 1 and num_parts[1].isdigit():
                        parts.append(int(num_parts[1]))
                    else:
                        parts.append(0)
                        parts.append(num_parts[1] if len(num_parts) > 1 else '')
                except (ValueError, IndexError):
                    parts.extend([0, ''])
            else:
                # Find tal i starten af delen
                numeric_prefix = ''
                alpha_suffix = ''
                for char in part:
                    if char.isdigit():
                        numeric_prefix += char
                    else:
                        alpha_suffix += char
                        
                # Konverter til tal hvis muligt og tilføj alfabetisk del
                try:
                    parts.append(int(numeric_prefix) if numeric_prefix else 0)
                except ValueError:
                    parts.append(0)
                parts.append(alpha_suffix if alpha_suffix else '')
                    
        return tuple(parts)
        
    def _find_columns(self) -> dict:
        """Finder relevante kolonner i Excel arket"""
        columns = {}
        last_col = self.raw_sheet.UsedRange.Columns.Count
        
        # Find kolonne indeks
        for col in range(1, last_col + 1):
            header = str(self.raw_sheet.Cells(1, col).Value).strip()
            if header in ['Part Number', 'REV', 'QTY', 'Description', 'BOM Structure']:
                columns[header] = col
                
        return columns
        
    def create_partlist(self) -> bool:
        """
        Opretter "Partlist" sheet med unikke part numbers og deres totale mængder
        """
        try:
            self.logger.info("Opretter Partlist sheet")
            
            # Find kolonner
            columns = self._find_columns()
            if not all(col in columns for col in ['Part Number', 'REV', 'QTY']):
                self.logger.error("Manglende påkrævede kolonner")
                return False
                
            # Dictionary til at holde unikke part numbers og deres data
            unique_parts = defaultdict(lambda: {'qty': 0, 'data': {}})
            
            # Scan alle rækker
            last_row = self.raw_sheet.UsedRange.Rows.Count
            for row in range(2, last_row + 1):
                part_number = str(self.raw_sheet.Cells(row, columns['Part Number']).Value)
                rev = str(self.raw_sheet.Cells(row, columns['REV']).Value)
                qty = self.raw_sheet.Cells(row, columns['QTY']).Value
                
                if not part_number or not rev:
                    continue
                    
                # Konverter qty til tal
                try:
                    qty = float(qty)
                except (ValueError, TypeError):
                    self.logger.warning(f"Ugyldig mængde i række {row}: {qty}")
                    continue
                    
                # Nøgle er kombinationen af part number og revision
                key = (part_number, rev)
                
                # Opdater mængde og gem øvrig data
                unique_parts[key]['qty'] += qty
                if not unique_parts[key]['data']:
                    unique_parts[key]['data'] = {
                        'Description': str(self.raw_sheet.Cells(row, columns.get('Description', 1)).Value),
                        'BOM Structure': str(self.raw_sheet.Cells(row, columns.get('BOM Structure', 1)).Value)
                    }
                    
            # Opret nyt ark
            if "Partlist" in [sheet.Name for sheet in self.workbook.Sheets]:
                self.workbook.Sheets("Partlist").Delete()
            partlist_sheet = self.workbook.Sheets.Add(After=self.workbook.Sheets(self.workbook.Sheets.Count))
            partlist_sheet.Name = "Partlist"
            
            # Indsæt headers
            headers = ['Part Number', 'REV', 'Total QTY', 'Description', 'BOM Structure']
            for col, header in enumerate(headers, 1):
                partlist_sheet.Cells(1, col).Value = header
                
            # Sortér part numbers numerisk
            sorted_parts = sorted(unique_parts.items(), key=lambda x: self._get_numeric_sort_key(x[0][0]))
            
            # Indsæt data
            for row, ((part_number, rev), data) in enumerate(sorted_parts, 2):
                partlist_sheet.Cells(row, 1).Value = part_number
                partlist_sheet.Cells(row, 2).Value = rev
                partlist_sheet.Cells(row, 3).Value = data['qty']
                partlist_sheet.Cells(row, 4).Value = data['data']['Description']
                partlist_sheet.Cells(row, 5).Value = data['data']['BOM Structure']
                
            self.logger.info(f"Partlist oprettet med {len(unique_parts)} unikke dele")
            return True
            
        except Exception as e:
            self.logger.error(f"Fejl under oprettelse af partlist: {str(e)}", exc_info=True)
            return False 